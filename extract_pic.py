#!/usr/bin/env python3
"""
convert_docx_images_noninteractive.py

非交互版：脚本内固定输入 .docx 路径与输出选项，直接运行 main() 会处理并退出。

请编辑变量 INPUT_DOCX_PATH 和 OUTPUT_FORMAT 在脚本顶部按需调整。
"""

from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
import sys
import shutil
import io
import tempfile
import subprocess
import os
from collections import OrderedDict

# Optional deps
try:
    from PIL import Image, UnidentifiedImageError
except Exception:
    print("错误: 需要安装 Pillow。请运行: pip install pillow")
    sys.exit(1)

try:
    import olefile
except Exception:
    olefile = None  # optional

# ----- Configuration: 修改这里来指定输入文件和输出选项 -----
INPUT_DOCX_PATH = r"example2.docx"   # <-- 把这里改成你要处理的 .docx 的真实路径（可以是相对或绝对路径）
OUTPUT_FORMAT = "png"             # 'png' 或 'jpg'
KEEP_ORIGINALS = False            # 是否把所有原始媒体也复制到 originals/ 下（True/False）
# ----------------------------------------------------------------

PAD_DIGITS = 3  # image001
NS = {
    'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    'a': "http://schemas.openxmlformats.org/drawingml/2006/main",
    'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    'wp': "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    'pic': "http://schemas.openxmlformats.org/drawingml/2006/picture",
    'v': "urn:schemas-microsoft-com:vml"
}
COMMON_RASTER_EXTS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff', '.webp'}
VECTOR_EXTS = {'.wmf', '.emf', '.svg'}
BIN_EXT = '.bin'

# detect external tools
def find_tool(names):
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    return None

IMAGEMAGICK = find_tool(['magick', 'convert'])
SOFFICE = find_tool(['soffice', 'libreoffice'])
INKSCAPE = find_tool(['inkscape'])

def normalize_target(t):
    t = t.replace('\\', '/')
    if '?' in t:
        t = t.split('?', 1)[0]
    if t.startswith('../'):
        t = t[3:]
    if not t.startswith('word/'):
        t = 'word/' + t
    return t

def parse_rels_from_zip(z: zipfile.ZipFile):
    rels = {}
    try:
        data = z.read('word/_rels/document.xml.rels')
    except KeyError:
        return rels
    root = ET.fromstring(data)
    for rel in root.findall('.//'):
        tag = rel.tag.lower()
        if tag.endswith('relationship'):
            rid = rel.get('Id') or rel.get('ID')
            target = rel.get('Target')
            if rid and target:
                rels[rid] = normalize_target(target)
    return rels

def collect_image_refs_in_doc(z: zipfile.ZipFile):
    results = []
    try:
        data = z.read('word/document.xml')
    except KeyError:
        return results
    root = ET.fromstring(data)
    for node in root.iter():
        drawing = node.find('.//w:drawing', NS)
        if drawing is not None:
            inline = drawing.find('.//wp:inline', NS) or drawing.find('.//wp:anchor', NS)
            if inline is not None:
                docPr = inline.find('.//wp:docPr', NS)
                img_name = docPr.get('name') if docPr is not None else None
                if not img_name:
                    img_name = 'Image'
                blip = inline.find('.//a:blip', NS)
                if blip is not None:
                    embed = blip.get(f'{{{NS["r"]}}}embed')
                    if embed:
                        results.append((embed, img_name))
        for im in node.findall('.//{*}imagedata'):
            rel_id = None
            for attr_name, attr_val in im.items():
                if attr_name.endswith('}id') or attr_name == 'r:id' or attr_name.lower().endswith(':id'):
                    rel_id = attr_val
                    break
            if rel_id:
                img_name = im.get('o:title') or im.get('title') or im.get('alt') or 'Image'
                results.append((rel_id, img_name))
    return results

def starts_with(b: bytes, sig: bytes):
    return b[:len(sig)] == sig

def guess_image_by_magic(b: bytes):
    if len(b) < 8:
        return None
    if starts_with(b, b'\x89PNG\r\n\x1a\n'):
        return ('PNG', '.png')
    if starts_with(b, b'\xff\xd8\xff'):
        return ('JPEG', '.jpg')
    if starts_with(b, b'GIF87a') or starts_with(b, b'GIF89a'):
        return ('GIF', '.gif')
    if starts_with(b, b'BM'):
        return ('BMP', '.bmp')
    if starts_with(b, b'II*\x00') or starts_with(b, b'MM\x00*'):
        return ('TIFF', '.tif')
    return None

def pillow_guess(b: bytes):
    try:
        im = Image.open(io.BytesIO(b))
        fmt = im.format
        im.close()
        ext = '.' + fmt.lower()
        return (fmt, ext)
    except Exception:
        return None

def save_with_pillow(b: bytes, out_path: Path, out_format: str):
    try:
        im = Image.open(io.BytesIO(b))
        if out_format.lower() == 'jpg' and im.mode in ('RGBA', 'LA', 'P'):
            im = im.convert('RGB')
        if out_format.lower() == 'png' and im.mode == 'P':
            im = im.convert('RGBA')
        im.save(out_path, 'PNG' if out_format.lower()=='png' else 'JPEG')
        im.close()
        return True
    except Exception as e:
        print(f"  [WARN] Pillow 保存失败: {e}")
        return False

def convert_vector_with_imagemagick(tmp_ext: str, bytes_src: bytes, out_path: Path):
    if not IMAGEMAGICK:
        return False
    with tempfile.NamedTemporaryFile(delete=False, suffix=tmp_ext) as tf:
        tf.write(bytes_src)
        tmpname = tf.name
    try:
        cmd = [IMAGEMAGICK, tmpname, str(out_path)]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return out_path.exists()
    except Exception as e:
        print(f"  [WARN] ImageMagick 转换失败: {e}")
        return False
    finally:
        try:
            os.unlink(tmpname)
        except Exception:
            pass

def convert_with_inkscape_if_svg(bytes_src: bytes, out_path: Path):
    if not INKSCAPE:
        return False
    with tempfile.NamedTemporaryFile(delete=False, suffix='.svg') as tf:
        tf.write(bytes_src)
        tmpname = tf.name
    try:
        cmd = [INKSCAPE, tmpname, '--export-type=png', '--export-filename', str(out_path)]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return out_path.exists()
    except Exception as e:
        print(f"  [WARN] Inkscape 转换 SVG 失败: {e}")
        return False
    finally:
        try:
            os.unlink(tmpname)
        except Exception:
            pass

def convert_with_soffice(bytes_src: bytes, src_ext: str, out_path: Path):
    if not SOFFICE:
        return False
    with tempfile.TemporaryDirectory() as td:
        tmp_in = Path(td) / ('in' + src_ext)
        tmp_in.write_bytes(bytes_src)
        try:
            cmd = [SOFFICE, '--headless', '--convert-to', 'png', '--outdir', td, str(tmp_in)]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            cand = tmp_in.with_suffix('.png')
            if cand.exists():
                shutil.move(str(cand), str(out_path))
                return True
            for f in Path(td).glob('*.png'):
                shutil.move(str(f), str(out_path))
                return True
        except Exception as e:
            print(f"  [WARN] soffice 转换失败: {e}")
            return False
    return False

def try_extract_from_ole(bin_bytes: bytes):
    results = []
    if olefile is None:
        return results
    with tempfile.NamedTemporaryFile(delete=False) as tf:
        tf.write(bin_bytes)
        tmpname = tf.name
    try:
        if not olefile.isOleFile(tmpname):
            return results
        ole = olefile.OleFileIO(tmpname)
        for entry in ole.listdir(streams=True, storages=False):
            try:
                data = ole.openstream(entry).read()
            except Exception:
                continue
            if pillow_guess(data) or guess_image_by_magic(data):
                results.append(("/".join(entry), data))
        ole.close()
    except Exception:
        pass
    finally:
        try:
            os.unlink(tmpname)
        except Exception:
            pass
    return results

def zero_pad(n):
    # return str(n).zfill(PAD_DIGITS)
    return str(n)

def convert_docx(docx_path: Path, out_format: str = 'png', keep_originals: bool = False):
    if not docx_path.exists():
        print("文件不存在:", docx_path)
        return 1
    doc_name = docx_path.stem
    out_root = Path("Images") / f"{doc_name}_converted"
    originals_dir = out_root / "originals"
    out_root.mkdir(parents=True, exist_ok=True)
    originals_dir.mkdir(parents=True, exist_ok=True)

    saved = []
    counter = 1
    processed_targets = set()

    with zipfile.ZipFile(docx_path, 'r') as z:
        rel_map = parse_rels_from_zip(z)
        refs = collect_image_refs_in_doc(z)

        for rid, title in refs:
            target = rel_map.get(rid)
            if not target:
                print(f"[WARN] document 引用 rId={rid} 未在 rels 中找到 target，跳过")
                continue
            if target.lower().startswith('http://') or target.lower().startswith('https://'):
                print(f"[WARN] rId={rid} 指向外部资源 {target}，跳过")
                continue
            if target in processed_targets:
                for e in saved:
                    if e.get('source') == target:
                        saved.append({
                            'source': target,
                            'rId': rid,
                            'title': title,
                            'status': 'duplicate_reference',
                            'outfile': e.get('outfile'),
                            'note': 'duplicate document reference to same target'
                        })
                        break
                continue

            try:
                b = z.read(target)
            except KeyError:
                print(f"[WARN] 无法从 zip 中读取 {target}（rId={rid}），跳过")
                processed_targets.add(target)
                saved.append({
                    'source': target,
                    'rId': rid,
                    'title': title,
                    'status': 'missing_in_zip',
                    'outfile': None
                })
                continue

            processed_targets.add(target)
            ext = Path(target).suffix.lower()
            print(f"[INFO] 处理 rId={rid} -> {target} (ext={ext}) title={title}")

            pillow_guess_res = pillow_guess(b)
            magic_guess = guess_image_by_magic(b)
            out_name = f"image{zero_pad(counter)}.{ 'png' if out_format.lower()=='png' else 'jpg' }"
            out_path = out_root / out_name

            if pillow_guess_res:
                ok = save_with_pillow(b, out_path, out_format)
                if ok:
                    print(f"  [OK] 使用 Pillow 保存为 {out_path.name}")
                    saved.append({
                        'source': target,
                        'rId': rid,
                        'title': title,
                        'status': 'converted',
                        'outfile': str(out_path),
                        'orig_ext': ext,
                        'orig_magic': pillow_guess_res[0]
                    })
                    counter += 1
                    if keep_originals:
                        (originals_dir / Path(target).name).write_bytes(b)
                    continue

            if ext in VECTOR_EXTS:
                converted = False
                if IMAGEMAGICK:
                    converted = convert_vector_with_imagemagick(ext, b, out_path)
                if not converted and ext == '.svg' and INKSCAPE:
                    converted = convert_with_inkscape_if_svg(b, out_path)
                if not converted and SOFFICE:
                    converted = convert_with_soffice(b, ext, out_path)
                if converted:
                    print(f"  [OK] 矢量文件转换成功 -> {out_path.name}")
                    saved.append({
                        'source': target,
                        'rId': rid,
                        'title': title,
                        'status': 'converted_vector',
                        'outfile': str(out_path),
                        'orig_ext': ext
                    })
                    counter += 1
                    if keep_originals:
                        (originals_dir / Path(target).name).write_bytes(b)
                    continue
                else:
                    orig_path = originals_dir / Path(target).name
                    orig_path.write_bytes(b)
                    print(f"  [WARN] 无法将矢量/特殊格式 {target} 自动转换，已保存原始文件到 {orig_path}")
                    saved.append({
                        'source': target,
                        'rId': rid,
                        'title': title,
                        'status': 'saved_original_vector',
                        'outfile': str(orig_path),
                        'orig_ext': ext
                    })
                    continue

            if ext == BIN_EXT:
                if magic_guess:
                    ok = save_with_pillow(b, out_path, out_format)
                    if ok:
                        print(f"  [OK] .bin 内含可识别图片，已保存 -> {out_path.name}")
                        saved.append({
                            'source': target,
                            'rId': rid,
                            'title': title,
                            'status': 'bin_magic_converted',
                            'outfile': str(out_path),
                            'orig_ext': ext,
                            'orig_magic': magic_guess[0]
                        })
                        counter += 1
                        if keep_originals:
                            (originals_dir / Path(target).name).write_bytes(b)
                        continue
                ole_extracted = try_extract_from_ole(b)
                if ole_extracted:
                    for sname, sbytes in ole_extracted:
                        inner_guess = pillow_guess(sbytes) or guess_image_by_magic(sbytes)
                        inner_out_name = f"image{zero_pad(counter)}.{ 'png' if out_format.lower()=='png' else 'jpg' }"
                        inner_out = out_root / inner_out_name
                        if inner_guess:
                            ok = save_with_pillow(sbytes, inner_out, out_format)
                            if ok:
                                print(f"  [OK] 从 OLE (.bin) 提取并转换流 {sname} -> {inner_out.name}")
                                saved.append({
                                    'source': target + '::' + sname,
                                    'rId': rid,
                                    'title': title,
                                    'status': 'ole_extracted_converted',
                                    'outfile': str(inner_out)
                                })
                                counter += 1
                                continue
                        raw_name = f"{Path(target).stem}_{sname.replace('/', '_')}.bin"
                        raw_path = originals_dir / raw_name
                        raw_path.write_bytes(sbytes)
                        print(f"  [INFO] 从 OLE 提取流 {sname}，但无法识别为图片，已保存 {raw_path}")
                        saved.append({
                            'source': target + '::' + sname,
                            'rId': rid,
                            'title': title,
                            'status': 'ole_extracted_raw',
                            'outfile': str(raw_path)
                        })
                    if keep_originals:
                        (originals_dir / Path(target).name).write_bytes(b)
                    continue

                orig_path = originals_dir / Path(target).name
                orig_path.write_bytes(b)
                print(f"  [WARN] .bin 未能解析，已保存原始 .bin 到 {orig_path}")
                saved.append({
                    'source': target,
                    'rId': rid,
                    'title': title,
                    'status': 'saved_bin',
                    'outfile': str(orig_path),
                    'orig_ext': ext
                })
                continue

            if magic_guess:
                ok = save_with_pillow(b, out_path, out_format)
                if ok:
                    print(f"  [OK] 基于魔数识别后保存为 {out_path.name}")
                    saved.append({
                        'source': target,
                        'rId': rid,
                        'title': title,
                        'status': 'magic_converted',
                        'outfile': str(out_path),
                        'orig_ext': ext,
                        'orig_magic': magic_guess[0]
                    })
                    counter += 1
                    if keep_originals:
                        (originals_dir / Path(target).name).write_bytes(b)
                    continue

            orig_path = originals_dir / Path(target).name
            orig_path.write_bytes(b)
            print(f"  [WARN] 无法识别或转换 {target}，已保存到 originals")
            saved.append({
                'source': target,
                'rId': rid,
                'title': title,
                'status': 'saved_original',
                'outfile': str(orig_path),
                'orig_ext': ext
            })

        media_files = [n for n in z.namelist() if n.startswith('word/media/') and not n.endswith('/')]
        unreferenced = []
        for m in media_files:
            if m in processed_targets:
                continue
            if any(Path(m).name == Path(entry.get('source','')).name for entry in saved):
                continue
            unreferenced.append(m)

        if unreferenced:
            print(f"[INFO] 发现 {len(unreferenced)} 个未在 document.xml 中引用的 media 文件，将尝试处理它们")

        for m in unreferenced:
            try:
                b = z.read(m)
            except KeyError:
                continue
            ext = Path(m).suffix.lower()
            print(f"[INFO] 处理未引用 media: {m} (ext={ext})")
            out_name = f"image{zero_pad(counter)}.{ 'png' if out_format.lower()=='png' else 'jpg' }"
            out_path = out_root / out_name

            pillow_guess_res = pillow_guess(b)
            magic_guess = guess_image_by_magic(b)

            if pillow_guess_res:
                ok = save_with_pillow(b, out_path, out_format)
                if ok:
                    print(f"  [OK] 保存为 {out_path.name}")
                    saved.append({
                        'source': m,
                        'rId': None,
                        'title': None,
                        'status': 'converted_unreferenced',
                        'outfile': str(out_path),
                        'orig_ext': ext
                    })
                    counter += 1
                    if keep_originals:
                        (originals_dir / Path(m).name).write_bytes(b)
                    continue

            if ext in VECTOR_EXTS:
                converted = False
                if IMAGEMAGICK:
                    converted = convert_vector_with_imagemagick(ext, b, out_path)
                if not converted and ext == '.svg' and INKSCAPE:
                    converted = convert_with_inkscape_if_svg(b, out_path)
                if not converted and SOFFICE:
                    converted = convert_with_soffice(b, ext, out_path)
                if converted:
                    print(f"  [OK] 矢量转换成功 -> {out_path.name}")
                    saved.append({
                        'source': m,
                        'rId': None,
                        'title': None,
                        'status': 'converted_unreferenced_vector',
                        'outfile': str(out_path),
                        'orig_ext': ext
                    })
                    counter += 1
                    if keep_originals:
                        (originals_dir / Path(m).name).write_bytes(b)
                    continue
                else:
                    outp = originals_dir / Path(m).name
                    outp.write_bytes(b)
                    print(f"  [WARN] 未能转换矢量文件，已保存原始 {outp}")
                    saved.append({
                        'source': m,
                        'rId': None,
                        'title': None,
                        'status': 'saved_unreferenced_vector',
                        'outfile': str(outp),
                        'orig_ext': ext
                    })
                    continue

            if ext == BIN_EXT:
                if magic_guess:
                    ok = save_with_pillow(b, out_path, out_format)
                    if ok:
                        print(f"  [OK] 未引用的 .bin 含图片，已保存 -> {out_path.name}")
                        saved.append({
                            'source': m,
                            'rId': None,
                            'title': None,
                            'status': 'bin_magic_converted_unreferenced',
                            'outfile': str(out_path),
                            'orig_ext': ext
                        })
                        counter += 1
                        if keep_originals:
                            (originals_dir / Path(m).name).write_bytes(b)
                        continue
                ole_extracted = try_extract_from_ole(b)
                if ole_extracted:
                    for sname, sbytes in ole_extracted:
                        inner_out_name = f"image{zero_pad(counter)}.{ 'png' if out_format.lower()=='png' else 'jpg' }"
                        inner_out = out_root / inner_out_name
                        ok = save_with_pillow(sbytes, inner_out, out_format)
                        if ok:
                            print(f"  [OK] 从未引用的 OLE (.bin) 提取并保存 {inner_out.name}")
                            saved.append({
                                'source': m + '::' + sname,
                                'rId': None,
                                'title': None,
                                'status': 'ole_extracted_converted_unreferenced',
                                'outfile': str(inner_out)
                            })
                            counter += 1
                            continue
                        else:
                            rp = originals_dir / f"{Path(m).stem}_{sname.replace('/', '_')}.bin"
                            rp.write_bytes(sbytes)
                            saved.append({
                                'source': m + '::' + sname,
                                'rId': None,
                                'title': None,
                                'status': 'ole_extracted_raw_unreferenced',
                                'outfile': str(rp)
                            })
                    if keep_originals:
                        (originals_dir / Path(m).name).write_bytes(b)
                    continue
                rp = originals_dir / Path(m).name
                rp.write_bytes(b)
                print(f"  [WARN] 未引用 .bin 无法解析，已保存原始 {rp}")
                saved.append({
                    'source': m,
                    'rId': None,
                    'title': None,
                    'status': 'saved_unreferenced_bin',
                    'outfile': str(rp),
                    'orig_ext': ext
                })
                continue

            if magic_guess:
                ok = save_with_pillow(b, out_path, out_format)
                if ok:
                    saved.append({
                        'source': m,
                        'rId': None,
                        'title': None,
                        'status': 'magic_converted_unreferenced',
                        'outfile': str(out_path),
                        'orig_ext': ext
                    })
                    counter += 1
                    if keep_originals:
                        (originals_dir / Path(m).name).write_bytes(b)
                    continue

            rp = originals_dir / Path(m).name
            rp.write_bytes(b)
            print(f"  [WARN] 未知未引用文件 {m}，已保存原始 {rp}")
            saved.append({
                'source': m,
                'rId': None,
                'title': None,
                'status': 'saved_unreferenced_original',
                'outfile': str(rp),
                'orig_ext': ext
            })

    total_converted = sum(1 for s in saved if s.get('status','').startswith('converted') or s.get('status','')=='magic_converted')
    total_saved_originals = sum(1 for s in saved if s.get('status','').startswith('saved') or s.get('status','').startswith('ole_extracted_raw'))
    print("\n=== 完成 ===")
    print(f"输出目录: {out_root}")
    print(f"已生成图像: {total_converted}, 原始/未能转换文件: {total_saved_originals}")
    report_path = out_root / 'report.txt'
    with report_path.open('w', encoding='utf-8') as rf:
        rf.write(f"Report for {docx_path.name}\n\n")
        for item in saved:
            rf.write(str(item) + "\n")
    print(f"报告已写入 {report_path}")
    return 0

def main():
    # no-interaction: use constants defined at top
    global INPUT_DOCX_PATH, OUTPUT_FORMAT, KEEP_ORIGINALS
    docx = Path(INPUT_DOCX_PATH)
    fmt = OUTPUT_FORMAT
    keep_orig = KEEP_ORIGINALS
    rc = convert_docx(docx, out_format=fmt, keep_originals=keep_orig)
    # exit with code rc
    sys.exit(rc)

if __name__ == '__main__':
    main()
