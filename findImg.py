#!/usr/bin/env python3
"""
count_docx_images.py（零交互版）

将要检查的 .docx 路径写在 INPUT_PATH 中，运行脚本会直接打印统计结果并退出。
"""

from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict, Counter

# ======== 只需修改这一行 ========
INPUT_PATH = r"example2.docx"
# ===============================

COMMON_IMAGE_EXTS = {
    '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff',
    '.wmf', '.emf', '.wmz', '.svg', '.ico', '.webp'
}

NS = {
    'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    'a': "http://schemas.openxmlformats.org/drawingml/2006/main",
    'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    'wp': "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    'pic': "http://schemas.openxmlformats.org/drawingml/2006/picture",
    'v': "urn:schemas-microsoft-com:vml"
}

def list_media_files(z):
    media = []
    for name in z.namelist():
        lname = name.lower()
        if lname.startswith('word/media/') and not name.endswith('/'):
            media.append(name)
        elif lname.startswith('word/embeddings/') and not name.endswith('/'):
            media.append(name)
    return sorted(media)

def classify_by_ext(media_list):
    by_ext = defaultdict(list)
    for path in media_list:
        ext = Path(path).suffix.lower()
        by_ext[ext].append(path)
    return by_ext

def parse_document_xml_for_rel_refs(z):
    refs = Counter()
    try:
        xml_bytes = z.read('word/document.xml')
    except KeyError:
        return refs

    try:
        root = ET.fromstring(xml_bytes)
    except Exception as e:
        print(f"[WARN] 无法解析 word/document.xml: {e}")
        return refs

    for blip in root.findall('.//{'+NS['a']+'}blip'):
        rid = blip.get('{'+NS['r']+'}embed')
        if rid:
            refs[rid] += 1

    for im in root.findall('.//{'+NS['v']+'}imagedata'):
        rid = None
        for k, v in im.items():
            if k.endswith('}id') or k == 'r:id' or k.lower().endswith(':id'):
                rid = v
                break
        if rid:
            refs[rid] += 1

    for blip in root.findall('.//{'+NS['pic']+'}pic//{'+NS['a']+'}blip'):
        rid = blip.get('{'+NS['r']+'}embed')
        if rid:
            refs[rid] += 1

    return refs

def parse_relationships_for_media(z):
    rels_name = 'word/_rels/document.xml.rels'
    mapping = {}
    try:
        rels_bytes = z.read(rels_name)
    except KeyError:
        return mapping

    try:
        root = ET.fromstring(rels_bytes)
    except Exception as e:
        print(f"[WARN] 无法解析 {rels_name}: {e}")
        return mapping

    for rel in root.findall('.//'):
        tag = rel.tag
        if tag is None:
            continue
        if tag.lower().endswith('relationship'):
            rid = rel.get('Id') or rel.get('ID') or rel.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            target = rel.get('Target')
            if rid and target:
                t = target.replace('\\', '/')
                if not t.startswith('word/'):
                    if t.startswith('../'):
                        t = t[3:]
                    if not t.startswith('word/'):
                        t = 'word/' + t
                mapping[rid] = t
    return mapping

def main(docx_path):
    p = Path(docx_path)
    if not p.exists():
        print(f"File not found: {p}")
        return 2

    with zipfile.ZipFile(p, 'r') as z:
        media_files = list_media_files(z)
        by_ext = classify_by_ext(media_files)
        refs = parse_document_xml_for_rel_refs(z)
        rel_map = parse_relationships_for_media(z)

        print(f"\nFound {len(media_files)} media/embeddings files in the .docx:")
        for m in media_files:
            ext = Path(m).suffix.lower() or '(no-ext)'
            print(f"  - {m} ({ext})")

        print("\nSummary by extension:")
        ext_counts = { (ext if ext else '(no-ext)'): len(lst) for ext, lst in by_ext.items() }
        for ext, cnt in sorted(ext_counts.items(), key=lambda x: -x[1]):
            print(f"  {ext}: {cnt}")

        if rel_map:
            print("\nRelationships in word/_rels/document.xml.rels (rId -> target):")
            for rid, target in rel_map.items():
                print(f"  {rid} -> {target}")

        if refs:
            print("\nReferences found in word/document.xml (rId -> count):")
            for rid, cnt in refs.items():
                mapped = rel_map.get(rid, '(no target in rels)')
                print(f"  {rid} -> referenced {cnt} time(s) ; target: {mapped}")

        referenced_targets = set(rel_map.get(rid) for rid in refs.keys() if rel_map.get(rid))
        referenced_targets = {t for t in referenced_targets if t}
        if referenced_targets:
            print("\nMedia files referenced from document.xml (via relationships):")
            for t in sorted(referenced_targets):
                print(f"  - {t}")
        else:
            print("\nNo media files referenced from document.xml via relationships were found.")

        known_image_counts = Counter()
        unknown_files = []
        for m in media_files:
            ext = Path(m).suffix.lower()
            if ext in COMMON_IMAGE_EXTS:
                known_image_counts[ext] += 1
            else:
                unknown_files.append(m)

        if known_image_counts:
            print("\nKnown image extension counts:")
            for ext, cnt in sorted(known_image_counts.items(), key=lambda x: -x[1]):
                print(f"  {ext}: {cnt}")
        if unknown_files:
            print("\nOther media/embeddings files (not in COMMON_IMAGE_EXTS):")
            for u in unknown_files:
                print(f"  - {u}")

        total_images = sum(known_image_counts.values())
        print(f"\nTotal known image files under word/media or word/embeddings: {total_images}")
    return 0

if __name__ == '__main__':
    # 直接以固定路径运行，无交互
    print(f"Scanning: {INPUT_PATH}")
    rc = main(INPUT_PATH)
    # 可根据需要退出码处理，但脚本会直接结束
