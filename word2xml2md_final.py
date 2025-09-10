#!/usr/bin/env python3
# convert_docx_xml_to_md.py


import xml.etree.ElementTree as ET
from pathlib import Path
import zipfile
import re
from docx import Document
from lxml import etree

# ==========================
# 只需修改这一行：传入你的 document.xml 或 .docx 路径（字符串）
INPUT_PATH = r"example2.docx"
# ==========================

# 定义输出目录
XML_OUTPUT_DIR = "XMLFile"
MD_OUTPUT_DIR = "MarkdownFile"
IMAGE_OUTPUT_DIR = "Images"
IMAGE_COUNT = 1
Path(IMAGE_OUTPUT_DIR).mkdir(exist_ok=True)
# 创建输出目录（如果不存在）
Path(XML_OUTPUT_DIR).mkdir(exist_ok=True)
Path(MD_OUTPUT_DIR).mkdir(exist_ok=True)
"""
这部分是把word转为xml文件的
"""
def dump_document_xml_via_python_docx(docx_path, output_xml_path, pretty_print=True):
    """
    使用 python-docx 打开文档，序列化 document.xml 并保存到文件。

    参数:
        docx_path: docx文件的路径
        output_xml_path: 输出XML文件的路径，默认为"document.xml"
        pretty_print: 是否格式化XML输出，默认为True
    """
    doc = Document(docx_path)
    # 获取document.xml的底层element
    doc_element = doc._element  # 这是 lxml 元素，表示 <w:document>

    # 将XML序列化为字节
    xml_bytes = etree.tostring(doc_element, pretty_print=pretty_print, encoding='utf-8', xml_declaration=True)
    xml_content = xml_bytes.decode('utf-8')

    # 保存到文件
    with open(output_xml_path, 'w', encoding='utf-8') as f:
        f.write(xml_content)

    print(f"\nXML内容已保存到 {output_xml_path}")


def extract_and_number_images(docx_path):
    """
    从.docx中提取所有图片并按顺序编号保存到Images/{docName}_images文件夹

    参数:
        docx_path: .docx文件路径

    返回:
        - 图片保存路径列表 (Images/{docName}_images/image1.png, ...)
        - 图片总数
    """
    # 获取文档名并创建对应的图片文件夹
    doc_name = Path(docx_path).stem
    images_dir = Path("Images") / f"{doc_name}_images"
    images_dir.mkdir(parents=True, exist_ok=True)

    print(f"[图片提取] 图片将保存到: {images_dir}")

    image_paths = []

    with zipfile.ZipFile(docx_path, 'r') as z:
        # 获取所有图片文件（支持多种格式）
        media_files = [f for f in z.namelist()
                       if 'media' in f.lower() and
                       any(f.lower().endswith(ext) for ext in ['.png'])]

        print(f"[图片提取] 找到 {len(media_files)} 个图片文件")

        # 分离有数字和无数字的文件
        numbered_files = []
        unnumbered_files = []

        for f in media_files:
            filename = Path(f).stem.lower()
            if re.search(r'\d+$', filename):  # 如果文件名以数字结尾
                numbered_files.append(f)
            else:
                unnumbered_files.append(f)

        # 首先处理无数字的文件（视为第一个）
        print("[图片提取] 处理无数字的图片文件...")
        for i, media_file in enumerate(unnumbered_files):
            # 获取文件扩展名
            ext = Path(media_file).suffix.lower()
            if not ext:  # 如果没有扩展名，默认使用.png
                ext = '.png'

            output_filename = f"image{i + 1}{ext}"
            output_path = images_dir / output_filename

            # 保存图片
            with open(output_path, 'wb') as f:
                f.write(z.read(media_file))

            image_paths.append(str(output_path))
            print(f"  - 保存: {media_file} -> {output_path}")

        # 然后处理有数字的文件，按数字排序
        print("[图片提取] 处理有数字的图片文件...")
        numbered_files.sort(key=lambda x: int(re.search(r'\d+$', Path(x).stem.lower()).group()))

        start_index = len(unnumbered_files) + 1
        for i, media_file in enumerate(numbered_files):
            # 保留原始文件扩展名
            ext = Path(media_file).suffix.lower()
            output_filename = f"image{start_index + i}{ext}"
            output_path = images_dir / output_filename

            # 保存图片
            with open(output_path, 'wb') as f:
                f.write(z.read(media_file))

            image_paths.append(str(output_path))

    print(f"[图片提取] 完成！共保存 {len(image_paths)} 张图片")
    return image_paths, len(image_paths)

NS = {
    'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    'm': "http://schemas.openxmlformats.org/officeDocument/2006/math",
    'mml': "http://www.w3.org/1998/Math/MathML",
    'wp': "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    'a': "http://schemas.openxmlformats.org/drawingml/2006/main",
    'pic': "http://schemas.openxmlformats.org/drawingml/2006/picture",
    'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)

def strip_ns(tag):
    if tag is None:
        return ''
    if '}' in tag:
        return tag.split('}', 1)[1]
    return tag


def has_underline_format(r):
    """判断<w:r>节点是否包含下划线格式（<w:u>标签）"""
    rPr = r.find('w:rPr', NS)
    if rPr is not None:
        # 存在<w:u>标签即视为有下划线格式
        return rPr.find('w:u', NS) is not None
    return False

def text_of_run(r):
    parts = []
    # 检查当前run是否有下划线格式
    is_underlined = has_underline_format(r)
    for t in r.findall('.//w:t', NS):
        if t.text:
            if is_underlined:
                # 有下划线时，将所有空格（包括全角空格）替换为下划线
                processed_text = re.sub(r'[\s　]', '_', t.text)
                parts.append(processed_text)
            else:
                parts.append(t.text)
    for br in r.findall('.//w:br', NS):
        parts.append('\n')
    return ''.join(parts)


def get_run_vertical_align(r):
    rPr = r.find('w:rPr', NS)
    if rPr is not None:
        vertAlign = rPr.find('w:vertAlign', NS)
        if vertAlign is not None:
            val = vertAlign.get(f'{{{NS["w"]}}}val')
            return val
    return None


def node_text_content(node):
    if node is None:
        return ''
    texts = []
    for t in node.findall('.//m:t', NS):
        if t.text:
            texts.append(t.text)
    if not texts:
        for t in node.findall('.//w:t', NS):
            if t.text:
                texts.append(t.text)
    return ''.join(texts)


def convert_math_operator(text):
    operator_map = {
        '×': r' \times ',
        '⋅': r' \cdot ',
        '÷': r' \div ',
        '±': r' \pm ',
        '∓': r' \mp ',
        '≤': r' \leq ',
        '≥': r' \geq ',
        '≠': r' \neq ',
        '≈': r' \approx ',
        '∞': r' \infty ',
        '∑': r' \sum ',
        '∏': r' \prod ',
        '∫': r' \int ',
        '√': r' \sqrt ',
        'α': r' \alpha ',
        'β': r' \beta ',
        'γ': r' \gamma ',
        'δ': r' \delta ',
        'π': r' \pi ',
        'θ': r' \theta ',
        'λ': r' \lambda ',
        'μ': r' \mu ',
        'σ': r' \sigma ',
        'φ': r' \phi ',
        'ω': r' \omega ',
    }
    for symbol, latex in operator_map.items():
        text = text.replace(symbol, latex)
    text = re.sub(r'÷(\d+)', r' \\div \1', text)
    return text


def omml_to_latex(node):
    if node is None:
        return ''
    tag = strip_ns(node.tag)

    if tag in ('oMath', 'oMathPara'):
        return ''.join(omml_to_latex(child) for child in node)

    if tag in ('r', 'm:r'):
        txt = node_text_content(node)
        rPr = node.find('w:rPr', NS)
        if rPr is not None:
            vertAlign = rPr.find('w:vertAlign', NS)
            if vertAlign is not None:
                val = vertAlign.get(f'{{{NS["w"]}}}val')
                if val == 'superscript':
                    return '^{' + convert_math_operator(txt) + '}'
                if val == 'subscript':
                    return '_{' + convert_math_operator(txt) + '}'
        if txt:
            return convert_math_operator(txt)
        return ''.join(omml_to_latex(child) for child in node)

    if tag in ('t', 'm:t'):
        text = node.text or ''
        return convert_math_operator(text)

    if tag in ('f', 'frac'):
        num_node = node.find('m:num', NS)
        den_node = node.find('m:den', NS)
        if num_node is None:
            for child in node:
                if strip_ns(child.tag) == 'num':
                    num_node = child
                    break
        if den_node is None:
            for child in node:
                if strip_ns(child.tag) == 'den':
                    den_node = child
                    break
        num_text = ''.join(omml_to_latex(c) for c in num_node) if num_node is not None else ''
        den_text = ''.join(omml_to_latex(c) for c in den_node) if den_node is not None else ''
        return r'\frac{' + num_text + '}{' + den_text + '}'

    if tag in ('num', 'den'):
        return ''.join(omml_to_latex(child) for child in node)

    if tag in ('sSup', 'sSub', 'sSupSub'):
        base = node.find('m:base', NS)
        if tag == 'sSup':
            sup_node = node.find('m:sup', NS)
            sup_l = ''.join(omml_to_latex(c) for c in sup_node) if sup_node is not None else ''
            base_l = ''.join(omml_to_latex(c) for c in base) if base is not None else ''
            return base_l + '^{' + sup_l + '}'
        if tag == 'sSub':
            sub_node = node.find('m:sub', NS)
            sub_l = ''.join(omml_to_latex(c) for c in sub_node) if sub_node is not None else ''
            base_l = ''.join(omml_to_latex(c) for c in base) if base is not None else ''
            return base_l + '_{' + sub_l + '}'
        sup_node = node.find('m:sup', NS)
        sub_node = node.find('m:sub', NS)
        sup_l = ''.join(omml_to_latex(c) for c in sup_node) if sup_node is not None else ''
        sub_l = ''.join(omml_to_latex(c) for c in sub_node) if sub_node is not None else ''
        base_l = ''.join(omml_to_latex(c) for c in base) if base is not None else ''
        return base_l + '_{' + sub_l + '}^{' + sup_l + '}'

    if tag == 'rad':
        deg = node.find('m:deg', NS)
        rad = node.find('m:radicand', NS)
        deg_l = ''.join(omml_to_latex(c) for c in deg) if deg is not None else ''
        rad_l = ''.join(omml_to_latex(c) for c in rad) if rad is not None else ''
        if deg_l:
            return r'\sqrt[' + deg_l + ']{' + rad_l + '}'
        return r'\sqrt{' + rad_l + '}'

    if tag == 'nary':
        op = node.find('.//m:chr', NS) or node.find('.//m:op', NS)
        op_l = node_text_content(op) if op is not None else ''
        lower = ''.join(omml_to_latex(child) for child in node.findall('.//m:low', NS))
        upper = ''.join(omml_to_latex(child) for child in node.findall('.//m:up', NS))
        map_op = {'∑': r'\sum', '∫': r'\int', 'Π': r'\prod'}
        op_tex = map_op.get(op_l.strip(), op_l.strip())
        if lower and upper:
            return op_tex + '_{' + lower + '}^{' + upper + '}'
        if lower:
            return op_tex + '_{' + lower + '}'
        return op_tex

    if tag == 'acc':
        base = node.find('m:e', NS)
        chr_node = node.find('.//m:chr', NS)
        chr_text = node_text_content(chr_node) if chr_node is not None else ''
        inner = ''.join(omml_to_latex(c) for c in base) if base is not None else ''
        if 'bar' in chr_text.lower() or '¯' in chr_text:
            return r'\overline{' + inner + '}'
        if 'hat' in chr_text.lower():
            return r'\hat{' + inner + '}'
        return r'\overset{' + chr_text + '}{' + inner + '}'

    return ''.join(omml_to_latex(child) for child in node)


def extract_paragraph_content(node, processed_nodes=None):
    if processed_nodes is None:
        processed_nodes = set()
    if id(node) in processed_nodes:
        return []
    processed_nodes.add(id(node))

    results = []
    tag = strip_ns(node.tag)

    if tag in ('oMath', 'oMathPara'):
        latex = omml_to_latex(node).strip()
        if latex:
            results.append(('math', latex))
        return results

    if tag == 'r':
        # 处理数学公式节点
        math_nodes = node.findall('.//m:oMath', NS)
        if math_nodes:
            for math_node in math_nodes:
                latex = omml_to_latex(math_node).strip()
                if latex:
                    results.append(('math', latex))
        else:
            # 处理文本节点
            text = text_of_run(node)
            if text:
                va = get_run_vertical_align(node)
                if va == 'superscript':
                    results.append(('superscript', text))
                elif va == 'subscript':
                    results.append(('subscript', text))
                else:
                    results.append(('text', text))

            # 处理图片节点 <w:drawing>
            drawing = node.find('.//w:drawing', NS)
            if drawing is not None:
                inline = drawing.find('.//wp:inline', NS)
                if inline is not None:
                    docPr = inline.find('.//wp:docPr', NS)
                    img_name = docPr.get('name', 'Image') if docPr is not None else 'Image'
                    blip = inline.find('.//a:blip', NS)
                    if blip is not None:
                        embed = blip.get(f'{{{NS["r"]}}}embed')
                        if embed:
                            # 返回图片ID、名称和类型标记
                            results.append(('image', (embed, img_name)))
        return results

    if tag == 'tbl':
        return []

    for child in node:
        results.extend(extract_paragraph_content(child, processed_nodes))

    return results


def merge_superscripts_subscripts(content_items):
    merged = []
    i = 0
    while i < len(content_items):
        ttype, cont = content_items[i]
        if ttype == 'text':
            base = cont
            sup = ''
            sub = ''
            j = i + 1
            while j < len(content_items):
                nt, nc = content_items[j]
                if nt == 'superscript':
                    sup += nc
                    j += 1
                elif nt == 'subscript':
                    sub += nc
                    j += 1
                else:
                    break
            if sup or sub:
                expr = base
                if sub and sup:
                    expr = expr + '_{' + sub + '}^{' + sup + '}'
                elif sub:
                    expr = expr + '_{' + sub + '}'
                elif sup:
                    expr = expr + '^{' + sup + '}'
                merged.append(('math', expr))
                i = j
            else:
                merged.append((ttype, cont))
                i += 1
        else:
            if ttype == 'math':
                merged.append(('math', cont))
            elif ttype == 'superscript':
                merged.append(('math', '^{' + cont + '}'))
            elif ttype == 'subscript':
                merged.append(('math', '_{' + cont + '}'))
            else:
                merged.append((ttype, cont))
            i += 1
    return merged


def paragraph_items_to_text(content_items, join_with_br=False):
    parts = []
    for ttype, cont in content_items:
        if ttype == 'text':
            parts.append(cont)
        elif ttype == 'math':
            parts.append('$' + cont + '$')
        else:
            parts.append(cont)
    text = ''.join(parts)
    if join_with_br:
        # 保留段落间换行为 <br/> 时使用 caller 指定 True
        text = text.replace('\n', '<br/>').strip()
    else:
        text = text.replace('\n', ' ').strip()
    return text


def extract_cell_text(tc):
    texts = []
    # 仅遍历直接子 p（避免跨单元格抓取）
    for p in tc.findall('./w:p', NS):
        items = extract_paragraph_content(p)
        items = merge_superscripts_subscripts(items)
        txt = paragraph_items_to_text(items, join_with_br=True)
        if txt:
            texts.append(txt)
    return '<br/>'.join(texts).strip()


def table_to_html(tbl):
    rows = []
    tr_list = tbl.findall('./w:tr', NS)
    for tr in tr_list:
        row_cells = []
        for tc in tr.findall('./w:tc', NS):
            cell_html = extract_cell_text(tc)
            row_cells.append(cell_html)
        if row_cells:
            rows.append(row_cells)

    if not rows:
        return ''

    max_cols = max(len(r) for r in rows)
    for r in rows:
        if len(r) < max_cols:
            r.extend([''] * (max_cols - len(r)))

    # 判断是否第一行可作为表头
    first_row = rows[0]
    header_is_present = all(cell.strip() != '' for cell in first_row)

    html_lines = []
    html_lines.append('<table border="1">')
    if header_is_present:
        html_lines.append('  <thead>')
        html_lines.append('    <tr>')
        for cell in first_row:
            html_lines.append('      <th>{}</th>'.format(cell))
        html_lines.append('    </tr>')
        html_lines.append('  </thead>')
        body_rows = rows[1:]
    else:
        body_rows = rows

    html_lines.append('  <tbody>')
    for r in body_rows:
        html_lines.append('    <tr>')
        for cell in r:
            html_lines.append('      <td>{}</td>'.format(cell))
        html_lines.append('    </tr>')
    html_lines.append('  </tbody>')
    html_lines.append('</table>')

    # 将 HTML 表格作为独立块返回（前后带空行以便 Markdown / 文本可读）
    return '\n\n' + '\n'.join(html_lines) + '\n\n'


def paragraph_to_md(p, docx_path=None, output_dir=None):
    """
    将段落转换为Markdown格式
    :param p: 段落节点
    :param docx_path: .docx文件路径（可选，用于提取图片）
    :param output_dir: 输出目录路径（用于保存图片）
    :return: Markdown格式的段落文本
    """
    texts = []
    pPr = p.find('w:pPr', NS)
    global IMAGE_COUNT  # 声明使用全局计数器

    # 处理标题样式
    if pPr is not None:
        pStyle = pPr.find('w:pStyle', NS)
        if pStyle is not None:
            val = pStyle.get(f'{{{NS["w"]}}}val')
            if val and val.lower().startswith('heading'):
                try:
                    level = int(''.join(ch for ch in val if ch.isdigit()) or 1)
                except:
                    level = 1
                texts.append('#' * min(level, 6) + ' ')

    # 提取段落内容
    content_items = extract_paragraph_content(p)
    content_items = merge_superscripts_subscripts(content_items)

    # 判断内容类型
    has_text = any(item[0] == 'text' and item[1].strip() for item in content_items)
    has_math = any(item[0] == 'math' for item in content_items)
    only_math = has_math and not has_text

    # 处理每个内容项
    for content_type, content in content_items:
        if content_type == 'text':
            texts.append(content)
        elif content_type == 'math':
            if only_math:
                texts.append('\n\n' + '$$' + '\n' + content + '\n' + '$$' + '\n\n')
            else:
                texts.append(' ' + '$' + ' ' + content + ' ' + '$' + ' ')
        elif content_type == 'image' and docx_path and output_dir:
            # 处理图片
            embed_id, img_name = content
            doc_name = Path(docx_path).stem
            print(f"\n[图片处理] 开始处理图片: {img_name} (ID: {embed_id})")
            texts.append(f"![{img_name}](../Images/{doc_name}_images/image{IMAGE_COUNT}.png)")
            IMAGE_COUNT += 1

        else:
            if content_type == 'superscript':
                texts.append('^{' + content + '}')
            elif content_type == 'subscript':
                texts.append('_{' + content + '}')
            else:
                texts.append(str(content))

    # 处理列表项
    line = ''.join(texts)
    if pPr is not None and pPr.find('w:numPr', NS) is not None:
        line = '- ' + line

    return line


def convert_document(xml_path, docx_path):
    """
    转换XML文档为Markdown格式，包含详细的调试信息

    参数:
        xml_path: XML文件路径
        docx_path: 原始.docx文件路径（用于提取图片）

    返回:
        Markdown格式的文本
    """
    # 初始化输出目录
    output_dir = Path(xml_path).parent
    print(f"\n[文档转换] 开始转换文档")
    print(f"[调试] XML文件路径: {xml_path}")
    if docx_path:
        print(f"[调试] DOCX文件路径: {docx_path}")
    print(f"[调试] 输出目录: {output_dir}")

    # 解析XML
    try:
        print("[调试] 正在解析XML文档...")
        tree = ET.parse(xml_path)
        root = tree.getroot()
        body = root.find('.//w:body', NS) or root
        print("[调试] XML解析成功")
    except Exception as e:
        print(f"[错误] XML解析失败: {str(e)}")
        raise

    # 初始化统计信息
    node_stats = {
        'paragraphs': 0,
        'tables': 0,
        'images': 0,
        'math': 0,
        'other': 0
    }

    out_lines = []

    print("\n[调试] 开始处理文档节点...")
    for i, node in enumerate(body):
        tag = strip_ns(node.tag)
        if i % 10 == 0:  # 每处理10个节点打印一次进度
            print(f"[进度] 正在处理第 {i + 1} 个节点 (类型: {tag})")

        try:
            if tag == 'tbl':
                print(f"[调试] 处理表格节点 #{node_stats['tables'] + 1}")
                md_tbl = table_to_html(node)
                if md_tbl:
                    out_lines.append(md_tbl)
                    node_stats['tables'] += 1
                    print(f"[调试] 表格转换成功 (累计: {node_stats['tables']})")

            elif tag == 'p':
                print(f"[调试] 处理段落节点 #{node_stats['paragraphs'] + 1}")
                md = paragraph_to_md(
                    node,
                    docx_path=docx_path,
                    output_dir=output_dir
                )
                if md:
                    out_lines.append(md)
                    node_stats['paragraphs'] += 1

                    # 统计段落中的特殊内容
                    content_items = extract_paragraph_content(node)
                    for item_type, _ in content_items:
                        if item_type == 'math':
                            node_stats['math'] += 1
                        elif item_type == 'image':
                            node_stats['images'] += 1

                    print(f"[调试] 段落转换成功 (累计: {node_stats['paragraphs']})")
                    print(f"[调试] 段落内容: {md[:100]}...")  # 打印前100个字符

            else:
                node_stats['other'] += 1
                print(f"[调试] 处理其他类型节点: {tag}")
                for sub in node:
                    stag = strip_ns(sub.tag)
                    if stag == 'tbl':
                        print(f"[调试] 处理子表格节点")
                        md_tbl = table_to_html(sub)
                        if md_tbl:
                            out_lines.append(md_tbl)
                            node_stats['tables'] += 1
                    elif stag == 'p':
                        print(f"[调试] 处理子段落节点")
                        md = paragraph_to_md(
                            sub,
                            docx_path=docx_path,
                            output_dir=output_dir
                        )
                        if md:
                            out_lines.append(md)
                            node_stats['paragraphs'] += 1

        except Exception as e:
            print(f"[错误] 处理节点时出错 (类型: {tag}): {str(e)}")
            print(f"[调试] 节点内容: {ET.tostring(node, encoding='unicode')[:200]}...")  # 打印前200个字符
            continue  # 跳过错误节点继续处理

    # 打印转换统计信息
    print("\n[统计] 文档转换完成")
    print(f" - 段落数量: {node_stats['paragraphs']}")
    print(f" - 表格数量: {node_stats['tables']}")
    print(f" - 数学公式: {node_stats['math']}")
    print(f" - 图片数量: {node_stats['images']}")
    print(f" - 其他节点: {node_stats['other']}")
    print(f" - 总输出行数: {len(out_lines)}")

    return '\n\n'.join(out_lines)

def extract_docx_document_xml(docx_path, dst_dir):
    with zipfile.ZipFile(docx_path, 'r') as z:
        names = z.namelist()
        if 'word/document.xml' not in names:
            raise FileNotFoundError("docx does not contain word/document.xml")
        z.extract('word/document.xml', dst_dir)
        return Path(dst_dir) / 'word' / 'document.xml'


def convert_word_to_xml(docx_path):
    """
    将Word文档转换为XML并保存到XMLFile目录
    返回生成的XML文件路径
    """
    src = Path(docx_path)
    if not src.exists():
        raise FileNotFoundError(f"Input not found: {src}")

    # 生成输出路径
    xml_filename = src.stem + ".xml"
    xml_path = Path(XML_OUTPUT_DIR) / xml_filename

    print(f"\n[Word转XML] 开始转换: {src} -> {xml_path}")

    try:
        # 使用python-docx转换
        dump_document_xml_via_python_docx(docx_path, xml_path)
        print("[Word转XML] 转换成功")
        return xml_path
    except Exception as e:
        print(f"[Word转XML] 错误: {str(e)}")
        raise


def runCode(input_path):
    """
    处理输入文件：
    - 如果是.docx，先转换为XML
    - 如果是.xml，直接处理
    最终输出Markdown到MarkdownFile目录
    """
    src = Path(input_path)
    if not src.exists():
        raise FileNotFoundError(f"Input not found: {src}")

    # 1. 处理输入文件
    if src.suffix.lower() == '.docx':
        # 先转换为XML
        xml_path = convert_word_to_xml(src)
    else:
        # 直接使用XML文件
        xml_path = src

    # 2. 准备输出路径
    md_filename = xml_path.stem + ".md"
    md_path = Path(MD_OUTPUT_DIR) / md_filename

    # 3. 转换XML为Markdown
    print(f"\n[XML转Markdown] 开始转换: {xml_path} -> {md_path}")
    try:
        md = convert_document(xml_path, docx_path=src)
        print("convert_document 转换成功")
        md_path.write_text(md, encoding='utf-8')
        print("[XML转Markdown] 转换完成")
    except Exception as e:
        print(f"[XML转Markdown] 错误: {str(e)}")
        raise

    return md_path


if __name__ == "__main__":
    # 1. 提取并编号所有图片
    image_paths, total_images = extract_and_number_images(INPUT_PATH)
    print(f"已成功提取 {total_images} 张图片")
    try:
        print(f"\n开始处理文件: {INPUT_PATH}")
        result = runCode(INPUT_PATH)
        print(f"\n处理完成! Markdown文件已保存到: {result}")
    except Exception as e:
        print(f"\n处理过程中出错: {str(e)}")