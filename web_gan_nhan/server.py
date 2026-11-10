import os
import io
import re
import json
import zipfile
import unicodedata
import xml.etree.ElementTree as ET
import urllib.request
import urllib.error
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, unquote

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_c1 = os.path.abspath(os.path.join(BASE_DIR, '..', 'Data'))
_c2 = os.path.abspath(os.path.join(BASE_DIR, '..', '..', 'Data'))
DATA_ROOT = _c1 if os.path.isdir(_c1) else (_c2 if os.path.isdir(_c2) else _c1)
W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def clean_glued_footer_text(txt):
    """Fix glued runs in legacy legal .docx files without losing any characters."""
    txt = re.sub(r'(\./\.)\s*(KT\.\s*BỘ TRƯỞNG|BỘ TRƯỞNG|THỨ TRƯỞNG|CHỦ TỊCH)', r'\1\n\2', txt)
    txt = re.sub(r'(KT\.\s*BỘ TRƯỞNG)\s+(THỨ TRƯỞNG)', r'\1\n\2', txt)
    txt = re.sub(r'(THỨ TRƯỞNG|BỘ TRƯỞNG|CHỦ TỊCH)\s*(\(Đã ký\))', r'\1\n\2', txt)
    txt = re.sub(r'(\(Đã ký\))\s*([^\n\s])', r'\1\n\2', txt)
    return txt


def parse_paragraph_xml(p_elem):
    """
    Extract 100% of text and formatting (align, bold, italic, style, indent) from a <w:p> element.
    Captures all <w:t>, <w:br/>, <w:cr/>, <w:tab/> inside runs, hyperlinks, smart tags, etc.
    """
    pPr = p_elem.find(f'{W_NS}pPr')
    jc = None
    p_style = None
    indent_left = 0
    if pPr is not None:
        jc_el = pPr.find(f'{W_NS}jc')
        if jc_el is not None:
            jc = jc_el.attrib.get(f'{W_NS}val')
        st_el = pPr.find(f'{W_NS}pStyle')
        if st_el is not None:
            p_style = st_el.attrib.get(f'{W_NS}val')
        ind_el = pPr.find(f'{W_NS}ind')
        if ind_el is not None:
            try:
                indent_left = int(ind_el.attrib.get(f'{W_NS}left', '0'))
            except ValueError:
                indent_left = 0

    parts = []
    bold_chars = 0
    italic_chars = 0
    total_chars = 0

    for elem in p_elem.iter():
        tag = elem.tag.replace(W_NS, '')
        if tag == 'r':
            rPr = elem.find(f'{W_NS}rPr')
            is_b = rPr is not None and (
                rPr.find(f'{W_NS}b') is not None and
                rPr.find(f'{W_NS}b').attrib.get(f'{W_NS}val', 'true') not in ('0', 'false')
            )
            is_i = rPr is not None and (
                rPr.find(f'{W_NS}i') is not None and
                rPr.find(f'{W_NS}i').attrib.get(f'{W_NS}val', 'true') not in ('0', 'false')
            )
            for child in elem:
                ctag = child.tag.replace(W_NS, '')
                if ctag == 't' and child.text:
                    parts.append(child.text)
                    c_len = len(child.text.strip())
                    total_chars += c_len
                    if is_b:
                        bold_chars += c_len
                    if is_i:
                        italic_chars += c_len
                elif ctag in ('br', 'cr'):
                    parts.append('\n')
                elif ctag == 'tab':
                    parts.append('\t')

    raw_text = ''.join(parts)
    cleaned_text = unicodedata.normalize('NFC', clean_glued_footer_text(raw_text).strip('\r'))
    if not cleaned_text.strip():
        return []

    is_bold = (bold_chars >= total_chars * 0.55) if total_chars > 0 else False
    is_italic = (italic_chars >= total_chars * 0.55) if total_chars > 0 else False

    align = 'left'
    if jc in ('center',):
        align = 'center'
    elif jc in ('right', 'end'):
        align = 'right'
    elif jc in ('both', 'distribute', 'justify'):
        align = 'justify'

    # Split if internal line breaks exist so each line gets proper alignment & offset
    sub_lines = [ln.strip() for ln in cleaned_text.split('\n') if ln.strip()]
    result_paras = []
    for ln in sub_lines:
        result_paras.append({
            'text': ln,
            'align': align,
            'bold': is_bold,
            'italic': is_italic,
            'style': p_style,
            'indent': indent_left
        })
    return result_paras


def extract_docx_structured(file_bytes):
    """
    Extract 100% of text and structural layout (paragraphs, alignments, bold/italic, tables, columns)
    from .docx bytes without losing a single character.
    Returns (full_text, blocks) where every paragraph in blocks has exact [start, end] offsets in full_text.
    """
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        xml_bytes = z.read('word/document.xml')
    root = ET.fromstring(xml_bytes)
    body = root.find(f'{W_NS}body')
    if body is None:
        return '', []

    raw_blocks = []
    for child in body:
        tag = child.tag.replace(W_NS, '')
        if tag == 'p':
            paras = parse_paragraph_xml(child)
            for p in paras:
                raw_blocks.append({'type': 'p', **p})
        elif tag == 'tbl':
            tblPr = child.find(f'{W_NS}tblPr')
            has_borders = False
            if tblPr is not None and tblPr.find(f'{W_NS}tblBorders') is not None:
                has_borders = True

            rows_data = []
            for tr in child.findall(f'{W_NS}tr'):
                row_cells = []
                for tc in tr.findall(f'{W_NS}tc'):
                    colspan = 1
                    tcPr = tc.find(f'{W_NS}tcPr')
                    if tcPr is not None:
                        gs = tcPr.find(f'{W_NS}gridSpan')
                        if gs is not None:
                            try:
                                colspan = int(gs.attrib.get(f'{W_NS}val', '1'))
                            except ValueError:
                                colspan = 1
                    cell_paras = []
                    for p_el in tc.findall(f'.//{W_NS}p'):
                        cell_paras.extend(parse_paragraph_xml(p_el))
                    row_cells.append({
                        'colspan': colspan,
                        'paras': cell_paras
                    })
                if any(len(c['paras']) > 0 for c in row_cells):
                    rows_data.append(row_cells)

            if len(rows_data) >= 3 or (len(rows_data) >= 1 and len(rows_data[0]) >= 3):
                has_borders = True

            if rows_data:
                raw_blocks.append({
                    'type': 'tbl',
                    'bordered': has_borders,
                    'rows': rows_data
                })

    # Apply intelligent Vietnamese Legal Document alignment heuristics where Word left jc=None
    in_title_zone = False
    for idx, blk in enumerate(raw_blocks):
        if blk['type'] == 'p':
            txt = blk['text'].strip()
            upper_txt = txt.upper()

            if re.match(r'^(QUYẾT ĐỊNH|QUYẾT ĐỊNH|NGHỊ ĐỊNH|THÔNG TƯ|CHỈ THỊ|NGHỊ QUYẾT|BỘ LUẬT|LUẬT|PHÁP LỆNH)$', upper_txt):
                blk['align'] = 'center'
                blk['bold'] = True
                in_title_zone = True
                continue

            if in_title_zone:
                if re.match(r'^[_—\-]{4,}$', txt):
                    blk['align'] = 'center'
                    in_title_zone = False
                    continue
                if txt == upper_txt and re.match(r'^(BỘ TRƯỞNG|THỦ TƯỚNG|CHỦ TỊCH|CHÍNH PHỦ|ỦY BAN|UỶ BAN|QUỐC HỘI)', upper_txt):
                    blk['align'] = 'center'
                    blk['bold'] = True
                    in_title_zone = False
                    continue
                if re.match(r'^(Căn\s+cứ|Theo\s+đề\s+nghị)', txt):
                    in_title_zone = False
                elif blk['bold'] or len(txt) < 120:
                    blk['align'] = 'center'
                    blk['bold'] = True
                    continue
                else:
                    in_title_zone = False

            if re.match(r'^(Điều\s+\d+|Khoản\s+\d+|\d+\.\s|[a-eđgh]\)\s|Căn\s+cứ\s|Theo\s+đề\s+nghị)', txt):
                blk['align'] = 'justify' if len(txt) > 85 else 'left'
            elif re.match(r'^[_—\-]{4,}$', txt):
                blk['align'] = 'center'
            elif txt == upper_txt and re.match(r'^(BỘ TRƯỞNG|THỦ TƯỚNG CHÍNH PHỦ|CHỦ TỊCH|HỘI ĐỒNG)', upper_txt) and len(txt) < 70:
                blk['align'] = 'center'
                blk['bold'] = True
            elif upper_txt in ('QUYẾT ĐỊNH:', 'NGHỊ ĐỊNH:', 'QUYẾT ĐỊNH:'):
                blk['align'] = 'center'
                blk['bold'] = True
            elif re.match(r'^(KT\.\s*BỘ TRƯỞNG|KT\.\s*THỦ TƯỚNG|KT\.\s*CHỦ TỊCH|THỨ TRƯỞNG|BỘ TRƯỞNG|CHỦ TỊCH|\(ĐÃ KÝ\))$', upper_txt):
                blk['align'] = 'right'
                blk['bold'] = True
            elif idx > 0 and raw_blocks[idx - 1].get('type') == 'p' and raw_blocks[idx - 1].get('text', '').strip() == '(Đã ký)':
                blk['align'] = 'right'
                blk['bold'] = True
            elif blk['align'] == 'left' and len(txt) > 95 and blk.get('style') != 'Heading1':
                blk['align'] = 'justify'

        elif blk['type'] == 'tbl' and not blk['bordered']:
            # Format 2-column header / footer tables cleanly like official legal documents
            for r in blk['rows']:
                if len(r) == 2:
                    for p in r[0]['paras']:
                        if not p['text'].startswith('-') and not p['text'].startswith('Nơi nhận'):
                            p['align'] = 'center'
                    for p in r[1]['paras']:
                        p['align'] = 'center'

    # Assign exact character offsets [start, end] in full_text for every paragraph
    all_lines = []
    current_offset = 0

    for blk in raw_blocks:
        if blk['type'] == 'p':
            txt = blk['text']
            blk['start'] = current_offset
            blk['end'] = current_offset + len(txt)
            all_lines.append(txt)
            current_offset = blk['end'] + 1
        elif blk['type'] == 'tbl':
            for r in blk['rows']:
                for cell in r:
                    for p in cell['paras']:
                        txt = p['text']
                        p['start'] = current_offset
                        p['end'] = current_offset + len(txt)
                        all_lines.append(txt)
                        current_offset = p['end'] + 1

    full_text = '\n'.join(all_lines)
    return full_text, raw_blocks


class LabelingRequestHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
        super().end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == '/api/folders':
            folders = []
            if os.path.isdir(DATA_ROOT):
                for entry in sorted(os.listdir(DATA_ROOT)):
                    full_p = os.path.join(DATA_ROOT, entry)
                    if os.path.isdir(full_p):
                        folders.append({
                            'name': entry,
                            'path': full_p
                        })
            self.send_json({'data_root': DATA_ROOT, 'folders': folders})
            return

        if parsed.path == '/api/samples':
            qs = parse_qs(parsed.query)
            search_q = qs.get('q', [''])[0].strip().lower()
            custom_dir = qs.get('dir', [''])[0].strip()
            limit = int(qs.get('limit', ['300'])[0])
            samples = []
            total_found = 0

            # Resolve target directories to scan securely under DATA_ROOT
            scan_dirs = []
            if custom_dir:
                cand = custom_dir if os.path.isabs(custom_dir) else os.path.join(DATA_ROOT, custom_dir)
                norm_cand = os.path.normpath(cand)
                if norm_cand == DATA_ROOT or norm_cand.startswith(DATA_ROOT + os.sep):
                    if os.path.isdir(norm_cand):
                        scan_dirs.append((os.path.basename(norm_cand.rstrip('\\/')), norm_cand))
            else:
                if os.path.isdir(DATA_ROOT):
                    for dname in sorted(os.listdir(DATA_ROOT)):
                        dp = os.path.join(DATA_ROOT, dname)
                        if os.path.isdir(dp):
                            scan_dirs.append((dname, dp))

            for cat_name, cat_dir in scan_dirs:
                try:
                    entries = sorted(os.listdir(cat_dir))
                except Exception:
                    continue

                # Case A: cat_dir directly contains .docx files
                direct_docx = [e for e in entries if e.lower().endswith('.docx') and not e.startswith('~$')]
                if direct_docx:
                    for docx_fn in direct_docx:
                        if search_q and search_q not in docx_fn.lower():
                            continue
                        total_found += 1
                        if len(samples) < limit:
                            samples.append({
                                'id': os.path.join(cat_dir, docx_fn),
                                'name': docx_fn,
                                'category': cat_name,
                                'has_json': any(e.lower().endswith('.json') for e in entries),
                                'json_files': [e for e in entries if e.lower().endswith('.json')]
                            })

                # Case B: cat_dir contains subfolders (each subfolder has noi_dung.docx + thuoc_tinh.json / luoc_do.json / muc_luc.json)
                for sub_name in entries:
                    sub_path = os.path.join(cat_dir, sub_name)
                    if not os.path.isdir(sub_path):
                        continue
                    if search_q and search_q not in sub_name.lower():
                        continue
                    try:
                        sub_files = os.listdir(sub_path)
                    except Exception:
                        continue
                    docx_candidates = [f for f in sub_files if f.lower().endswith('.docx') and not f.startswith('~$')]
                    if docx_candidates:
                        total_found += 1
                        if len(samples) < limit:
                            json_files = [f for f in sorted(sub_files) if f.lower().endswith('.json')]
                            samples.append({
                                'id': f"{cat_name}/{sub_name}" if not custom_dir else sub_path,
                                'name': sub_name,
                                'category': cat_name,
                                'has_json': len(json_files) > 0,
                                'json_files': json_files
                            })

            self.send_json({
                'samples': samples,
                'total': total_found,
                'active_dir': custom_dir or DATA_ROOT
            })
            return

        if parsed.path == '/api/load-sample':
            qs = parse_qs(parsed.query)
            folder = qs.get('folder', [''])[0].strip()
            if not folder:
                self.send_json({'error': 'Thiếu tham số folder'}, status=400)
                return

            # Resolve folder or file path securely under DATA_ROOT
            if os.path.isabs(folder):
                norm_cand = os.path.normpath(folder)
                if not (norm_cand == DATA_ROOT or norm_cand.startswith(DATA_ROOT + os.sep)):
                    self.send_json({'error': 'Đường dẫn không hợp lệ hoặc nằm ngoài thư mục Data'}, status=403)
                    return
                target_path = norm_cand
            else:
                target_path = os.path.normpath(os.path.join(DATA_ROOT, folder.replace('/', os.sep)))
                if not (target_path == DATA_ROOT or target_path.startswith(DATA_ROOT + os.sep)):
                    self.send_json({'error': 'Đường dẫn không hợp lệ hoặc nằm ngoài thư mục Data'}, status=403)
                    return

            if not os.path.exists(target_path):
                # Fallback: search inside subdirectories of DATA_ROOT
                base_f = os.path.basename(folder)
                if os.path.isdir(DATA_ROOT):
                    for dname in os.listdir(DATA_ROOT):
                        cand = os.path.join(DATA_ROOT, dname, base_f)
                        if os.path.exists(cand):
                            target_path = cand
                            break

            docx_path = None
            folder_dir = None
            display_name = os.path.basename(target_path)

            if os.path.isfile(target_path) and target_path.lower().endswith('.docx'):
                docx_path = target_path
                folder_dir = os.path.dirname(target_path)
            elif os.path.isdir(target_path):
                folder_dir = target_path
                cand_noi_dung = os.path.join(target_path, 'noi_dung.docx')
                if os.path.isfile(cand_noi_dung):
                    docx_path = cand_noi_dung
                else:
                    for f in os.listdir(target_path):
                        if f.lower().endswith('.docx') and not f.startswith('~$'):
                            docx_path = os.path.join(target_path, f)
                            break
                if not display_name.lower().endswith('.docx'):
                    display_name = f"{display_name}.docx"

            if not docx_path or not os.path.isfile(docx_path):
                self.send_json({'error': f'Không tìm thấy file .docx trong: {folder}'}, status=404)
                return

            with open(docx_path, 'rb') as f:
                full_text, blocks = extract_docx_structured(f.read())

            # Automatically read companion JSON files (thuoc_tinh.json, luoc_do.json, muc_luc.json) if present
            companion_json = {}
            if folder_dir and os.path.isdir(folder_dir):
                for jf in os.listdir(folder_dir):
                    if jf.lower().endswith('.json'):
                        jpath = os.path.join(folder_dir, jf)
                        try:
                            with open(jpath, 'r', encoding='utf-8-sig') as jfile:
                                key = os.path.splitext(jf)[0].lower().strip()
                                companion_json[key] = json.load(jfile)
                        except Exception:
                            pass

            self.send_json({
                'file_name': display_name,
                'folder': folder,
                'text': full_text,
                'blocks': blocks,
                'companion_json': companion_json
            })
            return

        return super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == '/api/parse-docx':
            length = int(self.headers.get('Content-Length', 0))
            file_bytes = self.rfile.read(length)
            raw_name = self.headers.get('X-File-Name', 'document.docx')
            filename = unquote(raw_name)
            try:
                full_text, blocks = extract_docx_structured(file_bytes)
                self.send_json({
                    'file_name': filename,
                    'text': full_text,
                    'blocks': blocks
                })
            except Exception as e:
                self.send_json({'error': str(e)}, status=400)
            return

        self.send_json({'error': 'Not found'}, status=404)

    def send_json(self, data, status=200):
        payload = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)


if __name__ == '__main__':
    port = 8000
    print("============================================================")
    print("  WEB GAN NHAN VAN BAN PHAP LUAT DANG CHAY TAI:")
    print(f"  http://localhost:{port}")
    print("============================================================")
    server = HTTPServer(('0.0.0.0', port), LabelingRequestHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nDa dung server.")
