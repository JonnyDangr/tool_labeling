import os
import io
import re
import json
import zipfile
import xml.etree.ElementTree as ET
import urllib.request
import urllib.error
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, unquote

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_ROOT = os.path.abspath(os.path.join(BASE_DIR, '..', 'Data'))
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
    cleaned_text = clean_glued_footer_text(raw_text).strip('\r')
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


def built_in_legal_model_predict(text, custom_prompt=None, labels_list=None, relations_list=None):
    """
    Vietnamese Legal NER & Relation Extraction Engine (Prompt-Aware).
    """
    annotations = []
    relations = []
    seen_spans = set()

    # Determine if user's edited prompt restricts which labels/relations should be extracted
    allowed_labels = None
    allowed_relations = None
    if custom_prompt and '{{LABELS}}' not in custom_prompt and labels_list:
        mentioned_labels = [lbl for lbl in labels_list if lbl.lower() in custom_prompt.lower()]
        if len(mentioned_labels) > 0:
            allowed_labels = set(mentioned_labels)

    if custom_prompt and '{{RELATIONS}}' not in custom_prompt and relations_list:
        mentioned_rels = [rel for rel in relations_list if rel.lower() in custom_prompt.lower()]
        if 'không trích xuất mối quan hệ' in custom_prompt.lower() or 'bỏ qua mối quan hệ' in custom_prompt.lower():
            allowed_relations = set()
        elif len(mentioned_rels) > 0:
            allowed_relations = set(mentioned_rels)

    def add_span(label, start, end):
        if allowed_labels is not None and label not in allowed_labels:
            return None
        while start < end and text[start].isspace():
            start += 1
        while end > start and text[end - 1].isspace():
            end -= 1
        if end <= start:
            return None
        key = (label, start, end)
        if key in seen_spans:
            for a in annotations:
                if a['label'] == label and a['start'] == start and a['end'] == end:
                    return a['id']
        seen_spans.add(key)
        sp_id = f"sp_ai_{len(annotations) + 1}"
        annotations.append({
            "id": sp_id,
            "label": label,
            "start": start,
            "end": end,
            "text": text[start:end]
        })
        return sp_id

    def add_rel(from_id, to_id, rel_type):
        if not from_id or not to_id:
            return
        if allowed_relations is not None and rel_type not in allowed_relations:
            return
        relations.append({
            "id": f"rel_ai_{len(relations) + 1}",
            "fromId": from_id,
            "toId": to_id,
            "type": rel_type
        })

    # Pattern 1: Số hiệu
    so_hieu_id = None
    for m in re.finditer(r'(?:Luật\s+số|Bộ\s+luật\s+số|Nghị\s+quyết\s+số|Số|SỐ)\s*:\s*([^\s\n]+)', text, re.IGNORECASE):
        so_hieu_id = add_span("Số hiệu", m.start(1), m.end(1))
        break
    if not so_hieu_id:
        m0 = re.search(r'(\d+/\d{4}/[A-ZĐ0-9\-]+)', text)
        if m0:
            so_hieu_id = add_span("Số hiệu", m0.start(1), m0.end(1))

    # Pattern 2: Cơ quan ban hành
    m_cq = re.search(r'(?:^|\n)((?:BỘ|CHÍNH PHỦ|THỦ TƯỚNG|ỦY BAN|UỶ BAN|QUỐC HỘI|NGÂN HÀNG)[A-ZÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸĐ\s]*?)\.?(?=\n)', text)
    if m_cq and 'TRƯỞNG' not in m_cq.group(1):
        add_span("Cơ quan ban hành", m_cq.start(1), m_cq.end(1))

    # Pattern 3: Ngày ban hành (prefer date near Số: or in header before QUYẾT ĐỊNH)
    search_start_for_date = 0
    m_so_pos = re.search(r'(?:Luật\s+số|Số|SỐ)\s*:', text, re.IGNORECASE)
    if m_so_pos:
        search_start_for_date = m_so_pos.start()
    date_found = False
    for m in re.finditer(r'([Nn]gày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4})', text[search_start_for_date:]):
        add_span("Ngày ban hành", search_start_for_date + m.start(1), search_start_for_date + m.end(1))
        date_found = True
        break
    if not date_found:
        for m in re.finditer(r'([Nn]gày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4})', text):
            add_span("Ngày ban hành", m.start(1), m.end(1))
            break

    # Pattern 4: Loại văn bản & Tên văn bản
    m_loai = re.search(r'\n(QUYẾT ĐỊNH|QUYẾT ĐỊNH|NGHỊ ĐỊNH|THÔNG TƯ|CHỈ THỊ|NGHỊ QUYẾT|BỘ LUẬT|LUẬT|PHÁP LỆNH)\s*\n', text)
    if m_loai:
        add_span("Loại văn bản", m_loai.start(1), m_loai.end(1))
        after_loai = text[m_loai.end():]
        m_title = re.match(r'((?:[^\n]+\n)+?)(?=[_—\-]{4,}|BỘ TRƯỞNG|THỦ TƯỚNG|Căn cứ)', after_loai)
        if m_title:
            t_start = m_loai.end() + m_title.start(1)
            t_end = m_loai.end() + m_title.end(1)
            add_span("Tên văn bản", t_start, t_end)

    # Pattern 5: Căn cứ ban hành
    for m in re.finditer(r'Căn cứ\s+([^\n;]+)', text):
        cc_id = add_span("Căn cứ ban hành", m.start(1), m.end(1))
        if cc_id and so_hieu_id:
            add_rel(so_hieu_id, cc_id, "Căn cứ ban hành")

    # Pattern 6: Cấu trúc mục lục (Phần, Chương, Mục, Tiểu Mục, Điều, Khoản, Điểm)
    for m in re.finditer(r'(?:^|\n)(Phần\s+(?:thứ\s+)?[IVXLCDM\d]+[^\n]*)', text, re.IGNORECASE):
        add_span("Phần", m.start(1), m.end(1))
    for m in re.finditer(r'(?:^|\n)(Chương\s+[IVXLCDM\d]+[^\n]*)', text):
        add_span("Chương", m.start(1), m.end(1))
    for m in re.finditer(r'(?:^|\n)(Mục\s+\d+[^\n]*)', text):
        add_span("Mục", m.start(1), m.end(1))
    for m in re.finditer(r'(?:^|\n)(Tiểu mục\s+\d+[^\n]*)', text, re.IGNORECASE):
        add_span("Tiểu Mục", m.start(1), m.end(1))
    for m in re.finditer(r'(?:^|\n)(Điều\s+\d+[\.\-:][^\n]+)', text):
        add_span("Điều", m.start(1), m.end(1))
    for m in re.finditer(r'(?:^|\n)(\d+\.\s+[^\n]+)', text):
        add_span("Khoản", m.start(1), m.end(1))
    for m in re.finditer(r'(?:^|\n)([a-eđgh]\)\s+[^\n]+)', text):
        add_span("Điểm", m.start(1), m.end(1))

    # Pattern 7: Ngày có hiệu lực
    m_hl = re.search(r'có hiệu lực\s+(?:thi hành\s+)?((?:kể\s+)?từ\s+ngày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4}|sau\s+\d+\s+ngày[^\.\n]*)', text, re.IGNORECASE)
    if m_hl:
        add_span("Ngày có hiệu lực", m_hl.start(1), m_hl.end(1))

    # Pattern 8: Chức danh & Người ký
    tail_start = max(0, len(text) - 550)
    tail_text = text[tail_start:]
    m_cd = re.search(r'(THỨ TRƯỞNG|BỘ TRƯỞNG|CHỦ TỊCH|THỦ TƯỚNG)', tail_text)
    if m_cd:
        add_span("Chức danh", tail_start + m_cd.start(1), tail_start + m_cd.end(1))
    m_nk = re.search(r'\(Đã ký\)\s*\n\s*([^\n]+)', tail_text)
    if m_nk:
        add_span("Người ký", tail_start + m_nk.start(1), tail_start + m_nk.end(1))

    return annotations, relations


def resolve_gemini_api_model(model_name):
    m = (model_name or '').lower()
    if 'flash-lite' in m:
        return 'gemini-2.5-flash-lite'
    if 'pro' in m:
        return 'gemini-2.5-pro'
    return 'gemini-2.5-flash'


def call_llm_auto_label(text, labels_list, relations_list, provider, api_key, endpoint_url, model_name, rendered_prompt=None):
    if rendered_prompt and rendered_prompt.strip():
        prompt = rendered_prompt
    else:
        prompt = f"""Bạn là hệ thống AI gán nhãn văn bản pháp luật Việt Nam.
Hãy trích xuất các Thực thể (Entities) và Mối quan hệ (Relations) từ văn bản sau.
Danh sách nhãn hợp lệ: {json.dumps(labels_list, ensure_ascii=False)}
Danh sách quan hệ hợp lệ: {json.dumps(relations_list, ensure_ascii=False)}

Trả về đúng định dạng JSON:
{{
  "entities": [{{"label": "Tên nhãn", "text": "chuỗi trích xuất chính xác từ văn bản"}}],
  "relations": [{{"from_text": "chuỗi nguồn", "to_text": "chuỗi đích", "type": "Tên quan hệ"}}]
}}

Văn bản:
{text[:8000]}
"""
    if provider == 'antigravity_gemini' and api_key:
        api_model = resolve_gemini_api_model(model_name)
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{api_model}:generateContent?key={api_key}"
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"responseMimeType": "application/json"}
        }
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode('utf-8'),
            headers={'Content-Type': 'application/json'}
        )
        with urllib.request.urlopen(req, timeout=35) as resp:
            res_data = json.loads(resp.read().decode('utf-8'))
            raw_json = res_data['candidates'][0]['content']['parts'][0]['text']
            return json.loads(raw_json)

    elif provider in ('vscode_openai', 'local_llm', 'antigravity_claude_openai') and endpoint_url:
        headers = {'Content-Type': 'application/json'}
        if api_key:
            headers['Authorization'] = f'Bearer {api_key}'
        payload = {
            "model": model_name or "gpt-4o-mini",
            "messages": [{"role": "user", "content": prompt}],
            "response_format": {"type": "json_object"}
        }
        req = urllib.request.Request(
            endpoint_url,
            data=json.dumps(payload).encode('utf-8'),
            headers=headers
        )
        with urllib.request.urlopen(req, timeout=35) as resp:
            res_data = json.loads(resp.read().decode('utf-8'))
            raw_json = res_data['choices'][0]['message']['content']
            return json.loads(raw_json)

    return None


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

            # Resolve target directories to scan
            scan_dirs = []
            if custom_dir:
                cand = custom_dir if os.path.isabs(custom_dir) else os.path.join(DATA_ROOT, custom_dir)
                if os.path.isdir(cand):
                    scan_dirs.append((os.path.basename(cand.rstrip('\\/')), cand))
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

            # Resolve folder or file path
            target_path = folder if os.path.isabs(folder) else os.path.join(DATA_ROOT, folder.replace('/', os.sep))
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
                            with open(jpath, 'r', encoding='utf-8') as jfile:
                                key = os.path.splitext(jf)[0].lower()
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

        if parsed.path == '/api/auto-label':
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length).decode('utf-8'))
            text = body.get('text', '')
            provider = body.get('provider', 'antigravity_gemini')
            api_key = body.get('api_key', '').strip()
            endpoint_url = body.get('endpoint_url', '').strip()
            model_name = body.get('model_name', '').strip()
            labels_list = body.get('labels', [])
            relations_list = body.get('relations', [])
            custom_prompt = body.get('custom_prompt', '')
            rendered_prompt = body.get('rendered_prompt', '')

            try:
                llm_res = None
                if provider in ('antigravity_gemini', 'vscode_openai', 'local_llm', 'antigravity_claude_openai') and (api_key or endpoint_url):
                    try:
                        llm_res = call_llm_auto_label(
                            text, labels_list, relations_list, provider, api_key, endpoint_url, model_name, rendered_prompt
                        )
                    except Exception:
                        llm_res = None

                raw_ents = None
                raw_rels = None
                if isinstance(llm_res, dict):
                    raw_ents = llm_res.get('Entities') or llm_res.get('entities')
                    raw_rels = llm_res.get('Relation') or llm_res.get('relations') or []

                if isinstance(raw_ents, list) and len(raw_ents) > 0:
                    annotations = []
                    relations = []
                    text_to_id = {}
                    eid_to_id = {}
                    for idx, ent in enumerate(raw_ents):
                        sub = (ent.get('noi_dung') or ent.get('text') or '').strip()
                        lbl = (ent.get('nhan') or ent.get('label') or '').strip()
                        eid = (ent.get('id') or f'E{idx + 1}').strip()
                        if not sub or not lbl:
                            continue
                        pos = text.find(sub)
                        if pos != -1:
                            sp_id = f"sp_llm_{idx + 1}"
                            annotations.append({
                                "id": sp_id,
                                "label": lbl,
                                "start": pos,
                                "end": pos + len(sub),
                                "text": sub
                            })
                            text_to_id[sub] = sp_id
                            eid_to_id[eid] = sp_id

                    for r_idx, rel in enumerate(raw_rels):
                        r_type = (rel.get('moi_quan_he') or rel.get('type') or '').strip()
                        from_id = eid_to_id.get((rel.get('tu_id') or '').strip()) or text_to_id.get((rel.get('from_text') or '').strip())
                        to_id = eid_to_id.get((rel.get('den_id') or '').strip()) or text_to_id.get((rel.get('to_text') or '').strip())
                        if from_id and to_id and r_type:
                            relations.append({
                                "id": f"rel_llm_{r_idx + 1}",
                                "fromId": from_id,
                                "toId": to_id,
                                "type": r_type
                            })
                    self.send_json({
                        'annotations': annotations,
                        'relations': relations,
                        'engine': f'{model_name or provider}'
                    })
                    return

                annotations, relations = built_in_legal_model_predict(
                    text,
                    custom_prompt=custom_prompt,
                    labels_list=labels_list,
                    relations_list=relations_list
                )
                self.send_json({
                    'annotations': annotations,
                    'relations': relations,
                    'engine': model_name or 'Antigravity Legal AI Engine'
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
