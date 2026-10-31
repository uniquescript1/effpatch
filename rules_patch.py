"""
RK Word Formatting Lab - Dynamic Rules & Evaluation Patch File
फाइल रीडिंग और सभी प्रश्न जाँचने के नियम इस फाइल में हैं।
"""
import time, re
from pathlib import Path

try:
    import docx
    from docx import Document
except ImportError:
    Document = None

def get_docx_format(path):
    if Document is None or not path or not Path(path).exists():
        return {
            "paragraphs": [], "tables": [], "sections": [],
            "has_header": False, "has_footer": False,
            "header_text": "", "footer_text": "",
            "has_page_border": False, "has_watermark": False, 
            "full_text": "", "all_runs": [], "table_structures": [],
            "orientation": "PORTRAIT", "margins": {"top": 0, "bottom": 0, "left": 0, "right": 0}
        }
    
    time.sleep(0.2)
    doc = Document(str(path))
    
    header_tokens = set()
    footer_tokens = set()
    
    for s in doc.sections:
        if s.header:
            try:
                for node in s.header._element.iter():
                    if node.text and node.text.strip():
                        header_tokens.add(node.text.strip())
            except Exception:
                pass
        if s.footer:
            try:
                for node in s.footer._element.iter():
                    if node.text and node.text.strip():
                        footer_tokens.add(node.text.strip())
            except Exception:
                pass

    header_text = " ".join(header_tokens).strip()
    footer_text = " ".join(footer_tokens).strip()
    
    xml_data = str(doc._element.xml)
    has_page_border = "pgBorders" in xml_data
    has_watermark = False
    all_xml_to_check = [xml_data]
    for s in doc.sections:
        if s.header and s.header._element is not None:
            all_xml_to_check.append(str(s.header._element.xml))
        if s.footer and s.footer._element is not None:
            all_xml_to_check.append(str(s.footer._element.xml))
            
    watermark_patterns = [
        r'PowerPlusWaterMarkObject', 
        r'WordPictureWatermark', 
        r'w:watermark', 
        r'watermark',
        r'v:shape',          
        r'w:drawing',        
        r'w:pict'            
    ]
    
    for xml in all_xml_to_check:
        if any(re.search(pattern, xml, re.I) for pattern in watermark_patterns):
            has_watermark = True
            break

    orientation = "PORTRAIT"
    margins = {"top": 0, "bottom": 0, "left": 0, "right": 0}
    for s in doc.sections:
        if s.orientation:
            orientation = str(s.orientation).upper()
        if s.top_margin: margins["top"] = round(s.top_margin.pt, 2)
        if s.bottom_margin: margins["bottom"] = round(s.bottom_margin.pt, 2)
        if s.left_margin: margins["left"] = round(s.left_margin.pt, 2)
        if s.right_margin: margins["right"] = round(s.right_margin.pt, 2)

    table_structures = []
    for t in doc.tables:
        t_struct = []
        for row in t.rows:
            r_struct = []
            for cell in row.cells:
                cell_xml = str(cell._tc.xml)
                cell_has_bullet = 'w:numPr' in cell_xml or 'w:bullet' in cell_xml or 'w:ilvl' in cell_xml
                r_struct.append({
                    "grid_span": "gridSpan" in cell_xml,
                    "v_merge": "vMerge" in cell_xml,
                    "text": cell.text.strip(),
                    "has_bullet": cell_has_bullet
                })
            t_struct.append(r_struct)
        table_structures.append(t_struct)

    paragraphs_data = []
    all_runs_summary = []
    for p in doc.paragraphs:
        runs_data = []
        for r in p.runs:
            sz = round(r.font.size.pt, 2) if r.font and r.font.size and r.font.size.pt is not None else 0.0
            
            # सुधार: बोल्ड (Bold) डिटेक्ट करने के लिए मजबूत और उन्नत तरीका (प्रॉपर्टी और XML दोनों चेक होंगे)[cite: 11]
            is_bold = bool(r.bold)
            try:
                r_xml = r._r.xml
                if '<w:b/>' in r_xml or '<w:b ' in r_xml:
                    if 'w:val="false"' not in r_xml and 'w:val="0"' not in r_xml:
                        is_bold = True
            except Exception:
                pass
            # फॉन्ट साइज निकालने का कोड
            sz = 0.0
            try:
                if r.font and r.font.size and r.font.size.pt is not None:
                    sz = round(r.font.size.pt, 2)
                else:
                    # XML से फॉन्ट साइज (sz टैग) निकालने का उन्नत तरीका
                    r_xml = r._r.xml
                    m_sz = re.search(r'w:sz\s+w:val="([^"]+)"', r_xml)
                    if m_sz:
                        sz = round(float(m_sz.group(1)) / 2.0, 2)
            except Exception:
                pass
            # सुपरस्क्रिप्ट (w:vertAlign w:val="superscript") डिटेक्ट करना
            superscript_val = False
            try:
                r_xml = r._r.xml
                if 'w:vertAlign' in r_xml and 'superscript' in r_xml:
                    superscript_val = True
                elif r.font and hasattr(r.font, 'superscript') and r.font.superscript:
                    superscript_val = True
            except Exception:
                pass    
            # हाईलाइट (Highlight) डिटेक्ट करना[cite: 11]
            highlight_val = ""
            try:
                r_xml = r._r.xml
                m_hl = re.search(r'w:highlight\s+w:val="([^"]+)"', r_xml)
                if m_hl:
                    highlight_val = m_hl.group(1).lower()
            except Exception:
                pass

            font_name_val = ""
            try:
                r_xml = r._r.xml
                m_font = re.search(r'w:rFonts\s+[^>]*w:([a-zA-Z]+)="([^"]+)"', r_xml)
                if m_font:
                    font_name_val = m_font.group(2).lower()
                elif r.font and r.font.name:
                    font_name_val = r.font.name.lower()
            except Exception:
                if r.font and r.font.name:
                    font_name_val = r.font.name.lower()

            font_color_val = ""
            try:
                r_xml = r._r.xml
                m_color = re.search(r'w:color\s+w:val="([^"]+)"', r_xml)
                if m_color:
                    font_color_val = m_color.group(1).lower()
                elif r.font and r.font.color and r.font.color.rgb:
                    font_color_val = str(r.font.color.rgb).lower()
            except Exception:
                pass

            dstrike_val = False
            try:
                r_xml = r._r.xml
                if 'w:dstrike' in r_xml and 'w:val="false"' not in r_xml:
                    dstrike_val = True
            except Exception:
                pass
            # बुलेट या लिस्ट (Bullet / Numbering) डिटेक्ट करना
            has_bullet = False
            try:
                p_xml = p._p.xml
                if 'w:numPr' in p_xml or 'w:bullet' in p_xml or 'w:ilvl' in p_xml:
                    has_bullet = True
            except Exception:
                pass
                # पैराग्राफ राइट बॉर्डर (Paragraph Right Border) डिटेक्ट करना
            # पैराग्राफ राइट बॉर्डर (Paragraph Right Border) डिटेक्ट करना
            right_border_val = False
            try:
                p_xml = p._p.xml
                if 'w:pBdr' in p_xml and 'w:right' in p_xml:
                    if 'w:val="none"' not in p_xml:
                        right_border_val = True
            except Exception:
                pass
            superscript_val = False
            try:
                r_xml = r._r.xml
                if 'w:vertAlign' in r_xml and 'superscript' in r_xml:
                    superscript_val = True
                elif r.font and hasattr(r.font, 'superscript') and r.font.superscript:
                    superscript_val = True
            except Exception:
                pass

            subscript_val = False
            try:
                r_xml = r._r.xml
                if 'w:vertAlign' in r_xml and 'subscript' in r_xml:
                    subscript_val = True
                elif r.font and hasattr(r.font, 'subscript') and r.font.subscript:
                    subscript_val = True
            except Exception:
                pass

            font_scale_val = "100"
            font_spacing_val = "normal"
            font_position_val = "normal"
            try:
                r_xml = r._r.xml
                m_scale = re.search(r'w:w\s+w:val="([^"]+)"', r_xml)
                if m_scale:
                    font_scale_val = m_scale.group(1)
                m_space = re.search(r'w:spacing\s+w:val="([^"]+)"', r_xml)
                if m_space:
                    font_spacing_val = m_space.group(1)
                m_pos = re.search(r'w:position\s+w:val="([^"]+)"', r_xml)
                if m_pos:
                    font_position_val = m_pos.group(1)
            except Exception:
                pass

            r_info = {
                "text": r.text,
                "bold": is_bold,
                "italic": bool(r.italic),
                "underline": bool(r.underline),
                "strike": bool(r.font.strike) if r.font else False,
                "double_strike": dstrike_val,
                "superscript": superscript_val,
                "subscript": subscript_val,
                "font_size": sz,
                "font_name": font_name_val,
                "highlight": highlight_val,
                "font_color": font_color_val,
                "font_scale": font_scale_val,        
                "font_spacing": font_spacing_val,    
                "font_position": font_position_val,  
            }
            runs_data.append(r_info)
            if r.text:
                all_runs_summary.append(r_info)
        
        left_indent_val = 0.0
        right_indent_val = 0.0
        first_line_indent_val = 0.0
        hanging_indent_val = 0.0
        space_before_val = 0.0
        space_after_val = 0.0
        line_spacing_val = 0.0
        
        try:
            if p.paragraph_format.left_indent and p.paragraph_format.left_indent.pt:
                left_indent_val = round(p.paragraph_format.left_indent.pt, 2)
            if p.paragraph_format.right_indent and p.paragraph_format.right_indent.pt:
                right_indent_val = round(p.paragraph_format.right_indent.pt, 2)
            if p.paragraph_format.first_line_indent and p.paragraph_format.first_line_indent.pt:
                val = p.paragraph_format.first_line_indent.pt
                if val > 0:
                    first_line_indent_val = round(val, 2)
                else:
                    hanging_indent_val = round(abs(val), 2)
            if p.paragraph_format.space_before and p.paragraph_format.space_before.pt:
                space_before_val = round(p.paragraph_format.space_before.pt, 2)
            if p.paragraph_format.space_after and p.paragraph_format.space_after.pt:
                space_after_val = round(p.paragraph_format.space_after.pt, 2)
            if p.paragraph_format.line_spacing is not None:
                line_spacing_val = float(p.paragraph_format.line_spacing)
        except Exception:
            pass

        align = p.alignment
        align_str = "LEFT"
        if align == docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER: align_str = "CENTER"
        elif align == docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT: align_str = "RIGHT"
        elif align == docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY: align_str = "JUSTIFY"
        elif str(align).upper() == "CENTER": align_str = "CENTER"
        elif str(align).upper() == "RIGHT": align_str = "RIGHT"
        elif str(align).upper() == "JUSTIFY": align_str = "JUSTIFY"

        last_word_info = {"word": "", "bold": False, "italic": False, "underline": False, "highlight": "", "strike": False}
        if p.text.strip():
            words_in_p = p.text.strip().split()
            if words_in_p:
                target_last_word = words_in_p[-1]
                last_word_info["word"] = target_last_word
                for r in reversed(p.runs):
                    if target_last_word in r.text or r.text.strip().endswith(target_last_word):
                        last_word_info["bold"] = bool(r.bold)
                        last_word_info["italic"] = bool(r.italic)
                        last_word_info["underline"] = bool(r.underline)
                        last_word_info["strike"] = bool(r.font.strike) if r.font else False
                        try:
                            r_xml = r._r.xml
                            m_hl = re.search(r'w:highlight\s+w:val="([^"]+)"', r_xml)
                            if m_hl:
                                last_word_info["highlight"] = m_hl.group(1).lower()
                        except Exception:
                            pass
                        break

        p_txt = p.text.strip()
        case_type = "normal"
        if p_txt:
            if p_txt == p_txt.capitalize():
                case_type = "sentence"
            elif p_txt == p_txt.lower():
                case_type = "lowercase"
            elif p_txt == p_txt.upper():
                case_type = "uppercase"
            elif p_txt == p_txt.title():
                case_type = "title"

        paragraphs_data.append({
            "text": p.text.strip(),
            "alignment": align_str,
            "runs": runs_data,
            "case_type": case_type,
            "left_indent": left_indent_val,
            "right_indent": right_indent_val,
            "first_line_indent": first_line_indent_val,
            "hanging_indent": hanging_indent_val,
            "space_before": space_before_val,
            "space_after": space_after_val,
            "line_spacing": line_spacing_val,
            "last_word": last_word_info,
            "has_bullet": has_bullet,
            "right_border": right_border_val
        })

    return {
        "paragraphs": paragraphs_data,
        "has_header": len(header_text.strip()) > 0,
        "has_footer": len(footer_text.strip()) > 0,
        "header_text": header_text,
        "footer_text": footer_text,
        "has_page_border": has_page_border,
        "has_watermark": has_watermark,
        "full_text": "".join([p.text for p in doc.paragraphs]),
        "all_runs": all_runs_summary,
        "table_structures": table_structures,
        "orientation": orientation,
        "margins": margins
    }


def evaluate_two_file_exact_match(practical_path, solved_path, user_path, questions, per_question, question_ticks):
    if Document is None:
        return 0, len(questions) * per_question, [(f"Q{i+1}", q, "Incorrect", "0", "python-docx missing") for i, q in enumerate(questions)]

    if not solved_path or not Path(solved_path).exists():
        solved_path = practical_path

    solved_fmt = get_docx_format(solved_path)
    actual_fmt = get_docx_format(user_path)

    max_score = len(questions) * per_question
    earned_score = 0
    details = []

    for i, q in enumerate(questions, 1):
        q_lower = q.lower()
        q_passed = False
        remark = "Incorrect or not matched with solved file"
        
        is_ticked = question_ticks.get(i, False)
        if not is_ticked:
            details.append((f"Q{i}", q, "Unticked", "0", "Question not checked/ticked"))
            continue
# पैराग्राफ के अंत में शब्द जोड़ने या लिखने वाले प्रश्न का सटीक नियम
        elif ("paragraph" in q_lower or "पैराग्राफ" in q_lower) and ("अंत" in q_lower or "end" in q_lower):
            p_match = re.search(r'(?:paragraph|पैराग्राफ)\s*(\d+)', q_lower)
            quotes_words = re.findall(r'["\']([^"\']+)["\']', q)
            
            p_idx = int(p_match.group(1)) - 1 if p_match else 3
            target_word = quotes_words[0].strip() if quotes_words else ""
            
            if p_idx < len(actual_fmt["paragraphs"]) and p_idx < len(solved_fmt["paragraphs"]):
                actual_p_text = actual_fmt["paragraphs"][p_idx]["text"].strip()
                solved_p_text = solved_fmt["paragraphs"][p_idx]["text"].strip()
                
                # जाँच करें कि क्या यूजर वाले पैराग्राफ के अंत में वह शब्द मौजूद है या साल्व्ड फाइल से मैच होता है
                if (target_word and actual_p_text.endswith(target_word)) or (actual_p_text == solved_p_text):
                    q_passed = True
                    remark = f"Paragraph {p_idx+1} end text 100% matched with solved file"
                else:
                    remark = f"Paragraph {p_idx+1} does not end with the required word '{target_word}'"
            else:
                if actual_fmt["full_text"].strip() == solved_fmt["full_text"].strip():
                    q_passed = True
                    remark = "Paragraph end text matched"
                else:
                    remark = "Paragraph index or text mismatch"
        # 1. Bold प्रश्न जाँचने का नियम (सटीक मिलान)[cite: 11]
        elif "bold" in q_lower or "बोल्ड" in q_lower:
            solved_has_bold = any(any(r["bold"] for r in p["runs"]) for p in solved_fmt["paragraphs"])
            actual_has_bold = any(any(r["bold"] for r in p["runs"]) for p in actual_fmt["paragraphs"])
            
            if solved_has_bold and actual_has_bold:
                match_bold = True
                # Para-by-para aur run-by-run strict comparison
                for sp, ap in zip(solved_fmt["paragraphs"], actual_fmt["paragraphs"]):
                    for sr, ar in zip(sp["runs"], ap["runs"]):
                        if sr["bold"] != ar["bold"]:
                            match_bold = False
                            break
                    if not match_bold:
                        break
                
                if match_bold:
                    q_passed = True
                    remark = "Bold formatting 100% matched with solved file"
                else:
                    remark = "Bold formatting mismatch with solved file (Kewal wahi line/paragraph bold hona chahiye jo solved file mein hai)"
            elif not solved_has_bold and not actual_has_bold:
                q_passed = True
                remark = "Bold condition matched"
            else:
                remark = "Bold formatting presence mismatch with solved file"
                
            # बुलेट लिस्ट (Bullet List / Bullets) जाँचने का सख्त नियम
        # बुलेट लिस्ट (Bullet List) जाँचने का 100% सटीक (Strict) नियम
        # बुलेट लिस्ट (Bullet List) जाँचने का 100% सटीक (Strict) नियम
        elif "bullet" in q_lower or "बुलेट" in q_lower or "list" in q_lower or "सूची" in q_lower:
            # 1. सामान्य पैराग्राफ बुलेट की सटीक तुलना
            solved_bullets = [p.get("has_bullet", False) for p in solved_fmt["paragraphs"]]
            actual_bullets = [p.get("has_bullet", False) for p in actual_fmt["paragraphs"]]
            
            paragraph_match = (solved_bullets == actual_bullets)
            
            # 2. टेबल के अंदर सेल्स की बुलेट की सटीक तुलना
            table_match = True
            if len(solved_fmt["table_structures"]) == len(actual_fmt["table_structures"]):
                for st, at in zip(solved_fmt["table_structures"], actual_fmt["table_structures"]):
                    for s_row, a_row in zip(st, at):
                        for s_cell, a_cell in zip(s_row, a_row):
                            if s_cell.get("has_bullet", False) != a_cell.get("has_bullet", False):
                                table_match = False
                                break
                        if not table_match:
                            break
                    if not table_match:
                        break
            else:
                if len(solved_fmt["table_structures"]) > 0 or len(actual_fmt["table_structures"]) > 0:
                    table_match = False

            # यदि पैराग्राफ और टेबल दोनों की बुलेट पोजीशन 100% मैच होती है तभी सही मानेंगे
            if paragraph_match and table_match:
                q_passed = True
                remark = "Bullet list 100% matched with solved file"
            else:
                remark = "Bullet list mismatch (बुलेट सही कॉलम, सेल या पैराग्राफ पर नहीं है)"    
                
                # 3. फॉन्ट साइज (Font Size) जाँचने का सख्त नियम
        elif "font size" in q_lower or "size" in q_lower or "फॉन्ट साइज" in q_lower or "कार आकार" in q_lower or "आकार" in q_lower:
            solved_has_size = any(any(r.get("font_size", 0) > 0 for r in p["runs"]) for p in solved_fmt["paragraphs"])
            actual_has_size = any(any(r.get("font_size", 0) > 0 for r in p["runs"]) for p in actual_fmt["paragraphs"])
            
            if solved_has_size and actual_has_size:
                match_size = True
                # पैरा-दर-पैरा और रन-दर-रन सटीक फॉन्ट साइज तुलना
                for sp, ap in zip(solved_fmt["paragraphs"], actual_fmt["paragraphs"]):
                    for sr, ar in zip(sp["runs"], ap["runs"]):
                        if sr.get("font_size", 0) > 0 and sr.get("font_size") != ar.get("font_size"):
                            match_size = False
                            break
                    if not match_size:
                        break
                
                if match_size:
                    q_passed = True
                    remark = "Font size 100% matched with solved file"
                else:
                    remark = "Font size mismatch with solved file (सही पैराग्राफ या लाइन का फॉन्ट साइज मेल नहीं खाता)"
            elif not solved_has_size and not actual_has_size:
                q_passed = True
                remark = "Font size condition matched"
            else:
                remark = "Font size presence mismatch with solved file"

        # 2. Highlight प्रश्न जाँचने का नियम[cite: 11]
        elif "highlight" in q_lower or "हाईलाइट" in q_lower or "पीला" in q_lower or "color" in q_lower:
            actual_has_hl = any(any(r["highlight"] != "" for r in p["runs"]) for p in actual_fmt["paragraphs"])
            solved_has_hl = any(any(r["highlight"] != "" for r in p["runs"]) for p in solved_fmt["paragraphs"])
            
            if actual_has_hl:
                if solved_has_hl:
                    match_hl = True
                    for sp, ap in zip(solved_fmt["paragraphs"], actual_fmt["paragraphs"]):
                        for sr, ar in zip(sp["runs"], ap["runs"]):
                            if sr["highlight"] and sr["highlight"] != ar["highlight"]:
                                match_hl = False
                    if match_hl:
                        q_passed = True
                        remark = "Highlight 100% matched with solved file"
                    else:
                        q_passed = True
                        remark = "Highlight applied successfully by user"
                else:
                    q_passed = True
                    remark = "Highlight applied successfully by user"
            else:
                remark = "Highlight missing (Question requires highlight)"

        elif "merge" in q_lower or "मर्ज" in q_lower or "cell" in q_lower:
            if actual_fmt["table_structures"] == solved_fmt["table_structures"] and len(solved_fmt["table_structures"]) > 0:
                q_passed = True
                remark = "Cell merge 100% matched with solved master file"
            else:
                remark = "Cell merge does not match solved file"
                
        elif "uppercase" in q_lower or "अपर केस" in q_lower or "बड़े अक्षरों" in q_lower:
            solved_uc = [p.get("case_type") == "uppercase" for p in solved_fmt["paragraphs"]]
            actual_uc = [p.get("case_type") == "uppercase" for p in actual_fmt["paragraphs"]]
            
            if any(solved_uc) and solved_uc == actual_uc:
                q_passed = True
                remark = "Uppercase formatting 100% matched with solved file"
            else:
                remark = "Uppercase formatting mismatch (अपर केस केवल उसी पैराग्राफ पर होना चाहिए जो साल्व्ड फाइल में है)"

        elif "lowercase" in q_lower or "लोअर केस" in q_lower or "छोटे अक्षरों" in q_lower:
            solved_lc = [p.get("case_type") == "lowercase" for p in solved_fmt["paragraphs"]]
            actual_lc = [p.get("case_type") == "lowercase" for p in actual_fmt["paragraphs"]]
            
            if any(solved_lc) and solved_lc == actual_lc:
                q_passed = True
                remark = "Lowercase formatting 100% matched with solved file"
            else:
                remark = "Lowercase formatting mismatch (लोअर केस केवल उसी पैराग्राफ पर होना चाहिए जो साल्व्ड फाइल में है)"
                
        elif "font" in q_lower or "फॉन्ट" in q_lower or "font family" in q_lower or "kruti" in q_lower or "mangal" in q_lower or "arial" in q_lower:
            solved_has_font = any(any(r["font_name"] != "" for r in p["runs"]) for p in solved_fmt["paragraphs"])
            actual_has_font = any(any(r["font_name"] != "" for r in p["runs"]) for p in actual_fmt["paragraphs"])
            if solved_has_font and actual_has_font:
                match_font = True
                for sp, ap in zip(solved_fmt["paragraphs"], actual_fmt["paragraphs"]):
                    for sr, ar in zip(sp["runs"], ap["runs"]):
                        if sr["font_name"] and sr["font_name"] != ar["font_name"]:
                            match_font = False
                if match_font:
                    q_passed = True
                    remark = "Font family/name 100% matched with solved file"
                else:
                    remark = "Font family mismatch with solved master file"
            elif not solved_has_font and not actual_has_font:
                q_passed = True
                remark = "Font condition matched"
            else:
                remark = "Font presence mismatch"
                
        elif "character spacing" in q_lower or "scale" in q_lower or "spacing" in q_lower or "advanced font" in q_lower or "स्केल" in q_lower or "स्पेसिंग" in q_lower:
            match_adv_font = True
            for sp, ap in zip(solved_fmt["paragraphs"], actual_fmt["paragraphs"]):
                for sr, ar in zip(sp["runs"], ap["runs"]):
                    if sr.get("font_scale") != ar.get("font_scale") or \
                       sr.get("font_spacing") != ar.get("font_spacing") or \
                       sr.get("font_position") != ar.get("font_position"):
                        match_adv_font = False
                        break
            if match_adv_font and len(solved_fmt["paragraphs"]) > 0:
                q_passed = True
                remark = "Advanced character spacing 100% matched with solved file"
            else:
                remark = "Advanced character spacing mismatch with solved file"        

        elif "color" in q_lower or "कलर" in q_lower or "रंग" in q_lower or "red" in q_lower or "blue" in q_lower or "green" in q_lower:
            solved_has_color = any(any(r["font_color"] != "" for r in p["runs"]) for p in solved_fmt["paragraphs"])
            actual_has_color = any(any(r["font_color"] != "" for r in p["runs"]) for p in actual_fmt["paragraphs"])
            if solved_has_color and actual_has_color:
                match_color = True
                for sp, ap in zip(solved_fmt["paragraphs"], actual_fmt["paragraphs"]):
                    for sr, ar in zip(sp["runs"], ap["runs"]):
                        if sr["font_color"] and sr["font_color"] != ar["font_color"]:
                            match_color = False
                if match_color:
                    q_passed = True
                    remark = "Font color 100% matched with solved file"
                else:
                    remark = "Font color mismatch with solved master file"
            elif not solved_has_color and not actual_has_color:
                q_passed = True
                remark = "Font color condition matched"
            else:
                remark = "Font color presence mismatch"
                
        elif "capitalize" in q_lower or "each word" in q_lower or "प्रत्येक शब्द" in q_lower or "title case" in q_lower:
            solved_tc = [p.get("case_type") == "title" for p in solved_fmt["paragraphs"]]
            actual_tc = [p.get("case_type") == "title" for p in actual_fmt["paragraphs"]]
            
            if any(solved_tc) and solved_tc == actual_tc:
                q_passed = True
                remark = "Capitalize Each Word 100% matched with solved file"
            else:
                remark = "Capitalize Each Word mismatch (सही पैराग्राफ पर टाइटल केस होना चाहिए)"

        elif "sentence case" in q_lower or "सेंटेंस केस" in q_lower:
            actual_sc = all(p.get("case_type") == "sentence" for p in actual_fmt["paragraphs"] if p["text"])
            solved_sc = all(p.get("case_type") == "sentence" for p in solved_fmt["paragraphs"] if p["text"])
            if actual_sc and actual_sc == solved_sc:
                q_passed = True
                remark = "Sentence case 100% matched"
            else:
                remark = "Sentence case mismatch"

        elif "italic" in q_lower or "इटैलिक" in q_lower:
            actual_it = any(any(r["italic"] for r in p["runs"]) for p in actual_fmt["paragraphs"])
            solved_it = any(any(r["italic"] for r in p["runs"]) for p in solved_fmt["paragraphs"])
            if actual_it and actual_it == solved_it:
                q_passed = True
                remark = "Italic formatting 100% matched"
            else:
                remark = "Italic formatting mismatch"

        elif "underline" in q_lower or "अंडरलाइन" in q_lower:
            actual_ul = any(any(r["underline"] for r in p["runs"]) for p in actual_fmt["paragraphs"])
            solved_ul = any(any(r["underline"] for r in p["runs"]) for p in solved_fmt["paragraphs"])
            if actual_ul and actual_ul == solved_ul:
                q_passed = True
                remark = "Underline 100% matched"
            else:
                remark = "Underline mismatch"

        elif "center" in q_lower or "मध्य" in q_lower:
            actual_center = any(p["alignment"] == "CENTER" for p in actual_fmt["paragraphs"])
            solved_center = any(p["alignment"] == "CENTER" for p in solved_fmt["paragraphs"])
            if actual_center and actual_center == solved_center:
                q_passed = True
                remark = "Center alignment 100% matched"
            else:
                remark = "Center alignment mismatch"

        elif "justify" in q_lower:
            actual_jst = any(p["alignment"] == "JUSTIFY" for p in actual_fmt["paragraphs"])
            solved_jst = any(p["alignment"] == "JUSTIFY" for p in solved_fmt["paragraphs"])
            if actual_jst and actual_jst == solved_jst:
                q_passed = True
                remark = "Justify alignment 100% matched"
            else:
                remark = "Justify alignment mismatch"

        elif "header" in q_lower or "हेडर" in q_lower:
            if actual_fmt["has_header"] and actual_fmt["header_text"] == solved_fmt["header_text"]:
                q_passed = True
                remark = "Header 100% matched with solved file"
            else:
                remark = "Header content or presence mismatch"
                
        elif ("indent" in q_lower or "special" in q_lower or "first line" in q_lower or "hanging" in q_lower) and not ("superscript" in q_lower or "subscript" in q_lower):
            match_indent_spacing = True
            for sp, ap in zip(solved_fmt["paragraphs"], actual_fmt["paragraphs"]):
                if sp.get("left_indent") != ap.get("left_indent") or \
                   sp.get("right_indent") != ap.get("right_indent") or \
                   sp.get("first_line_indent") != ap.get("first_line_indent") or \
                   sp.get("hanging_indent") != ap.get("hanging_indent") or \
                   sp.get("space_before") != ap.get("space_before") or \
                   sp.get("space_after") != ap.get("space_after") or \
                   sp.get("line_spacing") != ap.get("line_spacing"):
                    match_indent_spacing = False
                    break
            if match_indent_spacing and len(solved_fmt["paragraphs"]) > 0:
                q_passed = True
                remark = "Paragraph indentation and spacing 100% matched with solved file"
            else:
                remark = "Paragraph indentation or spacing mismatch with solved file"

        elif "footer" in q_lower or "फुटर" in q_lower:
            if actual_fmt["has_footer"] and actual_fmt["footer_text"] == solved_fmt["footer_text"]:
                q_passed = True
                remark = "Footer 100% matched with solved file"
            else:
                remark = "Footer content or presence mismatch"

        # बॉर्डर (Page or Paragraph Border) जाँचने का 100% सटीक नियम
        elif "border" in q_lower or "बॉर्डर" in q_lower:
            is_paragraph_border = ("paragraph" in q_lower or "पैराग्राफ" in q_lower or "right" in q_lower or "दाएं" in q_lower)
            
            if is_paragraph_border:
                # पैरा-दर-पैरा सटीक तुलना कि क्या बिल्कुल सही पैराग्राफ पर राइट बॉर्डर है
                solved_borders = [p.get("right_border", False) for p in solved_fmt["paragraphs"]]
                actual_borders = [p.get("right_border", False) for p in actual_fmt["paragraphs"]]
                
                if solved_borders == actual_borders and any(solved_borders):
                    q_passed = True
                    remark = "Paragraph right border 100% matched with solved file"
                else:
                    remark = "Paragraph right border mismatch (बॉर्डर केवल उसी पैराग्राफ पर होना चाहिए जो साल्व्ड फाइल में है)"
            else:
                # सामान्य पेज बॉर्डर की जाँच
                if actual_fmt["has_page_border"] == solved_fmt["has_page_border"]:
                    q_passed = True
                    remark = "Page border 100% matched"
                else:
                    remark = "Page border mismatch"

        elif "watermark" in q_lower or "वॉटरमार्क" in q_lower:
            if actual_fmt["has_watermark"] == solved_fmt["has_watermark"]:
                q_passed = True
                remark = "Watermark 100% matched"
            else:
                remark = "Watermark mismatch"

        elif "subscript" in q_lower or "सब्सक्रिप्ट" in q_lower or "पांव में" in q_lower:
            actual_sub = any(any(r.get("subscript") for r in p["runs"]) for p in actual_fmt["paragraphs"])
            solved_sub = any(any(r.get("subscript") for r in p["runs"]) for p in solved_fmt["paragraphs"])
            if actual_sub and actual_sub == solved_sub:
                q_passed = True
                remark = "Subscript 100% matched with solved file"
            else:
                remark = "Subscript formatting mismatch with solved file"                

        # सुपरस्क्रिप्ट (Superscript) का सटीक नियम
        elif "superscript" in q_lower or "सुपरस्क्रिप्ट" in q_lower or "घात" in q_lower:
            solved_has_super = any(any(r.get("superscript") for r in p["runs"]) for p in solved_fmt["paragraphs"])
            actual_has_super = any(any(r.get("superscript") for r in p["runs"]) for p in actual_fmt["paragraphs"])
            
            if solved_has_super and actual_has_super:
                match_super = True
                # पैरा-दर-पैरा और रन-दर-रन सटीक (Exact) तुलना
                for sp, ap in zip(solved_fmt["paragraphs"], actual_fmt["paragraphs"]):
                    for sr, ar in zip(sp["runs"], ap["runs"]):
                        if sr.get("superscript") != ar.get("superscript"):
                            match_super = False
                            break
                    if not match_super:
                        break
                
                if match_super:
                    q_passed = True
                    remark = "Superscript 100% matched with solved file"
                else:
                    remark = "Superscript formatting mismatch with solved file (केवल उसी शब्द या पैराग्राफ पर सुपरस्क्रिप्ट होना चाहिए जो साल्व्ड फाइल में है)"
            elif not solved_has_super and not actual_has_super:
                q_passed = True
                remark = "Superscript condition matched"
            else:
                remark = "Superscript presence mismatch with solved file"
                
        elif "replace" in q_lower or "find" in q_lower or "find what" in q_lower or "ढूंढकर" in q_lower or "बदलें" in q_lower or "रिप्लेस" in q_lower:
            quotes_words = re.findall(r'["\']([^"\']+)["\']', q)
            target_word = "the"
            if len(quotes_words) >= 1:
                target_word = quotes_words[0].lower()
            elif "the" in q_lower:
                target_word = "the"
            
            is_formatting_replace = "bold" in q_lower or "बोल्ड" in q_lower or "highlight" in q_lower or "हाईलाइट" in q_lower
            if is_formatting_replace:
                match_replace_fmt = True
                word_found_in_solved = False
                for sp, ap in zip(solved_fmt["paragraphs"], actual_fmt["paragraphs"]):
                    for sr, ar in zip(sp["runs"], ap["runs"]):
                        if target_word in sr["text"].lower():
                            word_found_in_solved = True
                            if sr.get("bold") != ar.get("bold") or sr.get("highlight") != ar.get("highlight"):
                                match_replace_fmt = False
                                break
                if word_found_in_solved and match_replace_fmt:
                    q_passed = True
                    remark = f"Advanced Find & Replace with formatting for '{target_word}' 100% matched"
                else:
                    q_passed = True
                    remark = f"Find & Replace formatting applied for '{target_word}'"
            else:
                solved_text = solved_fmt["full_text"].strip().lower()
                actual_text = actual_fmt["full_text"].strip().lower()
                if actual_text == solved_text and len(solved_text) > 0:
                    q_passed = True
                    remark = "Simple Text Replace 100% matched with solved file"
                else:
                    if target_word not in actual_text:
                        q_passed = True
                        remark = "Word successfully replaced"
                    else:
                        remark = f"Text replacement mismatch for '{target_word}'"
                
        elif "orientation" in q_lower or "landscape" in q_lower or "पोर्ट्रेट" in q_lower:
            if actual_fmt["orientation"] == solved_fmt["orientation"]:
                q_passed = True
                remark = "Page orientation 100% matched"
            else:
                remark = "Orientation mismatch"

        elif "margin" in q_lower or "मार्जिन" in q_lower:
            if actual_fmt["margins"] == solved_fmt["margins"]:
                q_passed = True
                remark = "Margins 100% matched"
            else:
                remark = "Margin settings mismatch"
                
        else:
            if actual_fmt["full_text"].strip() == solved_fmt["full_text"].strip() and len(solved_fmt["full_text"].strip()) > 0:
                q_passed = True
                remark = "Content & line formatting 100% matched with solved file"
            else:
                remark = "Text content or formatting does not match solved master file"

        if q_passed:
            earned_score += per_question
            details.append((f"Q{i}", q, "✓ Correct", f"+{per_question}", remark))
        else:
            details.append((f"Q{i}", q, "Incorrect", "0", remark))

    return round(earned_score, 2), max_score, details