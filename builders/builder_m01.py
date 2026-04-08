from docx import Document
from builders.common import *

def build(data):
    doc = Document()
    set_page_margins(doc)
    
    add_header_table(doc, data)
    add_title(doc, f"Thành lập đoàn {data['loai_tu_van']}")
    add_tgd(doc)
    add_can_cu(doc, data)
    add_quyet_dinh_label(doc)
    
    # Điều 1
    p1 = build_dieu1_header(doc, "Thành lập", data['loai_tu_van'], data)
    r_tail = p1.add_run(" gồm các cán bộ có tên sau đây:")
    fmt_body(r_tail)
    
    for i, m in enumerate(data["members"], 1):
        format_member_line(doc, i, m)
    
    add_dieu_2_3_4(doc, data["ten_tt"], data["loai_tu_van"])
    add_footer_table(doc)
    return doc
