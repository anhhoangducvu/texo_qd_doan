import streamlit as st
import io
import re
from core import ui
from builders import (
    builder_m01, builder_m01_2,
    builder_m02, builder_m02_2,
    builder_m03, builder_m03_2,
    builder_m04, builder_m04_2,
)

def build_ngay_qd_display(ngay_qd_raw: str) -> str:
    """
    Chuyển đổi ngày nhập sang định dạng văn bản hành chính VN.
    """
    s = ngay_qd_raw.strip()
    if not s:
        return "ngày ...... tháng ...... năm ......"

    m = re.match(r'^(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})$', s)
    if m:
        day, month, year = m.group(1), m.group(2), m.group(3)
        return f"ngày {int(day):02d} tháng {int(month):02d} năm {year}"

    if s.startswith("ngày"):
        return s

    return "ngày ...... tháng ...... năm ......"

APP_VERSION = "2.1"
TITLE = "Lập Quyết định Đoàn Tư vấn"
ui.setup("TEXO · " + TITLE, "📝")
ui.gate(TITLE)
ui.header(TITLE, "Phòng Kỹ thuật · Thành lập, bổ sung, thay thế, phân công đoàn tư vấn – xuất Word theo mẫu M01–M04")

# Sidebar
with st.sidebar:
    st.markdown("### 📋 Loại Quyết định")
    loai_qd = st.radio("Chọn loại văn bản cần lập:", [
        "Thành lập đoàn Tư vấn",
        "Bổ sung cán bộ (M02)",
        "Bổ sung và thay thế cán bộ (M03)",
        "Phân công nhiệm vụ (M04)",
    ])
    st.divider()
    ui.logout_button()
    st.caption(f"Lập QĐ đoàn tư vấn v{APP_VERSION} · Xây dựng: Hoàng Đức Vũ")

ui.callout("ℹ️ App chỉ hỗ trợ <b>tạo file Quyết định .docx</b>, không thay thế việc rà soát của người lập. "
           "Tải file về rồi <b>đọc lại toàn bộ</b>, đối chiếu quy định hiện hành của Công ty trước khi trình ký.", "blue")

with st.expander("📖 Hướng dẫn sử dụng", expanded=False):
    st.markdown("""
### Các bước thực hiện
1. **Chọn loại Quyết định** cần lập ở thanh bên trái.
2. **Điền đầy đủ thông tin** vào các ô trong phần *Thông tin Hợp đồng*.
3. **Thêm danh sách cán bộ** – nhấn ➕ để thêm dòng, 🗑️ để xoá.
4. Nhấn nút **🖨️ Tạo Quyết định** để tạo file.
5. Nhấn **📥 Tải xuống** để lưu file .docx về máy.
6. **Mở file bằng Microsoft Word** và kiểm tra lại toàn bộ nội dung trước khi in.

---

### Lưu ý về chọn mẫu
| Loại QĐ | Khi nào dùng |
|---|---|
| **M01** – Thành lập đoàn | Đoàn từ 1–3 thành viên (tự động chọn) |
| **M01.2** – Thành lập đoàn | Đoàn từ 4 thành viên trở lên (tự động chọn) |
| **M02** – Bổ sung cán bộ | Bổ sung 1–3 người (tự động chọn) |
| **M02.2** – Bổ sung cán bộ | Bổ sung 4 người trở lên (tự động chọn) |
| **M03** – Bổ sung & thay thế | Chỉ dùng khi có đúng 1 cán bộ thay đổi (tự động chọn) |
| **M03.2** – Bổ sung & thay thế | Từ 2 cán bộ thay đổi trở lên (tự động chọn) |
| **M04** – Phân công nhiệm vụ | Phân công 1–3 người (tự động chọn) |
| **M04.2** – Phân công nhiệm vụ | Phân công 4 người trở lên (tự động chọn) |

---

### Hướng dẫn nhập liệu nhanh
- **Số HĐ**: nhập đúng số hợp đồng, phân biệt HĐ của Chủ đầu tư và HĐ nội bộ TEXO.
- **Loại dịch vụ tư vấn**: nhập đầy đủ, ví dụ *Tư vấn giám sát*, *Tư vấn Quản lý dự án*,
  *Tư vấn Thẩm tra*... Nội dung Điều 2 và Điều 4 sẽ tự thay đổi theo.
- **Tên Trung tâm**: nhập đầy đủ, ví dụ *Trung tâm 09*. Không thêm "Trung tâm" thừa.
- **Trình độ chuyên môn**: ghi tắt chuẩn, ví dụ *KS Xây dựng*, *ThS Kinh tế*, *KTS*...
- **Số và ngày QĐ**: có thể để trống nếu chưa có số – ứng dụng sẽ để dạng
  `......` để điền tay sau. Nếu nhập ngày dạng `08/04/2026`, ứng dụng tự chuyển
  thành *ngày 08 tháng 04 năm 2026*.
""")

# PART A: CONTRACT INFORMATION
with st.expander("📋 Thông tin Hợp đồng", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        so_hd_cdt = st.text_input(
            "Số HĐ của Chủ đầu tư *",
            placeholder="VD: 123/HĐ-BQLDA"
        )
        ngay_hd = st.text_input(
            "Ngày ký Hợp đồng *",
            placeholder="VD: 01/01/2025"
        )
        ten_cdt = st.text_input(
            "Tên Chủ đầu tư / Ban Quản lý *",
            placeholder="VD: Ban QLDA Đầu tư xây dựng tỉnh X"
        )
    with col2:
        so_hd_texo = st.text_input(
            "Số HĐ nội bộ TEXO *",
            placeholder="VD: 45/HĐ-TEXO"
        )
        ten_tt = st.text_input(
            "Tên Trung tâm phụ trách *",
            placeholder="VD: Trung tâm 03",
            help="Nhập đầy đủ tên, ví dụ: Trung tâm 03, Trung tâm Hà Nội. Hệ thống tự ghép thành 'Giám đốc Trung tâm 03'."
        )
        loai_tu_van = st.text_input(
            "Loại dịch vụ tư vấn *",
            placeholder="VD: Tư vấn giám sát"
        )
        st.caption("💡 Ví dụ: Tư vấn giám sát (TVGS), Tư vấn Quản lý dự án (TVQLDA), Tư vấn Thẩm tra (TVTT), Kiểm định chất lượng...")

    noi_dung_hd = st.text_area(
        "Nội dung Hợp đồng (về việc...) *",
        placeholder="VD: tư vấn giám sát thi công xây dựng công trình Trường THCS Nguyễn Du, huyện X, tỉnh Y",
        height=80
    )
    
    st.divider()
    st.markdown("**Số và ngày Quyết định** *(tuỳ chọn – để trống nếu muốn điền tay sau)*")

    col_sqd, col_nqd = st.columns(2)
    with col_sqd:
        so_qd = st.text_input(
            "Số Quyết định",
            placeholder="VD: 123  →  sẽ ra: 123/QĐ-CT",
            help="Chỉ nhập phần số, hệ thống tự thêm /QĐ-CT"
        )
    with col_nqd:
        ngay_qd = st.text_input(
            "Ngày ký Quyết định",
            placeholder="VD: 08/04/2026",
            help="Nhập dạng ngày/tháng/năm. Hệ thống tự chuyển thành 'ngày 08 tháng 04 năm 2026' theo mẫu văn bản hành chính."
        )

# Logic xử lý hiển thị Số và Ngày QĐ
so_qd_display = f"{so_qd.strip()}/QĐ-CT" if so_qd.strip() else "........../QĐ-CT"
ngay_qd_display = build_ngay_qd_display(ngay_qd)

# PART B: MEMBERS LIST
CHUC_VU_OPTIONS = [
    "Giám sát trưởng",
    "Giám sát viên – Xây dựng",
    "Giám sát viên – Cơ điện",
    "Giám sát viên – Kiến trúc",
    "Giám sát viên – Hạ tầng kỹ thuật",
    "Giám sát viên – An toàn lao động",
    "Thành viên",
    "Nhập tay...",
]

if "M03" in loai_qd:
    # Logic cho M03 (Bổ sung và thay thế)
    if "members_m03" not in st.session_state:
        st.session_state.members_m03 = [{
            "trinh_do": "", "ho_ten": "", "chuc_vu": "",
            "la_thay_the": False,
            "trinh_do_cu": "", "ho_ten_cu": ""
        }]
    
    with st.expander("👥 Danh sách cán bộ Bổ sung/Thay thế", expanded=True):
        st.info("📌 Hướng dẫn: Đánh dấu vào ô 'Thay thế' nếu cán bộ này thay cho người cũ.")
        
        for i, member in enumerate(st.session_state.members_m03):
            st.markdown(f"**Cán bộ {i+1}**")
            col_td, col_ten, col_cv, col_del = st.columns([1.5, 3, 3, 0.5])
            with col_td:
                member["trinh_do"] = st.text_input("Trình độ", value=member["trinh_do"], key=f"m03_td_{i}")
            with col_ten:
                member["ho_ten"] = st.text_input("Họ tên", value=member["ho_ten"], key=f"m03_ten_{i}")
            with col_cv:
                cv_select = st.selectbox("Chức vụ", options=CHUC_VU_OPTIONS, key=f"m03_cvs_{i}")
                if cv_select == "Nhập tay...":
                    member["chuc_vu"] = st.text_input("Nhập chức vụ", key=f"m03_cvt_{i}")
                else:
                    member["chuc_vu"] = cv_select
            with col_del:
                if i > 0:
                    if st.button("🗑️", key=f"m03_del_{i}"):
                        st.session_state.members_m03.pop(i)
                        st.rerun()
            
            member["la_thay_the"] = st.checkbox("Đây là thay thế cho người cũ", value=member["la_thay_the"], key=f"m03_tt_{i}")
            if member["la_thay_the"]:
                col_td_cu, col_ten_cu = st.columns([1.5, 6.5])
                with col_td_cu:
                    member["trinh_do_cu"] = st.text_input("Trình độ (người cũ)", value=member["trinh_do_cu"], key=f"m03_tdcu_{i}")
                with col_ten_cu:
                    member["ho_ten_cu"] = st.text_input("Họ tên người cũ", value=member["ho_ten_cu"], key=f"m03_tencu_{i}")
            st.divider()
        
        if st.button("➕ Thêm cán bộ"):
            st.session_state.members_m03.append({
                "trinh_do": "", "ho_ten": "", "chuc_vu": "",
                "la_thay_the": False,
                "trinh_do_cu": "", "ho_ten_cu": ""
            })
            st.rerun()
    members_to_process = st.session_state.members_m03

else:
    # Logic cho các mẫu khác (M01, M01.2, M02, M04)
    if "members" not in st.session_state:
        st.session_state.members = [{"trinh_do": "", "ho_ten": "", "chuc_vu": ""}]
    
    with st.expander("👥 Danh sách cán bộ", expanded=True):
        st.info("""
        📌 Hướng dẫn nhập liệu:
        - **Trình độ CM**: KS Xây dựng, ThS Kinh tế, TS Địa kỹ thuật, KTS, KS Điện...
        - **Chức vụ**: Chọn từ danh sách hoặc nhập tay nếu cần cụ thể hạng mục.
        """)
        
        for i, member in enumerate(st.session_state.members):
            col_td, col_ten, col_cv, col_del = st.columns([1.5, 3, 3, 0.5])
            with col_td:
                member["trinh_do"] = st.text_input("Trình độ CM" if i==0 else "", value=member["trinh_do"], key=f"td_{i}", label_visibility="visible" if i==0 else "collapsed")
            with col_ten:
                member["ho_ten"] = st.text_input("Họ và tên" if i==0 else "", value=member["ho_ten"], key=f"ten_{i}", label_visibility="visible" if i==0 else "collapsed")
            with col_cv:
                cv_select = st.selectbox("Chức vụ" if i==0 else "", options=CHUC_VU_OPTIONS, key=f"cvs_{i}", label_visibility="visible" if i==0 else "collapsed")
                if cv_select == "Nhập tay...":
                    member["chuc_vu"] = st.text_input("Nhập chức vụ", key=f"cvt_{i}", label_visibility="collapsed")
                else:
                    member["chuc_vu"] = cv_select
            with col_del:
                if i > 0:
                    if st.button("🗑️", key=f"del_{i}"):
                        st.session_state.members.pop(i)
                        st.rerun()
        
        if st.button("➕ Thêm thành viên"):
            st.session_state.members.append({"trinh_do": "", "ho_ten": "", "chuc_vu": ""})
            st.rerun()
    members_to_process = st.session_state.members

# VALIDATION LOGIC
def validate(data):
    errs = []
    if not data["so_hd_cdt"]: errs.append("Thiếu số HĐ Chủ đầu tư")
    if not data["ten_cdt"]: errs.append("Thiếu tên Chủ đầu tư")
    if not data["noi_dung_hd"]: errs.append("Thiếu nội dung Hợp đồng")
    if not data["ten_tt"]: errs.append("Thiếu tên Trung tâm")
    if not data["loai_tu_van"]: errs.append("Thiếu loại dịch vụ tư vấn")
    
    valid_members = [m for m in data["members"] if m["ho_ten"].strip()]
    if not valid_members:
        errs.append("Chưa nhập ít nhất 1 thành viên")
    else:
        for m in valid_members:
            if not m["chuc_vu"]: errs.append(f"Chưa chọn chức vụ cho: {m['ho_ten']}")
    return errs, valid_members

data = {
    "so_hd_cdt": so_hd_cdt,
    "ngay_hd": ngay_hd,
    "ten_cdt": ten_cdt,
    "so_hd_texo": so_hd_texo,
    "ten_tt": ten_tt,
    "loai_tu_van": loai_tu_van,
    "noi_dung_hd": noi_dung_hd,
    "so_qd_display": so_qd_display,
    "ngay_qd_display": ngay_qd_display,
    "members": members_to_process
}

# PREVIEW & GENERATE
errors, members_valid = validate(data)

def _ma_mau(loai, n):
    if loai == "Thành lập đoàn Tư vấn":
        return "M01" if n <= 3 else "M01.2"
    if "thay thế" in loai.lower():
        return "M03" if n <= 1 else "M03.2"
    if "Bổ sung" in loai:
        return "M02" if n <= 3 else "M02.2"
    return "M04" if n <= 3 else "M04.2"


st.markdown("##### 📊 Tình trạng hồ sơ")
n_ok = len(members_valid)
k = st.columns(4)
_ngan = {"Thành lập đoàn Tư vấn": "Thành lập", "Bổ sung cán bộ (M02)": "Bổ sung",
         "Bổ sung và thay thế cán bộ (M03)": "Thay thế", "Phân công nhiệm vụ (M04)": "Phân công"}
ui.kpi(k[0], _ngan.get(loai_qd, "–"), "Loại Quyết định", loai_qd.split(" (")[0])
ui.kpi(k[1], _ma_mau(loai_qd, max(n_ok, 1)) if n_ok else "–", "Mẫu sẽ dùng", "tự chọn theo số người")
ui.kpi(k[2], n_ok, "Cán bộ đã nhập", ten_tt or "chưa nhập Trung tâm", "blue" if n_ok else "orange")
ui.kpi(k[3], len(errors), "Thông tin còn thiếu", "cần bổ sung" if errors else "đủ để tạo file",
       "orange" if errors else "green")
if errors:
    with st.expander("⚠️ Các thông tin cần bổ sung", expanded=False):
        for e in errors:
            st.write(f"- {e}")

# Nút tạo quyết định luôn hiển thị ở dưới cùng
if st.button("🖨️ Tạo Quyết định", type="primary", use_container_width=True):
    if errors:
        st.error("❌ Vui lòng sửa các lỗi thông tin trước khi tạo file.")
    else:
        # Cập nhật lại members chỉ lấy những người có tên
        data["members"] = members_valid
        n = len(members_valid)
        
        try:
            if loai_qd == "Thành lập đoàn Tư vấn":
                if n <= 3:
                    doc = builder_m01.build(data); ma_mau = "M01"
                else:
                    doc = builder_m01_2.build(data); ma_mau = "M01.2"
                
            elif "Bổ sung cán bộ" in loai_qd and "thay thế" not in loai_qd.lower():
                if n <= 3:
                    doc = builder_m02.build(data); ma_mau = "M02"
                else:
                    doc = builder_m02_2.build(data); ma_mau = "M02.2"

            elif "thay thế" in loai_qd.lower():
                if n <= 1:
                    doc = builder_m03.build(data); ma_mau = "M03"
                else:
                    doc = builder_m03_2.build(data); ma_mau = "M03.2"
                    st.info(f"ℹ️ Đoàn có {n} người → tự động dùng mẫu M03.2 (danh sách đính kèm)")

            elif "Phân công" in loai_qd:
                if n <= 3:
                    doc = builder_m04.build(data); ma_mau = "M04"
                else:
                    doc = builder_m04_2.build(data); ma_mau = "M04.2"
            else:
                st.error("❌ Không xác định được loại mẫu phù hợp.")
                st.stop()
            
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            
            ten_file = f"QD_{ma_mau}_{data['so_hd_texo'].replace('/', '-')}.docx"
            st.success(f"✅ Đã tạo xong file theo mẫu {ma_mau}!")
            st.download_button(
                label="📥 Tải xuống file .docx",
                data=buf,
                file_name=ten_file,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"❌ Có lỗi xảy ra: {str(e)}")

st.divider()
ui.callout(
    "⚠️ <b>App không tự xử lý được</b> – người lập bổ sung sau khi tải file: logo, dấu chìm, hình ảnh nhận diện; "
    "căn chỉnh tinh (tab, thụt lề đặc biệt); chữ ký số hoặc chữ ký scan; nội dung đặc thù theo từng hợp đồng. "
    "Người lập chịu trách nhiệm đọc lại, đối chiếu quy định Công ty và hoàn thiện file trước khi trình ký.", "orange")
