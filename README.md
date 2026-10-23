# TEXO – Lập Quyết định Đoàn Tư vấn

Ứng dụng Streamlit tạo file Quyết định (.docx) về đoàn tư vấn của Công ty Cổ phần TEXO Tư vấn và Đầu tư: thành lập đoàn, bổ sung cán bộ, bổ sung và thay thế cán bộ, phân công nhiệm vụ. Người dùng nhập thông tin hợp đồng và danh sách cán bộ, ứng dụng tự chọn mẫu phù hợp theo số người và xuất file Word đúng thể thức văn bản hành chính.

Tác giả: Hoàng Đức Vũ – Trưởng phòng Kỹ thuật TEXO.

---

## 1. Cài đặt và chạy

Yêu cầu: Python 3.10 trở lên, đã có `pip`.

**Cách nhanh (Windows):** nhấp đúp `run_app.bat`. File này tự cài thư viện trong `requirements.txt` rồi mở ứng dụng trên trình duyệt.

**Chạy thủ công:**

```bash
pip install -r requirements.txt
streamlit run app.py
```

Ứng dụng mở tại `http://localhost:8501`. Màn hình đầu yêu cầu mật khẩu truy cập (mặc định `texo2026`; đổi bằng `APP_PASSWORD` trong `.streamlit/secrets.toml`, mục Secrets của Streamlit Cloud hoặc biến môi trường).

**Bản 2.1 – đồng bộ giao diện** với bộ app Phòng Kỹ thuật (Đề cương, Khối lượng, Hồ sơ chất lượng, Legal Reviewer): đăng nhập gọn có logo TEXO, đầu trang logo + tiêu đề, ô tình trạng hồ sơ (loại QĐ, mẫu sẽ dùng, số cán bộ, thông tin còn thiếu), nút Đăng xuất. Giao diện dùng chung ở `core/ui.py`, logo ở `assets/logo_texo.png`.

Thư viện sử dụng: `streamlit >= 1.32.0`, `python-docx >= 1.1.0`.

## 2. Các mẫu Quyết định

Ứng dụng tự chọn mẫu theo loại Quyết định và số cán bộ nhập vào:

| Loại Quyết định | Mẫu | Điều kiện | Cách trình bày danh sách |
|---|---|---|---|
| Thành lập đoàn | M01 | 1–3 người | Liệt kê trong Điều 1 |
| | M01.2 | từ 4 người | Bảng danh sách đính kèm (trang 2) |
| Bổ sung cán bộ | M02 | 1–3 người | Liệt kê trong Điều 1 |
| | M02.2 | từ 4 người | Bảng danh sách đính kèm |
| Bổ sung và thay thế cán bộ | M03 | 1 người | Liệt kê trong Điều 1, kèm dòng "(thay thế …)" |
| | M03.2 | từ 2 người | Bảng đính kèm, có cột cán bộ bị thay thế |
| Phân công nhiệm vụ | M04 | 1–3 người | Liệt kê trong Điều 1 |
| | M04.2 | từ 4 người | Bảng danh sách đính kèm |

Nội dung Điều 2 và Điều 4, chức danh trưởng đoàn tự đổi theo loại dịch vụ tư vấn nhập vào:

| Loại dịch vụ (từ khóa nhận diện) | Chức danh trưởng đoàn |
|---|---|
| Giám sát, TVGS | Giám sát trưởng |
| Quản lý dự án, QLDA, TVQLDA | Giám đốc Quản lý dự án |
| Thẩm tra, TVTT | Chủ trì |
| Kiểm định, TVKĐ | Chủ trì |
| Khác | Trưởng đoàn |

## 3. Hướng dẫn sử dụng

1. Chọn loại Quyết định ở thanh bên trái.
2. Điền thông tin hợp đồng: số HĐ Chủ đầu tư, số HĐ nội bộ TEXO, ngày ký, tên Chủ đầu tư, tên Trung tâm, loại dịch vụ tư vấn, nội dung hợp đồng.
3. Số và ngày Quyết định có thể để trống, file sẽ để dạng `..........` để văn thư điền tay. Ngày nhập dạng `08/04/2026` được tự chuyển thành "ngày 08 tháng 04 năm 2026".
4. Nhập danh sách cán bộ: trình độ chuyên môn (ví dụ *KS Xây dựng*, *ThS Kinh tế*, *KTS*), họ và tên, chức vụ. Nhấn ➕ để thêm dòng, 🗑️ để xóa.
5. Nhấn **🖨️ Tạo Quyết định**, sau đó **📥 Tải xuống** file .docx.
6. Mở file bằng Microsoft Word, đọc lại toàn bộ nội dung và đối chiếu quy định Công ty trước khi trình ký.

## 4. Quy ước trình bày trong file xuất ra

- Khổ A4, lề trên 20 mm, dưới 20 mm, trái 30 mm, phải 20 mm.
- Font Times New Roman: quốc hiệu và khối ký 12pt, thân văn bản 13pt, tiêu đề QUYẾT ĐỊNH 14pt. Thân văn bản giãn dòng Exactly 17pt, căn đều hai lề.
- Quốc hiệu "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" (viết HÒA theo kiểu TEXO dùng).
- Dòng kẻ ngang sát chân chữ dưới "TƯ VẤN VÀ ĐẦU TƯ" (42 ký tự `_`) và dưới "Độc lập - Tự do - Hạnh phúc" (55 ký tự `_`), cỡ 5pt in đậm, giãn dòng Single, không cách trước/sau.
- Dòng cán bộ ghi liền trình độ và họ tên, không có dấu chấm ở giữa: `1. KS Xây dựng Nguyễn Văn A – Giám sát viên`.
- Dấu cuối dòng trong danh sách: đoàn 1 người thì dấu chấm (.); từ 2 người trở lên thì các dòng trên dấu chấm phẩy (;), dòng cuối dấu chấm (.). Ở mẫu M03, khi có dòng "(thay thế …)" thì dấu kết thúc đặt ở cuối dòng thay thế.
- Khối ký: CÔNG TY CỔ PHẦN TEXO / TƯ VẤN VÀ ĐẦU TƯ / KT. TỔNG GIÁM ĐỐC. Nơi nhận: Như điều 4; Chủ đầu tư; Lưu: VT.

## 5. Nội dung người lập cần tự bổ sung

Ứng dụng không xử lý các phần sau, người lập hoàn thiện sau khi tải file:

- Logo, con dấu chìm hoặc hình ảnh nhận diện.
- Căn chỉnh tinh (tab, indent đặc biệt) nếu khác định dạng mặc định.
- Chữ ký số hoặc chữ ký tay scan.
- Nội dung đặc thù theo từng hợp đồng mà mẫu chung không bao quát.

## 6. Cấu trúc thư mục

```
texo_qd_doan-main/
├── app.py               Giao diện Streamlit: đăng nhập, nhập liệu, kiểm tra, chọn mẫu, tải file
├── run_app.bat          Chạy nhanh trên Windows
├── requirements.txt     Thư viện cần cài
├── README.md
└── builders/
    ├── common.py        Hàm dùng chung: lề, font, quốc hiệu, dòng kẻ, căn cứ, Điều 1–4, khối ký
    ├── builder_m01.py   Thành lập đoàn (1–3 người)
    ├── builder_m01_2.py Thành lập đoàn (từ 4 người, bảng đính kèm)
    ├── builder_m02.py   Bổ sung cán bộ (1–3 người)
    ├── builder_m02_2.py Bổ sung cán bộ (từ 4 người)
    ├── builder_m03.py   Bổ sung và thay thế (1 người)
    ├── builder_m03_2.py Bổ sung và thay thế (từ 2 người)
    ├── builder_m04.py   Phân công nhiệm vụ (1–3 người)
    └── builder_m04_2.py Phân công nhiệm vụ (từ 4 người)
```

Muốn sửa định dạng chung (quốc hiệu, dòng kẻ, font, khối ký) thì sửa `builders/common.py`, thay đổi sẽ áp dụng cho cả 8 mẫu. Các hàm hay dùng:

- `add_header_table()` – bảng quốc hiệu hai cột.
- `add_underscore_divider(cell, char_count)` – dòng kẻ ngang sát chân chữ.
- `ten_can_bo(trinh_do, ho_ten)` – ghép trình độ và họ tên.
- `dau_cuoi_dong(idx, total)` – trả về `;` hoặc `.` cho từng dòng danh sách.
- `format_member_line(doc, idx, member, total)` – một dòng cán bộ trong Điều 1.

## 7. Lịch sử cập nhật

**20/10/2026**
- Thêm dòng kẻ ngang dưới "TƯ VẤN VÀ ĐẦU TƯ" và "Độc lập - Tự do - Hạnh phúc" cho cả 8 mẫu (theo cách của bộ mẫu văn bản hành chính TEXO).
- Đổi "CỘNG HOÀ" thành "CỘNG HÒA".
- Bỏ dấu chấm giữa trình độ chuyên môn và họ tên, áp dụng cho danh sách trong Điều 1, bảng đính kèm và dòng thay thế.
- Dấu cuối dòng danh sách: dấu chấm phẩy cho các dòng trên, dấu chấm cho dòng cuối.
- Bỏ ý nhắc "tự kẻ đường kẻ ngang" trong phần lưu ý cuối ứng dụng.
