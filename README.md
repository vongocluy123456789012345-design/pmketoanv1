# Phần mềm đồng bộ hóa đơn điện tử

Web app chạy trên máy (Python + FastAPI), có các chức năng:

1. Đăng nhập cổng **hoadondientu.gdt.gov.vn**. Captcha do người dùng tự nhập.
2. Đồng bộ hóa đơn **mua vào / bán ra**, gồm cả hóa đơn từ máy tính tiền, theo khoảng ngày (tự chia theo tháng).
3. Tải **XML** của từng hóa đơn và đọc các trường: người bán, người mua, hàng hóa, tiền, mã CQT, **mã tra cứu / link tra cứu**.
4. Nhận diện **nhà cung cấp HĐĐT** (VNPT, Viettel, MISA, BKAV, FPT, EasyInvoice, Thái Sơn…) qua MST tổ chức cung cấp giải pháp và tên miền trang tra cứu, rồi tải **PDF gốc**.
5. Không lấy được PDF gốc thì dựng **bản thể hiện từ XML**. Bản này có dòng ghi chú rõ là không phải PDF gốc. Người dùng cũng có thể upload PDF tải tay.
6. Lưu dữ liệu vào **Supabase**: bảng `invoices`, file XML/PDF lưu trong bucket riêng tư `invoices`.

## Tự động tải qua Chrome (khuyên dùng)

1. Bấm **"1. Mở trang hóa đơn điện tử"**. Phần mềm mở Chrome (hồ sơ riêng trong `data/chrome-profile`, không đụng Chrome thường ngày của bạn) tại hoadondientu.gdt.gov.vn.
2. **Bạn tự đăng nhập** trên cửa sổ Chrome đó (MST, mật khẩu, captcha). Phần mềm không nhìn thấy và không lưu mật khẩu.
3. Khi trạng thái chuyển sang "Đã đăng nhập MST …", chọn từ tháng / đến tháng và loại hóa đơn, rồi bấm **"2. Bắt đầu tải"**.

Phần mềm chạy lần lượt từng tháng và từng loại (mua vào có mã, mua vào không mã, bán ra, máy tính tiền mua vào/bán ra):
tra cứu danh sách → tải **bảng kê Excel** của cổng → tải **XML** từng hóa đơn → đọc mã/link tra cứu trong XML → tìm và tải **PDF gốc** ở trang của nhà cung cấp.
Lệnh được gọi ngay trong trang cổng thuế đang đăng nhập, với đúng header mà web cổng gửi, nên không bị chặn như khi gọi API từ ngoài.

Kết quả lưu tại `OUTPUT_DIR` (mặc định `Documents\HoaDonDienTu`):

```
<MST>/
  TongHop_2025-01_den_2025-03.xlsx          tổng hợp mọi hóa đơn, có link mở XML/PDF, NCC, mã tra cứu, lý do lỗi
  2025-01/
    Mua vao - Co ma/   BangKe_*.xlsx   XML/*.xml   PDF/*.pdf
    Ban ra/            ...
```

- Hóa đơn **máy tính tiền**: cổng bắt nhập captcha khi tra cứu. Phần mềm hiện ảnh captcha để bạn gõ. Bấm "Bỏ qua loại này" để bỏ qua.
- Chạy lại cùng khoảng tháng: file đã có sẽ được bỏ qua, chỉ tải phần còn thiếu.
- Hóa đơn chưa có PDF gốc: xem cột "Ghi chú / lỗi" và "Link tra cứu" trong file tổng hợp để tải tay.
- Luồng này chỉ lưu file trên máy, không cần Supabase.

## Nguồn dữ liệu hóa đơn

| Nguồn | Trạng thái | Ghi chú |
|---|---|---|
| **Nhập file tải từ cổng thuế** (khuyên dùng) | Hoạt động, đã test với file mẫu | Miễn phí. Kéo thả .xml / .zip / .xlsx / .xls tải từ hoadondientu.gdt.gov.vn. Excel cho danh sách nhanh, XML cho dữ liệu đầy đủ và mã tra cứu để lấy PDF gốc. Nhập lại nhiều lần không bị trùng hóa đơn. |
| **Tự nhập từ thư mục Downloads** | Hoạt động, đã test | Bật trên giao diện + khai báo danh sách MST khách hàng. Tải file trên cổng thuế là phần mềm tự nhập, tự gán đúng MST (theo MST người mua/bán trong XML, MST trong tiêu đề Excel). File không xác định được MST sẽ bị bỏ qua, ghi lý do. |
| **TConnect Open API** | Đã viết connector, chưa chạy thật | Dịch vụ bên thứ ba "lấy hộ" hóa đơn từ cổng thuế bằng MST + mật khẩu của bạn. Có phí theo hóa đơn. Trả JSON (không có XML gốc). Tài liệu: https://docs.tconnect.vn/docs/open-api/invoice-purchase/overview |
| Cổng thuế trực tiếp | **Bị cổng chặn** ("phát hiện hành vi không hợp lệ") | Giữ lại code tham khảo, không dùng được. |

Cấu hình TConnect: đăng ký tại https://tconnect.vn/vi/onboard, rồi điền các biến `TCONNECT_*` trong `.env`. Sau đó, trên giao diện, bấm **Kết nối qua TConnect** một lần cho mỗi MST.

## Cài đặt

1. Cài **Python 3.11+** từ https://www.python.org/downloads/ (khi cài, tick "Add python.exe to PATH").
2. Tạo project trên https://supabase.com, vào **SQL Editor** và chạy toàn bộ file [supabase/schema.sql](supabase/schema.sql).
3. Chạy `run.bat`. Lần đầu, file này sẽ:
   - tạo môi trường `.venv` và cài thư viện,
   - tạo file `.env` để bạn điền `SUPABASE_URL` và `SUPABASE_SERVICE_KEY` (Project Settings > API > service_role).
4. Chạy lại `run.bat`. Trình duyệt sẽ tự mở http://127.0.0.1:8000.

> `SUPABASE_SERVICE_KEY` có toàn quyền trên database. Chỉ để key này trong `.env` trên máy, không gửi cho ai và không đưa lên git.

## Cấu trúc

```
app/
  main.py               API web + phục vụ giao diện
  sync.py               điều phối: cổng thuế -> XML -> PDF -> Supabase
  db.py                 truy cập Supabase
  gdt/client.py         client API cổng thuế (captcha, đăng nhập, danh sách, export-xml)
  gdt/xml_parser.py     đọc XML hóa đơn TT78
  providers/
    providers.json      danh sách NCC hóa đơn + mẫu URL tải PDF  <- thêm/sửa NCC ở đây
    __init__.py         bộ máy nhận diện NCC + tải PDF (tự dò link PDF trong trang tra cứu)
    plugins.py          connector Python riêng cho NCC cần quy trình phức tạp
    fallback.py         dựng PDF từ XML
  templates/            giao diện web + mẫu PDF dự phòng
supabase/schema.sql     bảng + bucket
```

## Trạng thái PDF

| pdf_status | Ý nghĩa |
|---|---|
| `original` | Đã lấy được PDF gốc từ trang của NCC |
| `manual`   | PDF do người dùng tự upload |
| `rendered` | Không lấy được bản gốc, đang dùng bản dựng từ XML. Lý do xem ở `pdf_error` |
| `failed`   | Không lấy được bản gốc và cũng không dựng được bản dự phòng |

## Lưu ý quan trọng

- **Endpoint cổng thuế** (base `https://hoadondientu.gdt.gov.vn/api`: `/captcha`, `/security-taxpayer/authenticate`, `/{query|sco-query}/invoices/{purchase|sold|detail|export-xml}`) đã được đối chiếu với mã JS của cổng vào 11/2026. Cổng không có tài liệu chính thức. Nếu cổng thay đổi, sửa ở `QUERY_SETS` trong `app/gdt/client.py`. Lưu ý: địa chỉ cũ `:30000` không còn kết nối được.
- **Mẫu URL của các NCC trong `providers.json` đều đang `verified: false`**. Cần thử với hóa đơn thật của từng NCC, rồi sửa mẫu URL hoặc viết connector trong `plugins.py`. Ở mức tối thiểu, phần mềm luôn thử link tra cứu in trong XML và tự dò link PDF trong trang.
- Trang tra cứu nào **bắt buộc captcha** thì phần mềm không vượt qua. Hóa đơn đó sẽ hiện link "Trang tra cứu" để bạn tải PDF tay rồi bấm "Upload PDF".
- Phần mềm không lưu mật khẩu cổng thuế. Phiên đăng nhập chỉ nằm trong RAM và mất khi tắt chương trình.
