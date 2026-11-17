# AWBook – phần mềm hóa đơn điện tử, sao kê, đối chiếu thuế

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

## Lấy PDF gốc không cần đăng nhập cổng (Chrome chạy ẩn)

Khi tải PDF gốc, ngoài tải HTTP thường và connector riêng (MISA…), phần mềm mở Chrome/Edge **chạy ẩn** (hồ sơ trống, không đụng Chrome của bạn)
tại link tra cứu in trong XML hoặc trang tra cứu của NCC (`lookup_pages` trong `providers.json`), tự gõ mã tra cứu, bấm Tra cứu và bắt PDF
(file tải về, ZIP, JSON chứa PDF base64, tab mới, blob). Trang chỉ hiện hóa đơn dạng HTML thì in trang đó ra PDF. PDF phải chứa đúng số hóa đơn
+ ký hiệu/MST người bán mới được nhận. Trang **bắt captcha / kiểm tra chống robot** (EasyInvoice, Thái Sơn, BKAV reCAPTCHA, Cloudflare…) thì
phần mềm dừng và ghi "cần captcha" – không vượt captcha. Tắt bằng `HEADLESS_PDF=false` trong `.env`.

Trên mục **Tải hóa đơn** có thêm: *Lấy PDF* theo link + mã tra cứu (lưu vào `<nơi lưu>\TraCuuPDF`) và *Lấy bù PDF gốc còn thiếu* cho các tháng đã tải.

## Dự báo thuế GTGT (tờ khai 01/GTGT)

Mục **Dự báo VAT**: chọn MST, kỳ (tháng/quý), nhập **[22]** (số thuế còn được khấu trừ kỳ trước – tự gợi ý từ [43] của kỳ trước đã dự báo),
bấm **Dự báo VAT**. Phần mềm cộng hóa đơn đã tải (`.dulieu_YYYY-MM.json` + XML) theo từng chỉ tiêu [21]…[43] của mẫu 01/GTGT (TT80/2021):
bỏ hóa đơn bị thay thế/bị hủy; mua vào chỉ hóa đơn GTGT có thuế suất vào [23]/[24]; bán ra 8% cộng vào [32]/[33] và lập **Phụ lục giảm thuế GTGT**.
Kết quả lưu ở `<MST>\DuBaoVAT\<kỳ>\`: `ToKhai_01GTGT_*.pdf` (tờ khai, phụ lục 8%, bảng kê mua vào/bán ra, bảng kê chi tiết từng mặt hàng,
lưu ý), `.xlsx` và `.html` để in.

**Tự động:** từ 00:00 ngày đầu tháng (đầu quý nếu chọn kê khai theo quý), phần mềm tự dự báo kỳ vừa hết và báo lên chuông 🔔 + thông báo Windows.
Nếu Chrome của phần mềm đang đăng nhập đúng MST, phần mềm tải đủ hóa đơn kỳ đó trước rồi mới dự báo; chưa đăng nhập thì dự báo tạm bằng dữ liệu
đang có và nhắc đăng nhập (đăng nhập xong sẽ tự tải và dự báo lại). Phần mềm phải đang mở. Đây là **bản dự báo để kiểm tra**, không thay thế tờ khai nộp qua eTax/HTKK.

Mẫu tờ khai tự chọn theo kỳ: trước 01/7/2026 dùng mẫu 01/GTGT theo TT 80/2021; từ kỳ tháng 7/2026 (quý III/2026) dùng mẫu theo
**TT 89/2026/TT-BTC** (thêm [32b], [34a]; [27]=[29]+[30]+[32]+[32a]−[32b]; [34]=[26]+[27]+[34a]). Mỗi bản in có mục "Căn cứ pháp lý" kèm link.

## Trợ lý AI (bóng chat 💬 góc phải dưới)

- Khách mở là chat được ngay (dùng khóa dựng sẵn `app/ai_builtin.json` của nhà phát hành; bản .exe không hiện phần khóa API).
- Tra luật trực tuyến: bài mới nhất trên MISA AMIS / MISA SME / meInvoice (API tìm kiếm WordPress công khai) + kho kiến thức
  `app/ai_luat.md` đã kiểm chứng trên thuvienphapluat.vn (trang này chặn truy cập tự động nên không tra trực tiếp).
- 👥 Nhóm chat: Kế toán trưởng (Gemini Flash) trả lời + làm việc trên app; Chuyên viên thuế (Gemini Flash-Lite) soát căn cứ;
  Kiểm soát rủi ro (Gemma 4) nêu rủi ro, việc cần làm. Thêm/bớt thành viên: mục "group" trong `data/ai.json`.
- Gói Gemini miễn phí: Flash ~20 lượt/ngày; hết lượt app tự chuyển Flash-Lite. Nhiều khách dùng chung nên bật trả phí cho khóa.

## Trợ lý AI – chi tiết cài đặt

Chat tiếng Việt để làm việc trên app ("tải hóa đơn tháng 10 về cho tôi", "dự báo VAT quý 3") và hỏi thuế/luật. Câu hỏi luật được tra
trên thuvienphapluat.vn, vbpl.vn, chinhphu.vn, gdt.gov.vn, mof.gov.vn… (công cụ web search/fetch của Claude, giới hạn domain) và trả lời
kèm điều khoản, link. Chọn AI trong ô "Cài đặt AI":
- **Google Gemini – miễn phí (mặc định):** lấy khóa tại https://aistudio.google.com/apikey (đăng nhập Google, không cần thẻ). Gói free giới
  hạn số lượt/phút và /ngày; hết lượt model chính thì tự chuyển sang `gemini-flash-lite-latest`. Khóa Gemini cũng dùng để tra cứu luật
  (Google Search) cho mọi lựa chọn AI miễn phí.
- **Groq, OpenRouter (model `:free`)** – miễn phí có giới hạn; **Ollama** – chạy trên máy, miễn phí hoàn toàn (cần máy mạnh).
- **Claude (Anthropic)** – trả phí, mạnh nhất; tra cứu luật bằng web search riêng.
Khóa lưu ở `data/ai.json` trên máy (hoặc biến `GEMINI_API_KEY`, `GROQ_API_KEY`, `OPENROUTER_API_KEY`, `ANTHROPIC_API_KEY` trong `.env`).
Trợ lý không nhập mật khẩu và không tự giải captcha.

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

