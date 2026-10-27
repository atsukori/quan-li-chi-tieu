# Sổ Nhà — Hệ thống quản lý chi tiêu gia đình

Ứng dụng web Flask để ghi nhận thu nhập và chi tiêu của các thành viên trong gia đình. Giao diện dùng HTML/CSS và render ở phía máy chủ; không sử dụng JavaScript.

## Chức năng

- Đăng nhập, đăng xuất, phân quyền Admin và Thành viên.
- Lưu mật khẩu dưới dạng hash; hỗ trợ đổi mật khẩu.
- Ghi nhận khoản thu hoặc chi, ngày giao dịch, số tiền, mục đích và nguồn tiền (cá nhân/quỹ chung).
- Xem số dư ví cá nhân và quỹ chung; giao dịch quỹ chung được hiển thị cho các thành viên.
- Theo dõi hạn mức chi tiêu tháng, cảnh báo khi đã sử dụng từ 80% hoặc vượt hạn mức.
- Lọc lịch sử theo ngày, loại giao dịch và nguồn tiền; chỉnh sửa hoặc xóa giao dịch do mình tạo.
- Báo cáo theo ngày, tuần, tháng, năm hoặc khoảng ngày tùy chọn.
- Xuất các giao dịch đang lọc và bảng tổng hợp sang tệp Excel `.xlsx`.
- Admin tạo tài khoản thành viên, điều chỉnh hạn mức và khóa/mở khóa tài khoản.
- Giao diện responsive cho desktop và màn hình nhỏ.

## Công nghệ

- Python + Flask
- SQLite, không cần cài đặt máy chủ cơ sở dữ liệu riêng
- Jinja2 templates + CSS thuần, không dùng JavaScript
- openpyxl để xuất Excel

## Yêu cầu

- Python 3.8 trở lên; khuyến nghị Python 3.11 trở lên trên Windows.
- `pip`

## Cài đặt và chạy trên Windows

Mở Command Prompt hoặc PowerShell, chuyển vào thư mục dự án:

```powershell
cd QuanLyChiTieuGiaDinh
py -3 -m venv .venv
```

Kích hoạt môi trường ảo:

**PowerShell**

```powershell
.\.venv\Scripts\Activate.ps1
```

**Command Prompt (CMD)**

```bat
.venv\Scripts\activate.bat
```

Cài thư viện và khởi chạy:

```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python app.py
```

Mở trình duyệt tại <http://127.0.0.1:5000>.

Nếu lệnh `py -3` không tìm thấy Python, hãy cài Python 3 từ trang chính thức, chọn tùy chọn thêm Python vào PATH trong quá trình cài đặt, rồi mở một cửa sổ terminal mới.

## Tài khoản mặc định

- **Username:** `admin`
- **Password:** `123`

Tài khoản này được tự động tạo khi cơ sở dữ liệu khởi tạo lần đầu. Hãy đăng nhập và đổi mật khẩu ngay trong mục **Mật khẩu**. Mật khẩu thành viên khi được Admin tạo phải có ít nhất 6 ký tự; khi đổi mật khẩu cần ít nhất 8 ký tự.

## Cấu trúc thư mục

```text
QuanLyChiTieuGiaDinh/
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
├── setup_windows.bat
├── run_windows.bat
├── instance/
│   ├── .gitkeep
│   └── finance.db       # Tự sinh khi chạy; không commit lên GitHub
├── static/
│   └── style.css
└── templates/
    ├── base.html
    ├── login.html
    ├── dashboard.html
    ├── transactions.html
    ├── transaction_form.html
    ├── reports.html
    ├── admin_users.html
    ├── change_password.html
    ├── error.html
    └── partials_transaction_list.html
```

## Cách dữ liệu được tính

- **Ví cá nhân:** tổng thu trừ tổng chi có nguồn tiền cá nhân của chính thành viên.
- **Quỹ chung:** tổng thu trừ tổng chi có nguồn tiền chung của toàn gia đình.
- **Hạn mức tháng:** tổng các khoản chi được tài khoản đó ghi nhận trong tháng, bất kể giao dịch thuộc nguồn cá nhân hay quỹ chung. Đây là quy tắc theo người ghi nhận giao dịch, không tự động phân bổ một khoản chi chung cho nhiều thành viên.
- **Báo cáo thành viên:** hiển thị giao dịch cá nhân của tài khoản đang đăng nhập và các giao dịch thuộc quỹ chung. Dữ liệu cá nhân của thành viên khác không hiển thị cho thành viên thường; Admin có thể xem toàn bộ.

## Commit lên GitHub

Đứng ở thư mục cha của dự án:

```powershell
git init
git add .
git commit -m "Initial family expense tracker"
```

Sau đó tạo một repository rỗng trên GitHub và làm theo hướng dẫn GitHub để thêm remote rồi push. Không commit `instance/finance.db`, `.venv/`, secret hoặc mật khẩu thật.

## Lưu ý trước khi deploy

Đây là bản nền tảng chạy local, chưa phải cấu hình production hoàn chỉnh. Trước khi public lên Internet:

1. Đổi mật khẩu Admin mặc định.
2. Đặt biến môi trường `SECRET_KEY` thành chuỗi ngẫu nhiên dài, riêng tư; không đưa secret vào Git. Trong PowerShell có thể tạo nhanh bằng lệnh `$env:SECRET_KEY = (python -c "import secrets; print(secrets.token_hex(32))")` trước khi chạy ứng dụng.
3. Chạy sau HTTPS và bật cookie `Secure` trong cấu hình production.
4. Dùng máy chủ WSGI phù hợp, thiết lập sao lưu SQLite, giới hạn truy cập và bổ sung kiểm thử/bảo vệ vận hành.
5. Kiểm tra chính sách riêng tư và quyền xem dữ liệu nếu gia đình cần phân quyền chi tiết hơn.

Ứng dụng mặc định chỉ bind vào `127.0.0.1`, phù hợp chạy trên máy cá nhân. Không dùng máy chủ phát triển này làm máy chủ production trực tiếp.

## Ghi chú về tính năng mở rộng

Hiện tại hệ thống chỉ lưu số tiền VNĐ nguyên, không dùng phần thập phân. Không tự đưa ra tỷ giá ngoại tệ hoặc quy tắc tính thuế cố định vì các dữ liệu này cần nguồn cập nhật và phụ thuộc bối cảnh cụ thể. Có thể bổ sung sau bằng API tỷ giá có nguồn rõ ràng và mô-đun tính thuế theo quy định phù hợp.


Có thể chạy `setup_windows.bat` một lần để tạo môi trường và cài thư viện; những lần sau dùng `run_windows.bat`.
