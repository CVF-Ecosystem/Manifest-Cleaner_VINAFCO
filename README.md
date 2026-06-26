# VINAFCO Manifest Cleaner

> **Version:** 2.0.0  
> **Author:** Tien-Tan Thuan Port

Công cụ xử lý và làm sạch file Cargo Manifest từ hãng tàu VINAFCO & VIMC.

## ✨ Tính năng

- 📄 Hỗ trợ đọc các file manifest Excel từ hãng tàu VINAFCO & VIMC (.xls, .xlsx, .xlsm).
- 🔍 Tự động phân tích, bóc tách và trích xuất thông tin: Số BL, Số Container, Số Seal, Trọng lượng.
- 🧹 Chuẩn hóa tên chủ hàng (Consignee) và loại bỏ các từ khóa/thông tin thừa (tel, fax, address...).
- 📊 Phân loại F/E (Full/Empty) thông minh dựa trên cả mô tả hàng hóa và trọng lượng (VGM/Payload).
- 🗺️ Chuyển đổi mã kích cỡ container sang chuẩn VTOS ISO.
- 🎨 Tự động highlight các container rỗng (màu xanh nhạt) hoặc seal không hợp lệ (màu vàng).
- 📁 Xuất dữ liệu ra file Excel chuẩn định dạng phục vụ cho việc Import hệ thống cảng.

## 🚀 Cài đặt

### Yêu cầu hệ thống
- Python 3.8+
- Các thư viện Python phụ thuộc: `pandas`, `openpyxl`, `xlrd`, `xlsxwriter`

### Cài đặt dependencies

```bash
pip install pandas openpyxl xlrd xlsxwriter
```

## 💻 Hướng dẫn sử dụng

### Chạy ứng dụng từ mã nguồn

```bash
python vinafco_app.py
```

### Build thành file chạy độc lập (.exe)

Nếu muốn đóng gói ứng dụng để chạy trên môi trường không cài đặt Python:

```bash
pip install pyinstaller
pyinstaller VINAFCO_Manifest.spec
```

File thực thi `VINAFCO_Manifest.exe` sẽ được tạo ra tại thư mục `dist/`.

## 🔧 Cấu hình

Các quy tắc xử lý, bóc tách và định dạng màu sắc có thể tùy chỉnh thông qua file `config.ini` ở thư mục gốc:
- **Ngưỡng trọng lượng vỏ (Tare weights)** để tính toán F/E.
- **Quy tắc mapping** kích cỡ container (20GP, 40GP...) sang mã VTOS ISO.
- **Màu sắc hiển thị** đối với container rỗng hoặc seal lỗi.

## 📝 License

Thành viên phát triển nội bộ - Cảng Tiên Sa / Tân Thuận.
