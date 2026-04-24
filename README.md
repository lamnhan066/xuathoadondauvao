# Xuất Hóa Đơn Đầu Vào (MVP)

## Phạm vi hiện tại

* Luồng đăng nhập hybrid:
  * Nếu người dùng đã đăng nhập tại `https://hoadondientu.gdt.gov.vn/`, extension tái sử dụng session hiện có.
  * Nếu chưa đăng nhập, extension mở một cửa sổ đăng nhập riêng để người dùng xác thực trước.
* Nút hành động của extension mở giao diện toàn trang thay vì popup nhỏ.
* Bộ lọc ngày từ - đến.
* Tải và gộp 2 nguồn hóa đơn mua vào:
  * `/query/invoices/purchase` với `ttxly=5` mặc định
  * `/sco-query/invoices/purchase` với `ttxly=8` mặc định
* Danh sách hóa đơn gộp, xem chi tiết từng hóa đơn, chọn nhiều.
* Xuất từng hóa đơn đã chọn ra file `.xlsx` với header tương thích Tendoo.

## Cách cài đặt và kích hoạt extension

Extension hiện có 3 cách để tải và cài đặt: cài trực tiếp từ **Chrome Web Store**, clone project về máy, hoặc tải bản release được tạo từ workflow trên GitHub.

### Cách 1: Cài từ Chrome Web Store

1. Mở Chrome và truy cập trang tiện ích chính thức:

  https://chromewebstore.google.com/detail/bjfejbopdhigplifbjibfiejacbbgmkc

1. Nhấn **Thêm vào Chrome** (Add to Chrome) và xác nhận quyền yêu cầu.
2. Sau khi cài, mở menu `chrome://extensions` để ghim hoặc quản lý tiện ích.
3. Mở extension từ thanh công cụ để bắt đầu sử dụng.

### Cách 2: Tải từ GitHub Release

1. Vào trang Releases của repository trên GitHub.
2. Tải file release `.zip` do workflow tạo ra.
3. Giải nén file `.zip` ra một thư mục trên máy.
4. Mở Chrome và vào `chrome://extensions`.
5. Bật chế độ Nhà phát triển (Developer mode).
6. Chọn `Load unpacked`.
7. Trỏ tới thư mục đã giải nén, tức thư mục chứa file `manifest.json`.

### Cách 3: Clone project về máy (phát triển)

1. Clone repository về máy.
2. Mở terminal tại thư mục `hoadon/`.
3. Cài dependencies Node.js bằng lệnh:

```bash
npm install
```

Bước này sẽ tạo thư mục `node_modules/` để đồng bộ môi trường phát triển.

4. Mở Chrome và vào `chrome://extensions`.
5. Bật chế độ Nhà phát triển (Developer mode).
6. Chọn `Load unpacked`.
7. Trỏ tới thư mục `hoadon/` vừa clone, tức thư mục có file `manifest.json`.


## Ghi chú

* MVP này xuất trực tiếp `.xlsx` mà không cần workbook mẫu Tendoo đi kèm.
* API được gọi với `credentials: include`, nên người dùng cần đăng nhập trên cổng Hóa đơn điện tử chính thức trước.
* Nếu cài từ source để phát triển tiếp, hãy chạy lại `npm install` sau khi clone nếu `node_modules/` chưa tồn tại.
* Extension đã được phát hành trên Chrome Web Store: https://chromewebstore.google.com/detail/bjfejbopdhigplifbjibfiejacbbgmkc
