# Converted File Tools - Update 26/11/2024 ver 8.24.22

## Thư viện và Công cụ
- Python 3.x
- pandas
- pdf2docx
- python-docx
- openpyxl
- reportlab
- markdown
- PyPDF2
- python-pptx
- docx2pdf

## Chức năng
1. Chuyển đổi giữa các định dạng tệp sau:
   - PDF <-> DOCX
   - PDF <-> XLSX
   - PDF <-> TXT
   - DOCX <-> PDF
   - XLSX <-> DOCX
   - XLSX <-> PDF
   - XLSX <-> CSV
   - TXT <-> DOCX
   - TXT <-> PDF
   - TXT <-> MD
   - CSV <-> XLSX
   - PPTX -> PDF
   - PPTX -> DOCX
   - MD <-> TXT
   - MD -> HTML

2. Xử lý lỗi và ghi log cho mỗi quá trình chuyển đổi
3. Giao diện dòng lệnh cho người dùng chọn tùy chọn chuyển đổi
4. Hỗ trợ đường dẫn tệp với dấu phân cách thư mục '\'

## Yêu cầu phi chức năng
1. Xử lý ngoại lệ để tránh crash ứng dụng
2. Logging chi tiết để dễ dàng gỡ lỗi
3. Cấu trúc mã nguồn module hóa để dễ bảo trì và mở rộng
4. Sử dụng type hints để cải thiện khả năng đọc và bảo trì mã

## Giới hạn 
1. Một số chuyển đổi (như DOCX sang PDF) có thể yêu cầu Microsoft Word được cài đặt trên hệ thống
2. Chất lượng chuyển đổi có thể phụ thuộc vào độ phức tạp của tệp gốc
3. Hiệu suất có thể bị ảnh hưởng khi xử lý các tệp lớn
4. Vẫn sẽ lỗi ở các yêu cầu chuyển đổi không đúng định dạng nội dung

## Cập nhật
1. Fix lỗi chuyển đổi PDF -> XLSX
2. Điều chỉnh logic cho các chuyển đổi còn lại
3. Điều chỉnh lỗi chuyển đổi định dạng nội dung nâng cao (đang phát triển)

## Lưu ý

*Code vẫn đang được tôi hoàn thiện và nâng cấp, mọi đóng góp hay góp ý từ bạn rất quan trọng để tôi có thể điều chỉnh một cách hoàn thiện - Xin cảm ơn*

## Liên hệ với tôi 

[Telegram](https://t.me/tanbaycu)

[Email](mailto:tranminhtan4953@gmail.com)

