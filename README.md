# 🔄 Công Cụ Chuyển Đổi Định Dạng Tệp - tanbaycu

<img src="https://img.shields.io/badge/Phiên%20bản-12.13.58-blue?style=for-the-badge&logo=semver&logoColor=white" alt="Version"/>
<img src="https://img.shields.io/badge/Cập%20nhật-27%2F11%2F2024-green?style=for-the-badge&logo=clockify&logoColor=white" alt="Update"/>
<img src="https://img.shields.io/badge/Python-3.7%2B-yellow?style=for-the-badge&logo=python&logoColor=white" alt="Python"/>
<img src="https://img.shields.io/badge/Giấy%20phép-MIT-orange?style=for-the-badge&logo=license&logoColor=white" alt="License"/>
<img src="https://img.shields.io/badge/Trạng%20thái-Đang%20phát%20triển-red?style=for-the-badge&logo=github&logoColor=white" alt="Status"/>

> Công cụ chuyển đổi đa định dạng tệp tin với giao diện dòng lệnh đơn giản và hiệu quả

## 📚 Thư viện và Công cụ

Cài đặt các thư viện cần thiết bằng lệnh sau:

```bash
pip install pandas numpy pdf2docx python-docx openpyxl reportlab markdown PyPDF2 python-pptx docx2pdf mammoth html2text pdfplumber pywin32
```

## ✨ Tính năng

### 📄 Chuyển đổi giữa các định dạng tệp sau:

<table>
  <tr>
    <th>Từ \ Đến</th>
    <th>PDF</th>
    <th>DOCX</th>
    <th>XLSX</th>
    <th>TXT</th>
    <th>CSV</th>
    <th>MD</th>
    <th>HTML</th>
  </tr>
  <tr>
    <td><b>PDF</b></td>
    <td>-</td>
    <td>✅</td>
    <td>✅</td>
    <td>✅</td>
    <td>❌</td>
    <td>❌</td>
    <td>❌</td>
  </tr>
  <tr>
    <td><b>DOCX</b></td>
    <td>✅</td>
    <td>-</td>
    <td>✅</td>
    <td>❌</td>
    <td>❌</td>
    <td>❌</td>
    <td>❌</td>
  </tr>
  <tr>
    <td><b>XLSX</b></td>
    <td>✅</td>
    <td>✅</td>
    <td>-</td>
    <td>❌</td>
    <td>✅</td>
    <td>❌</td>
    <td>❌</td>
  </tr>
  <tr>
    <td><b>TXT</b></td>
    <td>✅</td>
    <td>✅</td>
    <td>❌</td>
    <td>-</td>
    <td>❌</td>
    <td>✅</td>
    <td>❌</td>
  </tr>
  <tr>
    <td><b>CSV</b></td>
    <td>❌</td>
    <td>❌</td>
    <td>✅</td>
    <td>❌</td>
    <td>-</td>
    <td>❌</td>
    <td>❌</td>
  </tr>
  <tr>
    <td><b>PPTX</b></td>
    <td>✅</td>
    <td>✅</td>
    <td>❌</td>
    <td>❌</td>
    <td>❌</td>
    <td>❌</td>
    <td>❌</td>
  </tr>
  <tr>
    <td><b>MD</b></td>
    <td>❌</td>
    <td>❌</td>
    <td>❌</td>
    <td>✅</td>
    <td>❌</td>
    <td>-</td>
    <td>✅</td>
  </tr>
</table>


### 🛠️ Tính năng bổ sung:

- ✅ Xử lý lỗi và ghi nhật ký cho mỗi quá trình chuyển đổi
- ✅ Giao diện dòng lệnh thân thiện với người dùng
- ✅ Hỗ trợ đường dẫn tệp với dấu phân cách thư mục (`\`)
- ✅ Chuyển đổi liên tục nhiều tệp tin
- ✅ Tự động phát hiện định dạng tệp nguồn


## 🔧 Yêu cầu phi chức năng

1. **Xử lý ngoại lệ**: Ngăn chặn ứng dụng bị crash khi gặp lỗi
2. **Ghi nhật ký chi tiết**: Giúp dễ dàng gỡ lỗi và theo dõi quá trình
3. **Cấu trúc mã mô-đun**: Dễ bảo trì và mở rộng
4. **Sử dụng gợi ý kiểu**: Nâng cao khả năng đọc và bảo trì mã


## ⚠️ Hạn chế

1. **Phụ thuộc phần mềm**: Một số chuyển đổi (ví dụ: DOCX sang PDF) có thể yêu cầu Microsoft Word được cài đặt trên hệ thống
2. **Chất lượng chuyển đổi**: Có thể phụ thuộc vào độ phức tạp của tệp gốc
3. **Hiệu suất**: Có thể giảm khi xử lý các tệp lớn
4. **Lỗi định dạng**: Có thể xảy ra đối với các tệp đầu vào được định dạng không chính xác


## 🆕 Cập nhật mới

### Phiên bản 12.13.58 (27/11/2024)

- ✅ Thêm vòng lặp cho chuyển đổi liên tục
- ✅ Thêm chuyển đổi DOCX -> XLSX
- ✅ Cải thiện hỗ trợ định dạng nội dung nâng cao (đang phát triển)
- ✅ Tối ưu hóa lỗi trong lần chuyển đổi đầu tiên
- ✅ Cải thiện xử lý ngoại lệ với decorator `chuyen_doi_an_toan`
- ✅ Thêm phương pháp thay thế cho chuyển đổi PDF sang DOCX



## 📝 Hướng dẫn sử dụng

1. **Chạy chương trình**:

```shellscript
python converter.py
```


2. **Nhập đường dẫn tệp**:

   1. Nhập đường dẫn đầy đủ đến tệp bạn muốn chuyển đổi
   2. Sử dụng dấu `\` làm dấu phân cách thư mục



3. **Chọn định dạng đích**:

   1. Chương trình sẽ hiển thị các tùy chọn chuyển đổi có sẵn dựa trên định dạng tệp nguồn
   2. Nhập số tương ứng với định dạng đích mong muốn



4. **Xem kết quả**:

   1. Chương trình sẽ thông báo kết quả chuyển đổi
   2. Tệp đã chuyển đổi sẽ được lưu trong cùng thư mục với tệp nguồn

## 👥 Đóng góp

Chúng tôi rất hoan nghênh mọi đóng góp! Dưới đây là cách bạn có thể tham gia:

<div class="contribution">
  <div class="contribution-item">
    <h4>🐛 Báo cáo lỗi</h4>
    <p>Tạo issue mới trên GitHub với mô tả chi tiết về lỗi và cách tái tạo</p>
  </div> <div class="contribution-item">
    <h4>💡 Đề xuất tính năng</h4>
    <p>Chia sẻ ý tưởng của bạn thông qua GitHub Issues hoặc Discussions</p>
  </div> <div class="contribution-item">
    <h4>🔧 Gửi Pull Request</h4>
    <p>Đóng góp mã nguồn bằng cách fork repository và tạo pull request</p>
  </div>
</div>

### Quy trình đóng góp

1. Fork repository
2. Tạo nhánh mới (`git checkout -b feature/amazing-feature`)
3. Commit thay đổi (`git commit -m 'Add some amazing feature'`)
4. Push lên nhánh của bạn (`git push origin feature/amazing-feature`)
5. Mở Pull Request

## 📌 Ghi chú

*Mã nguồn vẫn đang trong quá trình phát triển và cải tiến. Phản hồi và đề xuất của bạn rất quý giá để hoàn thiện công cụ này. Cảm ơn sự hỗ trợ của bạn!*

## 📞 Hỗ trợ & Liên hệ

<div class="contact-grid">
  <a href="https://t.me/tanbaycu" class="contact-item">
    <img src="https://img.shields.io/badge/Telegram-@tanbaycu-blue?style=for-the-badge&logo=telegram" alt="Telegram"/>

  </a>
  <a href="https://github.com/username/file-converter-tool/issues" class="contact-item">
    <img src="https://img.shields.io/badge/GitHub-Issues-gray?style=for-the-badge&logo=github" alt="GitHub Issues"/>
    
  </a>
</div>

<div align="center">
  <img src="https://img.shields.io/badge/Made%20with-%E2%9D%A4%EF%B8%8F-red?style=for-the-badge" alt="Made with love"/>
  <br/>
  <p>© 2025 tanbaycu</p>
  <p>
    <a href="#top">⬆️ Về đầu trang</a>
  </p>
</div>
<style>
.grid-container {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 20px;
}

.feature-card {
  padding: 15px;
  border-radius: 8px;
  background-color: #f8f9fa;
  border-left: 4px solid #0366d6;
}

.steps {
  display: flex;
  flex-direction: column;
  gap: 15px;
}

.step {
  display: flex;
  align-items: flex-start;
  gap: 15px;
}

.step-number {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 30px;
  height: 30px;
  border-radius: 50%;
  background-color: #0366d6;
  color: white;
  font-weight: bold;
}

.timeline {
  display: flex;
  flex-direction: column;
  gap: 20px;
}

.timeline-item {
  display: flex;
  gap: 20px;
}

.timeline-date {
  min-width: 120px;
  display: flex;
  flex-direction: column;
  align-items: flex-end;
}

.timeline-content {
  border-left: 2px solid #0366d6;
  padding-left: 20px;
}

.roadmap {
  display: flex;
  flex-direction: column;
  gap: 20px;
}

.roadmap-item {
  display: flex;
  gap: 20px;
}

.roadmap-icon {
  font-size: 24px;
}

.performance-card {
  margin-bottom: 20px;
}

.performance-chart {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.chart-bar {
  background-color: #0366d6;
  color: white;
  padding: 8px;
  border-radius: 4px;
}

.contact-grid {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 20px;
}

.contact-item {
  display: flex;
  flex-direction: column;
  align-items: center;
  text-decoration: none;
  padding: 15px;
  border-radius: 8px;
  background-color: #f8f9fa;
  transition: transform 0.2s;
}

.contact-item:hover {
  transform: translateY(-5px);
}

.tabs {
  display: flex;
  flex-direction: column;
}

.tab {
  overflow: hidden;
}

.tab input {
  display: none;
}

.tab label {
  display: inline-block;
  padding: 10px 20px;
  background-color: #f1f1f1;
  cursor: pointer;
}

.tab input:checked + label {
  background-color: #0366d6;
  color: white;
}

.content {
  display: none;
  padding: 15px;
  background-color: #f8f9fa;
}

.tab input:checked ~ .content {
  display: block;
}
</style>
