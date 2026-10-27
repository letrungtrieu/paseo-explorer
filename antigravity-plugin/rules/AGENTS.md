# Paseo Office Explorer — Reporting & Highlighting Guide

When inspecting or reporting issues, missing values, formula errors, or highlights in Office files (`.xlsx`, `.xls`, `.csv`, `.docx`, `.doc`, `.pdf`), always include clickable Office Explorer references so the user can click them in Paseo Chat to open the file in a new Workspace Tab and jump directly to the highlighted cell, row, column, range, or page:

## 1. Inline Interactive Link Tag (Recommended)

Format: `[[<workspace-relative-path>#<SheetName>!<Coordinate> | <Issue Description>]]`

- **Single Cell (`E3`)**:
  `[[examples/Ngân sách hàng tháng.xlsx#Tóm tắt!E3 | Ô E3 bị để trống]]`
  *(If omitting sheet name, the first sheet opens automatically: `[[examples/Ngân sách hàng tháng.xlsx#E3 | Ô E3 bị để trống]]`)*
- **Entire Row (`Row:5` or `Dòng 5`)**:
  `[[examples/Ngân sách hàng tháng.xlsx#Giao dịch!Row:5 | Dòng 5 thiếu chứng từ]]`
- **Entire Column (`Col:E` or `Cột E`)**:
  `[[examples/Ngân sách hàng tháng.xlsx#Tóm tắt!Col:E | Cột E sai định dạng số]]`
- **Cell Range (`B4:D8`)**:
  `[[examples/Ngân sách hàng tháng.xlsx#Tóm tắt!B4:D8 | Vùng B4:D8 lệch tổng cộng]]`
- **PDF / Word Page (`Trang 2` or `Page 2`)**:
  `[[docs/Bill_of_Lading.pdf#Trang 2 | Thiếu chữ ký xác nhận ở trang 2]]`

## 2. Natural-Language Sentences (Auto-Detected in Chat)

You can also write natural Vietnamese or English sentences containing the coordinate and quoted file path:

- `E3 trong file "examples/Ngân sách hàng tháng.xlsx" lỗi để trống`
- `Dòng 5 trong file "examples/Ngân sách hàng tháng.xlsx" thiếu mã HS`
- `Cột E trong file "examples/Ngân sách hàng tháng.xlsx" sai công thức`
- `Vùng B4:D8 trong file "examples/Ngân sách hàng tháng.xlsx" lệch số liệu`
