# Feature: UserForm Bulk Import & Validation (feature/userform)

Mục tiêu: Thêm UserForm để import CSV, validate, preview và apply bulk add/update/delete; export template; tạo workbook mẫu .xlsm; tách logic thành modules/class; ready để push vào branch feature/userform.

Các file đã chuẩn bị:
- BulkOperations.bas         (Module: core logic API)
- CreateWorkbook.bas         (Module: tạo workbook mẫu IT-Inventory.xlsm)
- ufmBulkImport.frm          (UserForm code — tạo UserForm tên ufmBulkImport trong VBA)
- cAssetManager.cls          (Class module để tách logic thao tác sheet)
- bulk_template.csv          (Template mẫu)
- INSTRUCTIONS.md            (hướng dẫn này)

Cách deploy vào workbook .xlsm (bước nhanh):
1) Tạo workbook .xlsm:
   - Cách nhanh nhất: mở Excel -> Alt+F11 -> Immediate Window -> paste and run:
     Application.Run "CreateITInventoryWorkbook"
   - Hoặc chạy thủ công: mở VBA, tạo Module, dán CreateWorkbook.bas, chạy Sub CreateITInventoryWorkbook.

2) Import modules & class:
   - Alt+F11 -> Project Explorer -> chuột phải vào VBAProject (IT-Inventory.xlsm) -> Import File... -> chọn BulkOperations.bas, CreateWorkbook.bas, cAssetManager.cls, ufmBulkImport.frm.
   - Tạo một UserForm mới (Insert -> UserForm) đặt tên là ufmBulkImport, tạo các controls như trong header và paste code ufmBulkImport.frm vào cửa sổ code của UserForm.

3) Chạy CreateMenuSheet (nếu muốn):
   - Bạn có thể thêm macro tạo menu (mã cũ CreateMenuSheet) hoặc tạo nút từ ribbon để gọi UserForm.
   - Để test nhanh: trong Immediate Window gõ:
     ufmBulkImport.Show

4) Thử workflow:
   - Export template: bấm nút "Export Template" hoặc gọi ExportTemplateCSV với đường dẫn.
   - Import CSV: Load CSV vào sheet Bulk Form qua UserForm hoặc thủ công.
   - Validate: Click Validate (sẽ tạo sheet 'Bulk Errors' nếu có vấn đề).
   - Apply Add/Update: Click "Apply Add/Update".
   - Apply Delete: Click "Apply Delete".
   - Kiểm tra sheet "Asset Movement History" để xem log thay đổi.

5) Git / Repository:
   - Nếu repo rỗng, tạo initial commit (ví dụ README) trước để có branch main.
   - Tạo branch feature/userform (local) hoặc trên GitHub:
     git checkout -b feature/userform
   - Thêm file text (BulkOperations.bas, CreateWorkbook.bas, cAssetManager.cls, ufmBulkImport.frm, INSTRUCTIONS.md, bulk_template.csv)
     git add BulkOperations.bas CreateWorkbook.bas cAssetManager.cls ufmBulkImport.frm INSTRUCTIONS.md bulk_template.csv
     git commit -m "Feature: add UserForm import/validate, modules and template"
     git push -u origin feature/userform
   - Tạo PR để review & merge.

6) Ghi chú an toàn & kiểm thử:
   - Luôn thử trên bản sao (copy) workbook trước khi chạy trên dữ liệu thật.
   - Kiểm tra kỹ sheet headers — các header phải chính xác như đã định nghĩa.
   - Bạn có thể mở rộng ValidateBulkSheet để kiểm tra định dạng ID, duplicate ID, hoặc mapping tên header linh hoạt.

Nếu bạn muốn, mình có thể:
- Push trực tiếp lên branch feature/userform và tạo PR nếu bạn chấp nhận hộp thoại xác nhận GitHub mà mình đã gửi.
- Hoặc bạn làm theo lệnh Git ở trên và mình hỗ trợ review/merge PR.
