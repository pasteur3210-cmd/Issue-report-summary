# CHANGELOG

## V1.2.2 - 2026-07-16

### Sorting
- 新增整份異常履歷依 `報告日期` 由新到舊排序。
- 既有資料與本次新增資料一起重新排序。
- 空白 / 無法辨識日期移至最後。
- 排序後重新編號 No. 1~N。
- 新增 Sort Verification Log。

### GUI / Build
- 沿用 OLT MA5800X2 Integrated Tool 的 `assets/HD.ico`。
- GUI Window Icon 使用 HD.ico。
- PyInstaller EXE Icon 使用 HD.ico。
- build.spec 將 HD.ico 打包至 assets。
- 新增 PyInstaller `_MEIPASS` runtime resource path 支援。

### Logging / Verification
- test.log 新增：
  - SORT_CHECK | 報告日期 | DESC
  - SORT_CHECK | Blank Date At End
  - RENUMBER_CHECK | No. 1~N
- verification_report.md 新增排序結果。

### Retained
- V1.2.1 Table-First Parsing
- Fixed Column Width / Row Height
- Batch Folder Import
- Duplicate Skip
- Illegal Character Sanitization
- Existing Excel Schema Migration
