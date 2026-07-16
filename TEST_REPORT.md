# TEST_REPORT V1.2.2

## 查核項目

### Date Sort
1. 既有 + 新增資料一起排序。
2. 有效報告日期依 DESC 排序。
3. 空白 / 無效日期置底。
4. 排序後 No. 重新編號 1~N。
5. test.log 寫入 Sort Check。
6. verification_report.md 寫入 Sort Result。

### Icon
7. assets/HD.ico 存在。
8. build.spec datas 包含 assets/HD.ico。
9. build.spec EXE icon 指向 assets/HD.ico。
10. GUI 使用 resource_path() 載入 HD.ico。
11. PyInstaller `_MEIPASS` runtime path 支援。

### Regression
12. Table-First Parser 保留。
13. Batch Folder Import 保留。
14. Fixed Column Width / Row Height 保留。
15. Illegal Character 清理保留。
16. Duplicate Report SKIP 保留。
17. Existing Excel Schema Migration 保留。
18. run/test/debug/error/verification logs 保留。

## Static Verification
- Python syntax check: PASS
- HD.ico found in OLT tool package: PASS
- build.spec icon configuration: PASS
- Package excludes __pycache__ / .pyc: PASS
