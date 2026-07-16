# 異常報告履歷擷取工具 V1.2.2

## V1.2.2 主要更新

### 報告日期排序
每次新增或更新履歷後，會將「既有資料 + 本次新增資料」全部依 `報告日期` 重新排序：
1. 最新日期在最上方
2. 最舊日期在下方
3. 空白或無法辨識日期統一置底
4. 排序完成後重新編號 `No.` 為 1~N

`test.log` 會新增：
- `SORT_CHECK | 報告日期 | DESC | PASS/FAIL`
- `SORT_CHECK | Blank Date At End | PASS/FAIL`
- `RENUMBER_CHECK | No. 1~N | PASS/FAIL`

### GUI / EXE 圖標
沿用 OLT MA5800X2 Integrated Tool 的 `assets/HD.ico`：
- Windows EXE 檔案圖標
- GUI 視窗左上角圖標
- PyInstaller source/EXE runtime 資源路徑均支援

### V1.2.1 功能全部保留
- Table-First 表格結構優先解析
- 固定欄寬 / 固定列高
- 移除測試數量 / 不良數量
- 保留不良率
- 批量資料夾匯入
- 包含子資料夾
- 重複報告 SKIP
- Illegal Character 清理
- Existing Excel Schema Migration
- run.log / test.log / debug.log / error.log / verification_report.md

## GitHub 編譯
解壓縮後，將所有檔案與 `.github`、`assets` 資料夾完整上傳到 GitHub Repository 根目錄。

必要結構：

```
main.py
build.spec
requirements.txt
README.md
CHANGELOG.md
TEST_REPORT.md
assets/
  HD.ico
.github/
  workflows/
    build.yml
```

進入 GitHub Actions 執行 Build Windows EXE。
