# PDF 支援說明

## 概述

從 v2.1 版本開始，Office 檔案中繼資料清理工具支援 PDF 檔案的中繼資料操作。

## 功能支援

### 讀取模式（Read Mode）
- ✅ **完整支援**：可使用系統內建功能讀取 PDF 中繼資料
- 📝 支援的屬性：Title、Subject、Author、Keywords
- 🚀 無需額外套件

### 清理模式（Clean Mode）與替換模式（Replace Mode）
- ⚠️ **需要 iTextSharp**：修改 PDF 中繼資料需要 iTextSharp 函式庫
- 📝 支援的屬性：Title、Subject、Author、Keywords、Creator、Producer
- 📦 需要下載並安裝 iTextSharp.dll

## iTextSharp 安裝指南

### 方法 1：下載預編譯的 DLL（推薦）

1. 訪問 [iTextSharp NuGet 頁面](https://www.nuget.org/packages/iTextSharp/)
2. 下載 iTextSharp 5.5.x 系列版本（推薦 5.5.13.3）
3. 從 NuGet 套件中提取 `itextsharp.dll`
4. 將 `itextsharp.dll` 複製到與 `OfficeClear_Tool.ps1` 相同的目錄

### 方法 2：使用 NuGet 套件管理器

如果您有 Visual Studio 或 .NET 開發環境：

```bash
Install-Package iTextSharp -Version 5.5.13.3
```

然後從套件目錄中複製 `itextsharp.dll`：
```
packages\iTextSharp.5.5.13.3\lib\itextsharp.dll
```

### 方法 3：手動下載

1. 前往 iTextSharp GitHub 發布頁面
2. 下載適合的版本
3. 將 `itextsharp.dll` 放在工具目錄

## 檔案結構

```
工作目錄/
├── OfficeClear.bat
├── OfficeClear_Tool.ps1
└── itextsharp.dll          ← 放在此處（用於 PDF 清理/替換）
```

## 使用說明

### 讀取 PDF 中繼資料（無需 iTextSharp）

1. 勾選「PDF (.pdf)」檔案類型
2. 選擇「讀取模式」
3. 勾選要查看的中繼資料項目
4. 點擊「開始執行」
5. 查看生成的 Metadata_Report_*.csv

### 清理 PDF 中繼資料（需要 iTextSharp）

1. 確認 `itextsharp.dll` 已放在正確位置
2. 勾選「PDF (.pdf)」檔案類型
3. 選擇「清理模式」
4. 勾選要清除的中繼資料項目
5. 點擊「開始執行」

### 替換 PDF 中繼資料（需要 iTextSharp）

1. 確認 `itextsharp.dll` 已放在正確位置
2. 勾選「PDF (.pdf)」檔案類型
3. 選擇「替換模式」
4. 在替換值輸入框中輸入新值
5. 點擊「開始執行」

## 故障排除

### 錯誤：「PDF cleaning not supported without iTextSharp library」

**原因**：嘗試清理或替換 PDF 中繼資料，但未找到 iTextSharp.dll

**解決方案**：
1. 確認 `itextsharp.dll` 存在於工具目錄
2. 檢查檔案名稱是否正確（必須是小寫的 `itextsharp.dll`）
3. 如果只需要讀取 PDF 中繼資料，切換到「讀取模式」

### 讀取模式使用備援方法

如果未找到 iTextSharp，讀取模式會自動使用 Windows Shell.Application COM 物件作為備援：
- 功能較有限
- 可能無法讀取所有 PDF 中繼資料
- 某些 PDF 格式可能不支援

## 注意事項

⚠️ **重要提醒**：
- PDF 處理比 Office 檔案處理慢
- 某些受保護的 PDF 無法修改中繼資料
- 加密的 PDF 需要先解密
- 建議先使用「讀取模式」確認 PDF 中繼資料
- 始終備份重要的 PDF 檔案

## 授權資訊

- **iTextSharp**：使用 AGPL 授權
- 商業使用可能需要購買授權
- 詳情請參閱 [iText 官方網站](https://itextpdf.com/)

## 版本相容性

| iTextSharp 版本 | 測試狀態 | 備註 |
|----------------|---------|------|
| 5.5.13.3 | ✅ 推薦 | 穩定版本 |
| 5.5.x | ✅ 支援 | 較舊但穩定 |
| 7.x+ | ⚠️ 未測試 | API 可能不相容 |

## 支援

如有問題或建議，請透過 GitHub Issues 回報。
