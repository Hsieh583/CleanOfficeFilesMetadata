# 需求規格書 v2.0 - Office 檔案中繼資料批次清理工具

## 1. 專案概述

開發一個基於 PowerShell + .NET Windows Forms 的免安裝工具，批次清除 Office 檔案（Word、Excel、PowerPoint）的個人資訊與中繼資料，並還原檔案的原始時間戳記。

---

## 2. 技術架構

| 項目 | 技術選擇 |
|------|---------|
| **開發語言** | PowerShell 5.1+ |
| **UI 框架** | .NET System.Windows.Forms |
| **Office 整合** | COM 物件（Word/Excel/PowerPoint.Application） |
| **部署方式** | 單一 .ps1 檔案 + .bat 啟動器 |

---

## 3. 核心功能需求

### 3.1 使用者介面（雙標籤設計）

#### 標籤 1：設置
- **目標資料夾選擇**
  - 文字輸入框 + 「瀏覽...」按鈕（FolderBrowserDialog）
  - 支援手動輸入或視覺化選擇

- **檔案類型選擇**（多選 CheckBox）
  - ☐ Word（.doc, .docx）
  - ☐ Excel（.xls, .xlsx）
  - ☐ PowerPoint（.ppt, .pptx）
  - 預設：全部勾選

- **中繼資料項目選擇**（多選 CheckBox）
  - ☐ 標題（Title）
  - ☐ 主旨（Subject）
  - ☐ 作者（Author）
  - ☐ 經理（Manager）
  - ☐ 公司（Company）
  - ☐ 類別（Category）
  - ☐ 關鍵字（Keywords）
  - ☐ 備註（Comments）
  - 預設：全部勾選
  - 提供「全選」/「取消全選」快速按鈕

- **處理選項**
  - ☑ 包含子資料夾（預設勾選）

#### 標籤 2：進度
- **進度指示**
  - ProgressBar：視覺化進度條
  - 狀態標籤：顯示「正在處理：[檔名] (n/總數)」

- **執行日誌**
  - TextBox（唯讀、黑底綠字）
  - 實時滾動顯示處理結果
  - 符號標示：[✓] 成功 / [✗] 失敗

#### 底部控制按鈕
- **開始執行**（綠色）：啟動清理流程
- **取消**（紅色）：中止執行
- **打開報告**（藍色）：開啟 CSV 報告

---

### 3.2 核心處理邏輯

#### 檔案掃描
```
1. 根據使用者選擇的檔案類型，建立搜尋模式
2. 使用 Get-ChildItem 遞迴搜尋指定資料夾
3. 彙總所有符合條件的檔案
```

#### 清理流程（每個檔案）
```
Step 1: 檔案存取檢查
  └─ 嘗試以獨占模式開啟檔案
  └─ 若被鎖定，跳過並記錄

Step 2: 保存原始時間戳記
  ├─ CreationTime（建立時間）
  └─ LastWriteTime（修改時間）

Step 3: 初始化 Office 應用
  ├─ 建立對應的 COM 物件（Word/Excel/PowerPoint.Application）
  ├─ 設定 Visible = false（背景執行）
  ├─ 設定 ScreenUpdating = false（加速處理）
  └─ 設定 DisplayAlerts = 0（無對話框）

Step 4: 開啟並處理文件
  ├─ Documents.Open() 開啟檔案
  ├─ 迴圈清除使用者選定的 BuiltInDocumentProperties
  ├─ 執行 RemoveDocumentInformation(1) 移除隱藏資訊
  └─ 存檔後關閉文件

Step 5: 還原時間戳記
  ├─ 設定 CreationTime = [原始建立時間]
  └─ 設定 LastWriteTime = [原始修改時間]

Step 6: 結果記錄
  └─ 成功：記入 ProcessedFiles
  └─ 失敗：記入 FailedFiles（含錯誤訊息）
```

#### 異常處理原則
- 單個檔案失敗不影響其他檔案
- 所有異常捕捉並記錄於日誌
- 檔案鎖定、權限不足、損毀檔案均跳過並標記

---

### 3.3 稽核報告

**檔案格式：** CSV（UTF8 編碼）

**命名規則：** `Cleanup_Report_YYYYMMDD_HHMMSS.csv`

**欄位結構：**
```csv
檔案路徑,處理狀態,錯誤訊息
"C:\path\file.docx",成功,
"C:\path\file2.xlsx",失敗,"檔案被鎖定"
```

**生成時機：** 執行完成後自動生成於當前目錄

---

## 4. 效能要求

| 指標 | 目標值 |
|------|-------|
| UI 啟動時間 | < 2 秒 |
| 檔案處理速度 | 5-10 檔案/秒 |
| 內存使用 | < 100 MB |
| COM 物件重複使用 | 同類型檔案共用一個實例 |

---

## 5. 安全性規範

### 5.1 權限控制
- 僅使用當前使用者權限
- 不修改系統設定或登錄檔
- 不寫入系統目錄

### 5.2 資源清理
- 使用 Try-Catch-Finally 確保資源釋放
- 程式結束時強制清理 Office 進程
- COM 物件使用 ReleaseComObject 釋放

### 5.3 資料保護
- 不修改文件內容，僅清除 metadata
- 還原原始時間戳記，避免留下修改痕跡
- 無臨時檔案，直接修改原檔案

---

## 6. 邊界情況處理

| 情況 | 處理策略 |
|------|---------|
| **檔案被鎖定** | 跳過，記錄於報告 |
| **檔案損毀** | 捕捉異常，記錄錯誤訊息 |
| **無效路徑** | 彈出 MessageBox 警告 |
| **未安裝 Office** | 啟動時檢查，彈出錯誤訊息 |
| **使用者中止** | 立即停止迴圈，生成部分報告 |
| **無檔案符合** | 顯示「未找到檔案」訊息 |
| **未選擇任何項目** | 驗證輸入，彈出警告 |

---

## 7. 使用者體驗原則

### 7.1 簡潔性
- 雙擊 .bat 即可啟動，無需 PowerShell 命令
- 雙標籤設計分離設置和執行，避免混淆
- 預設值合理（全選 + 遞迴），適合大多數場景

### 7.2 即時反饋
- 進度條準確反映處理進度（n/總數）
- 日誌實時輸出，彩色標示成功/失敗
- 按鈕狀態動態變化（執行中禁用設置）

### 7.3 容錯性
- 檔案失敗不中斷整體流程
- 提供取消按鈕，可隨時中止
- 完整的錯誤訊息幫助問題診斷

---

## 8. 部署規格

### 8.1 交付檔案

```
OfficeClear/
├── OfficeClear_Tool.ps1      # 主程式（730 行）
├── OfficeClear.bat            # 啟動器
├── README.md                  # 使用手冊
├── 快速開始.md                # 快速入門
├── 技術文檔.md                # 開發者文檔
└── CHANGELOG.md               # 版本記錄
```

### 8.2 系統需求

| 項目 | 要求 |
|------|------|
| 作業系統 | Windows 10/11 |
| PowerShell | v5.1+ |
| .NET Framework | 4.5+（Windows 內建） |
| Microsoft Office | 2013+ |
| 磁碟空間 | < 1 MB（腳本本身） |

### 8.3 啟動方式

**方式 1（推薦）：**
```
雙擊 OfficeClear.bat
```

**方式 2：**
```powershell
powershell -ExecutionPolicy Bypass -File ".\OfficeClear_Tool.ps1"
```

---

## 9. 驗收標準

### 9.1 功能驗收
- [x] UI 正常顯示，所有控制項可操作
- [x] 可選擇任意組合的 metadata 項目
- [x] 可選擇任意組合的檔案類型
- [x] 進度條和日誌實時更新
- [x] CSV 報告正確生成
- [x] 時間戳記成功還原
- [x] 取消功能正常運作

### 9.2 品質驗收
- [x] PowerShell 語法檢查通過
- [x] 無內存泄漏（COM 物件正確釋放）
- [x] 異常處理完整（Try-Catch-Finally）
- [x] 錯誤訊息清晰易懂

### 9.3 文檔驗收
- [x] README.md 包含完整使用說明
- [x] 快速開始.md 提供 30 秒入門
- [x] 技術文檔.md 包含架構和 API 說明
- [x] CHANGELOG.md 記錄版本更新

---

## 10. 版本資訊

- **版本號：** 2.0
- **發布日期：** 2024 年
- **開發工具：** PowerShell ISE / VS Code
- **測試環境：** Windows 10/11 + Office 2016/2019/365

---

## 11. 未來擴展方向（v2.1+）

- [ ] 支援更多 Office 應用（Access、OneNote）
- [ ] 配置檔保存常用設定
- [ ] 排程任務自動執行
- [ ] 批次檔案比對（清理前後對照）
- [ ] 多語言介面支援
- [ ] 黑暗模式 UI

---

**文檔狀態：✅ 最終定案**  
**最後更新：2024 年**
