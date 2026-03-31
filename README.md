# 案件結案流程自動化

自動完成品檢組的案件結案流程，包含 SharePoint 操作和本機網路磁碟操作。

## 前置條件

1. 安裝 **Antigravity AI**（VS Code 或獨立版本）
2. 電腦需在公司內網（可存取 `\\192.168.44.100\ts-qa\品檢組\`）
3. 擁有 SharePoint 帳號（IGS 帳號）

## 使用方式

### 步驟 1：開啟專案
用 Antigravity AI 開啟此資料夾。

### 步驟 2：登入 SharePoint
在瀏覽器中開啟 [SharePoint 測試交接](https://internationalgamessystem.sharepoint.com/teams/OLD_TS_QA/DocLib/Forms/AllItems.aspx?id=%2Fteams%2FOLD%5FTS%5FQA%2FDocLib%2F%E6%B8%AC%E8%A9%A6%E4%BA%A4%E6%8E%A5&viewid=1b2aa336%2Debf1%2D4209%2D8a1c%2D6d9f75023930)，完成帳號、密碼及 MFA 驗證碼認證。

### 步驟 3：執行結案
在 Antigravity AI 聊天中輸入：
```
/close-case <單號> <案件名稱>
```

**範例：**
```
/close-case 235927 滿貫大亨_活動序號兌換測試
```

AI 會自動完成以下操作：

| 步驟 | 說明 |
|------|------|
| SharePoint 測試交接 | 檔案加 `_DONE`，資料夾移至已結案 |
| SharePoint 線上文件 | 找到對應檔案，移至已結案/YYYYMM |
| 本機 0_GT測試交接 | 資料夾移至已結案 |
| 本機圖影片資料存放區 | 建立 GT\年\月\MMDD\案件名稱\問題圖檔 |

## 檔案說明

| 檔案 | 用途 |
|------|------|
| `Close-Case.ps1` | 本機操作腳本（步驟 5-6） |
| `config.json` | 路徑設定（所有腳本共用） |
| `.agents/skills/close-case/SKILL.md` | AI 自動化流程定義 |
| `.agents/workflows/close-case.md` | `/close-case` 指令定義 |

## 注意事項

- 案件單號必須為 **6 位數字**
- 執行前請確認 SharePoint 已登入
- 本機腳本需要網路磁碟存取權限
- 所有檔案使用 **UTF-8 BOM** 編碼
- `config.json` 為所有腳本的共用設定，路徑修改只需改此處
