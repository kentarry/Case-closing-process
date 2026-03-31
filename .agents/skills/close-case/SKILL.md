---
name: close-case
description: 案件結案流程自動化
---

# 結案流程

## 觸發
使用者輸入 6 碼單號 + 名稱，如 `123456 測試案件`。
前 6 碼 = 單號，其餘 = 名稱，YYYYMM = 當前年月。

## 架構
步驟 A：SP 登入 → 步驟 B（browser_subagent）→ 步驟 C（browser_subagent）→ 步驟 D（run_command）+ 步驟 E（browser_subagent）並行 → 完成報告

> [!IMPORTANT]
> 步驟 E 前，若使用者可能未登入 Google，先用 browser_subagent 開啟 GOOGLE_SHEET 確認是否可編輯。若不可編輯，通知使用者登入 Google 帳號後再繼續。

> [!IMPORTANT]
> - run_command 設 SafeToAutoRun=true
> - browser_subagent 直接用下方模板替換 {變數}
> - 結束後關閉瀏覽器，notify_user 報告。不做驗證截圖。

## 步驟 A：SP 登入
開啟 SP_URL_HANDOVER，若需登入 → notify_user 請使用者完成 MFA。

## URL 定義
- SP_URL_HANDOVER = https://internationalgamessystem.sharepoint.com/teams/OLD_TS_QA/DocLib/Forms/AllItems.aspx?id=%2Fteams%2FOLD%5FTS%5FQA%2FDocLib%2F%E6%B8%AC%E8%A9%A6%E4%BA%A4%E6%8E%A5&viewid=1b2aa336%2Debf1%2D4209%2D8a1c%2D6d9f75023930
- SP_URL_ONLINE_DOC = https://internationalgamessystem.sharepoint.com/teams/OLD_TS_QA/DocLib/Forms/AllItems.aspx?id=%2Fteams%2FOLD%5FTS%5FQA%2FDocLib%2F%E6%B8%AC%E8%A9%A6%E4%BA%A4%E6%8E%A5%2F%E7%B7%9A%E4%B8%8A%E6%96%87%E4%BB%B6&viewid=1b2aa336%2Debf1%2D4209%2D8a1c%2D6d9f75023930
- SP_URL_OUTSOURCE = https://internationalgamessystem.sharepoint.com/teams/OLD_TS_QA/DocLib/Forms/AllItems.aspx?id=%2Fteams%2FOLD%5FTS%5FQA%2FDocLib%2F%E6%B8%AC%E8%A9%A6%E4%BA%A4%E6%8E%A5%2F%E7%B7%9A%E4%B8%8A%E6%96%87%E4%BB%B6%2F%E5%A4%96%E5%8C%85%E6%B8%AC%E8%A9%A6&viewid=1b2aa336%2Debf1%2D4209%2D8a1c%2D6d9f75023930
- GOOGLE_SHEET = https://docs.google.com/spreadsheets/d/1h9y5EMPISDJUXae1YRSVIJ-FBW-JvB0spADmzrzcIs8/edit?gid=2067710872#gid=2067710872

## 步驟 B：SP 測試交接（browser_subagent）

```
你已登入 SharePoint。單號={單號}，名稱={名稱}，YYYYMM={YYYYMM}。

=== B1：進入案件資料夾 ===
1. 開啟 SP_URL_HANDOVER
2. 等 2 秒，點 {YYYYMM} 進入
3. 找包含 {單號} 的資料夾，點進去
4. 找不到 → 回上層找「外包」點進去再找
5. 還找不到 → 開啟 SP_URL_OUTSOURCE 再找

=== B2：重新命名檔案 ===
對每個「檔案」（跳過子資料夾）：
- 已含 _DONE → 跳過
- ⚠️ 絕對不要雙擊或點擊檔名文字！
- 在檔案那一行的空白區域（如檔案大小欄或修改日期欄）按右鍵，開啟右鍵選單
- 在右鍵選單中點擊「重新命名」
- 等 1 秒確認進入重新命名模式
- 按 End 到檔名末端 → 輸入 _DONE → 按 Enter
- 等 2 秒確認重新命名完成，再處理下一個
- ⚠️ 如果不慎開啟了檔案，關閉它，回到資料夾，繼續處理

=== B3：移動到已結案 ===
⚠️ 重要：已結案資料夾在 {YYYYMM} 裡面，不是在「測試交接」裡面。
1. 點頁面頂部麵包屑中的「{YYYYMM}」（確認回到 {YYYYMM} 層級，不要多往上跳到測試交接）
2. 點左側圓形選取框選取案件資料夾
3. 點工具列「移動至」
4. 對話框中目前位置應該是 {YYYYMM} 資料夾內，直接在這裡滾到最底部找「已結案」
5. 點擊「已結案」進入
6. 立即點「移到這裡」。不要再點任何子資料夾。
⚠️ 若對話框顯示的不是 {YYYYMM} 層級，手動導覽至：測試交接 → {YYYYMM} → 已結案。

完成後回報結果。
```

## 步驟 C：SP 線上文件（browser_subagent）

```
你已登入 SharePoint。單號={單號}，YYYYMM={YYYYMM}。

=== C1：線上文件 ===
1. 開啟 SP_URL_ONLINE_DOC
2. 找包含 {單號} 的檔案（可能有 3D_ 前綴）
3. 右鍵 → 移動至
4. 對話框找「已結案」→ 點進去 → 找 {YYYYMM} → 點進去
5. 點「移到這裡」
6. 看到通知後結束，不再搜尋移動。

完成後回報結果。
```

## 步驟 D：本機操作（run_command）
使用相對路徑，在專案資料夾中執行：
```powershell
powershell -ExecutionPolicy Bypass -File ".\Close-Case.ps1" -CaseNumber "<單號>" -CaseDescription "<名稱>"
```
Cwd 設為此專案的根資料夾路徑。

## 步驟 E：Google Sheet（browser_subagent）

```
單號={單號}。

=== E：Google Sheet 結案日 ===
1. 開啟 GOOGLE_SHEET
2. 等 5 秒載入（第一次可能需要更久）
3. 檢查頁面狀態：
   - 若出現 Google 登入頁面 → 回報 NEEDS_GOOGLE_LOGIN，請使用者登入 Google 帳號
   - 若顯示「僅供檢視」→ 回報 NEEDS_GOOGLE_LOGIN，請使用者登入有編輯權限的 Google 帳號
   - 若可編輯 → 繼續下一步
4. 確認在「案件反饋」頁籤
5. Ctrl+F 搜尋 {單號}
6. 點搜尋結果儲存格（C 欄）
7. Esc 關閉搜尋框
8. 右方向鍵 2 次到 E 欄
9. 輸入 {今天日期} 然後按 Enter

⛔ STOP — 任務完成 ⛔
按下 Enter 後，你的任務已 100% 完成。
- 禁止再點擊、修改、刪除、或查看任何儲存格
- 禁止驗證剛才輸入的日期是否正確
- 禁止做任何截圖
- 立即回報「Google Sheet 結案日填寫完成」並結束
```

## 完成報告
| 步驟 | 狀態 | 說明 |
|------|------|------|
| SP 測試交接 | ✅/❌ | 改名+移至已結案 |
| SP 線上文件 | ✅/❌ | 移至已結案/{YYYYMM} |
| 本機操作 | ✅/❌ | 移動+圖檔複製 |
| Google Sheet | ✅/❌ | 結案日填寫 |
