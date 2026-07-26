# RR5 UserForm 製作指引 — frmTransaction

這份指引帶你在 Excel VBA Editor 內，手動建立一個「新增交易」表單，並透過 `Ctrl+Shift+T` 或按鈕呼叫。表單會把欄位自動寫入 Transactions sheet。

---

## 0. 先決條件

1. 開啟 `Compound RR5 Derivatives.xlsm`
2. 啟用巨集（首次開檔點「啟用內容」）
3. **匯入 RR5_Core.bas（如尚未做）**：
   - 按 `Alt+F11` → File → Import File → 選 `RR5_Core.bas` → OK
4. 啟用 Excel 黑色主題（建議）：
   - File → Account → Office Theme → **黑色**

---

## 1. 建立 UserForm

1. 按 `Alt+F11` 進入 VBA Editor
2. 在左側 Project Explorer 找到 `VBAProject (Compound RR5 Derivatives.xlsm)`
3. 右鍵 → **Insert → UserForm**
4. 新增一個 `UserForm1`，會自動跳出表單設計畫面 + Toolbox
5. **在右下 Properties 視窗：**
   - `(Name)` 改為 **`frmTransaction`**
   - `Caption` 改為 `新增交易 — RR5 Transactions`
   - `Width` 設 `420`
   - `Height` 設 `580`
   - `BackColor` 可改為 `&H00141414&`（深灰）

---

## 2. 放置控制項（41 個）

從 Toolbox 拖曳對應控制項到表單，**每個拖好後在 Properties 立刻改 Name**。
建議分三欄：左欄、右欄、底部。

### 左欄（X 約 10-200，依序往下）

| Name | 控制類型 | Caption / 預設值 |
|------|---------|------------------|
| (Label1) | Label | `Date:` |
| **txtDate** | TextBox | （UserForm_Initialize 會填今日） |
| (Label2) | Label | `Ticker:` |
| **txtTicker** | TextBox | |
| (Label3) | Label | `Action:` |
| **cmbAction** | ComboBox | （程式碼自動填下拉） |
| (Label4) | Label | `Type:` |
| **cmbType** | ComboBox | |
| (Label5) | Label | `Strike:` |
| **txtStrike** | TextBox | |
| (Label6) | Label | `Expiry:` |
| **txtExpiry** | TextBox | |

### 右欄（X 約 220-410）

| Name | 控制類型 | Caption / 預設值 |
|------|---------|------------------|
| (Label7) | Label | `Shares/Contracts:` |
| **txtShares** | TextBox | |
| (Label8) | Label | `Multiplier:` |
| **txtMultiplier** | TextBox | （預設 1） |
| (Label9) | Label | `Price:` |
| **txtPrice** | TextBox | |
| (Label10) | Label | `Fee:` |
| **txtFee** | TextBox | （預設 0） |
| (Label11) | Label | `Tax:` |
| **txtTax** | TextBox | （預設 0） |
| (Label12) | Label | `Net Amount (自動):` |
| **txtNetAmount** | TextBox | **Properties: `Locked = True`、`BackColor = &H00404040&`（深灰唯讀）** |

### 底部（Y 約 350-510）

| Name | 控制類型 | Caption / 預設值 |
|------|---------|------------------|
| (Label13) | Label | `Sector:` |
| **txtSector** | TextBox | |
| (Label14) | Label | `Target:` |
| **txtTarget** | TextBox | |
| (Label15) | Label | `Strategy Tag:` |
| **cmbStrategy** | ComboBox | |
| (Label16) | Label | `IV at Entry:` |
| **txtIV** | TextBox | |
| (Label17) | Label | `Delta at Entry:` |
| **txtDelta** | TextBox | |
| (Label18) | Label | `Beta:` |
| **txtBeta** | TextBox | |
| (Label19) | Label | `Broker:` |
| **cmbBroker** | ComboBox | |
| (Label20) | Label | `Underlying:` |
| **txtUnderlying** | TextBox | （選擇權/權證填標的名稱，如 TWII、SPX、2330） |

> **txtUnderlying** 對應 Transactions col U（第 21 欄）。填入後 `RebuildPositionsFromTransactions` 會自動寫到 RR5 Options J（UNDERLYING）與 Warrants I（UNDERLYING）欄。期貨可留空。

### 三顆按鈕（最底部，Y 約 510）

| Name | 控制類型 | Caption | 額外設定 |
|------|---------|---------|---------|
| **btnSave** | CommandButton | `Save` | `Default = True`、`BackColor = &H00006400&`（深綠）、`ForeColor = &H00FFFFFF&` |
| **btnClear** | CommandButton | `Clear` | |
| **btnCancel** | CommandButton | `Cancel` | `Cancel = True` |

---

## 3. 貼入程式碼

1. 在表單設計畫面 **雙擊空白處**（不是任何控制項）→ 進入該 UserForm 的 code window
2. 全選預設的 stub 程式碼 → 刪除
3. 開啟同資料夾的 `frmTransaction_Code.txt`，全選複製
4. 貼到 code window
5. 按 `Ctrl+S` 儲存（會問你要存 .xlsm 還是其他格式，選 `Excel Macro-Enabled Workbook (*.xlsm)`）

---

## 4. 設定快捷鍵與工作表按鈕

### 選項 A：用 Ctrl+Shift+T 快捷鍵

1. 在 Project Explorer 找到 `ThisWorkbook`（在 `Microsoft Excel Objects` 下）
2. 雙擊打開 → 貼入：

```vba
Private Sub Workbook_Open()
    Call RR5_Core.ExpiryAlert
    Call RR5_Core.RefreshGreeks
    Call RR5_Core.AutoRefresh_Start
    Application.OnKey "^+t", "RR5_Core.ShowTransactionForm"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call RR5_Core.AutoRefresh_Stop
    Application.OnKey "^+t"
End Sub
```

3. 關閉 xlsm 再開一次 → 按 `Ctrl+Shift+T` 即跳出表單

### 選項 B：在 sheet 上放一顆按鈕

1. 回到 Excel → 開啟 `Transactions` sheet（或 `RR5` 主表）
2. Developer → Insert → **Form Control** 區的「按鈕」（Button）
3. 在 sheet 拉一個按鈕 → 跳出「指派巨集」對話框
4. 選 `RR5_Core.ShowTransactionForm` → 確定
5. 右鍵按鈕 → 編輯文字：`+ Add Transaction`

> 若沒看到 Developer 索引標籤：File → Options → Customize Ribbon → 勾選 Developer

---

## 5. 驗證流程

1. 按 `Ctrl+Shift+T`（或按按鈕）
2. ✅ 表單彈出、日期預設今日、Fee=0、Tax=0、Multiplier=1
3. 輸入：
   - Ticker=`MU`
   - Action=`Buy`
   - Type=`Equity`
   - Shares=`10`
   - Price=`120`
   - Multiplier=`1`
4. ✅ Net Amount 應自動顯示 `1,200.00`
5. 按 `Save` → ✅ 跳訊息確認、Transactions sheet 新增一列
6. 開啟 Transactions sheet 確認：
   - 第 1 欄 Transaction_ID 格式：`T-20260524-153012`
   - 第 2 欄 Date 為今日
   - 第 3 欄 `MU`、第 4 欄 `Buy`、第 5 欄 `Equity`
   - 第 13 欄 Net_Amount = `1200`

---

## 常見問題

| 問題 | 解法 |
|------|------|
| 按 Ctrl+Shift+T 沒反應 | 確認 Workbook_Open 已執行過（關檔重開）；或手動執行 `Application.OnKey "^+t", "RR5_Core.ShowTransactionForm"` |
| 表單跳「frmTransaction UserForm not found」 | UserForm 名稱拼錯，需精確為 `frmTransaction` |
| 按 Save 跳「Type mismatch」 | 數字欄位（Strike/Shares/Price）輸入了非數字。Strike/Expiry/Target/IV/Delta/Beta 可留空 |
| ComboBox 沒下拉箭頭 | Properties 設定 `Style = 0 - fmStyleDropDownCombo` |
| Net Amount 沒自動更新 | 確認 `txtPrice`、`txtShares`、`txtMultiplier`、`txtFee`、`txtTax` 名稱正確（程式碼有監聽這些） |
| 按 Cancel 沒關閉 | 確認該按鈕 Name 是 `btnCancel`，且 Cancel 屬性設 True |

---

## 視覺位置建議（420×580 表單）

```
┌──────────────────────────────────────────────────┐
│ 新增交易 — RR5 Transactions                       │
├────────────────────────┬─────────────────────────┤
│ Date:                  │ Shares/Contracts:        │
│ [txtDate         ]     │ [txtShares          ]    │
│ Ticker:                │ Multiplier:              │
│ [txtTicker       ]     │ [txtMultiplier      ]    │
│ Action:                │ Price:                   │
│ [cmbAction ▼     ]     │ [txtPrice           ]    │
│ Type:                  │ Fee:                     │
│ [cmbType ▼       ]     │ [txtFee             ]    │
│ Strike:                │ Tax:                     │
│ [txtStrike       ]     │ [txtTax             ]    │
│ Expiry:                │ Net Amount (自動):       │
│ [txtExpiry       ]     │ [txtNetAmount ////  ]    │
├────────────────────────┴─────────────────────────┤
│ Sector:        [txtSector            ]            │
│ Target:        [txtTarget            ]            │
│ Strategy Tag:  [cmbStrategy ▼        ]            │
│ IV at Entry:   [txtIV                ]            │
│ Delta:         [txtDelta             ]            │
│ Beta:          [txtBeta              ]            │
│ Broker:        [cmbBroker ▼          ]            │
├──────────────────────────────────────────────────┤
│       [  Save  ]   [  Clear  ]   [  Cancel  ]    │
└──────────────────────────────────────────────────┘
```
