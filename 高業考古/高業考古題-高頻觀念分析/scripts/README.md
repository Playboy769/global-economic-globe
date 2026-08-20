# 重複題規律分析管線

產生 `../專題_重複題年度規律分析.md` 所依據的全部統計。

## 重跑

```bash
python run_all.py
```

需要 `pymupdf`（其餘都是標準函式庫）。全程約 40 秒，輸出寫到 `out/`。

## 新的一屆考完之後

1. 把該屆 PDF 丟進 `高業考古/<年>/`。檔名隨意——年份與次數是從**考卷抬頭**（「115 年第3 次…」）讀的，不是從檔名。答案卡若命名為 `<同名>a.pdf` 會自動配對。
2. **編輯 `s2_cluster.py` 的 `EXAMS` 產生器**，把該年的場次數加進去。漏了這步，該屆會被解析但不會進入時間軸，統計等於沒更新。
3. `python run_all.py`。
4. 檢查 `out/parse_report.txt`：`blocks with !=50 questions` 應為 0。不為 0 表示切題規則沒吃到新版面，先修 `s1_extract.py` 的 `split_questions()` 再往下走。

## 檔案

| 腳本 | 做什麼 | 產出 |
|---|---|---|
| `s1_extract.py` | PDF 抽文字 → 切三科 → 切 50 題 → 題幹正規化 → 對答案卡 | `corpus.json`、`parse_report.txt` |
| `s2_cluster.py` | 相似度 ≥90% 分群；間隔分布 vs 蒙地卡羅虛無模型；屆次配對矩陣 | `analysis.json`、`analysis_report.txt` |
| `s3_pattern.py` | 同年/跨年控制變因對照；回收來源距離；經驗命中率初版 | `analysis2_report.txt` |
| `s4_calibrate.py` | 固定 8 屆視窗（去除機械效應）；含單次題的無偏命中率；回測；下一屆預測 | `analysis3.json`、`analysis3_report.txt` |
| `s5_options.py` | 題幹重複時選項是否也相同；套用回測偏誤校正 | `analysis4_report.txt` |
| `s6_mockexam.py` | 依命中機率加權隨機抽樣，產生下一屆模擬考卷 | `../115-3_模擬考卷.html` |
| `s7_allofabove.py` | 「以上皆是／皆非」選項的命中率（含偵測器覆核檔與反例清單） | `allofabove_report.txt`、`allofabove.json` |
| `s8_optionform.py` | 選項長度／絕對詞緩和詞／數字題中間值三項技巧 | `optionform_report.txt`、`_words_audit.txt` |
| `s9_penalty.py` | 罰則題「選最重」命中率（嚴重度排序為人工判定，含排除理由） | `penalty_report.txt`、`penalty.json` |

## 三個容易踩的地方

- **比對門檻**在 `s2_cluster.py` 的 `THRESH`（目前 0.90）。改成 `1.0` 就是「題幹一字不差」的嚴格口徑，重複率會從 73.0% 降到 61.1%。改了門檻後 `s3`–`s5` 都要重跑。
- **chart-tools 的「圖說」欄位會跨工作階段殘留。** 產圖前務必檢查並改掉——2026-08-13 發生過三張圖全部帶著上一個專案（ALAB 法說會）的圖說、就這樣被貼進報告，而且前一份專題的三張圖也同時中招。貼進 md 之後一定要把圖說撈出來逐條看過再定稿：

  ```bash
  grep -o 'font-size:12px;color:#666[^<]*<[^>]*>[^<]*' 專題_*.md
  ```
- **偏誤校正數字不可寫死。** `s5_options.py` 的 `BIAS`／`MAE`／`RAW` 是從 `analysis3.json` 讀 `s4` 當次回測的結果。若改成硬編碼，下次重跑會拿舊週期的校正值去修正新資料。
