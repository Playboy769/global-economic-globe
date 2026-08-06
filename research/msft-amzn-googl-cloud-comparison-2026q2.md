# Azure vs AWS vs Google Cloud — 2026年4–6月季度雲端業務比較

**依據**：Microsoft FY2026 Q4（財年4–6月，對應日曆Q2 2026）、Amazon Q2 FY2026（日曆Q2 2026）、Alphabet FY2026 Q2（日曆Q2 2026）三份法說會逐字稿＋SEC 10-K/10-Q交叉查證，三者財報期間完全對齊同一個日曆季度，可直接並列比較。完整背景與逐項論證請見對應的完整分析報告：
[MSFT FY26Q4](research/msft-analysis-fy26q4/MSFT_FY26Q4_Analysis/MSFT_FY26Q4_Analysis.html) ／
[AMZN FY2026Q2](research/amzn-analysis-2026q2/AMZN_FY2026Q2_Analysis/AMZN_FY2026Q2_Analysis.html) ／
[GOOGL FY26Q2](research/googl-analysis-2026q2/GOOG_FY26Q2_Analysis/GOOG_FY26Q2_Analysis.html)

---

## 一、總覽表（2026日曆Q2，即MSFT FY26財年Q4）

| | **Azure／Intelligent Cloud（MSFT）** | **AWS（AMZN）** | **Google Cloud（GOOGL）** |
|---|---|---|---|
| 分部營收 | $393億（Intelligent Cloud整體）；Azure單獨年增43% | $422.3億 | $248億 |
| YoY成長率 | Intelligent Cloud +32%（CC+31%）；**Azure本身+43%**，下季指引加速至CC~45% | **+36.8%** | **+81.8%**（QoQ +23.7%） |
| 分部營業利益 | 併入Intelligent Cloud，年度營益率41.3%（FY26） | $166.2億 | $88.1億 |
| 分部營業利益率 | 約41%（FY26年度，YoY大致持平） | **39.36%**（YoY+646bp，其中約100–130bp來自一次性能源衍生品公允價值利得） | **35.6%**（QoQ+2.7pp，YoY自20.7%大幅擴張達+1,490bp） |
| 待認列訂單／Backlog | RPO $6,780億（YoY+84%；扣除OpenAI合約仍+25%） | Backlog $4,960億（三位數年增） | 總Backlog約$5,140億（Alphabet層級，Cloud為主要驅動） |
| 資產負債表資本強度 | 伺服器/網路設備毛額年增+62.5%；PP&E淨額+52.7% | AWS分部資產+38.6%、PP&E+38.8%（半年） | 未單獨揭露Cloud分部PP&E，但TPU系統銷售本季首次認列，屬資本強度轉嫁給客戶的新模式 |

⚑ 三者分部定義並不完全對等：MSFT的「Intelligent Cloud」是Azure＋On-prem伺服器產品＋部分企業服務的混合分部，10-K本身不單獨揭露Azure的美元營收，只給成長率；AWS與Google Cloud則是乾淨的單一雲端分部（AWS在10-Q中是審計過的GAAP分部）。比較營收「絕對規模」時，AWS與Google Cloud的數字精確度高於Azure的推算值。

---

## 二、營收成長率與市佔率

### 成長率排序：Google Cloud > AWS > Azure（但基期差異巨大）

- **Google Cloud +81.8% YoY**是三者中最快，且是**加速中**的加速——上季（Q1 2026）年增已經很高，這季QoQ仍再+23.7%，營收半年內從$200億衝到$248億。加速的主因之一是**本季首度開始認列TPU系統銷售營收**——這代表Google Cloud的成長已經不只是雲端服務消費，而是疊加了硬體出貨的一次性/週期性收入，與AWS、Azure純服務型雲端收入的成長性質不完全可比。
- **AWS +36.8% YoY**，是「連續第五季加速」，且Andy Jassy特別強調這是「18季以來最快」的成長——在AWS這種已經是$1,690億年化營收規模的巨型基期上維持並加速這個成長率，其邊際美元增量（單季新增$46億）遠大於Google Cloud的邊際增量，是三者中規模效應最顯著的案例。
- **Azure +43%**（MSFT財年Q4）並指引下季加速到CC約45%，但Intelligent Cloud分部整體只有+32%——這個落差本身就是一個訊號：分部內的非Azure業務（On-prem伺服器產品）正在對沖Azure的加速度，拉低了分部總數。若只看Azure本身，43%與AWS的36.8%相近但略高；但由於MSFT不揭露Azure的美元營收，兩者的絕對規模無法直接比對，只能比成長率。

### 市佔率：無法從本輪三份逐字稿本身得出精確數字，但方向清楚

三份逐字稿與SEC申報都沒有揭露官方市佔率數字（這類數字通常來自第三方研究機構如Synergy Research／Canalys／Gartner，本輪分析未另外抓取這類資料，屬資訊缺口而非「市佔率不重要」）。但從**規模與加速度的交叉關係**可以合理推論方向：

- AWS絕對規模仍是三者中最大（$422億／季），是產業既有龍頭地位的體現；但36.8%的成長率若長期低於Azure（43%）與Google Cloud（81.8%），意味著**市佔率被侵蝕的速度取決於這個成長率落差能維持多久**——這正是AWS法說會Q&A中Brian Nowak追問「2027年capacity能否放緩」被Jassy用「barbell」框架迴避掉的問題核心。
- Google Cloud雖然絕對規模最小（$248億），但成長率是AWS的2.2倍、Azure的近2倍，且是**唯一同時公布「TPU系統對外銷售」這個新收入模式**的一家——如果這個模式持續擴大，Google Cloud的成長軌跡可能不只是「追上」AWS／Azure的雲端服務市佔，還疊加了一塊AWS（Trainium／Graviton自研晶片目前仍以雲端內部消費為主，僅開始探索對外銷售)、Azure（Maia/Cobalt完全未對外銷售）都還沒有的獨立硬體營收線。
- ⚑ 留白反推：三家管理層在法說會上都沒有主動引用任何市佔率數字——這本身是產業慣例（避免公開承認落後或授人以柄），但意味著任何市佔率結論都必須依賴第三方追蹤機構的獨立數據，而非管理層自述，這點在後續若要精確量化市佔率時務必補做。

---

## 三、獲利能力／營業利益率

### 營益率排序：Azure（~41%）> AWS（39.36%）> Google Cloud（35.6%），但**方向完全相反**——Google Cloud的擴張速度最猛

| | 一年前 | 上季 | 本季 | YoY變化 |
|---|---|---|---|---|
| Intelligent Cloud（MSFT，年度值） | 42.0%（FY25） | — | 41.3%（FY26） | −0.7pp（大致持平） |
| AWS | 32.90% | — | 39.36% | **+646bp** |
| Google Cloud | 20.7% | 32.9% | 35.6% | **+1,490bp**（近三倍毛利率擴張） |

- **Azure／Intelligent Cloud的營益率是三者中絕對水準最高，但成長已經見頂、甚至微幅倒退**——這符合一個已經高度規模化、成熟雲端業務的特徵：管理層本季法說會的敘事重心也從「毛利率擴張」轉向「CapEx紀律與效率工作」（Amy Hood反覆強調「短生命週期資產彈性」「效率是持續的grind work」），而非追加毛利率突破的空間。
- **AWS的+646bp擴張中，約100–130bp來自一次性能源衍生品公允價值利得**（$6億能源合約公允價值變動，主要計入AWS分部）——若扣除這個一次性項目，AWS的「乾淨」YoY擴張大約在520bp左右（Brian Olsavsky在法說會上自己主動揭露了這個拆分：「650bp，扣除衍生品利得後是520bp」），仍然是紮實的營運槓桿，但不宜把完整650bp都當作經常性毛利改善。
- **Google Cloud的擴張最誇張、也最年輕**——20.7%到35.6%，一年內營益率幾乎翻倍。這代表Google Cloud相對AWS／Azure仍處於「規模經濟正在兌現」的早期階段，尚未到達邊際毛利率改善趨緩的成熟期；但也正因為基期低，這個擴張速度未來能否延續存在較大不確定性——法說會中Eric Sheridan直接問「TPU在backlog中的佔比與毛利路徑」，管理層對Q3展望的回應是「因第三方過渡產能與Wiz整合仍有溫和毛利壓力」，暗示這個高速擴張軌跡在下一季可能會出現至少短暫的停頓或收斂，而非線性外推。
- ⚑ 三者毛利率的**驅動力本質不同**：Google Cloud本季的毛利率擴張主要來自「Cloud成本成長（+47.7%）明顯低於營收成長（+81.8%）」的乾淨營運槓桿（規模效應）；AWS則同時混合了「效率工作＋產能優化」與「一次性衍生品利得」兩種來源；Azure／Intelligent Cloud的毛利率已經進入「CapEx紀律決定毛利率能否維持」的階段——三者處在雲端毛利率曲線的三個不同位置，不宜用同一個「誰的毛利率管理比較好」的單一框架去比較。

---

## 四、客戶結構與產品差異化

### 客戶結構

- **AWS**：三者中客群最橫向、最不集中——涵蓋Amazon自身電商/物流基礎設施、傳統企業上雲遷移、新創、以及AI實驗室。法說會明確點名Anthropic與OpenAI是Trainium晶片的兩大多年期、多GW承諾客戶——**值得注意的交叉關係是：OpenAI同時是Azure雲端最大客戶（本季Microsoft Cloud總營收中「近90%來自Frontier Model公司以外的客戶」，意味著OpenAI雖非多數但顯然是集中度最高的單一客戶之一），也是AWS Trainium的簽約大戶**——這代表三大AI實驗室級客戶（OpenAI、Anthropic）並未把運算需求集中押注在單一雲端供應商上，而是同時向Azure與AWS下注，這種「AI實驗室多雲分散」的行為本身就是一個對三家雲端龍頭都適用的結構性風險訊號：任一家都無法把AI實驗室當作「自家專屬」的成長支柱。
- **Azure**：企業客群比重最高、涵蓋面最偏向大型跨國企業與政府（法說會列舉NHS England 50.5萬席、KPMG 27.6萬員工、HSBC 20萬席、EY 40萬員工等超大型部署案例），且10-K明確揭露「無任何單一客戶或美國以外國家佔總營收10%以上」——這是三者中唯一有此明確揭露的一家，代表其客戶集中度風險（至少在公司整體層級）是三者中最低、最分散的。
- **Google Cloud**：客戶結構呈現「兩條腿走路」——一是傳統企業AI採用（Gemini Enterprise本季具名客戶PepsiCo、Intel、HSBC、Bell Canada、Macy's、SIGNAL IDUNA，近90%Fortune 100採用率），二是**TPU硬體對外銷售的新客群**，透過Blackstone作為資料中心產能夥伴布局對外部署——這代表Google Cloud的客戶結構正在從「純軟體/雲端服務訂閱」擴張到「硬體＋資料中心產能」的新維度，客群輪廓與另外兩家的差異正在擴大而非收斂。

### 產品差異化

| 面向 | Azure | AWS | Google Cloud |
|---|---|---|---|
| 自研晶片 | Maia（AI，200世代，效能/成本+30%）、Cobalt（CPU）——**完全內部使用，未對外銷售** | Trainium（AI，年化營收已破$250億）、Graviton（CPU，滲透率98%）——**主要供雲端內部消費，法說會首度鬆口「未來有實質機會」對外銷售裸晶片** | TPU（Ironwood世代）——**本季首度認列對外系統銷售營收**，是三者中率先把自研AI晶片轉為獨立硬體營收線的一家 |
| 模型策略 | 「harness與模型解耦」架構哲學——11,000+模型目錄，涵蓋OpenAI／Anthropic／Mistral／xAI＋自家MAI家族；策略核心是「模型可替換」而非押注單一模型 | Bedrock模型市集——同樣強調多供應商選擇（不只Anthropic/OpenAI），並同步自研frontier模型（原因是成本控制與訓練優先權，而非要取代第三方模型） | Gemini（自家frontier模型）＋Vertex AI／Gemini Enterprise——**是三者中唯一同時擁有「自研晶片＋自研frontier模型＋自研雲端」完整垂直整合的一家**，不像MSFT/AWS主要以「多模型選擇平台」角色定位自己 |
| Agent／企業整合層 | Agent 365（治理控制平面，2個月內近4,000萬agent註冊）、Foundry（10萬客戶）、Copilot生態（M365 Copilot 3,000萬付費席、GitHub Copilot 5,000萬用戶） | Amazon Q（**唯一明確整合第三方SaaS生態的助理**——原生連接Slack、Salesforce、Jira、Teams、ServiceNow，而非只鎖定自家生態）、Bedrock Agents、Kiro（agentic coding） | Agent Development Kit（累計約7,000萬次下載）、Gemini Enterprise——企業級agent佈局的具名案例集中在傳統產業（保險、零售、電信），敘事上比MSFT/AWS更強調垂直產業導入 |
| 資料庫／分析層 | Cosmos DB／PostgreSQL（AI優化資料庫，PostgreSQL營收+55%連續三季加速）、Fabric（4萬+付費客戶） | 未在本季法說會特別著墨資料庫線，重心在Agent基礎設施（Bedrock Agents的policies／payments／web search新功能） | 未在本季法說會特別著墨，重心在TPU規模化與Gemini Enterprise擴張 |

**⚑ 三者最根本的策略分歧**：Azure與AWS都明確採取「模型無關（model-agnostic）」的平台定位——把自己包裝成「不管客戶用哪家模型都能賺到運算費」的中立基礎設施，即使兩者都各自有品牌模型（MAI、Bedrock自研模型）與AI實驗室的股權/商業關係（MSFT持有OpenAI與Anthropic雙重股權曝險，本季認列$32億Anthropic評價利得）。Google則相反，選擇「自研晶片＋自研模型＋自研雲端」三位一體的垂直整合路線，不強調模型選擇的中立性，而是靠Gemini本身的能力與TPU的成本效益作為差異化賣點。這代表三者其實不是在打同一場「雲端市佔率」的仗，而是在用三種不同的商業模式假設（中立基礎設施 vs 中立基礎設施 vs 垂直整合）同時競爭同一群客戶的AI運算支出。

---

## 五、綜合觀察

1. **成長率與毛利率呈現明顯的「後進者優勢」排列**：Google Cloud（規模最小）成長最快、毛利率擴張最猛；AWS（規模最大）成長與毛利率擴張都居中；Azure／Intelligent Cloud（分部定義最混雜、但Azure本身成長率其實也很快）毛利率絕對水準最高但擴張已趨緩。這是雲端產業典型的S曲線位置差異，而非單純「誰的技術／管理比較好」。
2. **三者的CapEx紀律敘事高度一致**，都在本季法說會強調「短生命週期資產可彈性延後」「效率工作是毛利率的主要槓桿」——這代表整個產業對2026–2027年「需求是否持續超過供給」有共同的謹慎轉向，與2025年那種近乎無條件樂觀的CapEx敘事略有不同。
3. **AI實驗室客戶（OpenAI、Anthropic）同時是三家的關鍵客戶與變數來源**——OpenAI對Azure是雲端大客戶、對AWS是Trainium簽約大戶；Anthropic對AWS是Trainium大戶、對MSFT是股權投資標的（且本季$32億評價利得直接墊高了MSFT的EPS成長率敘事）。這種交叉關係代表任何一家AI實驗室的資本支出決策變化，都會同時牽動三家雲端巨頭的財報，是三者財報品質的共同不確定性來源，而非各自獨立的風險。

## 資料來源與限制

- 主要依據三份法說會逐字稿的管理層發言，以及MSFT FY2026 10-K（accession 0001193125-26-323660）、AMZN Q2 2026 10-Q（accession 0001018724-26-000026）的SEC EDGAR審計數字交叉核對（MSFT／AMZN部分，詳見兩份SEC Cross-Reference Notes文件）；GOOGL部分沿用既有報告，本輪未重新抓取10-Q。
- 市佔率為方向性推論，**非**引用任何第三方市場研究機構（Synergy Research／Canalys／Gartner等）的官方數字——如需精確市佔率百分比，需另外查證這類資料源。
- Azure的美元營收無法從MSFT 10-K獨立取得（僅Intelligent Cloud分部合併數字＋MD&A成長率），與AWS／Google Cloud的審計分部數字精確度不完全對等，比較「絕對規模」時請留意此落差。
- 僅供資訊參考，非投資建議。
