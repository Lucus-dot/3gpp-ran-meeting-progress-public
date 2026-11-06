---
name: 3gpp-ran-meeting-progress
version: "1.0.0"
description: 3GPP RAN 单场会议进展分析技能，适用于每场 RAN 会议（RAN1/2/4 或 TSG-RAN 全会）结束后，从会议素材（Chairman note 主席报告为唯一落槌真相源、FL/Moderator Summary 为讨论脉络、WI/SID 文档的 Objective 段为 WI 条目来源）出发，按 WI 条目组织，产出单场会议的横向进展快照：每个 WI 条目下挂本场 FL 讨论的 Issue，含讨论脉络、Chairman note 逐字结论、未收口项与单场 WI 小结；必要时才递归检索公司 TDoc（起始观点真相源），不接基线报告、不跨会议演进。当用户需要在会后从素材而非既有报告出发、按 WI 条目铺开一场会的完整进展时。
agent_created: true
---

# 3GPP 单场会议进展分析（按 WI 条目组织）

## 0. 一句话定位

**会后才用，产出"这一场会推进到了哪"。** 输入是会议素材（不是一份既有报告），产出是一份
**独立的单场会议进展文档**，按 **WI 条目**（从 WI/SID 文档的 Objective 段原文提取）组织。
每个 WI 条目下挂本场 FL 讨论的 Issue，每个 Issue 给三段：①讨论脉络 ②结论（Chairman note
逐字）③未结论（未收口/FFS）。

## 1. 与相邻技能的边界（务必先判，别用错）

| 场景 | 用哪个 |
|---|---|
| 会前：把各公司 TDoc 归类成各家观点（Yes/No/中立 + 支持公司 + 提案号） | `3gpp-proposal-analysis` |
| **会后：按 WI 条目铺开一场会的完整进展**（Issue 脉络 + 落槌结论 + 未收口 + WI 小结） | **本技能** |
| 会后：把落槌结论**接进既有深度报告**，产出跨会议续篇 | `3gpp-report-extension` |

本质差异（别混）：

- 本技能 = **单场横向快照**：WI 条目来自**本场提供的** WI/SID 文档 Objective 段；**不跨会议、无基线、无演进分析**。
- `3gpp-report-extension` = **跨会议纵向续篇**：WI 条目来自基线报告、做跨会议关系类型判定（CONTINUE/NEW/DERIVED/SPLIT/RELATED）。本技能**没有**这些关系类型。
- `3gpp-proposal-analysis` = **会前提案视角**：重点是各家公司的立场，不是会议结论。

> 别人写的"会议进展"文档**不算素材**。落槌只认 Chairman note；拿二手综述当输入等于把别人的解读再解读一遍。

## 2. 输入要求

| # | 东西 | 性质 | 没有会怎样 |
|---|---|---|---|
| 1 | **Chairman note**（`Chair notes <meeting>_eom*.docx`） | **必须** ★★★ 唯一落槌来源 | **阻塞**。无它则无结论可溯源 |
| 2 | **FL Summary 文件夹**（`FL summary/*.docx`，多份） | **必须** ★★ 讨论脉络 + Issue 树 | 阻塞产出形态；至少 1 份，典型 2~5 份 |
| 3 | **WI / SID 文档**（`.docx`/`.doc(OLE2)`/`.md`；亦可为 RP 立项提案 zip 内正文） | **必须** | **阻塞**。提供 WI 条目结构（核心在 **Objective 段**）；工具**自主识别 WID/SID**，无需你告知 |
| 4 | 议程项编号（如 `9.8.1`） | 参数，每次运行给 | 可缺，缺了写进文件名占位 |
| 5 | **公司 TDoc**（`by_agenda/.../*.docx`，层级可能很深） | 按需 ★ | 可缺；**仅当 FL 对某 Issue 起始观点说得不清晰时**递归抽取原文——此时 TDoc 是公司**起始观点**真相源（观点可能演进） |

### 素材夹布局（一场会议一个夹，不要混场）

```
RAN1#127/
├── Chair notes RAN1#127_eom1.docx          ★★★ 唯一落槌来源
├── FL summary/                             ★★ 会中逐轮刷新的对齐稿，可能 2~5 份
│   ├── R1-261xxxx.docx                         ← 标题写着 "FL Summary #1"
│   └── R1-261yyyy.docx                         ← 标题写着 "FL Summary #2"
├── <WI/SID 文档>.doc(x)                       必须：WI 条目来源（核心=Objective 段；可为 .doc/.docx；工具自动识别 WID/SID）
└── by_agenda/9/9.4/9.4.1/                   公司 TDoc，按议程逐级归档（层级可能很深，工具须递归探测全部）
```

- **FL Summary 单独一个夹**：会中每轮刷新都发一版，别摊在根目录。
- **轮次不用人工确认**：由工具读文档标题（`FL Summary #k: …`）识别；读不出时退回"编号小的靠前"并标注"推定"。文稿号（`R1-26…`）是文稿号不是轮次，工具会报出不一致。
- **WI/SID 文档格式**：可能 `.docx`（OOXML）或 `.doc`（OLE2 旧格式，RP 全会立项提案常见）；工具**兼容两者**且**自动判别 WID/SID**。
- **TDoc 递归探测**：`by_agenda/` 下层级可能很深，`sources` 与 `tdoc` 工具必须**递归遍历全部子目录**收集 `.docx`，不能只扫一层。

## 3. 铁律（G1–G5，不可跳级）

- **G1｜没有基线报告，产出的就是"单场会议进展"，不是"续篇"。** 不抄基线、不接某一节、不写演进分析。按 WI 条目组织，每个 WI 条目下挂本场 Issue。
- **G2｜权威层级（不可跳级）。**
  - ★★★ **Chairman note = 唯一落槌真相源**：Agreement / Working assumption / Observation 等结论块**逐字照录**，不意译、不合并、不代为实现收敛。无论输入 WID 还是 SID，Chairman note 都是结论唯一真相源（SID 的落槌呈现为"研究结论/agreement to study"）。
  - ★★ **FL Summary = 讨论脉络**：提供本场 Issue 树与立场演进；其 "Agree / Proposal" **未落槌**。第一轮 FL 通常已聚合各公司提案分析。
  - ★ **公司 TDoc = 公司起始观点真相源（按需）**：只有 FL 对某 Issue 起始观点说得不清时才读；称"起始"是因为立场**可能随讨论演进**——TDoc 只证明"某公司起始主张过 X"，不能证明"会上达成了 X"。
- **G3｜对账单元 = Chairman note 结论块；组织单元 = WI 条目（从 WI/SID 文档 Objective 段原文提取）。**
  - 对账逐字可溯源 → 按结论块切，一块不能少。
  - 组织稳定可溯源 → 按 WI 条目排；WI 条目 = Objective 段的 scope items（如 `Specify the following enhancements` 下的子项）。FL 讨论的每个 Issue 都针对某个 WI 条目，**必须挂到该 WI 条目下**。
  - Objective 段是 SI/WI 核心，定义本场要组织的全部 WI 条目；工具只抽这一段，原文直取、不编造。
  - 每个 WI 条目给一个**单场内小结**：已收口 / 部分收口 / 未完。**不做跨会议关系类型判定**。
  - 约束：**每个结论块在正文出现且仅出现一次**（块数守恒），可机械校验。
- **G4｜模型不吃素材原文，只吃工具切好的切片。** 不读 TDoc 全文、不读 Chairman note 全文、不读整篇进展文档、不读 WI/SID 文档全文（只吃抽出的 `wid.json` 条目）。每个模型步骤 = 指名切片 + 指名原文，有硬上限（见 §5）。
- **G5｜可机械验证的交给工具，靠判断的交给模型，用标注符号分开。** 只对前者上硬门禁。

## 4. 工作流

> 工具链见 `scripts/skillctl.py`（子命令 `init/extract/blocks/issues/flow/rel/tdoc/sources/compose/check`）。
> 步骤与门禁以 `config.json` 的 `stages` 为唯一真值；本文件不引入那里没有的步骤。

### A1　提取 WI/SID 条目（工具）　`skillctl init`
读 WI/SID 文档 → **自主识别 WID/SID** → 定位并抽 **Objective 段** → 逐条抽 scope item（原文 + 负责 WG `[RAN1]/[RAN2]/[RAN4]`）落 `wid.json` → 生成进展骨架（替换 `references/skeleton.md` 占位符）。
- **WID/SID 识别**（工具按优先级匹配，任一命中即定）：标题/文档含 "New WID"/"Work Item Description"/"Objective of Core part **WI**"/"Specify…（出规范）" → **WID**；含 "New SID"/"Study Item Description"/"Objective of Core part **SI**"/"Study…（仅研究、不出规范）" → **SID**。结果只影响 `wid.json.type` 与条目性质标注，不改变 G2。
- **格式兼容**：`.docx` 走 `document.xml` 解析；`.doc(OLE2)` 走 cp1252/utf-16 粗暴解码（OLE2 正文多为单字节，英文/数字串可抓出）。
- 模型**不碰**此步（原文直取，不编造条目名）。

### A2　抽文本（工具）　`skillctl extract`
docx → 段落流（保 OMath 公式），落 `_eom_full.txt` 与 `_fl/*.txt`。模型读 `_eom_full.txt` 定位目标议程项、裁出 `_eom_verbatim.txt`（裁段是模型活）。

### A3　切块（轮次不在此定）（工具）　`skillctl blocks`
按可配置的结论块标记（`Agreement|Working assumption|Observation|Note`…）切块；抽块内 `FFS:`/`to discuss` → `ffs.json`。注意：**真实主席报告不内嵌 FL 分隔行**，故结论块的 `fl_round` 一律为 `N/A`（轮次来自 FL 文档本身，见 A4/A5）。

### A4　抽议题骨架（工具）　`skillctl issues`
从 `_fl/*.txt`（早→晚）抓 `Topic#k:` / `Issue#k-m:` 树 + 每个 Issue 的 `Companies:/Proponents:` 行、`Key technical arguments`、`Proposal` 段、FFS；以最末一份 FL 为骨架基准，与更早各份 diff。

### A5　对齐：结论块 → Issue（工具候选 + 模型判定）　`skillctl flow`
每块与每个 Issue 的 `Proposal` 段做 n-gram 余弦相似度取 top-3，零命中标 `UNMAPPED`；并串讨论链（各轮 FL 段落按轮排、Proposal diff 得 `delta`、比块原文得 `final_status`）。模型只判 `UNMAPPED` 与候选 ≥2 的块（预期 ≤5）。

### A5b　挂接：Issue → WI 条目（必须，工具候选 + 模型判定）　`skillctl rel`
双路候选（WI 关键词 + Issue 标题/Proposal 重叠），落空标 `UNMAPPED`。模型只判 `UNMAPPED` 与冲突的 Issue（预期 ≤6）；落不到既有条目的判是否派生新条目（N 系列），写明依据（WID 哪句 NOTE / 哪份 TS）。这是 report-extension 的 B1 下沉——WI 必须，挂接不可省；不做跨会议关系类型判定。

### A6　撰写：每个 Issue 三段（模型主笔，挂所属 WI 条目下）
一次只写一个 Issue：
1. **① 讨论脉络** — FL#1→FL#k 各轮立场演进 + delta（不写"达成了"）；FL 对起始观点不清时引 TDoc 原文标 `TDoc#R1-xxx · Issue#x-y`。
2. **② 结论** — 对应 Chairman note 结论块**逐字原文** + `【解读】` + `【映射】`。
3. **③ 未结论** — 无对应结论则标「未收口 / FFS / 后续轮次」，原样列 `ffs.json` 相关项。

`【原文】` 一字不改（排印错误用 `〔校〕` 注明）；`【解读】` 每个公司名须能在关联 FL 切片定位并标 `FL#k / Issue#x-y`；`【映射】` 写 WI 条目。

### A6b　WI 条目小结（模型判定）
对该 WI 条目下所有 Issue 的 `final_status` 给一句话判定：已收口 / 部分收口 / 未完 + 理由。不引入切片外事实。

### A7　组装正文（工具 + 模型）　`skillctl compose`
按 WI 条目组织：WI 条目节标题、`WI 小结`占位、每 Issue 三骨架、总表、块序（按 `fl_round`）、块标签、块数自证、未结论占位由工具发；`【解读】`/`【映射】`/WI 小结/未结论判别由模型填。**幂等**（重跑字节不变）、**块数守恒**（每结论块出现且仅出现一次）。

### A8　校验（工具）　`skillctl check`
块数 == `blocks.json` 条数 == `--eom` 切分数；每 `【原文】` 能在 Chairman note 逐字定位；**每个 Issue 都有 WI 归属**（无 UNMAPPED 残留）；**每个 Issue 要么有结论、要么标「未收口」**；每个 WI 条目都有小结判定；花名册声明篇数 == 主席报告 tdoc 列表行数；骨架占位符已填实；`【解读】` 公司名有出处；不出现"FL 的 Agree/Proposal 写成落槌"的反模式。

## 5. 模型上下文预算（硬约束）

| 步骤 | 允许喂 | 禁止喂 | 上限 |
|---|---|---|---|
| A5 块→Issue | 候选≥2 的块原文 + 各候选 Issue 的 `Proposal` 段 | 其余块、其他 Issue、TDoc | 5 KB |
| A5b Issue→WI | UNMAPPED/冲突 Issue + `wid.json` + WI/SID 文档段 | Chairman note 全文、TDoc 全文 | 8 KB |
| A6 单 Issue 三段 | 1 Issue 的 FL 脉络切片 + 其下各块原文 + `flow[k]` + 其 WI 条目表述 | 其他 Issue、Chairman note 全文、TDoc 全文 | 6 KB/次 |
| A8 定性 | A6/A6b 产物 + `ffs.json` | 素材原文全文 | 30 KB |

**为什么必须切**：全篇一次喂完必然"前面抄得准、后面开始编"。切片后每条独立可查：错一条重跑一条。

## 6. 标注四件套

`【原文】` 逐字 / `【解读】` 技术含义 / `【折算】` 本报告推导（仅当文档含数值且给了 TR 表时才出现）/ `【映射】` WI 条目。

## 7. 工具速查

| 命令 | 干什么 |
|---|---|
| `skillctl sources <素材夹> --agenda <项>` | 清点素材：主席报告在否、FL 几份、tdoc **递归**计数 |
| `skillctl init --wi <WI文档> --meeting <会议> --agenda <项>` | 抽 Objective 段 + 识别 WID/SID + 生成骨架 |
| `skillctl extract <docx> --out <txt>` | docx → 纯文本段落 |
| `skillctl blocks <verbatim.txt> --out-dir <_work>` | 结论块切块 + FFS（轮次来自 FL，不在此定） |
| `skillctl issues <FL1> <FL2> … --out-dir <_work>` | 抽 Issue 树（顺序早→晚） |
| `skillctl flow --blocks … --issues … --out-dir <_work>` | 块→Issue 候选 + 串讨论链 + final_status |
| `skillctl rel --flow … --wid … --out-dir <_work>` | Issue→WI 挂接（必须） |
| `skillctl tdoc <R1-xxx> --out <txt>` | 按需抽单篇 TDoc（须在 flow.json 记 tdoc_reason） |
| `skillctl compose --rel … --blocks … --target <文档>` | 按 WI 条目组装（幂等、块数守恒） |
| `skillctl check <文档> --eom <逐字> --workdir <_work>` | 机械校验 |

## 8. 踩坑速记（详见 `references/workflow.md`）

- 把相邻两个结论块并成一条（计数少 1）→ 以字符串 `Agreement:`/`Working assumption` 为切分锚逐块数。
- 沿用上一场的计数 → 每场重新从主席报告 tdoc 列表逐条数。
- 把 FL 的 "Agree" 写成落槌 → 一律降级为"仅 FL 级 Agree，未进主席报告"，记为未收口；`【原文】` 必须能在 Chairman note 逐字 grep 到。
- 只写正增量 → 未收口清单强制含停滞项，并点名结构性缺口。
- 给 WID 未明列的议题借用既有编号 → 另设 N 系列并写派生依据。
- 块标题重复 / 行内粗体被抹 / 分隔线跟着内容来回搬 / 用 blocks.json 覆盖已排版原文 → 这些是 `compose` 的幂等与"正文优先"纪律，重跑字节不变才算过关。
