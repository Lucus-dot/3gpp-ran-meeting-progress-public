# 3GPP RAN 会议进展分析 — 工作流与约定（详细参考）

本文件是 `3gpp-ran-meeting-progress` skill 的详细参考。SKILL.md 放核心规则与脚本入口；
这里放结构约定、踩坑速记、以及通用端到端流程。**本技能是单场横向快照**：WI 条目来自
本场提供的 WI/SID 文档 Objective 段，不接基线、不跨会议演进。

## 1. 端到端流程（Phase A，单场）

> 工具链 `scripts/skillctl.py`（子命令见 SKILL.md §7）。步骤与门禁以 `config.json` 的 `stages` 为唯一真值。

### A1　提取 WI/SID 条目　`skillctl init`
```
skillctl init --wi <WI文档> --meeting <会议> --agenda <项> [--out <骨架.md>]
```
- 工具**自主识别 WID/SID**（先看文件名，再看正文关键词；结果写 `wid.json.type`）。
- 工具定位并抽取 **Objective 段**（SI/WI 核心），逐条抽 scope item（原文 + 负责 WG `[RAN1]/[RAN2]/[RAN4]`）落 `wid.json`。
- 工具生成进展骨架（替换 `references/skeleton.md` 占位符，WI 名称来自 Objective 段原文）。
- ⚠️ **Objective 段抽取是启发式**：对 RP 级二进制 `.doc`，工具以"拉丁词数最多"选型（CP1252 / UTF-16LE），
  并跳过 lead-in（`For …,`）、NOTE 模板、RAN time budget / Work item leadership / Supporting members 等模板段。
  仍可能混入噪声或漏抽——**最终以 `wid.json` 为准，模型可手动增删条目**。
- SID 与 WID 处理一致：Chairman note 仍是唯一落槌真相源，只是落槌呈现为"研究结论/agreement to study"。

### A2　抽权威文本　`skillctl extract`
```
skillctl extract <Chair notes>.docx --out _work_<项>/_eom_full.txt
skillctl extract <FL>.docx --out _work_<项>/_fl/<名>.txt   # 每份 FL 各一次
```
- docx → 段落流（保 OMath 公式，见坑 P5）。模型读 `_eom_full.txt` 定位目标议程项、裁出 `_eom_verbatim.txt`。

### A3　切块（轮次不在此定）　`skillctl blocks`
```
skillctl blocks _eom_verbatim.txt --out-dir _work_<项>
```
- 结论块标记可配置（`Agreement|Working assumption|Observation|Note|Conclusion`…）。
- 抽块内 `FFS:`/`to discuss` → `ffs.json`。**轮次不在此定**：真实主席报告无内嵌 FL 分隔行，结论块 `fl_round` 恒为 `N/A`；轮次见 A4/A5（每份 FL 文档 = 一轮，按文件顺序 `round`）。

### A4　抽议题骨架　`skillctl issues`
```
skillctl issues _fl/FL1.txt FL2.txt … --out-dir _work_<项>   # 顺序必须早→晚
```
- 抓 `Topic#k:` / `Issue#k-m:` 树 + 每个 Issue 的 `Companies:/Proponents:` 行、`Proposal` 段、FFS；以最末 FL 为基准，与更早各份 diff。

### A5　对齐：结论块 → Issue　`skillctl flow`
```
skillctl flow --blocks blocks.json --issues issues.json --out-dir _work_<项>
```
- 每块与每个 Issue 的 `Proposal` 段做 n-gram 余弦相似度取 top-3，零命中标 `UNMAPPED`；串讨论链、`final_status`。
- 模型只判 `UNMAPPED` 与候选 ≥2 的块（预期 ≤5）。

### A5b　挂接：Issue → WI 条目（必须）　`skillctl rel`
```
skillctl rel --flow flow.json --wid wid.json --out-dir _work_<项>
```
- 双路候选（WI 关键词 + Issue 标题/Proposal 重叠），落空标 `UNMAPPED`。模型判 `UNMAPPED` 与冲突 Issue（预期 ≤6）；落不到既有条目的判是否派生新条目（N 系列），写明依据。

### A6　撰写：每个 Issue 三段（模型主笔）
- ① 讨论脉络（FL#1→FL#k 立场演进 + delta，不写"达成了"）② 结论（对应块【原文】逐字 + 【解读】+ 【映射】）③ 未结论（无结论标"未收口 / FFS / 后续轮次"）。
- `【原文】` 一字不改（排印错误用 `〔校〕` 注明）；`【解读】` 每个公司名须能在关联 FL 切片定位并标 `FL#k / Issue#x-y`；`【映射】` 写 WI 条目。

### A6b　WI 条目小结（模型判定）
- 对该 WI 下所有 Issue 的 `final_status` 给一句话判定：已收口 / 部分收口 / 未完 + 理由。不引入切片外事实。

### A7　组装正文　`skillctl compose`
```
skillctl compose --rel rel.json --blocks blocks.json --target <进展文档>
```
- 按 WI 条目组织，幂等、块数守恒、正文优先（不覆盖已写内容，只填占位符 + 生成 §3 建议结构）。

### A8　校验　`skillctl check`
```
skillctl check <进展文档> --eom <逐字> --workdir _work_<项>
```
- 块数守恒 / 原文溯源 / WI 归属 / 未收口覆盖 / 占位符填实 / 公司名出处 / FL 反模式。ERROR=0 才交付。

---

## 2. 素材夹布局（一场会议一个夹，不要混场）

```
RAN1#127/
├── Chair notes RAN1#127_eom1.docx          ★★★ 唯一落槌来源
├── FL summary/                             ★★ 会中逐轮刷新的对齐稿（可 2~5 份）
│   ├── R1-261xxxx.docx
│   └── R1-261yyyy.docx
├── <WI/SID 文档>.doc(x)                       必须：WI 条目来源（核心=Objective 段；可为 .doc/.docx；工具自动识别 WID/SID）
└── by_agenda/9/9.4/9.4.1/                   公司 TDoc，按议程逐级归档（层级可能很深，工具须递归探测全部）
```

`sources` 清点：
```
skillctl sources <素材夹> --agenda <项>
```
会报主席报告 / FL 份数 / WI 文档 / TDoc **递归**篇数。若布局不标准（如示例 `examples/9` 把讨论稿丢在 `9.4/9.4.1/` 下），`sources` 可能误判，需人工辨认。

---

## 3. 踩坑速记（全部真实会发生）

- **P1 把相邻两个结论块并成一个（计数少 1）** → 以字符串 `Agreement:` / `Working assumption` 为切分锚逐块数，不按语义数。
- **P2 沿用上一场的计数** → 每场重新从主席报告 tdoc 列表逐条数；花名册贴进文档当证据。
- **P3 把 FL Summary #1 当成 tdoc 清单** → 逐个核对 FL **标题**（`FL Summary #n: …`），不靠编号推。
- **P4 把 FL 的 "Agree" 写成落槌** → 一律降级为"仅 FL 级 Agree，未进主席报告"，记为未收口；`【原文】` 必须能在 Chairman note 逐字 grep 到。
- **P5 docx 公式坑（oMath）** → `extract` 只抽 `document.xml` 段落，公式（`<m:t>`）虽保留但上标/结构会丢；涉及公式处回 `document.xml` 核对，量级自检（如 FAC 粒度 `SCS/2^k` 别算成 `SCS×2^k`）。
- **P6 `.doc`（OLE2）编码坑** → RP 全会立项提案多为**单字节 CP1252**（不是 UTF-16LE！）；朴素解码会整体乱码。工具已按"拉丁词数最多"选型；若 Objective 段抽不出或明显错乱，**转成 .docx 后用 Word/LibreOffice 重新抽** 最稳。
- **P7 块标题重复 / 行内粗体被抹 / 分隔线跟着内容来回搬 / 用 blocks.json 覆盖已排版原文** → 这些是 `compose` 的"幂等"与"正文优先"纪律，重跑字节不变才算过关；只有正文里还没有的块才从 `blocks.json` 注入。
- **P8 只写正增量** → 未收口清单强制含停滞项，并点名结构性缺口（两条并列分支只收口了一条）。
- **P9 给 WID 未明列的议题借用既有编号** → 另设 **N 系列**并写派生依据（WID 的 NOTE / 受影响 TS）。
- **P10 块数守恒** → 每个结论块在正文出现且仅出现一次；复制原文跨 WI 会破坏守恒。

---

## 4. 校验清单（`check` 应覆盖）

| # | 检查 | 级别 |
|---|---|---|
| 1 | 结论块数 == `blocks.json` 条数 == `--eom` 切分数 | ERROR |
| 2 | 每个 `【原文】` 在 Chairman note 逐字可定位 | ERROR |
| 3 | **每个 Issue 都有 WI 归属**（A5b 覆盖），无 UNMAPPED 残留 | ERROR |
| 4 | **每个 Issue 都有结论或标「未收口」**，无遗漏、无伪造 | ERROR |
| 5 | 每个 WI 条目都有小结判定且与旗下 Issue 状态一致 | ERROR |
| 6 | 花名册声明篇数 == 主席报告 tdoc 列表行数 | ERROR |
| 7 | 骨架占位符已填实（无 `TODO`/`待补`/`{{…}}`） | ERROR |
| 8 | 每个 `【解读】` 公司名能在关联 FL 切片定位且标 `FL#k / Issue#x-y` | ERROR |
| 9 | 不出现"FL 的 Agree/Proposal 写成落槌" | ERROR |
| 10 | 块数守恒：每块在正文出现且仅出现一次 | ERROR |

---

## 5. 与相邻技能边界（再强调）

- `3gpp-proposal-analysis`：会前，按 TDoc 归类各家观点（Yes/No/中立 + 支持公司 + 提案号）。
- **本技能**：会后，按 WI 条目铺开**一场会**的完整进展（Issue 脉络 + 落槌 + 未收口 + WI 小结）。
- `3gpp-report-extension`：会后，把落槌结论接进既有深度报告，产出**跨会议续篇**（基线 + 关系类型 CONTINUE/NEW/DERIVED/SPLIT/RELATED + 演进分析）。本技能**没有**这些；本技能 WI 条目来自本场 WI/SID 文档 Objective 段，且**不跨会议**。

---

## 6. 实例：RP-261568（GNSS resilient NR-NTN WID）Objective 结构

`RP-261568 was 1553 New WID on GNSS resilient NR-NTN operation` 是 WID（文件名含 "New WID"，正文 Objective 段含 "Objective of Core part WI"）。其 Objective 段结构（工具抽取验证）：

- *Objective of Core part WI*：分 `For GNSS-resilient Initial Access in NR NTN`（PRACH 双序列 + 不同 roots、TA command/FAC in Msg2、initial transmit timing/frequency RRM）/ `For GNSS resilient NR-NTN operation in Connected Mode`（TA control、复用 legacy closed-loop TA、implicit/explicit FAC、connected-mode RRM）/ `For both connected and idle mode`（assistance information、mobility procedures、measurement RRM）三块，每块下是 `Specify the following enhancements` 的 scope items，条目尾标负责 WG（`[RAN1]`/`[RAN2]`/`[RAN4]`）。
- *Objective of Performance part WI*：demodulation / RRM 性能要求（PRACH、PUSCH）。
- *RAN time budget request*（附 `.xls`）：模板段，工具会跳过。

这些 scope items 即本技能的 WI 条目（`wid.json` 来源）。`init` 在该 `.doc` 上实测抽到 13 条核心条目并正确识别为 **WID**。
