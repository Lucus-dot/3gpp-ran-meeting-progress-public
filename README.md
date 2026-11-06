# 3gpp-ran-meeting-progress

**3GPP RAN 单场会议进展分析技能** —— 会后从会议素材（而不是一份既有报告）出发，按 **WI 条目**铺开"这一场会推进到了哪"。

![License](https://img.shields.io/badge/license-MIT-blue)
![Python](https://img.shields.io/badge/python-3.6%2B-3776AB)

> 一句话定位：**会后才用，产出单场横向快照。** 每个 WI 条目下挂本场的 FL 讨论 Issue，每个 Issue 三段：①讨论脉络 ②Chairman note 逐字结论 ③未收口项。不接基线报告、不跨会议演进。

---

## 这是什么 / 不是什么

| | 说明 |
|---|---|
| ✅ 是 | 输入一场会议的素材（主席报告 + FL Summary + WI/SID 文档 ± 公司 TDoc），输出一份**独立的单场会议进展文档** |
| ✅ 是 | **组织单元 = WI 条目**（从 WI/SID 文档 Objective 段原文提取），**对账单元 = 结论块**（Chairman note 逐字，块数守恒） |
| ❌ 不是 | 跨会议的"续篇"：无基线、无演进分析、无 CONTINUE/NEW/DERIVED/SPLIT/RELATED 关系判定 |
| ❌ 不是 | 会前提案归类（各家 Yes/No/中立）——那是 `3gpp-proposal-analysis` |
| ❌ 不是 | 二手综述的再综述。**别人写的"会议进展"文档不算素材**，落槌只认 Chairman note |

### 与相邻技能的边界

| 场景 | 用哪个 |
|---|---|
| 会前：把各公司 TDoc 归类成各家观点（Yes/No/中立 + 支持公司 + 提案号） | `3gpp-proposal-analysis` |
| **会后：按 WI 条目铺开一场会的完整进展**（Issue 脉络 + 落槌结论 + 未收口 + WI 小结） | **本技能** |
| 会后：把落槌结论**接进既有深度报告**，产出跨会议续篇 | `3gpp-report-extension` |

---

## 产出长什么样

`skillctl init` 用 `references/skeleton.md` 生成骨架，`compose` 按 WI 条目组装，最终文档结构固定为：

```
# <会议> · 议程 <项> 会议进展（<主题>）
§1  会议元信息与议程结构      （会议概况 / 公司贡献花名册 / FL 轮次时间线）
§2  议题映射表                （Issue ↔ WI 条目 ↔ 本场是否有 chairman 结论）
§3  按 WI 条目的本场进展      ← 核心
      WI-<n>｜<WI 条目原文>     （WI 小结：收口 / 部分收口 / 未完）
        Issue#x-y｜<标题>
          ① 讨论脉络   FL#1→FL#k 立场演进 + delta
          ② 结论       【原文】逐字 + 【解读】 + 【映射】
          ③ 未结论     未收口 / FFS / 后续轮次
§4  未收口与下一步            （FFS 清单，按阻塞性排序）
附录 A  文档索引
```

**硬约束**：本场每个结论块在全文**出现且仅出现一次**（块数守恒），可机械校验。

---

## 快速开始

### 依赖

- **Python 3.6+**，**只用标准库**（`argparse` / `json` / `re` / `zipfile` / `hashlib`），无需 `pip install` 任何东西。
- `.docx` 走 `zipfile` 解析 `document.xml`；`.doc`（OLE2 旧格式，RP 全会立项提案常见）走单字节粗暴解码——见「已知限制」。

### 作为技能使用

把整个目录放进技能目录（如 `~/.workbuddy/skills/3gpp-ran-meeting-progress/`），`SKILL.md` 即技能定义，`config.json` 的 `stages` 是步骤与门禁的**唯一真值**。素材夹准备好后直接说：

> 「RAN1#127 的 9.4.1 已经开完了，帮我出这一场的进展」

### 作为命令行工具使用

```bash
# 0) 先清点素材，缺什么当场知道
python scripts/skillctl.py sources /path/to/RAN1#127 --agenda 9.4.1

# 1) 抽 Objective 段、自动识别 WID/SID、生成骨架
python scripts/skillctl.py init --wi "/path/to/RP-261568_WID.doc" \
       --meeting "RAN1#127" --agenda 9.4.1 --out-dir _work_9.4.1

# 之后的 A2–A8 见下表
```

`sources` 的真实输出形如（缺素材直接报阻塞点）：

```
== 素材清点 ==
主席报告（★ 唯一落槌）：1 份
FL Summary：3 份（夹：FL summary）
WI/SID 文档：1 份
  - RP-261568_WID.doc
公司 TDoc（递归探测）：37 篇
```

---

## 素材夹布局

一场会议一个夹，**不要混场**：

```
RAN1#127/
├── Chair notes RAN1#127_eom1.docx      ★★★ 唯一落槌来源（必须，缺则阻塞）
├── FL summary/                          ★★ 会中逐轮刷新，典型 2~5 份（必须）
│   ├── R1-261xxxx.docx                 ← 标题 "FL Summary #1"
│   └── R1-261yyyy.docx                 ← 标题 "FL Summary #2"
├── <WI/SID 文档>.doc(x)                 必须：WI 条目来源（核心 = Objective 段）
└── by_agenda/9/9.4/9.4.1/              公司 TDoc，按议程逐级归档（层级可很深，工具递归探测）
```

- **FL Summary 单独一个夹**：会中每轮刷新都发一版，别摊在根目录。
- **轮次不用人工确认**：工具读文档标题 `FL Summary #k: …` 识别；读不出时退回"编号小的靠前"并标注"推定"。文稿号 `R1-26…` 是文稿号不是轮次。
- **WI/SID 文档**兼容 `.docx` / `.doc` / `.md`，也支持 RP 立项提案 zip 内正文；工具**自动判别 WID/SID**，无需你告知。

---

## 工作流（A1–A8）

步骤与门禁以 `config.json` 的 `stages` 为唯一真值，`SKILL.md` 不引入那里没有的步骤。

| 步 | 干什么 | 命令 | 谁做 |
|---|---|---|---|
| A1 | 抽 Objective 段 → 识别 WID/SID → 逐条 scope item 落 `wid.json` + 生成骨架 | `init --wi … --meeting … --agenda …` | 工具 |
| A2 | docx → 段落流（保 OMath），模型再裁出目标议程项的 `_eom_verbatim.txt` | `extract <docx> --out <txt>` | 工具 + 模型裁段 |
| A3 | 按 `Agreement` / `Working assumption` / `Observation` / `Note` 切结论块，抽 `FFS:` → `ffs.json` | `blocks <verbatim.txt> --out-dir _work` | 工具 |
| A4 | 从 `_fl/*.txt`（早→晚）抓 `Topic#k` / `Issue#k-m` 树 + 立场段 + Proposal + 跨轮 diff | `issues FL1 FL2 … --out-dir _work` | 工具 |
| A5 | 块 → Issue 候选（n-gram 余弦 top-3）+ 串讨论链 + `final_status` | `flow --blocks … --issues … --out-dir _work` | 工具候选 + 模型判 ≤5 块 |
| A5b | Issue → WI 条目挂接（**必须**，双路候选） | `rel --flow … --wid … --out-dir _work` | 工具候选 + 模型判 ≤6 项 |
| A6 | 每个 Issue 写三段（脉络 / 结论 / 未结论） | — | 模型主笔 |
| A6b | 每个 WI 条目一句话小结（收口 / 部分收口 / 未完） | — | 模型判定 |
| A7 | 按 WI 条目组装正文（幂等、块数守恒、正文优先） | `compose --rel … --blocks … --target <文档>` | 工具 + 模型填空 |
| A8 | 机械校验，ERROR 归零才许交 | `check <文档> --eom <逐字> --workdir _work` | 工具 |

按需补充：`skillctl tdoc <R1-xxx> --out <txt>` —— 仅当 FL 对某 Issue 的**起始观点**说不清时才用，且在 `flow.json` 记 `tdoc_reason`。

### 铁律（G1–G5）

1. **G1** 没有基线报告，产出的就是"单场会议进展"，不是"续篇"。
2. **G2** 权威层级不可跳级：Chairman note ★★★ 唯一落槌（逐字照录）＞ FL Summary ★★ 仅脉络（其 "Agree" **未落槌**）＞ TDoc ★ 公司**起始**观点（立场可能演进）。
3. **G3** 对账单元 = 结论块；组织单元 = WI 条目；块数守恒、WI 必须挂接。
4. **G4** 模型只吃工具切好的切片，不读素材全文（单步上限 5–8 KB）。
5. **G5** 可机械验证的交工具上硬门禁，靠判断的交模型并标注。

### 标注四件套

`【原文】` 逐字（排印错误用 `〔校〕` 注明，不在原文行里插话）/ `【解读】` 技术含义 / `【折算】` 本报告推导（仅当文档含数值且给了 TR 表）/ `【映射】` WI 条目。

---

## 目录结构

```
3gpp-ran-meeting-progress/
├── SKILL.md              技能定义（触发条件、边界、铁律、工作流、上下文预算）
├── config.json           步骤与门禁的唯一真值（stages / gates / required_sections / forbidden_patterns）
├── scripts/
│   └── skillctl.py       工具链，1008 行，纯标准库（sources/init/extract/blocks/issues/flow/rel/tdoc/compose/check）
├── references/
│   ├── skeleton.md       进展文档骨架模板（占位符由 init 替换）
│   └── workflow.md       端到端详版流程 + 踩坑清单 + 校验清单
├── NOTICE.md             版权与合规声明
└── LICENSE               MIT
```

**按项目改参数**：`config.local.json`（顶层键叠加在 `config.json` 上，如 `required_sections`），已在 `.gitignore` 中排除。

---

## 已知限制（如实说明）

- **`extract` 会丢**：公式（oMath 抽取不完整）、表格结构、批注。stderr 会报丢失项；涉及公式处须回 `document.xml` 核对。
- **`.doc`(OLE2) 是粗暴解码**：只可靠抓出英文/数字串（足够抽 WID 标题与 Objective 段），中文可能乱码——优先向来源索取 `.docx`。
- **轮次识别靠 FL 标题**：真实主席报告不内嵌 FL 分隔行，故结论块的 `fl_round` 恒为 `N/A`；轮次来自 FL 文档标题，读不出时标"推定"。
- **`check` 只做机械校验**：块数守恒、原文溯源、WI 归属、未收口覆盖、占位符填实、公司名出处。语义正确性仍需人复核。
- **WI 条目数量若与 Objective 段 scope item 数不符**，说明 Objective 段定位偏了，回 A1 重跑 `init`。

---

## 版权与合规

- 本仓库**不内置、不随包分发任何 3GPP 文档**（主席报告、FL Summary、TDoc、WI/SID 文档等）。所有素材由使用者自行提供到本地素材夹，且须已合法获得。
- 3GPP 会议文稿、技术规范受 3GPP 知识产权与版权政策约束；**不要**将受版权保护的原文随本技能或本技能的产出对外分发。
- 产出的进展分析是基于一手素材的二次分析，结论引用**以原始 Chairman note 为准**。
- 详见 [NOTICE.md](NOTICE.md)。

## License

[MIT](LICENSE) © 2026 Lucus
