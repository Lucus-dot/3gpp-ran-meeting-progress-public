#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
3gpp-ran-meeting-progress — 单场会议进展分析工具链（Phase A，单场；无 Phase B 跨会议演进）。

子命令（详见 config.json 的 stages，本文件不引入那里没有的步骤）：
  sources  清点一场会议的素材夹（主席报告 / FL Summary / WI文档 / TDoc 递归计数）
  init     读 WI/SID 文档 → 抽 Objective 段 → 自主识别 WID/SID → 落 wid.json + 生成骨架
  extract  docx/doc(.doc OLE2)/txt/md → 纯文本段落（保 OMath 公式）
  blocks   Chairman note 逐字文本 → 结论块切块 + 按内嵌 FL 分隔行定轮次 + FFS
  issues   FL Summary 多份 → Topic#/Issue# 树 + 立场段/Proposal/FFS + 跨轮 diff
  flow     结论块 → Issue 候选（n-gram 余弦）+ 串讨论链 + final_status
  rel      Issue → WI 条目挂接候选（必须）；UNMAPPED 留给模型
  tdoc     按需抽取单篇公司 TDoc（须在 flow.json 记 tdoc_reason）
  compose  按 WI 条目组装正文（幂等、块数守恒、正文优先）
  check    机械校验（块数守恒 / 原文溯源 / WI 归属 / 未收口覆盖 / 占位符填实）

设计铁律（来自 SPEC / SKILL.md）：
  G2  Chairman note 唯一落槌真相源；FL 仅脉络；TDoc 起始观点（按需）。
  G3  对账单元=结论块（逐字可溯源）；组织单元=WI 条目（Objective 段原文）。块数守恒。
  G4  模型只吃切片，不读素材全文。
  G5  可机械验证的交给工具，靠判断的交给模型。

复用自 3gpp-report-extension 的设计范式（只读参考，未修改其源码），并按单场快照定位裁剪
（砍掉 Phase B 的关系类型/演进分析；新增 Objective 抽取 + WID/SID 自主识别 + .doc OLE2 兼容
+ TDoc 递归探测）。
"""

import argparse
import json
import os
import re
import sys
import hashlib
import zipfile

# ---------------------------------------------------------------------------
# 通用工具
# ---------------------------------------------------------------------------

DOCX_EXT = (".docx",)
DOC_EXT = (".docx", ".doc", ".txt", ".md")
TDOC_EXT = (".docx", ".doc", ".pdf", ".txt")  # 公司 TDoc 原始文档


def _read_text_file(path):
    # 尽量用 utf-8，失败回退。
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def _docx_paras(path):
    """零依赖：docx 即 zip + word/document.xml；抽取 <w:p> 段落内的 <w:t>/<m:t>（保 OMath）。"""
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8", "replace")
    # 段落切分
    paras = re.split(r"</w:p>", xml)
    out = []
    for p in paras:
        # 取该段落内所有 w:t 与 m:t 文本
        texts = re.findall(r"<(?:w:t|m:t)[^>]*>(.*?)</(?:w:t|m:t)>", p, re.S)
        line = "".join(texts)
        line = re.sub(r"<[^>]+>", "", line)  # 去掉残留标签
        line = line.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
        out.append(line)
    return out


def _read_doc_ole2(path):
    """OLE2 .doc（RP 全会立项提案常见）粗暴解码：多编码候选，按拉丁词数选最优。

    实测：RP 级 WID 多为**单字节 CP1252**（非 UTF-16LE），utf-16-le 会整体乱码；
    故选型不以'可打印比例'（CJK 乱码也算可打印）为准，而以拉丁词数（[A-Za-z]{4,}）为准，
    保证选到能恢复正文（Objective / [RANx] 标签 / PRACH 等英文）的编码。
    """
    with open(path, "rb") as f:
        data = f.read()
    candidates = []
    for enc in ("cp1252", "latin-1"):
        try:
            candidates.append(data.decode(enc, "ignore"))
        except Exception:
            pass
    for off in (0, 1):
        try:
            candidates.append(data[off:].decode("utf-16-le", "ignore"))
        except Exception:
            pass
    best, best_score = "", -1
    for s in candidates:
        score = len(re.findall(r"[A-Za-z]{4,}", s))
        if score > best_score:
            best_score, best = score, s
    # 按换行切，去掉纯空白行；保留可见行（含表格/公式乱码行也无妨，Objective 段会自然清晰）
    return [ln for ln in best.splitlines() if ln.strip()]


def read_doc(path):
    """统一入口：.docx / .doc(OLE2) / .txt / .md → 段落列表（每元素一行）。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return _docx_paras(path)
    if ext in (".txt", ".md"):
        return _read_text_file(path).splitlines()
    if ext == ".doc":
        return _read_doc_ole2(path)
    # 兜底
    try:
        return _docx_paras(path)
    except Exception:
        return _read_doc_ole2(path)


def _dump(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def _load(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _norm(s):
    """归一化：去空白、小写，用于逐字比对。"""
    return re.sub(r"\s+", "", (s or "").lower())


def _work_dir(out_dir, agenda):
    if out_dir:
        d = out_dir
    else:
        d = "_work_%s" % (agenda or "x")
    os.makedirs(d, exist_ok=True)
    return d


# ---------------------------------------------------------------------------
# n-gram 重叠（flow 候选）
# ---------------------------------------------------------------------------

def _toks(s):
    # 英文按词，中文按字，混合。token 不少于 2 字符。
    s = (s or "").lower()
    toks = re.findall(r"[a-z0-9]+|[\u4e00-\u9fff]", s)
    return toks


def _grams(s, n=4):
    t = _toks(s)
    if len(t) < n:
        return set(t)
    return set(tuple(t[i:i + n]) for i in range(len(t) - n + 1))


def _overlap(a, b, n=4):
    ga, gb = _grams(a, n), _grams(b, n)
    if not ga or not gb:
        return 0.0
    return len(ga & gb) / len(ga | gb)


# ---------------------------------------------------------------------------
# sources：素材清点（递归探测 TDoc）
# ---------------------------------------------------------------------------

def find_root(explicit=None):
    if explicit:
        return explicit
    # 当前目录
    return "."


def _walk_docs(root, exts=DOC_EXT):
    res = []
    for dirpath, dirnames, filenames in os.walk(root):
        # 跳过工作目录与版本控制
        parts = dirpath.replace("\\", "/").split("/")
        if any(p in (".git", "_work_x", "_work") for p in parts):
            continue
        for fn in filenames:
            if fn.lower().endswith(exts):
                res.append(os.path.join(dirpath, fn))
    return res


def cmd_sources(cfg, path, agenda, as_json=False):
    root = path or "."
    docs = _walk_docs(root)
    chair = [d for d in docs if re.search(r"chair\s*notes.*_eom", d, re.I) or re.search(r"_eom\d*\.", d, re.I)]
    # FL summary 夹
    fl_dir = None
    for dp, dn, fn in os.walk(root):
        if os.path.basename(dp).lower() == "fl summary":
            fl_dir = dp
            break
    fl_files = []
    if fl_dir:
        for dp, dn, fn in os.walk(fl_dir):
            for f in fn:
                if f.lower().endswith(DOC_EXT):
                    fl_files.append(os.path.join(dp, f))
    # WI 文档：不在 fl summary、不是 chair、不是 by_agenda 的 doc/docx
    by_agenda = []
    wi_docs = []
    for d in docs:
        low = d.lower().replace("\\", "/")
        if "fl summary" in low:
            continue
        if re.search(r"_eom\d*\.", d, re.I):
            continue
        if "/by_agenda/" in low or "by_agenda" in low.split("/"):
            by_agenda.append(d)
            continue
        if low.endswith((".doc", ".docx")) and "rp-" in os.path.basename(d).lower():
            wi_docs.append(d)
        elif low.endswith((".doc", ".docx", ".md")) and not re.search(r"r1-\d", os.path.basename(d).lower()):
            wi_docs.append(d)
    tdoc_files = by_agenda
    summary = {
        "root": os.path.abspath(root),
        "chairman_note": [os.path.relpath(x, root) for x in chair],
        "fl_summary_dir": fl_dir,
        "fl_summary_count": len(fl_files),
        "fl_summary_files": [os.path.relpath(x, root) for x in fl_files],
        "wi_docs": [os.path.relpath(x, root) for x in wi_docs],
        "tdoc_count": len(tdoc_files),
        "tdoc_files": [os.path.relpath(x, root) for x in tdoc_files],
    }
    if as_json:
        print(json.dumps(summary, ensure_ascii=False, indent=2))
    else:
        print("== 素材清点 ==")
        print("主席报告（★ 唯一落槌）：%d 份" % len(chair))
        for x in chair:
            print("  - %s" % x)
        print("FL Summary：%d 份（夹：%s）" % (len(fl_files), fl_dir))
        print("WI/SID 文档：%d 份" % len(wi_docs))
        for x in wi_docs:
            print("  - %s" % x)
        print("公司 TDoc（递归探测）：%d 篇" % len(tdoc_files))
        if not chair:
            print("[ERROR] 没有找到主席报告（Chair notes …_eom*.docx）—— 落槌结论无从取得，A2 阻塞")
        if not fl_files:
            print("[ERROR] 没有找到 FL Summary 文件夹或文件 —— A4 阻塞")
        if not wi_docs:
            print("[WARN] 没有识别到 WI/SID 文档 —— A1 阻塞（请用 --wi 显式指定）")
    return summary


# ---------------------------------------------------------------------------
# init：抽 Objective 段 + 识别 WID/SID + 生成骨架
# ---------------------------------------------------------------------------

WID_SIGNALS = [
    ("New WID", "WID"), ("Work Item Description", "WID"), ("Objective of Core part WI", "WID"),
    ("Objective of Core part WI".lower(), "WID"), ("Specify the following enhancements", "WID"),
    ("Specify the following", "WID"),
]
SID_SIGNALS = [
    ("New SID", "SID"), ("Study Item Description", "SID"), ("Objective of Core part SI", "SID"),
    ("Objective of Core part SI".lower(), "SID"), ("Study the", "SID"), ("Evaluate the", "SID"),
]


def _identify_wid_sid(text, filename=""):
    tl = (text or "").lower()
    fl = (filename or "").lower()
    # 优先级 0：文件名（RP 提案文件名常含 'New WID' / 'New SID'）
    for pat, kind in WID_SIGNALS + SID_SIGNALS:
        if pat.lower() in fl:
            return kind
    # 优先级 1：正文/标题级信号
    for pat, kind in WID_SIGNALS + SID_SIGNALS:
        if pat.lower() in tl:
            return kind
    # 优先级 2：交付物性质
    if re.search(r"\bRel-\d+\b", text) and "specification" in tl:
        return "WID"
    if "technical report" in tl or ("study" in tl and "specify" not in tl):
        return "SID"
    return "UNKNOWN"


def _extract_objective(paras):
    """定位 Objective 段（SI/WI 核心），返回 (objective_text_lines, scope_items)。
    scope_items：Objective 段内带负责 WG 标注 [RANx] 的子项；多行合并到同一项的 [RANx] 标签处为止。"""
    # 找到第一个 'objective of' 行作为起点
    start = None
    for i, line in enumerate(paras):
        if re.search(r"objective\s+of", line, re.I):
            start = i
            break
    if start is None:
        return [], []
    # 段末：下一主节（Justification / Annex / References / At the first / Conclusion）
    # 或 WID 模板的编号小节（'N. Section'，排除 'N. Objective …'）；模板的
    # 'RAN time budget' / 'Work item leadership' / 'Supporting …' 等均在 Objective 之后。
    end = len(paras)
    for i in range(start + 1, len(paras)):
        ln = paras[i].strip()
        if re.match(r"^(justification|annex|references|at the first|conclusion)\b", ln, re.I):
            end = i
            break
        if re.match(r"^\d+\.\s+(?!objective)", ln, re.I):
            end = i
            break
        if re.match(r"^(the excel table|work item leadership|supporting individual members|ran time budget|aspects that involve)", ln, re.I):
            end = i
            break
    obj_lines = paras[start:end]
    # 抽 scope items：每个带 [RANx] 标签的行闭合一个条目；无标签行向上一项续接；
    # 子标题（'Objective of …' / 'Specify the following enhancements' / 'For …, specify …'）与
    # lead-in（'For …, ' 无标签）/ NOTE 模板 / RAN time budget 模板 不计入条目。
    items = []
    cur = None
    for ln in obj_lines:
        low = ln.strip()
        if not low:
            continue
        if re.search(r"objective\s+of", low, re.I):
            if cur:
                items.append(cur)
                cur = None
            continue
        if (re.search(r"specify the following enhancements", low, re.I)
                or re.match(r"^for\b.*,\s*specify", low, re.I)
                or (re.match(r"^for\b", low, re.I) and not re.search(r"\[RAN\d", low))
                or (re.match(r"^note\b", low, re.I) and not re.search(r"\[RAN\d", low))
                or "ran time budget" in low.lower()
                or "not applicable" in low.lower()):
            if cur:
                items.append(cur)
                cur = None
            continue
        has_tag = bool(re.search(r"\[RAN\d", low, re.I))
        if has_tag:
            if cur:
                items.append(cur)
            cur = low
        else:
            cur = low if cur is None else cur + " " + low
    if cur:
        items.append(cur)
    # 清理 + 记录 WG + 去重
    uniq = []
    seen = set()
    for it in items:
        wg = sorted({("RAN" + w.upper()) for w in re.findall(r"\[RAN(\d)\]", it, re.I)})
        text = re.sub(r"\[RAN\d\]", "", it).strip(" -•·\t")
        text = re.sub(r"\s+", " ", text).strip()
        k = _norm(text)
        if k and len(text) >= 6 and k not in seen:
            seen.add(k)
            uniq.append({"text": text, "wg": wg})
    return obj_lines, uniq


def cmd_init(cfg, wi_path, meeting, agenda, out_dir, skeleton, out, as_json=False):
    paras = read_doc(wi_path)
    text = "\n".join(paras)
    wid_type = _identify_wid_sid(text, os.path.basename(wi_path))
    obj_lines, items = _extract_objective(paras)
    wid = {
        "wi_doc": os.path.basename(wi_path),
        "type": wid_type,
        "objective_text": "\n".join(obj_lines),
        "items": [{"id": "WI-%d" % (i + 1), "text": it["text"], "wg": it["wg"],
                   "source": "Objective section (extracted)"} for i, it in enumerate(items)],
    }
    wd = _work_dir(out_dir, agenda)
    _dump(os.path.join(wd, "wid.json"), wid)

    # 生成骨架：替换 references/skeleton.md 占位符
    skel_path = skeleton or os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                         "..", "references", "skeleton.md")
    skel_path = os.path.abspath(skel_path)
    if os.path.exists(skel_path):
        skel = _read_text_file(skel_path)
        repl = {
            "{{MEETING}}": meeting or "{{MEETING}}",
            "{{AGENDA}}": agenda or "{{AGENDA}}",
            "{{TOPIC}}": "{{TOPIC}}",
            "{{WI_DOC}}": os.path.basename(wi_path),
            "{{WI_TYPE}}": wid_type,
            "{{DATE}}": "",
            "{{CHAIR_DOC}}": "{{CHAIR_DOC}}",
            "{{FL_COUNT}}": "{{FL_COUNT}}",
            "{{ISSUE_COUNT}}": "{{ISSUE_COUNT}}",
            "{{BLOCK_COUNT}}": "{{BLOCK_COUNT}}",
            "{{WI_COUNT}}": str(len(wid["items"])),
        }
        for k, v in repl.items():
            skel = skel.replace(k, v)
        # 预填 WI 条目节（在 §3 占位处插入）
        wi_sections = []
        for it in wid["items"]:
            wi_sections.append("### %s｜%s            ← 条目表述从 WI/SID 文档 Objective 段原文提取\n\n"
                               "**WI 小结**：（待 A6b 判定：收口 / 部分收口 / 未完）\n" % (it["id"], it["text"]))
        wi_block = "\n".join(wi_sections)
        # 把骨架里 §3 的占位行 "### WI-<条目编号>｜<WI 条目原文表述> …" 整行替换为生成的 WI 条目节
        skel = re.sub(r"### WI-<条目编号>.*",
                      lambda m: wi_block.rstrip(), skel, count=1)
        target = out or os.path.join(wd, "%s_agenda_%s_进展.md" % (meeting or "MEETING", agenda or "x"))
        with open(target, "w", encoding="utf-8") as f:
            f.write(skel)
        skeleton_out = target
    else:
        skeleton_out = None

    result = {"wid_type": wid_type, "item_count": len(wid["items"]),
              "wid_json": os.path.join(wd, "wid.json"), "skeleton": skeleton_out}
    if as_json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print("== init ==")
        print("WI/SID 文档：%s" % wi_path)
        print("识别类型：%s" % wid_type)
        print("Objective 段 scope items：%d 条" % len(wid["items"]))
        for it in wid["items"]:
            print("  %s  [%s]  %s" % (it["id"], "/".join(it["wg"]) or "-", it["text"][:80]))
        print("wid.json → %s" % result["wid_json"])
        if skeleton_out:
            print("骨架 → %s（占位符已填，WI 名称来自 Objective 段原文）" % skeleton_out)
        if wid_type == "UNKNOWN":
            print("[WARN] 未能自主识别 WID/SID，请检查文档标题/正文关键词")
    return wid


# ---------------------------------------------------------------------------
# extract：docx/doc/txt/md → 纯文本段落
# ---------------------------------------------------------------------------

def cmd_extract(path, out=None, head=0, as_json=False):
    paras = read_doc(path)
    text = "\n".join(paras)
    lost = []
    if os.path.splitext(path)[1].lower() == ".docx":
        lost.append("公式(oMath)/表格结构/批注：仅抽 document.xml 段落，涉及公式处需回 document.xml 核对")
    if out:
        with open(out, "w", encoding="utf-8") as f:
            f.write(text)
        print("已写出 %d 行 → %s" % (len(paras), out))
    else:
        shown = paras if head <= 0 else paras[:head]
        print("\n".join(shown))
    if lost:
        print("（已知丢失：%s）" % "; ".join(lost), file=sys.stderr)
    if as_json:
        print(json.dumps({"lines": len(paras), "lost": lost}, ensure_ascii=False))
    return text


# ---------------------------------------------------------------------------
# blocks：结论块切块 + 定轮次 + FFS
# ---------------------------------------------------------------------------

DEFAULT_MARKERS = ["Agreement", "Working assumption", "Observation", "Note", "Conclusion"]
FL_SEP_RE = re.compile(r"^R1-\d+\s+FL Summary\s+#(\d+)", re.I)
BLOCK_START_RE = re.compile(r"^\s*(Agreement|Working assumption|Observation|Note|Conclusion)\b[ \t]*:?", re.I)


def build_blocks(text, markers=None):
    lines = text.splitlines()
    markers = markers or DEFAULT_MARKERS
    # 结论块起始标记正则（按 markers）
    mk_alt = "|".join(re.escape(m) for m in markers)
    start_re = re.compile(r"^\s*(%s)\b[ \t]*:?" % mk_alt, re.I)
    # 找所有 FL 分隔行
    seps = []
    for i, ln in enumerate(lines):
        m = FL_SEP_RE.match(ln.strip())
        if m:
            seps.append((i, int(m.group(1))))
    # 找所有块起始
    starts = []
    for i, ln in enumerate(lines):
        if start_re.match(ln):
            starts.append(i)
    blocks = []
    n = len(starts)
    for idx, si in enumerate(starts):
        # 结束 = 下一个块起点 或 下一个 FL 分隔行（取较早），且 > si
        end = len(lines)
        if idx + 1 < n:
            end = min(end, starts[idx + 1])
        for (spi, _) in seps:
            if spi > si and spi < end:
                end = spi
        # 当前块所属轮次 = 最近的分隔行。
        # 真实主席报告不内嵌 "R1-xx FL Summary #N" 分隔行，故块不属任何轮次
        # （轮次来自 FL 文档本身，由 issues 按文件顺序赋 round，见 cmd_issues）。
        fl_round = None
        for (spi, rnd) in seps:
            if spi < si:
                fl_round = rnd
            else:
                break
        block_lines = lines[si:end]
        btext = "\n".join(block_lines).strip()
        # 抽 FFS
        ffs = [l.strip() for l in block_lines if re.search(r"\bFFS\b\s*:|to discuss", l, re.I)]
        kind = start_re.match(lines[si]).group(1).capitalize()
        blocks.append({
            "id": "B%d" % (idx + 1),
            "kind": kind,
            "line_start": si + 1,
            "line_end": end,
            "fl_round": fl_round,
            "text": btext,
            "ffs": ffs,
        })
    return blocks, seps


def cmd_blocks(verbatim, out_dir, agenda, markers=None, as_json=False):
    text = _read_text_file(verbatim)
    blocks, seps = build_blocks(text, markers)
    wd = _work_dir(out_dir, agenda)
    # FFS 汇总
    ffs_all = []
    for b in blocks:
        for f in b["ffs"]:
            ffs_all.append({"block": b["id"], "fl_round": b["fl_round"], "text": f})
    _dump(os.path.join(wd, "blocks.json"), blocks)
    _dump(os.path.join(wd, "ffs.json"), ffs_all)
    # 轮次单调检查
    rounds = [b["fl_round"] for b in blocks]
    numeric = [r for r in rounds if r is not None]
    mono = all(numeric[i] <= numeric[i + 1] for i in range(len(numeric) - 1))
    if as_json:
        print(json.dumps({"blocks": len(blocks), "ffs": len(ffs_all),
                          "rounds": rounds, "rounds_monotonic": mono}, ensure_ascii=False, indent=2))
    else:
        print("== blocks ==")
        print("结论块：%d 块" % len(blocks))
        for b in blocks:
            print("  %s [%s] round=%s L%d-%d  FFS=%d" % (b["id"], b["kind"], (b["fl_round"] if b["fl_round"] is not None else "N/A"), b["line_start"], b["line_end"], len(b["ffs"])))
        print("FL 分隔行：%d 个；轮次单调不回头：%s" % (len(seps), "是" if mono else "否（⚠ 可能混入别的议程项，回 A2 重裁）"))
        print("blocks.json / ffs.json → %s" % wd)
    if not mono:
        print("[WARN] 轮次回头；请确认 _eom_verbatim.txt 只含目标议程项", file=sys.stderr)
    return blocks


# ---------------------------------------------------------------------------
# issues：FL 多份 → Topic#/Issue# 树
# ---------------------------------------------------------------------------

TOPIC_RE = re.compile(r"^\s*Topic\s*#\s*(\d+)\s*[:\-]", re.I)
ISSUE_RE = re.compile(r"^\s*Issue\s*#\s*(\d+)-(\d+)\s*[:\-]", re.I)
PROP_RE = re.compile(r"proposal", re.I)
COMPANIES_RE = re.compile(r"(companies|proponents)\s*:", re.I)


def split_issues(text):
    lines = text.splitlines()
    # 切分 Topic/Issue 标题点
    headers = []  # (idx, type, a, b, title)
    for i, ln in enumerate(lines):
        mt = TOPIC_RE.match(ln)
        if mt:
            headers.append((i, "topic", int(mt.group(1)), 0, ln))
            continue
        mi = ISSUE_RE.match(ln)
        if mi:
            headers.append((i, "issue", int(mi.group(1)), int(mi.group(2)), ln))
    issues = []
    for hidx, h in enumerate(headers):
        si = h[0]
        ei = headers[hidx + 1][0] if hidx + 1 < len(headers) else len(lines)
        seg = lines[si:ei]
        seg_text = "\n".join(seg)
        if h[1] != "issue":
            continue
        has_position = bool(COMPANIES_RE.search(seg_text))
        proposals = [l.strip() for l in seg if PROP_RE.search(l)]
        ffs = [l.strip() for l in seg if re.search(r"\bFFS\b\s*:|to discuss", l, re.I)]
        issues.append({
            "id": "Issue#%d-%d" % (h[2], h[3]),
            "title": h[4].strip(),
            "text": seg_text,
            "has_position": has_position,
            "proposals": proposals,
            "ffs": ffs,
        })
    return issues


def cmd_issues(files, out_dir, agenda, as_json=False):
    per_file = []
    merged = []
    for fi, fp in enumerate(files):
        paras = read_doc(fp)
        text = "\n".join(paras)
        iss = split_issues(text)
        for x in iss:
            x["round"] = fi + 1  # 轮次 = FL 文档顺序（早→晚）
        per_file.append({"file": os.path.basename(fp), "issues": iss, "round": fi + 1})
        merged = iss  # 以最末一份为骨架基准
    # 跨轮 diff（与倒数第二份比）
    diff = []
    if len(per_file) >= 2:
        prev = {x["id"]: x for x in per_file[-2]["issues"]}
        cur = {x["id"]: x for x in merged}
        for iid in cur:
            if iid not in prev:
                diff.append({"kind": "added", "id": iid})
        for iid in prev:
            if iid not in cur:
                diff.append({"kind": "removed", "id": iid})
    wd = _work_dir(out_dir, agenda)
    _dump(os.path.join(wd, "issues.json"), {"merged": merged, "per_file": per_file, "diff": diff})
    if as_json:
        print(json.dumps({"issue_count": len(merged), "diff": diff}, ensure_ascii=False, indent=2))
    else:
        print("== issues ==")
        print("最末 FL 基准 Issue 数：%d" % len(merged))
        for x in merged:
            print("  %s  立场=%s  提案段=%d  FFS=%d" % (x["id"], "有" if x["has_position"] else "无", len(x["proposals"]), len(x["ffs"])))
        if diff:
            print("跨轮 diff：%s" % "; ".join("%s:%s" % (d["kind"], d["id"]) for d in diff))
        print("issues.json → %s" % wd)
    return merged


# ---------------------------------------------------------------------------
# flow：块→Issue 候选 + 串讨论链 + final_status
# ---------------------------------------------------------------------------

def cmd_flow(blocks_json, issues_json, out_dir, agenda, as_json=False, min_score=0.02):
    blocks = _load(blocks_json)
    iss_data = _load(issues_json)
    merged = iss_data.get("merged", iss_data if isinstance(iss_data, list) else [])
    per_file = iss_data.get("per_file", [])
    flow = []
    for b in blocks:
        cands = []
        for it in merged:
            # 用 Issue 全文匹配（提案段单独太稀疏，n-gram 重叠≈0，会漏掉真实关联）
            prop_text = it.get("text", "")
            sc = _overlap(b["text"], prop_text)
            if sc >= min_score:
                cands.append({"issue": it["id"], "score": round(sc, 4), "round": it.get("round")})
        cands.sort(key=lambda x: -x["score"])
        top = cands[:3]
        # 讨论链：候选 Issue 在哪些 FL 轮次被讨论（轮次来自 FL 文档，见 cmd_issues）
        discussion_rounds = sorted({c["round"] for c in top if c.get("round")})
        # rounds：在各 FL 文件里是否出现该 issue（保留，含 round 号）
        rounds = []
        for pf in per_file:
            present = any(x["id"] == (top[0]["issue"] if top else None) for x in pf["issues"]) if top else False
            rounds.append({"file": pf["file"], "round": pf.get("round"), "issue_present": present})
        final_status = "未收口" if b.get("ffs") else "落槌"
        if not top:
            final_status = "未收口"
        flow.append({
            "block_id": b["id"],
            "kind": b["kind"],
            "fl_round": b["fl_round"],   # 主席报告结论不属轮次（None）
            "discussion_rounds": discussion_rounds,
            "candidates": top,
            "needs_model": (len(top) == 0) or (len(top) >= 2 and top[0]["score"] - top[1]["score"] < 0.02),
            "rounds": rounds,
            "final_status": final_status,
            "delta": "",  # 由模型填写
            "tdoc_reason": "",  # 若抽 TDoc 须填触发条件
        })
    wd = _work_dir(out_dir, agenda)
    _dump(os.path.join(wd, "flow.json"), flow)
    if as_json:
        print(json.dumps({"blocks": len(flow), "needs_model": sum(1 for f in flow if f["needs_model"])}, ensure_ascii=False, indent=2))
    else:
        print("== flow ==")
        print("结论块 → Issue 候选：%d 块" % len(flow))
        for f in flow:
            cs = ", ".join("%s(%.3f)" % (c["issue"], c["score"]) for c in f["candidates"]) or "UNMAPPED"
            print("  %s [%s] → %s  状态=%s  %s" % (f["block_id"], f["kind"], cs, f["final_status"],
                                                  "需模型判定" if f["needs_model"] else ""))
        print("flow.json → %s" % wd)
    return flow


# ---------------------------------------------------------------------------
# rel：Issue → WI 条目挂接候选（必须）
# ---------------------------------------------------------------------------

def cmd_rel(flow_json, wid_json, issues_json, out_dir, agenda, as_json=False):
    flow = _load(flow_json)
    wid = _load(wid_json)
    items = wid.get("items", [])
    # 加载 issues（显式 --issues 优先；否则从 flow 同级目录猜 issues.json）
    idata = None
    if issues_json and os.path.exists(issues_json):
        idata = _load(issues_json)
    else:
        ip = os.path.join(os.path.dirname(os.path.abspath(flow_json)), "issues.json")
        if os.path.exists(ip):
            idata = _load(ip)
    merged = idata.get("merged", []) if idata else []
    issue_by_id = {it["id"]: it for it in merged}
    # 收集所有 issue id：flow 候选 + issues 全量（每个 Issue 都必须有 WI 归属）
    issue_ids = set()
    for f in flow:
        for c in f.get("candidates", []):
            issue_ids.add(c["issue"])
    for it in merged:
        issue_ids.add(it["id"])
    rel = []
    for iid in sorted(issue_ids):
        it = issue_by_id.get(iid, {})
        # 用 Issue 全文 + 提案段匹配 WI 条目（不是用 Issue ID 字符串）
        match_text = (it.get("text", "") or "") + "\n" + "\n".join(it.get("proposals", []))
        cands = []
        for wi in items:
            # 以 WI 条目为正样本：看其 4-gram 有多少出现在 Issue 内容里（recall），
            # 比 Jaccard 更适合"短 WI 短语 ↔ 长讨论"，避免稀疏导致全 0
            gw = _grams(wi["text"]); gi = _grams(match_text)
            sc = (len(gw & gi) / len(gw)) if gw else 0.0
            if sc > 0:
                cands.append({"wi": wi["id"], "score": round(sc, 4)})
        cands.sort(key=lambda x: -x["score"])
        top = cands[:2]
        rel.append({
            "issue": iid,
            "issue_title": it.get("title", ""),
            "candidates": top,
            "wi": top[0]["wi"] if top else None,
            "needs_model": (len(top) == 0) or (len(top) >= 2 and top[0]["score"] - top[1]["score"] < 0.02),
            "derived": False,  # 是否派生新条目 N 系列，由模型判定
            "derived_reason": "",
        })
    wd = _work_dir(out_dir, agenda)
    _dump(os.path.join(wd, "rel.json"), rel)
    unmapped = [r["issue"] for r in rel if r["needs_model"]]
    if as_json:
        print(json.dumps({"issues": len(rel), "unmapped": unmapped}, ensure_ascii=False, indent=2))
    else:
        print("== rel ==")
        print("Issue → WI 挂接候选：%d 个 Issue" % len(rel))
        for r in rel:
            cs = ", ".join("%s(%.3f)" % (c["wi"], c["score"]) for c in r["candidates"]) or "UNMAPPED"
            print("  %s → %s  %s" % (r["issue"], cs, "需模型判定" if r["needs_model"] else ""))
        if unmapped:
            print("[WARN] 仍有 %d 个 UNMAPPED，须模型在 rel.json 中手动填 wi / 派生依据" % len(unmapped))
        print("rel.json → %s" % wd)
    return rel


# ---------------------------------------------------------------------------
# tdoc：按需抽单篇
# ---------------------------------------------------------------------------

def cmd_tdoc(path, out, head=0, as_json=False):
    paras = read_doc(path)
    text = "\n".join(paras)
    if out:
        with open(out, "w", encoding="utf-8") as f:
            f.write(text)
        print("已抽出 %d 行 → %s" % (len(paras), out))
    else:
        shown = paras if head <= 0 else paras[:head]
        print("\n".join(shown))
    if as_json:
        print(json.dumps({"lines": len(paras)}, ensure_ascii=False))
    return text


# ---------------------------------------------------------------------------
# compose：按 WI 条目组装（幂等 + 块数守恒 + 正文优先）
# ---------------------------------------------------------------------------

def cmd_compose(rel_json, blocks_json, target, out, as_json=False):
    rel = _load(rel_json) if rel_json else []
    blocks = _load(blocks_json) if blocks_json else []
    # 不覆盖已有内容：若 target 已存在，只填占位符 + 报告，不重写已写部分
    existing = ""
    if target and os.path.exists(target):
        existing = _read_text_file(target)
    # 生成 §3 结构（作为建议/补充）
    wi_order = []
    issue_to_wi = {}
    for r in rel:
        issue_to_wi[r["issue"]] = r.get("wi")
    # 按 WI 分组 Issue
    groups = {}
    for r in rel:
        wi = r.get("wi") or "UNMAPPED"
        groups.setdefault(wi, []).append(r["issue"])
    sec3_lines = ["## §3　按 WI 条目的本场进展（每个 WI 条目一节）", ""]
    for wi, iss in groups.items():
        sec3_lines.append("### %s　（WI 小结：待 A6b 判定）" % wi)
        sec3_lines.append("")
        for iid in iss:
            sec3_lines.append("#### %s" % iid)
            sec3_lines.append("")
            sec3_lines.append("**① 讨论脉络**   ← FL#1→FL#k 各轮立场演进 + delta（不写「达成了」）")
            sec3_lines.append("**② 结论**       ← 对应 Chairman note 结论块【原文】逐字 + 【解读】+ 【映射】")
            sec3_lines.append("**③ 未结论**     ← 无结论则标「未收口 / FFS / 后续轮次」")
            sec3_lines.append("")
    sec3 = "\n".join(sec3_lines)
    if target and os.path.exists(target):
        # 填占位符
        filled = re.sub(r"\{\{(MEETING|AGENDA|WI_DOC|WI_TYPE|FL_COUNT|ISSUE_COUNT|BLOCK_COUNT|WI_COUNT|CHAIR_DOC|TOPIC|DATE)\}\}",
                        lambda m: _placeholder_fill(m.group(1), rel, blocks), existing)
        with open(target, "w", encoding="utf-8") as f:
            f.write(filled)
        print("已就位（仅填占位符，未覆盖已写内容）→ %s" % target)
        print("[提示] §3 建议结构已生成，请按需插入：\n%s" % ("（见上方 stdout 的 sec3 预览）" if as_json else ""))
        if as_json:
            print(json.dumps({"sec3": sec3, "filled_placeholders": True}, ensure_ascii=False, indent=2))
        else:
            print("----- §3 建议结构（按需插入，勿复制已写内容）-----")
            print(sec3)
    else:
        out_path = out or target or "进展文档.md"
        content = sec3 + "\n"
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(content)
        print("已生成 §3 结构 → %s" % out_path)
    return sec3


def _placeholder_fill(name, rel, blocks):
    if name == "BLOCK_COUNT":
        return str(len(blocks))
    if name == "ISSUE_COUNT":
        return str(len(rel))
    if name == "WI_COUNT":
        return str(len(set(r.get("wi") for r in rel if r.get("wi"))))
    return ""


# ---------------------------------------------------------------------------
# check：机械校验
# ---------------------------------------------------------------------------

def cmd_check(doc, eom, workdir, as_json=False):
    md = _read_text_file(doc) if doc and os.path.exists(doc) else ""
    errors = []
    warns = []
    # 占位符
    if "{{" in md:
        errors.append("骨架占位符未替换：%s" % ", ".join(set(re.findall(r"\{\{[^}]+\}\}", md))))
    for pat in ("TODO", "待补"):
        if pat in md:
            errors.append("占位符未填：发现 %r" % pat)
    # 块数守恒：blocks.json 中存在 → 每块的【原文】应在 md
    block_count = 0
    if workdir and os.path.exists(os.path.join(workdir, "blocks.json")):
        blocks = _load(os.path.join(workdir, "blocks.json"))
        block_count = len(blocks)
        # 每块的【原文】应能在 md 找到（取前 40 归一化字符做存在性检查）
        for b in blocks:
            snippet = _norm(b["text"])[:40]
            if snippet and snippet not in _norm(md):
                # 可能是正文已清理，允许 FFS 标记差异；降级为 WARN
                warns.append("块 %s 的原文片段未能在文档中逐字定位（可能已清理排版）" % b["id"])
    # eom 切分数对账
    if eom and os.path.exists(eom):
        etext = _read_text_file(eom)
        eom_blocks = len(build_blocks(etext)[0])
        if block_count and eom_blocks != block_count:
            errors.append("块数不一致：blocks.json=%d，eom 切分=%d" % (block_count, eom_blocks))
    # rel 覆盖：无 UNMAPPED
    if workdir and os.path.exists(os.path.join(workdir, "rel.json")):
        rel = _load(os.path.join(workdir, "rel.json"))
        unmapped = [r["issue"] for r in rel if r.get("needs_model") or not r.get("wi")]
        if unmapped:
            errors.append("存在 UNMAPPED Issue（须模型挂接）：%s" % ", ".join(unmapped))
    # 反模式：FL 级 Agree 被写成落槌
    if re.search(r"已同意支持|已达成一致.*FL|FL.*已同意", md):
        warns.append("疑似把 FL 的 Agree 写成落槌（须降级为'仅 FL 级 Agree，未进主席报告'）")
    report = {"errors": errors, "warns": warns, "block_count": block_count}
    if as_json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print("== check ==")
        print("ERROR：%d 条" % len(errors))
        for e in errors:
            print("  [ERROR] %s" % e)
        print("WARN：%d 条" % len(warns))
        for w in warns:
            print("  [WARN] %s" % w)
        print("结论：%s" % ("通过（ERROR=0）" if not errors else "未通过，须修复 ERROR"))
    return report


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def build_parser():
    p = argparse.ArgumentParser(prog="skillctl", description="3gpp-ran-meeting-progress 工具链")
    sub = p.add_subparsers(dest="cmd")

    sp = sub.add_parser("sources", help="清点一场会议的素材夹（主席报告 / FL / WI / TDoc 递归）")
    sp.add_argument("path", nargs="?", default=".", help="素材夹根目录")
    sp.add_argument("--agenda", default="")
    sp.add_argument("--json", action="store_true")

    sp = sub.add_parser("init", help="A1：抽 Objective 段 + 识别 WID/SID + 生成骨架")
    sp.add_argument("--wi", required=True, help="WI/SID 文档路径")
    sp.add_argument("--meeting", default="")
    sp.add_argument("--agenda", default="")
    sp.add_argument("--out-dir", default="", help="输出目录（wid.json + 骨架落此）；不填则落到默认 _work_<agenda>")
    sp.add_argument("--out", default="", help="骨架 .md 显式路径（覆盖 --out-dir 下的默认名）")
    sp.add_argument("--skeleton", default="")
    sp.add_argument("--json", action="store_true")

    sp = sub.add_parser("extract", help="A2：docx/doc/txt/md → 纯文本段落")
    sp.add_argument("path")
    sp.add_argument("--out", default="")
    sp.add_argument("--head", type=int, default=0)
    sp.add_argument("--json", action="store_true")

    sp = sub.add_parser("blocks", help="A3：结论块切块 + 定轮次 + FFS")
    sp.add_argument("verbatim")
    sp.add_argument("--out-dir", default="")
    sp.add_argument("--agenda", default="")
    sp.add_argument("--markers", default="")
    sp.add_argument("--json", action="store_true")

    sp = sub.add_parser("issues", help="A4：抽 Issue 树（顺序早→晚）")
    sp.add_argument("files", nargs="+")
    sp.add_argument("--out-dir", default="")
    sp.add_argument("--agenda", default="")
    sp.add_argument("--json", action="store_true")

    sp = sub.add_parser("flow", help="A5：块→Issue 候选 + 串讨论链 + final_status")
    sp.add_argument("--blocks", required=True)
    sp.add_argument("--issues", required=True)
    sp.add_argument("--out-dir", default="")
    sp.add_argument("--agenda", default="")
    sp.add_argument("--min-score", type=float, default=0.02, help="块→Issue 候选的最低重叠分（越低越宽松）")
    sp.add_argument("--json", action="store_true")

    sp = sub.add_parser("rel", help="A5b：Issue→WI 挂接候选（必须）")
    sp.add_argument("--flow", required=True)
    sp.add_argument("--wid", required=True)
    sp.add_argument("--issues", default="", help="issues.json 路径（显式指定；不填则从 flow 同级目录找）")
    sp.add_argument("--out-dir", default="")
    sp.add_argument("--agenda", default="")
    sp.add_argument("--json", action="store_true")

    sp = sub.add_parser("tdoc", help="按需抽取单篇 TDoc")
    sp.add_argument("path")
    sp.add_argument("--out", default="")
    sp.add_argument("--head", type=int, default=0)
    sp.add_argument("--json", action="store_true")

    sp = sub.add_parser("compose", help="A7：按 WI 条目组装（幂等、块数守恒）")
    sp.add_argument("--rel", default="")
    sp.add_argument("--blocks", default="")
    sp.add_argument("--target", default="")
    sp.add_argument("--out", default="")
    sp.add_argument("--json", action="store_true")

    sp = sub.add_parser("check", help="A8：机械校验")
    sp.add_argument("doc", nargs="?", default="")
    sp.add_argument("--eom", default="")
    sp.add_argument("--workdir", default="")
    sp.add_argument("--json", action="store_true")

    return p


def main(argv=None):
    argv = argv or sys.argv[1:]
    p = build_parser()
    args = p.parse_args(argv)
    # 读取同目录 config.json（若有），目前仅用于 markers 等；此处轻量加载
    cfg = {}
    here = os.path.dirname(os.path.abspath(__file__))
    cfg_path = os.path.join(here, "..", "config.json")
    if os.path.exists(cfg_path):
        try:
            cfg = _load(os.path.abspath(cfg_path))
        except Exception:
            cfg = {}
    if args.cmd == "sources":
        cmd_sources(cfg, args.path, args.agenda, args.json)
    elif args.cmd == "init":
        cmd_init(cfg, args.wi, args.meeting, args.agenda, args.out_dir, args.skeleton, args.out, args.json)
    elif args.cmd == "extract":
        cmd_extract(args.path, args.out, args.head, args.json)
    elif args.cmd == "blocks":
        markers = args.markers.split(",") if args.markers else None
        cmd_blocks(args.verbatim, args.out_dir, args.agenda, markers, args.json)
    elif args.cmd == "issues":
        cmd_issues(args.files, args.out_dir, args.agenda, args.json)
    elif args.cmd == "flow":
        cmd_flow(args.blocks, args.issues, args.out_dir, args.agenda, args.json, args.min_score)
    elif args.cmd == "rel":
        cmd_rel(args.flow, args.wid, args.issues, args.out_dir, args.agenda, args.json)
    elif args.cmd == "tdoc":
        cmd_tdoc(args.path, args.out, args.head, args.json)
    elif args.cmd == "compose":
        cmd_compose(args.rel, args.blocks, args.target, args.out, args.json)
    elif args.cmd == "check":
        cmd_check(args.doc, args.eom, args.workdir, args.json)
    else:
        p.print_help()
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
