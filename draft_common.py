"""
draft_common.py —— 一稿/二稿批改共用工具
============================================
职责：
1. 从一稿批改 docx 中提取评语（代码解析原始 run，不标准化，避免破坏格式标记）
2. 按文件名匹配一稿批改文件与二稿文件
3. 目录发现与增量扫描辅助

设计原则：
- 评语提取直接解析原始 run 结构，用颜色 + 【】识别评语，不跑 normalize_document_runs
  （normalize 会合并 run 并剥离格式标记，破坏评语的颜色/下划线/高亮信息）
- 评语与原文的绑定：评语 run 紧跟在被标记的原文 run 之后，取前一个原文 run 作为对应句子
"""

import re
from pathlib import Path
from typing import Optional

from docx import Document
from docx.oxml.ns import qn

# ============================================================
# 颜色常量（与 tools.py 一致）
# ============================================================
COLOR_GREEN = "00B050"   # 肯定评语
COLOR_RED = "EE0000"     # 纠错评语
COLOR_BLUE = "0070C0"    # 总评
COLOR_BLACK = "000000"   # 原文

# 评语颜色集合（用于识别评语 run）
COMMENT_COLORS = {COLOR_GREEN, COLOR_RED, COLOR_BLUE}


# ============================================================
# 评语提取：从一稿批改 docx 中提取评语
# ============================================================

def _run_color(run) -> Optional[str]:
    """获取 run 的字体颜色（十六进制，大写），无则返回 None"""
    rPr = run._element.rPr
    if rPr is None:
        return None
    c = rPr.find(qn("w:color"))
    if c is None:
        return None
    val = c.get(qn("w:val"))
    return val.upper() if val else None


def _run_underline(run) -> Optional[str]:
    """获取 run 的下划线类型，无则返回 None"""
    rPr = run._element.rPr
    if rPr is None:
        return None
    u = rPr.find(qn("w:u"))
    if u is None:
        return None
    return u.get(qn("w:val"))


def _run_highlight(run) -> Optional[str]:
    """获取 run 的高亮颜色，无则返回 None"""
    rPr = run._element.rPr
    if rPr is None:
        return None
    hl = rPr.find(qn("w:highlight"))
    if hl is None:
        return None
    return hl.get(qn("w:val"))


def _is_comment_run(run) -> bool:
    """判断 run 是否为本工具追加的评语（【开头或红/绿/蓝字色）"""
    text = (run.text or "").strip()
    if not text:
        return False
    if text.startswith("【"):
        return True
    color = _run_color(run)
    if color in COMMENT_COLORS:
        return True
    return False


def _is_final_review_para(text: str) -> bool:
    """判断段落是否属于总评部分（【总评】之后的所有蓝色段落）"""
    t = (text or "").strip()
    return t.startswith("【总评】")


def extract_comments_from_graded(docx_path: str) -> list[dict]:
    """
    从一稿批改 docx 中提取所有评语（含对应原文句子）。

    返回评语列表，每项：
    {
        "paragraph_index": 段落全文索引,
        "original_text": 评语对应的原文句子（评语 run 前的原文 run 拼接）,
        "comment": 评语内容（去掉【】）,
        "color": "green"/"red"/"blue",
        "raw_comment": 原始评语文本（含【】）,
        "has_underline": 原文是否有下划线（肯定标记）,
        "has_highlight": 原文是否有高亮（纠错标记）,
    }

    设计说明：
    - 直接解析原始 run，不标准化
    - 评语 run 紧跟在被标记的原文 run 之后，取前一个原文 run 作为对应句子
    - 跳过【题目】【材料】等非正文段落和【总评】部分
    """
    from word_agent.docx_reader import open_docx_safe
    doc = open_docx_safe(docx_path)
    comments = []

    for pi, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        # 跳过题目/材料/模板
        if text.startswith("【题目】") or text.startswith("【材料】") or text.startswith("【答案】"):
            continue
        # 跳过总评部分（【总评】及其后所有蓝色段落）
        if _is_final_review_para(text):
            break

        runs = list(para.runs)
        # 收集原文 run 和评语 run 的序列
        # 策略：遍历 run，遇到评语 run 时，向前找最近的原文 run 作为对应句子
        # 注意：一条评语可能被 Word 拆成多个连续 run（如【电头格式有误...】被拆成3段），
        #       需要把连续的评语 run 合并成一条评语
        pending_original = []  # 当前评语前的原文 run 文本累积
        pending_comment_runs = []  # 当前评语的连续 run 累积
        pending_comment_color = None  # 当前评语的颜色

        def _flush_comment():
            """把累积的评语 run 合并成一条评语"""
            nonlocal pending_comment_runs, pending_comment_color
            if not pending_comment_runs:
                return
            comment_text = "".join(r.text or "" for r in pending_comment_runs).strip()
            if not comment_text:
                pending_comment_runs = []
                pending_comment_color = None
                return
            # 去掉【】包裹
            inner = comment_text
            if inner.startswith("【") and inner.endswith("】"):
                inner = inner[1:-1]
            # 判断颜色
            color = pending_comment_color
            if color == COLOR_GREEN:
                color_label = "green"
            elif color == COLOR_RED:
                color_label = "red"
            elif color == COLOR_BLUE:
                color_label = "blue"
            else:
                # 无颜色但以【开头 → 按内容判断（✅ 视为 green）
                if "✅" in inner:
                    color_label = "green"
                else:
                    color_label = "red"

            original_text = "".join(pending_original).strip()
            comments.append({
                "paragraph_index": pi,
                "original_text": original_text,
                "comment": inner,
                "color": color_label,
                "raw_comment": comment_text,
                "has_underline": any(_run_underline(r) not in (None, "none") for r in runs if not _is_comment_run(r)),
                "has_highlight": any(_run_highlight(r) == "yellow" for r in runs if not _is_comment_run(r)),
            })
            # 评语后重置原文累积（评语本身不参与原文）
            pending_original.clear()
            pending_comment_runs = []
            pending_comment_color = None

        for run in runs:
            if _is_comment_run(run):
                # 评语 run：累积（连续评语 run 合并为一条）
                pending_comment_runs.append(run)
                color = _run_color(run)
                if color in COMMENT_COLORS:
                    pending_comment_color = color
            else:
                # 原文 run：先 flush 之前的评语，再累积原文
                _flush_comment()
                pending_original.append(run.text or "")

        # 段落末尾可能还有未 flush 的评语
        _flush_comment()

    return comments


def extract_final_review(docx_path: str) -> list[str]:
    """
    从一稿批改 docx 中提取总评部分（【总评】之后的所有蓝色段落文本）。
    返回段落文本列表。
    """
    from word_agent.docx_reader import open_docx_safe
    doc = open_docx_safe(docx_path)
    lines = []
    in_review = False
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        if text.startswith("【总评】"):
            in_review = True
            lines.append(text)
            continue
        if in_review:
            lines.append(text)
    return lines


# ============================================================
# 文件匹配：一稿批改 ↔ 二稿
# ============================================================

def normalize_stem(stem: str) -> str:
    """
    归一化文件名主干：去掉【已批改】前缀、一稿/二稿后缀、空格、括号序号。
    用于一稿/二稿匹配。
    例：
      "【已批改】01 消息 v27459 一稿" → "01消息v27459"
      "04 消息 v27459 二稿"          → "04消息v27459"
    """
    s = stem
    s = s.replace("【已批改】", "")
    s = re.sub(r"[（(]\d+[)）]", "", s)  # 去掉 (1) 等序号
    s = s.replace("一稿", "").replace("二稿", "")
    s = re.sub(r"\s+", "", s)
    return s


def match_first_draft(second_draft_path: Path, first_draft_dir: Path) -> Optional[Path]:
    """
    根据二稿文件名，在一稿批改目录中匹配对应的一稿批改文件。
    匹配规则：归一化主干相同。
    返回一稿批改文件路径，找不到返回 None。
    """
    if not first_draft_dir.exists():
        return None
    target = normalize_stem(second_draft_path.stem)
    for f in sorted(first_draft_dir.glob("*.docx")):
        if normalize_stem(f.stem) == target:
            return f
    return None


# ============================================================
# 目录发现与增量扫描
# ============================================================

def discover_draft_dirs(package_dir: Path) -> dict:
    """
    扫描 word_agent/ 下所有作业类型目录，发现一稿/二稿批改目录结构。

    新结构（二稿批改）：
      作业类型/二稿批改/
        ├── 一稿批改/        ← 一稿批改后的文件（用户手动放）
        ├── 二稿/            ← 二稿文件
        └── 二稿批改结果/    ← 二稿批改输出

    返回：
    {
        "消息": {
            "first_draft": {
                "homework": Path, "complete_homework": Path, "answer": Path,
            },
            "second_draft": {
                "first_draft_dir": Path,   # 二稿批改/一稿批改/
                "second_draft_dir": Path,  # 二稿批改/二稿/
                "output_dir": Path,        # 二稿批改/二稿批改结果/
            },
        },
        ...
    }
    """
    result = {}
    for sub in sorted(package_dir.iterdir()):
        if not sub.is_dir() or sub.name.startswith("__") or sub.name in ("temp", "standard"):
            continue
        first_dir = sub / "一稿批改"
        second_dir = sub / "二稿批改"
        if not first_dir.exists() and not second_dir.exists():
            continue
        entry = {}
        if first_dir.exists():
            entry["first_draft"] = {
                "homework": first_dir / "homework",
                "complete_homework": first_dir / "complete_homework",
                "answer": first_dir / "Answer",
            }
        if second_dir.exists():
            entry["second_draft"] = {
                "first_draft_dir": second_dir / "一稿批改",
                "second_draft_dir": second_dir / "二稿",
                "output_dir": second_dir / "二稿批改结果",
            }
        result[sub.name] = entry
    return result


def list_ungraded_second_drafts(second_draft_dir: Path, output_dir: Path) -> list[Path]:
    """
    列出二稿目录下所有未批改的二稿文件。
    增量逻辑：递归扫描输出目录（含日期子文件夹），已有【已批改】XX 二稿 则跳过。
    """
    if not second_draft_dir.exists():
        return []
    graded_stems = set()
    if output_dir.exists():
        for f in output_dir.rglob("*.docx"):
            if "【已批改】" in f.stem:
                graded_stems.add(normalize_stem(f.stem))

    result = []
    for f in sorted(second_draft_dir.glob("*.docx")):
        if "【已批改】" in f.stem:
            continue
        if normalize_stem(f.stem) in graded_stems:
            continue
        result.append(f)
    return result
