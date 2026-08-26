"""
Word Agent — 智能 Word 作业批改工具
运行方式：python -m word_agent 或 python main.py
行为：
1. 一稿批改：自动扫描 作业类型/一稿批改/homework/ 目录，增量批改未处理作业
2. 二稿批改：自动扫描 作业类型/二稿批改/{一稿批改,二稿,二稿批改结果}/ 目录，对比一稿评语批改二稿
"""

import os
import sys
from pathlib import Path

# 确保项目根目录在路径中
_project_root = str(Path(__file__).resolve().parent.parent)
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from word_agent.single_shot_grader import run_single_shot_grader
from word_agent.second_draft_grader import run_second_draft_grader
from word_agent.draft_common import (
    discover_draft_dirs,
    list_ungraded_second_drafts,
    match_first_draft,
)


# 项目根目录
PACKAGE_DIR = Path(__file__).resolve().parent

# ---- 作业类型自动发现（一稿批改） ----
HOMEWORK_TYPE_DIRS = {}


def _discover_homework_types():
    """扫描 word_agent/ 下所有子目录，发现作业类型。

    优先新结构：作业类型/一稿批改/{Answer,homework,complete_homework}
    兼容旧结构：作业类型/{Answer,homework,complete_homework}
    """
    global HOMEWORK_TYPE_DIRS
    HOMEWORK_TYPE_DIRS = {}
    for sub in sorted(PACKAGE_DIR.iterdir()):
        if not sub.is_dir() or sub.name.startswith("__") or sub.name in ("temp", "standard", "Answer", "homework", "complete_homework", "completed_homework"):
            continue
        # 新结构：作业类型/一稿批改/{Answer,homework,complete_homework}
        first_dir = sub / "一稿批改"
        answer_dir = first_dir / "Answer"
        hw_dir = first_dir / "homework"
        output_dir = first_dir / "complete_homework"
        if not (answer_dir.exists() and hw_dir.exists()):
            # 兼容旧结构：作业类型/{Answer,homework,complete_homework}
            answer_dir = sub / "Answer"
            hw_dir = sub / "homework"
            output_dir = sub / "complete_homework"
        if answer_dir.exists() and hw_dir.exists():
            type_name = sub.name
            HOMEWORK_TYPE_DIRS[type_name] = {
                "answer": answer_dir,
                "homework": hw_dir,
                "output": output_dir,
            }
    return HOMEWORK_TYPE_DIRS


def _list_docx(directory: Path) -> list[Path]:
    """列出目录下所有未批改的 docx 文件"""
    return sorted(f for f in directory.glob("*.docx") if "【已批改】" not in f.stem and "_v" not in f.stem)


def _is_graded(hw_stem: str, graded_stems: set) -> bool:
    """判断作业是否已被批改"""
    if hw_stem in graded_stems:
        return True
    for gs in graded_stems:
        if gs.startswith(hw_stem):
            return True
    return False


def _match_answer_in_dir(homework_path: Path, answer_dir: Path):
    """
    根据作业文件名自动匹配参考答案 PDF。
    只从当前作业类型的 answer_dir 中匹配（答案独立，不跨类型混淆）。
    匹配优先级：
      1. 作业序号（如文件名开头的 "02"）→ 优先匹配含该序号的 PDF
      2. 中文关键词重合 ≥2 个
      3. 兜底：返回第一个 PDF
    """
    if not answer_dir or not answer_dir.exists():
        return None
    answer_pdfs = sorted(answer_dir.glob("*.pdf"))
    if not answer_pdfs:
        return None
    if len(answer_pdfs) == 1:
        return answer_pdfs[0]

    hw_name = homework_path.stem
    import re

    # 1. 优先匹配作业序号（文件名开头的两位数序号，如 "02"）
    #    只取开头的序号，避免匹配到学号/版本号等干扰数字
    m = re.match(r"^\s*(\d{1,2})", hw_name)
    if m:
        hw_num = m.group(1)
        for pdf in answer_pdfs:
            if hw_num in pdf.stem:
                return pdf

    # 2. 中文关键词重合 ≥2 个
    for pdf in answer_pdfs:
        pdf_keywords = set(re.findall(r"[\u4e00-\u9fff]+", pdf.stem))
        hw_keywords = set(re.findall(r"[\u4e00-\u9fff]+", hw_name))
        if len(pdf_keywords & hw_keywords) >= 2:
            return pdf

    # 3. 兜底：返回第一个 PDF
    return answer_pdfs[0]


def auto_grade():
    """自动扫描所有作业类型，增量批改未处理作业（一稿）"""
    _discover_homework_types()
    if not HOMEWORK_TYPE_DIRS:
        print("📭 未发现作业类型目录（需要 Answer/ + homework/ 子目录）")
        print("   示例: word_agent/评论/Answer/ + word_agent/评论/homework/")
        return

    all_results = []
    for type_name, info in HOMEWORK_TYPE_DIRS.items():
        hw_dir = info["homework"]
        if not hw_dir.exists():
            continue
        hw_files = _list_docx(hw_dir)
        if not hw_files:
            continue

        # 收集已批改的作业名（递归扫描所有日期子文件夹，避免漏判）
        output_dir = info["output"]
        graded_stems = set()
        if output_dir.exists():
            for f in output_dir.rglob("*.docx"):
                if "【已批改】" in f.stem:
                    graded_stems.add(f.stem.replace("【已批改】", ""))

        ungraded = [f for f in hw_files if not _is_graded(f.stem, graded_stems)]
        if not ungraded:
            print(f"\n📂 [{type_name}] 所有作业已批改，无需处理")
            continue

        print(f"\n{'='*60}")
        print(f"📂 [{type_name}] 发现 {len(ungraded)}/{len(hw_files)} 份未批改作业")
        print(f"{'='*60}")

        for hw_path in ungraded:
            print(f"\n{'='*60}")
            print(f"📄 作业: {hw_path.name}")
            answer_pdf = _match_answer_in_dir(hw_path, info["answer"])
            if not answer_pdf:
                print(f"   ⚠️  跳过：{info['answer']} 中找不到匹配的参考答案")
                continue
            print(f"   🔗 参考答案: {answer_pdf.name}")
            result = run_single_shot_grader(
                answer_pdf_path=str(answer_pdf),
                homework_docx_path=str(hw_path),
                homework_type=type_name,
                author="时月学姐",
            )
            all_results.append(result)

    print(f"\n{'='*60}")
    print(f"✅ 一稿批改完成！共处理 {len(all_results)} 份作业")
    return all_results


def auto_grade_second_drafts():
    """自动扫描所有作业类型的二稿，对比一稿评语批改二稿"""
    draft_dirs = discover_draft_dirs(PACKAGE_DIR)
    if not draft_dirs:
        print("📭 未发现二稿批改目录（需要 作业类型/二稿批改/）")
        return []

    all_results = []
    for type_name, entry in draft_dirs.items():
        second = entry.get("second_draft")
        if not second:
            continue

        first_draft_dir = second["first_draft_dir"]
        second_draft_dir = second["second_draft_dir"]
        output_dir = second["output_dir"]

        ungraded = list_ungraded_second_drafts(second_draft_dir, output_dir)
        if not ungraded:
            print(f"\n📂 [{type_name}·二稿] 所有二稿已批改，无需处理")
            continue

        print(f"\n{'='*60}")
        print(f"📂 [{type_name}·二稿] 发现 {len(ungraded)} 份未批改二稿")
        print(f"{'='*60}")

        for second_path in ungraded:
            print(f"\n{'='*60}")
            print(f"📄 二稿: {second_path.name}")

            # 匹配一稿批改文件（从 二稿批改/一稿批改/ 目录）
            first_draft = match_first_draft(second_path, first_draft_dir)
            if not first_draft:
                print(f"   ⚠️  跳过：找不到对应的一稿批改文件（{first_draft_dir}）")
                continue
            print(f"   🔗 一稿批改: {first_draft.name}")

            result = run_second_draft_grader(
                first_draft_docx_path=str(first_draft),
                second_draft_docx_path=str(second_path),
                homework_type=type_name,
                author="时月学姐",
            )
            all_results.append(result)

    print(f"\n{'='*60}")
    print(f"✅ 二稿批改完成！共处理 {len(all_results)} 份二稿")
    return all_results


def _cleanup_temp_dir():
    """清空 word_agent/temp/ 下的所有内容（含子目录），批改结束后调用"""
    temp_dir = PACKAGE_DIR / "temp"
    if not temp_dir.exists():
        return
    import shutil
    import gc
    import time
    deleted = 0
    failed = 0
    for item in temp_dir.iterdir():
        removed = False
        # 多次重试以应对 Windows 文件锁
        for attempt in range(5):
            try:
                # 先清除只读属性（copy2 复制的文件可能带只读属性，导致无法删除）
                if item.is_dir():
                    for root, dirs, files in os.walk(item):
                        for f in files:
                            try:
                                os.chmod(os.path.join(root, f), 0o666)
                            except Exception:
                                pass
                    shutil.rmtree(item)
                else:
                    try:
                        os.chmod(item, 0o666)
                    except Exception:
                        pass
                    item.unlink()
                removed = True
                break
            except Exception:
                if attempt < 4:
                    time.sleep(0.2)
                    gc.collect()
        if removed:
            deleted += 1
        else:
            failed += 1
            print(f"   ⚠️  清理 temp 失败: {item.name}")
    if deleted or failed:
        print(f"\n🧹 temp/ 清理完成：删除 {deleted} 项" + (f"，{failed} 项因文件占用暂未删除" if failed else ""))


def main():
    """主入口 — 运行即自动扫描批改（一稿 + 二稿），结束清空 temp"""
    print("=" * 60)
    print("  Word Agent 智能批改系统")
    print("=" * 60)
    print("  ① 自动扫描 一稿批改/homework/ 目录，增量批改一稿...")
    auto_grade()
    print()
    print("  ② 自动扫描 二稿批改/ 目录，对比一稿批改二稿...")
    auto_grade_second_drafts()
    print()
    print("  ③ 清空 temp/ 临时文件夹...")
    _cleanup_temp_dir()


if __name__ == "__main__":
    main()