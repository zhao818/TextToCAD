#!/usr/bin/env python3
# t2cad_eval.py V2 — TextToCAD 评测引擎（项目级）
"""
Eval-driven development for TextToCAD tools — V2 Project-Level.

Directory (V2):
  ~/.text_to_cad/
  ├── t2cad_eval.py              # 引擎（通用）
  ├── config/
  │   └── t2cad.yaml             # 项目配置（可选）
  └── eval/
      ├── datasets/
      │   └── <project>/         # 如 t2cad, sql-gen
      │       ├── gold_set/      # 黄金集（必须全过）
      │       ├── edge_cases/    # 边界样本（测鲁棒性）
      │       └── regression/    # 回归测试（以前犯的错）
      └── results/               # 历史运行结果

Test case format (unchanged):
  { name, description, snapshot, user_input, checks, difficulty, tags }

Usage:
  python t2cad_eval.py run --project t2cad --tool excel
  python t2cad_eval.py run --project t2cad --tool excel --suite regression
  python t2cad_eval.py run --project t2cad --tool excel --suite all --dry
  python t2cad_eval.py compare --project t2cad --tool excel
  python t2cad_eval.py report --project t2cad --last 7days --format markdown
"""

import ast
import hashlib
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Optional

# ── Paths ──────────────────────────────────────────────
CONFIG_DIR = Path.home() / ".text_to_cad"
EVAL_DIR = CONFIG_DIR / "eval"
DATASETS_DIR = EVAL_DIR / "datasets"
RESULTS_DIR = EVAL_DIR / "results"

for d in [EVAL_DIR, DATASETS_DIR, RESULTS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

sys.path.insert(0, str(CONFIG_DIR))

# ── Constants ──────────────────────────────────────────
WEIGHTS = {
    "syntax": 0.20,
    "safety": 0.30,
    "task_fit": 0.30,
    "quality": 0.20,
}

PASS_THRESHOLD = 0.90
SUITES = ["gold_set", "edge_cases", "regression"]

# ── Safety: banned patterns ────────────────────────────
BANNED_PATTERNS = [
    (r'\bimport\s+', "禁止 import"),
    (r'\bfrom\s+\w+\s+import\b', "禁止 from-import"),
    (r'\bDispatch\b', "禁止 Dispatch（应使用已注入对象）"),
    (r'\bEnsureDispatch\b', "禁止 EnsureDispatch"),
    (r'\bopen\s*\(', "禁止 open() 文件操作"),
    (r'\bprint\s*\(', "禁止 print()"),
    (r'\bexit\s*\(', "禁止 exit()"),
    (r'\bsys\.exit\b', "禁止 sys.exit()"),
    (r'\bos\.system\b', "禁止 os.system()"),
    (r'\bsubprocess\b', "禁止 subprocess"),
    (r'\beval\s*\(', "禁止 eval()"),
    (r'\bexec\s*\(', "禁止 exec()（代码本身在 exec 中运行）"),
    (r'\bWorkbooks\.Open\b', "禁止 Workbooks.Open"),
    (r'\bPresentations\.Open\b', "禁止 Presentations.Open"),
    (r'\bDocuments\.Open\b', "禁止 Documents.Open"),
]

QUALITY_ANTI_PATTERNS = [
    (r'except\s*:', "裸 except: 应指定具体异常类型"),
    (r'except\s+Exception\s*:', "裸 catch Exception 太宽泛"),
    (r'\bwhile\s+True\b', "while True 可能死循环"),
    (r'[a-z]\s*=\s*[a-z]\s*=\s*[a-z]', "一行多个赋值影响可读性"),
]

QUALITY_POSITIVE = [
    (r'#', "包含注释"),
    (r'def\s+\w+', "使用函数封装"),
]


# ════════════════════════════════════════════════════════
# V2: Dataset management (project × suite × tool)
# ════════════════════════════════════════════════════════

def _dataset_path(project: str, suite: str, tool: str) -> Path:
    """Get path to a dataset file."""
    if suite == "all":
        return None  # special: load from all suites
    return DATASETS_DIR / project / suite / f"{tool}_cases.json"


def load_cases(project: str, tool: str, suite: str = "all") -> list[dict]:
    """Load test cases from project/dataset hierarchy.

    suite='all' merges gold_set + edge_cases + regression.
    """
    all_cases = []

    suites_to_load = SUITES if suite == "all" else [suite]

    for s in suites_to_load:
        path = DATASETS_DIR / project / s / f"{tool}_cases.json"
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            cases = data.get("cases", [])
            # Tag each case with its source suite
            for c in cases:
                c["_suite"] = s
            all_cases.extend(cases)

    return all_cases


def save_cases(project: str, tool: str, cases: list[dict], suite: str = "gold_set"):
    """Save test cases to project/dataset hierarchy."""
    suite_dir = DATASETS_DIR / project / suite
    suite_dir.mkdir(parents=True, exist_ok=True)
    case_file = suite_dir / f"{tool}_cases.json"

    # Strip internal _suite tags
    clean_cases = []
    for c in cases:
        clean = dict(c)
        clean.pop("_suite", None)
        clean_cases.append(clean)

    data = {
        "project": project,
        "tool": tool,
        "suite": suite,
        "version": 2,
        "updated": datetime.now().isoformat(),
        "count": len(clean_cases),
        "cases": clean_cases,
    }
    with open(case_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def migrate_v1_to_v2():
    """Migrate old eval/cases/<tool>_cases.json → new datasets/t2cad/gold_set/"""
    old_cases_dir = EVAL_DIR / "cases"
    if not old_cases_dir.exists():
        return

    for old_file in old_cases_dir.glob("*_cases.json"):
        tool = old_file.stem.replace("_cases", "")
        with open(old_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        cases = data.get("cases", [])

        # Move to new structure
        new_dir = DATASETS_DIR / "t2cad" / "gold_set"
        new_dir.mkdir(parents=True, exist_ok=True)
        new_file = new_dir / f"{tool}_cases.json"

        if not new_file.exists():
            save_cases("t2cad", tool, cases, "gold_set")
            print(f"  ✅ 已迁移: {tool} ({len(cases)} 用例) → datasets/t2cad/gold_set/")

    # Rename old dir to .bak
    old_cases_dir.rename(EVAL_DIR / "cases.v1.bak")
    print(f"  📦 旧目录已备份: cases.v1.bak")


# ════════════════════════════════════════════════════════
# Scoring Functions
# ════════════════════════════════════════════════════════

def check_syntax(code: str) -> tuple[float, str]:
    try:
        ast.parse(code)
        return 1.0, "语法正确"
    except SyntaxError as e:
        return 0.0, f"语法错误: {e}"


def check_safety(code: str) -> tuple[float, str]:
    violations = []
    for pattern, desc in BANNED_PATTERNS:
        if re.search(pattern, code):
            violations.append(f"  ❌ {desc}")
    if not violations:
        return 1.0, "安全: 无违规"
    score = max(0.0, 1.0 - len(violations) * 0.15)
    return score, "安全违规:\n" + "\n".join(violations)


def check_task_fit(code: str, case: dict) -> tuple[float, str]:
    checks = case.get("checks", {})
    score = 1.0
    details = []

    for func in checks.get("must_use", []):
        if func not in code:
            score -= 0.15
            details.append(f"  ❌ 缺少必要函数: {func}")
        else:
            details.append(f"  ✅ 使用: {func}")

    for pattern in checks.get("must_not_use", []):
        if pattern in code:
            score -= 0.2
            details.append(f"  ❌ 不应使用: {pattern}")

    expected = checks.get("expected_operations", [])
    found_any = False
    for op in expected:
        if op.lower() in code.lower():
            found_any = True
            details.append(f"  ✅ 包含操作: {op}")
            break
    if expected and not found_any:
        score -= 0.2
        details.append(f"  ⚠️ 未找到期望操作: {expected}")

    for item in checks.get("should_contain", []):
        if item in code:
            details.append(f"  ✅ 含推荐项: {item}")

    return max(0.0, min(1.0, score)), "任务匹配:\n" + "\n".join(details) if details else "任务匹配: 无检查项"


def check_quality(code: str) -> tuple[float, str]:
    score = 0.5
    details = []

    for pattern, desc in QUALITY_ANTI_PATTERNS:
        if re.search(pattern, code):
            score -= 0.1
            details.append(f"  ⚠️ {desc}")

    for pattern, desc in QUALITY_POSITIVE:
        if re.search(pattern, code):
            score += 0.1
            details.append(f"  ✅ {desc}")

    lines = code.strip().split("\n")
    if len(lines) < 3:
        score -= 0.2
        details.append("  ⚠️ 代码过短（可能不完整）")
    elif len(lines) > 200:
        score -= 0.1
        details.append("  ⚠️ 代码过长（>200行）")

    var_pattern = re.findall(r'^(\w+)\s*=', code, re.MULTILINE)
    single_char_vars = [v for v in var_pattern if len(v) == 1 and v not in ('x', 'y', 'z', 'i', 'j', 'k', 'n')]
    if len(single_char_vars) > 3:
        score -= 0.1
        details.append(f"  ⚠️ 过多单字母变量: {single_char_vars}")

    return max(0.0, min(1.0, score)), "代码质量:\n" + "\n".join(details) if details else "代码质量: 无特别发现"


# ════════════════════════════════════════════════════════
# Evaluation Engine
# ════════════════════════════════════════════════════════

class EvalEngine:
    def __init__(self, dry_run: bool = False):
        self.dry_run = dry_run
        self.client = None
        self.pipeline = None
        if not dry_run:
            from t2cad_llm import get_client
            from t2cad_pipeline import CodeGenPipeline
            self.client = get_client()
            self.pipeline = CodeGenPipeline(self.client)

    def evaluate_one(self, case: dict, system_prompt: str,
                     fixer_prompt: str = "", exec_ns: dict = None) -> dict:
        start_time = time.time()

        if self.dry_run:
            code = case.get("_dry_code", "")
            pipeline_result = {"success": bool(code), "code": code,
                              "result": "dry run", "attempt": 0, "all_codes": [code]}
        else:
            pipeline_result = self.pipeline.run(
                system_prompt=system_prompt,
                snapshot=case.get("snapshot", ""),
                user_input=case.get("user_input", ""),
                exec_namespace=exec_ns or {},
                fixer_prompt=fixer_prompt,
                max_retries=3,
            )

        code = pipeline_result.get("code", "")
        elapsed = time.time() - start_time

        syntax_score, syntax_detail = check_syntax(code)
        safety_score, safety_detail = check_safety(code)
        task_score, task_detail = check_task_fit(code, case)
        quality_score, quality_detail = check_quality(code)

        overall = (
            syntax_score * WEIGHTS["syntax"] +
            safety_score * WEIGHTS["safety"] +
            task_score * WEIGHTS["task_fit"] +
            quality_score * WEIGHTS["quality"]
        )

        return {
            "case_name": case["name"],
            "suite": case.get("_suite", "unknown"),
            "difficulty": case.get("difficulty", "unknown"),
            "tags": case.get("tags", []),
            "code": code,
            "code_hash": hashlib.sha256(code.encode()).hexdigest()[:12],
            "scores": {
                "syntax": round(syntax_score, 3),
                "safety": round(safety_score, 3),
                "task_fit": round(task_score, 3),
                "quality": round(quality_score, 3),
                "overall": round(overall, 3),
            },
            "details": {
                "syntax": syntax_detail,
                "safety": safety_detail,
                "task_fit": task_detail,
                "quality": quality_detail,
            },
            "pipeline": {
                "success": pipeline_result.get("success", False),
                "attempts": pipeline_result.get("attempt", 0),
                "result": pipeline_result.get("result", ""),
            },
            "passed": overall >= PASS_THRESHOLD,
            "elapsed_sec": round(elapsed, 1),
        }

    def run_suite(self, project: str, tool: str, system_prompt: str,
                  fixer_prompt: str = "", exec_ns: dict = None,
                  suite: str = "all", tag_filter: str = None,
                  limit: int = None) -> dict:
        cases = load_cases(project, tool, suite)
        if not cases:
            return {"error": f"未找到 {project}/{tool} 的测试用例，请先 init"}

        if tag_filter:
            cases = [c for c in cases if tag_filter in c.get("tags", [])]
        if limit:
            cases = cases[:limit]

        results = []
        passed = 0
        failed = 0

        for i, case in enumerate(cases):
            print(f"\n{'─'*60}")
            src = case.get("_suite", "?")
            print(f"[{i+1}/{len(cases)}] {case['name']} ({case.get('difficulty', '?')} / {src})")
            print(f"      输入: {case['user_input'][:80]}...")

            if not self.dry_run:
                print(f"      ⏳ 生成代码...")

            result = self.evaluate_one(case, system_prompt, fixer_prompt, exec_ns)
            results.append(result)

            status = "✅ PASS" if result["passed"] else "❌ FAIL"
            print(f"      {status} 综合: {result['scores']['overall']:.1%}  "
                  f"语{result['scores']['syntax']:.0%} 安{result['scores']['safety']:.0%}  "
                  f"任{result['scores']['task_fit']:.0%} 质{result['scores']['quality']:.0%}  "
                  f"{result['elapsed_sec']}s")

            if not result["passed"]:
                failed += 1
                for dim in ["syntax", "safety", "task_fit", "quality"]:
                    if result["scores"][dim] < 0.9:
                        print(f"      [{dim}] {result['details'][dim]}")
            else:
                passed += 1

        avg = sum(r["scores"]["overall"] for r in results) / len(results) if results else 0
        summary = {
            "project": project,
            "tool": tool,
            "suite": suite,
            "total": len(results),
            "passed": passed,
            "failed": failed,
            "pass_rate": round(passed / len(results), 3) if results else 0,
            "avg_overall": round(avg, 3),
            "avg_syntax": round(sum(r["scores"]["syntax"] for r in results) / len(results), 3) if results else 0,
            "avg_safety": round(sum(r["scores"]["safety"] for r in results) / len(results), 3) if results else 0,
            "avg_task_fit": round(sum(r["scores"]["task_fit"] for r in results) / len(results), 3) if results else 0,
            "avg_quality": round(sum(r["scores"]["quality"] for r in results) / len(results), 3) if results else 0,
            "avg_attempts": round(sum(r["pipeline"]["attempts"] for r in results) / len(results), 1) if results else 0,
            "avg_elapsed": round(sum(r["elapsed_sec"] for r in results) / len(results), 1) if results else 0,
            "tag_filter": tag_filter,
            "dry_run": self.dry_run,
        }

        # Per-suite breakdown
        suite_stats = {}
        for r in results:
            s = r.get("suite", "?")
            if s not in suite_stats:
                suite_stats[s] = {"total": 0, "passed": 0}
            suite_stats[s]["total"] += 1
            if r["passed"]:
                suite_stats[s]["passed"] += 1
        summary["by_suite"] = {
            s: {"pass_rate": round(v["passed"] / v["total"], 3), "total": v["total"]}
            for s, v in suite_stats.items()
        }

        return {"summary": summary, "results": results}


# ════════════════════════════════════════════════════════
# Results Management
# ════════════════════════════════════════════════════════

def save_result(project: str, tool: str, result: dict, label: str = ""):
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{project}_{tool}_{ts}.json"
    if label:
        filename = f"{project}_{tool}_{ts}_{label}.json"

    result["meta"] = {
        "project": project,
        "tool": tool,
        "timestamp": datetime.now().isoformat(),
        "label": label,
        "eval_version": 2,
    }

    filepath = RESULTS_DIR / filename
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    return filepath


def _parse_project_tool_from_filename(filename: str) -> tuple:
    """Infer project and tool from filename.
    V2 format: {project}_{tool}_{YYYYMMDD}_{HHMMSS}[_{label}].json
    V1 format: {tool}_{YYYYMMDD}_{HHMMSS}.json
    """
    import re as _re
    stem = filename.replace(".json", "")
    parts = stem.split("_")

    # Find the date part (8 digits)
    date_idx = None
    for i, p in enumerate(parts):
        if _re.match(r'^\d{8}$', p):
            date_idx = i
            break

    if date_idx is None:
        return "unknown", stem

    if date_idx >= 2:
        # V2: project_tool_YYYYMMDD_...
        return parts[0], parts[1]
    elif date_idx == 1:
        # V1: tool_YYYYMMDD_...
        return "t2cad", parts[0]
    return "unknown", parts[0]


def list_results(project: str = None, tool: str = None, days: int = None) -> list[dict]:
    """List saved eval results, optionally filtered. Handles both V1 and V2 formats."""
    all_files = sorted(RESULTS_DIR.glob("*.json"), reverse=True)

    # Filter by project/tool from filename and meta
    results = []
    cutoff = datetime.now() - timedelta(days=days) if days else None

    for f in all_files:
        with open(f, "r", encoding="utf-8") as fp:
            data = json.load(fp)
        meta = data.get("meta", {})
        ts = meta.get("timestamp", "")

        # Determine project/tool: meta first, then filename inference
        p = meta.get("project") or _parse_project_tool_from_filename(f.name)[0]
        t = meta.get("tool") or _parse_project_tool_from_filename(f.name)[1]

        if project and p != project:
            continue
        if tool and t != tool:
            continue

        if cutoff and ts:
            try:
                if datetime.fromisoformat(ts) < cutoff:
                    continue
            except ValueError:
                pass

        summary = data.get("summary", {})
        results.append({
            "file": f.name,
            "project": p,
            "tool": t,
            "timestamp": ts,
            "label": meta.get("label", ""),
            "pass_rate": summary.get("pass_rate", 0),
            "avg_overall": summary.get("avg_overall", 0),
            "total": summary.get("total", 0),
            "passed": summary.get("passed", 0),
            "failed": summary.get("failed", 0),
            "by_suite": summary.get("by_suite", {}),
            "full_data": data,
        })
    return results


def compare_results(project: str, tool: str) -> dict:
    """Compare the two most recent eval results."""
    results = list_results(project=project, tool=tool)
    if len(results) < 2:
        return {"error": f"需要至少2次评测结果才能对比，当前只有 {len(results)} 次"}

    new = results[0]["full_data"]
    old = results[1]["full_data"]
    new_sum = new["summary"]
    old_sum = old["summary"]

    diff = {
        "project": project,
        "tool": tool,
        "new": {"timestamp": new["meta"]["timestamp"], "label": new["meta"].get("label", "")},
        "old": {"timestamp": old["meta"]["timestamp"], "label": old["meta"].get("label", "")},
        "delta": {
            "pass_rate": round(new_sum["pass_rate"] - old_sum["pass_rate"], 3),
            "avg_overall": round(new_sum["avg_overall"] - old_sum["avg_overall"], 3),
            "avg_syntax": round(new_sum["avg_syntax"] - old_sum["avg_syntax"], 3),
            "avg_safety": round(new_sum["avg_safety"] - old_sum["avg_safety"], 3),
            "avg_task_fit": round(new_sum["avg_task_fit"] - old_sum["avg_task_fit"], 3),
            "avg_quality": round(new_sum["avg_quality"] - old_sum["avg_quality"], 3),
            "avg_attempts": round(new_sum["avg_attempts"] - old_sum["avg_attempts"], 1),
        },
        "per_case_delta": [],
        "regression_alert": [],
    }

    new_cases = {r["case_name"]: r for r in new["results"]}
    old_cases = {r["case_name"]: r for r in old["results"]}
    for name in new_cases:
        if name in old_cases:
            delta = round(new_cases[name]["scores"]["overall"] - old_cases[name]["scores"]["overall"], 3)
            item = {
                "case": name,
                "old": old_cases[name]["scores"]["overall"],
                "new": new_cases[name]["scores"]["overall"],
                "delta": delta,
                "direction": "📈" if delta > 0 else ("📉" if delta < 0 else "➡️"),
                "was_pass": old_cases[name]["passed"],
                "now_pass": new_cases[name]["passed"],
            }
            diff["per_case_delta"].append(item)
            # Regression alert: was passing, now failing
            if old_cases[name]["passed"] and not new_cases[name]["passed"]:
                diff["regression_alert"].append(item)

    return diff


# ════════════════════════════════════════════════════════
# V2: Weekly/Monthly Report Generator
# ════════════════════════════════════════════════════════

def generate_report(project: str, days: int = 7, format: str = "markdown") -> str:
    """Generate a comprehensive eval report for a time period.

    Args:
        project: Project name (e.g. 't2cad')
        days: Number of days to look back
        format: 'markdown' or 'text'

    Returns report as string.
    """
    results = list_results(project=project, days=days)
    if not results:
        return f"# {project} 评测报告\n\n暂无最近 {days} 天的评测数据。"

    now = datetime.now()
    start_date = now - timedelta(days=days)
    date_range = f"{start_date.strftime('%Y.%m.%d')} - {now.strftime('%Y.%m.%d')}"

    # Group by tool
    by_tool = {}
    for r in results:
        tool = r["tool"]
        if tool not in by_tool:
            by_tool[tool] = []
        by_tool[tool].append(r)

    lines = []
    lines.append(f"# FDE 项目评测报告 ({date_range})")
    lines.append("")
    lines.append(f"> 自动生成于 {now.strftime('%Y-%m-%d %H:%M')} · 数据来源: {len(results)} 次评测")
    lines.append("")

    # ── 1. Overall Health ──
    lines.append("## 1. 整体健康度")
    lines.append("")
    lines.append("| 工具 | 最新通过率 | 最新均分 | 变化 (vs上周) | 状态 |")
    lines.append("|------|:--:|:--:|:--:|:--:|")

    for tool, tool_results in sorted(by_tool.items()):
        latest = tool_results[0]
        rate = latest["pass_rate"]
        score = latest["avg_overall"]

        # Compare to previous run
        delta_str = "—"
        if len(tool_results) >= 2:
            prev = tool_results[1]
            delta = rate - prev["pass_rate"]
            arrow = "↑" if delta > 0 else ("↓" if delta < 0 else "→")
            delta_str = f"{arrow} {delta:+.1%}"

        status = "🟢" if rate >= 0.90 else ("🟡" if rate >= 0.70 else "🔴")
        lines.append(f"| **{tool}** | {rate:.1%} | {score:.1%} | {delta_str} | {status} |")

    lines.append("")

    # ── 2. Key Progress ──
    lines.append("## 2. 关键进展")
    lines.append("")

    progress_found = False
    for tool, tool_results in sorted(by_tool.items()):
        if len(tool_results) >= 2:
            newer = tool_results[0]["full_data"]
            older = tool_results[1]["full_data"]

            # Find cases that improved significantly
            new_cases = {r["case_name"]: r for r in newer["results"]}
            old_cases = {r["case_name"]: r for r in older["results"]}

            improved = []
            for name in new_cases:
                if name in old_cases:
                    delta = new_cases[name]["scores"]["overall"] - old_cases[name]["scores"]["overall"]
                    if delta > 0.05:
                        improved.append((name, delta, new_cases[name]["scores"]["overall"]))

            if improved:
                progress_found = True
                lines.append(f"### {tool}")
                for name, delta, final in sorted(improved, key=lambda x: -x[1])[:5]:
                    lines.append(f"- **{name}** ↑{delta:+.0%} → {final:.0%}")
                lines.append("")

    if not progress_found:
        lines.append("*本周无显著提升的用例。*")
        lines.append("")

    # ── 3. Regression Alerts ──
    lines.append("## 3. 回归警报")
    lines.append("")

    regression_found = False
    for tool, tool_results in sorted(by_tool.items()):
        if len(tool_results) >= 2:
            diff = compare_results(project, tool)
            if "regression_alert" in diff and diff["regression_alert"]:
                regression_found = True
                lines.append(f"### ⚠️ {tool}")
                for item in diff["regression_alert"]:
                    lines.append(f"- **{item['case']}** 📉 {item['old']:.1%} → {item['new']:.1%} ({item['delta']:+.1%}) —— 之前通过，现在失败！")
                lines.append("")

    if not regression_found:
        lines.append("*✅ 无回归问题。*")
        lines.append("")

    # ── 4. Suite Breakdown ──
    lines.append("## 4. 测试套件明细")
    lines.append("")

    for tool, tool_results in sorted(by_tool.items()):
        latest = tool_results[0]
        by_suite = latest.get("by_suite", {})
        if by_suite:
            lines.append(f"### {tool}")
            lines.append("")
            lines.append("| 套件 | 通过率 | 用例数 |")
            lines.append("|------|:--:|:--:|")
            for suite, stats in sorted(by_suite.items()):
                lines.append(f"| {suite} | {stats['pass_rate']:.1%} | {stats['total']} |")
            lines.append("")

    # ── 5. TODO ──
    lines.append("## 5. 待办事项")
    lines.append("")

    # Find all failed cases in the latest run
    todo_items = []
    for tool, tool_results in sorted(by_tool.items()):
        latest = tool_results[0]["full_data"]
        for r in latest["results"]:
            if not r["passed"]:
                weakness = []
                for dim in ["syntax", "safety", "task_fit", "quality"]:
                    if r["scores"][dim] < 0.9:
                        weakness.append(dim)
                todo_items.append(f"- [{tool}] **{r['case_name']}** ({r.get('suite', '?')}) — 弱项: {', '.join(weakness)} ({r['scores']['overall']:.0%})")

    if todo_items:
        for item in todo_items:
            lines.append(item)
    else:
        lines.append("*🎉 所有用例全部通过！*")

    lines.append("")
    lines.append("---")
    lines.append(f"*报告由 t2cad_eval.py V2 自动生成 · {now.strftime('%Y-%m-%d %H:%M:%S')}*")

    return "\n".join(lines)


# ════════════════════════════════════════════════════════
# Tool-specific helpers
# ════════════════════════════════════════════════════════

def get_system_prompt(tool_name: str) -> str:
    cfg_file = CONFIG_DIR / "config.json"
    if cfg_file.exists():
        with open(cfg_file, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        if "system_prompt_zh" in cfg:
            return cfg["system_prompt_zh"]

    fallbacks = {
        "excel": "你是 Excel 专家。操作 excel/wb/ws。用 V() FIX() BEAUTIFY() 等安全函数。只输出 Python 代码。",
        "cad": "你是 AutoCAD 专家。操作 acad 对象。用 L() C() PLINE() 等安全函数。只输出 Python 代码。",
        "word": "你是 Word 专家。操作 word/doc/sel。用 TB() MERGE() 等安全函数。只输出 Python 代码。",
        "ppt": "你是 PPT 专家。操作 ppt/pres/slide。用 TB_STYLE() SHAPE() 等安全函数。只输出 Python 代码。",
    }
    return fallbacks.get(tool_name, "你是技术专家。只输出 Python 代码。")


def get_fixer_prompt(tool_name: str) -> str:
    fallbacks = {
        "excel": "你是 Excel COM 专家。修正代码。只输出代码。",
        "cad": "你是 AutoCAD COM 专家。修正代码。只输出代码。",
        "word": "你是 Word COM 专家。修正代码。只输出代码。",
        "ppt": "你是 PPT COM 专家。修正代码。只输出代码。",
    }
    return fallbacks.get(tool_name, "你是调试专家。修正代码。只输出代码。")


def get_exec_namespace(tool_name: str) -> dict:
    base = {"math": __import__("math"), "re": __import__("re")}

    common = {
        "V": lambda cell=None: None,
        "FIX": lambda cell=None, formula=None: None,
        "ERR": lambda cell=None: None,
        "LOOP": lambda ws=None, col=None, start=None, end=None: [],
        "BEAUTIFY": lambda rng=None: None,
        "MERGE": lambda rng=None: None,
        "BORDER": lambda rng=None: None,
        "BG": lambda rng=None, color=None: None,
        "FONT_SET": lambda rng=None, name=None, size=None, bold=None: None,
        "ALIGN": lambda rng=None, align=None: None,
        "NUMFMT": lambda rng=None, fmt=None: None,
        "FREEZE": lambda r=None, c=None: None,
        "AUTOFIT": lambda: None,
        "COND_FMT": lambda rng=None, rule=None: None,
        "CHART": lambda type=None, rng=None, x=None, y=None, w=None, h=None: None,
        "NEW_SHEET": lambda name: type('Sheet', (), {'Name': name, 'Range': lambda *a: type('Range', (), {'Value': None})()}),
        "GET_SHEET": lambda name: type('Sheet', (), {'Name': name})(),
    }

    cad_funcs = {
        "FAST_MODE": lambda b: None, "L": lambda *a: None, "C": lambda *a: None,
        "R": lambda *a: None, "T": lambda *a: None, "DIM": lambda *a: None,
        "PLINE": lambda *a: None, "ARC": lambda *a: None, "HATCH": lambda *a: None,
        "MTEXT": lambda *a: None, "MOVE": lambda *a: None, "ROT": lambda *a: None,
        "DEL": lambda *a: None, "OBJS": lambda: [], "FIND_OBJS": lambda **kw: [],
        "DEL_BY_TYPE": lambda t: None, "UNDO": lambda: None, "LAYER": lambda *a: None,
        "SEND_CMD": lambda cmd: None, "ZOOM_EXT": lambda: None, "ZOOM_WIN": lambda *a: None,
        "BLOCK_INSERT": lambda *a: None, "SET_COLOR": lambda *a: None,
        "DIM_H": lambda *a: None, "DIM_STYLE": lambda: None,
    }

    ns = dict(base)
    ns.update(common)
    if tool_name == "cad": ns.update(cad_funcs)
    return ns


# ════════════════════════════════════════════════════════
# CLI (V2)
# ════════════════════════════════════════════════════════

def print_summary(result: dict):
    s = result["summary"]
    print(f"\n{'='*60}")
    print(f"📊 评测总结: {s.get('project', '?')}/{s['tool']}")
    print(f"{'='*60}")
    print(f"  通过率:    {s['pass_rate']:.1%} ({s['passed']}/{s['total']})")
    print(f"  综合均分:  {s['avg_overall']:.1%}")
    print(f"  语法: {s['avg_syntax']:.1%}  安全: {s['avg_safety']:.1%}  任务: {s['avg_task_fit']:.1%}  质量: {s['avg_quality']:.1%}")
    print(f"  平均尝试: {s['avg_attempts']}次  平均耗时: {s['avg_elapsed']}s")
    if s.get("dry_run"): print(f"  ⚠️  干跑模式")
    if "by_suite" in s:
        print(f"\n  按套件:")
        for suite, stats in s["by_suite"].items():
            print(f"    {suite}: {stats['pass_rate']:.1%} ({stats['total']}用例)")

    grade = "🏆" if s['pass_rate'] >= 0.9 else ("✅" if s['pass_rate'] >= 0.7 else "⚠️")
    verdict = "PASS" if s['pass_rate'] >= PASS_THRESHOLD else "FAIL"
    print(f"\n  {grade} 最终判定: {verdict} (阈值 {PASS_THRESHOLD:.0%})")

    print(f"\n{'─'*60}")
    print(f"📋 逐项明细:")
    for r in result["results"]:
        icon = "✅" if r["passed"] else "❌"
        print(f"  {icon} {r['case_name']:28s} {r['scores']['overall']:.1%}  "
              f"{r.get('suite', '?')}")


def print_compare(diff: dict):
    d = diff["delta"]
    print(f"\n{'='*60}")
    print(f"📊 评测对比: {diff['project']}/{diff['tool']}")
    print(f"{'='*60}")
    print(f"  旧: {diff['old']['timestamp']} ({diff['old']['label']})")
    print(f"  新: {diff['new']['timestamp']} ({diff['new']['label']})")
    print(f"\n  维度变化:")
    for key, label in [("pass_rate", "通过率"), ("avg_overall", "综合"),
                        ("avg_syntax", "语法"), ("avg_safety", "安全"),
                        ("avg_task_fit", "任务匹配"), ("avg_quality", "质量")]:
        delta = d[key]
        arrow = "📈" if delta > 0 else ("📉" if delta < 0 else "➡️")
        print(f"  {arrow} {label}: {delta:+.1%}")

    if diff.get("regression_alert"):
        print(f"\n  🚨 回归警报:")
        for item in diff["regression_alert"]:
            print(f"  ❌ {item['case']}: {item['old']:.1%} → {item['new']:.1%} —— 之前PASS，现在FAIL!")

    print(f"\n  逐用例:")
    for item in diff["per_case_delta"]:
        print(f"  {item['direction']} {item['case']:28s} {item['old']:.1%} → {item['new']:.1%} ({item['delta']:+.1%})")


def cmd_run(args: list):
    project = "t2cad"
    tool = None
    suite = "all"
    tag = None
    limit = None
    dry = False
    label = ""

    i = 1
    while i < len(args):
        if args[i] == "--project" and i + 1 < len(args):
            project = args[i + 1]; i += 2
        elif args[i] == "--tool" and i + 1 < len(args):
            tool = args[i + 1]; i += 2
        elif args[i] == "--suite" and i + 1 < len(args):
            suite = args[i + 1]; i += 2
        elif args[i] == "--tag" and i + 1 < len(args):
            tag = args[i + 1]; i += 2
        elif args[i] == "--limit" and i + 1 < len(args):
            limit = int(args[i + 1]); i += 2
        elif args[i] == "--dry":
            dry = True; i += 1
        elif args[i] == "--label" and i + 1 < len(args):
            label = args[i + 1]; i += 2
        elif not tool:
            tool = args[i]; i += 1
        else:
            i += 1

    if not tool:
        print("用法: t2cad_eval.py run --project <项目> --tool <工具> [--suite <套件>] [--tag <标签>] [--dry]")
        print("套件: gold_set, edge_cases, regression, all (默认)")
        return

    sys_prompt = get_system_prompt(tool)
    fixer = get_fixer_prompt(tool)
    ns = get_exec_namespace(tool)

    print(f"🚀 评测: {project}/{tool}  suite={suite}")
    print(f"   提示词: {sys_prompt[:80]}...")
    if dry: print(f"   ⚠️  干跑模式")

    engine = EvalEngine(dry_run=dry)
    result = engine.run_suite(project, tool, sys_prompt, fixer, ns, suite=suite, tag_filter=tag, limit=limit)

    if "error" in result:
        print(f"❌ {result['error']}")
        return

    print_summary(result)
    filepath = save_result(project, tool, result, label)
    print(f"\n💾 结果已保存: {filepath}")


def cmd_compare(args: list):
    project = "t2cad"
    tool = None

    i = 1
    while i < len(args):
        if args[i] == "--project" and i + 1 < len(args):
            project = args[i + 1]; i += 2
        elif not tool:
            tool = args[i]; i += 1
        else:
            i += 1

    if not tool:
        results = list_results(project=project)
        if results:
            tool = results[0]["tool"]
        else:
            print("未找到评测结果。先 run")
            return

    diff = compare_results(project, tool)
    if "error" in diff:
        print(f"❌ {diff['error']}")
        return
    print_compare(diff)


def cmd_report(args: list):
    project = "t2cad"
    days = 7
    fmt = "markdown"
    output = None

    i = 1
    while i < len(args):
        if args[i] == "--project" and i + 1 < len(args):
            project = args[i + 1]; i += 2
        elif args[i] == "--last" and i + 1 < len(args):
            val = args[i + 1]
            if val.endswith("days"):
                days = int(val.replace("days", ""))
            elif val.endswith("d"):
                days = int(val.replace("d", ""))
            else:
                days = int(val)
            i += 2
        elif args[i] == "--format" and i + 1 < len(args):
            fmt = args[i + 1]; i += 2
        elif args[i] == "--output" and i + 1 < len(args):
            output = args[i + 1]; i += 2
        else:
            i += 1

    report = generate_report(project, days, fmt)

    if output:
        out_path = Path(output)
        out_path.write_text(report, encoding="utf-8")
        print(f"📄 报告已保存: {out_path}")
    else:
        print(report)


def cmd_init(args: list):
    """Initialize test cases. V2: project + suite structure."""
    project = "t2cad"
    tool = None
    suite = "gold_set"

    i = 1
    while i < len(args):
        if args[i] == "--project" and i + 1 < len(args):
            project = args[i + 1]; i += 2
        elif args[i] == "--suite" and i + 1 < len(args):
            suite = args[i + 1]; i += 2
        elif not tool:
            tool = args[i]; i += 1
        else:
            i += 1

    if not tool:
        print("用法: t2cad_eval.py init --project <项目> --tool <工具> [--suite <套件>]")
        print("可用工具: excel, cad, word, ppt")
        print("可用套件: gold_set (默认), edge_cases, regression")
        return

    samples = {
        "excel": [
            {"name": "简单排序", "description": "按数值列降序排列",
             "snapshot": "Sheet 1: 销售表\nA1: 姓名 B1: 销售额\nA2: 张三 B2: 15000\nA3: 李四 B3: 23000\nA4: 王五 B4: 18000",
             "user_input": "按销售额从高到低排序，排完后美化表格",
             "checks": {"must_use": ["V", "BEAUTIFY"], "must_not_use": ["import ", "Dispatch"],
                        "expected_operations": ["sort", "Sort", "排序"], "should_contain": ["BEAUTIFY"]},
             "difficulty": "easy", "tags": ["sort", "basic"]},
            {"name": "分类汇总", "description": "按部门汇总工资",
             "snapshot": "Sheet 1: 工资表\nA1: 部门 B1: 姓名 C1: 工资\nA2: 技术部 B2: 张三 C2: 15000\nA3: 销售部 B3: 李四 C3: 12000\nA4: 技术部 B4: 王五 C4: 18000",
             "user_input": "按部门分类汇总工资总和，新建一个工作表放结果，表头叫'部门工资汇总'",
             "checks": {"must_use": ["V", "NEW_SHEET"], "must_not_use": ["import ", "Dispatch"],
                        "expected_operations": ["NEW_SHEET", "sum", "汇总"], "should_contain": ["BEAUTIFY"]},
             "difficulty": "medium", "tags": ["aggregate", "groupby"]},
            {"name": "条件格式化", "description": "高亮低于阈值的单元格",
             "snapshot": "Sheet 1: 质检表\nA1: 批次 B1: 合格率\nA2: B001 B2: 0.95\nA3: B002 B3: 0.82\nA4: B003 B4: 0.97",
             "user_input": "把合格率低于0.9的单元格标红底色，表头加粗居中，加边框",
             "checks": {"must_use": ["V", "BEAUTIFY", "BG"], "must_not_use": ["import ", "Dispatch"],
                        "expected_operations": ["BG", "条件", "0.9"], "should_contain": ["BEAUTIFY", "BG"]},
             "difficulty": "medium", "tags": ["format", "conditional"]},
            {"name": "创建图表", "description": "根据数据创建柱状图",
             "snapshot": "Sheet 1: 月度销售\nA1: 月份 B1: 销售额\nA2: 1月 B2: 50000\nA3: 2月 B3: 65000\nA4: 3月 B4: 48000",
             "user_input": "根据月份和销售额创建柱状图，标题叫'Q1销售趋势'",
             "checks": {"must_use": ["V", "CHART"], "must_not_use": ["import ", "Dispatch"],
                        "expected_operations": ["CHART", "chart"], "should_contain": ["CHART"]},
             "difficulty": "medium", "tags": ["chart", "visualization"]},
            {"name": "多表数据清洗", "description": "遍历多个工作表清洗数据",
             "snapshot": "Workbook: 生产报表.xlsx\nSheet 1: 原材料 (100行)\nSheet 2: 成品 (80行)\nSheet 3: 能耗 (60行)",
             "user_input": "遍历所有工作表，删除空行，数值列保留2位小数，表头加冻结和自动筛选",
             "checks": {"must_use": ["V", "LOOP", "BEAUTIFY"], "must_not_use": ["import ", "Dispatch", "open("],
                        "expected_operations": ["LOOP", "删除", "delete"], "should_contain": ["FREEZE", "AUTOFIT"]},
             "difficulty": "hard", "tags": ["clean", "multi-sheet", "loop"]},
        ],
        "cad": [
            {"name": "简单圆和阵列", "description": "画圆并环形阵列",
             "snapshot": "空图纸。单位: 毫米。",
             "user_input": "画一个直径50的圆，以原点为中心环形阵列6个，间距60度",
             "checks": {"must_use": ["C", "LAYER"], "must_not_use": ["import ", "Dispatch", "Autocad()"],
                        "expected_operations": ["array", "Array", "Copy", "copy"], "should_contain": ["C("]},
             "difficulty": "easy", "tags": ["circle", "array", "basic"]},
        ],
    }

    if tool not in samples:
        print(f"未知工具: {tool}")
        return

    cases = samples[tool]
    save_cases(project, tool, cases, suite)
    print(f"✅ 已为 {project}/{tool} 创建 {len(cases)} 个用例 → datasets/{project}/{suite}/")


def cmd_migrate(args: list):
    """Migrate V1 cases to V2 structure."""
    print("📦 迁移 V1 → V2 目录结构...")
    migrate_v1_to_v2()
    print("✅ 迁移完成")


def main():
    if len(sys.argv) < 2:
        print("TextToCAD Eval V2 — 项目级评测引擎")
        print("")
        print("命令:")
        print("  init     创建测试用例  [--project <p>] --tool <t> [--suite gold_set]")
        print("  run      跑评测       [--project <p>] --tool <t> [--suite all] [--dry]")
        print("  compare  对比最近两次  [--project <p>] <tool>")
        print("  report   生成周报      [--project <p>] [--last 7days] [--format markdown] [--output file.md]")
        print("  migrate  迁移 V1→V2")
        print("")
        print("套件: gold_set (黄金集) / edge_cases (边界) / regression (回归) / all (全部)")
        return

    cmd = sys.argv[1]
    if cmd == "init":        cmd_init(sys.argv[1:])
    elif cmd == "run":       cmd_run(sys.argv[1:])
    elif cmd == "compare":   cmd_compare(sys.argv[1:])
    elif cmd == "report":    cmd_report(sys.argv[1:])
    elif cmd == "migrate":   cmd_migrate(sys.argv[1:])
    else:
        print(f"未知命令: {cmd}")
        print("可用: init, run, compare, report, migrate")


if __name__ == "__main__":
    main()
