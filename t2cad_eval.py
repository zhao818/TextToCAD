#!/usr/bin/env python3
# t2cad_eval.py — TextToCAD 评测引擎
"""
Eval-driven development for TextToCAD tools.

Architecture:
  t2cad_eval.py run <tool>     → 跑全量测试用例
  t2cad_eval.py compare        → 对比两次评测结果（提示词迭代前后）
  t2cad_eval.py report         → 生成评测报告（含趋势图数据）

Test case format (eval/cases/<tool>_cases.json):
  {
    "name": "简单排序",
    "description": "把A列按数值降序排列",
    "snapshot": "Sheet 1: 数据表\\nA1: 姓名 B1: 销售额",
    "user_input": "按销售额从高到低排序",
    "checks": {
      "must_use": ["V"],
      "must_not_use": ["import ", "Dispatch", "open(", "print("],
      "expected_operations": ["sort", "Sort", "排序"],
      "should_contain": ["BEAUTIFY"]
    },
    "difficulty": "easy",
    "tags": ["sort", "basic"]
  }

Scoring:
  Syntax     (20%) — code passes ast.parse()
  Safety     (30%) — no banned patterns, uses safe functions
  Task Fit   (30%) — contains expected operations
  Quality    (20%) — readability, naming, structure
  Overall    ≥ 90% → PASS

Usage:
  python t2cad_eval.py run excel        # 跑 Excel 全部用例
  python t2cad_eval.py run excel --tag sort  # 只跑排序相关
  python t2cad_eval.py run excel --dry   # 不调 LLM，只评测已有代码
  python t2cad_eval.py compare           # 对比最近两次
  python t2cad_eval.py report --tool excel  # 生成报告
"""

import ast
import hashlib
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

# ── Paths ──────────────────────────────────────────────
CONFIG_DIR = Path.home() / ".text_to_cad"
EVAL_DIR = CONFIG_DIR / "eval"
CASES_DIR = EVAL_DIR / "cases"
RESULTS_DIR = EVAL_DIR / "results"
EVAL_CONFIG_FILE = EVAL_DIR / "eval_config.json"

for d in [EVAL_DIR, CASES_DIR, RESULTS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

sys.path.insert(0, str(CONFIG_DIR))

# ── Scoring weights ────────────────────────────────────
WEIGHTS = {
    "syntax": 0.20,
    "safety": 0.30,
    "task_fit": 0.30,
    "quality": 0.20,
}

PASS_THRESHOLD = 0.90

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
    (r'\bWorkbooks\.Open\b', "禁止 Workbooks.Open（工作簿已打开）"),
    (r'\bPresentations\.Open\b', "禁止 Presentations.Open"),
    (r'\bDocuments\.Open\b', "禁止 Documents.Open"),
]

# ── Quality: anti-patterns ─────────────────────────────
QUALITY_ANTI_PATTERNS = [
    (r'except\s*:', "裸 except: 应该指定具体异常类型"),
    (r'except\s+Exception\s*:', "裸 catch Exception 太宽泛"),
    (r'\bwhile\s+True\b', "while True 可能死循环"),
    (r'[a-z]\s*=\s*[a-z]\s*=\s*[a-z]', "一行多个赋值影响可读性"),
]

QUALITY_POSITIVE = [
    (r'#', "包含注释"),
    (r'def\s+\w+', "使用函数封装"),
    (r'[A-Z][a-z]+(?:[A-Z][a-z]+)+', "变量名使用驼峰/下划线"),
]


# ════════════════════════════════════════════════════════
# Test Case Management
# ════════════════════════════════════════════════════════

def load_cases(tool_name: str) -> list[dict]:
    """Load test cases for a tool."""
    case_file = CASES_DIR / f"{tool_name}_cases.json"
    if not case_file.exists():
        return []
    with open(case_file, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("cases", [])


def save_cases(tool_name: str, cases: list[dict]):
    """Save test cases for a tool."""
    CASES_DIR.mkdir(parents=True, exist_ok=True)
    case_file = CASES_DIR / f"{tool_name}_cases.json"
    data = {
        "tool": tool_name,
        "version": 1,
        "updated": datetime.now().isoformat(),
        "count": len(cases),
        "cases": cases,
    }
    with open(case_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ════════════════════════════════════════════════════════
# Scoring Functions
# ════════════════════════════════════════════════════════

def check_syntax(code: str) -> tuple[float, str]:
    """Check if code is valid Python syntax. Returns (score, detail)."""
    try:
        ast.parse(code)
        return 1.0, "语法正确"
    except SyntaxError as e:
        return 0.0, f"语法错误: {e}"


def check_safety(code: str) -> tuple[float, str]:
    """Check for banned patterns. Returns (score, detail)."""
    violations = []
    for pattern, desc in BANNED_PATTERNS:
        if re.search(pattern, code):
            violations.append(f"  ❌ {desc}")

    if not violations:
        return 1.0, "安全: 无违规"

    score = max(0.0, 1.0 - len(violations) * 0.15)
    return score, "安全违规:\n" + "\n".join(violations)


def check_task_fit(code: str, case: dict) -> tuple[float, str]:
    """Check if code matches expected task patterns. Returns (score, detail)."""
    checks = case.get("checks", {})
    score = 1.0
    details = []

    # Must-use functions
    must_use = checks.get("must_use", [])
    for func in must_use:
        if func not in code:
            score -= 0.15
            details.append(f"  ❌ 缺少必要函数: {func}")
        else:
            details.append(f"  ✅ 使用: {func}")

    # Must-not-use
    must_not = checks.get("must_not_use", [])
    for pattern in must_not:
        if pattern in code:
            score -= 0.2
            details.append(f"  ❌ 不应使用: {pattern}")

    # Expected operations
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

    # Should contain (bonus, not penalty)
    should = checks.get("should_contain", [])
    for item in should:
        if item in code:
            details.append(f"  ✅ 含推荐项: {item}")

    return max(0.0, min(1.0, score)), "任务匹配:\n" + "\n".join(details) if details else "任务匹配: 无检查项"


def check_quality(code: str) -> tuple[float, str]:
    """Check code quality. Returns (score, detail)."""
    score = 0.5  # baseline
    details = []

    # Anti-patterns
    for pattern, desc in QUALITY_ANTI_PATTERNS:
        if re.search(pattern, code):
            score -= 0.1
            details.append(f"  ⚠️ {desc}")

    # Positive patterns
    for pattern, desc in QUALITY_POSITIVE:
        if re.search(pattern, code):
            score += 0.1
            details.append(f"  ✅ {desc}")

    # Length check
    lines = code.strip().split("\n")
    if len(lines) < 3:
        score -= 0.2
        details.append("  ⚠️ 代码过短（可能不完整）")
    elif len(lines) > 200:
        score -= 0.1
        details.append("  ⚠️ 代码过长（>200行）")

    # Variable naming
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
    """Run test cases through LLM pipeline and score results."""

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
        """Evaluate a single test case."""
        start_time = time.time()

        if self.dry_run:
            # Dry run: use pre-existing code if provided
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
                max_retries=3,  # Eval uses fewer retries for speed
            )

        code = pipeline_result.get("code", "")
        elapsed = time.time() - start_time

        # Score
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

    def run_suite(self, tool_name: str, system_prompt: str,
                  fixer_prompt: str = "", exec_ns: dict = None,
                  tag_filter: str = None, limit: int = None) -> dict:
        """Run all test cases for a tool."""
        cases = load_cases(tool_name)
        if not cases:
            return {"error": f"未找到 {tool_name} 的测试用例，请先在 {CASES_DIR}/{tool_name}_cases.json 创建"}

        # Filter
        if tag_filter:
            cases = [c for c in cases if tag_filter in c.get("tags", [])]
        if limit:
            cases = cases[:limit]

        results = []
        passed = 0
        failed = 0

        for i, case in enumerate(cases):
            print(f"\n{'─'*60}")
            print(f"[{i+1}/{len(cases)}] {case['name']} ({case.get('difficulty', '?')})")
            print(f"      输入: {case['user_input'][:80]}...")

            if not self.dry_run:
                print(f"      ⏳ 生成代码...")

            result = self.evaluate_one(case, system_prompt, fixer_prompt, exec_ns)
            results.append(result)

            status = "✅ PASS" if result["passed"] else "❌ FAIL"
            print(f"      {status} 综合: {result['scores']['overall']:.1%}  "
                  f"语法:{result['scores']['syntax']:.0%} 安全:{result['scores']['safety']:.0%}  "
                  f"任务:{result['scores']['task_fit']:.0%} 质量:{result['scores']['quality']:.0%}  "
                  f"尝试:{result['pipeline']['attempts']}次 {result['elapsed_sec']}s")

            if not result["passed"]:
                failed += 1
                # Show failures
                for dim in ["syntax", "safety", "task_fit", "quality"]:
                    if result["scores"][dim] < 0.9:
                        print(f"      [{dim}] {result['details'][dim]}")
            else:
                passed += 1

        # Summary
        avg = sum(r["scores"]["overall"] for r in results) / len(results) if results else 0
        summary = {
            "tool": tool_name,
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

        return {"summary": summary, "results": results}


# ════════════════════════════════════════════════════════
# Results Management
# ════════════════════════════════════════════════════════

def save_result(tool_name: str, result: dict, label: str = ""):
    """Save eval result to disk."""
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{tool_name}_{ts}.json"
    if label:
        filename = f"{tool_name}_{ts}_{label}.json"

    result["meta"] = {
        "tool": tool_name,
        "timestamp": datetime.now().isoformat(),
        "label": label,
        "eval_version": 1,
    }

    filepath = RESULTS_DIR / filename
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    return filepath


def load_latest_result(tool_name: str) -> Optional[dict]:
    """Load the most recent eval result for a tool."""
    files = sorted(RESULTS_DIR.glob(f"{tool_name}_*.json"), reverse=True)
    if not files:
        return None
    with open(files[0], "r", encoding="utf-8") as f:
        return json.load(f)


def list_results(tool_name: str = None) -> list[dict]:
    """List all saved eval results."""
    pattern = f"{tool_name}_*.json" if tool_name else "*.json"
    files = sorted(RESULTS_DIR.glob(pattern), reverse=True)
    results = []
    for f in files:
        with open(f, "r", encoding="utf-8") as fp:
            data = json.load(fp)
            meta = data.get("meta", {})
            summary = data.get("summary", {})
            results.append({
                "file": f.name,
                "tool": meta.get("tool", "?"),
                "timestamp": meta.get("timestamp", "?"),
                "label": meta.get("label", ""),
                "pass_rate": summary.get("pass_rate", 0),
                "avg_overall": summary.get("avg_overall", 0),
                "total": summary.get("total", 0),
            })
    return results


def compare_results(tool_name: str) -> dict:
    """Compare the two most recent eval results."""
    files = sorted(RESULTS_DIR.glob(f"{tool_name}_*.json"), reverse=True)
    if len(files) < 2:
        return {"error": f"需要至少2次评测结果才能对比，当前只有 {len(files)} 次"}

    with open(files[0], "r", encoding="utf-8") as f:
        new = json.load(f)
    with open(files[1], "r", encoding="utf-8") as f:
        old = json.load(f)

    new_sum = new["summary"]
    old_sum = old["summary"]

    diff = {
        "tool": tool_name,
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
    }

    # Per-case comparison
    new_cases = {r["case_name"]: r for r in new["results"]}
    old_cases = {r["case_name"]: r for r in old["results"]}
    for name in new_cases:
        if name in old_cases:
            delta = round(new_cases[name]["scores"]["overall"] - old_cases[name]["scores"]["overall"], 3)
            diff["per_case_delta"].append({
                "case": name,
                "old": old_cases[name]["scores"]["overall"],
                "new": new_cases[name]["scores"]["overall"],
                "delta": delta,
                "direction": "📈" if delta > 0 else ("📉" if delta < 0 else "➡️"),
            })

    return diff


# ════════════════════════════════════════════════════════
# Tool-specific helpers
# ════════════════════════════════════════════════════════

def get_system_prompt(tool_name: str) -> str:
    """Get the default system prompt for a tool from its config."""
    # Try to load from config
    cfg_file = CONFIG_DIR / "config.json"
    if cfg_file.exists():
        with open(cfg_file, "r", encoding="utf-8") as f:
            cfg = json.load(f)

        key = f"system_prompt_zh"
        if key in cfg:
            return cfg[key]

    # Fallback: minimal prompts
    fallbacks = {
        "excel": "你是 Excel 专家。操作 excel/wb/ws。用 V() FIX() BEAUTIFY() 等安全函数。只输出 Python 代码。",
        "cad": "你是 AutoCAD 专家。操作 acad 对象。用 L() C() PLINE() 等安全函数。只输出 Python 代码。",
        "word": "你是 Word 专家。操作 word/doc/sel。用 TB() MERGE() 等安全函数。只输出 Python 代码。",
        "ppt": "你是 PPT 专家。操作 ppt/pres/slide。用 TB_STYLE() SHAPE() 等安全函数。只输出 Python 代码。",
    }
    return fallbacks.get(tool_name, "你是技术专家。只输出 Python 代码。")


def get_fixer_prompt(tool_name: str) -> str:
    """Get default fixer prompt for a tool."""
    fallbacks = {
        "excel": "你是 Excel COM 专家。修正代码。只输出代码。",
        "cad": "你是 AutoCAD COM 专家。修正代码。只输出代码。",
        "word": "你是 Word COM 专家。修正代码。只输出代码。",
        "ppt": "你是 PPT COM 专家。修正代码。只输出代码。",
    }
    return fallbacks.get(tool_name, "你是调试专家。修正代码。只输出代码。")


def get_exec_namespace(tool_name: str) -> dict:
    """Get a mock exec namespace for safe evaluation."""
    # Mock namespace — no real COM objects, just safe function stubs
    base = {
        "math": __import__("math"),
        "re": __import__("re"),
    }

    # Common safe functions (stub versions for syntax validation)
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
        "FAST_MODE": lambda b: None,
        "L": lambda x1, y1, x2, y2, color=None: None,
        "C": lambda x, y, r, color=None: None,
        "R": lambda x, y, w, h, color=None: None,
        "T": lambda x, y, text, size=None: None,
        "DIM": lambda *args: None,
        "PLINE": lambda *args: None,
        "ARC": lambda *args: None,
        "HATCH": lambda *args: None,
        "MTEXT": lambda *args: None,
        "MOVE": lambda obj, dx, dy: None,
        "ROT": lambda obj, angle: None,
        "DEL": lambda obj: None,
        "OBJS": lambda: [],
        "FIND_OBJS": lambda **kw: [],
        "DEL_BY_TYPE": lambda t: None,
        "UNDO": lambda: None,
        "LAYER": lambda name, color=None: None,
        "SEND_CMD": lambda cmd: None,
        "ZOOM_EXT": lambda: None,
        "ZOOM_WIN": lambda *args: None,
        "BLOCK_INSERT": lambda *args: None,
        "SET_COLOR": lambda obj, color: None,
        "DIM_H": lambda *args: None,
        "DIM_STYLE": lambda: None,
    }

    word_funcs = {
        "TB": lambda *args, **kw: None,
        "MERGE_CELLS": lambda t, r1, c1, r2, c2: None,
        "FMT_TABLE": lambda t, **kw: None,
        "SECTIONS": lambda doc: [],
        "IMG": lambda path, x, y, w, h: None,
    }

    ppt_funcs = {
        "NEW_SLIDE": lambda layout=12: None,
        "GOTO": lambda n: None,
        "DEL_SLIDE": lambda n: None,
        "CLEAR": lambda: None,
        "TB": lambda x, y, w, h, text, size=32, color=None: None,
        "TB_STYLE": lambda x, y, w, h, text, size, fc, fill, align=None: None,
        "SHAPE": lambda t, x, y, w, h, color=None: None,
        "FILL": lambda sh, color: None,
        "FONT_COLOR": lambda sh, color: None,
        "TABLE_SLIDE": lambda r, c, x, y, w, h: None,
        "CELL_SLIDE": lambda t, r, c, text: None,
        "TABLE_STYLE": lambda t, color: None,
        "IMG": lambda path, x, y, w, h: None,
        "ALIGN_SHAPE": lambda sh, align: None,
        "Z_ORDER": lambda sh, pos: None,
        "COLOR": lambda r, g, b: 0,
        "PTS": lambda mm: mm * 2.835,
    }

    ns = dict(base)
    ns.update(common)

    if tool_name == "cad":
        ns.update(cad_funcs)
    elif tool_name == "word":
        ns.update(word_funcs)
    elif tool_name == "ppt":
        ns.update(ppt_funcs)

    return ns


# ════════════════════════════════════════════════════════
# CLI
# ════════════════════════════════════════════════════════

def print_summary(result: dict):
    """Print a formatted summary."""
    s = result["summary"]
    print(f"\n{'='*60}")
    print(f"📊 评测总结: {s['tool']}")
    print(f"{'='*60}")
    print(f"  通过率:    {s['pass_rate']:.1%} ({s['passed']}/{s['total']})")
    print(f"  综合均分:  {s['avg_overall']:.1%}")
    print(f"  语法:      {s['avg_syntax']:.1%}")
    print(f"  安全:      {s['avg_safety']:.1%}")
    print(f"  任务匹配:  {s['avg_task_fit']:.1%}")
    print(f"  代码质量:  {s['avg_quality']:.1%}")
    print(f"  平均尝试:  {s['avg_attempts']} 次")
    print(f"  平均耗时:  {s['avg_elapsed']}s")
    if s.get("dry_run"):
        print(f"  ⚠️  干跑模式（未调LLM）")

    grade = "🏆" if s['pass_rate'] >= 0.9 else ("✅" if s['pass_rate'] >= 0.7 else "⚠️")
    verdict = "PASS" if s['pass_rate'] >= PASS_THRESHOLD else "FAIL"
    print(f"\n  {grade} 最终判定: {verdict} (阈值 {PASS_THRESHOLD:.0%})")

    # Per-case breakdown
    print(f"\n{'─'*60}")
    print(f"📋 逐项明细:")
    for r in result["results"]:
        icon = "✅" if r["passed"] else "❌"
        print(f"  {icon} {r['case_name']:30s} {r['scores']['overall']:.1%}  "
              f"语{r['scores']['syntax']:.0%} 安{r['scores']['safety']:.0%} "
              f"任{r['scores']['task_fit']:.0%} 质{r['scores']['quality']:.0%}")


def print_compare(diff: dict):
    """Print a comparison between two eval runs."""
    d = diff["delta"]
    print(f"\n{'='*60}")
    print(f"📊 评测对比: {diff['tool']}")
    print(f"{'='*60}")
    print(f"  旧: {diff['old']['timestamp']} ({diff['old']['label']})")
    print(f"  新: {diff['new']['timestamp']} ({diff['new']['label']}）")
    print(f"\n  维度变化:")
    for key, label in [("pass_rate", "通过率"), ("avg_overall", "综合"),
                        ("avg_syntax", "语法"), ("avg_safety", "安全"),
                        ("avg_task_fit", "任务匹配"), ("avg_quality", "质量")]:
        delta = d[key]
        arrow = "📈" if delta > 0 else ("📉" if delta < 0 else "➡️")
        print(f"  {arrow} {label}: {delta:+.1%}")

    print(f"\n  逐用例变化:")
    for item in diff["per_case_delta"]:
        print(f"  {item['direction']} {item['case']:30s} {item['old']:.1%} → {item['new']:.1%} ({item['delta']:+.1%})")


def cmd_run(args: list):
    """Run eval suite."""
    tool = None
    tag = None
    limit = None
    dry = False
    label = ""

    i = 1
    while i < len(args):
        if args[i] == "--tag" and i + 1 < len(args):
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
        print("用法: t2cad_eval.py run <工具名> [--tag <标签>] [--limit N] [--dry] [--label <备注>]")
        print("可用工具: excel, cad, word, ppt")
        return

    sys_prompt = get_system_prompt(tool)
    fixer = get_fixer_prompt(tool)
    ns = get_exec_namespace(tool)

    print(f"🚀 评测: {tool}")
    print(f"   提示词: {sys_prompt[:80]}...")
    if dry:
        print(f"   ⚠️  干跑模式（不调LLM）")

    engine = EvalEngine(dry_run=dry)
    result = engine.run_suite(tool, sys_prompt, fixer, ns, tag_filter=tag, limit=limit)

    if "error" in result:
        print(f"❌ {result['error']}")
        return

    print_summary(result)
    filepath = save_result(tool, result, label)
    print(f"\n💾 结果已保存: {filepath}")


def cmd_compare(args: list):
    """Compare two most recent eval runs."""
    tool = args[1] if len(args) > 1 else None
    if not tool:
        # Try to find from most recent results
        results = list_results()
        if results:
            tool = results[0]["tool"]
        else:
            print("未找到评测结果。先运行: python t2cad_eval.py run <工具名>")
            return

    diff = compare_results(tool)
    if "error" in diff:
        print(f"❌ {diff['error']}")
        return
    print_compare(diff)


def cmd_report(args: list):
    """Generate eval report."""
    tool = None
    for i, a in enumerate(args):
        if a == "--tool" and i + 1 < len(args):
            tool = args[i + 1]
            break
    if not tool and len(args) > 1:
        tool = args[1]

    results = list_results(tool)
    if not results:
        print("未找到评测结果")
        return

    print(f"\n📋 评测历史 ({len(results)} 次):")
    print(f"{'─'*80}")
    print(f"{'时间':<22} {'工具':<10} {'通过率':>8} {'均分':>8} {'用例':>5} {'标签'}")
    print(f"{'─'*80}")
    for r in results[:20]:
        print(f"{r['timestamp']:<22} {r['tool']:<10} {r['pass_rate']:>7.1%} {r['avg_overall']:>7.1%} {r['total']:>5} {r['label']}")


def cmd_init(args: list):
    """Initialize sample test cases for a tool."""
    tool = args[1] if len(args) > 1 else None
    if not tool:
        print("用法: t2cad_eval.py init <工具名>")
        print("可用: excel, cad, word, ppt")
        return

    samples = {
        "excel": [
            {
                "name": "简单排序",
                "description": "按数值列降序排列",
                "snapshot": "Sheet 1: 销售表\nA1: 姓名 B1: 销售额\nA2: 张三 B2: 15000\nA3: 李四 B3: 23000\nA4: 王五 B4: 18000",
                "user_input": "按销售额从高到低排序，排完后美化表格",
                "checks": {
                    "must_use": ["V", "BEAUTIFY"],
                    "must_not_use": ["import ", "Dispatch"],
                    "expected_operations": ["sort", "Sort", "排序"],
                    "should_contain": ["BEAUTIFY"]
                },
                "difficulty": "easy",
                "tags": ["sort", "basic"]
            },
            {
                "name": "分类汇总",
                "description": "按部门汇总工资",
                "snapshot": "Sheet 1: 工资表\nA1: 部门 B1: 姓名 C1: 工资\nA2: 技术部 B2: 张三 C2: 15000\nA3: 销售部 B3: 李四 C3: 12000\nA4: 技术部 B4: 王五 C4: 18000",
                "user_input": "按部门分类汇总工资总和，新建一个工作表放结果，表头叫'部门工资汇总'",
                "checks": {
                    "must_use": ["V", "NEW_SHEET"],
                    "must_not_use": ["import ", "Dispatch"],
                    "expected_operations": ["NEW_SHEET", "sum", "汇总"],
                    "should_contain": ["BEAUTIFY"]
                },
                "difficulty": "medium",
                "tags": ["aggregate", "groupby"]
            },
            {
                "name": "条件格式化",
                "description": "高亮大于阈值的单元格",
                "snapshot": "Sheet 1: 质检表\nA1: 批次 B1: 合格率\nA2: B001 B2: 0.95\nA3: B002 B3: 0.82\nA4: B003 B4: 0.97",
                "user_input": "把合格率低于0.9的单元格标红底色，表头加粗居中，加边框",
                "checks": {
                    "must_use": ["V", "BEAUTIFY", "BG"],
                    "must_not_use": ["import ", "Dispatch"],
                    "expected_operations": ["BG", "条件", "0.9"],
                    "should_contain": ["BEAUTIFY", "BG"]
                },
                "difficulty": "medium",
                "tags": ["format", "conditional"]
            },
            {
                "name": "创建图表",
                "description": "根据数据创建柱状图",
                "snapshot": "Sheet 1: 月度销售\nA1: 月份 B1: 销售额\nA2: 1月 B2: 50000\nA3: 2月 B3: 65000\nA4: 3月 B4: 48000",
                "user_input": "根据月份和销售额创建柱状图，标题叫'Q1销售趋势'",
                "checks": {
                    "must_use": ["V", "CHART"],
                    "must_not_use": ["import ", "Dispatch"],
                    "expected_operations": ["CHART", "chart"],
                    "should_contain": ["CHART"]
                },
                "difficulty": "medium",
                "tags": ["chart", "visualization"]
            },
            {
                "name": "多表数据清洗",
                "description": "遍历多个工作表清洗数据",
                "snapshot": "Workbook: 生产报表.xlsx\nSheet 1: 原材料 (100行: 日期/物料/数量/单价)\nSheet 2: 成品 (80行: 日期/产品/产量/合格数)\nSheet 3: 能耗 (60行: 日期/电/水/气)",
                "user_input": "遍历所有工作表，把每个表的空行删除，数值列统一保留2位小数，表头加冻结和自动筛选",
                "checks": {
                    "must_use": ["V", "LOOP", "BEAUTIFY"],
                    "must_not_use": ["import ", "Dispatch", "open("],
                    "expected_operations": ["LOOP", "删除", "delete", "Delete"],
                    "should_contain": ["FREEZE", "AUTOFIT"]
                },
                "difficulty": "hard",
                "tags": ["clean", "multi-sheet", "loop"]
            },
        ],
        "cad": [
            {
                "name": "简单圆和阵列",
                "description": "画圆并环形阵列",
                "snapshot": "空图纸。单位: 毫米。模型空间无对象。",
                "user_input": "画一个直径50的圆，以原点为中心环形阵列6个，间距60度",
                "checks": {
                    "must_use": ["C", "LAYER"],
                    "must_not_use": ["import ", "Dispatch", "Autocad()"],
                    "expected_operations": ["array", "Array", "Copy", "copy", "Rotate"],
                    "should_contain": ["C("]
                },
                "difficulty": "easy",
                "tags": ["circle", "array", "basic"]
            },
            {
                "name": "建筑平面图",
                "description": "画简单建筑平面",
                "snapshot": "空图纸。单位: 毫米。",
                "user_input": "画一个10米×8米的矩形建筑轮廓，墙厚240mm，在四个角放直径400mm的圆柱",
                "checks": {
                    "must_use": ["L", "C", "LAYER"],
                    "must_not_use": ["import ", "Dispatch"],
                    "expected_operations": ["L(", "C(", "rect"],
                    "should_contain": ["LAYER"]
                },
                "difficulty": "medium",
                "tags": ["architecture", "wall", "column"]
            },
        ],
        "word": [
            {
                "name": "生成标准报告",
                "description": "创建带格式的文档报告",
                "snapshot": "空白文档。页边距: 上下25.4mm 左右31.8mm。",
                "user_input": "创建报告：标题'Q2生产总结'居中加粗24pt，下面一个3列4行的表格(日期/产量/合格率)，填充示例数据，表格表头深蓝底白字",
                "checks": {
                    "must_use": [],
                    "must_not_use": ["import ", "Dispatch", "Documents.Open"],
                    "expected_operations": ["Table", "table", "表格", "Add"],
                    "should_contain": ["Bold", "Color", "Heading"]
                },
                "difficulty": "medium",
                "tags": ["report", "table", "basic"]
            },
        ],
        "ppt": [
            {
                "name": "生成三页演示",
                "description": "创建封面+内容+结束页",
                "snapshot": "空演示文稿。16:9 (960x540)。",
                "user_input": "生成3页PPT: 封面(深蓝背景+白色大标题'AI赋能制造业')，内容页(3个要点:降本/增效/提质)，结束页(谢谢+联系方式)",
                "checks": {
                    "must_use": ["NEW_SLIDE", "TB_STYLE"],
                    "must_not_use": ["import ", "Dispatch", "Presentations.Open"],
                    "expected_operations": ["NEW_SLIDE", "TB_STYLE", "封面", "标题"],
                    "should_contain": ["FILL", "FONT_COLOR"]
                },
                "difficulty": "medium",
                "tags": ["slides", "presentation", "basic"]
            },
        ],
    }

    if tool not in samples:
        print(f"未知工具: {tool}. 可用: {', '.join(samples.keys())}")
        return

    cases = samples[tool]
    save_cases(tool, cases)
    print(f"✅ 已为 {tool} 创建 {len(cases)} 个示例测试用例")
    print(f"   路径: {CASES_DIR / f'{tool}_cases.json'}")
    print(f"   下一步: python t2cad_eval.py run {tool}")


def main():
    if len(sys.argv) < 2:
        print("TextToCAD Eval 评测引擎")
        print("用法:")
        print("  python t2cad_eval.py init <工具名>         创建示例测试用例")
        print("  python t2cad_eval.py run <工具名>          跑全量评测")
        print("  python t2cad_eval.py run <工具名> --dry    干跑（不调LLM）")
        print("  python t2cad_eval.py compare [工具名]      对比最近两次")
        print("  python t2cad_eval.py report [--tool 工具名] 评测历史")
        print("")
        print("可用工具: excel, cad, word, ppt")
        return

    cmd = sys.argv[1]
    if cmd == "init":
        cmd_init(sys.argv[1:])
    elif cmd == "run":
        cmd_run(sys.argv[1:])
    elif cmd == "compare":
        cmd_compare(sys.argv[1:])
    elif cmd == "report":
        cmd_report(sys.argv[1:])
    else:
        print(f"未知命令: {cmd}")
        print("可用: init, run, compare, report")


if __name__ == "__main__":
    main()
