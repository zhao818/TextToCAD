# t2cad_fixer.py — shared error-fixing module for TextToCAD tools
"""Claude Code expert + web search + COM error decoding.
Imported by TextToCAD_Excel/PPT/Word/AutoCAD for error-driven code fixing."""

import re, subprocess, sys

# ── error explainer: show AI code line + decode COM HRESULT ──
def explain_error(trace, code=""):
    """Show which AI code line failed + decode COM HRESULT."""
    lines = trace.strip().split("\n")
    result = []

    # 1. Show the AI code line that failed
    for line in lines:
        m = re.search(r'File "<string>", line (\d+)', line)
        if m:
            ln = int(m.group(1))
            code_lines = code.split("\n") if code else []
            if ln <= len(code_lines):
                failed_line = code_lines[ln - 1].strip()
                result.append(f">> 出错代码(第{ln}行): {failed_line[:150]}")
                ctx_start = max(0, ln - 3)
                ctx_end = min(len(code_lines), ln + 1)
                for cl in range(ctx_start, ctx_end):
                    marker = "  >" if cl == ln - 1 else "   "
                    result.append(f"{marker} L{cl+1}: {code_lines[cl].strip()[:120]}")
                break

    # 2. Decode COM HRESULT
    if "-2147352565" in trace:  # DISP_E_MEMBERNOTFOUND
        result.append("[DISP_E_MEMBERNOTFOUND] 调用了不存在的方法/属性 — 对象类型可能不对")
    if "-2147352567" in trace:  # COM failure
        result.append("[COM失败] 操作被拒绝 — 可能原因: 参数无效/权限不足/对象状态异常")
    if "-2146827284" in trace:
        result.append("[文件未找到]")

    # 3. Last error line
    for i in range(len(lines) - 1, -1, -1):
        if "Error" in lines[i] or "error" in lines[i].lower():
            result.append(lines[i].strip())
            break
    if len(result) == 0 and lines:
        result.append(lines[-1].strip())

    return "\n".join(result) if result else trace.strip()


# ── web search ─────────────────────────────────────────────
def web_search(error_text, code_snippet=""):
    """Search DuckDuckGo for solutions related to the error."""
    try:
        from ddgs import DDGS
    except ImportError:
        try:
            from duckduckgo_search import DDGS
        except ImportError:
            return None

    query_parts = ["Excel", "COM", "win32com", "VBA"]
    for kw in ["NameError", "com_error", "pywintypes", "UsedRange",
               "Workbooks", "Worksheets", "DISP_E_MEMBERNOTFOUND", "AttributeError"]:
        if kw in error_text:
            query_parts.append(kw)
            break

    lines = error_text.strip().split("\n")
    last_line = lines[-1] if lines else error_text
    query_parts.append(last_line[:80])
    query = " ".join(query_parts[:8])

    try:
        results = []
        with DDGS() as ddgs:
            for r in ddgs.text(query, max_results=3, region="cn-zh"):
                snippet = r.get("body", "")[:200]
                results.append(f"- {snippet}\n  {r.get('href', '')}")
        if results:
            return "网上查到的相关资料:\n" + "\n\n".join(results)
    except Exception:
        pass
    return None


# ── Claude Code CLI fixer ──────────────────────────────────
def claude_fix(expert_prompt, error_text, failed_code, search_hint):
    """Use local Claude Code CLI as expert fixer."""
    prompt = expert_prompt + "\n\n## 报错信息\n" + error_text
    prompt += "\n\n## 失败的代码\n" + failed_code
    if search_hint:
        prompt += "\n\n## 网上查到的资料\n" + search_hint
    prompt += "\n\n只输出修正后的完整Python代码，不要解释，不要markdown标记。"

    try:
        result = subprocess.run(
            'claude -p --output-format text',
            input=prompt, capture_output=True, text=True, timeout=120,
            shell=True
        )
        if result.returncode == 0 and result.stdout.strip():
            fixed = result.stdout.strip()
            fixed = re.sub(r'^```(?:python)?\s*\n?', '', fixed)
            fixed = re.sub(r'\n?```\s*$', '', fixed)
            return fixed if fixed and len(fixed) > 20 else None
    except Exception:
        pass
    return None


# ── API fallback fixer ─────────────────────────────────────
def api_fix(expert_prompt, error_text, failed_code, search_hint, cfg, proxies):
    """Use API-based LLM as fallback fixer."""
    try:
        import requests as req

        user_msg = f"## 报错信息\n{error_text}\n\n## 失败的代码\n{failed_code}\n\n"
        if search_hint:
            user_msg += f"## 网上查到的资料\n{search_hint}\n\n"
        user_msg += "请提供修正后的完整代码。"

        fixer_api_base = cfg.get("fixer_api_base", "").strip() or cfg["api_base"]
        fixer_api_key = cfg.get("fixer_api_key", "").strip() or cfg["api_key"]
        fixer_model = cfg.get("fixer_model", "").strip() or cfg["model"]

        headers = {"Content-Type": "application/json",
                   "Authorization": f"Bearer {fixer_api_key}"}
        body = {"model": fixer_model, "messages": [
            {"role": "system", "content": expert_prompt},
            {"role": "user", "content": user_msg},
        ], "temperature": 0.0, "max_tokens": cfg.get("max_tokens", 4096)}

        url = f"{fixer_api_base.rstrip('/')}/chat/completions"
        resp = req.post(url, headers=headers, json=body, timeout=90, proxies=proxies)
        resp.raise_for_status()
        fixed = resp.json()["choices"][0]["message"]["content"]
        fixed = re.sub(r'^```(?:python)?\s*\n?', '', fixed.strip())
        fixed = re.sub(r'\n?```\s*$', '', fixed)
        return fixed if fixed else None
    except Exception:
        return None


# ── combined fixer (tries Claude CLI first, then API) ─────
def fix_code(expert_prompt, error_text, failed_code, search_hint, cfg, proxies):
    """Fix code using best available expert. Returns corrected code or None."""
    # 1. Claude Code CLI (local, no API key)
    fixed = claude_fix(expert_prompt, error_text, failed_code, search_hint)
    if fixed:
        return fixed
    # 2. API fallback
    return api_fix(expert_prompt, error_text, failed_code, search_hint, cfg, proxies)
