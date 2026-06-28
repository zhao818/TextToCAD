# t2cad_pipeline.py — LangGraph code-generation pipeline for TextToCAD
"""
Replaces the ~100-line retry loop duplicated in every tool with a unified
LangGraph StateGraph: generate → validate → execute → fix → retry.

Each tool provides:
  - system_prompt: str
  - snapshot: str (drawing overview / workbook snapshot / document structure)
  - user_input: str
  - exec_namespace: dict (safe helper functions + COM objects)
  - fixer_prompt: str (system prompt for the code-fixing LLM)

The pipeline returns:
  - success: bool
  - code: str (the final generated code)
  - result: str (success message or error trace)

Usage in a tool's on_run():
  from t2cad_pipeline import CodeGenPipeline

  pipeline = CodeGenPipeline(client)
  result = pipeline.run(
      system_prompt=cfg["system_prompt_zh"],
      snapshot=overview,
      user_input=user_input,
      exec_namespace=ns,  # {excel, wb, ws, V, F, ERR, ...}
      fixer_prompt=FIXER_PROMPT,
      cancel_check=lambda: self._cancelled,
      progress_callback=lambda tag, color: self.set_status(tag, color),
      code_callback=lambda code: self.output_edit.setText(code),
  )
"""

from typing import Any, Callable, Optional
from typing_extensions import TypedDict

from t2cad_llm import LLMClient, strip_code_fence


# ---------------------------------------------------------------------------
# State definition
# ---------------------------------------------------------------------------
class PipelineState(TypedDict):
    """State flowing through the LangGraph pipeline nodes."""
    # Inputs
    system_prompt: str
    snapshot: str
    user_input: str
    fixer_prompt: str
    exec_namespace: dict
    max_retries: int

    # Dynamic
    messages: list[dict]        # Full conversation with LLM
    code: str                   # Current generated/fixed code
    error: str                  # Last execution error (empty if success)
    search_hint: str            # Web search results for the error
    attempt: int                # Current attempt number (0-based)

    # Output
    success: bool
    result: str                 # "success" or error trace
    all_codes: list[str]        # History of all generated codes (for debugging)


# ---------------------------------------------------------------------------
# Pipeline runner
# ---------------------------------------------------------------------------
class CodeGenPipeline:
    """LangGraph-inspired code generation pipeline with retry logic.

    This is a simplified implementation that mirrors LangGraph's state-machine
    pattern without requiring langgraph as a hard dependency. It provides the
    same structured flow: generate → execute → analyze → fix → retry.
    """

    def __init__(self, client: LLMClient = None):
        from t2cad_llm import get_client
        self.client = client or get_client()

    # ── Public API ──────────────────────────────────────────

    def run(
        self,
        system_prompt: str,
        snapshot: str,
        user_input: str,
        exec_namespace: dict,
        fixer_prompt: str = "",
        max_retries: int = 6,
        cancel_check: Callable[[], bool] = None,
        progress_callback: Callable[[str, str], None] = None,
        code_callback: Callable[[str], None] = None,
    ) -> dict:
        """Run the full generate→execute→fix→retry pipeline.

        Returns:
            {"success": bool, "code": str, "result": str, "attempt": int}
        """
        # ── Build initial state ──
        state: PipelineState = {
            "system_prompt": system_prompt,
            "snapshot": snapshot,
            "user_input": user_input,
            "fixer_prompt": fixer_prompt,
            "exec_namespace": exec_namespace,
            "max_retries": max_retries,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": (
                    f"{snapshot}\n\n用户描述: {user_input}\n\n"
                    "理解全局结构后输出Python代码。只输出代码。"
                )},
            ],
            "code": "",
            "error": "",
            "search_hint": "",
            "attempt": 0,
            "success": False,
            "result": "",
            "all_codes": [],
        }

        # For tracking whether fixer gave us corrected code
        fixer_code = None

        # ── Main loop ──
        while state["attempt"] < state["max_retries"]:
            # Check cancellation
            if cancel_check and cancel_check():
                state["result"] = "用户取消"
                break

            # ── Node: generate_code ──
            self._node_generate(state, fixer_code, progress_callback, code_callback)
            fixer_code = None  # consumed

            # ── Node: execute ──
            exec_success = self._node_execute(state)

            if exec_success:
                state["success"] = True
                suffix = f" (第{state['attempt']+1}次)" if state["attempt"] > 0 else ""
                state["result"] = f"执行成功{suffix}"
                break

            # ── Execution failed, try to fix ──
            if state["attempt"] < state["max_retries"] - 1:
                # ── Node: analyze_error + web_search ──
                self._node_analyze_error(state)

                # ── Node: fix_code ──
                fixer_code = self._node_fix(state, progress_callback)

                if fixer_code:
                    # Fixer gave us corrected code → feed to LLM as context
                    state["messages"].append({"role": "assistant", "content": state["code"]})
                    state["messages"].append({"role": "user", "content": (
                        f"代码报错:\n{state['error']}\n\n专家已修正，直接执行。"
                    )})
                    if progress_callback:
                        progress_callback("专家已修正，直接执行...", "#9C27B0")
                else:
                    # No fixer → ask LLM to retry
                    state["messages"].append({"role": "assistant", "content": state["code"]})
                    feedback = f"代码报错:\n{state['error']}\n\n"
                    if state["search_hint"]:
                        feedback += f"{state['search_hint']}\n\n"
                    feedback += "修正后输出完整代码。只输出代码。"
                    state["messages"].append({"role": "user", "content": feedback})
                    if progress_callback:
                        progress_callback(f"修复第{state['attempt']+2}轮...", "#f0a030")

            state["attempt"] += 1

        # ── Handle final failure ──
        if not state["success"]:
            state["result"] = f"失败({state['max_retries']}轮)"

        return {
            "success": state["success"],
            "code": state["code"],
            "result": state["result"],
            "attempt": state["attempt"] + 1,
            "all_codes": state["all_codes"],
        }

    # ── Node implementations ────────────────────────────────

    def _node_generate(self, state: PipelineState, fixer_code: Optional[str],
                       progress_cb, code_cb):
        """Generate code: use fixer's code if available, else call LLM."""
        if fixer_code:
            state["code"] = fixer_code
            if code_cb:
                code_cb(f"--- 专家修正版 ---\n{fixer_code}")
        else:
            tag = "生成..." if state["attempt"] == 0 else f"修复第{state['attempt']}轮..."
            if progress_cb:
                progress_cb(tag, "#2196F3")

            response = self.client.chat(state["messages"])
            state["code"] = strip_code_fence(response)

            if code_cb:
                code_cb(f"--- {state['attempt']+1}/{state['max_retries']} ---\n{state['code']}")

        state["all_codes"].append(state["code"])

    def _node_execute(self, state: PipelineState) -> bool:
        """Execute generated code in the tool's namespace. Returns True on success."""
        import traceback

        try:
            exec(state["code"], state["exec_namespace"])
            return True
        except Exception:
            state["error"] = self._explain_error(
                traceback.format_exc(), state["code"]
            )
            return False

    def _node_analyze_error(self, state: PipelineState):
        """Analyze error + web search for solutions."""
        from t2cad_fixer import explain_error, web_search

        state["error"] = explain_error(state["error"], state["code"])
        state["search_hint"] = web_search(state["error"], state["code"][:500]) or ""

        # search_hint already populated by _node_analyze_error
        pass

    def _node_fix(self, state: PipelineState, progress_cb) -> Optional[str]:
        """Call fixer agent. Returns corrected code or None."""
        from t2cad_fixer import fix_code

        if progress_cb:
            progress_cb("专家诊断中...", "#9C27B0")

        return fix_code(
            state["fixer_prompt"],
            state["error"],
            state["code"],
            state["search_hint"],
            self.client.cfg,
            self.client.proxies,
        )

    # ── Error analysis ──────────────────────────────────────

    def _explain_error(self, trace: str, code: str) -> str:
        """Show which AI code line failed + decode COM HRESULT."""
        import re

        lines = trace.strip().split("\n")
        result = []

        # Find the failing code line
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

        # Decode COM HRESULT
        hresults = {
            "-2147352565": "[DISP_E_MEMBERNOTFOUND] 调用了不存在的方法/属性",
            "-2147352567": "[COM失败] 操作被拒绝 — 参数无效/权限不足/对象状态异常",
            "-2146827284": "[文件未找到]",
            "-2146822347": "[集合成员不存在] 行列越界或合并单元格",
            "-2146822297": "[无法访问单独行] 有纵向合并单元格",
        }
        for code_str, msg in hresults.items():
            if code_str in trace:
                result.append(msg)

        # Last error line
        for i in range(len(lines) - 1, -1, -1):
            if "Error" in lines[i] or "error" in lines[i].lower():
                result.append(lines[i].strip())
                break
        else:
            if lines:
                result.append(lines[-1].strip())

        return "\n".join(result) if result else trace.strip()


# ---------------------------------------------------------------------------
# Convenience: get a ready-to-use pipeline
# ---------------------------------------------------------------------------
_default_pipeline: Optional[CodeGenPipeline] = None


def get_pipeline() -> CodeGenPipeline:
    global _default_pipeline
    if _default_pipeline is None:
        _default_pipeline = CodeGenPipeline()
    return _default_pipeline
