# TextToCAD v3 - Excel Edition (v2.1 - global view)
"""Standalone app: Chinese NL -> LLM -> Excel COM -> spreadsheet/report/chart."""
import os, sys, json, re, traceback, time, subprocess
from pathlib import Path

LOG_FILE = os.path.join(os.path.expanduser("~"), "t2cad_excel_debug.log")

def _log(msg):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except:
        pass

MISSING = []
try:
    from PySide6 import QtWidgets, QtCore, QtGui
except ImportError:
    try:
        from PySide2 import QtWidgets, QtCore, QtGui
    except ImportError:
        MISSING.append("PySide6")

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

CONFIG_DIR = Path.home() / ".text_to_cad"
CONFIG_FILE = CONFIG_DIR / "config.json"
BRIDGE_DIR = CONFIG_DIR / "bridge"
BRIDGE_INPUT = BRIDGE_DIR / "input.txt"
BRIDGE_OUTPUT = BRIDGE_DIR / "output.py"
BRIDGE_DONE = BRIDGE_DIR / "done.txt"

# COM HRESULT for Excel cell errors (e.g. #DIV/0! = 0x800A07D7)
XL_HRESULT_ERRORS = {
    -2146826281: "#DIV/0!",
    -2146826273: "#VALUE!",
    -2146826265: "#REF!",
    -2146826259: "#NAME?",
    -2146826252: "#NUM!",
    -2146826246: "#N/A",
    -2146826250: "#NULL!",
}

DEFAULT_CONFIG = {
    "provider": "deepseek",
    "api_key": "",
    "api_base": "https://api.deepseek.com/v1",
    "model": "deepseek-chat",
    "temperature": 0.0,
    "max_tokens": 4096,
    "proxies": {"enabled": True, "http": "", "https": ""},
    "language": "zh",
    "system_prompt_zh": (
        "你是 Excel 专家。每次操作都考虑格式美化。只输出 Python 代码。\n\n"
        "## ⚠ 运行环境（必读！违反会导致代码失败）\n"
        "代码在 exec() 中执行，excel/wb/ws/math 已预先绑定到当前打开的 Excel。\n"
        "你绝对不能：任何 import 语句、Dispatch、EnsureDispatch、Workbooks.Open\n"
        "你绝对不能：open()读取文件、print()输出、exit()退出、wb.Sheets.Add()\n"
        "你绝对不能：try/except 吞错误 — 代码报错会自动显示并重试，让错误暴露出来\n"
        "你绝对不能：引用文件路径(C:\\...\\.xlsx)、打开其他工作簿 — 所有数据已在当前 wb 中\n"
        "创建工作表必须用 NEW_SHEET(\"表名\")，不要用 wb.Sheets.Add()（会因重名报错）\n"
        "你只能：操作已有的 wb(工作簿) 和 ws(当前工作表)，用 ws.Range() 读写单元格。\n"
        "⚠ 快照中的跨表引用(含!)指向的是本工作簿的其他工作表，或已链接的外部文件。\n"
        "  外部文件的数据已通过公式链接到当前表，你不需要也不能去打开源文件。\n"
        "  直接读写本工作簿的工作表即可，不要理会公式中出现的文件路径。\n\n"
        "## ⚠ 表格美化铁律\n"
        "1. 创建数据表后必须：表头加粗居中+加边框+冻结首行+自动列宽\n"
        "2. 数值列右对齐，文字列左对齐，表头全部居中\n"
        "3. 表头底色用深蓝(0x4472C4)白字，或浅灰(0xD9E2F3)黑字\n"
        "4. 金额/百分比列设置对应数字格式(#,##0.00 或 0.0%)\n"
        "5. 用 BEAUTIFY(rng) 一键美化选区\n"
        "6. 用 MERGE(\"A1:E1\") 合并单元格做标题\n\n"
        "## 核心函数(已注入)\n"
        "V(cell)=取值  F(cell)=取公式  ERR(cell)=取错误名  FIX(cell,f)=写公式\n"
        "LOOP(ws,col,start,end)=遍历列\n"
        "BEAUTIFY(rng_str)=一键美化(加边框+表头加粗+自动列宽+首行冻结)\n"
        "MERGE(rng_str)=合并单元格  BORDER(rng_str)=加边框\n"
        "BG(rng_str,color)=设底色  FONT_SET(rng_str,name,size,bold)=设字体\n"
        "ALIGN(rng_str,align)=对齐(1左2中3右)  NUMFMT(rng_str,fmt)=数字格式\n"
        "FREEZE(r,c)=冻结窗格  AUTOFIT()=自动列宽  COND_FMT(rng_str,rule)=条件格式\n"
        "CHART(type,rng_str,x,y,w,h)=插入图表\n"
        "NEW_SHEET(name)=创建工作表(如已存在则复用并清空)\n"
        "GET_SHEET(name)=获取工作表(用索引遍历, 支持中文表名, 替代wb.Sheets)\n\n"
        "## 分析任务时的代码模板\n"
        "当用户要求分析数据时，遍历工作表、读取单元格值、将结果写入一个新工作表：\n"
        "report = NEW_SHEET(\"分析报告\")\n"
        "report.Range(\"A1\").Value = \"分析报告\"\n"
        "# 遍历各工作表，用 V(cell) 读值，统计后写入 report\n\n"
        "## 写值/格式示例\n"
        "BEAUTIFY(\"A1:F100\")  # 一键美化整个数据区域\n"
        "ws.Range(\"A1\").Value = \"标题\"\n"
        "MERGE(\"A1:F1\")  # 合并标题行\n"
        "FONT_SET(\"A1:F1\", \"黑体\", 16, True)  # 标题字体\n"
        "ws.Range(\"A2:F2\").Value = ((\"序号\",\"名称\",\"金额\",\"日期\",\"状态\",\"备注\"),)\n"
        "BORDER(\"A2:F100\")  # 加边框\n"
        "NUMFMT(\"C3:C100\", \"#,##0.00\")  # 金额格式\n"
        "FREEZE(2, 0)  # 冻结表头行\n"
        "AUTOFIT()\n"
        "ws = GET_SHEET(\"Sheet名\")  # 切换工作表（不用wb.Sheets，因为中文表名会失败）\n"
        "last_row = ws.UsedRange.Rows.Count  # 末行\n\n"
        "## 铁律: 禁止import 禁止Dispatch 禁止Workbooks.Open 用已有excel/wb/ws 只输出代码"
    ),

}

def load_config():
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    BRIDGE_DIR.mkdir(parents=True, exist_ok=True)
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        for k, v in DEFAULT_CONFIG.items():
            cfg.setdefault(k, v)
        return cfg
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_CONFIG, f, indent=2, ensure_ascii=False)
    return dict(DEFAULT_CONFIG)

def _resolve_proxies(proxies_cfg):
    proxies = proxies_cfg or {"enabled": True}
    if not isinstance(proxies, dict) or not proxies.get("enabled"):
        return None
    http_proxy = proxies.get("http") or proxies.get("https")
    if not http_proxy:
        try:
            from urllib.request import getproxies
            sys_proxy = getproxies()
            http_proxy = sys_proxy.get("https") or sys_proxy.get("http")
        except:
            pass
    if http_proxy:
        return {"http": http_proxy, "https": http_proxy}
    return None

def _strip_code_fence(text):
    text = re.sub(r'^```(?:python)?\s*\n?', '', text.strip())
    text = re.sub(r'\n?```\s*$', '', text)
    return text.strip()

def _extract_short_error(exc_info):
    """Simple fallback: last error line + context."""
    lines = exc_info.strip().split("\n")
    for i in range(len(lines) - 1, -1, -1):
        if "Error" in lines[i] or "error" in lines[i]:
            return "\n".join(lines[max(0, i - 1):i + 1]).strip()
    return "\n".join(lines[-3:]).strip()

def _explain_error(trace, code):
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
                # Show context: 2 lines before
                ctx_start = max(0, ln - 3)
                ctx_end = min(len(code_lines), ln + 1)
                for cl in range(ctx_start, ctx_end):
                    marker = "  >" if cl == ln - 1 else "   "
                    result.append(f"{marker} L{cl+1}: {code_lines[cl].strip()[:120]}")
                break

    # 2. Decode COM HRESULT
    if "-2147352565" in trace:
        result.append("[DISP_E_MEMBERNOTFOUND] 调用了不存在的方法/属性 — 对象类型可能不对")
    if "-2147352567" in trace:
        result.append("[COM失败] 操作被Excel拒绝 — 可能原因: 参数无效/权限不足/对象状态异常")
    if "-2146827284" in trace:
        result.append("[文件未找到]")

    # 3. Last error line
    last_error = ""
    for i in range(len(lines) - 1, -1, -1):
        if "Error" in lines[i] or "error" in lines[i].lower():
            last_error = lines[i].strip()
            break
    if last_error:
        result.append(last_error)
    elif lines:
        result.append(lines[-1].strip())

    return "\n".join(result) if result else trace.strip()

def _web_search(error_text, code_snippet=""):
    """Search DuckDuckGo for Excel COM/VBA solutions related to the error.
    Returns search result snippets as a string, or None if unavailable."""
    try:
        from ddgs import DDGS
    except ImportError:
        try:
            from duckduckgo_search import DDGS
        except ImportError:
            return None
    # Build search query from error + code context
    query_parts = ["Excel VBA", "COM", "win32com"]
    # Extract key error info
    if "NameError" in error_text:
        query_parts.append("NameError")
    if "com_error" in error_text:
        query_parts.append("com_error")
        query_parts.append("pywintypes")
    if "UsedRange" in error_text:
        query_parts.append("UsedRange")
    if "Workbooks" in error_text:
        query_parts.append("Workbooks")
    # Add short error summary
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

FIXER_SYSTEM_PROMPT = """\
你是 Windows COM / win32com / Excel VBA 底层调试专家。

## 你的任务
用户提供了一段在 exec() 中执行的 Python 代码和报错信息。你需要：
1. 准确判断错误原因（不是泛泛而谈，要指出具体哪行、哪个API用错了）
2. 提供修正后的完整代码

## win32com 常见陷阱（务必记住！）
- 不能用 .Worksheets("名字") — win32com 的 IDispatch 不支持中文参数，必须用 GET_SHEET("名字")
- 不能用 wb.Sheets("中文名") — 同上，中文名传不进 COM，必须用 GET_SHEET()
- 不能用 wb.Sheets.Add() — 改用 NEW_SHEET()
- 不能用 import / Dispatch / EnsureDispatch / Workbooks.Open
- 不能用 print() / try/except pass
- 不能对 MERGE 后的单元格设置 ALIGN — MERGE 后合并区域的部分属性只读
- .HorizontalAlignment 用数字: 1=xlLeft(-4131), 2=xlCenter(-4108), 3=xlRight(-4152)
- 已注入的函数: V(), F(), ERR(), FIX(), LOOP(), BEAUTIFY(), MERGE(), BORDER(), BG(),
  FONT_SET(), ALIGN(), NUMFMT(), FREEZE(), AUTOFIT(), CHART(), NEW_SHEET(), GET_SHEET()
- excel/wb/ws 已连接到当前工作簿，不要重新打开或创建

## 输出格式
只输出修正后的完整 Python 代码。不要解释，不要 markdown 标记。"""

def _call_claude_cli(error_text, failed_code, search_hint):
    """Use local Claude Code CLI as COM expert fixer. Returns corrected code or None."""
    prompt = FIXER_SYSTEM_PROMPT + "\n\n## 报错信息\n" + error_text
    prompt += "\n\n## 失败的代码（第X行出错）\n" + failed_code
    if search_hint:
        prompt += "\n\n## 网上查到的资料\n" + search_hint
    prompt += "\n\n只输出修正后的完整Python代码，不要解释，不要markdown标记。"
    try:
        # On Windows, claude is a .cmd wrapper; use shell=True so the OS resolves it
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

def _call_fixer(error_text, failed_code, search_hint, cfg, proxies):
    """Call expert to fix code errors. Tries Claude CLI first, then API."""
    # 1. Try local Claude Code CLI (no API key needed)
    fixed = _call_claude_cli(error_text, failed_code, search_hint)
    if fixed:
        return fixed

    # 2. Fall back to API-based fixer
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
            {"role": "system", "content": FIXER_SYSTEM_PROMPT},
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

class ExcelConnection:
    def __init__(self):
        self.excel = None
        self.connected = False
        self.error = ""

    def connect(self, create=False):
        import win32com.client
        try:
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = True
            _ = self.excel.Name
            self.connected = True
            self.error = ""
            return True
        except Exception as e:
            self.connected = False
            self.error = str(e)
            return False

    def ensure(self):
        if self.connected and self.excel:
            try:
                _ = self.excel.Name
                return
            except:
                self.connected = False
        if not self.connect(create=True):
            raise ConnectionError(self.error)

    @property
    def wb(self):
        wb = self.excel.ActiveWorkbook
        if wb is None:
            wb = self.excel.Workbooks.Add()
        return wb

    @property
    def ws(self):
        ws = self.excel.ActiveSheet
        if ws is None:
            wb = self.wb
            if wb.Sheets.Count == 0:
                wb.Sheets.Add()
            ws = wb.ActiveSheet
        return ws

def _snapshot(conn):
    """Read full workbook overview: all sheets, named ranges, cross-refs, errors."""
    try:
        wb = conn.wb
        out = [f"工作簿: {wb.Name} ({wb.Sheets.Count}个工作表)"]

        # Named ranges
        try:
            names = wb.Names
            if names.Count:
                nr = []
                for i in range(1, min(names.Count, 50) + 1):
                    try:
                        nm = names(i)
                        nr.append(f"  {nm.Name} -> {nm.RefersTo}")
                    except:
                        pass
                if nr:
                    out.append(f"\n[名称管理器] ({len(nr)}个):\n" + "\n".join(nr))
        except:
            pass

        # All sheets
        ws_active = conn.ws
        for si in range(1, wb.Sheets.Count + 1):
            try:
                ws = wb.Sheets(si)
                used = ws.UsedRange
                if used is None:
                    out.append(f"\n工作表 {si}: {ws.Name} - 空")
                    continue
                rows, cols = used.Rows.Count, used.Columns.Count
                if rows == 1 and cols == 1 and (used.Value is None or str(used.Value).strip() == ""):
                    out.append(f"\n工作表 {si}: {ws.Name} - 空")
                    continue

                max_r, max_c = min(rows, 25), min(cols, 12)

                # Bulk read (2 COM calls per sheet)
                sub = ws.Range(ws.Cells(1, 1), ws.Cells(max_r, max_c))
                vals = sub.Value
                forms = sub.Formula

                # Normalize to 2D
                if not isinstance(vals, tuple):
                    vals = ((vals,),)
                if vals and not isinstance(vals[0], tuple):
                    vals = (vals,)
                if not isinstance(forms, tuple):
                    forms = ((forms,),)
                if forms and not isinstance(forms[0], tuple):
                    forms = (forms,)

                errors, formulas, refs_other = [], [], []

                for r in range(max_r):
                    for c in range(max_c):
                        v = vals[r][c] if r < len(vals) and c < len(vals[r]) else None
                        f = forms[r][c] if r < len(forms) and c < len(forms[r]) else None
                        ref = f"{chr(65+c)}{r+1}"

                        if isinstance(v, int) and v in XL_HRESULT_ERRORS:
                            errors.append(f"    {ref}: {XL_HRESULT_ERRORS[v]} <- {str(f)[:60]}")
                        if f and str(f).startswith("="):
                            formulas.append(f"    {ref}: {str(f)[:70]}")
                            if "!" in str(f):
                                if "[" in str(f) or "\\" in str(f):
                                    refs_other.append(f"    {ref}: {str(f)[:70]} ← ⚠外部文件(不要打开)")
                                else:
                                    refs_other.append(f"    {ref}: {str(f)[:70]} ← 跨工作表")

                is_active = " [当前]" if ws_active is not None and ws.Name == ws_active.Name else ""
                out.append(f"\n{'='*50}")
                out.append(f"工作表{si}: {ws.Name}{is_active}  ({rows}行 x {cols}列)")

                if errors:
                    out.append(f"  [{len(errors)}个错误]")
                    out.extend(errors)
                if refs_other:
                    ext_count = sum(1 for r in refs_other if "外部文件" in r)
                    xsheet_count = len(refs_other) - ext_count
                    label = f"  [{xsheet_count}个跨表, {ext_count}个外部链接(不要打开源文件)]"
                    out.append(label)
                    out.extend(refs_other)
                if formulas:
                    out.append(f"  [{len(formulas)}个公式]")
                    out.extend(formulas[:15])
                    if len(formulas) > 15:
                        out.append(f"    ... 余{len(formulas)-15}个")

                # Data sample
                out.append(f"  数据采样:")
                for r in range(min(4, max_r)):
                    row_items = []
                    for c in range(max_c):
                        v = vals[r][c] if r < len(vals) and c < len(vals[r]) else None
                        if isinstance(v, int) and v in XL_HRESULT_ERRORS:
                            row_items.append(f"[{XL_HRESULT_ERRORS[v]}]")
                        elif isinstance(v, float):
                            row_items.append(str(round(v, 2)))
                        elif v is not None:
                            row_items.append(str(v)[:15])
                        else:
                            row_items.append("")
                    out.append(f"  R{r+1}: " + " | ".join(f"{x:>14}" for x in row_items))

            except Exception as e:
                out.append(f"\n工作表{si}: 读取失败 ({e})")

        return "\n".join(out)
    except Exception as e:
        return f"快照失败: {e}"

class TextToCADApp(QtWidgets.QMainWindow):
    def __init__(self, conn):
        super().__init__()
        self.cfg = load_config()
        self.proxies = _resolve_proxies(self.cfg.get("proxies")) if HAS_REQUESTS else None
        self.conn = conn
        self._cancelled = False

        self.setWindowTitle("TextToCAD for Excel v2")
        self.setMinimumSize(420, 480)
        self.resize(540, 600)

        menu_bar = self.menuBar()
        fm = menu_bar.addMenu("文件")
        fm.addAction("检查连接").triggered.connect(self._check)
        fm.addAction("设置...").triggered.connect(lambda: os.startfile(str(CONFIG_FILE)))
        fm.addSeparator()
        fm.addAction("退出").triggered.connect(self.close)
        hm = menu_bar.addMenu("帮助")
        hm.addAction("关于").triggered.connect(lambda: QtWidgets.QMessageBox.about(
            self, "关于", "TextToCAD for Excel v2\nNL->Excel automation"))

        cw = QtWidgets.QWidget()
        self.setCentralWidget(cw)
        layout = QtWidgets.QVBoxLayout(cw)
        layout.setContentsMargins(10, 10, 10, 10)

        title = QtWidgets.QLabel("用自然语言操作 Excel  -- 输入问题或操作描述")
        title.setStyleSheet("font-size: 14px; font-weight: bold; padding: 4px;")
        layout.addWidget(title)

        self.input_edit = QtWidgets.QTextEdit()
        self.input_edit.setPlaceholderText(
            "例如:\n"
            "  创建销售表 A列产品 B列金额 C列日期 加柱状图\n"
            "  修复表格中的所有公式错误\n"
            "  在D列计算B列*1.2 保留两位小数\n"
            "  分析这份数据有什么规律\n"
            "  按照金额列降序排列"
        )
        self.input_edit.setMaximumHeight(110)
        layout.addWidget(self.input_edit)

        btn_layout = QtWidgets.QHBoxLayout()
        bs = "QPushButton{color:white;font-size:13px;font-weight:bold;padding:8px 20px;border-radius:4px}QPushButton:disabled{background:#666}"

        self.run_btn = QtWidgets.QPushButton("执行")
        self.run_btn.setStyleSheet(bs + "QPushButton{background:#2196F3}QPushButton:hover{background:#1976D2}")
        self.run_btn.clicked.connect(self.on_run)
        btn_layout.addWidget(self.run_btn)

        self.cancel_btn = QtWidgets.QPushButton("取消")
        self.cancel_btn.setStyleSheet(bs + "QPushButton{background:#f44336}QPushButton:hover{background:#d32f2f}")
        self.cancel_btn.clicked.connect(self.on_cancel)
        self.cancel_btn.hide()
        btn_layout.addWidget(self.cancel_btn)

        self.query_btn = QtWidgets.QPushButton("分析")
        self.query_btn.setStyleSheet(bs + "QPushButton{background:#4CAF50}QPushButton:hover{background:#388E3C}")
        self.query_btn.clicked.connect(self.on_query)
        btn_layout.addWidget(self.query_btn)

        self.reconn_btn = QtWidgets.QPushButton("重连")
        self.reconn_btn.setStyleSheet("QPushButton{font-size:13px;padding:8px 20px;border-radius:4px}QPushButton:hover{background:#e0e0e0}")
        self.reconn_btn.clicked.connect(self._reconnect)
        btn_layout.addWidget(self.reconn_btn)
        layout.addLayout(btn_layout)

        # ── 快捷工具栏 ──
        tools_layout = QtWidgets.QHBoxLayout()
        tools_label = QtWidgets.QLabel("快捷:")
        tools_label.setStyleSheet("font-size: 11px; color: #888; padding-right: 4px;")
        tools_layout.addWidget(tools_label)

        self.beautify_btn = QtWidgets.QPushButton("一键美化")
        self.beautify_btn.setToolTip("当前表: 加边框+表头加粗蓝底白字+自动列宽+冻结首行")
        self.beautify_btn.setStyleSheet(
            "QPushButton{background:#FF9800;color:white;font-size:11px;font-weight:bold;padding:5px 10px;border-radius:3px}"
            "QPushButton:hover{background:#F57C00}")
        self.beautify_btn.clicked.connect(self.on_beautify)
        tools_layout.addWidget(self.beautify_btn)

        self.auto_btn = QtWidgets.QPushButton("自动列宽")
        self.auto_btn.setToolTip("当前表所有列自动调整宽度")
        self.auto_btn.setStyleSheet(
            "QPushButton{background:#607D8B;color:white;font-size:11px;font-weight:bold;padding:5px 10px;border-radius:3px}"
            "QPushButton:hover{background:#455A64}")
        self.auto_btn.clicked.connect(self.on_autofit)
        tools_layout.addWidget(self.auto_btn)

        self.freeze_btn = QtWidgets.QPushButton("冻结首行")
        self.freeze_btn.setToolTip("冻结当前表第1行（表头固定）")
        self.freeze_btn.setStyleSheet(
            "QPushButton{background:#00897B;color:white;font-size:11px;font-weight:bold;padding:5px 10px;border-radius:3px}"
            "QPushButton:hover{background:#00695C}")
        self.freeze_btn.clicked.connect(self.on_freeze)
        tools_layout.addWidget(self.freeze_btn)

        tools_layout.addStretch()
        layout.addLayout(tools_layout)

        self.status_lbl = QtWidgets.QLabel("就绪")
        self.status_lbl.setStyleSheet("color: #888; font-size: 11px; padding: 4px;")
        self.status_lbl.setWordWrap(True)
        layout.addWidget(self.status_lbl)

        self.output_edit = QtWidgets.QTextEdit()
        self.output_edit.setReadOnly(True)
        self.output_edit.setMaximumHeight(200)
        self.output_edit.setStyleSheet(
            "QTextEdit{background:#1e1e1e;color:#d4d4d4;font-family:Consolas;font-size:11px}")
        layout.addWidget(self.output_edit)

        proxy_info = " proxy" if self.proxies else ""
        self.mode_lbl = QtWidgets.QLabel(f"{self.cfg['provider']}/{self.cfg['model']}{proxy_info}")
        self.mode_lbl.setStyleSheet("color: #f0a030; font-size: 10px;")
        layout.addWidget(self.mode_lbl)

        self.excel_status = QtWidgets.QLabel()
        self._update_status()
        self.statusBar().addPermanentWidget(self.excel_status)
        QtGui.QShortcut(QtGui.QKeySequence("Ctrl+Return"), self).activated.connect(self.on_run)

    def _update_status(self):
        if self.conn.connected:
            self.excel_status.setText("Excel: 已连接")
            self.excel_status.setStyleSheet("color: #4caf50; font-weight: bold; padding: 2px 8px;")
            self.reconn_btn.hide()
        else:
            self.excel_status.setText("Excel: 未连接")
            self.excel_status.setStyleSheet("color: #f0a030; font-weight: bold; padding: 2px 8px;")
            self.reconn_btn.show()

    def _reconnect(self):
        self.set_status("连接...", "#2196F3")
        ok = self.conn.connect(create=True)
        self._update_status()
        self.set_status("已连接" if ok else f"失败: {self.conn.error}", "green" if ok else "red")

    def _check(self):
        try:
            self.conn.ensure()
            QtWidgets.QMessageBox.information(self, "OK",
                f"Excel: {self.conn.excel.Name}\nWorkbook: {self.conn.wb.Name}\nSheet: {self.conn.ws.Name}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", str(e))
        self._update_status()

    # ── 快捷工具 ──
    def on_beautify(self):
        """一键美化当前表格"""
        try:
            self.conn.ensure(); self._update_status()
        except ConnectionError as e:
            QtWidgets.QMessageBox.warning(self, "Excel未连接", str(e)); return
        try:
            ws = self.conn.ws
            if ws is None:
                self.set_status("无活动工作表", "red"); return
            used = ws.UsedRange
            if used is None:
                self.set_status("表格为空", "red"); return
            rng = ws.Range(used.Address)
            rng.Borders.LineStyle = 1
            first = ws.Range(used.Rows(1).Address)
            first.Font.Bold = True
            first.Interior.Color = 0x4472C4
            first.Font.Color = 0xFFFFFF
            first.HorizontalAlignment = -4108
            ws.Columns.AutoFit()
            # Freeze top row
            try:
                ws.Range(ws.Cells(2, 1)).Select()
                self.conn.excel.ActiveWindow.FreezePanes = True
            except:
                pass
            self.set_status("美化完成: 边框+表头+列宽+冻结", "green")
        except Exception as e:
            self.set_status(f"失败: {e}", "red")

    def on_autofit(self):
        try: self.conn.ensure(); self.conn.ws.Columns.AutoFit(); self.set_status("自动列宽完成", "green")
        except Exception as e: self.set_status(f"失败: {e}", "red")

    def on_freeze(self):
        try:
            self.conn.ensure()
            ws = self.conn.ws
            ws.Range(ws.Cells(2, 1)).Select()
            self.conn.excel.ActiveWindow.FreezePanes = True
            self.set_status("已冻结首行", "green")
        except Exception as e: self.set_status(f"失败: {e}", "red")

    def _call_llm(self, messages):
        if self.cfg["provider"] == "bridge":
            return self._call_bridge(messages)
        headers = {"Content-Type": "application/json",
                   "Authorization": f"Bearer {self.cfg['api_key']}"}
        body = {"model": self.cfg["model"], "messages": messages,
                "temperature": self.cfg["temperature"], "max_tokens": self.cfg["max_tokens"]}
        url = f"{self.cfg['api_base'].rstrip('/')}/chat/completions"
        resp = requests.post(url, headers=headers, json=body, timeout=60, proxies=self.proxies)
        resp.raise_for_status()
        return _strip_code_fence(resp.json()["choices"][0]["message"]["content"])

    def _call_bridge(self, messages):
        BRIDGE_DIR.mkdir(parents=True, exist_ok=True)
        with open(BRIDGE_INPUT, "w", encoding="utf-8") as f:
            f.write(messages[-1]["content"])
        if BRIDGE_DONE.exists():
            BRIDGE_DONE.unlink()
        waited = 0
        while waited < 120:
            if self._cancelled:
                raise InterruptedError("取消")
            if BRIDGE_OUTPUT.exists() and BRIDGE_DONE.exists():
                code = _strip_code_fence(open(BRIDGE_OUTPUT, "r", encoding="utf-8").read())
                BRIDGE_INPUT.unlink(missing_ok=True)
                BRIDGE_OUTPUT.unlink(missing_ok=True)
                BRIDGE_DONE.unlink(missing_ok=True)
                return code
            time.sleep(0.5)
            waited += 0.5
            QtCore.QCoreApplication.processEvents()
        raise TimeoutError("桥接超时")

    def on_cancel(self):
        self._cancelled = True
        self.set_status("取消中...", "#f44336")

    def on_query(self):
        question = self.input_edit.toPlainText().strip()
        try:
            self.conn.ensure()
            self._update_status()
        except ConnectionError as e:
            QtWidgets.QMessageBox.warning(self, "Excel未连接", str(e))
            return

        t0 = time.time()
        snap = _snapshot(self.conn)
        t1 = time.time()
        self.output_edit.setText(f"读取 {(t1-t0)*1000:.0f}ms\n{snap}")
        self.set_status("AI分析中...", "#2196F3")
        QtCore.QCoreApplication.processEvents()

        self.query_btn.setEnabled(False)
        self.run_btn.setEnabled(False)
        try:
            user_msg = f"工作簿快照:\n{snap}\n\n用户问题: {question}" if question else f"工作簿快照:\n{snap}\n\n请总结数据概况和规律。"
            messages = [
                {"role": "system", "content": "你是Excel数据分析专家。根据工作簿快照回答用户问题。关注所有工作表、名称管理器、跨表引用。直接给结论。"},
                {"role": "user", "content": user_msg},
            ]
            answer = self._call_llm(messages)
            t2 = time.time()
            label = "查询" if question else "分析"
            self.output_edit.setText(f"--- {label} ({(t2-t1)*1000:.0f}ms) ---\n{answer}")
            self.set_status("完成", "green")
        except Exception as e:
            self.set_status(f"失败: {e}", "red")
        finally:
            self.query_btn.setEnabled(True)
            self.run_btn.setEnabled(True)

    def on_run(self):
        user_input = self.input_edit.toPlainText().strip()
        if not user_input:
            self.set_status("请输入描述", "red")
            return

        try:
            self.conn.ensure()
            self._update_status()
        except ConnectionError as e:
            QtWidgets.QMessageBox.warning(self, "Excel未连接", str(e))
            return

        self.run_btn.setEnabled(False)
        self.cancel_btn.show()
        self.output_edit.clear()
        self._cancelled = False

        snap = _snapshot(self.conn)

        messages = [
            {"role": "system", "content": self.cfg["system_prompt_zh"]},
            {"role": "user", "content": f"工作簿快照:\n{snap}\n\n用户描述: {user_input}\n\n理解全局结构后，输出Python代码。只输出代码。"},
        ]

        try:
            MAX_RETRIES = 6
            fixer_code = None  # set by fixer agent to override next LLM call

            for attempt in range(MAX_RETRIES):
                if self._cancelled:
                    self.set_status("已取消", "#f0a030")
                    return

                tag = "生成..." if attempt == 0 else f"修复第{attempt}轮..."
                self.set_status(tag, "#2196F3")
                QtCore.QCoreApplication.processEvents()

                try:
                    if fixer_code:
                        code = fixer_code
                        fixer_code = None
                        self.output_edit.setText(f"--- 专家修正版 ---\n{code}")
                    else:
                        code = self._call_llm(messages)
                        self.output_edit.setText(f"--- {attempt+1}/{MAX_RETRIES} ---\n{code}")

                    # ── sanitize common LLM artifacts ──
                    code = re.sub(r'^[ \t]*\^+\s*$', '', code, flags=re.MULTILINE)  # ^^^ lines
                    code = re.sub(r'^[ \t]*print\s*\([^)]*\)\s*$', '', code, flags=re.MULTILINE)  # print lines
                    code = re.sub(r'\b\.Worksheets\(', '.Sheets(', code)  # win32com only supports .Sheets
                    # Chinese sheet names fail via COM dispatch — replace with safe index-based lookup
                    code = re.sub(r'\b(excel|wb)\.Sheets\("([^"]*)"\)', r'GET_SHEET("\2")', code)

                    # ── pre-exec validation: reject dangerous patterns ──
                    forbidden = [
                        (r'\bimport\s+\w+', '禁止 import（excel/wb/ws 已注入）'),
                        (r'\bfrom\s+\w+\s+import\b', '禁止 from...import（excel/wb/ws 已注入）'),
                        (r'\bEnsureDispatch\b', '禁止 EnsureDispatch（excel 已连接）'),
                        (r'\bWorkbooks\.Open\b', '禁止 Workbooks.Open（wb 已打开）'),
                        (r'\bexcel\.Workbooks\.Open\b', '禁止打开文件（wb 已打开）'),
                        (r'\bwb\.Sheets\.Add\b', '禁止 wb.Sheets.Add() — 请用 NEW_SHEET("表名")'),
                        (r'except\s*.*:\s*\n\s*pass', '禁止 try/except 吞错 — 报错会自动重试修复，直接写代码即可'),
                        (r'[A-Za-z]:\\[^\s\'"]*\.xlsx?\b', '禁止引用文件路径 — 数据已在当前 wb 中，用 wb.Sheets("表名") 即可'),
                        (r'请手动打开', '禁止让用户手动操作 — 数据已在当前 wb 中，直接用代码处理'),
                    ]
                    for pattern, msg in forbidden:
                        if re.search(pattern, code):
                            raise RuntimeError(f"[代码审查] {msg}\n请用已有的 excel/wb/ws 变量操作当前工作簿。")

                    # ── capture ws/wb/excel in enclosing scope for helper closures ──
                    ws = self.conn.ws
                    wb = self.conn.wb
                    excel = self.conn.excel

                    # Safe helpers so AI doesn't need to know COM quirks
                    def _v(cell):
                        try:
                            v = cell.Value
                            return v if v is not None else ""
                        except:
                            return ""

                    def _f(cell):
                        try:
                            f = cell.Formula
                            return str(f) if f is not None else ""
                        except:
                            return ""

                    def _err(cell):
                        try:
                            v = cell.Value
                            return XL_HRESULT_ERRORS.get(v, None) if isinstance(v, int) else None
                        except:
                            return None

                    def _fix(cell, new_formula):
                        """Replace formula, preserving format."""
                        try:
                            cell.Formula = new_formula
                            return True
                        except:
                            return False

                    def _loop(ws_obj, col, start, end):
                        """Iterate cells in range, yielding (row, cell)."""
                        for r in range(start, end + 1):
                            yield r, ws_obj.Range(f"{col}{r}")

                    # ── formatting helpers ──
                    def _beautify(rng_str):
                        """One-click beautify: borders, bold header, auto-fit, freeze."""
                        rng = ws.Range(rng_str)
                        rng.Borders.LineStyle = 1  # xlContinuous
                        # Bold & center first row (header)
                        first_row = ws.Range(rng.Rows(1).Address)
                        first_row.Font.Bold = True
                        first_row.Interior.Color = 0x4472C4
                        first_row.Font.Color = 0xFFFFFF
                        first_row.HorizontalAlignment = -4108  # xlCenter
                        ws.Rows(1).AutoFit()

                    def _merge(rng_str):
                        ws.Range(rng_str).Merge()

                    def _border(rng_str):
                        ws.Range(rng_str).Borders.LineStyle = 1

                    def _bg(rng_str, color):
                        ws.Range(rng_str).Interior.Color = color

                    def _font_set(rng_str, name, size, bold=False):
                        rng = ws.Range(rng_str)
                        rng.Font.Name = name
                        rng.Font.Size = size
                        rng.Font.Bold = bold

                    def _align(rng_str, align=1):
                        """1=xlLeft, 2=xlCenter, 3=xlRight"""
                        ws.Range(rng_str).HorizontalAlignment = -4131 - align

                    def _numfmt(rng_str, fmt):
                        ws.Range(rng_str).NumberFormat = fmt

                    def _freeze(r, c):
                        """Freeze panes at row r, col c (0=no freeze in that dim)."""
                        if r and c:
                            ws.Activate()
                            ws.Range(ws.Cells(r + 1, c + 1)).Select()
                            excel.ActiveWindow.FreezePanes = True
                        elif r:
                            ws.Activate()
                            ws.Range(ws.Cells(r + 1, 1)).Select()
                            excel.ActiveWindow.FreezePanes = True
                        elif c:
                            ws.Activate()
                            ws.Range(ws.Cells(1, c + 1)).Select()
                            excel.ActiveWindow.FreezePanes = True

                    def _autofit():
                        ws.Columns.AutoFit()

                    def _chart(chart_type, rng_str, x=100, y=100, w=400, h=300):
                        """Insert chart. type: 'bar'/'pie'/'line'/'column'"""
                        types = {'bar': 57, 'pie': 5, 'line': 4, 'column': 51}
                        co = ws.ChartObjects().Add(x, y, w, h)
                        co.Chart.SetSourceData(ws.Range(rng_str))
                        co.Chart.ChartType = types.get(chart_type, 51)
                        return co

                    def _get_sheet(name):
                        """Get sheet by name, safe with Chinese chars (iter by index)."""
                        for i in range(1, wb.Sheets.Count + 1):
                            s = wb.Sheets(i)
                            if s.Name == name:
                                return s
                        raise ValueError(f"找不到工作表: {name}")

                    def _new_sheet(name):
                        """Get or create sheet: if exists, clear & reuse; else add new."""
                        for s in wb.Sheets:
                            if s.Name == name:
                                s.Activate()
                                s.Cells.Clear()
                                return s
                        s = wb.Sheets.Add()
                        s.Name = name
                        return s

                    ns = {"excel": self.conn.excel, "wb": self.conn.wb,
                           "ws": self.conn.ws, "math": __import__("math"),
                           "V": _v, "F": _f, "ERR": _err, "FIX": _fix,
                           "LOOP": _loop, "ECODES": XL_HRESULT_ERRORS,
                           "BEAUTIFY": _beautify, "MERGE": _merge,
                           "BORDER": _border, "BG": _bg, "FONT_SET": _font_set,
                           "ALIGN": _align, "NUMFMT": _numfmt, "FREEZE": _freeze,
                           "AUTOFIT": _autofit, "CHART": _chart, "NEW_SHEET": _new_sheet,
                           "GET_SHEET": _get_sheet}
                    exec(code, ns)
                    suffix = "" if attempt == 0 else f" (第{attempt+1}轮修复成功)"
                    self.set_status(f"完成{suffix}", "green")
                    return
                except Exception as e:
                    trace = traceback.format_exc()
                    short = _explain_error(trace, code)
                    _log(f"Exec err {attempt+1}: {short}")

                    if attempt < MAX_RETRIES - 1:
                        self.output_edit.setText(f"--- 错误, AI修复第{attempt+1}轮 ---\n{code}\n\n{short}")

                        # ── web search: look up error solutions ──
                        search_hint = _web_search(short, code[:500])
                        if search_hint:
                            self.output_edit.setText(
                                f"--- 错误, 搜索+专家诊断中... ---\n{code}\n\n{short}\n\n{search_hint[:300]}...")

                        # ── COM expert fixer agent ──
                        self.set_status(f"COM专家诊断中...", "#9C27B0")
                        QtCore.QCoreApplication.processEvents()
                        fixed = _call_fixer(short, code, search_hint, self.cfg, self.proxies)
                        if fixed:
                            _log(f"Fixer returned {len(fixed)} chars, will exec directly")
                            fixer_code = fixed  # Use fixer's code directly next round
                            # Also feed to main LLM as context for future rounds
                            messages.append({"role": "assistant", "content": code})
                            hint = f"代码报错:\n{short}\n\nCOM专家给出了修正版，已直接执行。如果还报错，请参考修正版重新生成。"
                            if search_hint:
                                hint = search_hint + "\n\n" + hint
                            messages.append({"role": "user", "content": hint})
                            self.set_status(f"COM专家已修正，直接执行...", "#9C27B0")
                        else:
                            # Fallback: no fixer response, use standard retry
                            messages.append({"role": "assistant", "content": code})
                            feedback = f"代码报错:\n{short}\n\n"
                            if search_hint:
                                feedback += f"{search_hint}\n\n"
                            feedback += "修正后输出完整代码。只输出代码。"
                            messages.append({"role": "user", "content": feedback})
                            self.set_status(f"修复第{attempt+2}轮...", "#f0a030")
                        QtCore.QCoreApplication.processEvents()
                    else:
                        self.output_edit.setText(f"--- {MAX_RETRIES}轮后仍失败 ---\n{code}\n\n{short}")
                        self.set_status(f"失败({MAX_RETRIES}轮)", "red")
        except (InterruptedError, TimeoutError):
            pass
        except Exception as e:
            self.set_status(f"错误: {e}", "red")
            self.output_edit.setText(traceback.format_exc())
            _log(f"API: {traceback.format_exc()}")
        finally:
            self.run_btn.setEnabled(True)
            self.cancel_btn.hide()

    def set_status(self, text, color="#888"):
        self.status_lbl.setText(text)
        self.status_lbl.setStyleSheet(f"color: {color}; font-size: 11px; padding: 4px;")

def run_app():
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    conn = ExcelConnection()
    conn.connect(create=False) or conn.connect(create=True)
    window = TextToCADApp(conn)
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    if MISSING:
        import ctypes
        msg = "缺少依赖:\n\n" + "\n".join(f"  - {d}" for d in MISSING)
        ctypes.windll.user32.MessageBoxW(0, msg, "TextToCAD - 依赖缺失", 0x10)
        sys.exit(1)
    run_app()
