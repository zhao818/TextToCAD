# TextToCAD v3 — Word Edition
"""Standalone app: Chinese NL → LLM → Word COM → document/report/table."""
import os, sys, json, re, traceback, subprocess
from pathlib import Path

LOG_FILE = os.path.join(os.path.expanduser("~"), "t2cad_word_debug.log")

CONFIG_DIR = Path.home() / ".text_to_cad"
sys.path.insert(0, str(CONFIG_DIR))
from t2cad_fixer import explain_error, web_search, fix_code
from t2cad_llm import LLMClient, strip_code_fence
from t2cad_pipeline import CodeGenPipeline

def _log(msg):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except:
        pass

# ── dependency check ─────────────────────────────────────
MISSING = []
try:
    from PySide6 import QtWidgets, QtCore, QtGui
except ImportError:
    try:
        from PySide2 import QtWidgets, QtCore, QtGui
    except ImportError:
        MISSING.append("PySide6 (pip install pyside6)")

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

# ── paths ────────────────────────────────────────────────
CONFIG_DIR = Path.home() / ".text_to_cad"
CONFIG_FILE = CONFIG_DIR / "config.json"
BRIDGE_DIR = CONFIG_DIR / "bridge"
BRIDGE_INPUT = BRIDGE_DIR / "input.txt"
BRIDGE_OUTPUT = BRIDGE_DIR / "output.py"
BRIDGE_DONE = BRIDGE_DIR / "done.txt"

WORD_FIXER_PROMPT = """\
你是 Word COM / win32com 底层调试专家。

## 任务
用户代码在 exec() 中执行报错。判断原因并输出修正后的完整 Python 代码。

## 环境说明
- 已注入: word=Word应用 doc=文档 sel=选择对象
- word/doc/sel 已连接，不要重新创建或打开文档
- 禁止: import / Dispatch / EnsureDispatch / Documents.Open

## Word COM 核心知识

── 表格操作 ──
- **访问单元格**: t.Range.Cells 遍历所有实际单元格（安全，支持合并单元格）
  t.Cell(r,c) 按行列访问（有纵向合并时抛异常或产生空行，尽量不用）
- **合并相邻表格**: doc.Range(t1.Range.End, t2.Range.Start).Delete() 删掉中间的¶
  两个表自动合并（需列数相同或兼容）
- **表格自适应**: t.AutoFitBehavior(2) 适配窗口; t.AutoFitBehavior(1) 适配内容
- **列宽**: t.Columns(c).Width 单位磅; t.Columns(c).PreferredWidthType=2 设百分比
- **单元格边距**: cell.TopPadding/BottomPadding/LeftPadding/RightPadding 单位磅
- **合并/拆分**: cell.Merge(other_cell) 合并; cell.Split(num_rows, num_cols) 拆分
- **跨页控制**: t.Rows.AllowBreakAcrossPages=False 禁止行内断页
  rng.ParagraphFormat.KeepWithNext=-1 与下一段保持同页
- **表头重复**: t.Rows(1).HeadingFormat=-1 每页重复表头
- **表格排序**: t.Sort(ExcludeHeader=True, FieldNumber=1, SortFieldType=0, SortOrder=0)
- **行列增删**: t.Rows.Add()/t.Columns.Add() 追加; t.Rows(i).Delete() 删除
- **表格样式**: t.Style = "网格型" 等内置样式名
- **单元格垂直对齐**: cell.VerticalAlignment(0顶1中2底)

── 段落与文字 ──
- **段落格式**: pf=sel.ParagraphFormat; pf.Alignment(0左1中2右3两端4分散)
  pf.LineSpacingRule(5=多倍行距) pf.LineSpacing(倍×12)
  pf.SpaceBefore/SpaceAfter 段间距(pt) pf.CharacterUnitFirstLineIndent 缩进(字符数)
  pf.FirstLineIndent 首行缩进(磅) pf.LeftIndent/RightIndent 左右缩进(磅)
- **字体**: sel.Font.Name/Size/Bold/Italic/Color/Underline
  Color=(B<<16)|(G<<8)|R 如蓝色=(0<<16)|(0<<8)|255=255
- **样式**: doc.Styles("Heading 1"); sel.Style = style_obj
  内置样式名: "Heading 1"~"Heading 9", "Normal", "Title", "TOC 1"
- **范围**: doc.Range(start, end) 创建; rng.Text 读写; rng.Delete() 删除
  rng.Start/End 字符位置; rng.Copy()/rng.Paste() 复制粘贴
  rng.Paragraphs.Count rng.Sentences.Count rng.Words.Count 统计
- **查找替换**: sel.Find.Execute(FindText, ReplaceWith, Replace=2, MatchCase=False)
  Replace: 0不替 1替1个 2全替; 支持通配符 MatchWildcards=True

── 页面布局 ──
- **页边距**: 1cm=28.346pt; ps=doc.PageSetup
  ps.TopMargin/BottomMargin/LeftMargin/RightMargin/HeaderDistance/FooterDistance
- **方向**: ps.Orientation(0竖1横); **纸张**: ps.PageWidth/PageHeight 磅
  ps.PaperSize=9 即A4; ps.PaperSize=7 即A3
- **分页**: sel.InsertBreak(Type=7) 分页符; Type=2 分节符(下一页)
  Type=3 分节符(连续) Type=4 分节符(偶数页)
- **节**: doc.Sections(i) 节集合; sec.PageSetup 节级页面设置
  不同节可有不同页边距/方向/页码
- **页数统计**: doc.Repaginate() 强制重分页; doc.ComputeStatistics(2) 总页数
  doc.ComputeStatistics(0) 总字数; (3) 总行数; (4) 总段落数

── 页眉页脚 ──
- **访问页眉**: sec.Headers(1) 首页页眉; sec.Headers(2) 偶数页; sec.Headers(3) 主页眉
- **页脚**: sec.Footers(1/2/3) 同上
- **页眉内容**: header.Range.Text = "xxx" ; header.Range.ParagraphFormat.Alignment=1
- **页码**: doc.Fields.Add(header.Range, -1, "PAGE") 插入页码域
  doc.Fields.Add(header.Range, -1, "NUMPAGES") 总页数域

── 常用 Information 常量 ──
- rng.Information(1)=字数 (2)=字符数 (3)=当前页号 (4)=总页数
  (5)=当前节号 (6)=当前列号 (7)=是否在页眉 (8)=是否在页脚
  (9)=行号 (10)=当前大纲级别 (11)=是否选中 (12)=是否在表格内
  (13)=是否在列表 (14)=是否在文本框 (15)=是否在脚注/尾注

── 常见 COM HRESULT ──
- -2146822347 (0x800A1735): 集合成员不存在（行列越界或合并单元格）
- -2146822297 (0x800A1757): 无法访问单独行（有纵向合并单元格）
- -2146827284 (0x800A040C): 文件未找到
- -2147352565 (0x8002000B): DISP_E_MEMBERNOTFOUND 方法名拼错或对象类型错
- -2146822342 (0x800A173A): 范围不能插入（表格外操作单元格时）
- -2146822496 (0x800A16A0): 命令不可用（文档保护或只读模式）

── 文档保护与属性 ──
- **只读检查**: doc.ProtectionType(-1=无保护 0=修订 1=批注 2=窗体 3=只读)
- **文档属性**: doc.BuiltInDocumentProperties("Title").Value
- **保存**: doc.Save()/doc.SaveAs2(filename); doc.Close(SaveChanges=0不保存1保存)
- **新建文档**: word.Documents.Add(Template) 新建空白文档

## 已注入的安全函数
WRITE(text) PARA() BOLD/SIZE/FONT/CENTER/LEFT/RIGHT/JUSTIFY
HEADING(n,text) TITLE(text) SUBTITLE(text)
TABLE(r,c) CELL(t,r,c,text) TABLE_BORDERS(t) TBOLD(t,row)
COL_WIDTHS(t,[w1,...]) CELL_MERGE CELL_ALIGN CELL_BG CELL_FONT
TABLE_HEADER(t,row,[texts])
PAGE_MARGINS PAGE_LANDSCAPE/PAGE_PORTRAIT PAGE_BREAK SECTION_BREAK
LINE_SPACING PARA_BEFORE/PARA_AFTER FIRST_LINE_INDENT
IMAGE FIND GOTO_END GOTO_START UNDO NEW_PAGE
COLOR(r,g,b)

只输出修正后的完整 Python 代码。"""

# ── config ───────────────────────────────────────────────
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
        "你是 Word 排版专家。你生成的每个文档都必须格式精美、符合中文正式文档标准。\n\n"
        "## ⚠ 排版铁律（每次生成文档都必须遵守）\n"
        "1. 创建文档第一件事：PAGE_MARGINS(top=2.54, bottom=2.54, left=3.17, right=3.17) 设置页边距\n"
        "2. 中文正式文档字体规范：标题用黑体，正文用宋体，英文/数字用Times New Roman\n"
        "3. 字号层级：大标题 22pt(二号) / 章节标题 16pt(三号) / 小标题 14pt(四号) / 正文 12pt(小四) 或 10.5pt(五号)\n"
        "4. 正文行距必须 1.5 倍（用 LINE_SPACING(1.5)），表格内可用 1.15 倍\n"
        "5. 正文段落首行缩进 2 字符（用 FIRST_LINE_INDENT(2)）\n"
        "6. 创建表格后立即用 TABLE_BORDERS(t) 加全边框，表头行用 TBOLD(t,1) 加粗\n"
        "7. 表格列宽用 COL_WIDTHS(t, [...]) 设为合理比例，不要全等宽\n"
        "8. 段落之间有适当间距（用 PARA_BEFORE/PARA_AFTER），标题段前 12pt，正文段前 0-6pt\n"
        "9. 页面方向：宽表用 PAGE_LANDSCAPE() 横版，普通文档用 PAGE_PORTRAIT() 竖版\n"
        "10. 颜色用RGB: COLOR(255,0,0)=蓝, COLOR(0,0,255)=红, COLOR(0,128,0)=绿\n\n"
        "## 所有可用函数\n"
        "── 页面设置 ──\n"
        "  PAGE_MARGINS(top, bottom, left, right) — 页边距(cm)，默认2.54\n"
        "  PAGE_LANDSCAPE() / PAGE_PORTRAIT() — 横/竖版\n"
        "  PAGE_BREAK() — 分页符  SECTION_BREAK() — 分节符\n"
        "── 文字与段落 ──\n"
        "  WRITE(text) PARA() — 写字、换行\n"
        "  BOLD(on=True) SIZE(n) FONT(name) — 加粗、字号(pt)、字体\n"
        "  CENTER() LEFT() RIGHT() JUSTIFY() — 对齐方式\n"
        "  COLOR(r,g,b) — 文字颜色(0-255)\n"
        "  LINE_SPACING(mult) — 行距倍数(1.0/1.25/1.5/2.0)\n"
        "  PARA_BEFORE(n) PARA_AFTER(n) — 段前/段后间距(pt)\n"
        "  FIRST_LINE_INDENT(n) — 首行缩进n字符\n"
        "── 标题与样式 ──\n"
        "  HEADING(n, text) — Word内置标题(1-3)\n"
        "  TITLE(text) — 文档大标题(自动黑体22pt居中)\n"
        "  SUBTITLE(text) — 副标题\n"
        "── 表格 ──\n"
        "  TABLE(rows, cols) — 创建表格\n"
        "  CELL(t, r, c, text) — 填单元格(r/c从1开始)\n"
        "  TABLE_BORDERS(t) — 给表格加全边框\n"
        "  TBOLD(t, row) — 表头行加粗\n"
        "  COL_WIDTHS(t, [w1,w2,...]) — 设列宽(cm)\n"
        "  CELL_MERGE(t, r1,c1, r2,c2) — 合并单元格\n"
        "  CELL_ALIGN(t, r, c, align) — 单元格对齐(0左1中2右)\n"
        "  CELL_BG(t, r, c, rgb) — 单元格底色，如(200,220,255)\n"
        "  CELL_FONT(t, r, c, name, size, bold) — 单元格字体\n"
        "── 其他 ──\n"
        "  IMAGE(path, w, h) FIND(old,new) GOTO_END() GOTO_START() UNDO()\n"
        "  doc=文档对象 word=Word应用 sel=选择对象\n\n"
        "## 完整示例：生成专业报告\n"
        "PAGE_MARGINS(2.54, 2.54, 3.17, 3.17)\n"
        "LINE_SPACING(1.5)\n"
        "TITLE(\"2026年度云县低效林改造项目 施工组织设计方案\")\n"
        "PARA_BEFORE(12); PARA_AFTER(12)\n"
        "PARA()\n"
        "FONT(\"宋体\"); SIZE(12); FIRST_LINE_INDENT(2)\n"
        "WRITE(\"一、项目概况。本项目位于云南省临沧市云县...\")\n"
        "PARA()\n"
        "WRITE(\"二、施工准备。施工单位应做好以下准备工作...\")\n"
        "PARA()\n"
        "t = TABLE(5, 4)\n"
        "TABLE_BORDERS(t)\n"
        "COL_WIDTHS(t, [1.5, 4, 3, 3])\n"
        "CELL(t,1,1,\"序号\"); CELL(t,1,2,\"工序名称\"); CELL(t,1,3,\"质量标准\"); CELL(t,1,4,\"备注\")\n"
        "TBOLD(t, 1)\n"
        "CELL(t,2,1,\"1\"); CELL(t,2,2,\"林地清理\"); CELL(t,2,3,\"全面清理\")\n"
        "CELL(t,3,1,\"2\"); CELL(t,3,2,\"整地挖穴\"); CELL(t,3,3,\"40×40×30cm\")\n"
        "GOTO_END()\n"
        "PARA()\n"
        "COLOR(128,128,128); SIZE(9)\n"
        "WRITE(\"编制单位：XXX  |  编制日期：2026年X月X日\")\n\n"
        "## 铁律\n"
        "禁止import 禁止Dispatch 禁止Client 禁止word.Visible\n"
        "word/doc/sel已注入直接用 表格行/列从1开始\n"
        "只输出Python代码 不要任何解释文字 不要markdown标记\n\n"
        "## ⚠ 表格操作进阶（必读！）\n"
        "- 合并两个相邻表格：删掉表格之间的空段落即可自动合并\n"
        "  doc.Range(t1.Range.End, t2.Range.Start).Delete()\n"
        "- 遍历表格所有单元格用 t.Range.Cells，不要用 t.Cell(r,c) 双重循环（合并单元格会出错）\n"
        "- 设置表头重复（每页显示）：t.Rows(1).HeadingFormat = -1\n"
        "- 表格自适应：t.AutoFitBehavior(2) 适配窗口宽度\n"
        "- 单元格对齐用 CELL_ALIGN(t,r,c,align): 0左 1中 2右\n"
        "- 修改已有表格：直接操作 wb.Tables(i) 获取，不要新建\n"
        "- 表格跨页控制：t.Rows.AllowBreakAcrossPages = False 禁止行内断页\n\n"
        "## ⚠ 文档操作进阶\n"
        "- 段落对齐：CENTER()=居中 LEFT()=左对齐 RIGHT()=右对齐 JUSTIFY()=两端对齐\n"
        "- 修改已有段落的格式：p = doc.Paragraphs(i); p.Range.Font.Size = 14\n"
        "- 分页符：PAGE_BREAK() 会插入到当前光标位置\n"
        "- 获取页数：doc.Repaginate(); pages = doc.ComputeStatistics(2)\n"
        "- 查找替换：FIND(\"旧文字\", \"新文字\") 替换全文\n"
        "- 插入目录前确保已有标题样式（HEADING函数会自动设置样式）"
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

# ── helpers ───────────────────────────────────────────────
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
    lines = exc_info.strip().split("\n")
    last_file = None
    for i, line in enumerate(lines):
        if line.strip().startswith("File "):
            last_file = i
    if last_file is not None:
        return "\n".join(lines[last_file:]).strip()
    return "\n".join(lines[-3:]).strip()

# ── Word connection ───────────────────────────────────────
class WordConnection:
    def __init__(self):
        self.word = None
        self.connected = False
        self.error = ""

    def connect(self, create=False):
        import win32com.client
        try:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = True
            _ = self.word.Name
            self.connected = True
            self.error = ""
            return True
        except Exception as e:
            self.connected = False
            self.error = str(e)
            return False

    def ensure(self):
        if self.connected and self.word:
            try:
                _ = self.word.Name
                return
            except:
                self.connected = False
        if not self.connect(create=True):
            raise ConnectionError(f"无法连接 Word\n{self.error}")

    @property
    def doc(self):
        doc = self.word.ActiveDocument
        if doc is None:
            doc = self.word.Documents.Add()
        return doc

    @property
    def sel(self):
        return self.word.Selection

# ── UI ────────────────────────────────────────────────────
class TextToCADApp(QtWidgets.QMainWindow):
    def __init__(self, conn):
        super().__init__()
        self.cfg = load_config()
        self.proxies = _resolve_proxies(self.cfg.get("proxies")) if HAS_REQUESTS else None
        self.conn = conn
        self._cancelled = False
        self.client = LLMClient(self.cfg)
        self.pipeline = CodeGenPipeline(self.client)

        self.setWindowTitle("TextToCAD for Word v2")
        self.setMinimumSize(400, 480)
        self.resize(580, 650)

        # ── menu ──
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("文件")
        act_check = file_menu.addAction("检查 Word 连接")
        act_check.triggered.connect(self._check_word)
        act_settings = file_menu.addAction("设置...")
        act_settings.triggered.connect(lambda: os.startfile(str(CONFIG_FILE)))
        file_menu.addSeparator()
        act_quit = file_menu.addAction("退出")
        act_quit.triggered.connect(self.close)
        help_menu = menu_bar.addMenu("帮助")
        act_about = help_menu.addAction("关于")
        act_about.triggered.connect(self._about)

        # ── central widget ──
        cw = QtWidgets.QWidget()
        self.setCentralWidget(cw)
        layout = QtWidgets.QVBoxLayout(cw)
        layout.setContentsMargins(10, 10, 10, 10)

        title = QtWidgets.QLabel("TextToCAD for Word — 用自然语言操作 Word + 一键美化")
        title.setStyleSheet("font-size: 14px; font-weight: bold; padding: 4px;")
        layout.addWidget(title)

        self.input_edit = QtWidgets.QTextEdit()
        self.input_edit.setPlaceholderText(
            "例如：\n"
            "  插入标题「2024年度报告」，居中，微软雅黑，20号字\n"
            "  创建5行3列的表格，填入示例数据，加边框\n"
            "  正文设置为宋体12号，行距1.5倍\n"
            "  把文章中的「AI」全部替换为「人工智能」\n"
            "  在文档末尾插入图片 C:\\logo.png，居中对齐\n"
            "  总结当前文档说了什么"
        )
        self.input_edit.setMaximumHeight(140)
        layout.addWidget(self.input_edit)

        # buttons
        btn_layout = QtWidgets.QHBoxLayout()
        self.run_btn = QtWidgets.QPushButton("执行")
        self.run_btn.setStyleSheet(
            "QPushButton { background-color: #2196F3; color: white; font-size: 13px;"
            " font-weight: bold; padding: 8px 20px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #1976D2; }"
            "QPushButton:disabled { background-color: #666; }"
        )
        self.run_btn.clicked.connect(self.on_run)
        btn_layout.addWidget(self.run_btn)

        self.cancel_btn = QtWidgets.QPushButton("取消")
        self.cancel_btn.setStyleSheet(
            "QPushButton { background-color: #f44336; color: white; font-size: 13px;"
            " font-weight: bold; padding: 8px 20px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #d32f2f; }"
        )
        self.cancel_btn.clicked.connect(self.on_cancel)
        self.cancel_btn.hide()
        btn_layout.addWidget(self.cancel_btn)

        self.query_btn = QtWidgets.QPushButton("分析文档")
        self.query_btn.setStyleSheet(
            "QPushButton { background-color: #4CAF50; color: white; font-size: 13px;"
            " font-weight: bold; padding: 8px 20px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #388E3C; }"
            "QPushButton:disabled { background-color: #666; }"
        )
        self.query_btn.clicked.connect(self.on_query)
        btn_layout.addWidget(self.query_btn)

        self.reconnect_btn = QtWidgets.QPushButton("重连Word")
        self.reconnect_btn.setStyleSheet(
            "QPushButton { font-size: 13px; padding: 8px 20px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #e0e0e0; }"
        )
        self.reconnect_btn.clicked.connect(self._reconnect)
        btn_layout.addWidget(self.reconnect_btn)
        layout.addLayout(btn_layout)

        # ── 快捷工具栏（一键美化 / 诊断 / 模板）──
        tools_layout = QtWidgets.QHBoxLayout()
        tools_label = QtWidgets.QLabel("快捷工具:")
        tools_label.setStyleSheet("font-size: 11px; color: #888; padding-right: 4px;")
        tools_layout.addWidget(tools_label)

        self.beautify_btn = QtWidgets.QPushButton(" 一键美化 ")
        self.beautify_btn.setToolTip("对整个文档套用标准排版：页边距/字体/行距/表格边框/首行缩进")
        self.beautify_btn.setStyleSheet(
            "QPushButton { background-color: #FF9800; color: white; font-size: 12px;"
            " font-weight: bold; padding: 6px 14px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #F57C00; }"
        )
        self.beautify_btn.clicked.connect(self.on_beautify)
        tools_layout.addWidget(self.beautify_btn)

        self.diag_btn = QtWidgets.QPushButton(" 文档诊断 ")
        self.diag_btn.setToolTip("扫描当前文档的排版问题：字体不一致/缺少边框/行距问题等")
        self.diag_btn.setStyleSheet(
            "QPushButton { background-color: #9C27B0; color: white; font-size: 12px;"
            " font-weight: bold; padding: 6px 14px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #7B1FA2; }"
        )
        self.diag_btn.clicked.connect(self.on_diagnose)
        tools_layout.addWidget(self.diag_btn)

        self.clean_btn = QtWidgets.QPushButton(" 清理空行 ")
        self.clean_btn.setToolTip("删除文档中多余的空段落")
        self.clean_btn.setStyleSheet(
            "QPushButton { background-color: #607D8B; color: white; font-size: 12px;"
            " font-weight: bold; padding: 6px 14px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #455A64; }"
        )
        self.clean_btn.clicked.connect(self.on_clean_empty)
        tools_layout.addWidget(self.clean_btn)

        self.toc_btn = QtWidgets.QPushButton(" 插入目录 ")
        self.toc_btn.setToolTip("在光标位置插入自动目录（需先有Heading标题）")
        self.toc_btn.setStyleSheet(
            "QPushButton { background-color: #00897B; color: white; font-size: 12px;"
            " font-weight: bold; padding: 6px 14px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #00695C; }"
        )
        self.toc_btn.clicked.connect(self.on_insert_toc)
        tools_layout.addWidget(self.toc_btn)

        tools_layout.addStretch()
        layout.addLayout(tools_layout)

        # status
        self.status_lbl = QtWidgets.QLabel("就绪 — Ctrl+Enter 快速执行")
        self.status_lbl.setStyleSheet("color: #888; font-size: 11px; padding: 4px;")
        self.status_lbl.setWordWrap(True)
        layout.addWidget(self.status_lbl)

        # code output
        self.output_edit = QtWidgets.QTextEdit()
        self.output_edit.setReadOnly(True)
        self.output_edit.setMaximumHeight(200)
        self.output_edit.setStyleSheet(
            "QTextEdit { background-color: #1e1e1e; color: #d4d4d4;"
            " font-family: Consolas, monospace; font-size: 11px; }"
        )
        layout.addWidget(self.output_edit)

        # mode bar
        proxy_info = " (代理已检测)" if self.proxies else ""
        mode_text = f"模式: {self.cfg['provider']} | 模型: {self.cfg['model']}{proxy_info}"
        self.mode_lbl = QtWidgets.QLabel(mode_text)
        self.mode_lbl.setStyleSheet("color: #f0a030; font-size: 10px;")
        self.mode_lbl.setWordWrap(True)
        layout.addWidget(self.mode_lbl)

        # ── status bar ──
        self.word_status = QtWidgets.QLabel()
        self._update_status()
        self.statusBar().addPermanentWidget(self.word_status)

        # shortcuts
        QtGui.QShortcut(QtGui.QKeySequence("Ctrl+Return"), self).activated.connect(self.on_run)

    # ── status ────────────────────────────────────────
    def _update_status(self):
        if self.conn.connected:
            self.word_status.setText("Word:  已连接")
            self.word_status.setStyleSheet("color: #4caf50; font-weight: bold; padding: 2px 8px;")
            self.reconnect_btn.hide()
        else:
            self.word_status.setText("Word:  未连接")
            self.word_status.setStyleSheet("color: #f0a030; font-weight: bold; padding: 2px 8px;")
            self.reconnect_btn.show()

    def _reconnect(self):
        self.set_status("正在连接 Word...", "#2196F3")
        QtCore.QCoreApplication.processEvents()
        ok = self.conn.connect(create=True)
        self._update_status()
        if ok:
            self.set_status("Word 已连接", "green")
        else:
            self.set_status(f"连接失败: {self.conn.error}", "red")

    def _check_word(self):
        try:
            self.conn.ensure()
            QtWidgets.QMessageBox.information(self, "Word 连接",
                f"已连接\n文档: {self.conn.doc.Name}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Word 连接", str(e))
        self._update_status()

    def _about(self):
        QtWidgets.QMessageBox.about(self, "关于 TextToCAD",
            "TextToCAD for Word v2\n\n用自然语言操作 Word\n"
            "LLM: DeepSeek / OpenAI 兼容 API\n\n"
            "快捷工具：一键美化 | 文档诊断 | 清理空行 | 插入目录")

    # ── 智能美化（AI先分析结构，再精准套格式）───────────────
    def on_beautify(self):
        """AI-powered smart beautify: 先理解文档结构，再针对性美化"""
        try:
            self.conn.ensure()
            self._update_status()
        except ConnectionError as e:
            QtWidgets.QMessageBox.warning(self, "Word 未连接", str(e)[:200])
            return

        self.set_status("AI 分析文档结构中...", "#2196F3")
        QtCore.QCoreApplication.processEvents()

        try:
            wd = self.conn.word
            doc = self.conn.doc

            # ── 第1步：读取文档快照供AI分析 ──
            paras_info = []
            for i in range(1, min(doc.Paragraphs.Count + 1, 200)):
                try:
                    p = doc.Paragraphs(i)
                    rng = p.Range
                    txt = rng.Text.strip()
                    if not txt:
                        paras_info.append({"idx": i, "text": "", "font": "", "size": 0, "bold": False, "align": 0, "style": ""})
                        continue
                    paras_info.append({
                        "idx": i,
                        "text": txt[:100],
                        "font": str(rng.Font.Name or ""),
                        "size": rng.Font.Size if rng.Font.Size else 0,
                        "bold": bool(rng.Font.Bold),
                        "align": p.Format.Alignment if p.Format else 0,
                        "style": str(rng.Style.NameLocal or ""),
                    })
                except:
                    pass

            tables_info = []
            for ti in range(1, doc.Tables.Count + 1):
                try:
                    t = doc.Tables(ti)
                    rows, cols = t.Rows.Count, t.Columns.Count
                    # Read first 2 rows + last row as sample
                    sample = []
                    for ri in [1, 2, rows]:
                        row_data = []
                        for ci in range(1, min(cols, 6) + 1):
                            try:
                                row_data.append(str(t.Cell(ri, ci).Range.Text)[:50].strip())
                            except:
                                row_data.append("")
                        sample.append(f"  R{ri}: " + " | ".join(row_data))
                    tables_info.append({
                        "id": ti, "rows": rows, "cols": cols,
                        "sample": "\n".join(sample),
                    })
                except:
                    pass

            # ── 第2步：AI分析结构 ──
            snapshot = f"文档: {doc.Name} ({doc.Paragraphs.Count}段, {doc.Tables.Count}个表格)\n\n"
            snapshot += "段落列表:\n"
            for p in paras_info:
                if p["text"]:
                    snapshot += f"P{p['idx']}: font={p['font']} size={p['size']} bold={p['bold']} align={p['align']} style={p['style']} | {p['text']}\n"
            if tables_info:
                snapshot += f"\n表格列表 ({len(tables_info)}个):\n"
                for t in tables_info:
                    snapshot += f"Table{t['id']}: {t['rows']}行×{t['cols']}列\n{t['sample']}\n\n"

            analysis_prompt = (
                "你是中文文档排版专家，精通GB/T 9704公文标准和WPS/Word排版规范。\n\n"
                "## 任务\n"
                "先判断文档类型和内容密度，再为每段/每表标注类型和建议格式。\n\n"
                "## 输出格式（严格按此格式）\n"
                "DOCTYPE: 文档类型(报告/日报/公文/通知/合同/论文/其他)\n"
                "DENSITY: compact(紧凑,内容多需省空间) / normal(正常) / spacious(宽松,内容少可放大)\n"
                "TITLE_MAX: 标题最大字号(pt) — compact=18-22 normal=28-38 spacious=38-44\n"
                "BODY_SIZE: 正文字号(pt) — compact=10.5 normal=12 spacious=14\n"
                "LINE_SPACING: 行距倍数 — compact=1.25 normal=1.5 spacious=1.5\n"
                "\nP1: 封面标题 | size=XX\nP2: 副标题 | size=XX\nP3-P8: 正文\nP9: 章标题 | size=XX\n"
                "Table1: 数据表,表头行1-2\nTable2: 纯数据表,无表头\n"
                "P10: 落款\n\n"
                "段落类型: 封面标题/副标题/章标题/节标题/正文/列表项/落款/签名行/日期/空段\n"
                "表格类型: 数据表(有表头加色)/清单表(有表头加色)/排版表(无底色)/纯数据表/表单(无底色) + 表头行x-y\n\n"
                "## 判断规则\n"
                "- 数据表=行列整齐的数值表格→加浅蓝表头底色\n"
                "- 清单表=条目列表→加浅蓝表头底色\n"
                "- 排版表=用表格排版的文字(无数据含义)→不加底色\n"
                "- 表单=填空表格→不加底色\n"
                "- 纯数据表=无表头的纯数据→不加底色\n"
                "- 日报/通知/公文→compact, 不要用大标题, 控制在一页内\n"
                "- 正式报告/论文→normal, 标题可略大\n"
                "- 宣传材料/封面→spacious, 标题要大\n"
                "- 看总段落数: >50段内容多→compact; 10-50段→normal; <10段→spacious\n"
                "- 看内容关键词: 出现'日报''通知''记录''汇报'→compact\n"
                "- 标题size根据内容多少决定: 内容越多标题越小\n"
                "- 表格第一行如和其他行格式不同→表头\n\n"
                f"文档快照:\n{snapshot}\n\n请先判断DOCTYPE/DENSITY/字号, 再标注每段每表:"
            )

            self.set_status("AI 分析中...", "#9C27B0")
            QtCore.QCoreApplication.processEvents()

            analysis = self.client.chat([
                {"role": "system", "content": "你是文档结构分析专家。只输出结构标签，不要解释。"},
                {"role": "user", "content": analysis_prompt},
            ])

            self.output_edit.setText(f"--- AI结构分析 ---\n{analysis}")
            self.set_status("正在应用美化...", "#FF9800")
            QtCore.QCoreApplication.processEvents()

            # ── 第3步：根据AI分析精准套用格式 ──
            CMPT = 28.3464567

            # Page setup
            ps = doc.PageSetup
            ps.TopMargin = 2.54 * CMPT
            ps.BottomMargin = 2.54 * CMPT
            ps.LeftMargin = 3.17 * CMPT
            ps.RightMargin = 3.17 * CMPT

            # Parse AI analysis
            p_types = {}  # {"3-8": "正文", "1": "封面标题|size=24", ...}
            t_types = {}  # {"1": "数据表,表头行1-2", ...}
            doc_meta = {"DOCTYPE": "通用文档", "DENSITY": "normal",
                        "TITLE_MAX": "22", "BODY_SIZE": "12", "LINE_SPACING": "1.5"}

            for line in analysis.strip().split("\n"):
                line = line.strip()
                # Document-level meta
                for key in ("DOCTYPE", "DENSITY", "TITLE_MAX", "BODY_SIZE", "LINE_SPACING"):
                    km = re.match(rf'{key}\s*[:：]\s*(.+)', line)
                    if km:
                        doc_meta[key] = km.group(1).strip()
                # Paragraph labels: P1: 封面标题 | size=24
                pm = re.match(r'P(\d+(?:-\d+)?)\s*[:：]\s*(.+)', line)
                if pm:
                    p_types[pm.group(1)] = pm.group(2).strip()
                # Table labels: Table1: 数据表,表头行1-2
                tm = re.match(r'Table(\d+)\s*[:：]\s*(.+)', line)
                if tm:
                    t_types[tm.group(1)] = tm.group(2).strip()

            # Extract fallback sizing from meta
            try:
                title_max = float(doc_meta.get("TITLE_MAX", "22"))
            except:
                title_max = 22
            try:
                body_size = float(doc_meta.get("BODY_SIZE", "12"))
            except:
                body_size = 12
            try:
                line_sp = float(doc_meta.get("LINE_SPACING", "1.5"))
            except:
                line_sp = 1.5

            def _parse_size(label, fallback):
                """Extract 'size=XX' from a label, or return fallback."""
                sm = re.search(r'size\s*=\s*(\d+(?:\.\d+)?)', label)
                return float(sm.group(1)) if sm else fallback

            def _is_type(pidx, type_kw):
                """Check if paragraph index matches a type keyword."""
                for key, label in p_types.items():
                    if '-' in key:
                        start, end = key.split('-')
                        if int(start) <= pidx <= int(end) and type_kw in label:
                            return True
                    elif int(key) == pidx and type_kw in label:
                        return True
                return False

            def _get_type(pidx):
                for key, label in p_types.items():
                    if '-' in key:
                        start, end = key.split('-')
                        if int(start) <= pidx <= int(end):
                            return label
                    elif int(key) == pidx:
                        return label
                return None

            # ── Apply formatting ──
            stats = {"标题": 0, "正文": 0, "落款": 0, "表格": 0, "空段": 0}

            for i in range(1, doc.Paragraphs.Count + 1):
                try:
                    p = doc.Paragraphs(i)
                    rng = p.Range
                    txt = rng.Text.strip()

                    # Empty paragraphs: minimize
                    if not txt:
                        pf = p.Format
                        pf.SpaceBefore = 0
                        pf.SpaceAfter = 0
                        pf.LineSpacingRule = 5
                        pf.LineSpacing = 2
                        stats["空段"] += 1
                        continue

                    ptype = _get_type(i) or ""
                    pf = p.Format

                    # ── AI-suggested sizes replace hardcoded values ──
                    if "封面标题" in ptype:
                        rng.Font.Name = '黑体'
                        rng.Font.Size = _parse_size(ptype, title_max)
                        rng.Font.Bold = True
                        pf.Alignment = 1  # center
                        pf.SpaceBefore = title_max * 1.5
                        pf.SpaceAfter = title_max * 0.3
                        stats["标题"] += 1

                    elif "副标题" in ptype:
                        rng.Font.Name = '宋体'
                        rng.Font.Size = _parse_size(ptype, title_max * 0.65)
                        pf.Alignment = 1
                        pf.SpaceBefore = 6
                        pf.SpaceAfter = title_max * 0.6
                        stats["标题"] += 1

                    elif "章标题" in ptype:
                        rng.Font.Name = '黑体'
                        rng.Font.Size = _parse_size(ptype, title_max * 0.7)
                        rng.Font.Bold = True
                        pf.Alignment = 0
                        pf.SpaceBefore = title_max * 0.5
                        pf.SpaceAfter = title_max * 0.3
                        pf.LineSpacingRule = 5
                        pf.LineSpacing = line_sp * 12
                        stats["标题"] += 1

                    elif "节标题" in ptype:
                        rng.Font.Name = '黑体'
                        rng.Font.Size = _parse_size(ptype, title_max * 0.55)
                        rng.Font.Bold = True
                        pf.Alignment = 0
                        pf.SpaceBefore = title_max * 0.3
                        pf.SpaceAfter = title_max * 0.15
                        pf.LineSpacingRule = 5
                        pf.LineSpacing = line_sp * 12
                        stats["标题"] += 1

                    elif "正文" in ptype or "列表项" in ptype:
                        rng.Font.Name = '宋体'
                        rng.Font.Size = _parse_size(ptype, body_size)
                        rng.Font.Bold = False
                        pf.Alignment = 3 if "列表" not in ptype else 0
                        pf.LineSpacingRule = 5
                        pf.LineSpacing = line_sp * 12
                        if pf.Alignment != 1:
                            try:
                                pf.CharacterUnitFirstLineIndent = 2
                            except:
                                pass
                        pf.SpaceBefore = 0
                        pf.SpaceAfter = 6
                        stats["正文"] += 1

                    elif "落款" in ptype or "签名行" in ptype or "日期" in ptype:
                        rng.Font.Name = '宋体'
                        rng.Font.Size = _parse_size(ptype, body_size)
                        pf.Alignment = 2  # right
                        pf.SpaceBefore = 6
                        pf.SpaceAfter = 3
                        try:
                            pf.CharacterUnitFirstLineIndent = 0
                        except:
                            pass
                        stats["落款"] += 1

                    else:
                        # Unknown type: gentle defaults using AI density
                        rng.Font.Name = rng.Font.Name or '宋体'
                        if not rng.Font.Size or rng.Font.Size > 72:
                            rng.Font.Size = body_size

                except:
                    continue

            # ── Intelligent table layout optimization ──
            page_width_pt = ps.PageWidth - ps.LeftMargin - ps.RightMargin  # usable width in pt

            for ti in range(1, doc.Tables.Count + 1):
                try:
                    t = doc.Tables(ti)
                    rows, cols = t.Rows.Count, t.Columns.Count

                    # ── Step 1: Tight cell margins (safe for merged cells) ──
                    try:
                        for cell in t.Range.Cells:
                            try:
                                cell.TopPadding = 2
                                cell.BottomPadding = 2
                                cell.LeftPadding = 5
                                cell.RightPadding = 5
                            except:
                                pass
                    except:
                        pass

                    # ── Step 2: AutoFit to window ──
                    try:
                        t.AutoFitBehavior(2)  # wdAutoFitWindow
                    except:
                        pass

                    # ── Step 3: Check if table overflows page ──
                    total_width = 0
                    for c in range(1, cols + 1):
                        try:
                            total_width += t.Columns(c).Width
                        except:
                            pass

                    # If table is wider than page, force proportional shrink
                    if total_width > page_width_pt * 1.05:
                        ratio = page_width_pt / total_width * 0.95
                        for c in range(1, cols + 1):
                            try:
                                t.Columns(c).Width = t.Columns(c).Width * ratio
                            except:
                                pass

                    # ── Step 4: Detect and fix table page spanning ──
                    # wdActiveEndPageNumber = 3
                    try:
                        first_page = t.Cell(1, 1).Range.Information(3)
                        last_page = t.Cell(rows, 1).Range.Information(3)
                        pages_used = (last_page or first_page or 2) - (first_page or 1) + 1
                    except:
                        pages_used = 1

                    # If table spans multiple pages, try progressive shrink to fit on fewer
                    if pages_used > 1:
                        for shrink_pt in [0.5, 1.0, 1.5]:
                            shrunk = False
                            try:
                                for cell in t.Range.Cells:
                                    try:
                                        cur = cell.Range.Font.Size
                                        if cur and cur > 8:
                                            cell.Range.Font.Size = max(8, cur - shrink_pt)
                                            shrunk = True
                                    except:
                                        pass
                            except:
                                pass

                            if not shrunk:
                                continue

                            try:
                                for cell in t.Range.Cells:
                                    try:
                                        cell.TopPadding = 1
                                        cell.BottomPadding = 1
                                    except:
                                        pass
                            except:
                                pass

                            try:
                                t.AutoFitBehavior(2)
                            except:
                                pass

                            # Re-check pages
                            try:
                                new_last = t.Cell(rows, 1).Range.Information(3)
                                if isinstance(new_last, int) and isinstance(first_page, int):
                                    if new_last - first_page + 1 < pages_used:
                                        pages_used = new_last - first_page + 1
                                        break  # improvement, keep shrinkage
                            except:
                                break

                    # ── Step 5: Prevent orphan rows ──
                    try:
                        t.Rows.AllowBreakAcrossPages = True
                    except:
                        pass

                    # ── Step 6: Style borders + header ──
                    t.Borders.Enable = True
                    t.Borders.InsideLineStyle = 1
                    t.Borders.OutsideLineStyle = 1

                    tlabel = t_types.get(str(ti), "")
                    header_match = re.search(r'表头行?(\d+)(?:-(\d+))?', tlabel)
                    if header_match:
                        h_start = int(header_match.group(1))
                        h_end = int(header_match.group(2)) if header_match.group(2) else h_start
                    else:
                        h_start, h_end = 1, 1

                    # Format header rows: bold + center. Only add background
                    # color if AI labeled it as a data table (not plain/form table)
                    use_color = "数据表" in tlabel or "清单表" in tlabel
                    try:
                        for cell in t.Range.Cells:
                            try:
                                ri = cell.RowIndex
                                if h_start <= ri <= h_end:
                                    cell.Range.Font.Bold = True
                                    cell.Range.ParagraphFormat.Alignment = 1
                                    if use_color:
                                        cell.Shading.BackgroundPatternColor = 0xD9E2F3
                            except:
                                pass
                    except:
                        pass

                    stats["表格"] += 1
                except:
                    continue

            # ── Step 7: Merge consecutive tables ──
            # In Word, deleting the ¶ between two tables merges them automatically
            if doc.Tables.Count >= 2:
                merged = 0
                # Iterate backwards to avoid index shift after merging
                for ti in range(doc.Tables.Count, 1, -1):
                    try:
                        t2 = doc.Tables(ti)       # lower table
                        t1 = doc.Tables(ti - 1)   # upper table

                        # Get the range BETWEEN the two tables
                        r_between = doc.Range(t1.Range.End, t2.Range.Start)

                        # Only merge if only empty paragraphs between them
                        txt_between = r_between.Text.strip()
                        if not txt_between:
                            # Tables have compatible column count? (optional check)
                            if t1.Columns.Count == t2.Columns.Count:
                                r_between.Delete()  # Delete the ¶ → tables merge!
                                merged += 1
                    except:
                        continue
                if merged > 0:
                    stats["表格合并"] = f"合并{merged}组相邻表格"

            # ── 第4步：页面拟合优化 ──
            self.set_status("页面拟合优化...", "#00897B")
            QtCore.QCoreApplication.processEvents()

            def _count_pages():
                """Force repaginate then count pages."""
                try:
                    doc.Repaginate()
                    return doc.ComputeStatistics(2)  # wdStatisticPages
                except:
                    return 1

            # 4a. Remove excessive blank paragraphs FIRST
            blank_removed = 0
            for i in range(doc.Paragraphs.Count, 0, -1):
                try:
                    p = doc.Paragraphs(i)
                    txt = p.Range.Text.strip()
                    if not txt:
                        # Check if previous 2 paragraphs are also empty → delete
                        prev_empty = True
                        for j in range(i - 1, max(0, i - 3), -1):
                            try:
                                if doc.Paragraphs(j).Range.Text.strip():
                                    prev_empty = False
                                    break
                            except:
                                pass
                        if prev_empty and i > 3:
                            p.Range.Delete()
                            blank_removed += 1
                except:
                    continue
            if blank_removed:
                stats["空白"] = f"删除{blank_removed}个多余空段"

            # 4b. Check pages, try to fit on fewer
            initial_pages = _count_pages()

            if initial_pages > 1:
                # Phase 1: Reduce margins (fastest, biggest impact)
                for margin_cm in [2.2, 2.0, 1.7, 1.5]:
                    ps.TopMargin = margin_cm * CMPT
                    ps.BottomMargin = margin_cm * CMPT
                    ps.LeftMargin = max(1.5, margin_cm) * CMPT
                    ps.RightMargin = max(1.5, margin_cm) * CMPT
                    pages = _count_pages()
                    if isinstance(pages, int) and pages < initial_pages:
                        initial_pages = pages
                        stats["页面优化"] = f"缩边距至{margin_cm:.1f}cm→{pages}页"
                        break

                # Phase 2: If still >1 page, also shrink font globally
                if initial_pages > 1:
                    for shrink_pt in [0.5, 1.0, 1.5]:
                        anything_shrunk = False
                        for i in range(1, doc.Paragraphs.Count + 1):
                            try:
                                p = doc.Paragraphs(i)
                                rng = p.Range
                                # Skip paragraphs inside table cells (handled separately via Range.Cells)
                                if rng.Information(12):  # wdWithInTable
                                    continue
                                if rng.Text.strip() and rng.Font.Size and rng.Font.Size > 7:
                                    rng.Font.Size = max(7, rng.Font.Size - shrink_pt)
                                    anything_shrunk = True
                            except:
                                continue
                        for ti in range(1, doc.Tables.Count + 1):
                            try:
                                t = doc.Tables(ti)
                                try:
                                    for cell in t.Range.Cells:
                                        try:
                                            cf = cell.Range.Font
                                            if cf.Size and cf.Size > 7:
                                                cf.Size = max(7, cf.Size - shrink_pt)
                                                anything_shrunk = True
                                        except:
                                            pass
                                except:
                                    pass
                            except:
                                pass
                        if not anything_shrunk:
                            break
                        pages = _count_pages()
                        if isinstance(pages, int) and pages < initial_pages:
                            initial_pages = pages
                            stats["页面优化"] = f"缩字号-{shrink_pt}pt→{pages}页"
                            break

                # Phase 3: Still not 1 page? Try tighter line spacing
                if initial_pages > 1:
                    for sp in [1.35, 1.25, 1.15, 1.0]:
                        for i in range(1, doc.Paragraphs.Count + 1):
                            try:
                                p = doc.Paragraphs(i)
                                rng = p.Range
                                if rng.Information(12):  # skip table cell paragraphs
                                    continue
                                if rng.Text.strip():
                                    ptype = _get_type(i) or ""
                                    if "正文" in ptype or "节标题" in ptype:
                                        pf = p.Format
                                        pf.LineSpacingRule = 5
                                        pf.LineSpacing = sp * 12
                            except:
                                continue
                        pages = _count_pages()
                        if isinstance(pages, int) and pages < initial_pages:
                            initial_pages = pages
                            stats["页面优化"] = f"缩行距{sp}倍→{pages}页"
                            break

                if initial_pages == 1:
                    stats["页面优化"] = stats.get("页面优化", "") + " → 已压缩至1页!"
                else:
                    stats["页面优化"] = stats.get("页面优化", f"初始{initial_pages}页") + f" → 现{initial_pages}页"
            else:
                stats["页面优化"] = "已1页,无需压缩"

            merge_info = f" ({stats['表格合并']})" if "表格合并" in stats else ""
            msg = (f"智能美化完成!\n\n"
                   f"  文档类型: {doc_meta.get('DOCTYPE', '?')}\n"
                   f"  密度: {doc_meta.get('DENSITY', '?')}  标题: {title_max:.0f}pt  正文: {body_size:.0f}pt  行距: {line_sp}倍\n\n"
                   f"  标题 → {stats['标题']}处\n"
                   f"  正文 → {stats['正文']}段\n"
                   f"  落款 → {stats['落款']}处\n"
                   f"  表格 → {stats['表格']}个{merge_info}\n"
                   f"  {stats.get('页面优化', '')}\n"
                   f"  {stats.get('空白', '')}")
            self.output_edit.setText(f"--- 智能美化 ---\n{msg}\n\nAI分析:\n{analysis[:500]}")
            self.set_status("智能美化完成", "green")
        except Exception as e:
            self.set_status(f"美化失败: {e}", "red")
            self.output_edit.setText(f"美化出错:\n{traceback.format_exc()}")

    def on_diagnose(self):
        """文档诊断：扫描排版问题"""
        try:
            self.conn.ensure()
            self._update_status()
        except ConnectionError as e:
            QtWidgets.QMessageBox.warning(self, "Word 未连接", str(e)[:200])
            return

        self.set_status("诊断中...", "#9C27B0")
        QtCore.QCoreApplication.processEvents()

        try:
            doc = self.conn.doc
            issues = []
            fonts_seen = set()

            # Check page setup
            ps = doc.PageSetup
            tm = ps.TopMargin
            if abs(tm - 2.54 * 28.3464567) > 10:
                issues.append(f"⚠ 页边距: 上={tm/28.35:.1f}cm (建议2.54cm)")

            # Scan paragraphs
            empty_count = 0
            no_indent_count = 0
            for i in range(1, min(doc.Paragraphs.Count + 1, 500)):
                try:
                    p = doc.Paragraphs(i)
                    rng = p.Range
                    txt = rng.Text.strip()

                    if not txt:
                        empty_count += 1
                        continue

                    fn = rng.Font.Name
                    if fn:
                        fonts_seen.add(fn)

                    # Check line spacing
                    pf = p.Format
                    if pf.LineSpacingRule != 5 or pf.LineSpacing < 15:
                        sname = str(rng.Style.NameLocal)
                        if "Heading" not in sname and "标题" not in sname and "TOC" not in sname:
                            no_indent_count += 1  # proxy
                except:
                    continue

            if len(fonts_seen) > 3:
                issues.append(f"⚠ 字体混杂: {len(fonts_seen)}种 ({', '.join(list(fonts_seen)[:5])})")

            if empty_count > doc.Paragraphs.Count * 0.3:
                issues.append(f"⚠ 空段落过多: {empty_count}个 (占{empty_count*100//max(1,doc.Paragraphs.Count)}%)")

            # Check tables
            no_border_tables = 0
            for i in range(1, doc.Tables.Count + 1):
                try:
                    t = doc.Tables(i)
                    if not t.Borders.Enable or t.Borders.InsideLineStyle == 0:
                        no_border_tables += 1
                except:
                    pass

            if no_border_tables > 0:
                issues.append(f"⚠ 表格缺边框: {no_border_tables}/{doc.Tables.Count}个")

            if not issues:
                issues.append("✓ 文档排版良好，未发现明显问题")

            result = "文档诊断结果\n" + "=" * 30 + "\n"
            result += f"段落: {doc.Paragraphs.Count} | 表格: {doc.Tables.Count} | 字体种类: {len(fonts_seen)}\n\n"
            result += "\n".join(issues)
            result += "\n\n提示: 点击「一键美化」可自动修复以上问题"

            self.output_edit.setText(f"--- 文档诊断 ---\n{result}")
            self.set_status("诊断完成", "green")
        except Exception as e:
            self.set_status(f"诊断失败: {e}", "red")

    def on_clean_empty(self):
        """清理多余空行：删连续>2的空段、表格间多余空段、首尾空段"""
        try:
            self.conn.ensure()
            self._update_status()
        except ConnectionError as e:
            QtWidgets.QMessageBox.warning(self, "Word 未连接", str(e)[:200])
            return

        self.set_status("清理中...", "#607D8B")
        QtCore.QCoreApplication.processEvents()

        try:
            doc = self.conn.doc
            deleted = 0

            # Go backwards to avoid index shift
            for i in range(doc.Paragraphs.Count, 0, -1):
                try:
                    p = doc.Paragraphs(i)
                    # Skip table cell paragraphs
                    if p.Range.Information(12):
                        continue
                    txt = p.Range.Text.strip()
                    if txt:
                        continue  # has content, keep it

                    # Empty paragraph - check if it's part of a consecutive run
                    prev_empty = i > 1 and not doc.Paragraphs(i - 1).Range.Text.strip()
                    next_empty = i < doc.Paragraphs.Count and not doc.Paragraphs(i + 1).Range.Text.strip() if i < doc.Paragraphs.Count else False

                    # Delete if: (consecutive empties: keep max 1 as spacer) OR (at doc start/end)
                    if prev_empty or next_empty or i == doc.Paragraphs.Count:
                        p.Range.Delete()
                        deleted += 1
                except:
                    continue

            self.output_edit.setText(f"--- 清理空行 ---\n删除了 {deleted} 个多余空段落\n保留单个空段作为自然分隔")
            self.set_status(f"清理完成 (删除{deleted}个空行)", "green")
        except Exception as e:
            self.set_status(f"清理失败: {e}", "red")

    def on_insert_toc(self):
        """插入自动目录"""
        try:
            self.conn.ensure()
            self._update_status()
        except ConnectionError as e:
            QtWidgets.QMessageBox.warning(self, "Word 未连接", str(e)[:200])
            return

        self.set_status("插入目录...", "#00897B")
        QtCore.QCoreApplication.processEvents()

        try:
            doc = self.conn.doc
            sel = self.conn.sel

            # Check if headings exist
            heading_count = 0
            for i in range(1, min(doc.Paragraphs.Count + 1, 100)):
                try:
                    sname = str(doc.Paragraphs(i).Range.Style.NameLocal)
                    if "Heading" in sname or "标题" in sname:
                        heading_count += 1
                except:
                    pass

            if heading_count == 0:
                self.output_edit.setText("--- 插入目录 ---\n⚠ 未检测到标题样式。\n请先用 Heading 样式设置标题，或使用 AI 生成带标题的文档。")
                self.set_status("无标题，无法生成目录", "#f0a030")
                return

            # Insert TOC at selection
            toc = doc.TablesOfContents.Add(sel.Range)
            toc.UseHeadingStyles = True
            toc.UpperHeadingLevel = 1
            toc.LowerHeadingLevel = 3

            self.output_edit.setText(f"--- 插入目录 ---\n✓ 已在光标位置插入自动目录\n检测到 {heading_count} 个标题")
            self.set_status("目录已插入", "green")
        except Exception as e:
            self.set_status(f"插入目录失败: {e}", "red")

    # ── LLM ───────────────────────────────────────────
    def _call_llm(self, messages):
        if self.cfg["provider"] == "bridge":
            return self._call_bridge(messages)
        if not HAS_REQUESTS:
            raise RuntimeError("requests 未安装")

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.cfg['api_key']}",
        }
        body = {
            "model": self.cfg["model"],
            "messages": messages,
            "temperature": self.cfg["temperature"],
            "max_tokens": self.cfg["max_tokens"],
        }
        url = f"{self.cfg['api_base'].rstrip('/')}/chat/completions"
        resp = requests.post(url, headers=headers, json=body, timeout=60, proxies=self.proxies)
        resp.raise_for_status()
        return _strip_code_fence(resp.json()["choices"][0]["message"]["content"])

    def _call_bridge(self, messages):
        import time
        BRIDGE_DIR.mkdir(parents=True, exist_ok=True)
        with open(BRIDGE_INPUT, "w", encoding="utf-8") as f:
            f.write(messages[-1]["content"])
        if BRIDGE_DONE.exists():
            BRIDGE_DONE.unlink()
        waited = 0
        while waited < 120:
            if self._cancelled:
                raise InterruptedError("用户取消")
            if BRIDGE_OUTPUT.exists() and BRIDGE_DONE.exists():
                code = _strip_code_fence(open(BRIDGE_OUTPUT, "r", encoding="utf-8").read())
                BRIDGE_INPUT.unlink(missing_ok=True)
                BRIDGE_OUTPUT.unlink(missing_ok=True)
                BRIDGE_DONE.unlink(missing_ok=True)
                return code
            time.sleep(0.5)
            waited += 0.5
            QtCore.QCoreApplication.processEvents()
        raise TimeoutError("桥接模式等待超时 (120s)")

    # ── Word document reader ──────────────────────────
    def _read_doc_snapshot(self):
        """Read full document overview: structure, styles, sections, tables + text."""
        try:
            doc = self.conn.doc
            out = [f"文档: {doc.Name}"]

            # Statistics
            try:
                pages = doc.ComputeStatistics(2)
                paras = doc.Paragraphs.Count
                chars = len(doc.Content.Text)
                out.append(f"页数={pages} 段落数={paras} 字符数={chars}")
            except:
                pass

            # Sections
            try:
                out.append(f"\n[节] ({doc.Sections.Count}个):")
                for i in range(1, min(doc.Sections.Count, 10) + 1):
                    sec = doc.Sections(i)
                    ps = sec.PageSetup
                    orient = "横向" if ps.Orientation == 1 else "纵向"
                    out.append(f"  节{i}: {orient} {ps.PageWidth}×{ps.PageHeight}磅")
            except:
                pass

            # Tables
            try:
                tables = doc.Tables
                if tables.Count:
                    table_info = []
                    for i in range(1, min(tables.Count, 20) + 1):
                        t = tables(i)
                        table_info.append(f"  表{i}: {t.Rows.Count}行×{t.Columns.Count}列")
                    out.append(f"\n[表格] ({tables.Count}个):\n" + "\n".join(table_info))
            except:
                pass

            # Images & shapes
            try:
                shapes = doc.InlineShapes
                if shapes.Count:
                    out.append(f"\n[图片/内联形状] {shapes.Count}个")
            except:
                pass

            # Styles in use
            try:
                style_samples = []
                for i in range(1, min(doc.Paragraphs.Count, 60) + 1):
                    try:
                        sname = str(doc.Paragraphs(i).Range.Style.NameLocal)
                        if sname not in style_samples:
                            style_samples.append(sname)
                    except:
                        pass
                if style_samples:
                    out.append(f"\n[使用的样式] ({len(style_samples)}种): " + ", ".join(style_samples[:15]))
            except:
                pass

            # Text content
            text = doc.Content.Text
            if text and text.strip():
                text = text[:2500] if len(text) > 2500 else text
                out.append(f"\n[正文内容]:\n{text}")
                if len(doc.Content.Text) > 2500:
                    out.append("...[已截断]")

            return "\n".join(out)
        except Exception as e:
            return f"读取文档失败: {e}"

    # ── actions ───────────────────────────────────────
    def on_cancel(self):
        self._cancelled = True
        self.set_status("正在取消...", "#f44336")

    def on_query(self):
        question = self.input_edit.toPlainText().strip()
        try:
            self.conn.ensure()
            self._update_status()
        except ConnectionError as e:
            QtWidgets.QMessageBox.warning(self, "Word 未连接",
                f"请先打开 Word 再试。\n\n{str(e)[:200]}")
            return

        snapshot = self._read_doc_snapshot()
        self.output_edit.setText(snapshot)
        self.set_status("AI 分析中...", "#2196F3")
        QtCore.QCoreApplication.processEvents()

        self.query_btn.setEnabled(False)
        self.run_btn.setEnabled(False)
        try:
            if question:
                user_msg = f"以下是 Word 文档的文本内容：\n\n{snapshot}\n\n用户问题：{question}\n\n请根据文档内容回答。用中文简洁回答。"
            else:
                user_msg = f"以下是 Word 文档的文本内容：\n\n{snapshot}\n\n请用中文总结文档内容。概括主题、要点、结构。"

            messages = [
                {"role": "system", "content": "你是文档分析专家。用户会给你文档文本，请根据内容回答问题。"},
                {"role": "user", "content": user_msg},
            ]
            answer = self.client.chat(messages)
            label = "查询结果" if question else "文档总结"
            self.output_edit.setText(f"--- {label} ---\n{answer}")
            self.set_status("分析完成", "green")
        except Exception as e:
            self.set_status(f"分析失败: {e}", "red")
        finally:
            self.query_btn.setEnabled(True)
            self.run_btn.setEnabled(True)

    def on_run(self):
        user_input = self.input_edit.toPlainText().strip()
        if not user_input:
            self.set_status("请输入描述文字", "red")
            return

        try:
            self.conn.ensure()
            self._update_status()
        except ConnectionError as e:
            QtWidgets.QMessageBox.warning(self, "Word 未连接",
                f"请先打开 Word 再试。\n\n{str(e)[:200]}")
            return

        self.run_btn.setEnabled(False)
        self.cancel_btn.show()
        self.output_edit.clear()
        self._cancelled = False

        snapshot = self._read_doc_snapshot()
        base = f"当前文档快照：\n{snapshot}\n\n" if "为空" not in snapshot else ""

        messages = [
            {"role": "system", "content": self.cfg["system_prompt_zh"]},
            {"role": "user", "content": f"{base}用户描述：{user_input}\n\n生成 win32com Python 代码操作 Word。只输出代码。"},
        ]

        MAX_RETRIES = 6
        fixer_code = None
        code = ""
        try:
            for attempt in range(MAX_RETRIES):
                if self._cancelled:
                    self.set_status("已取消", "#f0a030")
                    return

                tag = "生成..." if attempt == 0 else f"修复第{attempt}轮..."
                self.set_status(tag, "#2196F3")
                QtCore.QCoreApplication.processEvents()

                if fixer_code:
                    code = fixer_code
                    fixer_code = None
                    self.output_edit.setText(f"--- 专家修正版 ---\n{code}")
                else:
                    code = self.client.chat(messages)
                    self.output_edit.setText(f"--- {attempt+1}/{MAX_RETRIES} ---\n{code}")

                wd = self.conn.word
                doc = self.conn.doc
                sel = self.conn.sel

                def WRITE(text):
                    sel.TypeText(str(text))

                def PARA():
                    sel.TypeParagraph()

                def BOLD(on=True):
                    sel.Font.Bold = on

                def SIZE(n):
                    sel.Font.Size = n

                def FONT(name):
                    sel.Font.Name = name

                def CENTER():
                    sel.ParagraphFormat.Alignment = 1

                def LEFT():
                    sel.ParagraphFormat.Alignment = 0

                def RIGHT():
                    sel.ParagraphFormat.Alignment = 2

                def HEADING(n, text):
                    sel.Style = doc.Styles(f"Heading {n}")
                    sel.TypeText(text)
                    sel.TypeParagraph()

                def TABLE(rows, cols):
                    return doc.Tables.Add(sel.Range, rows, cols)

                def CELL(t, r, c, text):
                    try:
                        t.Cell(r, c).Range.Text = str(text)
                    except:
                        pass  # merged cell, skip

                def IMAGE(path, w=300, h=200):
                    doc.InlineShapes.AddPicture(path, False, True)

                def FIND(old, new):
                    sel.Find.Execute(old, ReplaceWith=new, Replace=2)

                def GOTO_END():
                    sel.EndKey(Unit=6)

                def GOTO_START():
                    sel.HomeKey(Unit=6)

                def UNDO():
                    try:
                        doc.Undo()
                    except:
                        pass

                # ── page setup ──
                def PAGE_MARGINS(top=2.54, bottom=2.54, left=3.17, right=3.17):
                    CMPT = 28.3464567
                    ps = doc.PageSetup
                    ps.TopMargin = top * CMPT
                    ps.BottomMargin = bottom * CMPT
                    ps.LeftMargin = left * CMPT
                    ps.RightMargin = right * CMPT

                def PAGE_LANDSCAPE():
                    doc.PageSetup.Orientation = 1

                def PAGE_PORTRAIT():
                    doc.PageSetup.Orientation = 0

                def PAGE_BREAK():
                    sel.InsertBreak(Type=7)

                def SECTION_BREAK():
                    sel.InsertBreak(Type=2)

                # ── paragraph formatting ──
                def LINE_SPACING(mult=1.5):
                    pf = sel.ParagraphFormat
                    pf.LineSpacingRule = 5  # wdLineSpacingMultiple
                    pf.LineSpacing = mult * 12

                def PARA_BEFORE(n):
                    sel.ParagraphFormat.SpaceBefore = n

                def PARA_AFTER(n):
                    sel.ParagraphFormat.SpaceAfter = n

                def FIRST_LINE_INDENT(chars=2):
                    pf = sel.ParagraphFormat
                    pf.CharacterUnitFirstLineIndent = chars

                def JUSTIFY():
                    sel.ParagraphFormat.Alignment = 3

                def COLOR(r, g, b):
                    sel.Font.Color = (b << 16) | (g << 8) | r

                # ── title helpers ──
                def TITLE(text):
                    FONT("黑体")
                    SIZE(22)
                    BOLD(True)
                    CENTER()
                    WRITE(text)

                def SUBTITLE(text):
                    FONT("宋体")
                    SIZE(14)
                    CENTER()
                    WRITE(text)

                # ── table formatting ──
                def TABLE_BORDERS(t):
                    t.Borders.Enable = True
                    t.Borders.InsideLineStyle = 1
                    t.Borders.OutsideLineStyle = 1

                def TBOLD(t, row):
                    for c in range(1, t.Columns.Count + 1):
                        try:
                            t.Cell(row, c).Range.Font.Bold = True
                        except:
                            pass

                def COL_WIDTHS(t, widths_cm):
                    CMPT = 28.3464567
                    for i, w in enumerate(widths_cm):
                        if i < t.Columns.Count:
                            try:
                                t.Columns(i + 1).Width = w * CMPT
                            except:
                                pass

                def CELL_MERGE(t, r1, c1, r2, c2):
                    try:
                        t.Cell(r1, c1).Merge(t.Cell(r2, c2))
                    except:
                        pass

                def CELL_ALIGN(t, r, c, align=1):
                    try:
                        t.Cell(r, c).Range.ParagraphFormat.Alignment = align
                    except:
                        pass

                def CELL_BG(t, r, c, rgb):
                    """rgb=(r,g,b) 0-255"""
                    try:
                        t.Cell(r, c).Shading.BackgroundPatternColor = \
                            (rgb[2] << 16) | (rgb[1] << 8) | rgb[0]
                    except:
                        pass

                def CELL_FONT(t, r, c, name, size, bold=False):
                    try:
                        cr = t.Cell(r, c).Range
                        cr.Font.Name = name
                        cr.Font.Size = size
                        cr.Font.Bold = bold
                    except:
                        pass

                # ── document helpers ──
                def TABLE_HEADER(t, row, texts):
                    """Fill a header row with texts and make it bold."""
                    for i, txt in enumerate(texts):
                        CELL(t, row, i + 1, txt)
                    TBOLD(t, row)
                    for i in range(1, len(texts) + 1):
                        CELL_ALIGN(t, row, i, 1)

                def NEW_PAGE():
                    PAGE_BREAK()

                ns = {
                    "word": wd, "doc": doc, "sel": sel, "math": __import__("math"),
                    # basic
                    "WRITE": WRITE, "PARA": PARA, "BOLD": BOLD, "SIZE": SIZE,
                    "FONT": FONT, "CENTER": CENTER, "LEFT": LEFT, "RIGHT": RIGHT,
                    "JUSTIFY": JUSTIFY, "COLOR": COLOR,
                    # heading
                    "HEADING": HEADING, "TITLE": TITLE, "SUBTITLE": SUBTITLE,
                    # table
                    "TABLE": TABLE, "CELL": CELL, "TABLE_BORDERS": TABLE_BORDERS,
                    "TBOLD": TBOLD, "COL_WIDTHS": COL_WIDTHS, "CELL_MERGE": CELL_MERGE,
                    "CELL_ALIGN": CELL_ALIGN, "CELL_BG": CELL_BG, "CELL_FONT": CELL_FONT,
                    "TABLE_HEADER": TABLE_HEADER,
                    # page
                    "PAGE_MARGINS": PAGE_MARGINS, "PAGE_LANDSCAPE": PAGE_LANDSCAPE,
                    "PAGE_PORTRAIT": PAGE_PORTRAIT, "PAGE_BREAK": PAGE_BREAK,
                    "SECTION_BREAK": SECTION_BREAK, "NEW_PAGE": NEW_PAGE,
                    # paragraph
                    "LINE_SPACING": LINE_SPACING, "PARA_BEFORE": PARA_BEFORE,
                    "PARA_AFTER": PARA_AFTER, "FIRST_LINE_INDENT": FIRST_LINE_INDENT,
                    # other
                    "IMAGE": IMAGE, "FIND": FIND, "GOTO_END": GOTO_END,
                    "GOTO_START": GOTO_START, "UNDO": UNDO,
                }
                try:
                    exec(code, ns)
                    suffix = f" (第{attempt+1}次)" if attempt > 0 else ""
                    self.set_status(f"执行成功{suffix}", "green")
                    return
                except Exception as exec_err:
                    trace = traceback.format_exc()
                    short = explain_error(trace, code)
                    _log(f"Exec err {attempt+1}: {short}")

                    if attempt < MAX_RETRIES - 1:
                        self.output_edit.setText(f"--- 错误, 修复第{attempt+1}轮 ---\n{code}\n\n{short}")
                        search_hint = web_search(short, code[:500])
                        self.set_status("Claude专家诊断中...", "#9C27B0")
                        QtCore.QCoreApplication.processEvents()
                        fixed = fix_code(WORD_FIXER_PROMPT, short, code, search_hint, self.cfg, self.proxies)
                        if fixed:
                            fixer_code = fixed
                            messages.append({"role": "assistant", "content": code})
                            messages.append({"role": "user", "content": f"代码报错:\n{short}\n\n专家已修正，直接执行。"})
                            self.set_status("专家已修正，直接执行...", "#9C27B0")
                        else:
                            messages.append({"role": "assistant", "content": code})
                            feedback = f"代码报错:\n{short}\n\n"
                            if search_hint:
                                feedback += f"{search_hint}\n\n"
                            feedback += "修正后输出完整代码。"
                            messages.append({"role": "user", "content": feedback})
                            self.set_status(f"修复第{attempt+2}轮...", "#f0a030")
                    else:
                        self.output_edit.setText(f"--- {MAX_RETRIES}轮后仍失败 ---\n{code}\n\n{short}")
                        self.set_status(f"失败({MAX_RETRIES}轮)", "red")
        except InterruptedError:
            self.set_status("已取消", "#f0a030")
        except TimeoutError as e:
            self.set_status(f"超时: {e}", "red")
        except Exception as e:
            self.set_status(f"错误: {e}", "red")
            self.output_edit.setText(traceback.format_exc())
            _log(f"API/System error:\n{traceback.format_exc()}")
        finally:
            self.run_btn.setEnabled(True)
            self.cancel_btn.hide()

    def set_status(self, text, color="#888"):
        self.status_lbl.setText(text)
        self.status_lbl.setStyleSheet(f"color: {color}; font-size: 11px; padding: 4px;")

# ── entry ─────────────────────────────────────────────────
def run_app():
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    conn = WordConnection()
    conn.connect(create=False)
    if not conn.connected:
        conn.connect(create=True)
    window = TextToCADApp(conn)
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    if MISSING:
        import ctypes
        msg = "缺少以下依赖:\n\n" + "\n".join(f"  - {d}" for d in MISSING)
        msg += "\n\n请用 pip 安装后重试。"
        ctypes.windll.user32.MessageBoxW(0, msg, "TextToCAD - 依赖缺失", 0x10)
        sys.exit(1)
    run_app()
