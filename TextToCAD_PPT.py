# TextToCAD v3 - PowerPoint Edition
"""Standalone app: Chinese NL -> LLM -> PowerPoint COM -> generate slides."""
import os, sys, json, re, traceback, subprocess
from pathlib import Path

LOG_FILE = os.path.join(os.path.expanduser("~"), "t2cad_ppt_debug.log")

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

CONFIG_DIR = Path.home() / ".text_to_cad"
CONFIG_FILE = CONFIG_DIR / "config.json"
BRIDGE_DIR = CONFIG_DIR / "bridge"
BRIDGE_INPUT = BRIDGE_DIR / "input.txt"
BRIDGE_OUTPUT = BRIDGE_DIR / "output.py"
BRIDGE_DONE = BRIDGE_DIR / "done.txt"

PPT_FIXER_PROMPT = """\
你是 PowerPoint COM / win32com 底层调试专家。

## 任务
用户代码在 exec() 中执行报错。判断原因并输出修正后的完整 Python 代码。

## 环境说明
- 已注入: ppt=PowerPoint应用 pres=演示文稿 slide=当前幻灯片
- 禁止: import / Dispatch / EnsureDispatch / Presentations.Open
- 幻灯片大小: 960x540 (16:9)

## 已注入的安全函数
NEW_SLIDE(12)=空白页 GOTO(n)=跳到第n页 DEL_SLIDE(n)=删除 CLEAR()=清空当前页
TB(x,y,w,h,text,size)=文本框 SHAPE(type,x,y,w,h,color)=形状
TABLE_SLIDE(r,c,x,y,w,h)=表格 CELL_SLIDE(t,r,c,text)=填单元格
TABLE_STYLE(t,color)=美化表格 IMG(path,x,y,w,h)=图片
FILL(sh,color)=填充 FONT_COLOR(sh,color)=文字颜色
TB_STYLE(x,y,w,h,text,size,font_color,fill_color,align)=带样式文本框
ALIGN_SHAPE(sh,align)=对齐 Z_ORDER(sh,pos)=层级 COLOR(r,g,b)=RGB PTS(mm)=毫米转磅

## 常见陷阱
- Shape.AddTextbox 第一个参数是 Orientation(1=水平)
- 颜色用 RGB 十六进制: 0x1A237E 不是 "1A237E"
- 设置文字后要设字体名和大小
- 不能用 ppt.Visible=True（已设置）

只输出修正后的完整 Python 代码。"""

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
        "你是 PowerPoint 设计专家。幻灯片 960x540 (16:9)。每次生成都考虑美观。\n\n"
        "## ⚠ 设计铁律\n"
        "1. 每页不超过5个要点，文字简洁，字号>=18pt\n"
        "2. 配色方案：深蓝(#1A237E)背景+白字，或白底+深蓝标题+灰色正文\n"
        "3. 专业色板: 蓝0x1A237E 橙0xFF6F00 绿0x2E7D32 灰0x616161\n"
        "4. 标题用微软雅黑 32-44pt，正文用微软雅黑 18-24pt\n"
        "5. 封面页: 深色背景+大标题居中，内页: 统一左上角标题+内容\n"
        "6. 表格要有表头深色行+边框，图片要居中\n"
        "7. 用 FILL(sh, color) 设形状填充, FONT_COLOR(sh, color) 设文字颜色\n\n"
        "## 核心函数(已注入)\n"
        "NEW_SLIDE(12)=空白页(1=标题页)  GOTO(n)=跳到第n页  DEL_SLIDE(n)=删除\n"
        "CLEAR()=清空当前页  UNDO()=撤销\n"
        "TB(x,y,w,h,text,size=32,color=None)=文本框\n"
        "TB_STYLE(x,y,w,h,text,size,font_color,fill_color,align)=带样式的文本框\n"
        "SHAPE(t,x,y,w,h,color=None)=形状(1=矩形,9=椭圆,14=圆角矩形)\n"
        "FILL(sh,color)=设填充  FONT_COLOR(sh,color)=设文字颜色\n"
        "TABLE_SLIDE(r,c,x,y,w,h)=表格  CELL_SLIDE(t,r,c,text)=填单元格\n"
        "TABLE_STYLE(t,header_color)=表格美化  IMG(path,x,y,w,h)=图片\n"
        "ALIGN_SHAPE(sh,align)=对齐(1左2中3右)  Z_ORDER(sh,pos)=层级(0最前1最后)\n"
        "COLOR(r,g,b)=RGB颜色  PTS(x,y)=坐标转换(毫米→磅)\n\n"
        "## 示例: 专业封面+内容页\n"
        "s = NEW_SLIDE(12)\n"
        "bg = SHAPE(1, 0, 0, 960, 540, 0x1A237E)  # 深蓝全幅背景\n"
        "TB(80, 160, 800, 100, \"2026年度报告\", 44, 0xFFFFFF)\n"
        "TB(80, 280, 800, 60, \"云县低效林改造项目\", 24, 0xB0BEC5)\n"
        "PARA()\n"
        "s2 = NEW_SLIDE(12)\n"
        "SHAPE(1, 0, 0, 960, 80, 0x1A237E)  # 顶部色条\n"
        "TB(60, 20, 840, 50, \"项目概况\", 32, 0xFFFFFF)\n"
        "TB(60, 120, 400, 300, \" 项目地点: 云南云县\\n 改造面积: 10334亩\\n 小班数量: 197个\", 22, 0x333333)\n\n"
        "## 铁律: 禁止import 禁止Dispatch 用NEW_SLIDE/TB/SHAPE 只输出代码 ppt=PowerPoint pres=文稿 slide=当前页"
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
    lines = exc_info.strip().split("\n")
    last_file = None
    for i, line in enumerate(lines):
        if line.strip().startswith("File "):
            last_file = i
    if last_file is not None:
        return "\n".join(lines[last_file:]).strip()
    return "\n".join(lines[-3:]).strip()

class PPTConnection:
    def __init__(self):
        self.ppt = None
        self.connected = False
        self.error = ""

    def connect(self, create=False):
        import win32com.client
        try:
            self.ppt = win32com.client.Dispatch("PowerPoint.Application")
            self.ppt.Visible = 1
            _ = self.ppt.Name
            self.connected = True
            self.error = ""
            return True
        except Exception as e:
            self.connected = False
            self.error = str(e)
            return False

    def ensure(self):
        if self.connected and self.ppt:
            try:
                _ = self.ppt.Name
                return
            except:
                self.connected = False
        if not self.connect(create=True):
            raise ConnectionError(self.error)

    @property
    def pres(self):
        pres = self.ppt.ActivePresentation
        if pres is None:
            pres = self.ppt.Presentations.Add()
        return pres

    @property
    def slide(self):
        pres = self.pres
        try:
            return self.ppt.ActiveWindow.View.Slide
        except:
            if pres.Slides.Count == 0:
                pres.Slides.Add(1, 12)
            return pres.Slides(1)

class TextToCADApp(QtWidgets.QMainWindow):
    def __init__(self, conn):
        super().__init__()
        self.cfg = load_config()
        self.proxies = _resolve_proxies(self.cfg.get("proxies")) if HAS_REQUESTS else None
        self.conn = conn
        self._cancelled = False
        self.client = LLMClient(self.cfg)
        self.pipeline = CodeGenPipeline(self.client)

        self.setWindowTitle("TextToCAD for PowerPoint v1")
        self.setMinimumSize(400, 480)
        self.resize(520, 580)

        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("文件")
        act_check = file_menu.addAction("检查 PowerPoint 连接")
        act_check.triggered.connect(self._check_ppt)
        act_settings = file_menu.addAction("设置...")
        act_settings.triggered.connect(lambda: os.startfile(str(CONFIG_FILE)))
        file_menu.addSeparator()
        file_menu.addAction("退出").triggered.connect(self.close)
        help_menu = menu_bar.addMenu("帮助")
        help_menu.addAction("关于").triggered.connect(self._about)

        cw = QtWidgets.QWidget()
        self.setCentralWidget(cw)
        layout = QtWidgets.QVBoxLayout(cw)
        layout.setContentsMargins(10, 10, 10, 10)

        title = QtWidgets.QLabel("用自然语言生成 PowerPoint")
        title.setStyleSheet("font-size: 14px; font-weight: bold; padding: 4px;")
        layout.addWidget(title)

        self.input_edit = QtWidgets.QTextEdit()
        self.input_edit.setPlaceholderText(
            "例如：\n"
            "  创建5页演示文稿：封面标题「2024年度总结」，第2页列出三个要点，第3页插入表格...\n"
            "  当前页加一个蓝色矩形背景，上面放白色标题文字\n"
            "  把第3页移到第1页后面\n"
            "  在第2页插入图片 C:\\chart.png，居中\n"
            "  删除当前幻灯片的所有内容\n"
            "  总结当前演示文稿的结构"
        )
        self.input_edit.setMaximumHeight(140)
        layout.addWidget(self.input_edit)

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

        self.query_btn = QtWidgets.QPushButton("分析文稿")
        self.query_btn.setStyleSheet(
            "QPushButton { background-color: #4CAF50; color: white; font-size: 13px;"
            " font-weight: bold; padding: 8px 20px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #388E3C; }"
            "QPushButton:disabled { background-color: #666; }"
        )
        self.query_btn.clicked.connect(self.on_query)
        btn_layout.addWidget(self.query_btn)

        self.reconnect_btn = QtWidgets.QPushButton("重连PPT")
        self.reconnect_btn.setStyleSheet(
            "QPushButton { font-size: 13px; padding: 8px 20px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #e0e0e0; }"
        )
        self.reconnect_btn.clicked.connect(self._reconnect)
        btn_layout.addWidget(self.reconnect_btn)
        layout.addLayout(btn_layout)

        # ── 快捷工具 ──
        tools_layout = QtWidgets.QHBoxLayout()
        tools_label = QtWidgets.QLabel("快捷:")
        tools_label.setStyleSheet("font-size: 11px; color: #888; padding-right: 4px;")
        tools_layout.addWidget(tools_label)

        self.font_btn = QtWidgets.QPushButton("统一字体")
        self.font_btn.setToolTip("全文统一为微软雅黑（标题36pt/正文20pt）")
        self.font_btn.setStyleSheet(
            "QPushButton{background:#FF9800;color:white;font-size:11px;font-weight:bold;padding:5px 10px;border-radius:3px}"
            "QPushButton:hover{background:#F57C00}")
        self.font_btn.clicked.connect(self.on_unify_fonts)
        tools_layout.addWidget(self.font_btn)

        self.align_btn = QtWidgets.QPushButton("居中对齐")
        self.align_btn.setToolTip("当前页所有形状水平居中")
        self.align_btn.setStyleSheet(
            "QPushButton{background:#607D8B;color:white;font-size:11px;font-weight:bold;padding:5px 10px;border-radius:3px}"
            "QPushButton:hover{background:#455A64}")
        self.align_btn.clicked.connect(self.on_center_all)
        tools_layout.addWidget(self.align_btn)

        self.cover_btn = QtWidgets.QPushButton("快速封面")
        self.cover_btn.setToolTip("在当前页插入标准封面: 深蓝背景+标题居中")
        self.cover_btn.setStyleSheet(
            "QPushButton{background:#00897B;color:white;font-size:11px;font-weight:bold;padding:5px 10px;border-radius:3px}"
            "QPushButton:hover{background:#00695C}")
        self.cover_btn.clicked.connect(self.on_quick_cover)
        tools_layout.addWidget(self.cover_btn)

        self.new_blank_btn = QtWidgets.QPushButton("新建空白页")
        self.new_blank_btn.setToolTip("在当前文稿末尾添加一页空白幻灯片")
        self.new_blank_btn.setStyleSheet(
            "QPushButton{background:#9C27B0;color:white;font-size:11px;font-weight:bold;padding:5px 10px;border-radius:3px}"
            "QPushButton:hover{background:#7B1FA2}")
        self.new_blank_btn.clicked.connect(self.on_new_blank)
        tools_layout.addWidget(self.new_blank_btn)

        tools_layout.addStretch()
        layout.addLayout(tools_layout)

        self.status_lbl = QtWidgets.QLabel("就绪 — Ctrl+Enter 快速执行")
        self.status_lbl.setStyleSheet("color: #888; font-size: 11px; padding: 4px;")
        self.status_lbl.setWordWrap(True)
        layout.addWidget(self.status_lbl)

        self.output_edit = QtWidgets.QTextEdit()
        self.output_edit.setReadOnly(True)
        self.output_edit.setMaximumHeight(200)
        self.output_edit.setStyleSheet(
            "QTextEdit { background-color: #1e1e1e; color: #d4d4d4;"
            " font-family: Consolas, monospace; font-size: 11px; }"
        )
        layout.addWidget(self.output_edit)

        proxy_info = " (代理已检测)" if self.proxies else ""
        mode_text = f"模式: {self.cfg['provider']} | 模型: {self.cfg['model']}{proxy_info}"
        self.mode_lbl = QtWidgets.QLabel(mode_text)
        self.mode_lbl.setStyleSheet("color: #f0a030; font-size: 10px;")
        self.mode_lbl.setWordWrap(True)
        layout.addWidget(self.mode_lbl)

        self.ppt_status = QtWidgets.QLabel()
        self._update_status()
        self.statusBar().addPermanentWidget(self.ppt_status)

        QtGui.QShortcut(QtGui.QKeySequence("Ctrl+Return"), self).activated.connect(self.on_run)

    def _update_status(self):
        if self.conn.connected:
            self.ppt_status.setText("PowerPoint:  已连接")
            self.ppt_status.setStyleSheet("color: #4caf50; font-weight: bold; padding: 2px 8px;")
            self.reconnect_btn.hide()
        else:
            self.ppt_status.setText("PowerPoint:  未连接")
            self.ppt_status.setStyleSheet("color: #f0a030; font-weight: bold; padding: 2px 8px;")
            self.reconnect_btn.show()

    def _reconnect(self):
        self.set_status("正在连接 PowerPoint...", "#2196F3")
        QtCore.QCoreApplication.processEvents()
        ok = self.conn.connect(create=True)
        self._update_status()
        self.set_status("已连接" if ok else f"连接失败: {self.conn.error}", "green" if ok else "red")

    def _check_ppt(self):
        try:
            self.conn.ensure()
            QtWidgets.QMessageBox.information(self, "PowerPoint 连接",
                f"已连接\n文稿: {self.conn.pres.Name}\n幻灯片数: {self.conn.pres.Slides.Count}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "PowerPoint 连接", str(e))
        self._update_status()

    def _about(self):
        QtWidgets.QMessageBox.about(self, "关于 TextToCAD",
            "TextToCAD for PowerPoint v2\n\n用自然语言生成演示文稿 + 快捷工具\nLLM: DeepSeek / OpenAI API")

    # ── 快捷工具 ──
    def on_unify_fonts(self):
        """统一全文字体"""
        try: self.conn.ensure(); self._update_status()
        except ConnectionError as e:
            QtWidgets.QMessageBox.warning(self, "PPT未连接", str(e)); return
        try:
            pres = self.conn.pres
            fixed = 0
            for si in range(1, pres.Slides.Count + 1):
                try:
                    for sh in pres.Slides(si).Shapes:
                        try:
                            if sh.HasTextFrame and sh.TextFrame.HasText:
                                tr = sh.TextFrame.TextRange
                                tr.Font.Name = "微软雅黑"
                                # Larger text → likely a title
                                if tr.Font.Size > 28:
                                    tr.Font.Size = 36
                                elif tr.Font.Size > 16:
                                    tr.Font.Size = 20
                                fixed += 1
                        except: continue
                except: continue
            self.set_status(f"统一字体完成 ({fixed}处)", "green")
        except Exception as e: self.set_status(f"失败: {e}", "red")

    def on_center_all(self):
        """当前页所有形状居中"""
        try: self.conn.ensure()
        except: pass
        try:
            slide = self.conn.slide
            cnt = 0
            for sh in slide.Shapes:
                try:
                    sh.Left = (960 - sh.Width) / 2
                    cnt += 1
                except: continue
            self.set_status(f"居中完成 ({cnt}个形状)", "green")
        except Exception as e: self.set_status(f"失败: {e}", "red")

    def on_quick_cover(self):
        """在当前页插入快速封面"""
        try: self.conn.ensure()
        except: pass
        try:
            slide = self.conn.slide
            # Clear current slide
            for sh in list(slide.Shapes):
                try: sh.Delete()
                except: pass
            # Full blue background
            bg = slide.Shapes.AddShape(1, 0, 0, 960, 540)
            bg.Fill.ForeColor.RGB = 0x1A237E
            bg.Line.Visible = 0
            # Title
            tb = slide.Shapes.AddTextbox(1, 60, 180, 840, 100)
            tb.TextFrame.TextRange.Text = "标题"
            tb.TextFrame.TextRange.Font.Size = 44
            tb.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
            tb.TextFrame.TextRange.Font.Name = "微软雅黑"
            # Subtitle
            tb2 = slide.Shapes.AddTextbox(1, 60, 300, 840, 50)
            tb2.TextFrame.TextRange.Text = "副标题"
            tb2.TextFrame.TextRange.Font.Size = 24
            tb2.TextFrame.TextRange.Font.Color.RGB = 0xB0BEC5
            tb2.TextFrame.TextRange.Font.Name = "微软雅黑"
            self.set_status("封面已创建", "green")
        except Exception as e: self.set_status(f"失败: {e}", "red")

    def on_new_blank(self):
        """添加空白幻灯片"""
        try: self.conn.ensure()
        except: pass
        try:
            pres = self.conn.pres
            pres.Slides.Add(pres.Slides.Count + 1, 12)
            self.set_status(f"已添加空白页 (共{pres.Slides.Count}页)", "green")
        except Exception as e: self.set_status(f"失败: {e}", "red")

    def _call_llm(self, messages):
        if self.cfg["provider"] == "bridge":
            return self._call_bridge(messages)
        if not HAS_REQUESTS:
            raise RuntimeError("requests 未安装")
        headers = {"Content-Type": "application/json",
                   "Authorization": f"Bearer {self.cfg['api_key']}"}
        body = {"model": self.cfg["model"], "messages": messages,
                "temperature": self.cfg["temperature"], "max_tokens": self.cfg["max_tokens"]}
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

    def _read_pres_snapshot(self):
        try:
            pres = self.conn.pres
            lines = [f"文稿: {pres.Name}, 幻灯片数: {pres.Slides.Count}"]
            for i in range(1, min(pres.Slides.Count, 30) + 1):
                try:
                    sl = pres.Slides(i)
                    shapes_info = []
                    for si in range(1, min(sl.Shapes.Count, 10) + 1):
                        try:
                            sh = sl.Shapes(si)
                            txt = ""
                            if sh.HasTextFrame and sh.TextFrame.HasText:
                                txt = sh.TextFrame.TextRange.Text[:80].replace('\r', ' ')
                            shapes_info.append(f"  [{si}] {sh.Name}({sh.Type}) \"{txt}\"")
                        except:
                            pass
                    lines.append(f"\n第{i}页 ({sl.Shapes.Count}个形状):")
                    lines.extend(shapes_info)
                except:
                    pass
            return "\n".join(lines)
        except Exception as e:
            return f"读取失败: {e}"

    def on_cancel(self):
        self._cancelled = True
        self.set_status("正在取消...", "#f44336")

    def on_query(self):
        question = self.input_edit.toPlainText().strip()
        try:
            self.conn.ensure()
            self._update_status()
        except ConnectionError as e:
            QtWidgets.QMessageBox.warning(self, "PowerPoint 未连接", str(e))
            return

        snapshot = self._read_pres_snapshot()
        self.output_edit.setText(snapshot)
        self.set_status("AI 分析中...", "#2196F3")
        QtCore.QCoreApplication.processEvents()

        self.query_btn.setEnabled(False)
        self.run_btn.setEnabled(False)
        try:
            if question:
                user_msg = f"演示文稿结构：\n\n{snapshot}\n\n问题：{question}"
            else:
                user_msg = f"演示文稿结构：\n\n{snapshot}\n\n请用中文总结这个PPT的结构和内容。"
            messages = [
                {"role": "system", "content": "你是演示文稿分析专家。请根据结构信息回答问题。"},
                {"role": "user", "content": user_msg},
            ]
            answer = self.client.chat(messages)
            label = "查询结果" if question else "文稿总结"
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
            QtWidgets.QMessageBox.warning(self, "PowerPoint 未连接", str(e))
            return

        self.run_btn.setEnabled(False)
        self.cancel_btn.show()
        self.output_edit.clear()
        self._cancelled = False

        snapshot = self._read_pres_snapshot()
        base = f"当前文稿快照：\n{snapshot}\n\n"

        messages = [
            {"role": "system", "content": self.cfg["system_prompt_zh"]},
            {"role": "user", "content": f"{base}用户描述：{user_input}\n\n生成 win32com Python 代码操作 PowerPoint。只输出代码。"},
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

                pt = self.conn.ppt
                pres = self.conn.pres
                slide = self.conn.slide

                def NEW_SLIDE(layout=12):
                    s = pres.Slides.Add(pres.Slides.Count + 1, layout)
                    return s

                def GOTO(n):
                    return pres.Slides(n)

                def DEL_SLIDE(n):
                    pres.Slides(n).Delete()

                def CLEAR():
                    for sh in list(slide.Shapes):
                        try:
                            sh.Delete()
                        except:
                            pass

                def TB(x, y, w, h, text, size=32):
                    tb = slide.Shapes.AddTextbox(1, x, y, w, h)
                    tr = tb.TextFrame.TextRange
                    tr.Text = str(text)
                    tr.Font.Size = size
                    return tb

                def SHAPE(stype, x, y, w, h, color=None):
                    s = slide.Shapes.AddShape(stype, x, y, w, h)
                    if color:
                        s.Fill.ForeColor.RGB = color
                    return s

                def TABLE_SLIDE(rows, cols, x, y, w, h):
                    return slide.Shapes.AddTable(rows, cols, x, y, w, h).Table

                def CELL_SLIDE(t, r, c, text):
                    t.Cell(r, c).Shape.TextFrame.TextRange.Text = str(text)

                def IMG(path, x, y, w, h):
                    slide.Shapes.AddPicture(path, 0, -1, x, y, w, h)

                def UNDO():
                    try:
                        pt.CommandBars.FindControl(Id=128).Execute()
                    except:
                        pass

                # ── formatting helpers ──
                def FILL(sh, color):
                    try: sh.Fill.ForeColor.RGB = color
                    except: pass

                def FONT_COLOR(sh, color):
                    try:
                        if sh.HasTextFrame and sh.TextFrame.HasText:
                            sh.TextFrame.TextRange.Font.Color.RGB = color
                    except: pass

                def TB_STYLE(x, y, w, h, text, size=32, font_color=None, fill_color=None, align=2):
                    """Styled textbox with optional fill and alignment."""
                    tb = slide.Shapes.AddTextbox(1, x, y, w, h)
                    if fill_color:
                        FILL(tb, fill_color)
                    tr = tb.TextFrame.TextRange
                    tr.Text = str(text)
                    tr.Font.Size = size
                    if font_color:
                        tr.Font.Color.RGB = font_color
                    if align == 2:
                        tr.ParagraphFormat.Alignment = 2  # center
                    return tb

                def ALIGN_SHAPE(sh, align=2):
                    """Align shape: 1=left, 2=center, 3=right relative to slide."""
                    try:
                        if align == 2:
                            sh.Left = (960 - sh.Width) / 2
                        elif align == 3:
                            sh.Left = 960 - sh.Width - 30
                        else:
                            sh.Left = 30
                    except: pass

                def Z_ORDER(sh, pos=1):
                    """pos: 0=front, 1=back"""
                    try:
                        if pos == 1: sh.ZOrder(2)  # msoSendToBack
                        else: sh.ZOrder(0)  # msoBringToFront
                    except: pass

                def TABLE_STYLE(t, header_color=0x1A237E):
                    """Style table: header row dark bg white text, borders."""
                    try:
                        for c in range(1, t.Columns.Count + 1):
                            cell = t.Cell(1, c)
                            cell.Shape.Fill.ForeColor.RGB = header_color
                            if cell.Shape.HasTextFrame:
                                cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
                                cell.Shape.TextFrame.TextRange.Font.Bold = True
                    except: pass

                def COLOR(r, g, b):
                    """RGB (0-255) -> BGR hex for COM."""
                    return (b << 16) | (g << 8) | r

                def PTS(mm_val):
                    """Convert mm to PowerPoint points (1pt ≈ 0.353mm)."""
                    return mm_val / 0.3528

                ns = {
                    "ppt": pt, "pres": pres, "slide": slide, "math": __import__("math"),
                    "NEW_SLIDE": NEW_SLIDE, "GOTO": GOTO, "DEL_SLIDE": DEL_SLIDE,
                    "CLEAR": CLEAR, "TB": TB, "TB_STYLE": TB_STYLE,
                    "SHAPE": SHAPE, "FILL": FILL, "FONT_COLOR": FONT_COLOR,
                    "TABLE_SLIDE": TABLE_SLIDE, "CELL_SLIDE": CELL_SLIDE,
                    "TABLE_STYLE": TABLE_STYLE, "ALIGN_SHAPE": ALIGN_SHAPE,
                    "Z_ORDER": Z_ORDER, "COLOR": COLOR, "PTS": PTS,
                    "IMG": IMG, "UNDO": UNDO,
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
                        fixed = fix_code(PPT_FIXER_PROMPT, short, code, search_hint, self.cfg, self.proxies)
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

def run_app():
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    conn = PPTConnection()
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
