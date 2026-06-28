# TextToCAD v3 - PySide6 + pyautocad + DeepSeek API
"""Standalone app: Chinese NL → LLM → pyautocad → AutoCAD drawing."""
import os, sys, json, re, traceback, subprocess
from pathlib import Path

LOG_FILE = os.path.join(os.path.expanduser("~"), "t2cad_autocad_debug.log")

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

try:
    from pyautocad import Autocad, APoint, aDouble
    HAS_PYAUTOCAD = True
except ImportError:
    HAS_PYAUTOCAD = False
    MISSING.append("pyautocad (pip install pyautocad)")

# ── paths ────────────────────────────────────────────────
CONFIG_DIR = Path.home() / ".text_to_cad"
CONFIG_FILE = CONFIG_DIR / "config.json"
BRIDGE_DIR = CONFIG_DIR / "bridge"
BRIDGE_INPUT = BRIDGE_DIR / "input.txt"
BRIDGE_OUTPUT = BRIDGE_DIR / "output.py"
BRIDGE_DONE = BRIDGE_DIR / "done.txt"

ACAD_FIXER_PROMPT = """\
你是 AutoCAD COM / pyautocad / win32com 底层调试专家。

## 任务
用户代码在 exec() 中执行报错。判断原因并输出修正后的完整 Python 代码。

## 环境说明
- 已注入: acad=pyautocad对象(acad.model=模型空间)
- APoint(x,y,z) 创建点 aDouble(*vals) 创建数组
- 禁止: import / Dispatch / Autocad() 实例化

## 已注入的安全函数
FAST_MODE(True/False) L() C() R() T() DIM() PLINE() ARC() HATCH() MTEXT()
MOVE() ROT() DEL() OBJS() FIND_OBJS() DEL_BY_TYPE() UNDO()
LAYER(name,color) SEND_CMD(cmd) ZOOM_EXT() ZOOM_WIN() BLOCK_INSERT()
颜色常量: RED=1 YELLOW=2 GREEN=3 CYAN=4 BLUE=5 MAGENTA=6 WHITE=7
SET_COLOR(obj,color) DIM_H() DIM_STYLE()

## 常见陷阱
- acad.model 是模型空间 acad.app.ActiveDocument 是当前文档
- 模型空间对象通过 ms.AddXxx() 创建
- 点坐标用 APoint(x,y,0)
- AddLightWeightPolyline 需要 aDouble(x1,y1,x2,y2,...)
- 旋转角度用 math.radians(deg)
- SendCommand 字符串要以 \\n 结尾

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
    "units": "mm",
    "auto_execute": True,
    "show_code": True,
    "system_prompt_zh": (
        "你是 AutoCAD + 天正建筑 绘图专家。输出 Python 代码。单位: 毫米(1m=1000, 1cm=10)。\n\n"
        "## ⚠ 绘图规范\n"
        "1. 建筑图优先用天正命令画专业对象(墙/门窗/柱/轴网),不要用基础线条硬画\n"
        "2. 每次绘图先建图层: LAYER(\"层名\", color=1-7)\n"
        "3. ⚠ 性能建议: 大量同类对象(如草地/树阵)用 HATCH 填充代替循环画圆, 速度差100倍\n"
        "   可用 BATCH(函数列表) 批量操作减少等待, 例如 BATCH([lambda: C(x,y,r,GREEN) for x in ...])\n"
        "4. 颜色用常量: RED=1 YELLOW=2 GREEN=3 CYAN=4 BLUE=5 MAGENTA=6 WHITE=7\n"
        "   所有绘图函数都支持 color=颜色: L(0,0,100,100,color=GREEN)\n"
        "   已画对象改色: SET_COLOR(obj, RED)\n"
        "   不同对象用不同颜色: 绿地绿植=GREEN, 道路=YELLOW, 水系=CYAN, 标注=GREEN, 轮廓=WHITE\n"
        "4. 文字高度: 图名500 标注300 说明200\n"
        "5. 画完后 ZOOM_EXT() 缩放全图\n\n"
        "## 天正建筑命令 (用 SEND_CMD 调用)\n"
        "SEND_CMD(\"命令名\")=发送AutoCAD命令(支持天正/源泉/任何CAD命令)\n"
        "── 轴网 ──\n"
        "  TRectAxis=直线轴网 TArcAxis=弧线轴网 TSingleAxis=单轴线 TAxisDote=轴网标注\n"
        "── 墙体 ──\n"
        "  TgWall=绘制墙体 TSWall=生线变墙 TDivWall=等分加墙 TFillet=倒墙角 TOffset=墙偏移\n"
        "  TWallThick=改墙厚 TChHeight=改高度 TAddInsulate=加保温层\n"
        "── 柱子 ──\n"
        "  TGColumn=标准柱 TCornColu=角柱 TPolyColu=多段柱 TAlignColu=柱齐墙\n"
        "── 门窗 ──\n"
        "  TOpening=插门窗 TBanWin=插门 TGroupOpening=组合门窗 TCornerWin=转角窗\n"
        "  TWin2Lib=门窗入库 TStatOp=门窗统计\n"
        "── 楼梯/台阶 ──\n"
        "  TLStair=直梯 TAStair=弧梯 TCStair=电梯 TStep=台阶 TBalcony=阳台\n"
        "── 房间/面积 ──\n"
        "  TSpArea=搜索房间 TApartArea=套内面积 TCountArea=统计面积 TSan=卫生洁具\n"
        "── 屋顶 ──\n"
        "  TRectRoof=矩形屋顶 TSlopeRoof=单坡屋顶 TDualSlopeRoof=双坡屋顶 TDormer=天窗\n"
        "── 标注 ──\n"
        "  TDim3=门窗标注 TDimWall=墙厚标注 TQuickDim=快速标注 TDimMP=两点标注 TCoord=坐标标注\n"
        "── 文字/表格 ──\n"
        "  TText=文字 TMText=多行文字 TWordLib=专业词库 TTextConv=文字转换\n"
        "  Sheet2Excel=转Excel Excel2Sheet=导Excel Sheet2Word=转Word\n\n"
        "## 基础图形函数(已注入, 全部支持 color=颜色常量)\n"
        "FAST_MODE(True)=关闭屏幕刷新加速绘图 FAST_MODE(False)=恢复并刷新\n"
        "SET_COLOR(obj,color)=改对象颜色 颜色常量: RED=1 YELLOW=2 GREEN=3 CYAN=4 BLUE=5 MAGENTA=6 WHITE=7\n"
        "L(x1,y1,x2,y2,color)=直线  PLINE([(x,y),...],color)=多段线  R(x,y,w,h,color)=矩形\n"
        "C(cx,cy,r,color)=圆  ARC(cx,cy,r,start,end,color)=圆弧(角度!)\n"
        "T(text,x,y,h,color)=文字  MTEXT(text,x,y,w,h,color)=多行文字\n"
        "DIM(x1,y1,x2,y2,tx,ty,color)=对齐标注  DIM_H(x1,x2,y,ty,color)=水平标注\n"
        "HATCH(pts,color)=填充  LAYER(name,color)=建/切图层\n"
        "MOVE(obj,dx,dy) ROT(obj,cx,cy,deg) DEL(obj) OBJS() UNDO()\n"
        "FIND_OBJS(type,color,layer)=按条件查找对象 DEL_BY_TYPE('Circle')=批量删指定类型\n"
        "ZOOM_EXT() ZOOM_WIN(x1,y1,x2,y2) SEND_CMD(cmd) BLOCK_INSERT(name,x,y,scale,rot)\n\n"
        "## 示例: 用天正画简单建筑平面\n"
        "LAYER(\"轴线\", 1)  # 红色轴线层\n"
        "SEND_CMD(\"TRectAxis\")  # 用户交互放置轴网\n"
        "LAYER(\"墙\", 7)\n"
        "SEND_CMD(\"TgWall\")  # 用户交互画墙\n"
        "LAYER(\"门窗\", 4)\n"
        "SEND_CMD(\"TOpening\")  # 用户交互插门窗\n"
        "PARA()  # 另起一段生成交互内容\n"
        "## 如果用纯Python绘制(不用天正交互):\n"
        "LAYER(\"墙\", 7)\n"
        "PLINE([(0,0),(12000,0),(12000,8000),(0,8000)])  # 外墙\n"
        "PLINE([(2000,2000),(6000,2000),(6000,6000),(2000,6000)])\n"
        "DIM(0,0,12000,0,6000,-800); ZOOM_EXT()\n\n"
        "## 铁律: 禁止import 禁止Dispatch 只用SEND_CMD+L/C/R等 只输出代码"
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

# ── AutoCAD connection ────────────────────────────────────
class AcadConnection:
    def __init__(self):
        self.acad = None
        self.connected = False
        self.error = ""

    def connect(self, create=False):
        if not HAS_PYAUTOCAD:
            self.error = "pyautocad 未安装"
            return False

        # Reset stale connection
        if self.acad is not None:
            try:
                del self.acad
            except:
                pass
            self.acad = None
        self.connected = False

        import time
        last_err = ""
        for attempt in range(3):
            try:
                self.acad = Autocad(create_if_not_exists=create)
                self.acad.app.Visible = True
                _ = self.acad.doc.Name
                self.connected = True
                self.error = ""
                return True
            except Exception as e:
                last_err = str(e)
                if attempt < 2:
                    time.sleep(1.5)  # Wait for AutoCAD to free up
                    continue

        self.connected = False
        self.error = last_err or "连接失败"
        return False

    def ensure(self):
        if self.connected and self.acad:
            try:
                _ = self.acad.doc.Name
                return
            except:
                self.connected = False
                self.acad = None
        if not self.connect(create=True):
            raise ConnectionError(f"无法连接 AutoCAD\n{self.error}")

# ── UI ────────────────────────────────────────────────────
class TextToCADApp(QtWidgets.QMainWindow):
    def __init__(self, acad_conn):
        super().__init__()
        self.cfg = load_config()
        self.proxies = _resolve_proxies(self.cfg.get("proxies")) if HAS_REQUESTS else None
        self.acad_conn = acad_conn
        self._cancelled = False
        self.client = LLMClient(self.cfg)
        self.pipeline = CodeGenPipeline(self.client)

        self.setWindowTitle("TextToCAD for AutoCAD v3")
        self.setMinimumSize(400, 480)
        self.resize(520, 580)

        # ── menu ──
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("文件")
        act_check = file_menu.addAction("检查 AutoCAD 连接")
        act_check.triggered.connect(self._check_acad)
        act_settings = file_menu.addAction("设置...")
        act_settings.triggered.connect(self._open_config)
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

        title = QtWidgets.QLabel("用自然语言在 AutoCAD 中画 2D 图")
        title.setStyleSheet("font-size: 14px; font-weight: bold; padding: 4px;")
        layout.addWidget(title)

        self.input_edit = QtWidgets.QTextEdit()
        self.input_edit.setPlaceholderText(
            "例如：\n"
            "  画一个 4 米 x 3 米的矩形，左下角在原点\n"
            "  给四条边标注尺寸\n"
            "  在矩形中心画一个半径 500mm 的圆\n"
            "  把矩形往右移动 1 米\n"
            "  旋转圆形 45 度"
        )
        self.input_edit.setMaximumHeight(140)
        layout.addWidget(self.input_edit)

        # buttons
        btn_layout = QtWidgets.QHBoxLayout()
        self.draw_btn = QtWidgets.QPushButton("生成并绘制")
        self.draw_btn.setStyleSheet(
            "QPushButton { background-color: #2196F3; color: white; font-size: 13px;"
            " font-weight: bold; padding: 8px 20px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #1976D2; }"
            "QPushButton:disabled { background-color: #666; }"
        )
        self.draw_btn.clicked.connect(self.on_draw)
        btn_layout.addWidget(self.draw_btn)

        self.cancel_btn = QtWidgets.QPushButton("取消")
        self.cancel_btn.setStyleSheet(
            "QPushButton { background-color: #f44336; color: white; font-size: 13px;"
            " font-weight: bold; padding: 8px 20px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #d32f2f; }"
        )
        self.cancel_btn.clicked.connect(self.on_cancel)
        self.cancel_btn.hide()
        btn_layout.addWidget(self.cancel_btn)

        self.query_btn = QtWidgets.QPushButton("查询图形")
        self.query_btn.setStyleSheet(
            "QPushButton { background-color: #4CAF50; color: white; font-size: 13px;"
            " font-weight: bold; padding: 8px 20px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #388E3C; }"
            "QPushButton:disabled { background-color: #666; }"
        )
        self.query_btn.clicked.connect(self.on_query)
        btn_layout.addWidget(self.query_btn)

        self.clear_btn = QtWidgets.QPushButton("清空画布")
        self.clear_btn.clicked.connect(self.on_clear)
        btn_layout.addWidget(self.clear_btn)

        self.reconnect_btn = QtWidgets.QPushButton("重连AutoCAD")
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

        self.zoom_btn = QtWidgets.QPushButton("全图缩放")
        self.zoom_btn.setToolTip("缩放至显示全部图形")
        self.zoom_btn.setStyleSheet(
            "QPushButton{background:#FF9800;color:white;font-size:11px;font-weight:bold;padding:5px 10px;border-radius:3px}"
            "QPushButton:hover{background:#F57C00}")
        self.zoom_btn.clicked.connect(self.on_zoom_ext)
        tools_layout.addWidget(self.zoom_btn)

        self.layer_btn = QtWidgets.QPushButton("列出图层")
        self.layer_btn.setToolTip("显示当前文档所有图层及其状态")
        self.layer_btn.setStyleSheet(
            "QPushButton{background:#607D8B;color:white;font-size:11px;font-weight:bold;padding:5px 10px;border-radius:3px}"
            "QPushButton:hover{background:#455A64}")
        self.layer_btn.clicked.connect(self.on_list_layers)
        tools_layout.addWidget(self.layer_btn)

        self.clean_btn = QtWidgets.QPushButton("清理空层")
        self.clean_btn.setToolTip("清理所有空图层和未使用块定义（PURGE）")
        self.clean_btn.setStyleSheet(
            "QPushButton{background:#00897B;color:white;font-size:11px;font-weight:bold;padding:5px 10px;border-radius:3px}"
            "QPushButton:hover{background:#00695C}")
        self.clean_btn.clicked.connect(self.on_purge)
        tools_layout.addWidget(self.clean_btn)

        self.obj_btn = QtWidgets.QPushButton("对象统计")
        self.obj_btn.setToolTip("统计模型空间各类对象数量")
        self.obj_btn.setStyleSheet(
            "QPushButton{background:#9C27B0;color:white;font-size:11px;font-weight:bold;padding:5px 10px;border-radius:3px}"
            "QPushButton:hover{background:#7B1FA2}")
        self.obj_btn.clicked.connect(self.on_obj_stats)
        tools_layout.addWidget(self.obj_btn)

        tools_layout.addStretch()
        layout.addLayout(tools_layout)

        # status
        self.status_lbl = QtWidgets.QLabel("就绪 — Ctrl+Enter 快速绘制")
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
        mode_text = f"模式: {self.cfg['provider']} | 模型: {self.cfg['model']} | 单位: {self.cfg['units']}{proxy_info}"
        if self.cfg["provider"] == "bridge":
            mode_text += "\n桥接模式：输入写入文件，等待 AI 响应"
        self.mode_lbl = QtWidgets.QLabel(mode_text)
        self.mode_lbl.setStyleSheet("color: #f0a030; font-size: 10px;")
        self.mode_lbl.setWordWrap(True)
        layout.addWidget(self.mode_lbl)

        # ── status bar ──
        self.acad_status = QtWidgets.QLabel()
        self._update_acad_status()
        self.statusBar().addPermanentWidget(self.acad_status)

        # shortcuts
        QtGui.QShortcut(QtGui.QKeySequence("Ctrl+Return"), self).activated.connect(self.on_draw)

    # ── status bar ────────────────────────────────────
    def _update_acad_status(self):
        if self.acad_conn.connected:
            self.acad_status.setText("AutoCAD:  已连接")
            self.acad_status.setStyleSheet("color: #4caf50; font-weight: bold; padding: 2px 8px;")
            self.reconnect_btn.hide()
        else:
            self.acad_status.setText("AutoCAD:  未连接")
            self.acad_status.setStyleSheet("color: #f0a030; font-weight: bold; padding: 2px 8px;")
            self.reconnect_btn.show()

    def _check_acad(self):
        try:
            self.acad_conn.ensure()
            QtWidgets.QMessageBox.information(
                self, "AutoCAD 连接",
                f"已连接\n文档: {self.acad_conn.acad.doc.Name}"
            )
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "AutoCAD 连接", str(e))
        self._update_acad_status()

    def _reconnect(self):
        self.set_status("正在连接 AutoCAD...", "#2196F3")
        QtCore.QCoreApplication.processEvents()
        ok = self.acad_conn.connect(create=False) or self.acad_conn.connect(create=True)
        self._update_acad_status()
        if ok:
            self.set_status("AutoCAD 已连接", "green")
            doc = self.acad_conn.acad.doc.Name
            self.reconnect_btn.hide()
        else:
            self.set_status(f"连接失败: {self.acad_conn.error}", "red")

    def _open_config(self):
        os.startfile(str(CONFIG_FILE))

    def _about(self):
        QtWidgets.QMessageBox.about(
            self, "关于 TextToCAD",
            "TextToCAD for AutoCAD v2\n\n"
            "用自然语言在 AutoCAD 中生成 2D 图形\n"
            "快捷工具: 全图缩放 | 图层管理 | 清理 | 统计\n"
            "LLM: DeepSeek / OpenAI 兼容 API"
        )

    # ── 快捷工具 ──
    def on_zoom_ext(self):
        """Zoom to extents"""
        try: self.acad_conn.ensure()
        except: pass
        try:
            self.acad_conn.acad.app.ZoomExtents()
            self.set_status("全图缩放完成", "green")
        except Exception as e: self.set_status(f"失败: {e}", "red")

    def on_list_layers(self):
        """List all layers"""
        try: self.acad_conn.ensure()
        except: pass
        try:
            doc = self.acad_conn.acad.app.ActiveDocument
            lines = []
            for i in range(doc.Layers.Count):
                try:
                    layer = doc.Layers.Item(i)
                    name = layer.Name
                    color = layer.Color
                    on_off = "ON" if layer.LayerOn else "OFF"
                    frozen = "冻结" if layer.Freeze else ""
                    lines.append(f"  {name} color={color} {on_off} {frozen}")
                except: continue
            self.output_edit.setText(f"--- 图层列表 ({len(lines)}个) ---\n" + "\n".join(lines))
            self.set_status(f"共{len(lines)}个图层", "green")
        except Exception as e: self.set_status(f"失败: {e}", "red")

    def on_purge(self):
        """Purge empty layers and unused blocks"""
        try: self.acad_conn.ensure()
        except: pass
        try:
            doc = self.acad_conn.acad.app.ActiveDocument
            doc.PurgeAll()
            self.set_status("清理完成", "green")
        except Exception as e: self.set_status(f"清理失败: {e}", "red")

    def on_obj_stats(self):
        """Count model space objects"""
        try: self.acad_conn.ensure()
        except: pass
        try:
            ms = self.acad_conn.acad.model
            stats = {}
            for i in range(ms.Count):
                try:
                    name = ms.Item(i).ObjectName
                    stats[name] = stats.get(name, 0) + 1
                except: continue
            lines = [f"  模型空间共 {ms.Count} 个对象:\n"]
            for k, v in sorted(stats.items(), key=lambda x: -x[1]):
                short = k.replace("AcDb", "").replace("Acad", "")
                lines.append(f"  {short}: {v}")
            self.output_edit.setText("--- 对象统计 ---\n" + "\n".join(lines))
            self.set_status(f"统计完成 ({ms.Count}个对象)", "green")
        except Exception as e: self.set_status(f"失败: {e}", "red")

    # ── LLM ───────────────────────────────────────────
    def _call_llm(self, messages):
        if self.cfg["provider"] == "bridge":
            return self._call_bridge(messages)
        if not HAS_REQUESTS:
            raise RuntimeError("requests 未安装且非桥接模式")

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

    # ── actions ───────────────────────────────────────
    def on_cancel(self):
        self._cancelled = True
        self.set_status("正在取消...", "#f44336")

    def on_clear(self):
        if not self.acad_conn.connected:
            QtWidgets.QMessageBox.warning(self, "未连接", "请先启动 AutoCAD 并确保已连接。")
            return

        reply = QtWidgets.QMessageBox.question(
            self, "确认清空",
            "确定要删除 AutoCAD 当前文档中的所有图形对象吗？\n此操作不可撤销。"
        )
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return

        try:
            count = 0
            for _, obj in self._iter_acad_objects():
                try:
                    obj.Delete()
                    count += 1
                except:
                    pass
            self.acad_conn.acad.Application.ZoomExtents()
            self.set_status(f"已删除 {count} 个对象", "green")
        except Exception as e:
            self.set_status(f"清空失败: {e}", "red")

    def _dwg_overview(self):
        """Read full drawing overview: layers, blocks, layouts + model space objects."""
        if not self.acad_conn.connected:
            return "未连接"
        try:
            doc = self.acad_conn.acad.app.ActiveDocument
            out = [f"图纸: {doc.Name}"]

            # Layers
            try:
                layers = doc.Layers
                layer_info = []
                for i in range(layers.Count):
                    try:
                        la = layers.Item(i)
                        state = []
                        if la.LayerOn:
                            state.append("ON")
                        else:
                            state.append("OFF")
                        if la.Freeze:
                            state.append("Frozen")
                        if la.Lock:
                            state.append("Locked")
                        color = la.TrueColor.ColorIndex if hasattr(la, "TrueColor") else "?"
                        layer_info.append(f"  {la.Name} ({','.join(state)}) color={color}")
                    except:
                        pass
                if layer_info:
                    out.append(f"\n[图层] ({len(layer_info)}个):\n" + "\n".join(layer_info[:20]))
            except:
                pass

            # Blocks
            try:
                blocks = doc.Blocks
                block_info = []
                for i in range(blocks.Count):
                    try:
                        blk = blocks.Item(i)
                        ct = blk.Count
                        if ct > 0 and not blk.Name.startswith("*"):
                            block_info.append(f"  {blk.Name}: {ct}个图元")
                    except:
                        pass
                if block_info:
                    out.append(f"\n[块定义] ({len(block_info)}个):\n" + "\n".join(block_info[:15]))
            except:
                pass

            # Layouts
            try:
                layouts = doc.Layouts
                layout_info = []
                for i in range(layouts.Count):
                    try:
                        lo = layouts.Item(i)
                        is_active = " [当前]" if lo.Name == doc.ActiveLayout.Name else ""
                        layout_info.append(f"  {lo.Name}{is_active} ({lo.Block.Count}个图元)")
                    except:
                        pass
                if layout_info:
                    out.append(f"\n[布局] ({len(layout_info)}个):\n" + "\n".join(layout_info))
            except:
                pass

            # Model space objects (fast summary)
            try:
                ms = doc.ModelSpace
                type_counts = {}
                for i in range(min(ms.Count, 500)):
                    try:
                        oname = str(ms.Item(i).ObjectName)
                        short = oname.replace("AcDb", "")
                        type_counts[short] = type_counts.get(short, 0) + 1
                    except:
                        pass
                if type_counts:
                    summary = ", ".join(f"{k}:{v}" for k, v in sorted(type_counts.items()))
                    out.append(f"\n[模型空间] ({ms.Count}个对象): {summary}")
            except:
                pass

            return "\n".join(out)
        except Exception as e:
            return f"图纸概览失败: {e}"

    def _safe_obj_info(self, obj):
        """Return a one-line description string for any COM object. Never raises."""
        try:
            oname = str(obj.ObjectName)
        except:
            return "?"
        parts = [oname]
        for attr, label in [("Radius", "R"), ("Length", "L"), ("Measurement", "标注"),
                             ("TextString", "文字"), ("Area", "面积")]:
            try:
                v = float(getattr(obj, attr, None))
                parts.append(f"{label}={v:.0f}")
            except:
                pass
        for attr, label in [("Center", "圆心"), ("StartPoint", "起点"), ("EndPoint", "终点"),
                             ("InsertionPoint", "位置"), ("TextPosition", "标注位置")]:
            try:
                pt = getattr(obj, attr, None)
                if pt is not None:
                    x, y = float(pt[0]), float(pt[1])
                    parts.append(f"{label}=({x:.0f},{y:.0f})")
            except:
                pass
        return " ".join(parts)

    def _get_existing_info(self):
        if not self.acad_conn.connected:
            return ""
        try:
            lines = []
            for oname, obj in self._iter_acad_objects():
                if oname == "AcDbBlockReference":
                    continue
                lines.append(f"  {self._safe_obj_info(obj)}")
            if lines:
                return "已有图形：\n" + "\n".join(lines) + "\n\n"
        except:
            pass
        return ""

    def on_query(self):
        question = self.input_edit.toPlainText().strip()
        try:
            self.acad_conn.ensure()
            self._update_acad_status()
        except ConnectionError as e:
            QtWidgets.QMessageBox.warning(self, "AutoCAD 未连接",
                f"请先启动 AutoCAD 再试。\n\n{str(e)[:200]}")
            return

        overview = self._dwg_overview()
        # Also collect detailed object info for deeper analysis
        objects = []
        for oname, obj in self._iter_acad_objects():
            if oname == "AcDbBlockReference":
                continue
            objects.append(self._safe_obj_info(obj))
        obj_text = "\n".join(f"  {o}" for o in objects[:100])

        self.output_edit.setText(f"--- 图纸概览 ---\n{overview}\n\n检测到 {len(objects)} 个对象")
        self.set_status("AI 分析中...", "#2196F3")
        QtCore.QCoreApplication.processEvents()

        self.query_btn.setEnabled(False)
        self.draw_btn.setEnabled(False)
        try:
            user_msg = f"图纸概览:\n{overview}\n\n"
            if objects:
                user_msg += f"模型空间对象细表 (前100个):\n{obj_text}\n\n"
            user_msg += f"用户问题: {question}" if question else "请用中文总结图纸内容、图层结构、块定义。"
            label = "查询结果" if question else "图纸总结"

            messages = [
                {"role": "system", "content": "你是 AutoCAD 图形分析专家。根据图纸概览（图层/块/布局/对象统计）和对象细表回答问题。关注全局结构。"},
                {"role": "user", "content": user_msg},
            ]
            answer = self.client.chat(messages)
            self.output_edit.setText(f"--- {label} ---\n{answer}")
            self.set_status("查询完成", "green")
        except Exception as e:
            self.set_status(f"查询失败: {e}", "red")
        finally:
            self.query_btn.setEnabled(True)
            self.draw_btn.setEnabled(True)

    def on_draw(self):
        user_input = self.input_edit.toPlainText().strip()
        if not user_input:
            self.set_status("请输入描述文字", "red")
            return

        try:
            self.acad_conn.ensure()
            self._update_acad_status()
        except ConnectionError as e:
            QtWidgets.QMessageBox.warning(self, "AutoCAD 未连接",
                f"请先启动 AutoCAD 再试。\n\n{str(e)[:200]}")
            return

        self.draw_btn.setEnabled(False)
        self.cancel_btn.show()
        self.output_edit.clear()
        self._cancelled = False

        overview = self._dwg_overview()
        self.output_edit.setText(f"--- 图纸概览 ---\n{overview}")
        messages = [
            {"role": "system", "content": self.cfg["system_prompt_zh"]},
            {"role": "user", "content": f"图纸概览:\n{overview}\n\n用户描述: {user_input}\n\n理解图纸全局结构后输出Python代码。只输出代码。"},
        ]

        MAX_RETRIES = 6
        fixer_code_acad = None
        code = ""
        try:
            for attempt in range(MAX_RETRIES):
                if self._cancelled:
                    self.set_status("已取消", "#f0a030")
                    return

                tag = "生成..." if attempt == 0 else f"修复第{attempt}轮..."
                self.set_status(tag, "#2196F3")
                QtCore.QCoreApplication.processEvents()

                if fixer_code_acad:
                    code = fixer_code_acad
                    fixer_code_acad = None
                    self.output_edit.setText(f"--- 专家修正版 ---\n{code}")
                else:
                    code = self.client.chat(messages)
                    self.output_edit.setText(f"--- {attempt+1}/{MAX_RETRIES} ---\n{code}")

                # Safe helpers injected into exec namespace
                ac = self.acad_conn.acad
                ms = ac.model  # Use pyautocad's standard model space property

                # Undo: remember object count before execution
                undo_count = ms.Count

                # Color constants
                RED, YELLOW, GREEN, CYAN, BLUE, MAGENTA, WHITE = 1, 2, 3, 4, 5, 6, 7

                def _setc(obj, color):
                    """Set object color if provided."""
                    if color is not None:
                        try:
                            obj.Color = color
                        except:
                            pass

                def L(x1, y1, x2, y2, color=None):
                    """Draw line from (x1,y1) to (x2,y2)."""
                    obj = ms.AddLine(APoint(x1, y1, 0), APoint(x2, y2, 0))
                    _setc(obj, color)
                    return obj

                def C(cx, cy, r, color=None):
                    """Draw circle center (cx,cy) radius r."""
                    obj = ms.AddCircle(APoint(cx, cy, 0), r)
                    _setc(obj, color)
                    return obj

                def R(x, y, w, h, color=None):
                    """Draw rectangle bottom-left (x,y) width w height h."""
                    pl = ms.AddLightWeightPolyline(aDouble(x, y, x + w, y, x + w, y + h, x, y + h))
                    pl.Closed = True
                    _setc(pl, color)
                    return pl

                def T(txt, x, y, h=100, color=None):
                    """Add text at (x,y) with height h."""
                    obj = ms.AddText(str(txt), APoint(x, y, 0), h)
                    _setc(obj, color)
                    return obj

                def DIM(x1, y1, x2, y2, tx, ty, color=None):
                    """Add aligned dimension."""
                    obj = ms.AddDimAligned(APoint(x1, y1, 0), APoint(x2, y2, 0), APoint(tx, ty, 0))
                    _setc(obj, color)
                    return obj

                def SET_COLOR(obj, color):
                    """Change color of an existing object. color: 1红2黄3绿4青5蓝6品红7白"""
                    obj.Color = color

                def MOVE(obj, dx, dy):
                    """Move object by (dx,dy)."""
                    obj.Move(APoint(0, 0, 0), APoint(dx, dy, 0))

                def ROT(obj, cx, cy, deg):
                    """Rotate object around (cx,cy) by deg degrees."""
                    obj.Rotate(APoint(cx, cy, 0), math.radians(deg))

                def DEL(obj):
                    """Safely delete an object."""
                    try:
                        obj.Delete()
                    except:
                        pass

                def OBJS():
                    """Return list of (name, obj) tuples for all model space objects."""
                    result = []
                    for i in range(ms.Count):
                        try:
                            o = ms.Item(i)
                            result.append((str(o.ObjectName), o))
                        except:
                            pass
                    return result

                def FIND_OBJS(type_kw=None, color=None, layer=None):
                    """Find objects by keywords. type_kw matches ObjectName substring, color=1-7, layer='name'.
                    Returns list of (object_name, obj). Example: FIND_OBJS('Circle', color=GREEN) finds green circles."""
                    result = []
                    for i in range(ms.Count):
                        try:
                            o = ms.Item(i)
                            oname = str(o.ObjectName)
                            match = True
                            if type_kw and type_kw.lower() not in oname.lower():
                                match = False
                            if match and color is not None:
                                try:
                                    if o.Color != color:
                                        match = False
                                except:
                                    pass
                            if match and layer:
                                try:
                                    if o.Layer.lower() != layer.lower():
                                        match = False
                                except:
                                    pass
                            if match:
                                result.append((oname, o))
                        except:
                            pass
                    return result

                def DEL_BY_TYPE(type_kw):
                    """Delete all objects matching a type keyword. Example: DEL_BY_TYPE('Circle') deletes all circles."""
                    deleted = 0
                    for i in range(ms.Count - 1, -1, -1):
                        try:
                            o = ms.Item(i)
                            if type_kw.lower() in str(o.ObjectName).lower():
                                o.Delete()
                                deleted += 1
                        except:
                            pass
                    return deleted

                def UNDO():
                    """Undo: delete objects created since this execution began."""
                    while ms.Count > undo_count:
                        try:
                            ms.Item(ms.Count - 1).Delete()
                        except:
                            break

                # ── performance helpers ──
                def FAST_MODE(on=True):
                    """Disable screen refresh for fast batch drawing. Call FAST_MODE(False) after done."""
                    try:
                        doc = ac.app.ActiveDocument
                        if on:
                            doc.SetVariable("CMDECHO", 0)
                        else:
                            doc.SetVariable("CMDECHO", 1)
                            doc.Regen(1)
                    except:
                        pass

                # ── command sender (for TArch/源泉/all CAD commands) ──
                def SEND_CMD(cmd):
                    """Send a raw command string to AutoCAD command line."""
                    doc = ac.app.ActiveDocument
                    doc.SendCommand(cmd + "\n")

                # ── additional helpers ──
                def PLINE(pts, color=None):
                    """Polyline from list of (x,y) tuples."""
                    flat = []
                    for p in pts:
                        flat.extend([p[0], p[1]])
                    pl = ms.AddLightWeightPolyline(aDouble(*flat))
                    _setc(pl, color)
                    return pl

                def ARC(cx, cy, r, start_deg, end_deg, color=None):
                    """Arc: center(cx,cy), radius r, start_deg to end_deg in DEGREES."""
                    import math as _m
                    sa = _m.radians(start_deg)
                    ea = _m.radians(end_deg)
                    obj = ms.AddArc(APoint(cx, cy, 0), r, sa, ea)
                    _setc(obj, color)
                    return obj

                def HATCH(pts, color=3):
                    """Hatch region defined by list of (x,y) points."""
                    flat = []
                    for p in pts:
                        flat.extend([p[0], p[1]])
                    pl = ms.AddLightWeightPolyline(aDouble(*flat))
                    pl.Closed = True
                    try:
                        hatch = ms.AddHatch(0, "SOLID", True)
                        hatch.AppendOuterLoop([pl])
                        hatch.Color = color
                    except:
                        pass
                    return pl

                def MTEXT(text, x, y, w=2000, h=400, color=None):
                    """Multiline text in a bounding box."""
                    obj = ms.AddMText(APoint(x, y, 0), w, str(text))
                    _setc(obj, color)
                    return obj

                def DIM_H(x1, x2, y, ty):
                    """Horizontal dimension."""
                    return ms.AddDimRotated(APoint(x1, y, 0), APoint(x2, y, 0), APoint((x1+x2)/2, ty, 0), 0)

                def LAYER(name, color=7, linetype="Continuous"):
                    """Create or switch to a layer."""
                    try:
                        doc = ac.app.ActiveDocument
                        layers = doc.Layers
                        try:
                            layer = layers.Item(name)
                        except:
                            layer = layers.Add(name)
                        layer.Color = color
                        if linetype != "Continuous":
                            try:
                                layer.Linetype = linetype
                            except:
                                pass
                        return layer
                    except:
                        pass

                def DIM_STYLE(name, arrow_size=200, text_size=300):
                    """Create/set dimension style."""
                    try:
                        styles = doc.DimStyles
                        try:
                            style = styles.Item(name)
                        except:
                            style = styles.Add(name)
                        doc.ActiveDimStyle = style
                    except:
                        pass

                def ZOOM_EXT():
                    """Zoom to extents."""
                    try:
                        ac.app.ZoomExtents()
                    except:
                        pass

                def ZOOM_WIN(x1, y1, x2, y2):
                    """Zoom to window."""
                    try:
                        ac.app.ZoomWindow(APoint(x1, y1, 0), APoint(x2, y2, 0))
                    except:
                        pass

                def BLOCK_INSERT(name, x, y, scale=1.0, rot=0):
                    """Insert block reference."""
                    try:
                        return ms.InsertBlock(APoint(x, y, 0), name, scale, scale, scale, math.radians(rot))
                    except:
                        return None

                ns = {
                    "acad": ac, "APoint": APoint, "aDouble": aDouble,
                    "math": __import__("math"),
                    "SEND_CMD": SEND_CMD, "SET_COLOR": SET_COLOR, "FAST_MODE": FAST_MODE,
                    "RED": RED, "YELLOW": YELLOW, "GREEN": GREEN,
                    "CYAN": CYAN, "BLUE": BLUE, "MAGENTA": MAGENTA, "WHITE": WHITE,
                    "L": L, "C": C, "R": R, "T": T, "DIM": DIM,
                    "MOVE": MOVE, "ROT": ROT, "DEL": DEL, "OBJS": OBJS,
                    "FIND_OBJS": FIND_OBJS, "DEL_BY_TYPE": DEL_BY_TYPE,
                    "UNDO": UNDO, "PLINE": PLINE, "ARC": ARC, "HATCH": HATCH,
                    "MTEXT": MTEXT, "DIM_H": DIM_H, "LAYER": LAYER,
                    "DIM_STYLE": DIM_STYLE, "ZOOM_EXT": ZOOM_EXT,
                    "ZOOM_WIN": ZOOM_WIN, "BLOCK_INSERT": BLOCK_INSERT,
                }
                try:
                    exec(code, ns)
                    self.acad_conn.acad.Application.ZoomExtents()
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
                        fixed = fix_code(ACAD_FIXER_PROMPT, short, code, search_hint, self.cfg, self.proxies)
                        if fixed:
                            fixer_code_acad = fixed
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
            self.draw_btn.setEnabled(True)
            self.cancel_btn.hide()

    def set_status(self, text, color="#888"):
        self.status_lbl.setText(text)
        self.status_lbl.setStyleSheet(f"color: {color}; font-size: 11px; padding: 4px;")

# ── entry ─────────────────────────────────────────────────
def run_app():
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")

    # COM must be initialized AFTER QApplication (Qt calls OleInitialize)
    # Use win32com which handles COM apartment correctly under Qt
    acad_conn = AcadConnection()
    acad_conn.connect(create=False)
    if not acad_conn.connected:
        acad_conn.connect(create=True)

    window = TextToCADApp(acad_conn)
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
