import sys, os, psutil, gc, time, subprocess
import pyqtgraph as pg
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *


def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(__file__), relative_path)


class WorkThread(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str, str)

    def __init__(self, t, p, pa=None):
        super().__init__()
        self.t = t;
        self.p = p;
        self.pa = pa
        self._is_active = True

    def run(self):
        try:
            if self.t == "BG":
                from rembg import remove, new_session
                os.environ["U2NET_HOME"] = get_resource_path("models")
                self.progress.emit("AI: 1/3 (Loading/加载中)")
                session = new_session("u2net")
                self.progress.emit("AI: 2/3 (Computing/处理中)")
                with open(self.p, "rb") as i:
                    res = remove(i.read(), session=session)
                out = self.p.rsplit(".", 1)[0] + "_nobg.png"
                with open(out, "wb") as f:
                    f.write(res)
                self.finished.emit(out)

            elif self.t == "PDF":
                from pdf2docx import Converter
                out = self.p.rsplit(".", 1)[0] + ".docx"
                self.progress.emit("PDF: 1/2 (Converting/转换中...)")
                cv = Converter(self.p)
                cv.convert(out, start=0, end=None, multi_processing=False)
                cv.close()
                self.finished.emit(out)

            elif self.t == "GIF":
                # --- 硬核修复：绕过 imageio 库，直接调用 ffmpeg.exe ---
                ffmpeg_bin = get_resource_path(os.path.join("bin", "ffmpeg.exe"))
                out = self.p.rsplit(".", 1)[0] + ".gif"
                self.progress.emit("GIF: Converting/转换中...")

                # 构建命令行：截取前 pa 秒，10fps，缩放到 480 宽(保证速度)
                cmd = [
                    ffmpeg_bin, "-y", "-ss", "0", "-t", str(self.pa),
                    "-i", self.p, "-vf", "fps=10,scale=480:-1", out
                ]

                # 隐藏控制台窗口执行
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

                process = subprocess.Popen(cmd, startupinfo=startupinfo, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                process.communicate()

                if os.path.exists(out):
                    self.finished.emit(out)
                else:
                    raise Exception("FFmpeg conversion failed/转换失败")

            gc.collect()
        except Exception as e:
            self.error.emit("Error/错误", str(e))

    def stop(self):
        self._is_active = False


class NanoDash(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(180, 240)
        self.worker = None
        self.init_ui()
        self.last_net = psutil.net_io_counters()
        self.timer = QTimer();
        self.timer.timeout.connect(self.refresh);
        self.timer.start(1000)

    def init_ui(self):
        self.panel = QFrame(self);
        self.panel.setGeometry(0, 0, 180, 240);
        self.panel.setObjectName("MainPanel")
        self.panel.setStyleSheet("""
            QFrame#MainPanel { background-color: rgba(30, 30, 35, 215); border: 1px solid rgba(255, 255, 255, 15); border-radius: 12px; }
            QLabel { color: #888; font-family: 'Segoe UI', 'Microsoft YaHei'; font-size: 10px; }
            QProgressBar { border: none; background: rgba(255,255,255,10); height: 2px; }
            QProgressBar::chunk { background: #5b9bd5; }
            QPushButton { background: rgba(255, 255, 255, 8); color: #ccc; border: none; border-radius: 4px; padding: 6px; font-size: 10px; }
            QPushButton:hover { background: #5b9bd5; color: #fff; }
        """)
        layout = QVBoxLayout(self.panel)
        header = QHBoxLayout()
        title = QLabel("NANO DASH");
        title.setStyleSheet("font-weight: bold; color: #5b9bd5;")
        btn_sw = QPushButton("⚙");
        btn_sw.setFixedSize(22, 22);
        btn_sw.clicked.connect(self.toggle_page)
        btn_ex = QPushButton("✕");
        btn_ex.setFixedSize(22, 22);
        btn_ex.clicked.connect(self.close)
        header.addWidget(title);
        header.addStretch();
        header.addWidget(btn_sw);
        header.addWidget(btn_ex)
        layout.addLayout(header)

        self.stack = QStackedWidget();
        layout.addWidget(self.stack)
        self.page_mon = QWidget();
        m_lyt = QVBoxLayout(self.page_mon)
        self.cpu_bar = self.add_stat("CPU Usage/占用", m_lyt);
        self.mem_bar = self.add_stat("MEM Usage/内存", m_lyt)
        self.net_lbl = QLabel("DL/下载: 0 KB/s");
        m_lyt.addWidget(self.net_lbl)
        self.plot = pg.PlotWidget();
        self.plot.setBackground(None);
        self.plot.setFixedHeight(35)
        self.plot.setMouseEnabled(x=False, y=False);
        self.plot.hideAxis('bottom');
        self.plot.hideAxis('left')
        self.curve = self.plot.plot(pen=pg.mkPen('#5b9bd5', width=1.2))
        self.net_data = [0] * 30;
        m_lyt.addWidget(self.plot);
        self.stack.addWidget(self.page_mon)

        self.page_tools = QWidget();
        t_lyt = QVBoxLayout(self.page_tools);
        t_lyt.setSpacing(6)
        btns = [("AI BG Remover/AI抠图", self.start_bg), ("PDF to Word/PDF转Word", self.start_pdf),
                ("Video to GIF/视频转GIF", self.start_gif)]
        for text, func in btns:
            b = QPushButton(text);
            b.clicked.connect(func);
            t_lyt.addWidget(b)
        self.stack.addWidget(self.page_tools)

        self.status = QLabel("Ready/就绪");
        self.status.setAlignment(Qt.AlignmentFlag.AlignCenter);
        layout.addWidget(self.status)

    def add_stat(self, n, l):
        l.addWidget(QLabel(n));
        b = QProgressBar();
        b.setRange(0, 100);
        l.addWidget(b);
        return b

    def toggle_page(self):
        self.stack.setCurrentIndex(1 - self.stack.currentIndex())

    def refresh(self):
        if self.stack.currentIndex() == 0:
            self.cpu_bar.setValue(int(psutil.cpu_percent()))
            self.mem_bar.setValue(int(psutil.virtual_memory().percent))
            curr = psutil.net_io_counters();
            dl = (curr.bytes_recv - self.last_net.bytes_recv) / 1024
            self.last_net = curr;
            self.net_lbl.setText(f"DL/下载: {dl:.1f} KB/s")
            self.net_data.pop(0);
            self.net_data.append(dl);
            self.curve.setData(self.net_data)

    def execute_task(self, t, p, pa=None):
        if self.worker and self.worker.isRunning(): return
        self.status.setText("Processing/处理中...")
        self.worker = WorkThread(t, p, pa)
        self.worker.progress.connect(self.status.setText)
        self.worker.finished.connect(lambda: self.status.setText("Done/转换完成！"))
        self.worker.error.connect(lambda t, c: QMessageBox.critical(self, t, c))
        self.worker.start()

    def start_bg(self):
        p, _ = QFileDialog.getOpenFileName(self, "Select Image/选择图片");
        if p: self.execute_task("BG", p)

    def start_pdf(self):
        p, _ = QFileDialog.getOpenFileName(self, "Select PDF/请选择PDF");
        if p: self.execute_task("PDF", p)

    def start_gif(self):
        p, _ = QFileDialog.getOpenFileName(self, "Select Video/请选择视频")
        if p:
            sec, ok = QInputDialog.getInt(self, "GIF Set/设置", "Duration/时长(1-10s):", 3, 1, 10)
            if ok: self.execute_task("GIF", p, sec)

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton: self.m_p = e.globalPosition().toPoint()

    def mouseMoveEvent(self, e):
        if hasattr(self, 'm_p'):
            self.move(self.pos() + e.globalPosition().toPoint() - self.m_p)
            self.m_p = e.globalPosition().toPoint()

    def closeEvent(self, event):
        if self.worker: self.worker.stop()
        self.timer.stop()
        os._exit(0)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = NanoDash();
    window.show()

    sys.exit(app.exec())

