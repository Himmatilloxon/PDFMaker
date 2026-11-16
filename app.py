"""Desktop app for convert.py: drop files, arrange pages, tweak, stamp, save a flattened PDF."""

import json
import shutil
import sys
from pathlib import Path

from PySide6.QtCore import QRectF, QSize, QSizeF, QStandardPaths, Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import (
    QAction, QColor, QDesktopServices, QIcon, QImage, QKeySequence, QPainter,
    QPalette, QPen, QPixmap,
)
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFormLayout, QGroupBox,
    QHBoxLayout, QInputDialog, QLabel, QLineEdit, QListWidget, QListWidgetItem,
    QMainWindow, QMessageBox, QProgressDialog, QPushButton, QSizePolicy, QSlider,
    QSplitter, QTabWidget, QVBoxLayout, QWidget,
)

import convert as core

EXTENSIONS = sorted(core.WORD_EXT | core.IMAGE_EXT | {".pdf"})
FILE_FILTER = f"Documents and images ({' '.join('*' + e for e in EXTENSIONS)})"
IMAGE_FILTER = f"Images ({' '.join('*' + e for e in sorted(core.IMAGE_EXT))})"
PRESETS = {  # label -> (dpi, JPEG quality; 0 = lossless PNG)
    "Small file (150 DPI)": (150, 75),
    "Balanced (200 DPI)": (200, 88),
    "High quality (300 DPI)": (300, 95),
    "Lossless (300 DPI, large file)": (300, 0),
}
MODES = {"Color": "color", "Grayscale": "gray", "Black and white": "bw"}
THUMB_DPI, PREVIEW_DPI = 20, 110
EMPTY_MESSAGE = "Drop Word, PDF or image files here\n\nor click  Add files"
STAMP_HINT = "Stamp: drag to move · corner handle or mouse wheel to resize · arrow keys nudge · Del removes"

BG, PANEL, CARD, LINE, TEXT, MUTED, ACCENT, ACCENT_HI = (
    "#12141a", "#191c24", "#20242e", "#2e3442", "#e9ebf2", "#98a1b3", "#7c5cff", "#9a82ff")
STYLE = f"""
* {{ font-family: "Segoe UI"; font-size: 10pt; }}
QMainWindow, QWidget#panel {{ background: {BG}; }}
QToolTip {{ background: {CARD}; color: {TEXT}; border: 1px solid {LINE}; padding: 4px 8px; }}
QToolBar {{ background: {PANEL}; border: none; border-bottom: 1px solid {LINE}; padding: 8px 10px; spacing: 4px; }}
QToolBar QLabel#brand {{ color: {TEXT}; font-size: 13pt; font-weight: 700; padding: 0 14px 0 4px; }}
QToolBar::separator {{ background: {LINE}; width: 1px; margin: 6px 8px; }}
QToolButton {{ color: {TEXT}; background: transparent; border: 1px solid transparent; border-radius: 8px; padding: 7px 12px; }}
QToolButton:hover {{ background: {CARD}; border-color: {LINE}; }}
QToolButton:pressed {{ background: {LINE}; }}
QToolButton:focus {{ border-color: {ACCENT}; }}
QToolButton#primary {{ background: {ACCENT}; color: white; font-weight: 600; padding: 8px 22px; }}
QToolButton#primary:hover {{ background: {ACCENT_HI}; }}
QListWidget {{ background: {PANEL}; border: 1px solid {LINE}; border-radius: 10px; padding: 6px; outline: none; }}
QListWidget::item {{ color: {TEXT}; border-radius: 8px; padding: 6px; margin: 2px 0; border: 1px solid transparent; }}
QListWidget::item:hover {{ background: {CARD}; }}
QListWidget::item:selected {{ background: #2b2552; border-color: {ACCENT}; color: white; }}
QTabWidget::pane {{ border: none; }}
QTabBar::tab {{ color: {MUTED}; background: transparent; padding: 8px 18px; margin-right: 4px;
    border-bottom: 2px solid transparent; font-weight: 600; }}
QTabBar::tab:hover {{ color: {TEXT}; }}
QTabBar::tab:selected {{ color: {TEXT}; border-bottom-color: {ACCENT}; }}
QGroupBox {{ background: {PANEL}; border: 1px solid {LINE}; border-radius: 10px; margin-top: 22px; padding: 12px 10px 10px; }}
QGroupBox::title {{ subcontrol-origin: margin; left: 4px; top: 2px; color: {MUTED}; font-weight: 600; }}
QLabel {{ color: {TEXT}; }}
QLabel#value {{ color: {MUTED}; }}
QLineEdit, QComboBox {{ color: {TEXT}; background: {CARD}; border: 1px solid {LINE}; border-radius: 7px; padding: 6px 8px; }}
QLineEdit:focus, QComboBox:focus {{ border-color: {ACCENT}; }}
QComboBox:hover, QLineEdit:hover {{ border-color: {MUTED}; }}
QComboBox QAbstractItemView {{ background: {CARD}; color: {TEXT}; border: 1px solid {LINE}; selection-background-color: {ACCENT}; }}
QPushButton {{ color: {TEXT}; background: {CARD}; border: 1px solid {LINE}; border-radius: 7px; padding: 7px 12px; }}
QPushButton:hover {{ border-color: {ACCENT}; }}
QPushButton:pressed {{ background: {LINE}; }}
QPushButton:focus {{ border-color: {ACCENT}; }}
QPushButton#accent {{ background: {ACCENT}; border-color: {ACCENT}; color: white; font-weight: 600; }}
QPushButton#accent:hover {{ background: {ACCENT_HI}; }}
QCheckBox {{ color: {TEXT}; spacing: 8px; }}
QCheckBox::indicator {{ width: 16px; height: 16px; border-radius: 5px; border: 1px solid {MUTED}; background: {CARD}; }}
QCheckBox::indicator:checked {{ background: {ACCENT}; border-color: {ACCENT_HI}; }}
QSlider::groove:horizontal {{ height: 4px; background: {LINE}; border-radius: 2px; }}
QSlider::sub-page:horizontal {{ background: {ACCENT}; border-radius: 2px; }}
QSlider::handle:horizontal {{ width: 16px; height: 16px; margin: -6px 0; border-radius: 8px; background: white; border: 2px solid {ACCENT}; }}
QSlider::handle:horizontal:hover {{ background: {ACCENT_HI}; }}
QSplitter::handle {{ background: {BG}; width: 8px; }}
QStatusBar {{ background: {PANEL}; color: {MUTED}; border-top: 1px solid {LINE}; }}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: {LINE}; border-radius: 4px; min-height: 30px; }}
QScrollBar::handle:vertical:hover {{ background: {MUTED}; }}
QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; }}
QProgressBar {{ background: {CARD}; border: 1px solid {LINE}; border-radius: 6px; text-align: center; color: {TEXT}; }}
QProgressBar::chunk {{ background: {ACCENT}; border-radius: 5px; }}
QDialog, QMessageBox {{ background: {PANEL}; }}
"""


def to_pixmap(img) -> QPixmap:
    img = img.convert("RGB")
    data = img.tobytes()
    return QPixmap.fromImage(QImage(data, img.width, img.height, img.width * 3, QImage.Format.Format_RGB888))


def dropped_files(event) -> list[str]:
    return [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]


class Cancelled(Exception):
    pass


class Worker(QThread):
    """Runs fn(tick) off the UI thread; tick(i, n) reports progress and honours Cancel."""

    progress = Signal(int, int)
    done = Signal(object)  # None when cancelled
    failed = Signal(str)

    def __init__(self, fn, parent):
        super().__init__(parent)
        self.fn = fn

    def tick(self, i, n):
        if self.isInterruptionRequested():
            raise Cancelled
        self.progress.emit(i, n)

    def run(self):
        try:
            self.done.emit(self.fn(self.tick))
        except Cancelled:
            self.done.emit(None)
        except Exception as e:
            self.failed.emit(str(e))


class PageList(QListWidget):
    """Page thumbnails: drag to reorder, drop files from Explorer to add."""

    files_dropped = Signal(list)

    def __init__(self):
        super().__init__()
        self.setIconSize(QSize(96, 132))
        self.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        self.setAccessibleName("Pages")
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)  # elide long names instead

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            self.files_dropped.emit(dropped_files(event))
            event.acceptProposedAction()
        else:
            super().dropEvent(event)


class StampLibrary(QGroupBox):
    """Stamp images kept in Documents\\PDFMaker Stamps, one sub-folder per category.

    sizes.json in that folder remembers the last size each stamp was used at.
    """

    place = Signal(str)  # path of the stamp to put on the current page

    def __init__(self):
        super().__init__("Stamp library (place by hand)")
        docs = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
        self.root = Path(docs) / "PDFMaker Stamps"
        (self.root / "General").mkdir(parents=True, exist_ok=True)
        self.sizes_file = self.root / "sizes.json"
        try:
            self.sizes = json.loads(self.sizes_file.read_text("utf-8"))
        except (OSError, ValueError):
            self.sizes = {}

        self.folders = QComboBox()
        self.folders.setAccessibleName("Stamp folder")
        self.folders.currentIndexChanged.connect(self.reload_stamps)
        self.stamps = QListWidget()
        self.stamps.setIconSize(QSize(56, 56))
        self.stamps.setDragDropMode(QListWidget.DragDropMode.DragOnly)
        self.stamps.setMinimumHeight(170)
        self.stamps.setAccessibleName("Stamps in this folder")
        self.stamps.setToolTip("Drag a stamp onto the page, or double-click it")
        self.stamps.itemActivated.connect(lambda item: self.place.emit(item.data(Qt.ItemDataRole.UserRole)))

        def button(text, slot, tip, name=""):
            b = QPushButton(text)
            b.clicked.connect(slot)
            b.setToolTip(tip)
            b.setObjectName(name)
            return b

        top = QHBoxLayout()
        top.addWidget(self.folders, 1)
        top.addWidget(button("New folder", self.new_folder, "Create a folder for a type of stamp"))
        bottom = QHBoxLayout()
        bottom.addWidget(button("Add images…", self.add_images, "Copy stamp images into this folder"))
        bottom.addWidget(button("Open folder", lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.folder()))),
                                "Rename, move or delete stamps in Explorer"))
        bottom.addWidget(button("Refresh", self.reload_folders, "Reload after changing files in Explorer"))
        layout = QVBoxLayout(self)
        layout.addLayout(top)
        layout.addWidget(self.stamps, 1)
        layout.addWidget(button("Place on page", self.place_current, "Put the selected stamp on the page you are viewing", "accent"))
        layout.addLayout(bottom)
        self.reload_folders()

    def folder(self) -> Path:
        return self.root / self.folders.currentText()

    def reload_folders(self, *_):
        current = self.folders.currentText()
        self.folders.blockSignals(True)
        self.folders.clear()
        self.folders.addItems(sorted(d.name for d in self.root.iterdir() if d.is_dir()))
        self.folders.setCurrentIndex(max(self.folders.findText(current), 0))
        self.folders.blockSignals(False)
        self.reload_stamps()

    def reload_stamps(self, *_):
        self.stamps.clear()
        if not self.folder().is_dir():
            return
        for f in sorted(self.folder().iterdir()):
            if f.suffix.lower() in core.IMAGE_EXT:
                tile = QPixmap(112, 112)  # white tile: dark ink on a transparent stamp is invisible on the dark list
                tile.fill(QColor("white"))
                art = QPixmap(str(f)).scaled(100, 100, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                painter = QPainter(tile)
                painter.drawPixmap((112 - art.width()) // 2, (112 - art.height()) // 2, art)
                painter.end()
                item = QListWidgetItem(QIcon(tile), f.stem)
                item.setData(Qt.ItemDataRole.UserRole, str(f))
                self.stamps.addItem(item)

    def new_folder(self):
        name, ok = QInputDialog.getText(self, "New stamp folder", "Folder name (e.g. Signatures, Seals):")
        name = name.strip().strip(".")
        if not ok or not name:
            return
        if any(c in name for c in '\\/:*?"<>|'):
            QMessageBox.warning(self, "PDF Maker", 'A folder name cannot contain  \\ / : * ? " < > |')
            return
        (self.root / name).mkdir(exist_ok=True)
        self.reload_folders()
        self.folders.setCurrentText(name)

    def add_images(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Add stamp images", "", IMAGE_FILTER)
        for f in map(Path, files):
            target, n = self.folder() / f.name, 2
            while target.exists():  # never overwrite a stamp already in the library
                target, n = self.folder() / f"{f.stem} {n}{f.suffix}", n + 1
            shutil.copy2(f, target)
        self.reload_stamps()

    def place_current(self):
        item = self.stamps.currentItem()
        if item:
            self.place.emit(item.data(Qt.ItemDataRole.UserRole))

    def key(self, path) -> str:
        return Path(path).relative_to(self.root).as_posix()

    def size_of(self, path) -> float:
        return self.sizes.get(self.key(path), 0.25)

    def remember_size(self, stamp):
        self.sizes[self.key(stamp.path)] = round(stamp.w, 4)
        self.sizes_file.write_text(json.dumps(self.sizes, indent=1), "utf-8")


class PageView(QWidget):
    """Page preview with hand-placed stamps drawn on top: drag to move, handle to resize."""

    files_dropped = Signal(list)
    stamp_dropped = Signal(float, float)  # drop point as fractions of the page
    stamp_resized = Signal(object)
    HANDLE = 14

    def __init__(self, stamp_source):
        super().__init__()
        self.stamp_source = stamp_source  # the library list; drags from it are stamps
        self.pix = self.page = self.selected = self.drag = None
        self.message = ""
        self.hot = False  # something is being dragged over us
        self.cache = {}
        self.setAcceptDrops(True)
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setMinimumSize(320, 320)
        self.setAccessibleName("Page preview")

    def show_page(self, pix, page, message=""):
        if page is not self.page:
            self.selected = None
        self.pix, self.page, self.message = pix, page, message
        self.update()

    # ---------- geometry ----------

    def page_rect(self) -> QRectF:
        area = QRectF(self.rect()).adjusted(24, 24, -24, -24)
        size = QSizeF(self.pix.size()).scaled(area.size(), Qt.AspectRatioMode.KeepAspectRatio)
        return QRectF(area.center().x() - size.width() / 2, area.center().y() - size.height() / 2,
                      size.width(), size.height())

    def stamp_pixmap(self, path) -> QPixmap:
        if path not in self.cache:
            self.cache[path] = QPixmap(path)
        return self.cache[path]

    def stamp_rect(self, st) -> QRectF:
        r, pm = self.page_rect(), self.stamp_pixmap(st.path)
        w = st.w * r.width()
        h = w * pm.height() / pm.width() if not pm.isNull() else w
        return QRectF(r.x() + st.x * r.width(), r.y() + st.y * r.height(), w, h)

    def handle_rect(self, st) -> QRectF:
        c, h = self.stamp_rect(st).bottomRight(), self.HANDLE
        return QRectF(c.x() - h / 2, c.y() - h / 2, h, h)

    def stamp_at(self, pos):
        return next((st for st in reversed(self.page.stamps) if self.stamp_rect(st).contains(pos)), None)

    # ---------- editing ----------

    def place(self, st, x, y, w=None):
        """Move/resize a stamp, keeping it inside the page."""
        if w is not None:
            st.w = min(max(w, 0.02), 1.0)
        height = self.stamp_rect(st).height() / self.page_rect().height()
        st.x = min(max(x, 0.0), max(0.0, 1 - st.w))
        st.y = min(max(y, 0.0), max(0.0, 1 - height))
        self.update()

    def add_stamp(self, path, w, cx=0.5, cy=0.5):
        st = core.Stamp(path, 0, 0, w)
        self.page.stamps.append(st)
        height = self.stamp_rect(st).height() / self.page_rect().height()
        self.place(st, cx - w / 2, cy - height / 2)
        self.selected = st
        self.setFocus()

    def remove_selected(self):
        self.page.stamps.remove(self.selected)
        self.selected = None
        self.update()

    def mousePressEvent(self, event):
        if not self.pix or event.button() != Qt.MouseButton.LeftButton:
            return
        pos = event.position()
        if self.selected and self.handle_rect(self.selected).contains(pos):
            self.drag = ("resize", None)
        else:
            self.selected = self.stamp_at(pos)
            self.drag = ("move", pos - self.stamp_rect(self.selected).topLeft()) if self.selected else None
        self.update()

    def mouseMoveEvent(self, event):
        if not self.pix:
            return
        pos = event.position()
        if not self.drag:
            if self.selected and self.handle_rect(self.selected).contains(pos):
                self.setCursor(Qt.CursorShape.SizeFDiagCursor)
            elif self.stamp_at(pos):
                self.setCursor(Qt.CursorShape.SizeAllCursor)
            else:
                self.unsetCursor()
            return
        r, st = self.page_rect(), self.selected
        if self.drag[0] == "move":
            corner = pos - self.drag[1]
            self.place(st, (corner.x() - r.x()) / r.width(), (corner.y() - r.y()) / r.height())
        else:
            self.place(st, st.x, st.y, (pos.x() - self.stamp_rect(st).left()) / r.width())

    def mouseReleaseEvent(self, event):
        if self.drag and self.drag[0] == "resize":
            self.stamp_resized.emit(self.selected)
        self.drag = None

    def wheelEvent(self, event):
        if not self.selected:
            return super().wheelEvent(event)
        st = self.selected
        self.place(st, st.x, st.y, st.w * 1.05 ** (event.angleDelta().y() / 120))
        self.stamp_resized.emit(st)

    def keyPressEvent(self, event):
        st, step = self.selected, 0.02 if event.modifiers() & Qt.KeyboardModifier.ShiftModifier else 0.004
        moves = {Qt.Key.Key_Left: (-step, 0), Qt.Key.Key_Right: (step, 0), Qt.Key.Key_Up: (0, -step), Qt.Key.Key_Down: (0, step)}
        if st and event.key() in moves:
            dx, dy = moves[event.key()]
            self.place(st, st.x + dx, st.y + dy)
        elif st and event.key() in (Qt.Key.Key_Plus, Qt.Key.Key_Equal, Qt.Key.Key_Minus):
            self.place(st, st.x, st.y, st.w * (0.95 if event.key() == Qt.Key.Key_Minus else 1.05))
            self.stamp_resized.emit(st)
        else:
            super().keyPressEvent(event)

    # ---------- drag and drop ----------

    def accepts(self, event) -> bool:
        return event.mimeData().hasUrls() or (self.pix is not None and event.source() is self.stamp_source)

    def dragEnterEvent(self, event):
        if self.accepts(event):
            self.hot = True
            self.update()
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if self.accepts(event):
            event.acceptProposedAction()

    def dragLeaveEvent(self, event):
        self.hot = False
        self.update()

    def dropEvent(self, event):
        self.hot = False
        self.update()
        if event.mimeData().hasUrls():
            self.files_dropped.emit(dropped_files(event))
        else:
            r, pos = self.page_rect(), event.position()
            self.stamp_dropped.emit((pos.x() - r.x()) / r.width(), (pos.y() - r.y()) / r.height())
        event.acceptProposedAction()

    # ---------- painting ----------

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHints(QPainter.RenderHint.Antialiasing | QPainter.RenderHint.SmoothPixmapTransform)
        accent = QColor(ACCENT_HI)
        if not self.pix:  # empty state doubles as the drop zone
            zone = QRectF(self.rect()).adjusted(28, 28, -28, -28)
            p.setPen(QPen(accent if self.hot else QColor(LINE), 2, Qt.PenStyle.DashLine))
            p.setBrush(QColor(124, 92, 255, 30) if self.hot else QColor(255, 255, 255, 5))
            p.drawRoundedRect(zone, 18, 18)
            font = p.font()
            font.setPointSize(13)
            p.setFont(font)
            p.setPen(QColor(TEXT if self.hot else MUTED))
            p.drawText(zone, Qt.AlignmentFlag.AlignCenter, self.message)
            return
        r = self.page_rect()
        p.setPen(Qt.PenStyle.NoPen)
        for i in range(1, 9):  # soft drop shadow under the page
            p.fillRect(r.adjusted(-i, -i + 3, i, i + 3), QColor(0, 0, 0, 14))
        p.drawPixmap(r, self.pix, QRectF(self.pix.rect()))
        for st in self.page.stamps:
            sr, pm = self.stamp_rect(st), self.stamp_pixmap(st.path)
            if pm.isNull():
                p.fillRect(sr, QColor(220, 60, 60, 120))  # image file went missing
            else:
                p.drawPixmap(sr, pm, QRectF(pm.rect()))
        if self.selected:
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.setPen(QPen(accent, 1.5, Qt.PenStyle.DashLine))
            p.drawRect(self.stamp_rect(self.selected))
            p.setPen(QPen(QColor("white"), 1.5))
            p.setBrush(accent)
            p.drawEllipse(self.handle_rect(self.selected))
        if self.hot:
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.setPen(QPen(accent, 3))
            p.drawRect(r)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Maker")
        self.resize(1240, 800)
        self.setAcceptDrops(True)
        self.worker = None
        self.stamp_path = ""

        self.pages = PageList()
        self.pages.files_dropped.connect(self.add_files)
        self.pages.currentItemChanged.connect(lambda *_: self.refresh_preview())
        self.pages.model().rowsMoved.connect(lambda *_: self.update_status())
        self.pages.model().rowsRemoved.connect(lambda *_: (self.update_status(), self.refresh_preview()))

        self.library = StampLibrary()
        self.library.place.connect(self.place_stamp)
        self.view = PageView(self.library.stamps)
        self.view.files_dropped.connect(self.add_files)
        self.view.stamp_resized.connect(self.library.remember_size)
        self.view.stamp_dropped.connect(self.drop_stamp)

        splitter = QSplitter()
        splitter.addWidget(self.pages)
        splitter.addWidget(self.view)
        splitter.addWidget(self.build_panel())
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([240, 640, 360])
        splitter.setContentsMargins(8, 8, 8, 8)
        self.setCentralWidget(splitter)
        self.build_toolbar()

        # Sliders fire many events while dragging; redraw once they settle.
        self.preview_timer = QTimer(self, singleShot=True, interval=120)
        self.preview_timer.timeout.connect(self.refresh_preview)
        self.refresh_preview()
        self.update_status()

    # ---------- UI construction ----------

    def build_toolbar(self):
        bar = self.addToolBar("Main")
        bar.setMovable(False)
        bar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
        brand = QLabel("PDF Maker")
        brand.setObjectName("brand")
        bar.addWidget(brand)

        def action(text, slot, shortcut=None, tip=""):
            act = QAction(text, self)
            act.triggered.connect(slot)
            if shortcut:
                act.setShortcut(QKeySequence(shortcut))
            act.setToolTip(f"{tip or text.strip('＋⟲⟳↑↓✕ ')} ({shortcut})" if shortcut else tip or text)
            bar.addAction(act)
            return act

        action("＋  Add files", self.choose_files, "Ctrl+O")
        bar.addSeparator()
        action("⟲  Rotate left", lambda: self.rotate(-90), "Ctrl+L")
        action("⟳  Rotate right", lambda: self.rotate(90), "Ctrl+R")
        action("↑  Move up", lambda: self.move(-1), "Ctrl+Up")
        action("↓  Move down", lambda: self.move(1), "Ctrl+Down")
        action("✕  Delete", self.delete_selected, "Del", "Remove the selected stamp, or the selected pages")
        action("Clear all", self.pages.clear, tip="Remove every page")
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        bar.addWidget(spacer)
        save = action("Save PDF", self.save_pdf, "Ctrl+S", "Convert and save the flattened PDF")
        bar.widgetForAction(save).setObjectName("primary")

    def build_panel(self) -> QWidget:
        def slider(name, lo, hi, value):
            s = QSlider(Qt.Orientation.Horizontal, minimum=lo, maximum=hi, value=value)
            s.setAccessibleName(name)
            label = QLabel(f"{value}%", alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            label.setObjectName("value")
            label.setMinimumWidth(42)
            s.valueChanged.connect(lambda v: label.setText(f"{v}%"))
            s.valueChanged.connect(self.settings_changed)
            row = QWidget()
            box = QHBoxLayout(row)
            box.setContentsMargins(0, 0, 0, 0)
            box.addWidget(s, 1)
            box.addWidget(label)
            return s, row

        def combo(items, current=0):
            c = QComboBox()
            c.addItems(items)
            c.setCurrentIndex(current)
            c.currentIndexChanged.connect(self.settings_changed)
            return c

        self.preset = combo(PRESETS, 1)
        output = QGroupBox("Output")
        QFormLayout(output).addRow("Quality", self.preset)

        self.mode = combo(MODES)
        self.brightness, brightness_row = slider("Brightness", 50, 150, 100)
        self.contrast, contrast_row = slider("Contrast", 50, 150, 100)
        self.crop, crop_row = slider("Crop edges", 0, 20, 0)
        self.clean_bg = QCheckBox("Clean page background (whiten)")
        self.clean_bg.setToolTip("Turns grey, yellow or shadowed paper into white. Best for scans and photos of documents.")
        self.clean_bg.toggled.connect(self.settings_changed)
        reset = QPushButton("Reset adjustments")
        reset.clicked.connect(self.reset_adjustments)
        image = QGroupBox("Image")
        form = QFormLayout(image)
        form.addRow("Color", self.mode)
        form.addRow("Brightness", brightness_row)
        form.addRow("Contrast", contrast_row)
        form.addRow("Crop edges", crop_row)
        form.addRow(self.clean_bg)
        form.addRow(reset)

        self.wm_text = QLineEdit(placeholderText="e.g. CONFIDENTIAL", clearButtonEnabled=True)
        self.wm_text.textChanged.connect(self.settings_changed)
        self.stamp_button = QPushButton("Choose image…")
        self.stamp_button.clicked.connect(self.choose_stamp)
        stamp_clear = QPushButton("Remove")
        stamp_clear.clicked.connect(lambda: self.set_stamp(""))
        stamp_row = QHBoxLayout()
        stamp_row.addWidget(self.stamp_button, 1)
        stamp_row.addWidget(stamp_clear)
        self.stamp_pos = combo(core.STAMP_POS)
        self.stamp_size, size_row = slider("Stamp size", 5, 80, 25)
        self.stamp_opacity, opacity_row = slider("Stamp opacity", 10, 100, 100)
        auto = QGroupBox("Automatic (same on all pages)")
        form = QFormLayout(auto)
        form.addRow("Watermark", self.wm_text)
        form.addRow("Stamp", stamp_row)
        form.addRow("Position", self.stamp_pos)
        form.addRow("Size", size_row)
        form.addRow("Opacity", opacity_row)

        def tab(*boxes, stretch_last=False):
            page = QWidget()
            page.setObjectName("panel")
            layout = QVBoxLayout(page)
            layout.setContentsMargins(0, 4, 0, 0)
            for box in boxes:
                layout.addWidget(box, 1 if stretch_last and box is boxes[-1] else 0)
            if not stretch_last:
                layout.addStretch()
            return page

        tabs = QTabWidget()
        tabs.addTab(tab(output, image), "Adjust")
        tabs.addTab(tab(auto, self.library, stretch_last=True), "Stamps")
        tabs.setMinimumWidth(330)
        return tabs

    # ---------- state ----------

    def settings(self) -> core.Settings:
        dpi, quality = PRESETS[self.preset.currentText()]
        return core.Settings(
            dpi=dpi,
            quality=quality,
            mode=MODES[self.mode.currentText()],
            brightness=self.brightness.value() / 100,
            contrast=self.contrast.value() / 100,
            crop=self.crop.value() / 100,
            clean_bg=self.clean_bg.isChecked(),
            wm_text=self.wm_text.text().strip(),
            stamp_path=self.stamp_path,
            stamp_pos=self.stamp_pos.currentText(),
            stamp_size=self.stamp_size.value() / 100,
            stamp_opacity=self.stamp_opacity.value() / 100,
        )

    def settings_changed(self, *_):
        self.preview_timer.start()

    def reset_adjustments(self):
        self.mode.setCurrentIndex(0)
        self.brightness.setValue(100)
        self.contrast.setValue(100)
        self.crop.setValue(0)
        self.clean_bg.setChecked(False)

    def all_pages(self) -> list[core.Page]:
        return [self.pages.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.pages.count())]

    def update_status(self):
        n = self.pages.count()
        self.statusBar().showMessage(f"{n} page{'s' if n != 1 else ''}" if n else "No pages yet")

    # ---------- background work ----------

    def run_task(self, title, fn, on_done):
        if self.worker and self.worker.isRunning():
            self.statusBar().showMessage("Still finishing the previous task, try again in a moment.")
            return
        dialog = QProgressDialog(title, "Cancel", 0, 0, self)
        dialog.setWindowTitle("PDF Maker")
        dialog.setWindowModality(Qt.WindowModality.WindowModal)
        dialog.setMinimumDuration(300)
        dialog.setAutoClose(False)
        dialog.setAutoReset(False)

        def progress(i, n):
            dialog.setMaximum(n)
            dialog.setValue(i)

        def finished(result):
            dialog.close()
            if result is not None:
                on_done(result)

        def failed(message):
            dialog.close()
            QMessageBox.critical(self, "PDF Maker", message)

        self.worker = Worker(fn, self)
        self.worker.progress.connect(progress)
        self.worker.done.connect(finished)
        self.worker.failed.connect(failed)
        dialog.canceled.connect(self.worker.requestInterruption)
        self.worker.start()

    # ---------- adding files ----------

    def choose_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Add files", "", FILE_FILTER)
        self.add_files(files)

    def add_files(self, files):
        if not files:
            return

        def load(tick):
            loaded, errors = [], []
            for i, f in enumerate(files):
                tick(i, len(files))
                try:
                    loaded += [(p, core.render(p, THUMB_DPI)) for p in core.load(Path(f))]
                except Exception as e:
                    errors.append(f"{Path(f).name}: {e}")
            return loaded, errors

        def loaded(result):
            pages, errors = result
            for page, thumb in pages:
                item = QListWidgetItem(QIcon(to_pixmap(thumb)), page.name)
                item.setData(Qt.ItemDataRole.UserRole, page)
                item.setToolTip(page.name)
                self.pages.addItem(item)
            if pages and not self.pages.currentItem():
                self.pages.setCurrentRow(0)
            self.update_status()
            if errors:
                QMessageBox.warning(self, "Some files were skipped", "\n".join(errors))

        self.run_task("Opening files…", load, loaded)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        self.add_files(dropped_files(event))

    # ---------- page management ----------

    def rotate(self, degrees):
        for item in self.pages.selectedItems():
            page = item.data(Qt.ItemDataRole.UserRole)
            page.rotation = (page.rotation + degrees) % 360
            item.setIcon(QIcon(to_pixmap(core.render(page, THUMB_DPI))))
        self.refresh_preview()

    def move(self, step):
        row = self.pages.currentRow()
        target = row + step
        if row < 0 or not 0 <= target < self.pages.count():
            return
        self.pages.insertItem(target, self.pages.takeItem(row))
        self.pages.setCurrentRow(target)

    def delete_selected(self):
        # Del is one window-wide shortcut: a selected stamp wins while the preview has focus.
        if self.focusWidget() is self.view and self.view.selected:
            self.view.remove_selected()
            return
        for item in self.pages.selectedItems():
            self.pages.takeItem(self.pages.row(item))

    # ---------- stamps ----------

    def place_stamp(self, path, cx=0.5, cy=0.5):
        if not self.view.pix:
            self.statusBar().showMessage("Add a document first, then place the stamp on a page.")
            return
        self.view.add_stamp(path, self.library.size_of(path), cx, cy)
        self.statusBar().showMessage(STAMP_HINT)

    def drop_stamp(self, cx, cy):
        item = self.library.stamps.currentItem()
        if item:
            self.place_stamp(item.data(Qt.ItemDataRole.UserRole), cx, cy)

    def choose_stamp(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose stamp image", str(self.library.folder()), IMAGE_FILTER)
        if path:
            self.set_stamp(path)

    def set_stamp(self, path):
        self.stamp_path = path
        self.stamp_button.setText(Path(path).name if path else "Choose image…")
        self.settings_changed()

    # ---------- preview ----------

    def refresh_preview(self):
        item = self.pages.currentItem()
        if not item:
            self.view.show_page(None, None, EMPTY_MESSAGE)
            return
        page = item.data(Qt.ItemDataRole.UserRole)
        try:  # hand-placed stamps are drawn live by the view, so they are left out here
            pix = to_pixmap(core.process(core.render(page, PREVIEW_DPI), self.settings()))
        except Exception as e:
            self.view.show_page(None, None, f"Cannot preview this page:\n{e}")
            return
        self.view.show_page(pix, page)

    # ---------- saving ----------

    def save_pdf(self):
        pages = self.all_pages()
        if not pages:
            self.statusBar().showMessage("Add some files first.")
            return
        first = Path(pages[0].name.split("  p")[0]).stem
        out, _ = QFileDialog.getSaveFileName(self, "Save PDF", f"{first}.flattened.pdf", "PDF (*.pdf)")
        if not out:
            return
        settings = self.settings()

        def saved(_):
            size = Path(out).stat().st_size / 1e6
            self.statusBar().showMessage(f"Saved {out} ({size:.1f} MB)")
            box = QMessageBox(QMessageBox.Icon.Information, "PDF Maker",
                              f"Saved {len(pages)} pages, {size:.1f} MB\n{out}", parent=self)
            open_button = box.addButton("Open file", QMessageBox.ButtonRole.AcceptRole)
            box.addButton(QMessageBox.StandardButton.Close)
            box.exec()
            if box.clickedButton() is open_button:
                QDesktopServices.openUrl(QUrl.fromLocalFile(out))

        self.run_task("Converting pages…", lambda tick: core.build(pages, settings, out, tick) or True, saved)


def dark_palette() -> QPalette:
    pal = QPalette()
    roles = QPalette.ColorRole
    for role, colour in {
        roles.Window: PANEL, roles.Base: CARD, roles.AlternateBase: PANEL, roles.Button: CARD,
        roles.WindowText: TEXT, roles.Text: TEXT, roles.ButtonText: TEXT, roles.ToolTipText: TEXT,
        roles.ToolTipBase: CARD, roles.PlaceholderText: MUTED, roles.Highlight: ACCENT,
        roles.HighlightedText: "#ffffff",
    }.items():
        pal.setColor(role, QColor(colour))
    return pal


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setPalette(dark_palette())
    app.setStyleSheet(STYLE)
    window = MainWindow()
    window.show()
    window.add_files([a for a in sys.argv[1:] if Path(a).exists()])  # "Open with" / drag onto the exe
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
