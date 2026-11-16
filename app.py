"""Desktop app for convert.py: drop files, arrange pages, tweak, save a flattened PDF."""

import sys
from pathlib import Path

from PySide6.QtCore import QSize, Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import QAction, QDesktopServices, QIcon, QImage, QKeySequence, QPixmap
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFormLayout, QGroupBox,
    QHBoxLayout, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
    QMessageBox, QProgressDialog, QPushButton, QSizePolicy, QSlider, QSplitter,
    QVBoxLayout, QWidget,
)

import convert as core

EXTENSIONS = sorted(core.WORD_EXT | core.IMAGE_EXT | {".pdf"})
FILE_FILTER = f"Documents and images ({' '.join('*' + e for e in EXTENSIONS)})"
PRESETS = {  # label -> (dpi, JPEG quality; 0 = lossless PNG)
    "Small file (150 DPI)": (150, 75),
    "Balanced (200 DPI)": (200, 88),
    "High quality (300 DPI)": (300, 95),
    "Lossless (300 DPI, large file)": (300, 0),
}
MODES = {"Color": "color", "Grayscale": "gray", "Black and white": "bw"}
THUMB_DPI, PREVIEW_DPI = 20, 80


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


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Maker")
        self.resize(1150, 760)
        self.setAcceptDrops(True)
        self.worker = None
        self.preview_pix = None
        self.stamp_path = ""

        self.pages = PageList()
        self.pages.files_dropped.connect(self.add_files)
        self.pages.currentItemChanged.connect(lambda *_: self.refresh_preview())
        self.pages.model().rowsMoved.connect(lambda *_: self.update_status())

        self.preview = QLabel(alignment=Qt.AlignmentFlag.AlignCenter)
        self.preview.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)
        self.preview.setMinimumSize(300, 300)

        splitter = QSplitter()
        splitter.addWidget(self.pages)
        splitter.addWidget(self.preview)
        splitter.addWidget(self.build_panel())
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([230, 620, 300])
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

        def action(text, slot, shortcut=None, tip=""):
            act = QAction(text, self)
            act.triggered.connect(slot)
            if shortcut:
                act.setShortcut(QKeySequence(shortcut))
            act.setToolTip(f"{tip or text} ({shortcut})" if shortcut else tip or text)
            bar.addAction(act)
            return act

        action("Add files…", self.choose_files, "Ctrl+O")
        bar.addSeparator()
        action("Rotate left", lambda: self.rotate(-90), "Ctrl+L")
        action("Rotate right", lambda: self.rotate(90), "Ctrl+R")
        action("Move up", lambda: self.move(-1), "Ctrl+Up")
        action("Move down", lambda: self.move(1), "Ctrl+Down")
        action("Delete", self.delete_selected, "Del", "Remove selected pages")
        action("Clear all", self.pages.clear)
        self.pages.model().rowsRemoved.connect(lambda *_: (self.update_status(), self.refresh_preview()))
        bar.addSeparator()
        action("Save PDF…", self.save_pdf, "Ctrl+S", "Convert and save the flattened PDF")

    def build_panel(self) -> QWidget:
        def slider(lo, hi, value):
            s = QSlider(Qt.Orientation.Horizontal, minimum=lo, maximum=hi, value=value)
            s.valueChanged.connect(self.settings_changed)
            return s

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
        self.brightness = slider(50, 150, 100)
        self.contrast = slider(50, 150, 100)
        self.crop = slider(0, 20, 0)
        self.clean_bg = QCheckBox("Clean page background (whiten)")
        self.clean_bg.setToolTip("Turns grey, yellow or shadowed paper into white. Best for scans and photos of documents.")
        self.clean_bg.toggled.connect(self.settings_changed)
        reset = QPushButton("Reset adjustments")
        reset.clicked.connect(self.reset_adjustments)
        image = QGroupBox("Image")
        form = QFormLayout(image)
        form.addRow("Color", self.mode)
        form.addRow("Brightness", self.brightness)
        form.addRow("Contrast", self.contrast)
        form.addRow("Crop edges", self.crop)
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
        self.stamp_size = slider(5, 80, 25)
        self.stamp_opacity = slider(10, 100, 100)
        mark = QGroupBox("Watermark and stamp (all pages)")
        form = QFormLayout(mark)
        form.addRow("Text", self.wm_text)
        form.addRow("Stamp", stamp_row)
        form.addRow("Position", self.stamp_pos)
        form.addRow("Size", self.stamp_size)
        form.addRow("Opacity", self.stamp_opacity)

        panel = QWidget()
        layout = QVBoxLayout(panel)
        for box in (output, image, mark):
            layout.addWidget(box)
        layout.addStretch()
        return panel

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
        for item in self.pages.selectedItems():
            self.pages.takeItem(self.pages.row(item))

    # ---------- watermark ----------

    def choose_stamp(self):
        images = " ".join("*" + e for e in sorted(core.IMAGE_EXT))
        path, _ = QFileDialog.getOpenFileName(self, "Choose stamp image", "", f"Images ({images})")
        if path:
            self.set_stamp(path)

    def set_stamp(self, path):
        self.stamp_path = path
        self.stamp_button.setText(Path(path).name if path else "Choose image…")
        self.settings_changed()

    # ---------- preview ----------

    def refresh_preview(self):
        item = self.pages.currentItem()
        self.preview_pix = None
        if not item:
            self.preview.setText("Drop Word, PDF or image files here\nor click “Add files…”")
            return
        try:
            page = item.data(Qt.ItemDataRole.UserRole)
            self.preview_pix = to_pixmap(core.process(core.render(page, PREVIEW_DPI), self.settings()))
        except Exception as e:
            self.preview.setText(f"Cannot preview this page:\n{e}")
            return
        self.fit_preview()

    def fit_preview(self):
        if self.preview_pix:
            self.preview.setPixmap(self.preview_pix.scaled(
                self.preview.size(), Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.fit_preview()

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


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    window.add_files([a for a in sys.argv[1:] if Path(a).exists()])  # "Open with" / drag onto the exe
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
