"""Self-check for the window logic: python test_app.py (runs hidden, touches no real user folders)."""

import os
import tempfile
import time
from pathlib import Path

os.environ["QT_QPA_PLATFORM"] = "offscreen"

import pymupdf
from PIL import Image
from PySide6.QtCore import QPoint, Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

import app
import convert as core

qa = QApplication([])


def pump(n=10):
    for _ in range(n):
        qa.processEvents()
        time.sleep(0.01)


def wait(w):
    while w.worker and w.worker.isRunning():
        pump(1)
    pump()


with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src = pymupdf.open()
    for _ in range(3):
        src.new_page(width=core.A4[0], height=core.A4[1])
    src.save(tmp / "doc.pdf")
    src.close()

    app.STAMP_DIR = tmp / "stamps"  # never touch the user's real stamp library
    w = app.MainWindow()
    w.resize(1200, 800)
    w.show()
    lib = w.library
    assert lib.root == app.STAMP_DIR
    Image.new("RGBA", (300, 300), (30, 70, 200, 255)).save(lib.root / "General" / "seal.png")
    lib.reload_folders()
    lib.stamps.setCurrentRow(0)

    w.add_files([str(tmp / "doc.pdf")])
    wait(w)
    assert w.pages.count() == 3 and w.statusBar().currentMessage() == "3 pages"
    page, v = w.all_pages()[0], w.view

    # place, move with the mouse, resize with the corner handle; the size is remembered
    lib.place_current()
    st = page.stamps[0]
    assert st.w == 0.25
    r = v.page_rect()
    c = v.stamp_rect(st).center().toPoint()
    x0, y0 = st.x, st.y
    QTest.mousePress(v, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier, c)
    QTest.mouseMove(v, c + QPoint(100, 120))
    QTest.mouseRelease(v, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier, c + QPoint(100, 120))
    assert abs((st.x - x0) * r.width() - 100) < 2 and abs((st.y - y0) * r.height() - 120) < 2
    h = v.handle_rect(st).center().toPoint()
    width_px = st.w * r.width()
    QTest.mousePress(v, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier, h)
    QTest.mouseMove(v, h - QPoint(40, 40))
    QTest.mouseRelease(v, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier, h - QPoint(40, 40))
    assert abs(width_px - st.w * r.width() - 40) < 2
    assert lib.sizes == {"General/seal.png": round(st.w, 4)}
    lib.place_current()
    assert page.stamps[1].w == lib.sizes["General/seal.png"], "next placement should reuse the size"

    # Delete: with a stamp selected in the preview it removes the stamp, never pages
    v.setFocus()
    w.delete_selected()
    assert len(page.stamps) == 1 and w.pages.count() == 3
    # ...and with focus anywhere else (e.g. the stamp library) it removes nothing
    w.pages.selectAll()
    lib.stamps.setFocus()
    w.delete_selected()
    assert w.pages.count() == 3, "Del outside the page list must not delete pages"

    # a stamp near the bottom stays inside the page after the page is rotated
    w.pages.clearSelection()
    w.pages.setCurrentRow(0)
    w.pages.item(0).setSelected(True)
    v.place(st, 0.4, 0.9)
    w.rotate(90)
    height = v.stamp_rect(st).height() / v.page_rect().height()
    assert 0 <= st.y and st.y + height <= 1 + 1e-6 and st.x + st.w <= 1 + 1e-6

    # the saved PDF has the stamp on page 1 only
    out = tmp / "out.pdf"
    core.build(w.all_pages(), w.settings(), out)
    doc = pymupdf.open(out)
    pm = doc[0].get_pixmap(dpi=40)
    cx, cy = int((st.x + st.w / 2) * pm.width), int((st.y + height / 2) * pm.height)
    r_, g, b = pm.pixel(cx, cy)
    assert b > 150 and r_ < 80, f"stamp not where it was placed: {(r_, g, b)}"
    assert pymupdf.open(out)[1].get_pixmap(dpi=40).pixel(cx % 100, cy % 100) == (255, 255, 255)
    doc.close()

    # Delete in the page list removes pages; Clear all empties everything and unlocks the file
    w.pages.setFocus()
    w.pages.clearSelection()
    w.pages.item(2).setSelected(True)
    w.delete_selected()
    pump()
    assert w.pages.count() == 2 and w.statusBar().currentMessage() == "2 pages"
    w.pages.clear()
    pump()
    assert w.statusBar().currentMessage() == "No pages yet" and v.pix is None
    (tmp / "doc.pdf").unlink()  # fails on Windows if the app still holds the PDF open

    w.close()
    core._cleanup()

print("ok")
