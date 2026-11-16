#!/usr/bin/env python3
"""
Flatten Word / PDF / image files into one PDF where every page is a picture
(no live text, no fonts, nothing editable).

Engine for app.py, also usable from the command line:
  python convert.py input.docx more.pdf scan.jpg [-o out.pdf] [--dpi 200]
                    [--quality 88] [--gray] [--clean]
"""

import argparse
import atexit
import io
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import uuid
from dataclasses import dataclass, field
from pathlib import Path

import pymupdf
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageMath, ImageOps

A4 = (595.276, 841.89)  # points
WORD_OPEN_TIMEOUT = 30  # seconds; a normal document opens in a few
WORD_EXT = {".docx", ".doc", ".rtf", ".odt"}
IMAGE_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
STAMP_POS = {
    "Bottom right": (1, 1),
    "Bottom left": (0, 1),
    "Top right": (1, 0),
    "Top left": (0, 0),
    "Center": (0.5, 0.5),
}


@dataclass(eq=False)
class Stamp:
    """An image placed by hand on one page. Position and width are fractions of the page."""

    path: str
    x: float  # left edge / page width
    y: float  # top edge / page height
    w: float  # width / page width


@dataclass
class Page:
    path: str  # the PDF or image actually rendered (the converted PDF for Word files)
    index: int
    name: str
    is_image: bool = False
    rotation: int = 0  # clockwise degrees, multiple of 90
    stamps: list = field(default_factory=list)  # hand-placed Stamp objects


@dataclass
class Settings:
    dpi: int = 200
    quality: int = 88  # JPEG quality; 0 = lossless PNG
    mode: str = "color"  # color | gray | bw
    brightness: float = 1.0
    contrast: float = 1.0
    crop: float = 0.0  # fraction trimmed from every side
    clean_bg: bool = False
    wm_text: str = ""
    stamp_path: str = ""
    stamp_pos: str = "Bottom right"
    stamp_size: float = 0.25  # fraction of page width
    stamp_opacity: float = 1.0


_docs: dict[str, pymupdf.Document] = {}
_lock = threading.Lock()  # MuPDF documents must not be used from two threads at once
_tmp = None


def _cleanup():
    for doc in _docs.values():
        doc.close()
    _docs.clear()
    if _tmp:
        shutil.rmtree(_tmp, ignore_errors=True)


atexit.register(_cleanup)


def release(in_use: set):
    """Close source PDFs that no page uses any more, so Windows unlocks the files."""
    with _lock:
        for key in [k for k in _docs if k not in in_use]:
            _docs.pop(key).close()


def word_to_pdf(path: Path) -> Path:
    """Convert with Microsoft Word (best fidelity); fall back to LibreOffice."""
    global _tmp
    _tmp = _tmp or tempfile.mkdtemp(prefix="flatpdf_")
    out_dir = Path(tempfile.mkdtemp(dir=_tmp))
    out = out_dir / (path.stem + ".pdf")
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
    except Exception as e:
        word, word_err = None, e
    if word:
        # Word is invisible, so a password or repair prompt would wait forever. Find the
        # process of *this* Word (by a unique caption) and kill it if opening takes too long.
        import win32gui
        import win32process

        word.Caption = tag = f"pdfmaker-{uuid.uuid4().hex}"
        pid = win32process.GetWindowThreadProcessId(win32gui.FindWindow("OpusApp", tag))[1]
        watchdog = threading.Timer(WORD_OPEN_TIMEOUT, lambda: subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)], capture_output=True))
        try:
            word.DisplayAlerts = 0
            word.AutomationSecurity = 3  # force-disable macros: never run code from files we convert
            watchdog.start()
            doc = word.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
            watchdog.cancel()
            doc.SaveAs2(str(out), FileFormat=17)  # 17 = PDF
            doc.Close(False)
        except Exception as e:
            raise RuntimeError(f"Word could not open {path.name}. It may be password-protected or damaged.\n({e})")
        finally:
            watchdog.cancel()
            try:
                word.Quit()
            except Exception:
                pass  # already killed by the watchdog
    else:
        candidates = [
            shutil.which("soffice"),
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        soffice = next((c for c in candidates if c and os.path.exists(c)), None)
        if not soffice:
            raise RuntimeError(
                f"Could not convert {path.name}: Microsoft Word or LibreOffice is "
                f"required. (Word said: {word_err})"
            )
        cmd = [soffice, "--headless", "--norestore", "--convert-to", "pdf",
               "--outdir", str(out_dir), str(path)]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        except subprocess.TimeoutExpired:
            raise RuntimeError(f"LibreOffice took too long on {path.name}. It may be password-protected or damaged.")
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice failed on {path.name}:\n{result.stdout}{result.stderr}")
    if not out.exists():
        raise RuntimeError(f"No PDF was produced for {path.name}.")
    return out


def load(path: Path) -> list[Page]:
    """Open one input file and return its pages."""
    path = Path(path).resolve()
    ext = path.suffix.lower()
    if ext in IMAGE_EXT:
        with Image.open(path) as im:
            frames = getattr(im, "n_frames", 1) if ext in (".tif", ".tiff") else 1  # scanners put many pages in one TIFF
            im.verify()
        if frames == 1:
            return [Page(str(path), 0, path.name, is_image=True)]
        return [Page(str(path), i, f"{path.name}  p{i + 1}", is_image=True) for i in range(frames)]
    if ext in WORD_EXT:
        pdf = word_to_pdf(path)
    elif ext == ".pdf":
        pdf = path
    else:
        raise ValueError(f"Unsupported file type: {ext or path.name}")
    with _lock:
        doc = _docs.get(str(pdf)) or pymupdf.open(pdf)
        if doc.needs_pass:
            doc.close()
            raise ValueError("PDF is password-protected.")
        _docs[str(pdf)] = doc
        count = doc.page_count
    return [Page(str(pdf), i, f"{path.name}  p{i + 1}") for i in range(count)]


def render(page: Page, dpi: int) -> Image.Image:
    """Rasterize one page to an RGB image, rotation applied."""
    if page.is_image:
        with Image.open(page.path) as f:
            f.seek(page.index)
            img = ImageOps.exif_transpose(f)
        if img.mode.startswith("I;16") or img.mode == "I":  # 16-bit scans: scale down, don't clip to white
            img = img.convert("I").point(lambda v: v / 256).convert("L")
        if img.mode in ("RGBA", "LA", "P"):  # transparent areas become white, not black
            img = img.convert("RGBA")
            white = Image.new("RGB", img.size, "white")
            white.paste(img, mask=img.getchannel("A"))
            img = white
        img = img.convert("RGB")
        limit = int(A4[1] / 72 * dpi)  # no point keeping more pixels than A4 at this DPI
        img.thumbnail((limit, limit), Image.LANCZOS)
    else:
        with _lock:
            pix = _docs[page.path][page.index].get_pixmap(dpi=dpi, alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    if page.rotation % 360:
        img = img.rotate(-page.rotation, expand=True)
    return img


def clean_background(img: Image.Image) -> Image.Image:
    """Turn grey / yellow / shadowed paper into white.

    Estimates the paper colour around every pixel (brightest nearby value,
    blurred) and divides it out, so uneven lighting flattens to white.
    """
    small = img.resize((400, max(1, round(400 * img.height / img.width))), Image.BILINEAR)
    paper = small.filter(ImageFilter.MaxFilter(9)).filter(ImageFilter.GaussianBlur(4))
    paper = paper.resize(img.size, Image.BILINEAR)
    k = 255 * 255 / 235  # push near-white to white

    def divide(ink, bg):
        return ImageMath.lambda_eval(
            lambda e: e["convert"](e["min"](e["float"](e["i"]) * k / e["max"](e["p"], 1), 255), "L"),
            i=ink, p=bg)

    return Image.merge("RGB", [divide(i, p) for i, p in zip(img.split(), paper.split())])


def _font(size: int):
    try:
        return ImageFont.truetype("arial.ttf", size)
    except OSError:
        return ImageFont.load_default(size)


def overlay(img: Image.Image, s: Settings) -> Image.Image:
    """Draw the diagonal text watermark and/or the stamp image."""
    base = img.convert("RGBA")
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    if s.wm_text:
        width_at_100 = max(_font(100).getlength(s.wm_text), 1)
        size = min(int(100 * 0.75 * base.width / width_at_100), base.height // 6)
        font = _font(max(size, 8))
        l, t, r, b = font.getbbox(s.wm_text)
        text = Image.new("RGBA", (r - l + 20, b - t + 20), (0, 0, 0, 0))
        ImageDraw.Draw(text).text((10 - l, 10 - t), s.wm_text, font=font, fill=(128, 128, 128, 80))
        text = text.rotate(30, expand=True, resample=Image.BICUBIC)
        layer.paste(text, ((base.width - text.width) // 2, (base.height - text.height) // 2), text)
    if s.stamp_path:
        stamp = Image.open(s.stamp_path).convert("RGBA")
        w = max(1, int(base.width * s.stamp_size))
        stamp = stamp.resize((w, max(1, round(stamp.height * w / stamp.width))), Image.LANCZOS)
        if s.stamp_opacity < 1:
            stamp.putalpha(stamp.getchannel("A").point(lambda a: int(a * s.stamp_opacity)))
        margin = int(base.width * 0.04)
        fx, fy = STAMP_POS[s.stamp_pos]
        x = margin + int((base.width - w - 2 * margin) * fx)
        y = margin + int((base.height - stamp.height - 2 * margin) * fy)
        layer.paste(stamp, (x, y), stamp)
    return Image.alpha_composite(base, layer).convert("RGB")


def process(img: Image.Image, s: Settings, stamps=()) -> Image.Image:
    """Apply crop, clean-up, adjustments, colour mode, watermark and hand-placed stamps."""
    if s.crop:
        dx, dy = int(img.width * s.crop), int(img.height * s.crop)
        img = img.crop((dx, dy, img.width - dx, img.height - dy))
    if s.clean_bg:
        img = clean_background(img)
    if s.brightness != 1:
        img = ImageEnhance.Brightness(img).enhance(s.brightness)
    if s.contrast != 1:
        img = ImageEnhance.Contrast(img).enhance(s.contrast)
    if s.mode != "color":
        img = img.convert("L")
        if s.mode == "bw":
            img = img.point(lambda v: 255 if v > 150 else 0)
    if s.wm_text or s.stamp_path:
        img = overlay(img, s)
        if s.mode != "color":
            img = img.convert("L")
    if stamps:  # pasted last so they keep their own colours, exactly as previewed
        img = img.convert("RGB")
        for st in stamps:
            im = Image.open(st.path).convert("RGBA")
            w = max(1, round(img.width * st.w))
            im = im.resize((w, max(1, round(im.height * w / im.width))), Image.LANCZOS)
            # Clamp: rotating or cropping after placing can push a stamp past the edge.
            x = min(max(round(st.x * img.width), 0), max(img.width - im.width, 0))
            y = min(max(round(st.y * img.height), 0), max(img.height - im.height, 0))
            img.paste(im, (x, y), im)
    return img


def build(pages: list[Page], s: Settings, out_path: Path, progress=lambda i, n: None):
    """Render every page and write the flattened PDF."""
    out_path = Path(out_path).resolve()
    if any(Path(p.path) == out_path for p in pages):
        raise ValueError("The output file cannot be one of the input files.")
    out = pymupdf.open()
    tmp = out_path.with_name(out_path.name + ".part")  # never leave a half-written PDF
    try:
        _write_pages(out, pages, s, progress)
        out.save(tmp, deflate=True, garbage=3)
        out.close()
        os.replace(tmp, out_path)
    except PermissionError:
        raise RuntimeError(f"Could not write {out_path.name}. If it is open in another program, "
                           "close it and try again.") from None
    finally:
        if not out.is_closed:
            out.close()
        tmp.unlink(missing_ok=True)


def _write_pages(out, pages, s, progress):
    # shortcut: pages are rendered one at a time; go multi-process if huge jobs feel slow
    for i, page in enumerate(pages):
        progress(i, len(pages))
        img = process(render(page, s.dpi), s, page.stamps)
        buf = io.BytesIO()
        if s.quality and (s.mode != "bw" or page.stamps):  # black-and-white is far smaller as PNG
            img.save(buf, "JPEG", quality=s.quality)
        else:
            img.save(buf, "PNG")
        w, h = img.width * 72 / s.dpi, img.height * 72 / s.dpi
        if page.is_image:  # photos have no paper size: centre them on an A4 sheet
            pw, ph = A4 if h >= w else A4[::-1]
            k = min(pw / w, ph / h)
            w, h = w * k, h * k
            rect = pymupdf.Rect((pw - w) / 2, (ph - h) / 2, (pw + w) / 2, (ph + h) / 2)
        else:
            pw, ph, rect = w, h, pymupdf.Rect(0, 0, w, h)
        out.new_page(width=pw, height=ph).insert_image(rect, stream=buf.getvalue())
    progress(len(pages), len(pages))


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("inputs", type=Path, nargs="+", help="Word, PDF or image files, in order")
    parser.add_argument("-o", "--output", type=Path, help="Output PDF path")
    parser.add_argument("--dpi", type=int, default=200)
    parser.add_argument("--quality", type=int, default=88, help="JPEG quality 1-100, 0 = lossless PNG")
    parser.add_argument("--gray", action="store_true", help="Grayscale output")
    parser.add_argument("--clean", action="store_true", help="Whiten page backgrounds")
    args = parser.parse_args()

    pages = []
    for path in args.inputs:
        if not path.exists():
            sys.exit(f"Input file not found: {path}")
        print(f"Loading {path.name} ...")
        pages += load(path)
    out = args.output or args.inputs[0].with_suffix(".flattened.pdf")
    settings = Settings(dpi=args.dpi, quality=args.quality, clean_bg=args.clean,
                        mode="gray" if args.gray else "color")
    build(pages, settings, out, lambda i, n: print(f"\rPage {i}/{n}", end=""))
    print(f"\nDone -> {out}")


if __name__ == "__main__":
    main()
