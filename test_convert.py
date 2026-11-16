"""Self-check for the engine: python test_convert.py"""

import tempfile
from pathlib import Path

import pymupdf
from PIL import Image

import convert as core

with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)

    src = pymupdf.open()
    for text in ("first page", "second page"):
        src.new_page(width=core.A4[0], height=core.A4[1]).insert_text((72, 100), text, fontsize=30)
    src.save(tmp / "in.pdf")
    src.close()
    Image.new("RGBA", (1600, 900), (200, 180, 150, 255)).save(tmp / "photo.png")
    Image.new("RGBA", (300, 120), (200, 0, 0, 255)).save(tmp / "stamp.png")

    pages = core.load(tmp / "in.pdf") + core.load(tmp / "photo.png")
    assert len(pages) == 3
    pages[1].rotation = 90
    pages[0], pages[1] = pages[1], pages[0]  # reorder

    for s in (
        core.Settings(),
        core.Settings(quality=0, dpi=100),
        core.Settings(mode="bw", crop=0.05, brightness=1.1, contrast=1.2),
        core.Settings(mode="gray", clean_bg=True, wm_text="DRAFT",
                      stamp_path=str(tmp / "stamp.png"), stamp_opacity=0.5),
    ):
        out = tmp / "out.pdf"
        core.build(pages, s, out)
        doc = pymupdf.open(out)
        assert doc.page_count == 3
        assert all(not p.get_text().strip() for p in doc), "live text survived"
        assert all(len(p.get_images()) == 1 and not p.get_fonts() for p in doc)
        if not s.crop:
            rotated, plain, photo = (p.rect for p in doc)
            assert abs(rotated.width - core.A4[1]) < 2 and abs(rotated.height - core.A4[0]) < 2
            assert abs(plain.width - core.A4[0]) < 2 and abs(plain.height - core.A4[1]) < 2
            assert abs(photo.width - core.A4[1]) < 2, "landscape photo should sit on landscape A4"
        doc.close()

    # whitening: grey paper with dark ink -> white paper, ink stays dark
    grey = Image.new("RGB", (800, 1000), (170, 165, 150))
    grey.paste((20, 20, 20), (100, 100, 400, 110))
    cleaned = core.clean_background(grey)
    assert min(cleaned.getpixel((600, 600))) > 245
    assert max(cleaned.getpixel((200, 105))) < 80

    try:
        core.build(pages, core.Settings(), tmp / "in.pdf")
        raise AssertionError("overwriting an input must be refused")
    except ValueError:
        pass

    core._cleanup()

print("ok")
