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

    # hand-placed stamp keeps its colour and lands where asked, even in grayscale mode
    stamp = core.Stamp(str(tmp / "stamp.png"), 0.5, 0.5, 0.2)
    img = core.process(Image.new("RGB", (1000, 1400), "white"), core.Settings(mode="gray"), [stamp])
    assert img.getpixel((520, 720)) == (200, 0, 0) and img.getpixel((480, 720)) == (255, 255, 255)
    assert img.getpixel((690, 720)) == (200, 0, 0) and img.getpixel((710, 720)) == (255, 255, 255)
    pages[2].stamps.append(stamp)
    core.build(pages, core.Settings(mode="bw"), tmp / "stamped.pdf")

    # a stamp pushed past the edge (e.g. by rotating after placing) is pulled back inside, not cut off
    edge = core.Stamp(str(tmp / "stamp.png"), 0.95, 0.98, 0.2)
    img = core.process(Image.new("RGB", (1000, 1400), "white"), core.Settings(), [edge])
    assert img.getpixel((999, 1399)) == (200, 0, 0) and img.getpixel((799, 1399)) == (255, 255, 255)

    try:
        core.build(pages, core.Settings(), tmp / "in.pdf")
        raise AssertionError("overwriting an input must be refused")
    except ValueError:
        pass

    # multi-page TIFF: every frame becomes a page, in order
    frames = [Image.new("RGB", (200, 280), c) for c in ("red", "green", "blue")]
    frames[0].save(tmp / "scan.tif", save_all=True, append_images=frames[1:])
    tif = core.load(tmp / "scan.tif")
    assert len(tif) == 3
    assert [core.render(p, 72).getpixel((5, 5)) for p in tif] == [(255, 0, 0), (0, 128, 0), (0, 0, 255)]

    # 16-bit scan: grey paper and dark ink stay grey and dark, not blown out to white
    deep = Image.new("I;16", (200, 280), 30000)
    deep.paste(Image.new("I;16", (200, 10), 2000), (0, 100))
    deep.save(tmp / "deep.tif")
    img = core.render(core.load(tmp / "deep.tif")[0], 72)
    assert 100 < img.getpixel((50, 50))[0] < 135 and img.getpixel((50, 105))[0] < 15

    # a failed write (output open in another program) leaves no .part file and a clear message
    out = tmp / "locked.pdf"
    out.write_bytes(b"old")
    with open(out, "rb"):  # Windows: an open file cannot be replaced
        try:
            core.build(pages[:1], core.Settings(), out)
            raise AssertionError("writing over an open file should fail")
        except RuntimeError as e:
            assert "close it" in str(e)
    assert not (tmp / "locked.pdf.part").exists() and out.read_bytes() == b"old"

    # releasing unused PDFs unlocks them, so the user can replace or delete them
    core.release(set())
    (tmp / "in.pdf").unlink()

    core._cleanup()

# Word: password-protected files fail fast with a clear error instead of hanging
try:
    import subprocess
    import threading
    import win32com.client

    word = win32com.client.DispatchEx("Word.Application")
except Exception:
    print("(Word not installed: Word checks skipped)")
else:
    with tempfile.TemporaryDirectory() as tmp:
        locked, plain = Path(tmp) / "locked.docx", Path(tmp) / "plain.docx"
        d = word.Documents.Add()
        d.Content.Text = "hello"
        d.SaveAs2(str(plain), FileFormat=16)
        d.Close(False)
        d = word.Documents.Add()
        d.Content.Text = "secret"
        d.Password = "secret"
        d.SaveAs2(str(locked), FileFormat=16)
        d.Close(False)
        word.Quit()
        result = []

        def convert_locked():
            try:
                core.load(locked)
            except RuntimeError as e:
                result.append(str(e))

        core.WORD_OPEN_TIMEOUT = 5  # keep the test quick
        t = threading.Thread(target=convert_locked, daemon=True)
        t.start()
        t.join(60)
        assert result and "password" in result[0], f"protected Word file did not fail cleanly: {result}"
        assert len(core.load(plain)) == 1, "a normal Word file must still convert"
        core._cleanup()
    leftovers = subprocess.run(
        ["powershell", "-NoProfile", "-Command",
         "(Get-CimInstance Win32_Process -Filter \"Name='WINWORD.EXE'\").CommandLine"],
        capture_output=True, text=True).stdout
    assert "/Automation" not in leftovers, "a hidden Word process was left running"

print("ok")
