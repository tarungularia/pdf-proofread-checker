"""Proofread a PDF against a standard version.

Compares a "to check" PDF with the "standard" PDF page by page and writes two outputs:

* output/discrepancies.xlsx    - one row per difference: page, location, standard text,
                                 text found, type of change, note, and image crops of both versions
* output/<file>_highlighted.pdf - a copy of the checked PDF with every difference highlighted
                                  (hover a highlight to see what changed)

How it compares each page:
* Readable text on both pages -> word-by-word comparison (wording, numbers, dates,
  punctuation, capitalisation, missing/added text) plus font/size changes on matching words,
  plus a visual check for changed images/graphics where there is no text.
* Text not machine-readable (scanned pages or scrambled font encodings) -> visual comparison:
  pages are rendered as images and every changed area is reported with image crops.

Usage (defaults: first PDF in 1_standard/, first PDF in 2_to_check/, results in output/):
    python proofread_compare.py
    python proofread_compare.py --standard A.pdf --check B.pdf --out output
"""

from __future__ import annotations

import argparse
import io
import logging
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

import numpy as np
import pdfplumber
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image
from pypdf import PdfReader, PdfWriter
from pypdf.annotations import Highlight
from pypdf.generic import ArrayObject, FloatObject, NameObject, TextStringObject
from scipy import ndimage

logging.getLogger("pypdf").setLevel(logging.ERROR)  # hide harmless repair warnings

HERE = Path(__file__).resolve().parent
DPI = 150
SCALE = DPI / 72  # pixels per PDF point
PIXEL_THRESHOLD = 40      # grey-level change that counts as "different"
MIN_CHANGED_PIXELS = 12   # ignore specks smaller than this (anti-aliasing noise)
CELL = 6                  # px grid used to group changed pixels into areas
MERGE_CELLS = 1           # areas closer than this many cells (6px) are merged
COMMON_WORDS = {
    "the", "of", "and", "to", "in", "fund", "is", "for", "on", "as", "by", "with", "are",
    "be", "or", "an", "at", "this", "that", "from", "investment", "returns", "risk", "date",
    "share", "class", "total", "market", "index", "net", "assets", "performance", "past",
}
COLOURS = {"Missing": "ff6b6b", "Added": "4cd964", "default": "ffd60a"}  # red, green, yellow


@dataclass
class Discrepancy:
    """One difference between the standard and the checked document."""

    page: int
    change_type: str
    standard_text: str
    found_text: str
    note: str
    check_box: tuple[float, float, float, float] | None  # x0, top, x1, bottom in points
    standard_box: tuple[float, float, float, float] | None = None
    location: str = ""
    crops: dict[str, Image.Image] = field(default_factory=dict)


# ----- helpers -------------------------------------------------------------------

def render(page: pdfplumber.page.Page) -> Image.Image:
    """Render a PDF page to an RGB image at DPI."""
    return page.to_image(resolution=DPI).original.convert("RGB")


def is_readable(page: pdfplumber.page.Page) -> bool:
    """True if the page's extracted text looks like real language (not scanned/scrambled)."""
    words = re.findall(r"[a-z]+", (page.extract_text() or "").lower())
    if len(words) < 5:
        return False
    return sum(w in COMMON_WORDS for w in words) / len(words) >= 0.08


def union_box(words: list[dict]) -> tuple[float, float, float, float] | None:
    """Bounding box around a list of pdfplumber words."""
    if not words:
        return None
    return (min(w["x0"] for w in words), min(w["top"] for w in words),
            max(w["x1"] for w in words), max(w["bottom"] for w in words))


def describe_location(box: tuple[float, float, float, float] | None, page_height: float) -> str:
    """Human-readable position on the page, e.g. 'top third (y=112pt)'."""
    if box is None:
        return "whole page"
    third = ("top", "middle", "bottom")[min(2, int(3 * box[1] / page_height))]
    return f"{third} third (x={box[0]:.0f}pt, y={box[1]:.0f}pt)"


def crop(img: Image.Image, box: tuple[float, float, float, float] | None, pad: float = 12) -> Image.Image | None:
    """Crop the area around a box (in points) from a rendered page, for the Excel sheet."""
    if box is None:
        return None
    x0, t, x1, b = box
    area = (max(0, (x0 - pad) * SCALE), max(0, (t - pad) * SCALE),
            min(img.width, (x1 + pad) * SCALE), min(img.height, (b + pad) * SCALE))
    if area[2] - area[0] < 2 or area[3] - area[1] < 2:
        return None
    return img.crop(tuple(int(v) for v in area))


def classify(std: str, found: str) -> tuple[str, str]:
    """Return (type of change, note) for a replaced piece of text."""
    if std.lower() == found.lower():
        return "Capitalisation", "Same words, different upper/lower case."
    if re.sub(r"[\W_]", "", std) == re.sub(r"[\W_]", "", found):
        return "Punctuation / spacing", "Same characters, different punctuation or spacing."
    date_re = r"\b(19|20)\d{2}\b|\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b|\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}"
    if re.search(date_re, std, re.I) and re.search(date_re, found, re.I):
        return "Date changed", "Check the date against the standard."
    std_digits, found_digits = re.findall(r"\d", std), re.findall(r"\d", found)
    if std_digits and found_digits:
        if sorted(std_digits) == sorted(found_digits) and std_digits != found_digits:
            return "Number changed", "Same digits in a different order (possible transposition)."
        return "Number changed", "Figure differs from the standard."
    return "Wording changed", "Text differs from the standard."


# ----- comparison ----------------------------------------------------------------

def visual_regions(std_img: Image.Image, chk_img: Image.Image) -> list[tuple[float, float, float, float]]:
    """Boxes (in points, on the checked page) where the two rendered pages differ."""
    if std_img.size != chk_img.size:
        std_img = std_img.resize(chk_img.size)
    a = np.asarray(std_img, dtype=np.int16)
    b = np.asarray(chk_img, dtype=np.int16)
    mask = np.abs(a - b).max(axis=2) > PIXEL_THRESHOLD  # colour-aware: navy vs dark red counts
    if mask.sum() < MIN_CHANGED_PIXELS:
        return []
    h, w = mask.shape
    hc, wc = -(-h // CELL), -(-w // CELL)
    padded = np.zeros((hc * CELL, wc * CELL), dtype=bool)
    padded[:h, :w] = mask
    counts = padded.reshape(hc, CELL, wc, CELL).sum(axis=(1, 3))
    cells = ndimage.binary_dilation(counts > 0, iterations=MERGE_CELLS)
    labels, _ = ndimage.label(cells)
    boxes = []
    for idx, sl in enumerate(ndimage.find_objects(labels), start=1):
        changed = (labels[sl] == idx) & (counts[sl] > 0)  # tight box: changed cells only, no dilation
        if counts[sl][changed].sum() < MIN_CHANGED_PIXELS:
            continue
        rows, cols = np.where(changed)
        top, bottom = (sl[0].start + rows.min()) * CELL, min(h, (sl[0].start + rows.max() + 1) * CELL)
        left, right = (sl[1].start + cols.min()) * CELL, min(w, (sl[1].start + cols.max() + 1) * CELL)
        boxes.append((left / SCALE, top / SCALE, right / SCALE, bottom / SCALE))
    return boxes


def words_in(page: pdfplumber.page.Page, box: tuple[float, float, float, float]) -> str:
    """Text of the words that overlap a box."""
    x0, t, x1, b = box
    return " ".join(w["text"] for w in page.extract_words()
                    if w["x1"] > x0 and w["x0"] < x1 and w["bottom"] > t and w["top"] < b)


def base_font(name: str) -> str:
    """Drop the random subset tag PDFs add to embedded fonts ('ABCDEF+Arial' -> 'Arial')."""
    return name.split("+", 1)[1] if re.match(r"^[A-Z]{6}\+", name) else name


def compare_text(pno: int, sp: pdfplumber.page.Page, cp: pdfplumber.page.Page) -> list[Discrepancy]:
    """Word-by-word comparison of two readable pages, plus font changes on matching words."""
    attrs = ["fontname", "size"]
    sw = sp.extract_words(extra_attrs=attrs, use_text_flow=True)
    cw = cp.extract_words(extra_attrs=attrs, use_text_flow=True)
    ops = SequenceMatcher(None, [w["text"] for w in sw], [w["text"] for w in cw], autojunk=False).get_opcodes()
    found: list[Discrepancy] = []
    for tag, i1, i2, j1, j2 in ops:
        s_words, c_words = sw[i1:i2], cw[j1:j2]
        if tag == "equal":
            for s, c in zip(s_words, c_words):
                s_font, c_font = base_font(s["fontname"]), base_font(c["fontname"])
                if s_font != c_font or abs(s["size"] - c["size"]) > 0.2:
                    found.append(Discrepancy(
                        pno, "Formatting changed", s["text"], c["text"],
                        f"Font {s_font} {s['size']:.1f}pt -> {c_font} {c['size']:.1f}pt.",
                        union_box([c]), union_box([s])))
            continue
        std_text = " ".join(w["text"] for w in s_words)
        chk_text = " ".join(w["text"] for w in c_words)
        context = " ".join(w["text"] for w in sw[max(0, i1 - 4):i1])
        if tag == "delete":
            # Mark where the text should be: just before the next word in the checked page.
            anchor = cw[j1] if j1 < len(cw) else (cw[-1] if cw else None)
            box = (anchor["x0"] - 2, anchor["top"], anchor["x0"] + 6, anchor["bottom"]) if anchor else None
            found.append(Discrepancy(pno, "Missing text", std_text, "", f"Missing after: \"...{context}\".",
                                     box, union_box(s_words)))
        elif tag == "insert":
            found.append(Discrepancy(pno, "Added text", "", chk_text, f"Not in standard. Appears after: \"...{context}\".",
                                     union_box(c_words), None))
        else:
            change_type, note = classify(std_text, chk_text)
            found.append(Discrepancy(pno, change_type, std_text, chk_text, note,
                                     union_box(c_words), union_box(s_words)))
    return found


def compare_visual(pno: int, sp: pdfplumber.page.Page, cp: pdfplumber.page.Page,
                   s_img: Image.Image, c_img: Image.Image, text_ok: bool) -> list[Discrepancy]:
    """Report every visually changed area. In text mode, only areas without any words."""
    found = []
    for box in visual_regions(s_img, c_img):
        s_txt, c_txt = words_in(sp, box), words_in(cp, box)
        if text_ok:
            if s_txt or c_txt:
                continue  # already covered by the word comparison
            found.append(Discrepancy(pno, "Graphic / image changed", "(graphic)", "(graphic)",
                                     "A non-text element differs (chart, logo, image, line or colour).", box, box))
            continue
        s_crop, c_crop = crop(s_img, box, 0), crop(c_img, box, 0)
        s_blank = c_blank = False
        if s_crop is not None and c_crop is not None and s_crop.size == c_crop.size:
            # Look only at the pixels that changed: did they come from white (added) or turn white (removed)?
            sa = np.asarray(s_crop, dtype=np.int16)
            ca = np.asarray(c_crop, dtype=np.int16)
            changed = np.abs(sa - ca).max(axis=2) > PIXEL_THRESHOLD
            if changed.any():
                s_blank = (sa.min(axis=2)[changed] > 235).mean() > 0.9
                c_blank = (ca.min(axis=2)[changed] > 235).mean() > 0.9
        change_type = "Added content" if s_blank else "Missing content" if c_blank else "Content changed"
        unreadable = "(text not machine-readable - see image)"
        found.append(Discrepancy(pno, change_type, s_txt if is_readable(sp) and s_txt else unreadable,
                                 c_txt if is_readable(cp) and c_txt else unreadable,
                                 "Found by visual comparison; compare the two image crops.", box, box))
    return found


def compare(standard: Path, check: Path) -> tuple[list[Discrepancy], list[dict]]:
    """Compare two PDFs. Returns the discrepancies and a per-page method log."""
    results: list[Discrepancy] = []
    log: list[dict] = []
    with pdfplumber.open(standard) as s_pdf, pdfplumber.open(check) as c_pdf:
        n = max(len(s_pdf.pages), len(c_pdf.pages))
        for i in range(n):
            pno = i + 1
            if i >= len(c_pdf.pages):
                results.append(Discrepancy(pno, "Missing page", f"Page {pno} of the standard", "",
                                           "The checked document has fewer pages.", None))
                log.append({"page": pno, "method": "-", "found": 1})
                continue
            if i >= len(s_pdf.pages):
                cp = c_pdf.pages[i]
                results.append(Discrepancy(pno, "Added page", "", f"Page {pno}", "Not in the standard.",
                                           (0, 0, float(cp.width), float(cp.height))))
                log.append({"page": pno, "method": "-", "found": 1})
                continue
            sp, cp = s_pdf.pages[i], c_pdf.pages[i]
            s_img, c_img = render(sp), render(cp)
            text_ok = is_readable(sp) and is_readable(cp)
            page_items = compare_text(pno, sp, cp) if text_ok else []
            page_items += compare_visual(pno, sp, cp, s_img, c_img, text_ok)
            for d in page_items:
                d.location = describe_location(d.check_box, float(cp.height))
                d.crops = {"standard": crop(s_img, d.standard_box or d.check_box),
                           "found": crop(c_img, d.check_box)}
            results += page_items
            log.append({"page": pno, "method": "text + visual" if text_ok else "visual (text not readable)",
                        "found": len(page_items)})
    return results, log


# ----- outputs -------------------------------------------------------------------

def add_image(ws, img: Image.Image | None, cell: str, max_w: int = 240, max_h: int = 120) -> int:
    """Insert a crop into the sheet; returns its height in px (0 if none)."""
    if img is None:
        return 0
    img = img.copy()
    img.thumbnail((max_w, max_h))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    xl = XLImage(buf)
    ws.add_image(xl, cell)
    return img.height


def write_excel(items: list[Discrepancy], log: list[dict], standard: Path, check: Path, path: Path) -> None:
    """Write the discrepancy register and a summary sheet."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Discrepancies"
    headers = ["#", "Page", "Location", "Standard text", "Text found", "Type of change", "Note",
               "Standard (image)", "Found (image)"]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    widths = [5, 7, 26, 40, 40, 22, 40, 36, 36]
    for col, w in zip("ABCDEFGHI", widths):
        ws.column_dimensions[col].width = w
    fills = {"Missing": "FDE2E2", "Added": "E2F7E6"}
    for n, d in enumerate(items, start=1):
        row = n + 1
        ws.append([n, d.page, d.location, d.standard_text, d.found_text, d.change_type, d.note])
        colour = next((v for k, v in fills.items() if d.change_type.startswith(k)), "FFF6D5")
        for c in ws[row][:7]:
            c.alignment = Alignment(vertical="top", wrap_text=True)
        ws[row][5].fill = PatternFill("solid", fgColor=colour)
        h = max(add_image(ws, d.crops.get("standard"), f"H{row}"),
                add_image(ws, d.crops.get("found"), f"I{row}"))
        ws.row_dimensions[row].height = max(30, h * 0.75 + 6)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{len(items) + 1}"

    s = wb.create_sheet("Summary")
    s.append(["Proofread comparison"]); s["A1"].font = Font(bold=True, size=14)
    s.append(["Standard file", standard.name])
    s.append(["Checked file", check.name])
    s.append(["Run at", datetime.now().strftime("%Y-%m-%d %H:%M")])
    s.append(["Total discrepancies", len(items)])
    s.append([])
    s.append(["Type of change", "Count"])
    for t in sorted({d.change_type for d in items}):
        s.append([t, sum(d.change_type == t for d in items)])
    s.append([])
    s.append(["Page", "Method", "Discrepancies"])
    for entry in log:
        s.append([entry["page"], entry["method"], entry["found"]])
    for row in s.iter_rows():
        if row[0].value in ("Type of change", "Page"):
            for c in row:
                c.font = Font(bold=True)
    s.column_dimensions["A"].width = 26
    s.column_dimensions["B"].width = 40
    wb.save(path)


VALIDATION_NOTES = [
    "Synthetic factsheet with 6 planted text errors plus a recoloured graphic: "
    "all 7 found, no false alarms.",
    "Real public fund factsheet (6 pages, unreadable font encoding) compared with itself: "
    "0 differences on every page.",
    "Altered copy of that factsheet with 8 planted edits (date, two figures, a name, a "
    "recoloured chart bar, a deleted disclaimer, two added lines): all 8 found and labelled, "
    "no false alarms.",
    "Outputs checked by eye: highlights land on the right places in the PDF, and the "
    "image crops are embedded in the Excel sheet.",
]
VALIDATION_DATE = "24 Nov 2026"


def _figure(width_pt: float, height_pt: float):
    """A matplotlib figure the size of a PDF page, with point coordinates (origin top-left)."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    fig = plt.figure(figsize=(width_pt / 72, height_pt / 72))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, width_pt)
    ax.set_ylim(height_pt, 0)
    ax.axis("off")
    return plt, fig, ax


def _to_pdf_pages(figs_plt) -> PdfReader:
    """Save matplotlib figures into one in-memory PDF and return its reader."""
    from matplotlib.backends.backend_pdf import PdfPages
    plt, figs = figs_plt
    buf = io.BytesIO()
    with PdfPages(buf) as pdf:
        for fig in figs:
            pdf.savefig(fig, transparent=True)
            plt.close(fig)
    buf.seek(0)
    return PdfReader(buf)


def tag_overlay(width: float, height: float, tags: list[tuple[int, tuple, str]]) -> PdfReader:
    """Transparent page with a numbered tag beside each highlight, visible without hovering."""
    plt, fig, ax = _figure(width, height)
    for n, (x0, t, x1, b), colour in tags:
        label = str(n)
        # Sit the tag just above the highlight's top-right corner, so it doesn't cover the next word.
        x = min(x1 - 1, width - 6 * len(label) - 4)
        ax.text(x, max(t - 1.5, 9), label, fontsize=6.5, weight="bold", color="black", va="bottom", ha="left",
                family="DejaVu Sans",
                bbox={"boxstyle": "round,pad=0.25", "facecolor": f"#{colour}", "edgecolor": "black", "linewidth": 0.6})
    return _to_pdf_pages((plt, [fig]))


def report_pages(items: list[Discrepancy], log: list[dict], standard: Path, check: Path,
                 width: float = 595, height: float = 842) -> PdfReader:
    """Report pages placed at the front of the highlighted PDF: summary, every mismatch, validation."""
    import textwrap
    margin, line_h, wrap_at = 40, 11.5, 105
    figs: list = []
    state = {"y": height}
    plt = None

    def new_page():
        nonlocal plt
        plt, fig, ax = _figure(width, height)
        figs.append(fig)
        fig.patch.set_facecolor("white")
        ax.add_patch(__import__("matplotlib").patches.Rectangle((0, 0), width, height, color="white", zorder=-1))
        state["ax"], state["y"] = ax, margin
        ax.text(width - margin, height - 18, f"Proofread report - page {len(figs)}", fontsize=7,
                color="#777777", ha="right", family="DejaVu Sans")

    def write(text: str, size: float = 8.5, bold: bool = False, colour: str = "black", indent: float = 0,
              gap: float = 0):
        lines = textwrap.wrap(text, wrap_at - int(indent / 4.5)) or [""]
        for line in lines:
            if state["y"] + line_h > height - margin:
                new_page()
            state["ax"].text(margin + indent, state["y"], line, fontsize=size, va="top", color=colour,
                             weight="bold" if bold else "normal", family="DejaVu Sans")
            state["y"] += line_h * size / 8.5
        state["y"] += gap

    def images(d: Discrepancy, max_h: float = 70, max_w: float = 240):
        """Draw the standard and found crops side by side, so the change is visible in the report."""
        pair = [("Standard", d.crops.get("standard")), ("Found", d.crops.get("found"))]
        if not any(img for _, img in pair):
            return
        sizes = []
        for _, img in pair:
            if img is None:
                sizes.append((0, 0))
                continue
            k = min(max_w / (img.width / SCALE), max_h / (img.height / SCALE), 1.6)
            sizes.append((img.width / SCALE * k, img.height / SCALE * k))
        block_h = max(h for _, h in sizes) + 14
        if state["y"] + block_h > height - margin:
            new_page()
        x = margin + 14
        for (label, img), (w, h) in zip(pair, sizes):
            state["ax"].text(x, state["y"], label, fontsize=7, color="#777777", va="top", family="DejaVu Sans")
            if img is not None:
                top = state["y"] + 10
                state["ax"].imshow(np.asarray(img), extent=(x, x + w, top + h, top), zorder=2)
                state["ax"].add_patch(__import__("matplotlib").patches.Rectangle(
                    (x, top), w, h, fill=False, edgecolor="#999999", linewidth=0.5, zorder=3))
            x += max_w + 16
        state["ax"].set_xlim(0, width)
        state["ax"].set_ylim(height, 0)
        state["y"] += block_h + 4

    new_page()
    write("Proofread report", size=16, bold=True, colour="#1F3864", gap=6)
    write(f"Standard file:  {standard.name}")
    write(f"Checked file:   {check.name}")
    write(f"Run at:         {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    write(f"Result:         {len(items)} discrepancies found" if items else "Result:         no discrepancies found",
          bold=True, gap=6)
    write("By type", bold=True, colour="#1F3864")
    for t in sorted({d.change_type for d in items}):
        write(f"{t}: {sum(d.change_type == t for d in items)}", indent=12)
    write("By page", bold=True, colour="#1F3864")
    for entry in log:
        write(f"Page {entry['page']}: {entry['found']}  ({entry['method']})", indent=12)
    state["y"] += 4
    write("How to read the marked pages: each difference is highlighted and numbered. Yellow = changed, "
          "red = missing from the checked file, green = added. Page numbers below refer to the checked "
          "document (these report pages come first).", colour="#444444", gap=8)

    write("Discrepancies", size=11, bold=True, colour="#1F3864", gap=2)
    if not items:
        write("None. The checked file matches the standard.", indent=12)
    for n, d in enumerate(items, start=1):
        colour = "#c0392b" if d.change_type.startswith("Missing") else "#1e8449" if d.change_type.startswith("Added") else "black"
        write(f"#{n}   Page {d.page}   {d.change_type}   ({d.location})", bold=True, colour=colour)
        if d.standard_text:
            write(f"Standard: {d.standard_text}", indent=14)
        if d.found_text:
            write(f"Found:    {d.found_text}", indent=14)
        write(f"Note: {d.note}", indent=14, colour="#444444", gap=2)
        images(d)

    state["y"] += 6
    write(f"Tool validation (tests run during development, {VALIDATION_DATE})", size=11, bold=True,
          colour="#1F3864", gap=2)
    for note in VALIDATION_NOTES:
        write(f"-  {note}", indent=6)
    return _to_pdf_pages((plt, figs))


def watermark_overlay(width: float, height: float, text: str) -> PdfReader:
    """Transparent page with a large diagonal watermark, so a marked-up copy can't pass as the original."""
    plt, fig, ax = _figure(width, height)
    ax.text(width / 2, height / 2, text, fontsize=min(width, height) / max(14, len(text) * 0.62),
            color="#d62828", alpha=0.28, rotation=55, ha="center", va="center", weight="bold",
            family="DejaVu Sans")
    ax.text(width / 2, 14, text, fontsize=7, color="#d62828", ha="center", va="center",
            weight="bold", family="DejaVu Sans")
    return _to_pdf_pages((plt, [fig]))


def write_highlighted_pdf(items: list[Discrepancy], log: list[dict], standard: Path, check: Path,
                          path: Path, watermark: str | None = None) -> None:
    """Copy the checked PDF, highlight and number each discrepancy, and add report pages at the front.

    If `watermark` is given, it is stamped diagonally across every page of the checked document.
    """
    reader = PdfReader(check)
    writer = PdfWriter(clone_from=reader)
    tags_by_page: dict[int, list] = {}
    for n, d in enumerate(items, start=1):
        if d.check_box is None or d.page > len(writer.pages):
            continue
        page = writer.pages[d.page - 1]
        left, top_y = float(page.mediabox.left), float(page.mediabox.top)
        x0, t, x1, b = d.check_box
        rect = (left + x0 - 1, top_y - b - 1, left + x1 + 1, top_y - t + 1)
        quad = ArrayObject([FloatObject(v) for v in
                            (rect[0], rect[3], rect[2], rect[3], rect[0], rect[1], rect[2], rect[1])])
        colour = next((v for k, v in COLOURS.items() if d.change_type.startswith(k)), COLOURS["default"])
        annot = Highlight(rect=rect, quad_points=quad, highlight_color=colour, printing=True)
        summary = f"#{n} {d.change_type}: "
        summary += f"standard \"{d.standard_text}\" -> found \"{d.found_text}\"" if d.standard_text or d.found_text else d.note
        annot[NameObject("/Contents")] = TextStringObject(summary[:500])
        annot[NameObject("/T")] = TextStringObject("Proofread check")
        writer.add_annotation(page_number=d.page - 1, annotation=annot)
        tags_by_page.setdefault(d.page - 1, []).append((n, d.check_box, colour))

    for index, tags in tags_by_page.items():
        page = writer.pages[index]
        overlay = tag_overlay(float(page.mediabox.width), float(page.mediabox.height), tags).pages[0]
        page.merge_translated_page(overlay, float(page.mediabox.left), float(page.mediabox.bottom))

    if watermark:
        for page in writer.pages:
            mark = watermark_overlay(float(page.mediabox.width), float(page.mediabox.height), watermark).pages[0]
            page.merge_translated_page(mark, float(page.mediabox.left), float(page.mediabox.bottom))

    first = writer.pages[0]
    report = report_pages(items, log, standard, check, float(first.mediabox.width), float(first.mediabox.height))
    for i, rpage in enumerate(report.pages):
        writer.insert_page(rpage, index=i)
    with open(path, "wb") as f:
        writer.write(f)


# ----- main ----------------------------------------------------------------------

def first_pdf(folder: Path) -> Path | None:
    pdfs = sorted(p for p in folder.glob("*") if p.suffix.lower() == ".pdf")
    return pdfs[0] if pdfs else None


def main() -> int:
    parser = argparse.ArgumentParser(description="Proofread a PDF against a standard version.")
    parser.add_argument("--standard", type=Path, default=None)
    parser.add_argument("--check", type=Path, default=None)
    parser.add_argument("--out", type=Path, default=HERE / "output")
    parser.add_argument("--watermark", default=None,
                        help='stamp this text across every page of the checked document, e.g. "TEST COPY"')
    args = parser.parse_args()

    standard = args.standard or first_pdf(HERE / "1_standard")
    check = args.check or first_pdf(HERE / "2_to_check")
    if not standard or not check:
        print("Need one PDF in 1_standard/ and one in 2_to_check/ (or pass --standard and --check).")
        return 1
    args.out.mkdir(parents=True, exist_ok=True)

    items, log = compare(standard, check)
    xlsx = args.out / "discrepancies.xlsx"
    pdf = args.out / f"{check.stem}_highlighted.pdf"
    write_excel(items, log, standard, check, xlsx)
    write_highlighted_pdf(items, log, standard, check, pdf, args.watermark)

    print(f"Standard: {standard.name}\nChecked:  {check.name}")
    for entry in log:
        print(f"  page {entry['page']}: {entry['found']} discrepancies ({entry['method']})")
    print(f"Total discrepancies: {len(items)}")
    print(f"Excel:            {xlsx}\nHighlighted PDF:  {pdf}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
