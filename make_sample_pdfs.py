"""Generate a FICTIONAL sample factsheet and an edited copy with planted errors.

Lets anyone try the proofreading tool without real documents. "Example Global
Equity Fund" and every name and number in it are made up.

Planted differences in the edited copy (page 1 unless noted):
  1. Date 2026 -> 2025
  2. Fund size 1,245.2 -> 1,425.2 (digits swapped)
  3. ", guaranteed" added to the returns sentence
  4. "Kumar" -> "kumar" (capitalisation)
  5. Benchmark "World" -> "ACWI"
  6. Risk warning "Past performance is not a guide to future returns." removed
  7. Page 2: chart box recoloured (graphic change)

Usage:
    python make_sample_pdfs.py            # writes examples/sample_standard.pdf and examples/sample_to_check.pdf
"""

from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

HERE = Path(__file__).resolve().parent

STANDARD_LINES = [
    "Example Global Equity Fund - Monthly Factsheet (fictional)",
    "Data as at 31 October 2026",
    "The fund aims to provide long-term growth of capital by investing in global shares.",
    "Fund size: USD 1,245.2 million. Ongoing charges: 0.85%.",
    "One year return: 12.4% and three year return: 7.3% per year.",
    "Portfolio manager: Jane Smith and Ravi Kumar.",
    "Benchmark: MSCI World Index.",
    "Past performance is not a guide to future returns.",
    "The value of investments can go down as well as up.",
]


def edited_lines() -> list[str]:
    """The standard text with six planted errors."""
    lines = list(STANDARD_LINES)
    lines[1] = "Data as at 31 October 2025"
    lines[3] = "Fund size: USD 1,425.2 million. Ongoing charges: 0.85%."
    lines[4] = "One year return: 12.4% and three year return: 7.3% per year, guaranteed."
    lines[5] = "Portfolio manager: Jane Smith and Ravi kumar."
    lines[6] = "Benchmark: MSCI ACWI Index."
    del lines[7]
    return lines


def make_pdf(lines: list[str], path: Path, chart_colour: str) -> None:
    """Two A4 pages: page 1 text, page 2 a short note and a coloured chart box."""
    with PdfPages(path) as pdf:
        for page in range(2):
            fig = plt.figure(figsize=(8.27, 11.69))
            page_lines = lines if page == 0 else ["Page two: risk profile and portfolio breakdown for investors."]
            for k, line in enumerate(page_lines):
                fig.text(0.08, 0.92 - k * 0.05, line, fontsize=11, family="DejaVu Sans")
            if page == 1:
                fig.patches.append(plt.Rectangle((0.1, 0.4), 0.3, 0.1, transform=fig.transFigure, color=chart_colour))
            pdf.savefig(fig)
            plt.close(fig)


def main() -> None:
    out = HERE / "examples"
    out.mkdir(exist_ok=True)
    make_pdf(STANDARD_LINES, out / "sample_standard.pdf", "navy")
    make_pdf(edited_lines(), out / "sample_to_check.pdf", "darkred")
    print(f"Wrote {out / 'sample_standard.pdf'} and {out / 'sample_to_check.pdf'}")


if __name__ == "__main__":
    main()
