#!/usr/bin/env python3
"""Assemble the final lease packet by appending static PDFs to the generated lease.

Usage:
    python3 assemble_packet.py <lease.pdf> <property.json> <tenancy.json> <output.pdf>

What gets appended (in order):
  1. Mold pamphlet                — always
  2. Lead paint pamphlet          — when property["lead_paint_disclosure_required"] is true
  3. Trash calendar               — when tenancy["lease_kind"] == "new" and
                                    property["utilities"]["trash"]["calendar_pdf_path"]
                                    points to an existing file

The static pamphlets live in  <repo_root>/static/.
"""

import json
import sys
from pathlib import Path

from pypdf import PdfReader, PdfWriter


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

# build_scripts/ → parent → repo root
REPO_ROOT = Path(__file__).resolve().parent.parent
STATIC_DIR = REPO_ROOT / "static"

MOLD_PAMPHLET    = STATIC_DIR / "mold-pamphlet.pdf"
LEAD_PAMPHLET    = STATIC_DIR / "lead-paint-pamphlet.pdf"


# ---------------------------------------------------------------------------
# Core function
# ---------------------------------------------------------------------------

def assemble(lease_pdf: Path, property_: dict, tenancy: dict, output_path: Path) -> Path:
    """Merge lease PDF with applicable static pamphlets.

    Parameters
    ----------
    lease_pdf    : Path to the already-converted lease PDF.
    property_    : Parsed property JSON dict.
    tenancy      : Parsed tenancy JSON dict.
    output_path  : Destination path for the assembled packet PDF.

    Returns
    -------
    Path to the written packet PDF (same as output_path).

    Raises
    ------
    FileNotFoundError if lease_pdf or a required static asset is missing.
    """
    lease_pdf = Path(lease_pdf)
    output_path = Path(output_path)

    if not lease_pdf.exists():
        raise FileNotFoundError(f"Lease PDF not found: {lease_pdf}")
    if not MOLD_PAMPHLET.exists():
        raise FileNotFoundError(
            f"Mold pamphlet not found at {MOLD_PAMPHLET}. "
            "Place mold-pamphlet.pdf in the static/ directory."
        )

    # Decide which PDFs to append
    attachments: list[Path] = []

    # 1. Mold pamphlet — always
    attachments.append(MOLD_PAMPHLET)

    # 2. Lead paint pamphlet — property-driven
    if property_.get("lead_paint_disclosure_required"):
        if not LEAD_PAMPHLET.exists():
            raise FileNotFoundError(
                f"Lead paint pamphlet not found at {LEAD_PAMPHLET}. "
                "Place lead-paint-pamphlet.pdf in the static/ directory."
            )
        attachments.append(LEAD_PAMPHLET)

    # 3. Trash calendar — new leases only, when a local PDF has been saved
    if tenancy.get("lease_kind") == "new":
        calendar_path = (
            property_
            .get("utilities", {})
            .get("trash", {})
            .get("calendar_pdf_path")
        )
        if calendar_path:
            calendar_pdf = Path(calendar_path)
            if not calendar_pdf.is_absolute():
                calendar_pdf = REPO_ROOT / calendar_pdf
            if calendar_pdf.exists():
                attachments.append(calendar_pdf)
            else:
                print(
                    f"[assemble_packet] Warning: trash calendar PDF listed in property "
                    f"profile ({calendar_pdf}) not found — skipping.",
                    file=sys.stderr,
                )

    # Merge
    writer = PdfWriter()
    for pdf_path in [lease_pdf] + attachments:
        reader = PdfReader(str(pdf_path))
        for page in reader.pages:
            writer.add_page(page)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)

    pamphlet_names = [p.stem for p in attachments]
    print(
        f"[assemble_packet] Wrote {output_path.name} "
        f"({_page_count(output_path)} pages, appended: {', '.join(pamphlet_names)})"
    )
    return output_path


def _page_count(pdf_path: Path) -> int:
    return len(PdfReader(str(pdf_path)).pages)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) != 5:
        print(
            "Usage: assemble_packet.py <lease.pdf> <property.json> "
            "<tenancy.json> <output.pdf>"
        )
        sys.exit(2)

    lease_pdf    = Path(sys.argv[1])
    property_    = json.loads(Path(sys.argv[2]).read_text())
    tenancy      = json.loads(Path(sys.argv[3]).read_text())
    output_path  = Path(sys.argv[4])

    assemble(lease_pdf, property_, tenancy, output_path)


if __name__ == "__main__":
    main()
