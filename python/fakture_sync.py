# Fakture — LibreOffice Calc extension for invoice management.
# Copyright (C) 2026  Fortunacommerc d.o.o.
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""
fakture_sync — Read source .ods files and write to hidden _Sifrarnik sheet.

Reads 3 source files from <base>/Sifrarnik/:
  - proizvodi.ods
  - domaci_kupci.ods
  - ino_kupci.ods

Writes 3 sections to the hidden _Sifrarnik sheet and defines 6 Named Ranges.

Hidden sheet layout:
  [PROIZVODI] header + data
  (2 blank rows)
  [DOMACI] header + data
  (2 blank rows)
  [INO] header + data
"""

import os
from decimal import Decimal, ROUND_HALF_UP

import logging

import uno

_log = logging.getLogger("fakture.sync")

EUR_TO_BAM = 1.955830


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _round4(value):
    """Round to 4 decimal places (ROUND_HALF_UP)."""
    return float(Decimal(str(value)).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP))


def _generate_barcode(product_id):
    """Generate barcode: 00 + zeros + ID, min 8 max 14 digits."""
    prefix = "00"
    suffix = product_id
    total = len(prefix) + len(suffix)
    if total > 14:
        return ""
    pad = max(8 - total, 0)
    return prefix + "0" * pad + suffix


def _open_hidden_doc(path):
    """Open an ODS document hidden (Hidden=True) and return it."""
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    pv = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    pv.Name = "Hidden"
    pv.Value = True
    return desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(path), "_blank", 0, (pv,)
    )


def _read_sheet_data(doc, sheet_index=0):
    """Read all rows from the given sheet. Returns (headers, rows) where rows is a list of dicts."""
    sheet = doc.Sheets.getByIndex(sheet_index)

    # Read header row
    headers = []
    c = 0
    while True:
        h = sheet.getCellByPosition(c, 0).getString().strip()
        if not h:
            break
        headers.append(h)
        c += 1

    if not headers:
        return headers, []

    # Read data rows
    rows = []
    row = 1
    while True:
        first_val = sheet.getCellByPosition(0, row).getString().strip()
        if not first_val:
            break
        row_data = {}
        for ci, h in enumerate(headers):
            cell = sheet.getCellByPosition(ci, row)
            row_data[h] = cell.getString().strip()
            # Store numeric value for price columns
            row_data["_num_" + h] = cell.getValue()
        rows.append(row_data)
        row += 1

    return headers, rows


# ─────────────────────────────────────────────────────────────────────────────
# Source file loaders
# ─────────────────────────────────────────────────────────────────────────────

def _load_products(filepath):
    """Load products from proizvodi.ods. Returns list of dicts."""
    doc = _open_hidden_doc(filepath)
    try:
        headers, rows = _read_sheet_data(doc)
    finally:
        doc.close(False)

    products = []
    for r in rows:
        pid = r.get("ID", "")
        bar_kod = r.get("Bar kod", "")
        naziv = r.get("Naziv", "")
        jm = r.get("Jed. Mjere", "")
        cijena_raw = r.get("_num_Cijena bez PDV-a", 0.0)
        valuta = r.get("Valuta", "")

        if not naziv:
            continue

        # EUR/BAM conversion
        if valuta.upper() == "EUR":
            cijena_eur = cijena_raw
            cijena_bam = _round4(cijena_raw * EUR_TO_BAM)
        else:
            cijena_bam = cijena_raw
            cijena_eur = _round4(cijena_raw / EUR_TO_BAM) if cijena_raw else 0.0

        # Auto-generate barcode if missing
        if not bar_kod and pid:
            bar_kod = _generate_barcode(pid)

        products.append({
            "id": pid,
            "bar_kod": bar_kod,
            "naziv": naziv,
            "jed_mjere": jm,
            "cijena_bam": cijena_bam,
            "cijena_eur": cijena_eur,
        })
    return products


def _load_domestic_customers(filepath):
    """Load domestic customers from domaci_kupci.ods. Returns list of dicts."""
    doc = _open_hidden_doc(filepath)
    try:
        headers, rows = _read_sheet_data(doc)
    finally:
        doc.close(False)

    customers = []
    for r in rows:
        firma = r.get("Firma", "")
        if not firma:
            continue
        jib = r.get("JIB", "")
        pdv = r.get("PDV", "")
        if jib:
            buyer_id = "VP:" + jib
        elif pdv:
            buyer_id = "VP:4" + pdv
        else:
            buyer_id = ""
        pb = r.get("PB", "")
        grad = r.get("Grad", "")
        pb_grad = (pb + " " + grad).strip()
        customers.append({
            "firma": firma,
            "podruznica": r.get("Podruznica", ""),
            "ulica": r.get("Ulica", ""),
            "pb_grad": pb_grad,
            "jib": jib,
            "pdv": pdv,
            "buyer_id": buyer_id,
        })
    return customers


def _load_foreign_customers(filepath):
    """Load foreign customers from ino_kupci.ods. Returns list of dicts."""
    doc = _open_hidden_doc(filepath)
    try:
        headers, rows = _read_sheet_data(doc)
    finally:
        doc.close(False)

    customers = []
    for r in rows:
        firma = r.get("Firma", "")
        if not firma:
            continue
        pb = r.get("PB", "")
        grad = r.get("Grad", "")
        pb_grad = (pb + " " + grad).strip()
        customers.append({
            "firma": firma,
            "ulica": r.get("Ulica", ""),
            "pb_grad": pb_grad,
            "drzava": r.get("Drzava", ""),
            "vat": r.get("VAT", ""),
            "buyer_id": "VP:9999999999999",
        })
    return customers


# ─────────────────────────────────────────────────────────────────────────────
# Hidden sheet writer
# ─────────────────────────────────────────────────────────────────────────────

SHEET_NAME = "_Sifrarnik"


def _ensure_hidden_sheet(doc):
    """Create or clear the _Sifrarnik sheet. Returns the sheet object."""
    sheets = doc.Sheets
    if sheets.hasByName(SHEET_NAME):
        sheet = sheets.getByName(SHEET_NAME)
        # Clear contents but keep the sheet (Named Ranges remain valid)
        # clearContents flags: 7 = VALUE(1) + DATETIME(2) + STRING(4)
        used = sheet.createCursor()
        used.gotoStartOfUsedArea(False)
        used.gotoEndOfUsedArea(True)
        used.clearContents(7)
    else:
        sheets.insertNewByName(SHEET_NAME, sheets.Count)
        sheet = sheets.getByName(SHEET_NAME)

    sheet.IsVisible = False
    return sheet


def _write_section(sheet, start_row, label, headers, data, value_cols=None):
    """
    Write a section to the sheet.
    start_row: first row (0-based)
    label: string for the [LABEL] marker row
    headers: list of header strings
    data: list of dicts with row data
    value_cols: set of header names whose values should be written as setValue (numeric)
    Returns (data_start_row, data_end_row) — data rows (excluding header).
    """
    if value_cols is None:
        value_cols = set()

    # [LABEL] row
    sheet.getCellByPosition(0, start_row).setString("[" + label + "]")

    # Header row
    header_row = start_row + 1
    for ci, h in enumerate(headers):
        sheet.getCellByPosition(ci, header_row).setString(h)

    # Data rows
    data_start = header_row + 1
    for ri, row_data in enumerate(data):
        row_idx = data_start + ri
        for ci, h in enumerate(headers):
            cell = sheet.getCellByPosition(ci, row_idx)
            val = row_data.get(h, "")
            if h in value_cols and isinstance(val, (int, float)):
                cell.setValue(val)
            else:
                cell.setString(str(val) if val else "")

    data_end = data_start + len(data) - 1 if data else data_start - 1
    return data_start, data_end


def _define_named_ranges(doc, ranges_info):
    """
    Create or update Named Ranges.
    If a range already exists, update it with setContent(). Otherwise create new.
    ranges_info: list of (name, sheet_index, col_start, row_start, col_end, row_end)
    """
    named_ranges = doc.NamedRanges

    for name, sheet_idx, col_start, row_start, col_end, row_end in ranges_info:
        if row_start > row_end:
            # No data — remove if exists, don't create
            if named_ranges.hasByName(name):
                named_ranges.removeByName(name)
            continue

        base_cell = uno.createUnoStruct("com.sun.star.table.CellAddress")
        base_cell.Sheet = sheet_idx
        base_cell.Column = col_start
        base_cell.Row = row_start

        # Absolute reference string: $'_Sifrarnik'.$A$3:$F$10
        col_start_letter = _col_letter(col_start)
        col_end_letter = _col_letter(col_end)
        content = "$'{}'.${}${}:${}${}".format(
            SHEET_NAME,
            col_start_letter, row_start + 1,
            col_end_letter, row_end + 1,
        )

        if named_ranges.hasByName(name):
            existing = named_ranges.getByName(name)
            existing.setContent(content)
            existing.setReferencePosition(base_cell)
        else:
            named_ranges.addNewByName(name, content, base_cell, 0)


def _col_letter(col_idx):
    """Convert 0-based column index to letter(s): 0=A, 1=B, ..., 25=Z, 26=AA."""
    result = ""
    idx = col_idx
    while True:
        result = chr(65 + idx % 26) + result
        idx = idx // 26 - 1
        if idx < 0:
            break
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Main sync
# ─────────────────────────────────────────────────────────────────────────────

def sync_to_hidden_sheet(doc, base_path):
    """
    Read 3 source files from <base_path>/Sifrarnik/ and write to the _Sifrarnik sheet.
    Creates 6 Named Ranges for dropdowns and VLOOKUP.
    Returns (n_products, n_domestic, n_foreign).
    """
    _log.info("sync start: base_path=%s", base_path)
    sifrarnik_dir = os.path.join(base_path, "Sifrarnik")

    products_file = os.path.join(sifrarnik_dir, "proizvodi.ods")
    domaci_file   = os.path.join(sifrarnik_dir, "domaci_kupci.ods")
    ino_file      = os.path.join(sifrarnik_dir, "ino_kupci.ods")

    products = _load_products(products_file) if os.path.exists(products_file) else []
    domaci   = _load_domestic_customers(domaci_file) if os.path.exists(domaci_file) else []
    ino      = _load_foreign_customers(ino_file) if os.path.exists(ino_file) else []

    # Prepare sheet
    sheet = _ensure_hidden_sheet(doc)
    sheet_idx = doc.Sheets.Count - 1
    for i in range(doc.Sheets.Count):
        if doc.Sheets.getByIndex(i).Name == SHEET_NAME:
            sheet_idx = i
            break

    # ── Section: PROIZVODI ──
    prod_headers = ["Naziv", "ID", "Bar kod", "Jed. Mjere", "Cijena BAM", "Cijena EUR"]
    prod_data = [
        {
            "Naziv": p["naziv"],
            "ID": p["id"],
            "Bar kod": p["bar_kod"],
            "Jed. Mjere": p["jed_mjere"],
            "Cijena BAM": p["cijena_bam"],
            "Cijena EUR": p["cijena_eur"],
        }
        for p in products
    ]
    prod_data_start, prod_data_end = _write_section(
        sheet, 0, "PROIZVODI", prod_headers, prod_data,
        value_cols={"Cijena BAM", "Cijena EUR"},
    )

    # ── Section: DOMACI (2 blank rows after products) ──
    domaci_start_row = prod_data_end + 3 if products else 4
    domaci_headers = ["Firma", "Podruznica", "Ulica", "PB Grad", "JIB", "PDV", "BuyerID"]
    domaci_data = [
        {
            "Firma": c["firma"],
            "Podruznica": c["podruznica"],
            "Ulica": c["ulica"],
            "PB Grad": c["pb_grad"],
            "JIB": c["jib"],
            "PDV": c["pdv"],
            "BuyerID": c["buyer_id"],
        }
        for c in domaci
    ]
    domaci_data_start, domaci_data_end = _write_section(
        sheet, domaci_start_row, "DOMACI", domaci_headers, domaci_data,
    )

    # ── Section: INO (2 blank rows after domestic) ──
    ino_start_row = domaci_data_end + 3 if domaci else domaci_start_row + 4
    ino_headers = ["Firma", "Ulica", "PB Grad", "Drzava", "VAT", "BuyerID"]
    ino_data = [
        {
            "Firma": c["firma"],
            "Ulica": c["ulica"],
            "PB Grad": c["pb_grad"],
            "Drzava": c["drzava"],
            "VAT": c["vat"],
            "BuyerID": c["buyer_id"],
        }
        for c in ino
    ]
    ino_data_start, ino_data_end = _write_section(
        sheet, ino_start_row, "INO", ino_headers, ino_data,
    )

    # ── Named Ranges ──
    # (name, sheet_idx, col_start, row_start, col_end, row_end)
    ranges = [
        # Products — all rows, all columns (Naziv..Cijena EUR = 0..5)
        ("Proizvodi",       sheet_idx, 0, prod_data_start,   5, prod_data_end),
        # Products — column A (Naziv) only
        ("ProizvodiNazivi", sheet_idx, 0, prod_data_start,   0, prod_data_end),
        # Domestic customers — all rows, all columns (Firma..BuyerID = 0..6)
        ("DomaciKupci",       sheet_idx, 0, domaci_data_start, 6, domaci_data_end),
        # Domestic customers — column A (Firma) only
        ("DomaciKupciNazivi", sheet_idx, 0, domaci_data_start, 0, domaci_data_end),
        # Foreign customers — all rows, all columns (Firma..BuyerID = 0..5)
        ("InoKupci",       sheet_idx, 0, ino_data_start, 5, ino_data_end),
        # Foreign customers — column A (Firma) only
        ("InoKupciNazivi", sheet_idx, 0, ino_data_start, 0, ino_data_end),
    ]
    _define_named_ranges(doc, ranges)

    # Remove obsolete Named Ranges from previous versions
    named_ranges = doc.NamedRanges
    for obsolete in ("ProizvodiSifre",):
        if named_ranges.hasByName(obsolete):
            named_ranges.removeByName(obsolete)

    _log.info("sync done: products=%d, domestic=%d, foreign=%d",
              len(products), len(domaci), len(ino))
    return len(products), len(domaci), len(ino)
