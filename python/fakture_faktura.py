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
fakture_faktura — Invoice creation from templates with auto-numbering.

Flow:
  1. Check config (base path)
  2. get_next_rb() — scan base folder for next sequential number
  3. show_identifier_dialog() — user enters identifier
  4. Copy template from Obrasci/
  5. Open the copy
  6. sync_to_hidden_sheet() — automatic import
  7. doc.store() — save
"""

import os
import re
import shutil

import logging

import uno

import fakture_config
import fakture_dialogs
import fakture_sync

_log = logging.getLogger("fakture.faktura")


# ─────────────────────────────────────────────────────────────────────────────
# Sequential number
# ─────────────────────────────────────────────────────────────────────────────

def get_next_rb(base_path):
    r"""
    Scan base folder (not subdirs) for files matching:
      Faktura-{RB}-{GG}__.*\.ods
    Returns (next_rb, year_suffix).
    """
    year_suffix = fakture_config.detect_year_from_folder(base_path)
    pattern = re.compile(
        rf"^Faktura-(\d+)-{re.escape(year_suffix)}__.*\.ods$"
    )
    max_rb = 0

    if os.path.exists(base_path):
        for fname in os.listdir(base_path):
            match = pattern.match(fname)
            if match:
                rb = int(match.group(1))
                max_rb = max(max_rb, rb)

    return max_rb + 1, year_suffix


# ─────────────────────────────────────────────────────────────────────────────
# Identifier sanitization
# ─────────────────────────────────────────────────────────────────────────────

def sanitize_identifier(text):
    """
    Sanitize identifier for use in filename:
    - Spaces → _
    - Allowed: letters (incl. šćčžđ), digits, dash, underscore
    - Max 50 characters
    """
    text = text.strip()
    text = text.replace(" ", "_")
    allowed = re.compile(r'[a-zA-ZšćčžđŠĆČŽĐ0-9_\-]')
    text = "".join(c for c in text if allowed.match(c))
    return text[:50]


# ─────────────────────────────────────────────────────────────────────────────
# Template type mapping
# ─────────────────────────────────────────────────────────────────────────────

TEMPLATE_MAP = {
    "domaci":   "faktura_domaci.ods",
    "ino":      "faktura_ino.ods",
}


# ─────────────────────────────────────────────────────────────────────────────
# Invoice creation
# ─────────────────────────────────────────────────────────────────────────────

def create_invoice(ctx, tip, frame=None):
    """Create a new invoice of the given type (domaci/ino)."""
    _log.info("create_invoice: type=%s", tip)

    # 1. Check config
    base_path = fakture_config.get_base_path()
    if not base_path:
        fakture_dialogs.show_msgbox(ctx,
            "Bazni folder nije podešen.\n\nOtvoriće se podešavanja.",
            "Fakture", "info", frame=frame)
        if not fakture_dialogs.show_settings(ctx, frame):
            return
        base_path = fakture_config.get_base_path()
        if not base_path:
            return

    # 2. Check that folder exists
    if not os.path.isdir(base_path):
        fakture_dialogs.show_msgbox(ctx,
            "Bazni folder ne postoji:\n" + base_path +
            "\n\nProvjerite podešavanja.",
            "Fakture — Greška", "error", frame=frame)
        return

    # 3. Next sequential number
    next_rb, year_suffix = get_next_rb(base_path)

    # 4. Identifier dialog
    identifier = fakture_dialogs.show_identifier_dialog(ctx, next_rb, frame)
    if identifier is None:
        _log.info("create_invoice: user cancelled")
        return

    # 5. Sanitize
    identifier = sanitize_identifier(identifier)
    if not identifier:
        fakture_dialogs.show_msgbox(ctx,
            "Identifikator je prazan ili sadrži samo nedozvoljene znakove.",
            "Fakture — Greška", "error", frame=frame)
        return

    # 6. Check template
    template_name = TEMPLATE_MAP.get(tip)
    if not template_name:
        fakture_dialogs.show_msgbox(ctx,
            "Nepoznat tip fakture: " + str(tip),
            "Fakture — Greška", "error", frame=frame)
        return

    template_path = os.path.join(base_path, "Obrasci", template_name)
    if not os.path.exists(template_path):
        fakture_dialogs.show_msgbox(ctx,
            "Obrazac '{}' nije pronađen u folderu Obrasci/".format(template_name),
            "Fakture — Greška", "error", frame=frame)
        return

    # 7. Copy template
    dest_name = "Faktura-{}-{}__{}.ods".format(next_rb, year_suffix, identifier)
    dest_path = os.path.join(base_path, dest_name)
    _log.info("copying template: %s -> %s", template_name, dest_name)

    try:
        shutil.copy2(template_path, dest_path)
    except Exception as e:
        _log.error("template copy failed: %s", e, exc_info=True)
        fakture_dialogs.show_msgbox(ctx,
            "Greška pri kopiranju obrasca:\n" + str(e),
            "Fakture — Greška", "error", frame=frame)
        return

    # 8. Open the copy
    try:
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        file_url = uno.systemPathToFileUrl(dest_path)
        doc = desktop.loadComponentFromURL(file_url, "_blank", 0, ())
    except Exception as e:
        _log.error("failed to open invoice: %s", e, exc_info=True)
        fakture_dialogs.show_msgbox(ctx,
            "Greška pri otvaranju fakture:\n" + str(e),
            "Fakture — Greška", "error", frame=frame)
        return

    # 9. Auto-sync
    try:
        fakture_sync.sync_to_hidden_sheet(doc, base_path)
    except Exception as e:
        _log.error("sync failed: %s", e, exc_info=True)
        fakture_dialogs.show_msgbox(ctx,
            "Šifrarnik uvezen, ali sa greškom:\n" + str(e),
            "Fakture — Upozorenje", "info", frame=frame)

    # 10. Save
    try:
        doc.store()
    except Exception as e:
        _log.warning("save failed: %s", e)

    _log.info("create_invoice done: %s", dest_name)
