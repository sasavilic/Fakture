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
fakture_dialogs — UNO dialogs for the Fakture extension.

Functions:
  show_msgbox             — info/error/question message box (5-param LO 26.x)
  show_settings           — base folder settings (XDL: Settings.xdl)
  show_identifier_dialog  — invoice identifier input (XDL: IdentifierDialog.xdl)
  show_folder_picker      — system folder picker
"""

import logging

import uno
import unohelper
from com.sun.star.awt import XActionListener

import fakture_config

_log = logging.getLogger("fakture.dialogs")

_EXT_ID = "com.fortunacommerc.fakture"


# ─────────────────────────────────────────────────────────────────────────────
# Dialog loader
# ─────────────────────────────────────────────────────────────────────────────

def _load_dialog(ctx, frame, dialog_name):
    """Load XDL dialog from extension package, center, return dialog."""
    dp = ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    dialog = dp.createDialog(
        f"vnd.sun.star.extension://{_EXT_ID}/dialogs/{dialog_name}.xdl")
    parent = frame.getContainerWindow()
    toolkit = parent.getToolkit()
    dialog.createPeer(toolkit, parent)
    parent_size = parent.getSize()
    dialog_size = dialog.getSize()
    x = max(0, (parent_size.Width - dialog_size.Width) // 2)
    y = max(0, (parent_size.Height - dialog_size.Height) // 2)
    dialog.setPosSize(x, y, dialog_size.Width, dialog_size.Height,
                      uno.getConstantByName("com.sun.star.awt.PosSize.POS"))
    return dialog


# ─────────────────────────────────────────────────────────────────────────────
# Message box
# ─────────────────────────────────────────────────────────────────────────────

def show_msgbox(ctx, message, title="Fakture", msgtype="info", frame=None):
    """
    Show a LibreOffice message box (5-param API, LO 26.x).
    msgtype: "info" | "error" | "question"
    Returns result code (2=YES, 1=OK, 0=Cancel/No).
    """
    type_map = {"info": 1, "error": 3, "question": 2}
    btn_map  = {"info": 1, "error": 1, "question": 3}
    smgr = ctx.ServiceManager
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    parent = frame.getContainerWindow() if frame is not None else None
    msgbox = toolkit.createMessageBox(
        parent, type_map.get(msgtype, 1), btn_map.get(msgtype, 1), title, message)
    return msgbox.execute()


# ─────────────────────────────────────────────────────────────────────────────
# Settings dialog
# ─────────────────────────────────────────────────────────────────────────────

def show_settings(ctx, frame=None):
    """
    Show the base folder settings dialog (Settings.xdl).
    Returns True if user saved, False if cancelled.
    """
    current_path = fakture_config.get_base_path()
    dialog = _load_dialog(ctx, frame, "Settings")
    dialog.getControl("txtPath").setText(current_path)

    class _BrowseListener(unohelper.Base, XActionListener):
        def actionPerformed(self, ev):
            picked = show_folder_picker(ctx, "Izaberite bazni folder")
            if picked:
                dialog.getControl("txtPath").setText(picked)
        def disposing(self, ev):
            pass

    dialog.getControl("btnBrowse").addActionListener(_BrowseListener())

    result_code = dialog.execute()
    if result_code == 1:
        new_path = dialog.getControl("txtPath").getText().strip()
        dialog.dispose()
        if new_path:
            fakture_config.set_base_path(new_path)
            created = fakture_config.ensure_folder_structure(new_path)
            if created:
                show_msgbox(ctx,
                    "Kreirani podfolderi: " + ", ".join(created),
                    "Fakture", frame=frame)
        return True
    else:
        dialog.dispose()
        return False


# ─────────────────────────────────────────────────────────────────────────────
# Identifier dialog
# ─────────────────────────────────────────────────────────────────────────────

def show_identifier_dialog(ctx, next_rb, frame=None):
    """
    Show the invoice identifier input dialog (IdentifierDialog.xdl).
    Returns identifier string or None if cancelled.
    """
    dialog = _load_dialog(ctx, frame, "IdentifierDialog")
    dialog.getControl("lblInfo").setText("Sljedeći redni broj: " + str(next_rb))

    result_code = dialog.execute()
    if result_code == 1:
        identifier = dialog.getControl("txtIdentifier").getText().strip()
        dialog.dispose()
        return identifier if identifier else None
    else:
        dialog.dispose()
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Folder picker
# ─────────────────────────────────────────────────────────────────────────────

def show_folder_picker(ctx, title="Izaberite folder"):
    """Show system folder picker. Returns path or None."""
    smgr = ctx.ServiceManager
    picker = smgr.createInstanceWithContext("com.sun.star.ui.dialogs.FolderPicker", ctx)
    picker.setTitle(title)
    result = picker.execute()
    if result == 1:
        url = picker.getDirectory()
        return uno.fileUrlToSystemPath(url)
    return None
