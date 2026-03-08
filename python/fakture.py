#!/usr/bin/env python3
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
fakture.py — LibreOffice UNO ProtocolHandler for the Fakture extension.

Protocol: vnd.fortunacommerc.fakture:*
Menu:     Fakture > ...

Commands:
  nova_faktura_domaci  — new domestic invoice
  nova_faktura_ino     — new foreign invoice
  sync                 — refresh codebook in the active document
  open_domaci_kupci    — open domestic customers codebook
  open_ino_kupci       — open foreign customers codebook
  open_proizvodi       — open products codebook
  settings             — system settings
"""

import gzip
import os
import shutil
import sys
import logging
import logging.handlers
import platform

# Add the extension's python/ directory to sys.path so fakture_*.py modules are importable
_EXT_PYTHON_DIR = os.path.dirname(os.path.abspath(__file__))
if _EXT_PYTHON_DIR not in sys.path:
    sys.path.insert(0, _EXT_PYTHON_DIR)

# ──────────────────────────────────────────────────────────────────────────────
# Logging
# ──────────────────────────────────────────────────────────────────────────────

_EXT_NAME = "fakture"
_LOGGING_CONFIG_NODE = "/com.fortunacommerc.fakture/Logging"


def _setup_logger():
    if platform.system() == "Windows":
        log_dir = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), _EXT_NAME)
    else:
        log_dir = os.path.join(os.path.expanduser("~"), ".config", _EXT_NAME)
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "extension.log")
    logger = logging.getLogger(_EXT_NAME)
    if logger.handlers:
        return logger
    handler = logging.handlers.TimedRotatingFileHandler(
        log_file, when="W0", backupCount=0, encoding="utf-8")

    def _rotator(source, dest):
        with open(source, "rb") as f_in, gzip.open(dest, "wb") as f_out:
            shutil.copyfileobj(f_in, f_out)
        os.remove(source)

    handler.rotator = _rotator
    handler.namer = lambda name: name + ".gz"

    handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s"))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    return logger


log = _setup_logger()

log.info("Logger ready")

import uno
import unohelper
from com.sun.star.frame import XDispatch, XDispatchProvider
from com.sun.star.lang import XServiceInfo, XInitialization

log.info("Loading fakture_config")
import fakture_config
log.info("Loading fakture_dialogs")
import fakture_dialogs
log.info("Loading fakture_faktura")
import fakture_faktura
log.info("Loading fakture_sync")
import fakture_sync
log.info("Internal modules loaded")




def _apply_log_level(ctx):
    try:
        pv = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        pv.Name = "nodepath"
        pv.Value = _LOGGING_CONFIG_NODE
        cp = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx)
        node = cp.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationAccess", (pv,))
        level_name = node.getByName("LogLevel") or "INFO"
        log.setLevel(getattr(logging, level_name.upper(), logging.INFO))
    except Exception as e:
        log.warning("Could not read LogLevel from registry: %s", e)


# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────

DOC_MAP = {
    "domaci_kupci": "Sifrarnik/domaci_kupci.ods",
    "ino_kupci":    "Sifrarnik/ino_kupci.ods",
    "proizvodi":    "Sifrarnik/proizvodi.ods",
}


def _open_or_focus(ctx, file_path):
    """Open a file or focus its window if already open."""
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    file_url = uno.systemPathToFileUrl(file_path)

    components = desktop.getComponents()
    if components:
        enum = components.createEnumeration()
        while enum.hasMoreElements():
            comp = enum.nextElement()
            try:
                if comp.getURL() == file_url:
                    frame = comp.getCurrentController().getFrame()
                    cw = frame.getContainerWindow()
                    cw.setFocus()
                    frame.activate()
                    return
            except Exception:
                continue

    desktop.loadComponentFromURL(file_url, "_blank", 0, ())


# ──────────────────────────────────────────────────────────────────────────────
# Command implementations
# ──────────────────────────────────────────────────────────────────────────────

def _cmd_nova_faktura(ctx, frame, tip):
    log.info("nova_faktura: type=%s", tip)
    fakture_faktura.create_invoice(ctx, tip, frame)


def _cmd_sync(ctx, frame):
    log.info("sync")
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

    doc = frame.getController().getModel()
    if doc is None:
        fakture_dialogs.show_msgbox(ctx,
            "Nema otvorenog dokumenta.", "Fakture — Greška", "error", frame=frame)
        return

    try:
        _ = doc.Sheets
    except Exception:
        fakture_dialogs.show_msgbox(ctx,
            "Aktivni dokument nije Calc spreadsheet.",
            "Fakture — Greška", "error", frame=frame)
        return

    n_prod, n_dom, n_ino = fakture_sync.sync_to_hidden_sheet(doc, base_path)
    doc.store()
    fakture_dialogs.show_msgbox(ctx,
        "Šifrarnik osvježen.\n\n"
        "Proizvodi: {}\nDomaći kupci: {}\nIno kupci: {}".format(n_prod, n_dom, n_ino),
        "Fakture", frame=frame)


def _cmd_open(ctx, frame, doc_name):
    log.info("open: %s", doc_name)
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

    rel_path = DOC_MAP.get(doc_name)
    if not rel_path:
        fakture_dialogs.show_msgbox(ctx,
            "Nepoznat dokument: " + doc_name,
            "Fakture — Greška", "error", frame=frame)
        return

    file_path = os.path.join(base_path, rel_path)
    if not os.path.exists(file_path):
        fakture_dialogs.show_msgbox(ctx,
            "Fajl '{}' nije pronađen u folderu Sifrarnik/".format(
                os.path.basename(file_path)),
            "Fakture — Greška", "error", frame=frame)
        return

    _open_or_focus(ctx, file_path)


def _cmd_settings(ctx, frame):
    log.info("settings")
    fakture_dialogs.show_settings(ctx, frame)


# ──────────────────────────────────────────────────────────────────────────────
# UNO ProtocolHandler
# ──────────────────────────────────────────────────────────────────────────────

class FaktureDispatch(unohelper.Base, XDispatch):
    def __init__(self, context, frame, command):
        self.context = context
        self.frame = frame
        self.command = command

    def dispatch(self, url, args):
        try:
            ctx = self.context
            frame = self.frame
            cmd = self.command
            if cmd == "nova_faktura_domaci":
                _cmd_nova_faktura(ctx, frame, "domaci")
            elif cmd == "nova_faktura_ino":
                _cmd_nova_faktura(ctx, frame, "ino")
            elif cmd == "sync":
                _cmd_sync(ctx, frame)
            elif cmd == "open_domaci_kupci":
                _cmd_open(ctx, frame, "domaci_kupci")
            elif cmd == "open_ino_kupci":
                _cmd_open(ctx, frame, "ino_kupci")
            elif cmd == "open_proizvodi":
                _cmd_open(ctx, frame, "proizvodi")
            elif cmd == "settings":
                _cmd_settings(ctx, frame)
            else:
                log.warning("Unknown command: %s", cmd)
        except Exception as e:
            log.error("Dispatch [%s] failed: %s", self.command, e, exc_info=True)

    def addStatusListener(self, listener, url):
        pass

    def removeStatusListener(self, listener, url):
        pass


class FaktureProtocolHandler(unohelper.Base, XDispatchProvider, XServiceInfo, XInitialization):
    def __init__(self, context):
        self.context = context
        self.frame = None
        _apply_log_level(context)
        log.info("FaktureProtocolHandler initialized")

    def initialize(self, args):
        if args:
            self.frame = args[0]

    def queryDispatch(self, url, target_frame, search_flags):
        if url.Protocol == "vnd.fortunacommerc.fakture:":
            return FaktureDispatch(self.context, self.frame, url.Path)
        return None

    def queryDispatches(self, requests):
        return [self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags)
                for r in requests]

    def getImplementationName(self):
        return "com.fortunacommerc.fakture.ProtocolHandler"

    def supportsService(self, name):
        return name == "com.sun.star.frame.ProtocolHandler"

    def getSupportedServiceNames(self):
        return ("com.sun.star.frame.ProtocolHandler",)

log.info("Publishing handlers ...")
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    FaktureProtocolHandler,
    "com.fortunacommerc.fakture.ProtocolHandler",
    ("com.sun.star.frame.ProtocolHandler",))
log.info("Publishing handlers completed")
