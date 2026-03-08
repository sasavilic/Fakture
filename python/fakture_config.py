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
fakture_config — Read/write LibreOffice configuration for the Fakture extension.

Config node: /com.fortunacommerc.fakture/Settings
Properties:  BasePath (string)
"""

import os
import re
import datetime

import logging

import uno

CONFIG_NODE = "/com.fortunacommerc.fakture/Settings"

_log = logging.getLogger("fakture.config")


def get_base_path():
    """Read base folder from LO configuration. Returns empty string if not set."""
    try:
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        config_provider = smgr.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx
        )
        node = config_provider.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationAccess",
            (uno.createUnoStruct("com.sun.star.beans.NamedValue", "nodepath", CONFIG_NODE),)
        )
        result = node.getByName("BasePath")
        return result
    except Exception as e:
        _log.error("get_base_path failed: %s", e, exc_info=True)
        return ""


def set_base_path(path):
    """Save base folder to LO configuration."""
    _log.info("set_base_path: %r", path)
    try:
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        config_provider = smgr.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx
        )
        node = config_provider.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationUpdateAccess",
            (uno.createUnoStruct("com.sun.star.beans.NamedValue", "nodepath", CONFIG_NODE),)
        )
        node.replaceByName("BasePath", path)
        node.commitChanges()
    except Exception as e:
        _log.error("set_base_path failed: %s", e, exc_info=True)
        raise


def detect_year_from_folder(base_path):
    """Detect year from base folder name (last 2 digits). Fallback: current year."""
    folder_name = os.path.basename(base_path.rstrip(os.sep))
    match = re.search(r'(\d{2})$', folder_name)
    if match:
        return match.group(1)
    return datetime.datetime.now().strftime("%y")


def ensure_folder_structure(base_path):
    """Create Obrasci/ and Sifrarnik/ subdirs if they don't exist. Returns list of created dirs."""
    created = []
    for sub in ("Obrasci", "Sifrarnik"):
        path = os.path.join(base_path, sub)
        if not os.path.exists(path):
            os.makedirs(path)
            created.append(sub)
    return created
