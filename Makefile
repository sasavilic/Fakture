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

EXT = Fakture
UNOPKG = unopkg

SOURCES = \
	META-INF/manifest.xml \
	description.xml \
	LICENSE \
	Addons.xcu \
	ProtocolHandler.xcu \
	Fakture.xcs \
	Fakture.xcu \
	python/fakture.py \
	python/fakture_config.py \
	python/fakture_sync.py \
	python/fakture_faktura.py \
	python/fakture_dialogs.py \
	dialogs/dialog.xlb \
	dialogs/Settings.xdl \
	dialogs/IdentifierDialog.xdl \
	dialogs/TemplatePickerDialog.xdl

.PHONY: build install reinstall uninstall clean

build: $(EXT).oxt

$(EXT).oxt: $(SOURCES)
	rm -f $(EXT).oxt
	zip -r $(EXT).oxt $(SOURCES) dialogs/

install: build
	$(UNOPKG) remove com.fortunacommerc.fakture 2>/dev/null || true
	$(UNOPKG) add --suppress-license $(EXT).oxt
	@echo "Restart LibreOffice to apply changes."

reinstall: build
	$(UNOPKG) add --force --suppress-license $(EXT).oxt
	@echo "Restart LibreOffice to apply changes."

uninstall:
	$(UNOPKG) remove com.fortunacommerc.fakture 2>/dev/null || true

clean:
	rm -f $(EXT).oxt
