# Fakture v2 — LibreOffice Extension

## Pregled

LibreOffice Calc extension za upravljanje šifarnicima (proizvodi, domaći i ino klijenti)
i kreiranje faktura iz predložaka sa auto-numeracijom. Dodaje "Fakture" meni u Calc menu bar.
Implementiran kao UNO ProtocolHandler (`vnd.fortunacommerc.fakture:*`).

## Fajlovi

| Fajl | Opis |
|------|------|
| `python/fakture.py` | UNO ProtocolHandler entry point: `FaktureProtocolHandler` + dispatch logika |
| `python/fakture_config.py` | Čitanje/pisanje LO konfiguracije (BasePath) |
| `python/fakture_sync.py` | Čita 3 .ods fajla → piše na skriveni `_Sifrarnik` sheet + Named Ranges |
| `python/fakture_faktura.py` | Kreiranje fakture: skeniranje RB, kopiranje template-a, sync |
| `python/fakture_dialogs.py` | UNO dijalozi: `_load_dialog` helper + msgbox, settings, identifier, folder picker |
| `dialogs/Settings.xdl` | Dialog za podešavanje baznog foldera |
| `dialogs/IdentifierDialog.xdl` | Dialog za unos identifikatora fakture |
| `dialogs/dialog.xlb` | Library manifest za dialogs/ |
| `Addons.xcu` | Meni definicija: 7 stavki sa `vnd.fortunacommerc.fakture:` URL-ovima |
| `ProtocolHandler.xcu` | Registracija `vnd.fortunacommerc.fakture:*` protokola sa LO |
| `Fakture.xcs` | LO konfiguraciona šema: Settings (BasePath) + Logging (LogLevel) grupe |
| `Fakture.xcu` | Default konfiguracione vrijednosti |
| `META-INF/manifest.xml` | Registracija fajlova; fakture.py registrovan kao uno-component |
| `description.xml` | Identitet: `com.fortunacommerc.fakture` v2.0.0 |
| `Makefile` | `make build` → `Fakture.oxt` / `make install` → `unopkg add` |

## Build & instalacija

```bash
make build      # kreira Fakture.oxt
make install    # build + unopkg remove/add (restart LO nakon toga)
make reinstall  # build + unopkg add --force (brži od install)
make uninstall  # unopkg remove
make clean      # briše .oxt
```

## Arhitektura — UNO ProtocolHandler

Ekstenzija koristi **UNO ProtocolHandler** (`vnd.fortunacommerc.fakture:*`).

### ProtocolHandler pattern

`fakture.py` je registrovan kao UNO komponenta u `manifest.xml` (type=Python).
LO ga učita pri startu, `FaktureProtocolHandler` prima frame u `initialize()`,
`FaktureDispatch` poziva odgovarajuću Python funkciju sa `(ctx, frame)`.

Klik na stavku menija → `vnd.fortunacommerc.fakture:<command>` → `queryDispatch()` →
`FaktureDispatch.dispatch()` → odgovarajuća `_cmd_*` funkcija.

### sys.path i modularni uvoz

`fakture.py` dodaje `os.path.dirname(__file__)` u `sys.path` na startu. Ostali moduli
(`fakture_config`, `fakture_dialogs`, `fakture_faktura`, `fakture_sync`) se uvoze
normalnim Python `import` mehanizmom — svi su pakovani u `.oxt` u `python/` direktoriju.

### Dispatch mapiranje

| Komanda | Funkcija |
|---------|----------|
| `nova_faktura_domaci` | `_cmd_nova_faktura(ctx, frame, "domaci")` |
| `nova_faktura_ino` | `_cmd_nova_faktura(ctx, frame, "ino")` |
| `sync` | `_cmd_sync(ctx, frame)` |
| `open_domaci_kupci` | `_cmd_open(ctx, frame, "domaci_kupci")` |
| `open_ino_kupci` | `_cmd_open(ctx, frame, "ino_kupci")` |
| `open_proizvodi` | `_cmd_open(ctx, frame, "proizvodi")` |
| `settings` | `_cmd_settings(ctx, frame)` |

## Logging

Log fajl: `~/.config/fakture/extension.log` (TimedRotatingFileHandler, sedmična rotacija `W0`, gzip kompresija, bez backupCount limita).
Logger se inicijalizuje **prije** UNO importa — tako se greške pri uvozu loguju.
Log level se čita iz LO registry: `/com.fortunacommerc.fakture/Logging/LogLevel` (default `INFO`).
Primjenjuje se u `FaktureProtocolHandler.__init__()` → `_apply_log_level(ctx)`.

Sub-logeri: `fakture.config`, `fakture.dialogs`, `fakture.faktura`, `fakture.sync`
(nasljeđuju handler od root `fakture` loggera).

## Konfiguracija

Globalna konfiguracija (ne per-document) kroz LO Configuration API:

| Čvor | Property | Opis |
|------|----------|------|
| `/com.fortunacommerc.fakture/Settings` | `BasePath` | Putanja do baznog foldera |
| `/com.fortunacommerc.fakture/Logging` | `LogLevel` | Log nivo: DEBUG/INFO/WARNING/ERROR |

## Struktura baznog foldera

```
<bazni_folder>/
├── Faktura-1-26__MojKupac.ods       # Fakture direktno u baznom folderu
├── Faktura-2-26__DrugaFirma.ods
├── Obrasci/
│   ├── faktura_domaci.ods           # Template za domaće fakture
│   └── faktura_ino.ods              # Template za ino fakture
└── Sifrarnik/
    ├── proizvodi.ods                # Šifarnik proizvoda i usluga
    ├── domaci_kupci.ods             # Šifarnik domaćih kupaca
    └── ino_kupci.ods                # Šifarnik ino kupaca
```

## Meni struktura

LO ne renderuje addon meni separatore — nisu uključeni.

```
Fakture
├── Nova domaća faktura              [uvijek]
├── Nova ino faktura                 [uvijek]
├── Osvježi šifrarnik                [samo sa otvorenim Calc dokumentom]
├── Domaći kupci                     [uvijek]
├── Ino kupci                        [uvijek]
├── Proizvodi i usluge               [uvijek]
└── Podešavanja sistema              [uvijek]
```

## Kreiranje fakture — tok

```
1. Provjeri config (BasePath) → ako nema, otvori settings
2. get_next_rb() → skenira bazni folder za Faktura-{RB}-{GG}__*.ods
3. show_identifier_dialog() → korisnik unosi identifikator
4. sanitize_identifier() → razmaci→_, samo dozvoljeni znakovi, max 50
5. shutil.copy2(template, dest) → Faktura-{RB}-{GG}__{ID}.ods
6. loadComponentFromURL() → otvori kopiju
7. sync_to_hidden_sheet() → automatski, bez pitanja
8. doc.store() → sačuvaj
```

## Detekcija godine

Iz naziva baznog foldera — zadnje 2 cifre. Fallback: tekuća godina.
- `Fakture-God26/` → `"26"`
- `MojeFakture/` → `datetime.now().strftime("%y")`

## Skriveni sheet `_Sifrarnik`

Sync čita 3 .ods fajla i piše 3 sekcije sa 2 prazna reda razmaka:

```
Row 0: [PROIZVODI]
Row 1: Naziv | ID | Bar kod | Jed. Mjere | Cijena BAM | Cijena EUR  (header)
Row 2+: podaci
       (2 prazna reda)
Row N: [DOMACI]
Row N+1: Firma | Podruznica | Ulica | PB Grad | JIB | PDV | BuyerID  (header)
Row N+2+: podaci
       (2 prazna reda)
Row M: [INO]
Row M+1: Firma | Ulica | PB Grad | Drzava | VAT | BuyerID  (header)
Row M+2+: podaci
```

### Named Ranges (6 komada)

| Naziv | Opis | Korištenje |
|-------|------|------------|
| `Proizvodi` | Cijela tabela proizvoda (bez headera), Naziv prva kolona | VLOOKUP po nazivu |
| `ProizvodiNazivi` | Kolona Naziv proizvoda (A) | Dropdown (Data Validation) |
| `DomaciKupci` | Cijela tabela domaćih kupaca | VLOOKUP izvor |
| `DomaciKupciNazivi` | Kolona Firma domaćih kupaca | Dropdown |
| `InoKupci` | Cijela tabela ino kupaca | VLOOKUP izvor |
| `InoKupciNazivi` | Kolona Firma ino kupaca | Dropdown |

## Izvorni fajlovi — struktura

**proizvodi.ods** (prvi sheet, header u redu 1):
```
ID | Bar kod | Naziv | Jed. Mjere | Cijena bez PDV-a | Valuta
```

**domaci_kupci.ods** (prvi sheet):
```
Firma | Podruznica | Ulica | PB | Grad | JIB | PDV
```

**ino_kupci.ods** (prvi sheet):
```
Firma | Ulica | PB | Grad | Drzava | VAT
```

## Auto-generisanje bar koda

Ako proizvod nema bar kod ali ima ID: `00` + nule + ID, min 8, max 14 cifara.

| ID | Bar kod | Dužina |
|----|---------|--------|
| `5` | `00000005` | 8 |
| `123` | `00000123` | 8 |
| `12345` | `00012345` | 8 |
| `123456` | `00123456` | 8 |
| `1234567` | `001234567` | 9 |
| `123456789012` | `00123456789012` | 14 |
| `1234567890123` | *(prazan)* | — |

## BuyerID logika

- Domaći: JIB → `"VP:" + JIB`, samo PDV → `"VP:4" + PDV`, ni jedno → `""`
- Ino: fiksno `"VP:9999999999999"`

## EUR/BAM konverzija

`EUR_TO_BAM = 1.955830`, zaokruženje na 4 decimale (ROUND_HALF_UP).

## Error handling

| Situacija | Reakcija |
|-----------|---------|
| Config ne postoji | Otvori dijalog Podešavanja |
| Bazni folder ne postoji | MsgBox greška + otvori Podešavanja |
| Obrazac ne postoji | MsgBox greška |
| Šifarnik .ods ne postoji | Preskače (prazna sekcija) |
| Prazan identifikator | MsgBox upozorenje |
| Fajl već otvoren | Fokusiraj postojeći prozor |

## LibreOffice API napomene

- `createMessageBox(parent, type, buttons, title, message)` — **5 parametara** (LO 26.x)
- `getCellByPosition(col, row)` — oba indeksa 0-bazirani
- `ConfigurationUpdateAccess` koristi `replaceByName()`, **ne** `setByName()`
- `frame` se uvijek proslijedi kao parametar — dolazi iz `FaktureDispatch.frame`
- `_load_dialog(ctx, frame, name)` — učitava XDL iz .oxt, kreira peer, centrira; vraća dialog
  - Extension URL: `vnd.sun.star.extension://com.fortunacommerc.fakture/dialogs/{name}.xdl`
- `fakture_config` koristi `uno.getComponentContext()` direktno (globalni LO kontekst)

## Okruženje

- LibreOffice 26.2+ (dozvoljena novija bugfix verzija), Python 3.12.12 (ugrađen u LibreOffice, ne sistemski), Linux (Ubuntu)
- Bez eksternih Python zavisnosti — samo stdlib + UNO bindings
