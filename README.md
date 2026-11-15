# BME Document Navigator

A private, local document workspace for biomedical equipment technicians and
clinical engineers. Bring your own manuals, SOPs, calibration references, and
datasheets; find them by equipment metadata or text inside the document.

The new HTML interface runs on the same computer as your library. It uses no
cloud service, account, external fonts, or analytics, and needs no internet after
installation. Windows, macOS, and Linux share the same Python application.

## Start here

Install Python **3.10 or newer**, then open a terminal in this repository:

```sh
python -m venv .venv
```

Activate it on Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Or on macOS/Linux:

```sh
source .venv/bin/activate
```

Install the document readers and launch:

```sh
python -m pip install -r requirements.txt
python navigator_web.py
```

Your default browser opens the app. Keep the terminal open; press **Ctrl+C** to
stop it. If your system uses `python3`, use that command to create the environment.
The browser interface uses Python's standard library; PDF and Word indexing need
the installed readers. No Node.js installation is needed to run the app.

1. Open **Library folders** and paste a full folder path from your file manager.
2. Adding a folder starts a scan, including its subfolders.
3. Search for a model, fault code, manufacturer, or words from the text.
4. Open a result, then choose **Read PDF** to use the in-app reader. You can also
   edit metadata, add page-specific notes, or save a favorite.
5. Use **Scan details** to see extraction errors and unavailable folders.

## What works now

| Capability | Behavior |
| --- | --- |
| Local HTML workspace | Responsive layout, automatic light/dark appearance, totals, pagination, and `/` to focus search |
| Search | Words across filenames, BME metadata, and document text; matching page and context for text hits |
| Filters | Manufacturer, document type, and favorites |
| Saved searches | Named searches stored with their words, manufacturer/type filters, and favorites view |
| Metadata | Manufacturer, model, revision/date, status, applicable models, test equipment, keywords |
| Batch editing | Select a results page, choose fields, and review current/new values before applying |
| Folder reconnection | Preview and verify identical files at a new folder path while keeping document IDs and history |
| Notes | Timestamped notes with page references, saved independently of source files |
| Favorites and page bookmarks | Shared with the original desktop application; removing a document star retains its page bookmarks |
| Automatic indexing | Incremental scans on startup and every five minutes while running; manual scan available |
| PDF reader | Page navigation, outline, zoom, phrase highlighting, extracted text, page bookmarks, and saved reading position |
| Related documents | Link manuals, SOPs, and references with descriptions; links appear from either document |
| Other files | Downloads the original for use in its native application |
| CSV export | Entire library metadata with spreadsheet formula protection |
| Database backup | Consistent SQLite snapshot including index, notes, favorites, and existing document links |
| Existing data | Uses the original database and retains the original desktop viewer |

Automatic metadata is a filename/folder heuristic; verify it and fill in the model
manually. Search treats punctuation as separators and combines words with AND;
it does not accept raw SQLite FTS operators. A result matches either its metadata
or its indexed document text.

### File support

**Content indexing:** PDF text, Word `.docx` paragraphs and tables, `.txt`, `.md`,
`.csv`, and `.html`/`.htm` visible text. PDF matches retain page numbers. Word and
plain-text content are treated as one text block.

**Metadata listing:** legacy Office files, spreadsheets, presentations, images,
RTF, and ZIP/7z/RAR archives. Archives are not unpacked; executables are not indexed
or launched. HTML downloads as a file rather than executing in the workspace.
Scanned/image-only PDFs need OCR, which is not included yet.

### Reading and connecting references

Choose **Read PDF** in a document's details to open it inside the workspace. Use
Previous/Next, the page field, or a Contents entry to navigate. Left/right arrow
keys also navigate when focus is outside a text field or select control. The app
remembers the last successfully opened page and zoom for each document. Search
hits and note-page links open their specific page instead of the saved position.

**Find phrase on this page** highlights up to 200 literal phrase matches on the
current page. It is separate from the library's word-based full-text search.
**Page text** provides selectable extracted text; the rendered page itself is an
image. Scanned pages may have no extractable text until OCR is available. Page
images are bounded to roughly 12 million pixels and a 6000-pixel maximum dimension.
Highlighting and rendering never save changes to the source PDF.

Use **Bookmark page** to save a named reference, then load it from the reader or
**Page bookmarks** in the sidebar. Names are unique across the shared favorites
table. Existing desktop page bookmarks remain visible. The document star controls
the app's document-level favorite; removing it leaves page bookmarks intact.
The Favorites view and saved-reference total include documents saved by either
method. Password-protected PDFs can still be opened in your own PDF application.

In document details, **Related documents** shows existing links and lets you find
another document by its filename, metadata, or indexed text. Choose a match, add
an optional description, and save the link. Either document can then open the
other. Removing a link changes only the relationship, not either source file,
library record, note, or bookmark.

## Data preservation

### Batch editing and saved searches

Use **Save search** to name the current search and filters. Saved searches persist
in your library database and can be loaded from the sidebar or **Saved searches**.
Search names are unique without regard to case. A saved search stores criteria,
so its results update as your library changes.

Select individual documents or **Select this page**, then choose **Edit selected
metadata**. Only checked fields are edited; a checked field with a blank value
explicitly clears that field. Review the current/new values before applying.
Selection is limited to the current results page and clears when results refresh.

Batch edits and folder reconnections create a database snapshot in the `backups`
folder beside the library database. All changes apply in one transaction. If an
affected value or file changes after the preview, the operation stops without
partially applying changes. Previews are single-use, expire after ten minutes,
and cannot be applied after restarting the app. Backups are retained until you
remove them; periodically archive or clean them up to manage disk usage.

### Reconnecting a moved folder

In **Library folders**, choose **Reconnect** beside the previous folder, or
**Reconnect a moved folder** to enter a previously removed folder path. Enter the
new folder path and preview the matches. Keep filenames and the relative folder
structure unchanged. Each file must match the indexed file's content fingerprint.
Resolve missing files, content differences, and already-indexed destinations
before applying; this workflow does not merge duplicate library records.

The app updates document paths and registered scan folders while retaining IDs,
notes, favorites, links, metadata, and searchable text. It never moves files on
disk. Reconnect before adding/scanning the destination as a separate library folder.
If an enclosing scan folder still includes an old copy, that copy can be indexed
again on a later scan; remove the old copy from your scan scope when moving.

The first scan after this update refreshes legacy indexed documents and records
their fingerprints. If a legacy original is already offline and has no recorded
fingerprint, reconnect the original and scan it first. The app cannot safely
verify such a document using its filename alone. Individual file renames and
changed folder layouts still require future relinking support.

### Offline files and backups

Scans never modify source documents. An unavailable drive, removed scan folder,
deleted document, or partial traversal does **not** delete records, notes, or
favorites. Failed text extraction retains the previous index and is retried on
the next scan. Edited metadata survives future scans.

Missing documents keep their last indexed text, so search may show an offline
reference. Its details report that the original is unavailable. Scanning a renamed
or moved file without first using folder reconnection creates a new record.

The database and original desktop settings live in:

| System | Application data folder |
| --- | --- |
| Windows | `%LOCALAPPDATA%\BMEDocumentNavigator` |
| macOS | `~/Library/Application Support/BMEDocumentNavigator` |
| Linux | `$XDG_DATA_HOME/BMEDocumentNavigator`, falling back to `~/.local/share/BMEDocumentNavigator` |

Existing database/settings beside the original script are copied on first use;
originals remain intact. Set `BMEDOCUMENTNAVIGATOR_DATA_DIR` to choose another data
folder, or pass `--database` to choose a specific database.

Use **Back up library** before major changes. Backups contain library data, not
source documents or desktop window settings. To restore, stop both apps, keep a
copy of the current database, and put the backup at the database path as
`bme_doc_index.db`. Sources must remain accessible at their recorded paths.
Keep backups private: indexed text and notes may contain sensitive material.

## Launch options

```sh
python navigator_web.py --port 8765
python navigator_web.py --scan-interval 60
python navigator_web.py --scan-interval 0 --no-browser
python navigator_web.py --database /full/path/to/library.db
```

The default chooses a free port. Automatic rescans run only while the app runs;
this is not an installed background service. The server binds to `127.0.0.1`,
validates the host, and requires a session token for library access. There is no
LAN sharing or multi-user authentication. Do not expose it through a public proxy.
SQLite is local storage, not encryption at rest.

## Portable applications

Build **on the operating system you intend to distribute to**:

```sh
python -m pip install pyinstaller
python build_app.py
```

Existing bundles are protected from replacement by default. Close an existing
portable app and use `python build_app.py --replace` to rebuild into its folder.
To keep it running, build separately with `python build_app.py --output dist-updated`
and verify that bundle with `python smoke_app.py --output dist-updated`.

Distribute the **entire** `dist/BME-Navigator` folder. Windows launches
`BME-Navigator.exe`; macOS/Linux launch `BME-Navigator` from a terminal. It opens
the browser and keeps a console for stopping the server. These are portable
bundles, not signed installers or a macOS `.app` package.

The **Build portable applications** GitHub Actions workflow builds, smoke-tests,
and uploads separate Windows, macOS, and Linux artifacts for application pull
requests, version tags, and manual runs. Regression checks run on all
three platforms for pushes and pull requests. Artifacts inherit the runner's CPU
architecture; they do not cover every architecture. Regression tests, builds, and
packaged-app checks passed on Windows, macOS, and Linux GitHub runners. Windows
was also built and exercised locally. Signing and installers remain planned.

GitHub artifact downloads do not preserve executable permissions. After extracting
a macOS/Linux artifact, run `chmod +x BME-Navigator` inside its application folder
before launching `./BME-Navigator`. Extract the whole artifact, including `_internal`.

## Original desktop viewer

```sh
python bme_navigator.py
```

The original Tkinter viewer remains available with tabbed PDFs, zoom, outline,
document links, batch metadata editing, notes, bookmarks, and session restoration.
Its scan commands now use the repaired shared engine. Linux may need its
Python Tk package. The HTML workspace now includes outline navigation, page
bookmarks, document links, and batch editing; it currently uses a single reader
instead of the desktop viewer's multiple document tabs.

## Development and verification

```sh
python -m unittest discover -v
node --check web/app.js
node --check web/reader.js
```

Node is used only for the optional JavaScript syntax check. Tests cover storage
migration, PDF pages, DOCX tables, metadata preservation, offline drives,
extraction failures, overlapping folders, filters, browser access boundaries,
safe downloads, CSV export, backups, saved searches, batch rollback, folder
identity, stale previews, and reconnection history. Packaged-app checks also
exercise saved searches, batch edits, reconnection, PDF page rendering, outlines,
highlighting, reading positions, bookmarks, and related documents. Tests use
temporary libraries and verify that rendering leaves source PDFs unchanged.

`navigator_core.py` owns library services; `navigator_web.py` owns the local HTTP
app; `document_reader.py` handles PDF reading; `pdf_support.py` keeps new indexing
and reader PDF operations on a single worker, in line with
[PyMuPDF's threading guidance](https://pymupdf.readthedocs.io/en/latest/recipes-multiprocessing.html).
`web/` contains HTML/CSS/JavaScript; `app_storage.py` handles storage and migration.
See [ROADMAP.md](ROADMAP.md) for future stages and current limits.

## Credits

Created by **Aly Sherif**. The original application was developed with assistance
from Google's Gemini Pro. This modernization was developed with OpenAI Codex.
The original concept was inspired by Al-Maktaba Al-Shamela and adapted to
biomedical engineering documentation workflows.
