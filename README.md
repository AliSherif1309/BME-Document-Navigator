# BME Document Navigator Pro+

**Project Overview:**

This application is a desktop tool designed specifically for Biomedical Equipment Technicians (BMEs) and Clinical Engineers to effectively manage, search, and access technical documentation crucial for equipment service and maintenance. It provides features for indexing large collections of manuals, Standard Operating Procedures (SOPs), datasheets, and other relevant files, allowing for quick retrieval based on metadata and powerful full-text content searching.

The core concept was inspired by the functionality of scholarly text databases like Al-Maktaba Al-Shamela but has been adapted and tailored to the specific needs of the technical BME workflow.

**Development:**

*   **Developer:** Aly Sherif
*   **AI Assistance:** Significant assistance provided by Google's Gemini Pro language model for code generation, debugging, feature brainstorming, implementation strategies, and refinement.

## Core Features

*   **Indexed Document Library:** Scans user-configured folders for various document types (`.pdf`, `.docx`, `.txt`, `.html`, common images, archives like `.zip`, `.7z`, etc.).
*   **File Tree Browser:** Navigate your document library through a familiar hierarchical folder structure with icons.
*   **Metadata Management:**
    *   Automatic (basic) extraction of Manufacturer/Model/Type.
    *   Prompts for default Manufacturer during folder scanning.
    *   Manual editing of individual document metadata (Manufacturer, Model, Type, Revision, Rev Date, Status, Applicable Models, Test Equipment, Keywords).
    *   Batch editing capability to apply metadata changes to multiple selected files.
*   **Full-Text Search (FTS):** Searches within the content of indexed PDFs, DOCX, TXT, and HTML files using SQLite FTS5.
*   **Integrated Search Results Tab:** Displays combined metadata and FTS results in a dedicated notebook tab, showing context snippets and the page number for the best match per document.
*   **Tabbed Document Viewer:** Open multiple documents concurrently in separate tabs.
    *   Internal PDF viewer with page rendering.
    *   Internal text viewer for DOCX, TXT, HTML (tags stripped).
    *   Page navigation controls (Next, Previous).
    *   Browser-style Back/Forward navigation history within each tab.
    *   Zoom functionality for PDF viewing.
*   **Outline Navigation (PDF):** Extracts and displays PDF Bookmarks/Table of Contents in a dedicated details panel tab, allowing direct navigation to sections. Includes expand/collapse all and filtering capabilities.
*   **Document Linking:**
    *   Manually create links between related documents with optional descriptions.
    *   Semi-automatic link suggestion feature that scans document text for potential references (filenames, codes, PNs) and suggests links to create.
*   **Notes/Annotations:**
    *   Add timestamped text notes associated with specific documents and pages.
    *   View all notes for a selected document or view all notes across the library.
    *   Edit and Delete existing notes.
    *   Double-click a note to open the corresponding document and page.
*   **Favorites/Bookmarks:** Bookmark specific document pages for quick access via a dedicated menu. Includes management (Rename/Delete) functionality.
*   **Session Persistence:** Remembers window size/position, side pane layout (sash positions), and restores previously open document tabs (including page number and zoom level) on startup via an `.ini` configuration file.
*   **Collapsible Panes:** Side panels (File Tree, Details) can be collapsed via the View menu or dedicated buttons to maximize the document viewing area.
*   **Customizable Appearance:** Supports switching between available system Tkinter/ttk themes via the View menu.

## Requirements

*   Python 3.8+ (Developed primarily using 3.10)
*   Required Python packages (install via pip):
    *   `Pillow`
    *   `PyMuPDF`
    *   `python-docx`

## Setup and Usage (from Source Code)

1.  **Clone/Download:** Get the source code repository.
2.  **Navigate:** Open a terminal/command prompt in the project directory.
3.  **Environment:** Create and activate a Python virtual environment (recommended):
    ```bash
    python -m venv .venv
    # Activate (Windows CMD): .\.venv\Scripts\activate
    # Activate (Windows PowerShell): .\.venv\Scripts\Activate.ps1
    # Activate (macOS/Linux): source .venv/bin/activate
    ```
4.  **Install Dependencies:** Install the required packages:
    ```bash
    pip install --upgrade pip
    pip install Pillow PyMuPDF python-docx
    # Or if requirements.txt exists: pip install -r requirements.txt
    ```
5.  **Run Script:** Execute the main Python file:
    ```bash
    python bme_navigator.py # Or your script's name
    ```
6.  **Configure Paths:** On first run, navigate to `File -> Manage Scan Paths...` and add the root directories containing your technical documents.
7.  **Initial Scan:** Navigate to `File -> Scan/Update Index`. This builds the database and search index and will take time initially. Assign default manufacturers per folder when prompted if desired.
8.  **Use:** Browse the file tree, use the search bar (press Enter or click "Search"), open documents in tabs, view details/notes/outline/links, add favorites, edit metadata.

## Creating `requirements.txt`

In your activated virtual environment after installing packages:
```bash
pip freeze > requirements.txt

## Building the Executable (using PyInstaller)

1.  Install PyInstaller in your virtual environment:
    ```bash
    pip install pyinstaller
    ```
2.  (Optional) Ensure your icon file (e.g., `icons/favicon.ico`) exists in an `icons` subfolder relative to your script.
3.  Navigate to the project directory in your terminal (with the virtual environment active).
4.  Run PyInstaller once to generate the `.spec` file (adjust script name and icon path if needed):
    ```bash
    pyinstaller --noconsole --icon=icons/favicon.ico bme_navigator.py --name AlyNavigationFiles
    ```
5.  Edit the generated `AlyNavigationFiles.spec` file:
    *   Ensure `datas=[('icons', 'icons')]` is present under `Analysis` (to include the icons folder).
    *   Ensure `windowed=True` is set under `EXE`.
    *   Ensure `icon='icons/favicon.ico'` (as a string) is set under `EXE`.
6.  Delete the generated `build/` and `dist/` folders.
7.  Run the build using the spec file:
    ```bash
    pyinstaller AlyNavigationFiles.spec
    ```
8.  The distributable application folder will be `dist/AlyNavigationFiles`.

## Deployment Note

The application relies on accessing the original document files at the paths stored in its database (`bme_doc_index.db`). For the viewer and content search to function correctly, the computer running the application must have access to these documents at the **exact same paths** (e.g., via mapped network drives or identical local directory structures).

## Future Enhancements

*   FTS result highlighting within the viewer.
*   Advanced search/metadata filtering UI.
*   DOCX outline support.
*   Improved Note/Outline interaction.
*   Document version comparison.
*   Further UI/UX refinements (more icons, progress indicators, keyboard navigation).
