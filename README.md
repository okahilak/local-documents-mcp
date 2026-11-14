# Local Documents

**Work with local PDF, Word, and Excel files directly from Claude Desktop.**

Local Documents is a Claude Desktop extension that lets Claude read documents, search Word and Excel files, edit Excel workbooks, and organise files and folders on your computer.

Supports **macOS (Apple Silicon)** and **Windows (x64)**. No Python or separate document-processing software required.

## What you can do

- **Read PDFs:** Extract text, inspect document details, and render individual pages as images (useful for scanned pages, charts and diagrams).
- **Work with Word documents:** Read text and tables, and search document contents.
- **Read and edit Excel workbooks:** Inspect sheets, read cell ranges, search values, update cells and formulas, and add sheets.
- **Organise files:** List directories, create folders, and move or rename files and folders individually or in batches.

### Example requests

Ask Claude to:

- "Which of the Word documents in this folder mention the Riverside project?"
- "Find every row in this workbook that mentions Helsinki."
- "Read the tables from these Word documents."
- "Update the prices in this Excel workbook and save a new copy."
- "Create folders for each year and move these documents into the appropriate folders."
- "Show me page 5 of this PDF."

## Installation

Local Documents works with Claude Desktop on macOS (Apple Silicon) and Windows (x64).

1. Download `LocalDocuments.mcpb` from the [latest release](https://github.com/okahilak/local-documents-mcp/releases/latest).
2. In Claude Desktop, open **Settings → Desktop app → Extensions**.
3. Drag the downloaded `.mcpb` file into the Extensions window.
4. Review the extension details and confirm the installation.
5. Under **Allowed folders**, choose which folders Claude may access. The default is your `Documents` folder; use **Add directory** to add more.
6. Click **Save** to finish setup.

Local Documents is now ready to use. Start a new conversation and ask Claude to work with files in your selected folders. See [Example requests](#example-requests) for ideas.

**Folder access:** Local Documents can only access files inside the folders you select, including their subfolders. You can change these folders later in the extension's settings.

## Safety and file handling

**You control which folders Claude can access.** Files outside the selected folders are inaccessible to the extension.

- Excel edits are saved to a new file next to the original, named `<name> (edited).xlsx`, unless you ask for a different name. Existing files are overwritten only when explicitly requested.
- Other workbook content, including charts, pivot tables, and macros, is preserved.
- File organisation operations never overwrite or delete existing files.
- Batch moves are transactional: if an operation fails, earlier changes are rolled back.

## Supported files and limitations

- **PDF** (`.pdf`): text is read from the existing text layer; there is no OCR. Scanned pages can be viewed with `pdf_render_page`.
- **Word** (`.docx`): the older `.doc` format is not supported.
- **Excel** (`.xlsx`, `.xlsm`): the older `.xls` format is not supported. Formulas are stored, not calculated; Excel recalculates them when the file is opened.
- Searching works within one Word or Excel file at a time; there is no search across files or within PDFs.

## Available tools

| Category | Tools |
|---|---|
| PDF | `pdf_info`, `pdf_read_text`, `pdf_render_page` |
| Word | `docx_info`, `docx_read_text`, `docx_read_tables`, `docx_search` |
| Excel | `xlsx_info`, `xlsx_read_range`, `xlsx_search`, `xlsx_edit` |
| Files and folders | `list_directory`, `create_folder`, `rename`, `move`, `move_batch` |

### Excel editing

`xlsx_edit` supports:

- Setting cell values and formulas
- Clearing cells
- Adding worksheets
- Renaming Excel table columns (setting a table's header cell also updates the table's column name)

All edits in one call are applied together: if any edit is invalid, nothing is written.

## Build from source

If you prefer to build the extension yourself, you need Node.js 22.13 or later. The build downloads platform-specific binaries, so it needs network access, and produces a single bundle that works on both macOS and Windows.

```bash
npm install
npm test
npm run build
```

The build produces `LocalDocuments.mcpb`, which you can install using the same steps as above, starting from step 2.

## License

MIT
