# Local Documents

Claude Desktop extension for reading PDF, Word and Excel files, editing Excel workbooks, and organising files and folders.

Works on macOS (Apple Silicon) and Windows (x64). Nothing else to install.

## Install

Build `LocalDocuments.mcpb` (see below), open it in Claude Desktop, and choose the folders Claude may access.

## Tools

- `pdf_info`, `pdf_read_text`, `pdf_render_page`
- `docx_info`, `docx_read_text`, `docx_read_tables`, `docx_search`
- `xlsx_info`, `xlsx_read_range`, `xlsx_search`
- `xlsx_edit`: set values and formulas, clear cells, add sheets, rename table column headers

Edits are saved to a new file unless you ask to overwrite the original. Everything else in the workbook, including charts, pivot tables and macros, is left as is.

- `list_directory`: see what's in a folder, optionally with all its subfolders
- `create_folder`, `rename`, `move`: organise files and folders of any type
- `move_batch`: many moves and renames in one call, e.g. to reorganise a folder

Organising never overwrites or deletes anything. If one move in a `move_batch` fails, the earlier ones are undone and nothing changes.

Files outside the chosen folders can't be accessed.

## Development

```sh
npm install
npm test
npm run build
```

## License

MIT
