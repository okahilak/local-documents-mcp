# Local Documents

Claude Desktop extension for reading PDF, Word and Excel files, and editing Excel workbooks.

Works on macOS (Apple Silicon) and Windows (x64). Nothing else to install.

## Install

Build `LocalDocuments.mcpb` (see below), open it in Claude Desktop, and choose the folders Claude may access.

## Tools

- `pdf_info`, `pdf_read_text`, `pdf_render_page`
- `docx_info`, `docx_read_text`, `docx_read_tables`, `docx_search`
- `xlsx_info`, `xlsx_read_range`, `xlsx_search`
- `xlsx_edit`: set values and formulas, clear cells, add sheets

Edits are saved to a new file unless you ask to overwrite the original. Everything else in the workbook, including charts, pivot tables and macros, is left as is.

Files outside the chosen folders can't be accessed.

## Development

```sh
npm install
npm test
npm run build
```

## License

MIT
