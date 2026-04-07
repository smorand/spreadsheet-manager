# Spreadsheet Manager

A comprehensive CLI and MCP server for managing Google Sheets spreadsheets with support for creating, formatting, styling, and data operations.

## Features

- **Create spreadsheets** : Create new spreadsheets or copy from templates
- **Data management** : Add data, import/export CSV files
- **Cell formatting** : Format cells as NUMBER, CURRENCY, DATE, PERCENT, TIME, or TEXT
- **Cell styling** : Apply colors, fonts, bold, italic, and font sizes
- **Sheet operations** : Create, rename, and list sheets
- **Layout control** : Freeze rows/columns, set column width, text wrap, alignment
- **Banding** : Alternating row colors
- **Notes** : Add notes to individual cells
- **MCP Server** : HTTP Streamable MCP server with OAuth 2.1 for AI agent integration

## Installation

### Prerequisites

- Go 1.25 or later
- Google Cloud project with Sheets API enabled
- OAuth2 credentials

### Build from source

```bash
make build
```

### Install to system

```bash
# Install to /usr/local/bin
make install

# Or install to custom location
TARGET=/path/to/bin make install
```

## Setup

### 1. Google Cloud Credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing one
3. Enable the Google Sheets API and Google Drive API
4. Create OAuth 2.0 credentials (Desktop application)
5. Download the credentials JSON file

### 2. Configure credentials

```bash
mkdir -p ~/.credentials
cp /path/to/downloaded/credentials.json ~/.credentials/google_credentials.json
```

### 3. First run authentication

On first run, the tool will prompt you to authenticate via browser and save the token to `~/.credentials/google_token.json`.

## Usage

### Create a new spreadsheet

```bash
spreadsheet-manager create "My Spreadsheet"
```

### Create from template

```bash
spreadsheet-manager create "New Document" --template TEMPLATE_ID --folder FOLDER_ID
```

### Add data to cells

```bash
spreadsheet-manager add-data SPREADSHEET_ID "Sheet1" "A1:B2" '[["Name","Age"],["John",30]]'
```

With formulas disabled (raw values):

```bash
spreadsheet-manager add-data SPREADSHEET_ID "Sheet1" "A1" '[["=SUM(1,2)"]]' --formula=false
```

### Import CSV data

```bash
spreadsheet-manager import-csv SPREADSHEET_ID "Sheet1" data.csv --start A1
```

### Format cells

```bash
# Format as currency
spreadsheet-manager format-cells SPREADSHEET_ID "Sheet1" "B2:B10" CURRENCY

# Custom pattern
spreadsheet-manager format-cells SPREADSHEET_ID "Sheet1" "A1:A10" DATE --pattern "dd/mm/yyyy"
```

Supported format types:
- `NUMBER` - Numeric values with decimal places
- `CURRENCY` - Currency formatting with symbols
- `DATE` - Date formatting
- `PERCENT` - Percentage values
- `TIME` - Time formatting
- `TEXT` - Text format

### Style cells

```bash
# Apply background and font colors
spreadsheet-manager style-cells SPREADSHEET_ID "Sheet1" "A1:B1" \
  --bg-color "#4285f4" \
  --font-color "#ffffff" \
  --bold \
  --font-size 12

# Make text italic
spreadsheet-manager style-cells SPREADSHEET_ID "Sheet1" "A2:A10" --italic
```

### Export to CSV

```bash
spreadsheet-manager export-csv SPREADSHEET_ID "Sheet1" output.csv
```

### Sheet operations

```bash
# Create a new sheet
spreadsheet-manager create-sheet SPREADSHEET_ID "New Sheet"

# Rename a sheet
spreadsheet-manager rename-sheet SPREADSHEET_ID "Old Name" "New Name"

# List all sheets
spreadsheet-manager list-sheets SPREADSHEET_ID
```

### Add notes to cells

```bash
spreadsheet-manager add-note SPREADSHEET_ID "Sheet1" "A1" "This is a note"
```

## MCP Server

The project includes an HTTP Streamable MCP server that exposes all spreadsheet operations as tools for AI agents (e.g., Claude).

### Starting the server

```bash
# Local development (with local credential file)
spreadsheet-manager mcp \
  --credential-file ~/.credentials/scm-pwd-web.json \
  --base-url http://localhost:8080

# With all options
spreadsheet-manager mcp \
  --port 8080 \
  --host 0.0.0.0 \
  --base-url https://spreadsheet-manager.scm-platform.org \
  --secret-project my-gcp-project \
  --secret-name scm-pwd-web
```

### Docker deployment

```bash
docker build -t spreadsheet-manager .
docker run -p 8080:8080 \
  -e BASE_URL=https://spreadsheet-manager.scm-platform.org \
  -e CREDENTIAL_FILE=/data/scm-pwd-web.json \
  -v /path/to/credentials:/data \
  spreadsheet-manager
```

### Available MCP tools

| Tool | Description |
|------|-------------|
| `spreadsheet_create` | Create a new spreadsheet |
| `spreadsheet_add_data` | Add or update cell data |
| `spreadsheet_import_csv` | Import CSV file |
| `spreadsheet_export_csv` | Export to CSV file |
| `spreadsheet_format_cells` | Apply number formatting |
| `spreadsheet_style_cells` | Apply visual styling |
| `spreadsheet_create_sheet` | Create a new sheet tab |
| `spreadsheet_rename_sheet` | Rename a sheet |
| `spreadsheet_list_sheets` | List all sheets |
| `spreadsheet_add_note` | Add a note to a cell |
| `spreadsheet_freeze` | Freeze rows/columns |
| `spreadsheet_set_column_width` | Set column width |
| `spreadsheet_set_text_wrap` | Set text wrapping |
| `spreadsheet_set_alignment` | Set cell alignment |
| `spreadsheet_alternate_row_colors` | Add banded row colors |

## Output Format

All commands return JSON output for easy parsing:

```json
{
  "id": "1abc123...",
  "url": "https://docs.google.com/spreadsheets/d/1abc123.../edit",
  "status": "success"
}
```

## Development

### Build

```bash
make build
```

### Run tests

```bash
make test
```

### Format code

```bash
make fmt
```

### Run linter

```bash
make vet
```

### Run all checks

```bash
make check
```

### Clean build artifacts

```bash
make clean

# Clean including go.mod and go.sum
make clean-all
```

### Rebuild from scratch

```bash
make rebuild
```

## Project Structure

```
spreadsheet-manager/
├── Makefile              # Build automation
├── README.md             # This file
├── CLAUDE.md             # AI-oriented documentation
├── spreadsheet-manager   # Compiled binary
└── src/
    ├── main.go           # Entry point and command registration
    ├── cli.go            # Command definitions and implementations
    ├── auth.go           # Google OAuth2 authentication
    ├── helpers.go        # Helper functions
    ├── go.mod            # Go module definition
    └── go.sum            # Go module checksums
```

## Troubleshooting

### Authentication issues

If you encounter authentication problems:

1. Delete the token file: `rm ~/.credentials/google_token.json`
2. Run any command again to re-authenticate

### Permission errors

Ensure your OAuth2 credentials have the following scopes:
- `https://www.googleapis.com/auth/spreadsheets`
- `https://www.googleapis.com/auth/drive.file`

### API quota exceeded

Google Sheets API has usage limits. If exceeded, wait or request quota increase in Google Cloud Console.

## License

This project is provided as-is for personal and commercial use.

## Contributing

Contributions are welcome! Please ensure code follows Go coding standards:
- Run `make fmt` before committing
- Run `make vet` to check for issues
- Add tests for new functionality
- Update documentation as needed
