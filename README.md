# Spreadsheet Manager

A comprehensive command-line tool for managing Google Sheets spreadsheets with support for creating, formatting, styling, and data operations.

## Features

- **Create spreadsheets** - Create new spreadsheets or copy from templates
- **Data management** - Add data, import/export CSV files
- **Cell formatting** - Format cells as NUMBER, CURRENCY, DATE, PERCENT, TIME, or TEXT
- **Cell styling** - Apply colors, fonts, bold, italic, and font sizes
- **Sheet operations** - Create, rename, and list sheets
- **Notes** - Add notes to individual cells

## Installation

### Prerequisites

- Go 1.21 or later
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
