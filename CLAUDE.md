# Spreadsheet Manager - AI Documentation

## Project Overview

**Type**: Command-line tool
**Language**: Go 1.21
**Purpose**: Comprehensive Google Sheets management via CLI
**Primary Dependencies**:
- `github.com/spf13/cobra` - CLI framework
- `google.golang.org/api/sheets/v4` - Google Sheets API
- `google.golang.org/api/drive/v3` - Google Drive API
- `golang.org/x/oauth2` - OAuth2 authentication

## Architecture

### Code Organization

```
src/
├── main.go      - Entry point, command registration
├── cli.go       - All command definitions and implementations
├── auth.go      - OAuth2 authentication logic
├── helpers.go   - Utility functions (A1 notation, colors, etc.)
├── go.mod       - Module definition
└── go.sum       - Dependency checksums
```

### Authentication Flow

1. Check for credentials at `~/.credentials/google_credentials.json`
2. Load existing token from `~/.credentials/google_token.json` or initiate OAuth flow
3. Token is cached and reused for subsequent requests
4. Context is properly passed through all authentication functions

### Command Structure

All commands follow a consistent pattern using immediately-invoked function expressions (IIFE) to avoid `init()` functions:

```go
var commandCmd = func() *cobra.Command {
    cmd := &cobra.Command{
        Use:   "command-name <args>",
        Short: "Description",
        Args:  cobra.ExactArgs(n),
        RunE:  runCommand,
    }
    cmd.Flags().StringVar(&flagVar, "flag", "default", "description")
    return cmd
}()
```

## Key Implementation Details

### A1 Notation Handling

- `a1ToGrid(cell string)` - Converts "A1" to (0,0) grid coordinates
- `parseRange(rangeA1 string)` - Parses "A1:B10" to start/end coordinates
- All coordinates are 0-indexed internally

### Color Handling

- `parseColor(hexColor string)` - Converts "#ff0000" to RGB values (0.0-1.0 range)
- Colors in Google Sheets API use float values from 0.0 to 1.0

### Format Patterns

Default patterns stored in `getDefaultFormatPattern()`:
- NUMBER: `#,##0.00`
- CURRENCY: `$#,##0.00`
- DATE: `yyyy-mm-dd`
- PERCENT: `0.00%`
- TIME: `hh:mm:ss`

## Command Reference

### create
Creates spreadsheet or copies from template. Can place in specific folder using Drive API.

**Flags**:
- `--template` - Template spreadsheet ID to copy
- `--folder` - Folder ID for placement

**Output**: JSON with `id` and `url`

### add-data
Updates cell values with JSON array data.

**Flags**:
- `--formula` (default: true) - Use USER_ENTERED mode for formulas

**Input format**: `'[["row1col1", "row1col2"], ["row2col1", "row2col2"]]'`

### import-csv
Reads CSV file and imports to sheet.

**Flags**:
- `--start` (default: "A1") - Starting cell position

**Process**: CSV → [][]interface{} → Sheets API

### format-cells
Applies number formatting to cells.

**Format types**: NUMBER, CURRENCY, DATE, PERCENT, TIME, TEXT

**Flags**:
- `--pattern` - Custom format pattern (overrides defaults)

**Implementation**: Uses `RepeatCellRequest` with `NumberFormat`

### style-cells
Applies visual styling to cells.

**Flags**:
- `--bg-color` - Background color (hex)
- `--font-color` - Font color (hex)
- `--font-size` - Font size (int)
- `--bold` - Bold text (bool)
- `--italic` - Italic text (bool)

**Implementation**: Uses `RepeatCellRequest` with `CellFormat.TextFormat`

### export-csv
Exports sheet data to CSV file.

**Process**: Sheets API → [][]interface{} → CSV Writer

### create-sheet
Adds new sheet to existing spreadsheet.

**Implementation**: Uses `AddSheetRequest` with `BatchUpdate`

### rename-sheet
Renames existing sheet.

**Implementation**: Uses `UpdateSheetPropertiesRequest` with `BatchUpdate`

### list-sheets
Lists all sheets with IDs, titles, and indices.

**Output**: JSON array of sheet objects

### add-note
Adds note/comment to specific cell.

**Implementation**: Uses `UpdateCellsRequest` with note field

## Error Handling

- All errors use `fmt.Errorf()` with `%w` for proper error wrapping
- Authentication errors provide helpful messages with setup instructions
- API errors are wrapped with context about the operation
- Context is properly propagated (no `context.TODO()` in production code)

## Common Patterns

### JSON Output

All commands return JSON via `printJSON()` helper:
```go
output := map[string]string{
    "status": "success",
    "key": "value",
}
return printJSON(output)
```

### Sheet ID Resolution

Helper `getSheetID()` resolves sheet name to numeric ID:
```go
sheetID, err := getSheetID(service, spreadsheetID, sheetName)
```

### Range Operations

For range-based operations:
1. Parse range: `startCol, startRow, endCol, endRow, err := parseRange(rangeA1)`
2. Get sheet ID: `sheetID, err := getSheetID(service, spreadsheetID, sheetName)`
3. Create GridRange with 0-indexed coordinates
4. Apply operation via BatchUpdate

## Development Guidelines

### Adding New Commands

1. Define flag variables at package level
2. Create command using IIFE pattern (no `init()` functions)
3. Implement `runCommandName()` function
4. Register in `main.go` with `rootCmd.AddCommand()`
5. Return JSON output for consistency
6. Wrap errors with context using `%w`

### Testing Considerations

- Mock Google Sheets API for unit tests
- Test A1 notation parsing edge cases
- Test color parsing with invalid inputs
- Test CSV import/export with various encodings
- Test error handling paths

### Code Style

- Functions ordered alphabetically within helpers.go
- Commands defined before run functions in cli.go
- All package-level variables grouped at top
- Context passed as first parameter
- Use `context.Background()` only in command entry points
- Prefer `strings.Join()` over manual concatenation

## Future Enhancements

Potential improvements:
- Add batch operations support
- Implement conditional formatting
- Add data validation rules
- Support for charts and images
- Implement sharing/permissions management
- Add support for protected ranges
- Implement filter and sort operations
- Add support for named ranges
- Implement pivot tables

## Debugging

### Common Issues

**Authentication failures**:
- Check credentials file exists at `~/.credentials/google_credentials.json`
- Verify OAuth scopes include spreadsheets and drive.file
- Delete token file to re-authenticate

**API errors**:
- Enable verbose logging by adding debug output
- Check spreadsheet/sheet IDs are valid
- Verify user has edit permissions on spreadsheet

**Format issues**:
- JSON input must use double quotes
- Hex colors must include # or be 6 characters
- Range notation must be valid A1 format (e.g., "A1:B10")

## Build System

Makefile targets:
- `build` - Compile binary
- `rebuild` - Clean all and rebuild
- `install` - Install to system (default: /usr/local/bin)
- `uninstall` - Remove from system
- `clean` - Remove binary
- `clean-all` - Remove binary, go.mod, go.sum
- `test` - Run tests
- `fmt` - Format code
- `vet` - Run go vet
- `check` - Run fmt, vet, and test

Custom install location:
```bash
TARGET=/custom/path make install
```

## API Rate Limits

Google Sheets API quotas (as of writing):
- 100 requests per 100 seconds per user
- 500 requests per 100 seconds per project

Handle rate limits by:
- Implementing exponential backoff
- Batching operations when possible
- Using batch update instead of individual updates

## Security Considerations

- Credentials stored in `~/.credentials/` with restrictive permissions (0600)
- OAuth2 tokens have limited lifetime and auto-refresh
- No credentials in code or repository
- Scopes limited to necessary permissions only

## Performance Tips

- Use batch updates for multiple operations
- Import large CSV files in chunks if needed
- Cache sheet IDs to avoid repeated lookups
- Use USER_ENTERED mode only when formulas needed
- Combine style operations into single batch request
