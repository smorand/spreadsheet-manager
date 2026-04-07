package mcp

import (
	"github.com/modelcontextprotocol/go-sdk/mcp"
)

// Handler provides MCP tool implementations for Google Sheets operations.
type Handler struct{}

// NewHandler creates a new Handler instance.
func NewHandler() *Handler {
	return &Handler{}
}

// RegisterTools registers all spreadsheet tools on the MCP server.
func (h *Handler) RegisterTools(s *mcp.Server) {
	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_create",
		Description: "Create a new Google Sheets spreadsheet, optionally from a template and/or in a specific Drive folder",
	}, h.handleCreate)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_add_data",
		Description: "Add or update cell data using a 2D JSON array of values",
	}, h.handleAddData)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_import_csv",
		Description: "Import a CSV file into a sheet",
	}, h.handleImportCSV)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_export_csv",
		Description: "Export sheet data to a CSV file",
	}, h.handleExportCSV)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_format_cells",
		Description: "Apply number formatting (NUMBER, CURRENCY, DATE, PERCENT, TIME, TEXT) to a cell range",
	}, h.handleFormatCells)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_style_cells",
		Description: "Apply visual styling (colors, font size, bold, italic) to a cell range",
	}, h.handleStyleCells)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_create_sheet",
		Description: "Create a new sheet (tab) in an existing spreadsheet",
	}, h.handleCreateSheet)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_rename_sheet",
		Description: "Rename an existing sheet",
	}, h.handleRenameSheet)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_list_sheets",
		Description: "List all sheets in a spreadsheet with their IDs, titles, and indices",
	}, h.handleListSheets)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_add_note",
		Description: "Add a note/comment to a specific cell",
	}, h.handleAddNote)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_freeze",
		Description: "Freeze rows and/or columns in a sheet",
	}, h.handleFreeze)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_set_column_width",
		Description: "Set the pixel width of one or more columns",
	}, h.handleSetColumnWidth)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_set_text_wrap",
		Description: "Set text wrapping mode (WRAP, CLIP, OVERFLOW) for a cell range",
	}, h.handleSetTextWrap)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_set_alignment",
		Description: "Set horizontal and/or vertical alignment for a cell range",
	}, h.handleSetAlignment)

	mcp.AddTool(s, &mcp.Tool{
		Name:        "spreadsheet_alternate_row_colors",
		Description: "Add alternating row colors (banded rows) to a cell range",
	}, h.handleAlternateRowColors)
}

// CreateInput holds parameters for the create spreadsheet tool.
type CreateInput struct {
	Title      string `json:"title"`
	TemplateID string `json:"template_id,omitempty"`
	FolderID   string `json:"folder_id,omitempty"`
}

// CreateSheetInput holds parameters for the create sheet tool.
type CreateSheetInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	SheetName     string `json:"sheet_name"`
}

// AddDataInput holds parameters for the add data tool.
type AddDataInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	SheetName     string `json:"sheet_name"`
	Range         string `json:"range"`
	ValuesJSON    string `json:"values_json"`
	FormulaMode   *bool  `json:"formula_mode,omitempty"`
}

// AddNoteInput holds parameters for the add note tool.
type AddNoteInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	SheetName     string `json:"sheet_name"`
	Cell          string `json:"cell"`
	Note          string `json:"note"`
}

// ImportCSVInput holds parameters for the import CSV tool.
type ImportCSVInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	SheetName     string `json:"sheet_name"`
	CSVPath       string `json:"csv_path"`
	StartCell     string `json:"start_cell,omitempty"`
}

// ExportCSVInput holds parameters for the export CSV tool.
type ExportCSVInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	SheetName     string `json:"sheet_name"`
	OutputPath    string `json:"output_path"`
}

// FormatCellsInput holds parameters for the format cells tool.
type FormatCellsInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	SheetName     string `json:"sheet_name"`
	Range         string `json:"range"`
	FormatType    string `json:"format_type"`
	Pattern       string `json:"pattern,omitempty"`
}

// StyleCellsInput holds parameters for the style cells tool.
type StyleCellsInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	SheetName     string `json:"sheet_name"`
	Range         string `json:"range"`
	BgColor       string `json:"bg_color,omitempty"`
	FontColor     string `json:"font_color,omitempty"`
	FontSize      int    `json:"font_size,omitempty"`
	Bold          bool   `json:"bold,omitempty"`
	Italic        bool   `json:"italic,omitempty"`
}

// ListSheetsInput holds parameters for the list sheets tool.
type ListSheetsInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
}

// RenameSheetInput holds parameters for the rename sheet tool.
type RenameSheetInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	OldName       string `json:"old_name"`
	NewName       string `json:"new_name"`
}

// FreezeInput holds parameters for the freeze tool.
type FreezeInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	SheetName     string `json:"sheet_name"`
	Rows          int    `json:"rows,omitempty"`
	Cols          int    `json:"cols,omitempty"`
}

// SetColumnWidthInput holds parameters for the set column width tool.
type SetColumnWidthInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	SheetName     string `json:"sheet_name"`
	Columns       string `json:"columns"`
	Width         int    `json:"width"`
}

// SetTextWrapInput holds parameters for the set text wrap tool.
type SetTextWrapInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	SheetName     string `json:"sheet_name"`
	Range         string `json:"range"`
	Mode          string `json:"mode,omitempty"`
}

// SetAlignmentInput holds parameters for the set alignment tool.
type SetAlignmentInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	SheetName     string `json:"sheet_name"`
	Range         string `json:"range"`
	Horizontal    string `json:"horizontal,omitempty"`
	Vertical      string `json:"vertical,omitempty"`
}

// AlternateRowColorsInput holds parameters for the alternate row colors tool.
type AlternateRowColorsInput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	SheetName     string `json:"sheet_name"`
	Range         string `json:"range"`
	Color1        string `json:"color1,omitempty"`
	Color2        string `json:"color2,omitempty"`
	HeaderColor   string `json:"header_color,omitempty"`
}
