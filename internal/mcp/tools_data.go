package mcp

import (
	"context"
	"encoding/json"
	"fmt"

	"github.com/modelcontextprotocol/go-sdk/mcp"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
	"spreadsheet-manager/internal/helpers"
)

const (
	valueInputModeFormula = "USER_ENTERED"
	valueInputModeRaw     = "RAW"
)

// handleAddData updates cells with JSON array data.
// Parses values_json as [][]interface{}. Default formula mode is true (USER_ENTERED).
// Returns {status, range}.
func (h *Handler) handleAddData(ctx context.Context, _ *mcp.CallToolRequest, input AddDataInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}
	if input.SheetName == "" {
		return nil, nil, fmt.Errorf("sheet_name is required")
	}
	if input.Range == "" {
		return nil, nil, fmt.Errorf("range is required")
	}
	if input.ValuesJSON == "" {
		return nil, nil, fmt.Errorf("values_json is required")
	}

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get Sheets service: %w", err)
	}

	var values [][]interface{}
	if err := json.Unmarshal([]byte(input.ValuesJSON), &values); err != nil {
		return nil, nil, fmt.Errorf("invalid JSON values: %w", err)
	}

	// Default formula mode is true
	formulaMode := true
	if input.FormulaMode != nil {
		formulaMode = *input.FormulaMode
	}

	valueInputOption := valueInputModeFormula
	if !formulaMode {
		valueInputOption = valueInputModeRaw
	}

	valueRange := &sheets.ValueRange{Values: values}
	rangeNotation := fmt.Sprintf("%s!%s", input.SheetName, input.Range)

	_, err = service.Spreadsheets.Values.Update(
		input.SpreadsheetID,
		rangeNotation,
		valueRange,
	).ValueInputOption(valueInputOption).Do()

	if err != nil {
		return nil, nil, fmt.Errorf("unable to update cells: %w", err)
	}

	return nil, map[string]string{
		"status": "success",
		"range":  rangeNotation,
	}, nil
}

// handleAddNote adds a note to a specific cell using UpdateCells with the note field.
// Returns {status, cell, note_length}.
func (h *Handler) handleAddNote(ctx context.Context, _ *mcp.CallToolRequest, input AddNoteInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}
	if input.SheetName == "" {
		return nil, nil, fmt.Errorf("sheet_name is required")
	}
	if input.Cell == "" {
		return nil, nil, fmt.Errorf("cell is required")
	}
	if input.Note == "" {
		return nil, nil, fmt.Errorf("note is required")
	}

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get Sheets service: %w", err)
	}

	sheetID, err := helpers.GetSheetID(service, input.SpreadsheetID, input.SheetName)
	if err != nil {
		return nil, nil, err
	}

	col, row, err := helpers.A1ToGrid(input.Cell)
	if err != nil {
		return nil, nil, err
	}

	req := &sheets.Request{
		UpdateCells: &sheets.UpdateCellsRequest{
			Range: &sheets.GridRange{
				SheetId:          sheetID,
				StartRowIndex:    int64(row),
				EndRowIndex:      int64(row + 1),
				StartColumnIndex: int64(col),
				EndColumnIndex:   int64(col + 1),
			},
			Rows: []*sheets.RowData{
				{
					Values: []*sheets.CellData{
						{Note: input.Note},
					},
				},
			},
			Fields: "note",
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(input.SpreadsheetID, batchReq).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to add note: %w", err)
	}

	return nil, map[string]interface{}{
		"status":      "success",
		"cell":        input.Cell,
		"note_length": len(input.Note),
	}, nil
}
