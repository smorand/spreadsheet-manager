package mcp

import (
	"context"
	"fmt"

	"github.com/modelcontextprotocol/go-sdk/mcp"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
	"spreadsheet-manager/internal/helpers"
)

// handleFormatCells applies number formatting to cells.
// Supports format types: NUMBER, CURRENCY, DATE, PERCENT, TIME, TEXT.
// Uses RepeatCellRequest with NumberFormat.
// Returns {status, format}.
func (h *Handler) handleFormatCells(ctx context.Context, _ *mcp.CallToolRequest, input FormatCellsInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}
	if input.SheetName == "" {
		return nil, nil, fmt.Errorf("sheet_name is required")
	}
	if input.Range == "" {
		return nil, nil, fmt.Errorf("range is required")
	}
	if input.FormatType == "" {
		return nil, nil, fmt.Errorf("format_type is required")
	}

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get Sheets service: %w", err)
	}

	sheetID, err := helpers.GetSheetID(service, input.SpreadsheetID, input.SheetName)
	if err != nil {
		return nil, nil, err
	}

	startCol, startRow, endCol, endRow, err := helpers.ParseRange(input.Range)
	if err != nil {
		return nil, nil, err
	}

	pattern := input.Pattern
	if pattern == "" {
		pattern = helpers.GetDefaultFormatPattern(input.FormatType)
	}

	numFormat := &sheets.NumberFormat{
		Type:    input.FormatType,
		Pattern: pattern,
	}

	req := &sheets.Request{
		RepeatCell: &sheets.RepeatCellRequest{
			Range: &sheets.GridRange{
				SheetId:          sheetID,
				StartRowIndex:    int64(startRow),
				EndRowIndex:      int64(endRow + 1),
				StartColumnIndex: int64(startCol),
				EndColumnIndex:   int64(endCol + 1),
			},
			Cell: &sheets.CellData{
				UserEnteredFormat: &sheets.CellFormat{
					NumberFormat: numFormat,
				},
			},
			Fields: "userEnteredFormat.numberFormat",
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(input.SpreadsheetID, batchReq).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to format cells: %w", err)
	}

	return nil, map[string]string{
		"status": "success",
		"format": input.FormatType,
	}, nil
}
