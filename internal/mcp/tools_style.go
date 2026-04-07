package mcp

import (
	"context"
	"fmt"
	"strings"

	"github.com/modelcontextprotocol/go-sdk/mcp"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
	"spreadsheet-manager/internal/helpers"
)

// handleStyleCells applies visual styling to cells.
// Builds CellFormat based on provided fields (bg_color, font_color, font_size, bold, italic).
// Uses RepeatCellRequest.
// Returns {status}.
func (h *Handler) handleStyleCells(ctx context.Context, _ *mcp.CallToolRequest, input StyleCellsInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}
	if input.SheetName == "" {
		return nil, nil, fmt.Errorf("sheet_name is required")
	}
	if input.Range == "" {
		return nil, nil, fmt.Errorf("range is required")
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

	cellFormat, fields := buildCellFormatFromInput(input)

	if len(fields) == 0 {
		return nil, nil, fmt.Errorf("at least one style property is required (bg_color, font_color, font_size, bold, or italic)")
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
				UserEnteredFormat: cellFormat,
			},
			Fields: strings.Join(fields, ","),
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(input.SpreadsheetID, batchReq).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to style cells: %w", err)
	}

	return nil, map[string]string{
		"status": "success",
	}, nil
}

// buildCellFormatFromInput constructs a CellFormat and field mask from StyleCellsInput.
func buildCellFormatFromInput(input StyleCellsInput) (*sheets.CellFormat, []string) {
	cellFormat := &sheets.CellFormat{}
	var fields []string

	if input.BgColor != "" {
		cellFormat.BackgroundColor = helpers.ParseColor(input.BgColor)
		fields = append(fields, "userEnteredFormat.backgroundColor")
	}

	if input.FontColor != "" || input.FontSize > 0 || input.Bold || input.Italic {
		textFormat := &sheets.TextFormat{}
		if input.FontColor != "" {
			textFormat.ForegroundColor = helpers.ParseColor(input.FontColor)
		}
		if input.FontSize > 0 {
			textFormat.FontSize = int64(input.FontSize)
		}
		if input.Bold {
			textFormat.Bold = true
		}
		if input.Italic {
			textFormat.Italic = true
		}
		cellFormat.TextFormat = textFormat
		fields = append(fields, "userEnteredFormat.textFormat")
	}

	return cellFormat, fields
}
