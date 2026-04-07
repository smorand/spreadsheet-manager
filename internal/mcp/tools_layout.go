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

// handleFreeze freezes rows and/or columns using UpdateSheetProperties with GridProperties.
// At least one of rows or cols must be provided.
// Returns {status, frozen_rows, frozen_cols}.
func (h *Handler) handleFreeze(ctx context.Context, _ *mcp.CallToolRequest, input FreezeInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}
	if input.SheetName == "" {
		return nil, nil, fmt.Errorf("sheet_name is required")
	}
	if input.Rows == 0 && input.Cols == 0 {
		return nil, nil, fmt.Errorf("at least one of rows or cols is required")
	}

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get Sheets service: %w", err)
	}

	sheetID, err := helpers.GetSheetID(service, input.SpreadsheetID, input.SheetName)
	if err != nil {
		return nil, nil, err
	}

	gridProps := &sheets.GridProperties{}
	var fields []string

	if input.Rows > 0 {
		gridProps.FrozenRowCount = int64(input.Rows)
		fields = append(fields, "gridProperties.frozenRowCount")
	}
	if input.Cols > 0 {
		gridProps.FrozenColumnCount = int64(input.Cols)
		fields = append(fields, "gridProperties.frozenColumnCount")
	}

	req := &sheets.Request{
		UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
			Properties: &sheets.SheetProperties{
				SheetId:        sheetID,
				GridProperties: gridProps,
			},
			Fields: strings.Join(fields, ","),
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(input.SpreadsheetID, batchReq).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to freeze rows/columns: %w", err)
	}

	return nil, map[string]interface{}{
		"status":      "success",
		"frozen_rows": input.Rows,
		"frozen_cols": input.Cols,
	}, nil
}

// handleSetColumnWidth sets the pixel width of columns using UpdateDimensionProperties.
// Returns {status, columns, width}.
func (h *Handler) handleSetColumnWidth(ctx context.Context, _ *mcp.CallToolRequest, input SetColumnWidthInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}
	if input.SheetName == "" {
		return nil, nil, fmt.Errorf("sheet_name is required")
	}
	if input.Columns == "" {
		return nil, nil, fmt.Errorf("columns is required")
	}
	if input.Width <= 0 {
		return nil, nil, fmt.Errorf("width must be a positive integer")
	}

	startCol, endCol, err := helpers.ParseColumns(input.Columns)
	if err != nil {
		return nil, nil, err
	}

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get Sheets service: %w", err)
	}

	sheetID, err := helpers.GetSheetID(service, input.SpreadsheetID, input.SheetName)
	if err != nil {
		return nil, nil, err
	}

	req := &sheets.Request{
		UpdateDimensionProperties: &sheets.UpdateDimensionPropertiesRequest{
			Range: &sheets.DimensionRange{
				SheetId:    sheetID,
				Dimension:  "COLUMNS",
				StartIndex: int64(startCol),
				EndIndex:   int64(endCol + 1),
			},
			Properties: &sheets.DimensionProperties{
				PixelSize: int64(input.Width),
			},
			Fields: "pixelSize",
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(input.SpreadsheetID, batchReq).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to set column width: %w", err)
	}

	return nil, map[string]interface{}{
		"status":  "success",
		"columns": input.Columns,
		"width":   input.Width,
	}, nil
}

// handleSetTextWrap sets the text wrapping mode (WRAP, CLIP, OVERFLOW) using RepeatCellRequest.
// Returns {status, wrap_strategy}.
func (h *Handler) handleSetTextWrap(ctx context.Context, _ *mcp.CallToolRequest, input SetTextWrapInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}
	if input.SheetName == "" {
		return nil, nil, fmt.Errorf("sheet_name is required")
	}
	if input.Range == "" {
		return nil, nil, fmt.Errorf("range is required")
	}

	mode := input.Mode
	if mode == "" {
		mode = "WRAP"
	}
	mode = strings.ToUpper(mode)

	validModes := map[string]string{
		"WRAP":          "WRAP",
		"CLIP":          "CLIP",
		"OVERFLOW":      "OVERFLOW_CELL",
		"OVERFLOW_CELL": "OVERFLOW_CELL",
	}
	wrapStrategy, ok := validModes[mode]
	if !ok {
		return nil, nil, fmt.Errorf("invalid wrap mode: %s (use WRAP, CLIP, or OVERFLOW)", input.Mode)
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
					WrapStrategy: wrapStrategy,
				},
			},
			Fields: "userEnteredFormat.wrapStrategy",
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(input.SpreadsheetID, batchReq).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to set text wrap: %w", err)
	}

	return nil, map[string]string{
		"status":        "success",
		"wrap_strategy": wrapStrategy,
	}, nil
}

// handleSetAlignment sets horizontal and/or vertical alignment using RepeatCellRequest.
// At least one of horizontal or vertical must be provided.
// Returns {status}.
func (h *Handler) handleSetAlignment(ctx context.Context, _ *mcp.CallToolRequest, input SetAlignmentInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}
	if input.SheetName == "" {
		return nil, nil, fmt.Errorf("sheet_name is required")
	}
	if input.Range == "" {
		return nil, nil, fmt.Errorf("range is required")
	}
	if input.Horizontal == "" && input.Vertical == "" {
		return nil, nil, fmt.Errorf("at least one of horizontal or vertical is required")
	}

	cellFormat := &sheets.CellFormat{}
	var fields []string

	if input.Horizontal != "" {
		hAlign := strings.ToUpper(input.Horizontal)
		validH := map[string]bool{"LEFT": true, "CENTER": true, "RIGHT": true}
		if !validH[hAlign] {
			return nil, nil, fmt.Errorf("invalid horizontal alignment: %s (use LEFT, CENTER, or RIGHT)", input.Horizontal)
		}
		cellFormat.HorizontalAlignment = hAlign
		fields = append(fields, "userEnteredFormat.horizontalAlignment")
	}

	if input.Vertical != "" {
		vAlign := strings.ToUpper(input.Vertical)
		validV := map[string]bool{"TOP": true, "MIDDLE": true, "BOTTOM": true}
		if !validV[vAlign] {
			return nil, nil, fmt.Errorf("invalid vertical alignment: %s (use TOP, MIDDLE, or BOTTOM)", input.Vertical)
		}
		cellFormat.VerticalAlignment = vAlign
		fields = append(fields, "userEnteredFormat.verticalAlignment")
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
		return nil, nil, fmt.Errorf("unable to set alignment: %w", err)
	}

	return nil, map[string]string{
		"status": "success",
	}, nil
}
