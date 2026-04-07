package mcp

import (
	"context"
	"fmt"

	"github.com/modelcontextprotocol/go-sdk/mcp"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
	"spreadsheet-manager/internal/helpers"
)

const (
	defaultBandingColor1 = "#FFFFFF"
	defaultBandingColor2 = "#F5F5F5"
)

// handleAlternateRowColors adds banded rows using AddBanding.
// Default color1 is #FFFFFF, default color2 is #F5F5F5, optional header_color.
// Returns {status}.
func (h *Handler) handleAlternateRowColors(ctx context.Context, _ *mcp.CallToolRequest, input AlternateRowColorsInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}
	if input.SheetName == "" {
		return nil, nil, fmt.Errorf("sheet_name is required")
	}
	if input.Range == "" {
		return nil, nil, fmt.Errorf("range is required")
	}

	color1Hex := input.Color1
	if color1Hex == "" {
		color1Hex = defaultBandingColor1
	}
	color2Hex := input.Color2
	if color2Hex == "" {
		color2Hex = defaultBandingColor2
	}

	c1 := helpers.ParseColor(color1Hex)
	if c1 == nil {
		return nil, nil, fmt.Errorf("invalid color1: %s", color1Hex)
	}
	c2 := helpers.ParseColor(color2Hex)
	if c2 == nil {
		return nil, nil, fmt.Errorf("invalid color2: %s", color2Hex)
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

	rowProps := &sheets.BandingProperties{
		FirstBandColor:  c1,
		SecondBandColor: c2,
	}

	if input.HeaderColor != "" {
		hc := helpers.ParseColor(input.HeaderColor)
		if hc == nil {
			return nil, nil, fmt.Errorf("invalid header_color: %s", input.HeaderColor)
		}
		rowProps.HeaderColor = hc
	}

	req := &sheets.Request{
		AddBanding: &sheets.AddBandingRequest{
			BandedRange: &sheets.BandedRange{
				Range: &sheets.GridRange{
					SheetId:          sheetID,
					StartRowIndex:    int64(startRow),
					EndRowIndex:      int64(endRow + 1),
					StartColumnIndex: int64(startCol),
					EndColumnIndex:   int64(endCol + 1),
				},
				RowProperties: rowProps,
			},
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(input.SpreadsheetID, batchReq).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to add alternating row colors: %w", err)
	}

	return nil, map[string]string{
		"status": "success",
	}, nil
}
