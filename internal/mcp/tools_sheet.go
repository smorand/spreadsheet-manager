package mcp

import (
	"context"
	"fmt"

	"github.com/modelcontextprotocol/go-sdk/mcp"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
	"spreadsheet-manager/internal/helpers"
)

// handleListSheets lists all sheets in a spreadsheet with their id, title, and index.
// Returns {status, sheets: [...]}.
func (h *Handler) handleListSheets(ctx context.Context, _ *mcp.CallToolRequest, input ListSheetsInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get Sheets service: %w", err)
	}

	spreadsheet, err := service.Spreadsheets.Get(input.SpreadsheetID).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get spreadsheet: %w", err)
	}

	var sheetsList []map[string]interface{}
	for _, sheet := range spreadsheet.Sheets {
		sheetsList = append(sheetsList, map[string]interface{}{
			"sheet_id": sheet.Properties.SheetId,
			"title":    sheet.Properties.Title,
			"index":    sheet.Properties.Index,
		})
	}

	return nil, map[string]interface{}{
		"status": "success",
		"sheets": sheetsList,
	}, nil
}

// handleRenameSheet renames an existing sheet using UpdateSheetProperties.
// Returns {status, old_name, new_name}.
func (h *Handler) handleRenameSheet(ctx context.Context, _ *mcp.CallToolRequest, input RenameSheetInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}
	if input.OldName == "" {
		return nil, nil, fmt.Errorf("old_name is required")
	}
	if input.NewName == "" {
		return nil, nil, fmt.Errorf("new_name is required")
	}

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get Sheets service: %w", err)
	}

	sheetID, err := helpers.GetSheetID(service, input.SpreadsheetID, input.OldName)
	if err != nil {
		return nil, nil, err
	}

	req := &sheets.Request{
		UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
			Properties: &sheets.SheetProperties{
				SheetId: sheetID,
				Title:   input.NewName,
			},
			Fields: "title",
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(input.SpreadsheetID, batchReq).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to rename sheet: %w", err)
	}

	return nil, map[string]string{
		"status":   "success",
		"old_name": input.OldName,
		"new_name": input.NewName,
	}, nil
}
