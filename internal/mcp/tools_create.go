package mcp

import (
	"context"
	"fmt"

	"github.com/modelcontextprotocol/go-sdk/mcp"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
)

const (
	googleSheetsURLPattern = "https://docs.google.com/spreadsheets/d/%s/edit"
)

// handleCreate creates a new spreadsheet or copies from a template.
// If template_id is provided, uses Drive API to copy. If folder_id is provided, moves to folder.
// Returns {id, url}.
func (h *Handler) handleCreate(ctx context.Context, _ *mcp.CallToolRequest, input CreateInput) (*mcp.CallToolResult, any, error) {
	if input.Title == "" {
		return nil, nil, fmt.Errorf("title is required")
	}

	if input.TemplateID != "" {
		return h.createFromTemplate(ctx, input.Title, input.TemplateID, input.FolderID)
	}

	return h.createNew(ctx, input.Title, input.FolderID)
}

func (h *Handler) createFromTemplate(ctx context.Context, title, templateID, folderID string) (*mcp.CallToolResult, any, error) {
	client, err := auth.GetClient(ctx)
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get client: %w", err)
	}

	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, nil, fmt.Errorf("unable to create Drive service: %w", err)
	}

	file := &drive.File{Name: title}
	if folderID != "" {
		file.Parents = []string{folderID}
	}

	result, err := driveService.Files.Copy(templateID, file).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to copy template: %w", err)
	}

	return nil, map[string]string{
		"id":  result.Id,
		"url": fmt.Sprintf(googleSheetsURLPattern, result.Id),
	}, nil
}

func (h *Handler) createNew(ctx context.Context, title, folderID string) (*mcp.CallToolResult, any, error) {
	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get Sheets service: %w", err)
	}

	spreadsheet := &sheets.Spreadsheet{
		Properties: &sheets.SpreadsheetProperties{
			Title: title,
		},
	}

	result, err := service.Spreadsheets.Create(spreadsheet).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to create spreadsheet: %w", err)
	}

	if folderID != "" {
		if err := moveToFolder(ctx, result.SpreadsheetId, folderID); err != nil {
			// Non-fatal: spreadsheet was created but could not be moved
			return nil, map[string]any{
				"id":      result.SpreadsheetId,
				"url":     fmt.Sprintf(googleSheetsURLPattern, result.SpreadsheetId),
				"warning": fmt.Sprintf("unable to move to folder: %v", err),
			}, nil
		}
	}

	return nil, map[string]string{
		"id":  result.SpreadsheetId,
		"url": fmt.Sprintf(googleSheetsURLPattern, result.SpreadsheetId),
	}, nil
}

// handleCreateSheet creates a new sheet in an existing spreadsheet.
// Returns {status, sheet_name}.
func (h *Handler) handleCreateSheet(ctx context.Context, _ *mcp.CallToolRequest, input CreateSheetInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}
	if input.SheetName == "" {
		return nil, nil, fmt.Errorf("sheet_name is required")
	}

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get Sheets service: %w", err)
	}

	req := &sheets.Request{
		AddSheet: &sheets.AddSheetRequest{
			Properties: &sheets.SheetProperties{
				Title: input.SheetName,
			},
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(input.SpreadsheetID, batchReq).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to create sheet: %w", err)
	}

	return nil, map[string]string{
		"status":     "success",
		"sheet_name": input.SheetName,
	}, nil
}

// moveToFolder moves a spreadsheet to a specific Google Drive folder.
func moveToFolder(ctx context.Context, spreadsheetID, folderID string) error {
	client, err := auth.GetClient(ctx)
	if err != nil {
		return err
	}

	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("unable to create Drive service: %w", err)
	}

	_, err = driveService.Files.Update(spreadsheetID, &drive.File{}).AddParents(folderID).Do()
	return err
}
