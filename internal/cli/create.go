package cli

import (
	"context"
	"fmt"
	"os"

	"github.com/spf13/cobra"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
	"spreadsheet-manager/internal/helpers"
)

var (
	createTemplateID string
	createFolderID   string
)

var createCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "create <title>",
		Short: "Create a new spreadsheet",
		Args:  cobra.ExactArgs(1),
		RunE:  runCreate,
	}
	cmd.Flags().StringVar(&createTemplateID, "template", "", "Template spreadsheet ID to copy from")
	cmd.Flags().StringVar(&createFolderID, "folder", "", "Folder ID to create spreadsheet in")
	return cmd
}()

func runCreate(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	title := args[0]

	if createTemplateID != "" {
		return createFromTemplate(ctx, title, createTemplateID, createFolderID)
	}

	return createNew(ctx, title, createFolderID)
}

func createFromTemplate(ctx context.Context, title, templateID, folderID string) error {
	client, err := auth.GetClient(ctx)
	if err != nil {
		return err
	}

	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("unable to create Drive service: %w", err)
	}

	file := &drive.File{Name: title}
	if folderID != "" {
		file.Parents = []string{folderID}
	}

	result, err := driveService.Files.Copy(templateID, file).Do()
	if err != nil {
		return fmt.Errorf("unable to copy template: %w", err)
	}

	return helpers.PrintJSON(map[string]string{
		"id":  result.Id,
		"url": fmt.Sprintf(GoogleSheetsURLPattern, result.Id),
	})
}

func createNew(ctx context.Context, title, folderID string) error {
	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return err
	}

	spreadsheet := &sheets.Spreadsheet{
		Properties: &sheets.SpreadsheetProperties{
			Title: title,
		},
	}

	result, err := service.Spreadsheets.Create(spreadsheet).Do()
	if err != nil {
		return fmt.Errorf("unable to create spreadsheet: %w", err)
	}

	if folderID != "" {
		if err := moveToFolder(ctx, result.SpreadsheetId, folderID); err != nil {
			fmt.Fprintf(os.Stderr, "Warning: unable to move to folder: %v\n", err)
		}
	}

	return helpers.PrintJSON(map[string]string{
		"id":  result.SpreadsheetId,
		"url": fmt.Sprintf(GoogleSheetsURLPattern, result.SpreadsheetId),
	})
}

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
