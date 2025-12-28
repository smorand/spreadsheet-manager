package cli

import (
	"context"
	"fmt"

	"github.com/spf13/cobra"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
	"spreadsheet-manager/internal/helpers"
)

var createSheetCmd = &cobra.Command{
	Use:   "create-sheet <spreadsheet-id> <sheet-name>",
	Short: "Create a new sheet in the spreadsheet",
	Args:  cobra.ExactArgs(2),
	RunE:  runCreateSheet,
}

func runCreateSheet(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return err
	}

	req := &sheets.Request{
		AddSheet: &sheets.AddSheetRequest{
			Properties: &sheets.SheetProperties{
				Title: sheetName,
			},
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetID, batchReq).Do()
	if err != nil {
		return fmt.Errorf("unable to create sheet: %w", err)
	}

	return helpers.PrintJSON(map[string]string{
		"status":     "success",
		"sheet_name": sheetName,
	})
}

var renameSheetCmd = &cobra.Command{
	Use:   "rename-sheet <spreadsheet-id> <old-name> <new-name>",
	Short: "Rename a sheet",
	Args:  cobra.ExactArgs(3),
	RunE:  runRenameSheet,
}

func runRenameSheet(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	oldName := args[1]
	newName := args[2]

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return err
	}

	sheetID, err := helpers.GetSheetID(service, spreadsheetID, oldName)
	if err != nil {
		return err
	}

	req := &sheets.Request{
		UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
			Properties: &sheets.SheetProperties{
				SheetId: sheetID,
				Title:   newName,
			},
			Fields: "title",
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetID, batchReq).Do()
	if err != nil {
		return fmt.Errorf("unable to rename sheet: %w", err)
	}

	return helpers.PrintJSON(map[string]string{
		"status":   "success",
		"old_name": oldName,
		"new_name": newName,
	})
}

var listSheetsCmd = &cobra.Command{
	Use:   "list-sheets <spreadsheet-id>",
	Short: "List all sheets in the spreadsheet",
	Args:  cobra.ExactArgs(1),
	RunE:  runListSheets,
}

func runListSheets(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return err
	}

	spreadsheet, err := service.Spreadsheets.Get(spreadsheetID).Do()
	if err != nil {
		return fmt.Errorf("unable to get spreadsheet: %w", err)
	}

	var sheetsList []map[string]interface{}
	for _, sheet := range spreadsheet.Sheets {
		sheetsList = append(sheetsList, map[string]interface{}{
			"sheet_id": sheet.Properties.SheetId,
			"title":    sheet.Properties.Title,
			"index":    sheet.Properties.Index,
		})
	}

	return helpers.PrintJSON(map[string]interface{}{
		"status": "success",
		"sheets": sheetsList,
	})
}

var addNoteCmd = &cobra.Command{
	Use:   "add-note <spreadsheet-id> <sheet-name> <cell> <note>",
	Short: "Add a note to a cell",
	Args:  cobra.ExactArgs(4),
	RunE:  runAddNote,
}

func runAddNote(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	cell := args[2]
	note := args[3]

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return err
	}

	sheetID, err := helpers.GetSheetID(service, spreadsheetID, sheetName)
	if err != nil {
		return err
	}

	col, row, err := helpers.A1ToGrid(cell)
	if err != nil {
		return err
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
						{Note: note},
					},
				},
			},
			Fields: "note",
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetID, batchReq).Do()
	if err != nil {
		return fmt.Errorf("unable to add note: %w", err)
	}

	return helpers.PrintJSON(map[string]interface{}{
		"status":      "success",
		"cell":        cell,
		"note_length": len(note),
	})
}
