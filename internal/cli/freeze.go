package cli

import (
	"context"
	"fmt"

	"github.com/spf13/cobra"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
	"spreadsheet-manager/internal/helpers"
)

var (
	freezeRows int
	freezeCols int
)

var freezeCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "freeze <spreadsheet-id> <sheet-name>",
		Short: "Freeze rows and/or columns",
		Args:  cobra.ExactArgs(2),
		RunE:  runFreeze,
	}
	cmd.Flags().IntVar(&freezeRows, "rows", 0, "Number of rows to freeze")
	cmd.Flags().IntVar(&freezeCols, "cols", 0, "Number of columns to freeze")
	return cmd
}()

func runFreeze(cmd *cobra.Command, args []string) error {
	if freezeRows == 0 && freezeCols == 0 {
		return fmt.Errorf("at least one of --rows or --cols is required")
	}

	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return err
	}

	sheetID, err := helpers.GetSheetID(service, spreadsheetID, sheetName)
	if err != nil {
		return err
	}

	gridProps := &sheets.GridProperties{}
	var fields []string

	if freezeRows > 0 {
		gridProps.FrozenRowCount = int64(freezeRows)
		fields = append(fields, "gridProperties.frozenRowCount")
	}
	if freezeCols > 0 {
		gridProps.FrozenColumnCount = int64(freezeCols)
		fields = append(fields, "gridProperties.frozenColumnCount")
	}

	req := &sheets.Request{
		UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
			Properties: &sheets.SheetProperties{
				SheetId:        sheetID,
				GridProperties: gridProps,
			},
			Fields: joinFields(fields),
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetID, batchReq).Do()
	if err != nil {
		return fmt.Errorf("unable to freeze rows/columns: %w", err)
	}

	return helpers.PrintJSON(map[string]interface{}{
		"status":      "success",
		"frozen_rows": freezeRows,
		"frozen_cols": freezeCols,
	})
}

func joinFields(fields []string) string {
	result := ""
	for i, f := range fields {
		if i > 0 {
			result += ","
		}
		result += f
	}
	return result
}
