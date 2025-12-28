package cli

import (
	"context"
	"fmt"

	"github.com/spf13/cobra"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
	"spreadsheet-manager/internal/helpers"
)

var formatCellsPattern string

var formatCellsCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "format-cells <spreadsheet-id> <sheet-name> <range> <format-type>",
		Short: "Format cells (NUMBER, CURRENCY, DATE, PERCENT, TIME, TEXT)",
		Args:  cobra.ExactArgs(4),
		RunE:  runFormatCells,
	}
	cmd.Flags().StringVar(&formatCellsPattern, "pattern", "", "Custom format pattern")
	return cmd
}()

func runFormatCells(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	rangeA1 := args[2]
	formatType := args[3]

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return err
	}

	sheetID, err := helpers.GetSheetID(service, spreadsheetID, sheetName)
	if err != nil {
		return err
	}

	startCol, startRow, endCol, endRow, err := helpers.ParseRange(rangeA1)
	if err != nil {
		return err
	}

	pattern := formatCellsPattern
	if pattern == "" {
		pattern = helpers.GetDefaultFormatPattern(formatType)
	}

	numFormat := &sheets.NumberFormat{
		Type:    formatType,
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

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetID, batchReq).Do()
	if err != nil {
		return fmt.Errorf("unable to format cells: %w", err)
	}

	return helpers.PrintJSON(map[string]string{
		"status": "success",
		"format": formatType,
	})
}
