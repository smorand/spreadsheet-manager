package cli

import (
	"context"
	"fmt"
	"strings"

	"github.com/spf13/cobra"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
	"spreadsheet-manager/internal/helpers"
)

var setTextWrapMode string

var setTextWrapCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "set-text-wrap <spreadsheet-id> <sheet-name> <range>",
		Short: "Set text wrapping mode (WRAP, CLIP, OVERFLOW)",
		Args:  cobra.ExactArgs(3),
		RunE:  runSetTextWrap,
	}
	cmd.Flags().StringVar(&setTextWrapMode, "mode", "WRAP", "Wrap mode: WRAP, CLIP, or OVERFLOW_CELL")
	return cmd
}()

func runSetTextWrap(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	rangeA1 := args[2]

	mode := strings.ToUpper(setTextWrapMode)
	validModes := map[string]string{
		"WRAP":          "WRAP",
		"CLIP":          "CLIP",
		"OVERFLOW":      "OVERFLOW_CELL",
		"OVERFLOW_CELL": "OVERFLOW_CELL",
	}
	wrapStrategy, ok := validModes[mode]
	if !ok {
		return fmt.Errorf("invalid wrap mode: %s (use WRAP, CLIP, or OVERFLOW)", setTextWrapMode)
	}

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

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetID, batchReq).Do()
	if err != nil {
		return fmt.Errorf("unable to set text wrap: %w", err)
	}

	return helpers.PrintJSON(map[string]string{
		"status":        "success",
		"wrap_strategy": wrapStrategy,
	})
}
