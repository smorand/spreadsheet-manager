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

var (
	alignmentHorizontal string
	alignmentVertical   string
)

var setAlignmentCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "set-alignment <spreadsheet-id> <sheet-name> <range>",
		Short: "Set horizontal and/or vertical alignment",
		Args:  cobra.ExactArgs(3),
		RunE:  runSetAlignment,
	}
	cmd.Flags().StringVar(&alignmentHorizontal, "horizontal", "", "Horizontal alignment: LEFT, CENTER, RIGHT")
	cmd.Flags().StringVar(&alignmentVertical, "vertical", "", "Vertical alignment: TOP, MIDDLE, BOTTOM")
	return cmd
}()

func runSetAlignment(cmd *cobra.Command, args []string) error {
	if alignmentHorizontal == "" && alignmentVertical == "" {
		return fmt.Errorf("at least one of --horizontal or --vertical is required")
	}

	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	rangeA1 := args[2]

	cellFormat := &sheets.CellFormat{}
	var fields []string

	if alignmentHorizontal != "" {
		h := strings.ToUpper(alignmentHorizontal)
		validH := map[string]bool{"LEFT": true, "CENTER": true, "RIGHT": true}
		if !validH[h] {
			return fmt.Errorf("invalid horizontal alignment: %s (use LEFT, CENTER, or RIGHT)", alignmentHorizontal)
		}
		cellFormat.HorizontalAlignment = h
		fields = append(fields, "userEnteredFormat.horizontalAlignment")
	}

	if alignmentVertical != "" {
		v := strings.ToUpper(alignmentVertical)
		validV := map[string]bool{"TOP": true, "MIDDLE": true, "BOTTOM": true}
		if !validV[v] {
			return fmt.Errorf("invalid vertical alignment: %s (use TOP, MIDDLE, or BOTTOM)", alignmentVertical)
		}
		cellFormat.VerticalAlignment = v
		fields = append(fields, "userEnteredFormat.verticalAlignment")
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
				UserEnteredFormat: cellFormat,
			},
			Fields: strings.Join(fields, ","),
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetID, batchReq).Do()
	if err != nil {
		return fmt.Errorf("unable to set alignment: %w", err)
	}

	return helpers.PrintJSON(map[string]string{
		"status": "success",
	})
}
