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
	styleCellsBgColor   string
	styleCellsFontColor string
	styleCellsFontSize  int
	styleCellsBold      bool
	styleCellsItalic    bool
)

var styleCellsCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "style-cells <spreadsheet-id> <sheet-name> <range>",
		Short: "Style cells with colors, fonts, and formatting",
		Args:  cobra.ExactArgs(3),
		RunE:  runStyleCells,
	}
	cmd.Flags().StringVar(&styleCellsBgColor, "bg-color", "", "Background color (hex)")
	cmd.Flags().StringVar(&styleCellsFontColor, "font-color", "", "Font color (hex)")
	cmd.Flags().IntVar(&styleCellsFontSize, "font-size", 0, "Font size")
	cmd.Flags().BoolVar(&styleCellsBold, "bold", false, "Bold text")
	cmd.Flags().BoolVar(&styleCellsItalic, "italic", false, "Italic text")
	return cmd
}()

func runStyleCells(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	rangeA1 := args[2]

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

	cellFormat, fields := buildCellFormat()

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
		return fmt.Errorf("unable to style cells: %w", err)
	}

	return helpers.PrintJSON(map[string]string{
		"status": "success",
	})
}

func buildCellFormat() (*sheets.CellFormat, []string) {
	cellFormat := &sheets.CellFormat{}
	var fields []string

	if styleCellsBgColor != "" {
		cellFormat.BackgroundColor = helpers.ParseColor(styleCellsBgColor)
		fields = append(fields, "userEnteredFormat.backgroundColor")
	}

	if styleCellsFontColor != "" || styleCellsFontSize > 0 || styleCellsBold || styleCellsItalic {
		textFormat := &sheets.TextFormat{}
		if styleCellsFontColor != "" {
			textFormat.ForegroundColor = helpers.ParseColor(styleCellsFontColor)
		}
		if styleCellsFontSize > 0 {
			textFormat.FontSize = int64(styleCellsFontSize)
		}
		if styleCellsBold {
			textFormat.Bold = true
		}
		if styleCellsItalic {
			textFormat.Italic = true
		}
		cellFormat.TextFormat = textFormat
		fields = append(fields, "userEnteredFormat.textFormat")
	}

	return cellFormat, fields
}
