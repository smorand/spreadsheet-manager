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
	bandingColor1      string
	bandingColor2      string
	bandingHeaderColor string
)

var alternateRowColorsCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "alternate-row-colors <spreadsheet-id> <sheet-name> <range>",
		Short: "Add alternating row colors (banded rows)",
		Args:  cobra.ExactArgs(3),
		RunE:  runAlternateRowColors,
	}
	cmd.Flags().StringVar(&bandingColor1, "color1", "#FFFFFF", "First row color (hex)")
	cmd.Flags().StringVar(&bandingColor2, "color2", "#F5F5F5", "Second row color (hex)")
	cmd.Flags().StringVar(&bandingHeaderColor, "header-color", "", "Header row color (hex, optional)")
	return cmd
}()

func runAlternateRowColors(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	rangeA1 := args[2]

	c1 := helpers.ParseColor(bandingColor1)
	if c1 == nil {
		return fmt.Errorf("invalid color1: %s", bandingColor1)
	}
	c2 := helpers.ParseColor(bandingColor2)
	if c2 == nil {
		return fmt.Errorf("invalid color2: %s", bandingColor2)
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

	rowProps := &sheets.BandingProperties{
		FirstBandColor:  c1,
		SecondBandColor: c2,
	}

	if bandingHeaderColor != "" {
		hc := helpers.ParseColor(bandingHeaderColor)
		if hc == nil {
			return fmt.Errorf("invalid header-color: %s", bandingHeaderColor)
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

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetID, batchReq).Do()
	if err != nil {
		return fmt.Errorf("unable to add alternating row colors: %w", err)
	}

	return helpers.PrintJSON(map[string]string{
		"status": "success",
	})
}
