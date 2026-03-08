package cli

import (
	"context"
	"fmt"
	"strconv"

	"github.com/spf13/cobra"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
	"spreadsheet-manager/internal/helpers"
)

var setColumnWidthCmd = &cobra.Command{
	Use:   "set-column-width <spreadsheet-id> <sheet-name> <columns> <width>",
	Short: "Set pixel width of columns (e.g., columns: A, B:D)",
	Args:  cobra.ExactArgs(4),
	RunE:  runSetColumnWidth,
}

func runSetColumnWidth(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	columnsArg := args[2]
	widthArg := args[3]

	width, err := strconv.Atoi(widthArg)
	if err != nil {
		return fmt.Errorf("invalid width: %s", widthArg)
	}

	startCol, endCol, err := helpers.ParseColumns(columnsArg)
	if err != nil {
		return err
	}

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return err
	}

	sheetID, err := helpers.GetSheetID(service, spreadsheetID, sheetName)
	if err != nil {
		return err
	}

	req := &sheets.Request{
		UpdateDimensionProperties: &sheets.UpdateDimensionPropertiesRequest{
			Range: &sheets.DimensionRange{
				SheetId:    sheetID,
				Dimension:  "COLUMNS",
				StartIndex: int64(startCol),
				EndIndex:   int64(endCol + 1),
			},
			Properties: &sheets.DimensionProperties{
				PixelSize: int64(width),
			},
			Fields: "pixelSize",
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetID, batchReq).Do()
	if err != nil {
		return fmt.Errorf("unable to set column width: %w", err)
	}

	return helpers.PrintJSON(map[string]interface{}{
		"status":  "success",
		"columns": columnsArg,
		"width":   width,
	})
}
