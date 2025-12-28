package cli

import (
	"context"
	"encoding/json"
	"fmt"

	"github.com/spf13/cobra"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
	"spreadsheet-manager/internal/helpers"
)

var addDataFormulaMode bool

var addDataCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "add-data <spreadsheet-id> <sheet-name> <range> <values-json>",
		Short: "Add data to cells",
		Args:  cobra.ExactArgs(4),
		RunE:  runAddData,
	}
	cmd.Flags().BoolVar(&addDataFormulaMode, "formula", true, "Enable formula mode (USER_ENTERED)")
	return cmd
}()

func runAddData(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	rangeA1 := args[2]
	valuesJSON := args[3]

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return err
	}

	var values [][]interface{}
	if err := json.Unmarshal([]byte(valuesJSON), &values); err != nil {
		return fmt.Errorf("invalid JSON values: %w", err)
	}

	valueInputOption := ValueInputModeFormula
	if !addDataFormulaMode {
		valueInputOption = ValueInputModeRaw
	}

	valueRange := &sheets.ValueRange{Values: values}

	_, err = service.Spreadsheets.Values.Update(
		spreadsheetID,
		fmt.Sprintf("%s!%s", sheetName, rangeA1),
		valueRange,
	).ValueInputOption(valueInputOption).Do()

	if err != nil {
		return fmt.Errorf("unable to update cells: %w", err)
	}

	return helpers.PrintJSON(map[string]string{
		"status": "success",
		"range":  fmt.Sprintf("%s!%s", sheetName, rangeA1),
	})
}
