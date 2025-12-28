package cli

import (
	"context"
	"encoding/csv"
	"fmt"
	"os"

	"github.com/spf13/cobra"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
	"spreadsheet-manager/internal/helpers"
)

var importCSVStartCell string

var importCSVCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "import-csv <spreadsheet-id> <sheet-name> <csv-path>",
		Short: "Import CSV data into sheet",
		Args:  cobra.ExactArgs(3),
		RunE:  runImportCSV,
	}
	cmd.Flags().StringVar(&importCSVStartCell, "start", DefaultStartCell, "Starting cell")
	return cmd
}()

func runImportCSV(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	csvPath := args[2]

	values, err := readCSV(csvPath)
	if err != nil {
		return err
	}

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return err
	}

	valueRange := &sheets.ValueRange{Values: values}

	_, err = service.Spreadsheets.Values.Update(
		spreadsheetID,
		fmt.Sprintf("%s!%s", sheetName, importCSVStartCell),
		valueRange,
	).ValueInputOption(ValueInputModeFormula).Do()

	if err != nil {
		return fmt.Errorf("unable to import CSV: %w", err)
	}

	return helpers.PrintJSON(map[string]interface{}{
		"status": "success",
		"rows":   len(values),
	})
}

var exportCSVCmd = &cobra.Command{
	Use:   "export-csv <spreadsheet-id> <sheet-name> <output-path>",
	Short: "Export sheet to CSV file",
	Args:  cobra.ExactArgs(3),
	RunE:  runExportCSV,
}

func runExportCSV(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	outputPath := args[2]

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return err
	}

	resp, err := service.Spreadsheets.Values.Get(spreadsheetID, sheetName).Do()
	if err != nil {
		return fmt.Errorf("unable to get sheet data: %w", err)
	}

	if err := writeCSV(outputPath, resp.Values); err != nil {
		return err
	}

	return helpers.PrintJSON(map[string]string{
		"status": "success",
		"file":   outputPath,
	})
}

func readCSV(path string) ([][]interface{}, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("unable to open CSV file: %w", err)
	}
	defer file.Close()

	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		return nil, fmt.Errorf("unable to read CSV: %w", err)
	}

	values := make([][]interface{}, len(records))
	for i, record := range records {
		row := make([]interface{}, len(record))
		for j, v := range record {
			row[j] = v
		}
		values[i] = row
	}

	return values, nil
}

func writeCSV(path string, values [][]interface{}) error {
	file, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("unable to create CSV file: %w", err)
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	defer writer.Flush()

	for _, row := range values {
		record := make([]string, len(row))
		for i, cell := range row {
			record[i] = fmt.Sprintf("%v", cell)
		}
		if err := writer.Write(record); err != nil {
			return fmt.Errorf("unable to write CSV row: %w", err)
		}
	}

	return nil
}
