package mcp

import (
	"context"
	"encoding/csv"
	"fmt"
	"os"

	"github.com/modelcontextprotocol/go-sdk/mcp"
	"google.golang.org/api/sheets/v4"

	"spreadsheet-manager/internal/auth"
)

const (
	defaultStartCell = "A1"
)

// handleImportCSV reads a CSV file and imports its contents to a sheet.
// Default start cell is A1. Returns {status, rows}.
func (h *Handler) handleImportCSV(ctx context.Context, _ *mcp.CallToolRequest, input ImportCSVInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}
	if input.SheetName == "" {
		return nil, nil, fmt.Errorf("sheet_name is required")
	}
	if input.CSVPath == "" {
		return nil, nil, fmt.Errorf("csv_path is required")
	}

	startCell := input.StartCell
	if startCell == "" {
		startCell = defaultStartCell
	}

	values, err := readCSV(input.CSVPath)
	if err != nil {
		return nil, nil, err
	}

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get Sheets service: %w", err)
	}

	valueRange := &sheets.ValueRange{Values: values}

	_, err = service.Spreadsheets.Values.Update(
		input.SpreadsheetID,
		fmt.Sprintf("%s!%s", input.SheetName, startCell),
		valueRange,
	).ValueInputOption(valueInputModeFormula).Do()

	if err != nil {
		return nil, nil, fmt.Errorf("unable to import CSV: %w", err)
	}

	return nil, map[string]interface{}{
		"status": "success",
		"rows":   len(values),
	}, nil
}

// handleExportCSV exports sheet data to a CSV file.
// Returns {status, file}.
func (h *Handler) handleExportCSV(ctx context.Context, _ *mcp.CallToolRequest, input ExportCSVInput) (*mcp.CallToolResult, any, error) {
	if input.SpreadsheetID == "" {
		return nil, nil, fmt.Errorf("spreadsheet_id is required")
	}
	if input.SheetName == "" {
		return nil, nil, fmt.Errorf("sheet_name is required")
	}
	if input.OutputPath == "" {
		return nil, nil, fmt.Errorf("output_path is required")
	}

	service, err := auth.GetSheetsService(ctx)
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get Sheets service: %w", err)
	}

	resp, err := service.Spreadsheets.Values.Get(input.SpreadsheetID, input.SheetName).Do()
	if err != nil {
		return nil, nil, fmt.Errorf("unable to get sheet data: %w", err)
	}

	if err := writeCSV(input.OutputPath, resp.Values); err != nil {
		return nil, nil, err
	}

	return nil, map[string]string{
		"status": "success",
		"file":   input.OutputPath,
	}, nil
}

// readCSV reads a CSV file and returns its contents as [][]interface{}.
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

// writeCSV writes data to a CSV file.
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
