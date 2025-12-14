package main

import (
	"context"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"os"
	"strings"

	"github.com/spf13/cobra"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

var (
	// createCmd flags
	templateID string
	folderID   string

	// addDataCmd flags
	formulaMode bool

	// importCSVCmd flags
	startCell string

	// formatCellsCmd flags
	pattern string

	// styleCellsCmd flags
	bgColor    string
	fontColor  string
	fontSize   int
	bold       bool
	italic     bool
)

// createCmd creates a new spreadsheet
var createCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "create <title>",
		Short: "Create a new spreadsheet",
		Args:  cobra.ExactArgs(1),
		RunE:  runCreate,
	}
	cmd.Flags().StringVar(&templateID, "template", "", "Template spreadsheet ID to copy from")
	cmd.Flags().StringVar(&folderID, "folder", "", "Folder ID to create spreadsheet in")
	return cmd
}()

func runCreate(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	title := args[0]

	if templateID != "" {
		// Create from template using Drive API
		client, err := getClient(ctx)
		if err != nil {
			return err
		}

		driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
		if err != nil {
			return fmt.Errorf("unable to create Drive service: %w", err)
		}

		file := &drive.File{
			Name: title,
		}
		if folderID != "" {
			file.Parents = []string{folderID}
		}

		result, err := driveService.Files.Copy(templateID, file).Do()
		if err != nil {
			return fmt.Errorf("unable to copy template: %w", err)
		}

		output := map[string]string{
			"id":  result.Id,
			"url": fmt.Sprintf("https://docs.google.com/spreadsheets/d/%s/edit", result.Id),
		}
		return printJSON(output)
	}

	// Create new spreadsheet
	service, err := getSheetsService(ctx)
	if err != nil {
		return err
	}

	ss := &sheets.Spreadsheet{
		Properties: &sheets.SpreadsheetProperties{
			Title: title,
		},
	}

	result, err := service.Spreadsheets.Create(ss).Do()
	if err != nil {
		return fmt.Errorf("unable to create spreadsheet: %w", err)
	}

	// Move to folder if specified
	if folderID != "" {
		client, err := getClient(ctx)
		if err != nil {
			return err
		}

		driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
		if err != nil {
			return fmt.Errorf("unable to create Drive service: %w", err)
		}

		_, err = driveService.Files.Update(result.SpreadsheetId, &drive.File{}).AddParents(folderID).Do()
		if err != nil {
			fmt.Fprintf(os.Stderr, "Warning: unable to move to folder: %v\n", err)
		}
	}

	output := map[string]string{
		"id":  result.SpreadsheetId,
		"url": fmt.Sprintf("https://docs.google.com/spreadsheets/d/%s/edit", result.SpreadsheetId),
	}
	return printJSON(output)
}

// addDataCmd adds data to cells
var addDataCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "add-data <spreadsheet-id> <sheet-name> <range> <values-json>",
		Short: "Add data to cells",
		Args:  cobra.ExactArgs(4),
		RunE:  runAddData,
	}
	cmd.Flags().BoolVar(&formulaMode, "formula", true, "Enable formula mode (USER_ENTERED)")
	return cmd
}()

func runAddData(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	rangeA1 := args[2]
	valuesJSON := args[3]

	service, err := getSheetsService(ctx)
	if err != nil {
		return err
	}

	var values [][]interface{}
	if err := json.Unmarshal([]byte(valuesJSON), &values); err != nil {
		return fmt.Errorf("invalid JSON values: %w", err)
	}

	valueInputOption := "USER_ENTERED"
	if !formulaMode {
		valueInputOption = "RAW"
	}

	vr := &sheets.ValueRange{
		Values: values,
	}

	_, err = service.Spreadsheets.Values.Update(
		spreadsheetID,
		fmt.Sprintf("%s!%s", sheetName, rangeA1),
		vr,
	).ValueInputOption(valueInputOption).Do()

	if err != nil {
		return fmt.Errorf("unable to update cells: %w", err)
	}

	output := map[string]string{
		"status": "success",
		"range":  fmt.Sprintf("%s!%s", sheetName, rangeA1),
	}
	return printJSON(output)
}

// importCSVCmd imports CSV data
var importCSVCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "import-csv <spreadsheet-id> <sheet-name> <csv-path>",
		Short: "Import CSV data into sheet",
		Args:  cobra.ExactArgs(3),
		RunE:  runImportCSV,
	}
	cmd.Flags().StringVar(&startCell, "start", "A1", "Starting cell")
	return cmd
}()

func runImportCSV(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	csvPath := args[2]

	// Read CSV file
	f, err := os.Open(csvPath)
	if err != nil {
		return fmt.Errorf("unable to open CSV file: %w", err)
	}
	defer f.Close()

	reader := csv.NewReader(f)
	records, err := reader.ReadAll()
	if err != nil {
		return fmt.Errorf("unable to read CSV: %w", err)
	}

	// Convert to interface{}[]
	var values [][]interface{}
	for _, record := range records {
		row := make([]interface{}, len(record))
		for i, v := range record {
			row[i] = v
		}
		values = append(values, row)
	}

	service, err := getSheetsService(ctx)
	if err != nil {
		return err
	}

	vr := &sheets.ValueRange{
		Values: values,
	}

	_, err = service.Spreadsheets.Values.Update(
		spreadsheetID,
		fmt.Sprintf("%s!%s", sheetName, startCell),
		vr,
	).ValueInputOption("USER_ENTERED").Do()

	if err != nil {
		return fmt.Errorf("unable to import CSV: %w", err)
	}

	output := map[string]interface{}{
		"status": "success",
		"rows":   len(values),
	}
	return printJSON(output)
}

// formatCellsCmd formats cells
var formatCellsCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "format-cells <spreadsheet-id> <sheet-name> <range> <format-type>",
		Short: "Format cells (NUMBER, CURRENCY, DATE, PERCENT, TIME, TEXT)",
		Args:  cobra.ExactArgs(4),
		RunE:  runFormatCells,
	}
	cmd.Flags().StringVar(&pattern, "pattern", "", "Custom format pattern")
	return cmd
}()

func runFormatCells(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	rangeA1 := args[2]
	formatType := args[3]

	service, err := getSheetsService(ctx)
	if err != nil {
		return err
	}

	sheetID, err := getSheetID(service, spreadsheetID, sheetName)
	if err != nil {
		return err
	}

	startCol, startRow, endCol, endRow, err := parseRange(rangeA1)
	if err != nil {
		return err
	}

	formatPattern := pattern
	if formatPattern == "" {
		formatPattern = getDefaultFormatPattern(formatType)
	}

	numFormat := &sheets.NumberFormat{
		Type:    formatType,
		Pattern: formatPattern,
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

	output := map[string]string{
		"status": "success",
		"format": formatType,
	}
	return printJSON(output)
}

// styleCellsCmd styles cells
var styleCellsCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "style-cells <spreadsheet-id> <sheet-name> <range>",
		Short: "Style cells with colors, fonts, and formatting",
		Args:  cobra.ExactArgs(3),
		RunE:  runStyleCells,
	}
	cmd.Flags().StringVar(&bgColor, "bg-color", "", "Background color (hex)")
	cmd.Flags().StringVar(&fontColor, "font-color", "", "Font color (hex)")
	cmd.Flags().IntVar(&fontSize, "font-size", 0, "Font size")
	cmd.Flags().BoolVar(&bold, "bold", false, "Bold text")
	cmd.Flags().BoolVar(&italic, "italic", false, "Italic text")
	return cmd
}()

func runStyleCells(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	rangeA1 := args[2]

	service, err := getSheetsService(ctx)
	if err != nil {
		return err
	}

	sheetID, err := getSheetID(service, spreadsheetID, sheetName)
	if err != nil {
		return err
	}

	startCol, startRow, endCol, endRow, err := parseRange(rangeA1)
	if err != nil {
		return err
	}

	cellFormat := &sheets.CellFormat{}
	var fields []string

	if bgColor != "" {
		cellFormat.BackgroundColor = parseColor(bgColor)
		fields = append(fields, "userEnteredFormat.backgroundColor")
	}

	if fontColor != "" || fontSize > 0 || bold || italic {
		textFormat := &sheets.TextFormat{}
		if fontColor != "" {
			textFormat.ForegroundColor = parseColor(fontColor)
		}
		if fontSize > 0 {
			textFormat.FontSize = int64(fontSize)
		}
		if bold {
			textFormat.Bold = true
		}
		if italic {
			textFormat.Italic = true
		}
		cellFormat.TextFormat = textFormat
		fields = append(fields, "userEnteredFormat.textFormat")
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
		return fmt.Errorf("unable to style cells: %w", err)
	}

	output := map[string]string{
		"status": "success",
	}
	return printJSON(output)
}

// exportCSVCmd exports sheet to CSV
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

	service, err := getSheetsService(ctx)
	if err != nil {
		return err
	}

	resp, err := service.Spreadsheets.Values.Get(spreadsheetID, sheetName).Do()
	if err != nil {
		return fmt.Errorf("unable to get sheet data: %w", err)
	}

	// Write to CSV
	f, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("unable to create CSV file: %w", err)
	}
	defer f.Close()

	writer := csv.NewWriter(f)
	defer writer.Flush()

	for _, row := range resp.Values {
		record := make([]string, len(row))
		for i, cell := range row {
			record[i] = fmt.Sprintf("%v", cell)
		}
		if err := writer.Write(record); err != nil {
			return fmt.Errorf("unable to write CSV row: %w", err)
		}
	}

	output := map[string]string{
		"status": "success",
		"file":   outputPath,
	}
	return printJSON(output)
}

// renameSheetCmd renames a sheet
var renameSheetCmd = &cobra.Command{
	Use:   "rename-sheet <spreadsheet-id> <old-name> <new-name>",
	Short: "Rename a sheet",
	Args:  cobra.ExactArgs(3),
	RunE:  runRenameSheet,
}

func runRenameSheet(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	oldName := args[1]
	newName := args[2]

	service, err := getSheetsService(ctx)
	if err != nil {
		return err
	}

	sheetID, err := getSheetID(service, spreadsheetID, oldName)
	if err != nil {
		return err
	}

	req := &sheets.Request{
		UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
			Properties: &sheets.SheetProperties{
				SheetId: sheetID,
				Title:   newName,
			},
			Fields: "title",
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetID, batchReq).Do()
	if err != nil {
		return fmt.Errorf("unable to rename sheet: %w", err)
	}

	output := map[string]string{
		"status":   "success",
		"old_name": oldName,
		"new_name": newName,
	}
	return printJSON(output)
}

// createSheetCmd creates a new sheet
var createSheetCmd = &cobra.Command{
	Use:   "create-sheet <spreadsheet-id> <sheet-name>",
	Short: "Create a new sheet in the spreadsheet",
	Args:  cobra.ExactArgs(2),
	RunE:  runCreateSheet,
}

func runCreateSheet(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]

	service, err := getSheetsService(ctx)
	if err != nil {
		return err
	}

	req := &sheets.Request{
		AddSheet: &sheets.AddSheetRequest{
			Properties: &sheets.SheetProperties{
				Title: sheetName,
			},
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetID, batchReq).Do()
	if err != nil {
		return fmt.Errorf("unable to create sheet: %w", err)
	}

	output := map[string]string{
		"status":     "success",
		"sheet_name": sheetName,
	}
	return printJSON(output)
}

// listSheetsCmd lists all sheets
var listSheetsCmd = &cobra.Command{
	Use:   "list-sheets <spreadsheet-id>",
	Short: "List all sheets in the spreadsheet",
	Args:  cobra.ExactArgs(1),
	RunE:  runListSheets,
}

func runListSheets(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]

	service, err := getSheetsService(ctx)
	if err != nil {
		return err
	}

	ss, err := service.Spreadsheets.Get(spreadsheetID).Do()
	if err != nil {
		return fmt.Errorf("unable to get spreadsheet: %w", err)
	}

	var sheetsList []map[string]interface{}
	for _, sheet := range ss.Sheets {
		sheetsList = append(sheetsList, map[string]interface{}{
			"sheet_id": sheet.Properties.SheetId,
			"title":    sheet.Properties.Title,
			"index":    sheet.Properties.Index,
		})
	}

	output := map[string]interface{}{
		"status": "success",
		"sheets": sheetsList,
	}
	return printJSON(output)
}

// addNoteCmd adds a note to a cell
var addNoteCmd = &cobra.Command{
	Use:   "add-note <spreadsheet-id> <sheet-name> <cell> <note>",
	Short: "Add a note to a cell",
	Args:  cobra.ExactArgs(4),
	RunE:  runAddNote,
}

func runAddNote(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	spreadsheetID := args[0]
	sheetName := args[1]
	cell := args[2]
	note := args[3]

	service, err := getSheetsService(ctx)
	if err != nil {
		return err
	}

	sheetID, err := getSheetID(service, spreadsheetID, sheetName)
	if err != nil {
		return err
	}

	col, row, err := a1ToGrid(cell)
	if err != nil {
		return err
	}

	req := &sheets.Request{
		UpdateCells: &sheets.UpdateCellsRequest{
			Range: &sheets.GridRange{
				SheetId:          sheetID,
				StartRowIndex:    int64(row),
				EndRowIndex:      int64(row + 1),
				StartColumnIndex: int64(col),
				EndColumnIndex:   int64(col + 1),
			},
			Rows: []*sheets.RowData{
				{
					Values: []*sheets.CellData{
						{
							Note: note,
						},
					},
				},
			},
			Fields: "note",
		},
	}

	batchReq := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{req},
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetID, batchReq).Do()
	if err != nil {
		return fmt.Errorf("unable to add note: %w", err)
	}

	output := map[string]interface{}{
		"status":      "success",
		"cell":        cell,
		"note_length": len(note),
	}
	return printJSON(output)
}

// Helper functions

func printJSON(v interface{}) error {
	data, err := json.MarshalIndent(v, "", "  ")
	if err != nil {
		return err
	}
	fmt.Println(string(data))
	return nil
}
