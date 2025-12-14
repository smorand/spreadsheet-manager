package main

import (
	"fmt"
	"strconv"
	"strings"

	"google.golang.org/api/sheets/v4"
)

// a1ToGrid converts A1 notation to grid coordinates (0-indexed)
func a1ToGrid(cell string) (col int, row int, err error) {
	i := 0
	// Parse column letters
	for i < len(cell) && cell[i] >= 'A' && cell[i] <= 'Z' {
		col = col*26 + int(cell[i]-'A'+1)
		i++
	}

	// Parse row number
	if i < len(cell) {
		rowNum, err := strconv.Atoi(cell[i:])
		if err != nil {
			return 0, 0, fmt.Errorf("invalid cell reference: %s", cell)
		}
		row = rowNum
	}

	// Convert to 0-indexed
	col--
	row--

	return col, row, nil
}

// getDefaultFormatPattern returns default pattern for format type
func getDefaultFormatPattern(formatType string) string {
	patterns := map[string]string{
		"NUMBER":   "#,##0.00",
		"CURRENCY": "$#,##0.00",
		"DATE":     "yyyy-mm-dd",
		"PERCENT":  "0.00%",
		"TIME":     "hh:mm:ss",
	}

	if pattern, ok := patterns[formatType]; ok {
		return pattern
	}
	return ""
}

// getSheetID retrieves sheet ID by name
func getSheetID(service *sheets.Service, spreadsheetID, sheetName string) (int64, error) {
	ss, err := service.Spreadsheets.Get(spreadsheetID).Do()
	if err != nil {
		return 0, fmt.Errorf("unable to retrieve spreadsheet: %w", err)
	}

	for _, sheet := range ss.Sheets {
		if sheet.Properties.Title == sheetName {
			return sheet.Properties.SheetId, nil
		}
	}

	return 0, fmt.Errorf("sheet '%s' not found", sheetName)
}

// parseColor converts hex color to RGB (0-1 range)
func parseColor(hexColor string) *sheets.Color {
	hexColor = strings.TrimPrefix(hexColor, "#")
	if len(hexColor) != 6 {
		return nil
	}

	r, err1 := strconv.ParseInt(hexColor[0:2], 16, 64)
	g, err2 := strconv.ParseInt(hexColor[2:4], 16, 64)
	b, err3 := strconv.ParseInt(hexColor[4:6], 16, 64)

	if err1 != nil || err2 != nil || err3 != nil {
		return nil
	}

	return &sheets.Color{
		Red:   float64(r) / 255.0,
		Green: float64(g) / 255.0,
		Blue:  float64(b) / 255.0,
	}
}

// parseRange parses A1 notation range into grid coordinates
func parseRange(rangeA1 string) (startCol, startRow, endCol, endRow int, err error) {
	parts := strings.Split(rangeA1, ":")

	startCol, startRow, err = a1ToGrid(parts[0])
	if err != nil {
		return 0, 0, 0, 0, err
	}

	if len(parts) > 1 {
		endCol, endRow, err = a1ToGrid(parts[1])
		if err != nil {
			return 0, 0, 0, 0, err
		}
	} else {
		endCol, endRow = startCol, startRow
	}

	return startCol, startRow, endCol, endRow, nil
}
