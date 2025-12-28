package helpers

import (
	"fmt"

	"google.golang.org/api/sheets/v4"
)

// GetSheetID retrieves the numeric sheet ID for a given sheet name
func GetSheetID(service *sheets.Service, spreadsheetID, sheetName string) (int64, error) {
	spreadsheet, err := service.Spreadsheets.Get(spreadsheetID).Do()
	if err != nil {
		return 0, fmt.Errorf("unable to retrieve spreadsheet: %w", err)
	}

	for _, sheet := range spreadsheet.Sheets {
		if sheet.Properties.Title == sheetName {
			return sheet.Properties.SheetId, nil
		}
	}

	return 0, fmt.Errorf("sheet '%s' not found", sheetName)
}
