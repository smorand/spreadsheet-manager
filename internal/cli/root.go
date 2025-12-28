package cli

import "github.com/spf13/cobra"

var RootCmd = &cobra.Command{
	Use:   "spreadsheet-manager",
	Short: "Google Sheets Spreadsheet Manager",
	Long:  "Comprehensive spreadsheet operations: create, format, style, import/export",
}

func init() {
	RootCmd.AddCommand(addDataCmd)
	RootCmd.AddCommand(addNoteCmd)
	RootCmd.AddCommand(createCmd)
	RootCmd.AddCommand(createSheetCmd)
	RootCmd.AddCommand(exportCSVCmd)
	RootCmd.AddCommand(formatCellsCmd)
	RootCmd.AddCommand(importCSVCmd)
	RootCmd.AddCommand(listSheetsCmd)
	RootCmd.AddCommand(renameSheetCmd)
	RootCmd.AddCommand(styleCellsCmd)
}
