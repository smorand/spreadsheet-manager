package main

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

var rootCmd = &cobra.Command{
	Use:   "spreadsheet-manager",
	Short: "Google Sheets Spreadsheet Manager",
	Long:  "Comprehensive spreadsheet operations: create, format, style, import/export",
}

func main() {
	// Register all commands
	rootCmd.AddCommand(createCmd)
	rootCmd.AddCommand(addDataCmd)
	rootCmd.AddCommand(importCSVCmd)
	rootCmd.AddCommand(formatCellsCmd)
	rootCmd.AddCommand(styleCellsCmd)
	rootCmd.AddCommand(exportCSVCmd)
	rootCmd.AddCommand(renameSheetCmd)
	rootCmd.AddCommand(createSheetCmd)
	rootCmd.AddCommand(listSheetsCmd)
	rootCmd.AddCommand(addNoteCmd)

	if err := rootCmd.Execute(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
