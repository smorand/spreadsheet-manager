package main

import (
	"fmt"
	"os"

	"spreadsheet-manager/internal/cli"
	"spreadsheet-manager/internal/observability"
)

func main() {
	observability.InitLogger(os.Getenv("LOG_LEVEL"))

	if err := cli.RootCmd.Execute(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
