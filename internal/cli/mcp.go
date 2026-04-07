package cli

import (
	"context"
	"os"
	"strconv"

	"github.com/spf13/cobra"

	mcpserver "spreadsheet-manager/internal/mcp"
)

var (
	mcpPort           int
	mcpHost           string
	mcpBaseURL        string
	mcpSecretName     string
	mcpSecretProject  string
	mcpCredentialFile string
)

var mcpCmd = func() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "mcp",
		Short: "Start the MCP server",
		Long: `Start the MCP (Model Context Protocol) server for remote access.

The MCP server enables AI assistants to manage Google Sheets remotely
using the standard MCP protocol over HTTP Streamable transport.

Available tools:
  - spreadsheet_create: Create a new spreadsheet
  - spreadsheet_add_data: Add or update cell data
  - spreadsheet_import_csv: Import CSV file into a sheet
  - spreadsheet_export_csv: Export sheet data to CSV
  - spreadsheet_format_cells: Apply number formatting
  - spreadsheet_style_cells: Apply visual styling
  - spreadsheet_create_sheet: Create a new sheet tab
  - spreadsheet_rename_sheet: Rename a sheet
  - spreadsheet_list_sheets: List all sheets
  - spreadsheet_add_note: Add a note to a cell
  - spreadsheet_freeze: Freeze rows/columns
  - spreadsheet_set_column_width: Set column width
  - spreadsheet_set_text_wrap: Set text wrapping
  - spreadsheet_set_alignment: Set cell alignment
  - spreadsheet_alternate_row_colors: Add banded row colors

Authentication:
  The server implements OAuth 2.1 with Dynamic Client Registration
  for MCP-compatible clients. OAuth credentials are loaded from
  GCP Secret Manager (priority) or a local file (fallback).`,
		RunE: runMCP,
	}
	cmd.Flags().IntVarP(&mcpPort, "port", "p", 8080, "Port to listen on")
	cmd.Flags().StringVarP(&mcpHost, "host", "H", "localhost", "Host to bind to")
	cmd.Flags().StringVar(&mcpBaseURL, "base-url", "", "Base URL for OAuth callbacks (e.g., https://example.com)")
	cmd.Flags().StringVar(&mcpSecretName, "secret-name", "", "Secret Manager secret name for OAuth credentials")
	cmd.Flags().StringVar(&mcpSecretProject, "secret-project", "", "GCP project for Secret Manager")
	cmd.Flags().StringVar(&mcpCredentialFile, "credential-file", "", "Local OAuth credential file path (fallback)")
	return cmd
}()

func runMCP(cmd *cobra.Command, args []string) error {
	// Determine host: flag value, then HOST env var, then default
	host := mcpHost
	if host == "localhost" {
		if envHost := os.Getenv("HOST"); envHost != "" {
			host = envHost
		}
	}

	// Determine port: flag value, then PORT env var, then default
	port := mcpPort
	if envPort := os.Getenv("PORT"); envPort != "" {
		if p, err := strconv.Atoi(envPort); err == nil {
			port = p
		}
	}

	// Get Secret Manager project from env if not set via flag
	secretProject := mcpSecretProject
	if secretProject == "" {
		secretProject = os.Getenv("SECRET_PROJECT")
	}

	// Get Secret Manager secret name from env if not set via flag
	secretName := mcpSecretName
	if secretName == "" {
		secretName = os.Getenv("SECRET_NAME")
	}

	// Get base URL from env if not set via flag
	baseURL := mcpBaseURL
	if baseURL == "" {
		baseURL = os.Getenv("BASE_URL")
	}

	// Get credential file from env if not set via flag
	credentialFile := mcpCredentialFile
	if credentialFile == "" {
		credentialFile = os.Getenv("CREDENTIAL_FILE")
	}

	// Get Vault config from env
	vaultAddr := os.Getenv("VAULT_ADDR")
	vaultToken := os.Getenv("VAULT_TOKEN")
	vaultSecretPath := os.Getenv("VAULT_SECRET_PATH")

	// Create MCP server configuration
	cfg := &mcpserver.Config{
		Host:            host,
		Port:            port,
		BaseURL:         baseURL,
		SecretName:      secretName,
		SecretProject:   secretProject,
		VaultAddr:       vaultAddr,
		VaultToken:      vaultToken,
		VaultSecretPath: vaultSecretPath,
		CredentialFile:  credentialFile,
	}

	// Create server and register tools
	server := mcpserver.NewServer(cfg)
	handler := mcpserver.NewHandler()
	handler.RegisterTools(server.MCPServer())

	// Run the MCP server
	return server.Run(context.Background())
}
