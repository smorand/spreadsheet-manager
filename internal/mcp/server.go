// Package mcp provides the MCP (Model Context Protocol) server implementation
// for spreadsheet-manager, enabling AI assistants to manage spreadsheets remotely.
package mcp

import (
	"context"
	"fmt"
	"log/slog"
	"net/http"
	"os"
	"os/signal"
	"strings"
	"syscall"
	"time"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"spreadsheet-manager/internal/auth"
)

// Config holds the MCP server configuration.
type Config struct {
	Host            string
	Port            int
	BaseURL         string // Base URL for OAuth callbacks (e.g., https://example.com)
	SecretName      string // Secret Manager secret name for OAuth credentials
	SecretProject   string // GCP project for Secret Manager
	VaultAddr       string // HashiCorp Vault address (e.g., http://vault:8200)
	VaultToken      string // HashiCorp Vault token
	VaultSecretPath string // Vault secret path (e.g., secret/data/spreadsheet-manager/scm-pwd-web)
	CredentialFile  string // Local credential file path (fallback)
}

// Server wraps the MCP server and HTTP server.
type Server struct {
	config       *Config
	mcpServer    *mcp.Server
	httpServer   *http.Server
	oauth2Server *OAuth2Server
}

// NewServer creates a new MCP server with the given configuration.
func NewServer(cfg *Config) *Server {
	// Create the MCP server
	mcpServer := mcp.NewServer(&mcp.Implementation{
		Name:    "spreadsheet-manager",
		Version: "1.0.0",
	}, nil)

	return &Server{
		config:    cfg,
		mcpServer: mcpServer,
	}
}

// MCPServer returns the underlying MCP server for tool registration.
func (s *Server) MCPServer() *mcp.Server {
	return s.mcpServer
}

// extractBearerToken extracts the token from the Authorization header.
// Expected format: "Bearer <token>"
func extractBearerToken(r *http.Request) string {
	authHeader := r.Header.Get("Authorization")
	if authHeader == "" {
		return ""
	}

	const bearerPrefix = "Bearer "
	if !strings.HasPrefix(authHeader, bearerPrefix) {
		return ""
	}

	return strings.TrimPrefix(authHeader, bearerPrefix)
}

// authMiddleware wraps an HTTP handler with OAuth2 Bearer token authentication.
// When no token is provided, returns 401 with WWW-Authenticate header pointing to the
// OAuth2 protected resource metadata endpoint (RFC 9728).
func (s *Server) authMiddleware(next http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		ctx := r.Context()

		// Extract Bearer token from Authorization header
		accessToken := extractBearerToken(r)

		// If no token provided, return 401 with proper WWW-Authenticate header
		if accessToken == "" {
			// RFC 9728: WWW-Authenticate header with resource_metadata URL
			w.Header().Set("WWW-Authenticate", fmt.Sprintf(
				`Bearer resource_metadata="%s/.well-known/oauth-protected-resource"`,
				s.config.BaseURL,
			))
			http.Error(w, "Unauthorized: Bearer token required", http.StatusUnauthorized)
			return
		}

		// Validate the access token and get OAuth config
		if s.oauth2Server == nil {
			http.Error(w, "OAuth not configured", http.StatusInternalServerError)
			return
		}

		oauthConfig, token, err := s.oauth2Server.ValidateAccessToken(ctx, accessToken)
		if err != nil {
			slog.Warn("token validation failed", "error", err)
			w.Header().Set("WWW-Authenticate", fmt.Sprintf(
				`Bearer error="invalid_token", resource_metadata="%s/.well-known/oauth-protected-resource"`,
				s.config.BaseURL,
			))
			http.Error(w, "Unauthorized: invalid token", http.StatusUnauthorized)
			return
		}

		// Inject OAuth config and token into context for Sheets API
		ctx = auth.WithOAuthConfig(ctx, oauthConfig)
		ctx = auth.WithAccessToken(ctx, token)

		r = r.WithContext(ctx)
		next.ServeHTTP(w, r)
	})
}

// Run starts the HTTP server and blocks until shutdown.
func (s *Server) Run(ctx context.Context) error {
	// Create the streamable HTTP handler for MCP
	mcpHandler := mcp.NewStreamableHTTPHandler(func(r *http.Request) *mcp.Server {
		return s.mcpServer
	}, &mcp.StreamableHTTPOptions{
		Stateless: false, // Enable session tracking
	})

	// Create HTTP mux for routing
	mux := http.NewServeMux()

	// Determine credential file path (default to local credentials)
	credFile := s.config.CredentialFile
	if credFile == "" {
		credFile = LocalCredentialsPath()
	}

	// Initialize OAuth2 server
	s.oauth2Server = NewOAuth2Server(&OAuth2ServerConfig{
		BaseURL:         s.config.BaseURL,
		SecretProject:   s.config.SecretProject,
		SecretName:      s.config.SecretName,
		VaultAddr:       s.config.VaultAddr,
		VaultToken:      s.config.VaultToken,
		VaultSecretPath: s.config.VaultSecretPath,
		CredentialFile:  credFile,
	})

	// Register OAuth2 routes (not protected by auth)
	s.oauth2Server.SetupRoutes(mux)
	slog.Info("oauth2 endpoints enabled",
		"endpoints", []string{
			"/.well-known/oauth-protected-resource",
			"/.well-known/oauth-authorization-server",
			"/oauth/register",
			"/oauth/authorize",
			"/oauth/callback",
			"/oauth/token",
		},
	)

	// Health check endpoint (not protected by auth)
	mux.HandleFunc("/health", func(w http.ResponseWriter, r *http.Request) {
		w.WriteHeader(http.StatusOK)
		w.Write([]byte("OK"))
	})

	// Wrap MCP handler with authentication middleware
	authedMCPHandler := s.authMiddleware(mcpHandler)

	// MCP endpoint (protected by OAuth2 Bearer token auth)
	mux.Handle("/mcp", authedMCPHandler)

	slog.Info("authentication mode set", "mode", "oauth2_bearer_tokens")

	// Create HTTP server
	addr := fmt.Sprintf("%s:%d", s.config.Host, s.config.Port)
	s.httpServer = &http.Server{
		Addr:    addr,
		Handler: mux,
	}

	// Setup graceful shutdown
	shutdown := make(chan os.Signal, 1)
	signal.Notify(shutdown, os.Interrupt, syscall.SIGTERM)

	// Start server in goroutine
	errChan := make(chan error, 1)
	go func() {
		slog.Info("starting mcp server", "addr", addr, "base_url", s.config.BaseURL)
		if err := s.httpServer.ListenAndServe(); err != nil && err != http.ErrServerClosed {
			errChan <- err
		}
	}()

	// Wait for shutdown signal or error
	select {
	case err := <-errChan:
		return fmt.Errorf("server error: %w", err)
	case sig := <-shutdown:
		slog.Info("received shutdown signal", "signal", sig)
	case <-ctx.Done():
		slog.Info("context cancelled, shutting down")
	}

	// Graceful shutdown with timeout
	shutdownCtx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()

	if err := s.httpServer.Shutdown(shutdownCtx); err != nil {
		return fmt.Errorf("shutdown error: %w", err)
	}

	slog.Info("mcp server stopped")
	return nil
}

// LocalCredentialsPath returns the default local credentials path for the MCP server.
func LocalCredentialsPath() string {
	return auth.GetCredentialsPath() + "/scm-pwd-web.json"
}
