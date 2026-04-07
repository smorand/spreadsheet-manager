package auth

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"time"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

const (
	CredentialsFile = "google_credentials.json"
	TokenFile       = "google_token_sheets.json"
)

var Scopes = []string{
	sheets.SpreadsheetsScope,
	sheets.DriveScope,
}

// contextKey is a type for context keys used in this package.
type contextKey string

// Context keys for storing authentication data.
const (
	oauthConfigKey contextKey = "oauth_config"
	accessTokenKey contextKey = "access_token"
)

// WithOAuthConfig returns a new context with the OAuth2 configuration stored.
// Used by the MCP server to pass the OAuth config loaded from Secret Manager.
func WithOAuthConfig(ctx context.Context, config *oauth2.Config) context.Context {
	return context.WithValue(ctx, oauthConfigKey, config)
}

// GetOAuthConfigFromContext retrieves the OAuth2 config from context, if present.
func GetOAuthConfigFromContext(ctx context.Context) (*oauth2.Config, bool) {
	config, ok := ctx.Value(oauthConfigKey).(*oauth2.Config)
	return config, ok
}

// WithAccessToken returns a new context with the OAuth2 token stored.
// Used by the MCP server to pass the validated access token.
func WithAccessToken(ctx context.Context, token *oauth2.Token) context.Context {
	return context.WithValue(ctx, accessTokenKey, token)
}

// GetAccessTokenFromContext retrieves the OAuth2 token from context, if present.
func GetAccessTokenFromContext(ctx context.Context) (*oauth2.Token, bool) {
	token, ok := ctx.Value(accessTokenKey).(*oauth2.Token)
	return token, ok
}

// GetCredentialsPath returns the path to the credentials directory.
func GetCredentialsPath() string {
	home, err := os.UserHomeDir()
	if err != nil {
		return ""
	}
	return filepath.Join(home, ".credentials")
}

// GetClient retrieves an OAuth2 HTTP client.
// In MCP mode: uses OAuth config and access token from context.
// In CLI mode: falls back to file-based credentials and token.
func GetClient(ctx context.Context) (*http.Client, error) {
	// Check if OAuth config is provided via context (MCP server mode)
	if ctxConfig, ok := GetOAuthConfigFromContext(ctx); ok && ctxConfig != nil {
		// Check if an access token is provided via context
		if token, ok := GetAccessTokenFromContext(ctx); ok && token != nil {
			return ctxConfig.Client(ctx, token), nil
		}
	}

	// Fall back to file-based credentials (CLI mode)
	credPath := filepath.Join(GetCredentialsPath(), CredentialsFile)
	tokenPath := filepath.Join(GetCredentialsPath(), TokenFile)

	credentials, err := os.ReadFile(credPath)
	if err != nil {
		return nil, fmt.Errorf("unable to read credentials file %s: %w\nSee README.md for setup instructions", credPath, err)
	}

	config, err := google.ConfigFromJSON(credentials, Scopes...)
	if err != nil {
		return nil, fmt.Errorf("unable to parse credentials: %w", err)
	}

	token, err := loadToken(tokenPath)
	if err != nil {
		token, err = requestTokenFromWeb(config)
		if err != nil {
			return nil, err
		}
		if err := saveToken(tokenPath, token); err != nil {
			fmt.Fprintf(os.Stderr, "Warning: unable to save token: %v\n", err)
		}
	}

	return config.Client(ctx, token), nil
}

// GetSheetsService creates an authenticated Google Sheets service
func GetSheetsService(ctx context.Context) (*sheets.Service, error) {
	client, err := GetClient(ctx)
	if err != nil {
		return nil, err
	}

	service, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("unable to create Sheets service: %w", err)
	}

	return service, nil
}

func loadToken(path string) (*oauth2.Token, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	token := &oauth2.Token{}
	err = json.NewDecoder(file).Decode(token)
	return token, err
}

func requestTokenFromWeb(config *oauth2.Config) (*oauth2.Token, error) {
	config.RedirectURL = "http://localhost:8080/oauth2callback"

	codeChan := make(chan string)
	errChan := make(chan error)

	server := &http.Server{Addr: ":8080"}
	http.HandleFunc("/oauth2callback", func(w http.ResponseWriter, r *http.Request) {
		code := r.URL.Query().Get("code")
		if code == "" {
			errChan <- fmt.Errorf("no code in callback")
			return
		}

		w.Header().Set("Content-Type", "text/html")
		fmt.Fprintf(w, `
			<html>
			<body>
				<h1>Authentication successful!</h1>
				<p>You can close this window and return to the terminal.</p>
			</body>
			</html>
		`)

		codeChan <- code
	})

	go func() {
		if err := server.ListenAndServe(); err != nil && err != http.ErrServerClosed {
			errChan <- err
		}
	}()

	time.Sleep(100 * time.Millisecond)

	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline, oauth2.ApprovalForce)
	fmt.Fprintf(os.Stderr, "Opening browser for authentication...\n")
	fmt.Fprintf(os.Stderr, "If browser doesn't open, visit:\n%v\n\n", authURL)

	var cmd *exec.Cmd
	switch runtime.GOOS {
	case "darwin":
		cmd = exec.Command("open", authURL)
	case "linux":
		cmd = exec.Command("xdg-open", authURL)
	case "windows":
		cmd = exec.Command("rundll32", "url.dll,FileProtocolHandler", authURL)
	}

	if cmd != nil {
		_ = cmd.Start()
	}

	var code string
	select {
	case code = <-codeChan:
	case err := <-errChan:
		return nil, err
	case <-time.After(3 * time.Minute):
		return nil, fmt.Errorf("authentication timeout after 3 minutes")
	}

	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	_ = server.Shutdown(ctx)

	tok, err := config.Exchange(context.Background(), code)
	if err != nil {
		return nil, fmt.Errorf("unable to retrieve token from web: %w", err)
	}

	fmt.Fprintln(os.Stderr, "\nAuthentication successful!")
	return tok, nil
}

func saveToken(path string, token *oauth2.Token) error {
	fmt.Fprintf(os.Stderr, "Saving credentials to: %s\n", path)

	if err := os.MkdirAll(filepath.Dir(path), 0700); err != nil {
		return err
	}

	file, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		return err
	}
	defer file.Close()

	return json.NewEncoder(file).Encode(token)
}
