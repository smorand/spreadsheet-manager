package main

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"os"
	"path/filepath"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

const (
	credentialsFile = "google_credentials.json"
	tokenFile       = "google_token.json"
)

var scopes = []string{
	sheets.SpreadsheetsScope,
	sheets.DriveFileScope,
}

// getCredentialsPath returns the path to credentials directory
func getCredentialsPath() string {
	home, err := os.UserHomeDir()
	if err != nil {
		return ""
	}
	return filepath.Join(home, ".credentials")
}

// getClient retrieves an OAuth2 HTTP client
func getClient(ctx context.Context) (*http.Client, error) {
	credPath := filepath.Join(getCredentialsPath(), credentialsFile)
	tokenPath := filepath.Join(getCredentialsPath(), tokenFile)

	// Read credentials file
	b, err := os.ReadFile(credPath)
	if err != nil {
		return nil, fmt.Errorf("unable to read credentials file %s: %w\n"+
			"See README.md for setup instructions", credPath, err)
	}

	// Parse credentials
	config, err := google.ConfigFromJSON(b, scopes...)
	if err != nil {
		return nil, fmt.Errorf("unable to parse credentials: %w", err)
	}

	// Try to load token from file
	token, err := tokenFromFile(tokenPath)
	if err != nil {
		// Get new token from user
		token, err = getTokenFromWeb(ctx, config)
		if err != nil {
			return nil, err
		}
		// Save token
		if err := saveToken(tokenPath, token); err != nil {
			fmt.Fprintf(os.Stderr, "Warning: unable to save token: %v\n", err)
		}
	}

	return config.Client(ctx, token), nil
}

// getSheetsService creates an authenticated Sheets service
func getSheetsService(ctx context.Context) (*sheets.Service, error) {
	client, err := getClient(ctx)
	if err != nil {
		return nil, err
	}

	service, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("unable to create Sheets service: %w", err)
	}

	return service, nil
}

// getTokenFromWeb requests a token from the web, then returns the retrieved token
func getTokenFromWeb(ctx context.Context, config *oauth2.Config) (*oauth2.Token, error) {
	// Channel to receive the authorization code
	codeChan := make(chan string, 1)
	errChan := make(chan error, 1)

	// Start local HTTP server to handle OAuth callback
	server := &http.Server{Addr: ":8080"}

	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		code := r.URL.Query().Get("code")
		if code == "" {
			errChan <- fmt.Errorf("no authorization code in callback")
			http.Error(w, "No authorization code received", http.StatusBadRequest)
			return
		}

		// Send success response to browser
		w.Header().Set("Content-Type", "text/html")
		fmt.Fprintf(w, `
			<html>
			<head><title>Authentication Successful</title></head>
			<body>
				<h1>Authentication successful!</h1>
				<p>You can close this window and return to the terminal.</p>
			</body>
			</html>
		`)

		// Send code to channel
		codeChan <- code
	})

	// Start server in background
	go func() {
		if err := server.ListenAndServe(); err != nil && err != http.ErrServerClosed {
			errChan <- fmt.Errorf("failed to start server: %w", err)
		}
	}()

	// Generate auth URL and prompt user
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser:\n%v\n\n", authURL)
	fmt.Println("Waiting for authentication...")

	// Wait for code or error
	var authCode string
	select {
	case authCode = <-codeChan:
		// Got the code
	case err := <-errChan:
		server.Shutdown(ctx)
		return nil, err
	case <-ctx.Done():
		server.Shutdown(ctx)
		return nil, fmt.Errorf("authentication cancelled")
	}

	// Shutdown server
	server.Shutdown(ctx)

	// Exchange code for token
	token, err := config.Exchange(ctx, authCode)
	if err != nil {
		return nil, fmt.Errorf("unable to retrieve token from web: %w", err)
	}

	return token, nil
}

// tokenFromFile retrieves a token from a local file
func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	token := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(token)
	return token, err
}

// saveToken saves a token to a file path
func saveToken(path string, token *oauth2.Token) error {
	fmt.Fprintf(os.Stderr, "Saving credentials to: %s\n", path)

	// Ensure directory exists
	if err := os.MkdirAll(filepath.Dir(path), 0700); err != nil {
		return err
	}

	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		return err
	}
	defer f.Close()

	return json.NewEncoder(f).Encode(token)
}
