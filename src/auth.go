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
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser:\n%v\n\n", authURL)
	fmt.Printf("Enter authorization code: ")

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		return nil, fmt.Errorf("unable to read authorization code: %w", err)
	}

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
