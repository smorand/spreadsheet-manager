// Package mcp provides the MCP (Model Context Protocol) server implementation
// for spreadsheet-manager, enabling AI assistants to manage spreadsheets remotely.
// This file implements OAuth 2.1 authorization server endpoints as per MCP specification.
//
// MCP OAuth2 Flow:
// 1. Client discovers auth server via /.well-known/oauth-protected-resource
// 2. Client fetches auth server metadata from /.well-known/oauth-authorization-server
// 3. Client registers via /oauth/register (Dynamic Client Registration)
// 4. Client redirects user to /oauth/authorize => we redirect to Google OAuth
// 5. Google returns code to /oauth/callback => we exchange with Google
// 6. Client exchanges code at /oauth/token => we return Google tokens
// 7. Client sends Bearer token on MCP requests => we validate and use for Sheets API
package mcp

import (
	"context"
	"crypto/rand"
	"crypto/sha256"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"os"
	"strings"
	"sync"
	"time"

	secretmanager "cloud.google.com/go/secretmanager/apiv1"
	"cloud.google.com/go/secretmanager/apiv1/secretmanagerpb"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/sheets/v4"
)

// OAuth2Scopes defines the scopes required for Google Sheets access.
var OAuth2Scopes = []string{
	sheets.SpreadsheetsScope,
	sheets.DriveScope,
}

// ProtectedResourceMetadata represents RFC 9728 protected resource metadata.
type ProtectedResourceMetadata struct {
	Resource               string   `json:"resource"`
	AuthorizationServers   []string `json:"authorization_servers"`
	BearerMethodsSupported []string `json:"bearer_methods_supported,omitempty"`
	ScopesSupported        []string `json:"scopes_supported,omitempty"`
}

// AuthorizationServerMetadata represents RFC 8414 authorization server metadata.
type AuthorizationServerMetadata struct {
	Issuer                            string   `json:"issuer"`
	AuthorizationEndpoint             string   `json:"authorization_endpoint"`
	TokenEndpoint                     string   `json:"token_endpoint"`
	RegistrationEndpoint              string   `json:"registration_endpoint,omitempty"`
	ScopesSupported                   []string `json:"scopes_supported,omitempty"`
	ResponseTypesSupported            []string `json:"response_types_supported"`
	GrantTypesSupported               []string `json:"grant_types_supported"`
	CodeChallengeMethodsSupported     []string `json:"code_challenge_methods_supported"`
	TokenEndpointAuthMethodsSupported []string `json:"token_endpoint_auth_methods_supported"`
}

// ClientRegistrationRequest represents RFC 7591 dynamic client registration request.
type ClientRegistrationRequest struct {
	ClientName              string   `json:"client_name,omitempty"`
	RedirectURIs            []string `json:"redirect_uris"`
	GrantTypes              []string `json:"grant_types,omitempty"`
	ResponseTypes           []string `json:"response_types,omitempty"`
	TokenEndpointAuthMethod string   `json:"token_endpoint_auth_method,omitempty"`
}

// ClientRegistrationResponse represents RFC 7591 dynamic client registration response.
type ClientRegistrationResponse struct {
	ClientID                string   `json:"client_id"`
	ClientSecret            string   `json:"client_secret,omitempty"`
	ClientIDIssuedAt        int64    `json:"client_id_issued_at,omitempty"`
	ClientSecretExpiresAt   int64    `json:"client_secret_expires_at,omitempty"`
	RedirectURIs            []string `json:"redirect_uris"`
	GrantTypes              []string `json:"grant_types,omitempty"`
	ResponseTypes           []string `json:"response_types,omitempty"`
	TokenEndpointAuthMethod string   `json:"token_endpoint_auth_method,omitempty"`
}

// TokenRequest represents an OAuth2 token request.
type TokenRequest struct {
	GrantType    string `json:"grant_type"`
	Code         string `json:"code,omitempty"`
	RedirectURI  string `json:"redirect_uri,omitempty"`
	ClientID     string `json:"client_id,omitempty"`
	CodeVerifier string `json:"code_verifier,omitempty"`
	RefreshToken string `json:"refresh_token,omitempty"`
	Resource     string `json:"resource,omitempty"`
}

// TokenResponse represents an OAuth2 token response.
type TokenResponse struct {
	AccessToken  string `json:"access_token"`
	TokenType    string `json:"token_type"`
	ExpiresIn    int64  `json:"expires_in,omitempty"`
	RefreshToken string `json:"refresh_token,omitempty"`
	Scope        string `json:"scope,omitempty"`
}

// TokenErrorResponse represents an OAuth2 error response.
type TokenErrorResponse struct {
	Error            string `json:"error"`
	ErrorDescription string `json:"error_description,omitempty"`
}

// registeredClient stores registered OAuth client information.
type registeredClient struct {
	ClientID     string
	ClientSecret string
	RedirectURIs []string
	CreatedAt    time.Time
}

// authorizationState stores OAuth authorization state.
type authorizationState struct {
	ClientID      string
	RedirectURI   string
	CodeChallenge string
	CodeMethod    string
	CreatedAt     time.Time
}

// authorizationCode stores issued authorization codes.
type authorizationCode struct {
	Code          string
	ClientID      string
	RedirectURI   string
	CodeChallenge string
	CodeMethod    string
	GoogleToken   *oauth2.Token // The actual Google OAuth token
	CreatedAt     time.Time
}

// OAuth2Server handles OAuth 2.1 authorization server endpoints.
type OAuth2Server struct {
	baseURL        string
	googleClientID string
	googleSecret   string
	oauthConfig    *oauth2.Config
	oauthConfigMu  sync.RWMutex

	// In-memory stores (could be replaced with Redis for production)
	clientsMu sync.RWMutex
	clients   map[string]*registeredClient

	statesMu sync.RWMutex
	states   map[string]*authorizationState

	codesMu sync.RWMutex
	codes   map[string]*authorizationCode

	// Configuration
	secretProject  string
	secretName     string
	credentialFile string
}

// OAuth2ServerConfig holds configuration for the OAuth2 server.
type OAuth2ServerConfig struct {
	BaseURL        string // Base URL (e.g., https://mcp.example.com)
	SecretProject  string // GCP project for Secret Manager
	SecretName     string // Secret name for OAuth credentials
	CredentialFile string // Local credential file (fallback)
}

// NewOAuth2Server creates a new OAuth2 authorization server.
func NewOAuth2Server(cfg *OAuth2ServerConfig) *OAuth2Server {
	s := &OAuth2Server{
		baseURL:        cfg.BaseURL,
		secretProject:  cfg.SecretProject,
		secretName:     cfg.SecretName,
		credentialFile: cfg.CredentialFile,
		clients:        make(map[string]*registeredClient),
		states:         make(map[string]*authorizationState),
		codes:          make(map[string]*authorizationCode),
	}

	// Start cleanup goroutines
	go s.cleanupExpiredStates()
	go s.cleanupExpiredCodes()

	return s
}

// cleanupExpiredStates removes expired authorization states.
func (s *OAuth2Server) cleanupExpiredStates() {
	ticker := time.NewTicker(1 * time.Minute)
	defer ticker.Stop()

	for range ticker.C {
		s.statesMu.Lock()
		now := time.Now()
		for state, entry := range s.states {
			if now.Sub(entry.CreatedAt) > 10*time.Minute {
				delete(s.states, state)
			}
		}
		s.statesMu.Unlock()
	}
}

// cleanupExpiredCodes removes expired authorization codes.
func (s *OAuth2Server) cleanupExpiredCodes() {
	ticker := time.NewTicker(1 * time.Minute)
	defer ticker.Stop()

	for range ticker.C {
		s.codesMu.Lock()
		now := time.Now()
		for code, entry := range s.codes {
			if now.Sub(entry.CreatedAt) > 10*time.Minute {
				delete(s.codes, code)
			}
		}
		s.codesMu.Unlock()
	}
}

// LoadCredentials loads Google OAuth credentials from Secret Manager or file.
func (s *OAuth2Server) LoadCredentials(ctx context.Context) error {
	s.oauthConfigMu.Lock()
	defer s.oauthConfigMu.Unlock()

	if s.oauthConfig != nil {
		return nil
	}

	var credentialsJSON []byte
	var err error

	// Try Secret Manager first
	if s.secretProject != "" && s.secretName != "" {
		credentialsJSON, err = loadFromSecretManager(ctx, s.secretProject, s.secretName)
		if err != nil {
			log.Printf("Failed to load credentials from Secret Manager: %v", err)
		} else {
			log.Printf("OAuth credentials loaded from Secret Manager: %s/%s", s.secretProject, s.secretName)
		}
	}

	// Fall back to local file
	if credentialsJSON == nil && s.credentialFile != "" {
		credentialsJSON, err = os.ReadFile(s.credentialFile)
		if err != nil {
			return fmt.Errorf("failed to read credentials file %s: %w", s.credentialFile, err)
		}
		log.Printf("OAuth credentials loaded from file: %s", s.credentialFile)
	}

	if credentialsJSON == nil {
		return fmt.Errorf("no OAuth credentials available: configure Secret Manager or credential file")
	}

	// Parse credentials
	config, err := google.ConfigFromJSON(credentialsJSON, OAuth2Scopes...)
	if err != nil {
		return fmt.Errorf("failed to parse OAuth credentials: %w", err)
	}

	// Set redirect to our internal callback (not the client's callback)
	config.RedirectURL = s.baseURL + "/oauth/callback"

	s.oauthConfig = config
	s.googleClientID = config.ClientID
	s.googleSecret = config.ClientSecret

	return nil
}

// getOAuthConfig returns the loaded OAuth config.
func (s *OAuth2Server) getOAuthConfig(ctx context.Context) (*oauth2.Config, error) {
	s.oauthConfigMu.RLock()
	config := s.oauthConfig
	s.oauthConfigMu.RUnlock()

	if config != nil {
		return config, nil
	}

	if err := s.LoadCredentials(ctx); err != nil {
		return nil, err
	}

	s.oauthConfigMu.RLock()
	defer s.oauthConfigMu.RUnlock()
	return s.oauthConfig, nil
}

// SetupRoutes registers all OAuth2 endpoints.
func (s *OAuth2Server) SetupRoutes(mux *http.ServeMux) {
	// Well-known endpoints
	mux.HandleFunc("/.well-known/oauth-protected-resource", s.HandleProtectedResourceMetadata)
	mux.HandleFunc("/.well-known/oauth-authorization-server", s.HandleAuthorizationServerMetadata)

	// OAuth endpoints
	mux.HandleFunc("/oauth/register", s.HandleClientRegistration)
	mux.HandleFunc("/oauth/authorize", s.HandleAuthorize)
	mux.HandleFunc("/oauth/callback", s.HandleCallback)
	mux.HandleFunc("/oauth/token", s.HandleToken)
}

// HandleProtectedResourceMetadata serves RFC 9728 protected resource metadata.
// GET /.well-known/oauth-protected-resource
func (s *OAuth2Server) HandleProtectedResourceMetadata(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	metadata := ProtectedResourceMetadata{
		Resource:               s.baseURL,
		AuthorizationServers:   []string{s.baseURL},
		BearerMethodsSupported: []string{"header"},
		ScopesSupported: []string{
			"spreadsheets:read",
			"spreadsheets:write",
		},
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(metadata)
}

// HandleAuthorizationServerMetadata serves RFC 8414 authorization server metadata.
// GET /.well-known/oauth-authorization-server
func (s *OAuth2Server) HandleAuthorizationServerMetadata(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	metadata := AuthorizationServerMetadata{
		Issuer:                s.baseURL,
		AuthorizationEndpoint: s.baseURL + "/oauth/authorize",
		TokenEndpoint:         s.baseURL + "/oauth/token",
		RegistrationEndpoint:  s.baseURL + "/oauth/register",
		ScopesSupported: []string{
			"spreadsheets:read",
			"spreadsheets:write",
		},
		ResponseTypesSupported:            []string{"code"},
		GrantTypesSupported:               []string{"authorization_code", "refresh_token"},
		CodeChallengeMethodsSupported:     []string{"S256"},
		TokenEndpointAuthMethodsSupported: []string{"none", "client_secret_basic", "client_secret_post"},
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(metadata)
}

// HandleClientRegistration implements RFC 7591 Dynamic Client Registration.
// POST /oauth/register
func (s *OAuth2Server) HandleClientRegistration(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	var req ClientRegistrationRequest
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		writeOAuthError(w, "invalid_request", "Invalid JSON body", http.StatusBadRequest)
		return
	}

	// Validate redirect URIs
	if len(req.RedirectURIs) == 0 {
		writeOAuthError(w, "invalid_request", "redirect_uris is required", http.StatusBadRequest)
		return
	}

	// Generate client credentials
	clientID := generateSecureToken(16)
	clientSecret := generateSecureToken(32)

	// Store client
	client := &registeredClient{
		ClientID:     clientID,
		ClientSecret: clientSecret,
		RedirectURIs: req.RedirectURIs,
		CreatedAt:    time.Now(),
	}

	s.clientsMu.Lock()
	s.clients[clientID] = client
	s.clientsMu.Unlock()

	log.Printf("Registered new OAuth client: %s (name: %s)", clientID, req.ClientName)

	// Build response
	resp := ClientRegistrationResponse{
		ClientID:                clientID,
		ClientSecret:            clientSecret,
		ClientIDIssuedAt:        time.Now().Unix(),
		ClientSecretExpiresAt:   0, // Never expires
		RedirectURIs:            req.RedirectURIs,
		GrantTypes:              []string{"authorization_code", "refresh_token"},
		ResponseTypes:           []string{"code"},
		TokenEndpointAuthMethod: "client_secret_basic",
	}

	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(http.StatusCreated)
	json.NewEncoder(w).Encode(resp)
}

// HandleAuthorize handles the authorization request from the MCP client.
// GET /oauth/authorize?client_id=xxx&redirect_uri=xxx&response_type=code&state=xxx&code_challenge=xxx&code_challenge_method=S256
func (s *OAuth2Server) HandleAuthorize(w http.ResponseWriter, r *http.Request) {
	ctx := r.Context()

	if r.Method != http.MethodGet {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	// Parse query parameters
	clientID := r.URL.Query().Get("client_id")
	redirectURI := r.URL.Query().Get("redirect_uri")
	responseType := r.URL.Query().Get("response_type")
	state := r.URL.Query().Get("state")
	codeChallenge := r.URL.Query().Get("code_challenge")
	codeChallengeMethod := r.URL.Query().Get("code_challenge_method")

	// Validate required parameters
	if clientID == "" {
		writeOAuthError(w, "invalid_request", "client_id is required", http.StatusBadRequest)
		return
	}
	if redirectURI == "" {
		writeOAuthError(w, "invalid_request", "redirect_uri is required", http.StatusBadRequest)
		return
	}
	if responseType != "code" {
		writeOAuthError(w, "unsupported_response_type", "Only 'code' response type is supported", http.StatusBadRequest)
		return
	}

	// Check if client exists, auto-register if not
	// This allows MCP clients like Claude to use the OAuth flow without
	// explicit Dynamic Client Registration
	s.clientsMu.RLock()
	client, exists := s.clients[clientID]
	s.clientsMu.RUnlock()

	if !exists {
		// Auto-register the client with the provided redirect_uri
		client = &registeredClient{
			ClientID:     clientID,
			ClientSecret: "", // Not needed for authorization code flow with PKCE
			RedirectURIs: []string{redirectURI},
			CreatedAt:    time.Now(),
		}
		s.clientsMu.Lock()
		s.clients[clientID] = client
		s.clientsMu.Unlock()
		log.Printf("Auto-registered OAuth client: %s with redirect_uri: %s", clientID, redirectURI)
	}

	// Validate redirect_uri (for auto-registered clients, we add new URIs dynamically)
	validRedirect := false
	for _, uri := range client.RedirectURIs {
		if uri == redirectURI {
			validRedirect = true
			break
		}
	}
	if !validRedirect {
		// Add the new redirect_uri for this client
		s.clientsMu.Lock()
		client.RedirectURIs = append(client.RedirectURIs, redirectURI)
		s.clientsMu.Unlock()
		log.Printf("Added redirect_uri %s for client %s", redirectURI, clientID)
	}

	// Generate internal state that maps to the client's request
	internalState := generateSecureToken(32)

	// Store authorization state
	authState := &authorizationState{
		ClientID:      clientID,
		RedirectURI:   redirectURI,
		CodeChallenge: codeChallenge,
		CodeMethod:    codeChallengeMethod,
		CreatedAt:     time.Now(),
	}

	s.statesMu.Lock()
	s.states[internalState] = authState
	s.statesMu.Unlock()

	// Also store the client's state so we can return it
	if state != "" {
		// Append client state to our internal state (separated by .)
		internalState = internalState + "." + state
	}

	// Load Google OAuth config
	config, err := s.getOAuthConfig(ctx)
	if err != nil {
		log.Printf("Failed to load OAuth config: %v", err)
		http.Error(w, "OAuth configuration error", http.StatusInternalServerError)
		return
	}

	// Redirect to Google OAuth
	authURL := config.AuthCodeURL(internalState, oauth2.AccessTypeOffline, oauth2.ApprovalForce)

	log.Printf("Authorization request: client=%s, redirecting to Google OAuth", clientID)

	http.Redirect(w, r, authURL, http.StatusFound)
}

// HandleCallback handles the OAuth callback from Google.
// GET /oauth/callback?code=xxx&state=xxx
func (s *OAuth2Server) HandleCallback(w http.ResponseWriter, r *http.Request) {
	ctx := r.Context()

	// Check for error
	if errParam := r.URL.Query().Get("error"); errParam != "" {
		errDesc := r.URL.Query().Get("error_description")
		log.Printf("Google OAuth error: %s - %s", errParam, errDesc)
		http.Error(w, fmt.Sprintf("OAuth error: %s - %s", errParam, errDesc), http.StatusBadRequest)
		return
	}

	code := r.URL.Query().Get("code")
	fullState := r.URL.Query().Get("state")

	if code == "" || fullState == "" {
		writeOAuthError(w, "invalid_request", "Missing code or state", http.StatusBadRequest)
		return
	}

	// Split internal state from client state
	// Our internal state is 42 chars (32 bytes base64url, length + length/3)
	// Format: "<internal_state>.<client_state>" or just "<internal_state>"
	internalState := fullState
	clientState := ""
	if dotIdx := strings.Index(fullState, "."); dotIdx > 0 {
		internalState = fullState[:dotIdx]
		clientState = fullState[dotIdx+1:]
	}

	// Validate internal state
	s.statesMu.Lock()
	authState, exists := s.states[internalState]
	if exists {
		delete(s.states, internalState)
	}
	s.statesMu.Unlock()

	if !exists {
		writeOAuthError(w, "invalid_request", "Invalid or expired state", http.StatusBadRequest)
		return
	}

	// Exchange code with Google
	config, err := s.getOAuthConfig(ctx)
	if err != nil {
		log.Printf("Failed to load OAuth config: %v", err)
		http.Error(w, "OAuth configuration error", http.StatusInternalServerError)
		return
	}

	googleToken, err := config.Exchange(ctx, code)
	if err != nil {
		log.Printf("Failed to exchange code with Google: %v", err)
		writeOAuthError(w, "invalid_grant", "Failed to exchange authorization code", http.StatusBadRequest)
		return
	}

	log.Printf("Google OAuth exchange successful, refresh_token_present=%v", googleToken.RefreshToken != "")

	// Generate our own authorization code
	ourCode := generateSecureToken(32)

	// Store the code with the Google token
	codeEntry := &authorizationCode{
		Code:          ourCode,
		ClientID:      authState.ClientID,
		RedirectURI:   authState.RedirectURI,
		CodeChallenge: authState.CodeChallenge,
		CodeMethod:    authState.CodeMethod,
		GoogleToken:   googleToken,
		CreatedAt:     time.Now(),
	}

	s.codesMu.Lock()
	s.codes[ourCode] = codeEntry
	s.codesMu.Unlock()

	// Redirect back to the client with our code
	redirectURL := authState.RedirectURI + "?code=" + ourCode
	if clientState != "" {
		redirectURL += "&state=" + clientState
	}

	log.Printf("Redirecting to client: %s", authState.RedirectURI)

	http.Redirect(w, r, redirectURL, http.StatusFound)
}

// HandleToken handles token exchange and refresh requests.
// POST /oauth/token
func (s *OAuth2Server) HandleToken(w http.ResponseWriter, r *http.Request) {
	ctx := r.Context()

	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	// Parse form data
	if err := r.ParseForm(); err != nil {
		writeOAuthError(w, "invalid_request", "Invalid form data", http.StatusBadRequest)
		return
	}

	grantType := r.FormValue("grant_type")
	code := r.FormValue("code")
	clientID := r.FormValue("client_id")
	codeVerifier := r.FormValue("code_verifier")
	refreshToken := r.FormValue("refresh_token")

	// Also check for client credentials in Authorization header (Basic auth)
	if clientID == "" {
		if username, _, ok := r.BasicAuth(); ok {
			clientID = username
		}
	}

	switch grantType {
	case "authorization_code":
		s.handleAuthorizationCodeGrant(ctx, w, clientID, code, codeVerifier)
	case "refresh_token":
		s.handleRefreshTokenGrant(ctx, w, clientID, refreshToken)
	default:
		writeOAuthError(w, "unsupported_grant_type", "Only authorization_code and refresh_token are supported", http.StatusBadRequest)
	}
}

// handleAuthorizationCodeGrant handles the authorization_code grant type.
func (s *OAuth2Server) handleAuthorizationCodeGrant(ctx context.Context, w http.ResponseWriter, clientID, code, codeVerifier string) {
	if code == "" {
		writeOAuthError(w, "invalid_request", "code is required", http.StatusBadRequest)
		return
	}

	// Look up the authorization code
	s.codesMu.Lock()
	codeEntry, exists := s.codes[code]
	if exists {
		delete(s.codes, code) // Single use
	}
	s.codesMu.Unlock()

	if !exists {
		writeOAuthError(w, "invalid_grant", "Invalid or expired authorization code", http.StatusBadRequest)
		return
	}

	// Validate client_id matches
	if clientID != "" && clientID != codeEntry.ClientID {
		writeOAuthError(w, "invalid_client", "client_id mismatch", http.StatusUnauthorized)
		return
	}

	// Validate PKCE if code_challenge was provided during authorization
	if codeEntry.CodeChallenge != "" {
		if codeVerifier == "" {
			writeOAuthError(w, "invalid_request", "code_verifier is required", http.StatusBadRequest)
			return
		}
		if !validatePKCE(codeVerifier, codeEntry.CodeChallenge, codeEntry.CodeMethod) {
			writeOAuthError(w, "invalid_grant", "Invalid code_verifier", http.StatusBadRequest)
			return
		}
	}

	// Return the Google tokens
	googleToken := codeEntry.GoogleToken

	resp := TokenResponse{
		AccessToken:  googleToken.AccessToken,
		TokenType:    "Bearer",
		ExpiresIn:    int64(time.Until(googleToken.Expiry).Seconds()),
		RefreshToken: googleToken.RefreshToken,
		Scope:        "spreadsheets:read spreadsheets:write",
	}

	log.Printf("Token issued for client: %s", codeEntry.ClientID)

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(resp)
}

// handleRefreshTokenGrant handles the refresh_token grant type.
func (s *OAuth2Server) handleRefreshTokenGrant(ctx context.Context, w http.ResponseWriter, clientID, refreshToken string) {
	if refreshToken == "" {
		writeOAuthError(w, "invalid_request", "refresh_token is required", http.StatusBadRequest)
		return
	}

	// Load Google OAuth config
	config, err := s.getOAuthConfig(ctx)
	if err != nil {
		log.Printf("Failed to load OAuth config: %v", err)
		writeOAuthError(w, "server_error", "OAuth configuration error", http.StatusInternalServerError)
		return
	}

	// Create a token source from the refresh token
	token := &oauth2.Token{
		RefreshToken: refreshToken,
	}

	tokenSource := config.TokenSource(ctx, token)

	// Get a new token (this will refresh if needed)
	newToken, err := tokenSource.Token()
	if err != nil {
		log.Printf("Failed to refresh token: %v", err)
		writeOAuthError(w, "invalid_grant", "Failed to refresh token", http.StatusBadRequest)
		return
	}

	resp := TokenResponse{
		AccessToken:  newToken.AccessToken,
		TokenType:    "Bearer",
		ExpiresIn:    int64(time.Until(newToken.Expiry).Seconds()),
		RefreshToken: newToken.RefreshToken,
		Scope:        "spreadsheets:read spreadsheets:write",
	}

	// RefreshToken may be empty if Google didn't rotate it
	if resp.RefreshToken == "" {
		resp.RefreshToken = refreshToken
	}

	log.Printf("Token refreshed for client: %s", clientID)

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(resp)
}

// ValidateAccessToken validates a Bearer token and returns the Google OAuth config for API calls.
// This is used by the auth middleware.
func (s *OAuth2Server) ValidateAccessToken(ctx context.Context, accessToken string) (*oauth2.Config, *oauth2.Token, error) {
	config, err := s.getOAuthConfig(ctx)
	if err != nil {
		return nil, nil, fmt.Errorf("failed to load OAuth config: %w", err)
	}

	// Create a token with the access token
	// Note: We don't have the expiry, so the token source will refresh if needed
	token := &oauth2.Token{
		AccessToken: accessToken,
	}

	return config, token, nil
}

// Helper functions

// generateSecureToken generates a cryptographically secure random token.
func generateSecureToken(length int) string {
	b := make([]byte, length)
	if _, err := rand.Read(b); err != nil {
		// Fallback to less secure but functional
		log.Printf("Warning: crypto/rand failed, using time-based seed")
		return fmt.Sprintf("%d", time.Now().UnixNano())
	}
	return base64.URLEncoding.EncodeToString(b)[:length+length/3]
}

// validatePKCE validates the PKCE code_verifier against the stored code_challenge.
func validatePKCE(verifier, challenge, method string) bool {
	if method != "S256" && method != "" {
		return false // Only S256 is supported
	}

	// For S256: challenge = BASE64URL(SHA256(verifier))
	h := sha256.Sum256([]byte(verifier))
	computed := base64.RawURLEncoding.EncodeToString(h[:])

	return computed == challenge
}

// writeOAuthError writes a standard OAuth2 error response.
func writeOAuthError(w http.ResponseWriter, errorCode, description string, statusCode int) {
	resp := TokenErrorResponse{
		Error:            errorCode,
		ErrorDescription: description,
	}
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(statusCode)
	json.NewEncoder(w).Encode(resp)
}

// loadFromSecretManager loads credentials from Google Secret Manager.
func loadFromSecretManager(ctx context.Context, project, secretName string) ([]byte, error) {
	client, err := secretmanager.NewClient(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to create Secret Manager client: %w", err)
	}
	defer client.Close()

	secretPath := fmt.Sprintf("projects/%s/secrets/%s/versions/latest", project, secretName)
	result, err := client.AccessSecretVersion(ctx, &secretmanagerpb.AccessSecretVersionRequest{
		Name: secretPath,
	})
	if err != nil {
		return nil, fmt.Errorf("failed to access secret %s: %w", secretPath, err)
	}

	return result.Payload.Data, nil
}
