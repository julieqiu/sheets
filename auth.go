package sheets

import (
	"context"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"os"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/sheets/v4"
)

func GoogleSheetsService(ctx context.Context, credentialsFile, tokenFile string) (*sheets.Service, error) {
	// Read the user's credentials file.
	b, err := ioutil.ReadFile(credentialsFile)
	if err != nil {
		return nil, err
	}
	// If modifying these scopes, delete your previously saved token.json.
	config, err := google.ConfigFromJSON(b, "https://www.googleapis.com/auth/spreadsheets")
	if err != nil {
		return nil, err
	}
	tok, err := getOauthToken(ctx, tokenFile, config)
	if err != nil {
		return nil, err
	}
	return sheets.New(config.Client(ctx, tok))
}

func getOauthToken(ctx context.Context, tokenFile string, config *oauth2.Config) (*oauth2.Token, error) {
	// token.json stores the user's access and refresh tokens, and is created
	// automatically when the authorization flow completes for the first time.
	f, err := os.Open(tokenFile)
	if err == nil {
		defer f.Close()
		tok := &oauth2.Token{}
		if err := json.NewDecoder(f).Decode(tok); err != nil {
			return nil, err
		}
		return tok, nil
	}
	if !os.IsNotExist(err) {
		return nil, err
	}
	// If the token file isn't available, create one.
	// Request a token from the web, then returns the retrieved token.
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	log.Printf("Go to the following link in your browser then type the "+
		"authorization code: \n%v\n", authURL)

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		return nil, err
	}
	tok, err := config.Exchange(ctx, authCode)
	if err != nil {
		return nil, err
	}
	// Save the token for future use.
	log.Printf("Saving credential file to: %s\n", tokenFile)
	f, err = os.OpenFile(tokenFile, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	if err := json.NewEncoder(f).Encode(tok); err != nil {
		return nil, err
	}
	return tok, nil
}
