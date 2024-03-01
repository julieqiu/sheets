package sheets

import (
	"context"
	"fmt"
	"regexp"
	"strings"

	"google.golang.org/api/sheets/v4"
)

// GetSpreadsheetID returns the Google Sheets spreadsheet ID for the given URL.
// If the URL is an invalid format, an error will be returned.
func GetSpreadsheetID(url string) (string, error) {
	var spreadsheetID string
	// Trim the extra pieces that the URL may contain.
	trimmed := strings.TrimPrefix(url, "https://docs.google.com")
	trimmed = strings.TrimSuffix(trimmed, "edit#gid=0")

	// Source: https://developers.google.com/sheets/api/guides/concepts.
	re, err := regexp.Compile("/spreadsheets/d/(?P<ID>([a-zA-Z0-9-_]+))")
	if err != nil {
		return "", err
	}
	match := re.FindStringSubmatch(trimmed)
	for i, name := range re.SubexpNames() {
		if name == "ID" {
			spreadsheetID = match[i]
		}
	}
	return spreadsheetID, nil
}

func GetValues(ctx context.Context, srv *sheets.Service, spreadsheetId, readRange string) error {
	resp, err := srv.Spreadsheets.Values.Get(spreadsheetId, readRange).Do()
	if err != nil {
		return fmt.Errorf("Unable to retrieve data from sheet: %v", err)
	}
	for _, row := range resp.Values {
		// Print columns A and E, which correspond to indices 0 and 4.
		fmt.Printf("%s, %s\n", row[0], row[4])
	}
	return nil
}
