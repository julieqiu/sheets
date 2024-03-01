package sheets

import (
	"context"
	"fmt"
	"regexp"
	"strings"
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

func (s *Spreadsheet) GetValues(ctx context.Context, readRange string) ([][]interface{}, error) {
	resp, err := s.service.Spreadsheets.Values.Get(s.id, readRange).Do()
	if err != nil {
		return nil, fmt.Errorf("unable to retrieve data from sheet: %v", err)
	}
	return resp.Values, nil
}
