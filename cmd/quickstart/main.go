package main

import (
	"context"
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/julieqiu/sheets"
)

const (
	exampleSheetsURL = "https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit"
	exampleReadRange = "Class Data!A2:E"
)

var (
	credentialsFile = flag.String("credentials", os.Getenv("GOOGLE_SHEETS_CREDENTIALS"), "path to credentials file for Google Sheets")
	tokenFile       = flag.String("token", os.Getenv("GOOGLE_SHEETS_TOKEN"), "path to token file for authentication in Google sheets")
	sheetsURL       = flag.String("url", exampleSheetsURL, "url of Google sheet to be read")
	readRange       = flag.String("range", exampleReadRange, "range of sheet to be read")
)

func main() {
	ctx := context.Background()
	flag.Parse()

	sheetsID, err := sheets.GetSpreadsheetID(*sheetsURL)
	if err != nil {
		log.Fatal(err)
	}
	s, err := sheets.Open(ctx, *credentialsFile, *tokenFile, sheetsID)
	if err != nil {
		log.Fatal(err)
	}
	values, err := s.GetValues(ctx, *readRange)
	if err != nil {
		log.Fatal(err)
	}
	for _, v := range values {
		fmt.Println(v)
	}
}
