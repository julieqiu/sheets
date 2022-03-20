package sheets

import (
	"context"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strings"

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

func Write(_ context.Context, outputDir string, data map[string][]*Row, rowData map[string][]*sheets.RowData) error {
	// Write output to disk first.
	var filenames []string
	for filename, cells := range data {
		if len(cells) == 0 {
			continue
		}
		fullpath := filepath.Join(outputDir, fmt.Sprintf("%s.csv", filename))
		file, err := os.Create(fullpath)
		if err != nil {
			return err
		}
		defer file.Close()

		writer := csv.NewWriter(file)
		defer writer.Flush()

		for _, row := range cells {
			if err := writer.Write(row.ToCells()); err != nil {
				return err
			}
		}
		filenames = append(filenames, fullpath)
	}
	for _, filename := range filenames {
		log.Printf("Wrote output to %s.\n", filename)
	}
	// Add a new sheet and write output to it.
	for title, cells := range data {
		if len(cells) == 0 {
			continue
		}
		var rd []*sheets.RowData
		for _, row := range cells {
			var values []*sheets.CellData
			for _, cell := range row.Cells {
				cd := &sheets.CellData{
					UserEnteredFormat: &sheets.CellFormat{
						TextFormat: &sheets.TextFormat{
							Bold: row.BoldText,
						},
					},
				}
				if row.Color != nil {
					r, g, b, _ := row.Color.RGBA()
					cd.UserEnteredFormat.BackgroundColor = &sheets.Color{
						Blue:  float64(b) / 255.0,
						Green: float64(g) / 255.0,
						Red:   float64(r) / 255.0,
					}
				}
				if cell.Hyperlink != "" {
					cd.UserEnteredValue = &sheets.ExtendedValue{
						FormulaValue: newStrPtr(cell.HyperlinkFormula()),
					}
				} else {
					cd.UserEnteredValue = &sheets.ExtendedValue{
						StringValue: newStrPtr(cell.Text),
					}
				}
				values = append(values, cd)
			}
			rd = append(rd, &sheets.RowData{
				Values: values,
			})
		}
		rowData[title] = rd
	}
	return nil
}

func newStrPtr(text string) *string {
	strValuePtr := new(string)
	*strValuePtr = text
	return strValuePtr
}

func CreateSheet(ctx context.Context, srv *sheets.Service, title string, rowData map[string][]*sheets.RowData) (*sheets.Spreadsheet, error) {
	var newSheets []*sheets.Sheet
	for title, data := range rowData {
		newSheets = append(newSheets, &sheets.Sheet{
			Properties: &sheets.SheetProperties{
				Title: title,
				GridProperties: &sheets.GridProperties{
					FrozenRowCount: 1,
				},
			},
			Data: []*sheets.GridData{{RowData: data}},
		})
	}
	spreadsheet := &sheets.Spreadsheet{
		Properties: &sheets.SpreadsheetProperties{
			Title: title,
		},
		Sheets: newSheets,
	}
	return srv.Spreadsheets.Create(spreadsheet).Context(ctx).Do()
}

func AppendToSheet(ctx context.Context, srv *sheets.Service, spreadsheetID string, rowData map[string][]*sheets.RowData) (*sheets.Spreadsheet, error) {
	// First, create the new sheets in spreadsheet.
	var createRequests []*sheets.Request
	for title := range rowData {
		createRequests = append(createRequests, &sheets.Request{
			AddSheet: &sheets.AddSheetRequest{
				Properties: &sheets.SheetProperties{
					Title: title,
					GridProperties: &sheets.GridProperties{
						FrozenRowCount: 1,
					},
				},
			},
		})
	}
	response, err := srv.Spreadsheets.BatchUpdate(spreadsheetID, &sheets.BatchUpdateSpreadsheetRequest{
		IncludeSpreadsheetInResponse: true,
		Requests:                     createRequests,
	}).Context(ctx).Do()
	if err != nil {
		return nil, err
	}
	// Now, add the data to the spreadsheets.
	var dataRequests []*sheets.Request
	for _, sheet := range response.UpdatedSpreadsheet.Sheets {
		dataRequests = append(dataRequests, &sheets.Request{
			AppendCells: &sheets.AppendCellsRequest{
				SheetId: sheet.Properties.SheetId,
				Rows:    rowData[sheet.Properties.Title],
				Fields:  "*",
			},
		})
	}
	response, err = srv.Spreadsheets.BatchUpdate(spreadsheetID, &sheets.BatchUpdateSpreadsheetRequest{
		IncludeSpreadsheetInResponse: true,
		Requests:                     dataRequests,
	}).Context(ctx).Do()
	if err != nil {
		return nil, err
	}
	return response.UpdatedSpreadsheet, nil
}

// Return the Google Sheets spreadsheet ID for the given URL. If the URL is an
// invalid format, an error will be returned.
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
