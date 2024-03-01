package sheets

import (
	"context"

	"google.golang.org/api/sheets/v4"
)

// Spreadsheet represents a Google Spreadsheet, which can contain multiple
// sheets, each with structured information contained in cells.  A Spreadsheet
// has a unique spreadsheetID value, which can be found in a Google Sheets URL.
// See https://developers.google.com/sheets/api/guides/concepts for more
// information.
type Spreadsheet struct {
	id          string
	spreadsheet *sheets.Spreadsheet
	service     *sheets.Service
}

// Open opens an exist Spreadsheet.
func Open(ctx context.Context, credentialsFile, tokenFile, id string) (*Spreadsheet, error) {
	s, err := newSpreadsheet(ctx, credentialsFile, tokenFile)
	if err != nil {
		return nil, err
	}
	s.id = id
	return s, nil
}

// Create creates a blank spreadsheet.
func Create(ctx context.Context, credentialsFile, tokenFile, title string) (*Spreadsheet, error) {
	s, err := newSpreadsheet(ctx, credentialsFile, tokenFile)
	if err != nil {
		return nil, err
	}
	rowData := make(map[string][]*sheets.RowData)
	sheet, err := createSheet(ctx, s.service, title, rowData)
	if err != nil {
		return nil, err
	}
	s.id = sheet.SpreadsheetId
	s.spreadsheet = sheet
	return s, nil
}

func newSpreadsheet(ctx context.Context, credentialsFile, tokenFile string) (*Spreadsheet, error) {
	srv, err := GoogleSheetsService(ctx, credentialsFile, tokenFile)
	if err != nil {
		return nil, err
	}
	return &Spreadsheet{
		service: srv,
	}, nil
}

func createSheet(ctx context.Context, srv *sheets.Service, title string, rowData map[string][]*sheets.RowData) (*sheets.Spreadsheet, error) {
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
	sheet, err := srv.Spreadsheets.Create(spreadsheet).Context(ctx).Do()
	if err != nil {
		return nil, err
	}
	return sheet, nil
}
