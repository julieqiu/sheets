package sheets

import (
	"context"

	"google.golang.org/api/sheets/v4"
)

type GoogleSheet struct {
	id          string
	spreadsheet *sheets.Spreadsheet
	service     *sheets.Service
}

func New(ctx context.Context, credentialsFile, tokenFile, title string) (*GoogleSheet, error) {
	srv, err := GoogleSheetsService(ctx, credentialsFile, tokenFile)
	if err != nil {
		return nil, err
	}
	rowData := make(map[string][]*sheets.RowData)
	sheet, err := createSheet(ctx, srv, title, rowData)
	if err != nil {
		return nil, err
	}
	return &GoogleSheet{
		id:          sheet.SpreadsheetId,
		spreadsheet: sheet,
		service:     srv,
	}, nil
}

func Open(ctx context.Context, credentialsFile, tokenFile, id string) (*GoogleSheet, error) {
	srv, err := GoogleSheetsService(ctx, credentialsFile, tokenFile)
	if err != nil {
		return nil, err
	}
	return &GoogleSheet{
		id:      id,
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
