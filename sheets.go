package sheets

import (
	"context"
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"google.golang.org/api/sheets/v4"
)

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

func ResizeColumns(ctx context.Context, srv *sheets.Service, spreadsheet sheets.Spreadsheet) error {
	// Final sheet updates:
	// - Auto-resize the  columns of the spreadsheet to fit.
	var requests []*sheets.Request
	for _, sheet := range spreadsheet.Sheets {
		requests = append(requests, &sheets.Request{
			AutoResizeDimensions: &sheets.AutoResizeDimensionsRequest{
				Dimensions: &sheets.DimensionRange{
					Dimension: "COLUMNS",
					SheetId:   sheet.Properties.SheetId,
				},
			},
		})
	}
	_, err := srv.Spreadsheets.BatchUpdate(spreadsheet.SpreadsheetId, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: requests,
	}).Context(ctx).Do()
	return err
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

// Write populates the given rowData with the given data.
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
