package sheets

import (
	"context"
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"google.golang.org/api/sheets/v4"
)

func (s *Spreadsheet) Append(ctx context.Context, data map[string][]*Row) error {
	rowData, err := convertToRowData(data)
	if err != nil {
		return err
	}
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
	response, err := s.service.Spreadsheets.BatchUpdate(s.id, &sheets.BatchUpdateSpreadsheetRequest{
		IncludeSpreadsheetInResponse: true,
		Requests:                     createRequests,
	}).Context(ctx).Do()
	if err != nil {
		return err
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
	response, err = s.service.Spreadsheets.BatchUpdate(s.id, &sheets.BatchUpdateSpreadsheetRequest{
		IncludeSpreadsheetInResponse: true,
		Requests:                     dataRequests,
	}).Context(ctx).Do()
	if err != nil {
		return err
	}
	s.spreadsheet = response.UpdatedSpreadsheet
	return nil
}

func (s *Spreadsheet) ResizeColumns(ctx context.Context) error {
	// Final sheet updates:
	// - Auto-resize the  columns of the spreadsheet to fit.
	var requests []*sheets.Request
	for _, sheet := range s.spreadsheet.Sheets {
		requests = append(requests, &sheets.Request{
			AutoResizeDimensions: &sheets.AutoResizeDimensionsRequest{
				Dimensions: &sheets.DimensionRange{
					Dimension: "COLUMNS",
					SheetId:   sheet.Properties.SheetId,
				},
			},
		})
	}
	_, err := s.service.Spreadsheets.BatchUpdate(s.id, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: requests,
	}).Context(ctx).Do()
	return err
}

// convertToRowData returns a rowData with the given data.
func convertToRowData(data map[string][]*Row) (map[string][]*sheets.RowData, error) {
	rowData := map[string][]*sheets.RowData{}

	// Write output to disk first.
	var filenames []string
	for filename, cells := range data {
		if len(cells) == 0 {
			continue
		}
		fullpath := filepath.Join("/tmp", fmt.Sprintf("julieqiusheets-%s.csv", filename))
		file, err := os.Create(fullpath)
		if err != nil {
			return nil, err
		}
		defer file.Close()

		writer := csv.NewWriter(file)
		defer writer.Flush()

		for _, row := range cells {
			if err := writer.Write(row.ToCells()); err != nil {
				return nil, err
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
	return rowData, nil
}

func newStrPtr(text string) *string {
	strValuePtr := new(string)
	*strValuePtr = text
	return strValuePtr
}
