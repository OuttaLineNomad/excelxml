// Package excelxml converts xml files with excel schemas.
//
// Converts excel files that are formated in xml to a slice.
// The slice can be saved in to a CSV file. Either saving each
// sheet as a different file, or saving all sheets to one file.
//
package excelxml

import (
	"encoding/csv"
	"encoding/xml"
	"errors"
	"io"
	"os"
)

// Workbook is the whole workbook with multiple sheets.
type Workbook struct {
	Worksheets []Worksheet
}

// Worksheet holds sheet data.
// Name is sheet name and Table is all data.
type Worksheet struct {
	Name  string
	Table [][]string
}

// xmlSheet xml structure of a worksheet in file.
// XML data only works for excel XML files
type xmlSheet struct {
	Name  string `ss:"Name"`
	Table []row  `xml:"Table>Row"`
}

// row xml structure of a row in file.
// Represents one fow of data in XML file.
// In "Cell" is also type data, but due to
// the nature of this project we don't need it
type row struct {
	Cells []struct {
		Data string `xml:"Data"`
	} `xml:"Cell"`
}

// Error type to share error messages from package.
type Error struct {
	Func string
	Msg  string
	Err  error
}

// Error is function to make format error message.
func (er *Error) Error() string {
	return `excelxml.` + er.Func + `: ` + er.Msg + ` : ` + er.Err.Error()
}

// SliceXML takes excel xml reader converts to an slice of sheets.
// *Workbook is returned with which holds the slice of sheets. In
// it has the sheet name, and a slice of data from the sheet
func SliceXML(r io.Reader) (*Workbook, error) {
	decoder := xml.NewDecoder(r)
	sheets := &Workbook{}
	for {
		t, err := decoder.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			return nil, &Error{"SliceXML", "getToken", err}
		}

		if t == nil {

		}
		switch ty := t.(type) {
		case xml.StartElement:
			if ty.Name.Local == "Worksheet" {
				ws := xmlSheet{}
				sheet := [][]string{}
				wSheet := Worksheet{}
				decoder.DecodeElement(&ws, &ty)
				for _, row := range ws.Table {
					newRow := []string{}
					for _, cell := range row.Cells {
						newRow = append(newRow, cell.Data)
					}
					sheet = append(sheet, newRow)
				}
				wSheet.Name = ty.Attr[0].Value
				wSheet.Table = sheet
				sheets.Worksheets = append(sheets.Worksheets, wSheet)
			}

		}
	}
	if len(sheets.Worksheets) == 0 {
		return nil, &Error{"SliceXML", "finding sheets", errors.New("file is not an excel xml file")}
	}
	return sheets, nil
}

// SaveCSV saves Workbook to a CSV file.
// If Multiple has true multiple sheets will be saved to multiple files.
func (wb *Workbook) SaveCSV(Multiple ...bool) error {
	multi := false
	if len(Multiple) > 0 {
		multi = Multiple[0]
	}

	// creating multiple files from multiple sheets.
	if multi {
		for _, sheet := range wb.Worksheets {
			f, err := os.Create(sheet.Name + ".csv")
			if err != nil {
				return &Error{"SaveCSV", "create multiple files", err}
			}
			cv := csv.NewWriter(f)
			cv.WriteAll(sheet.Table)
			cv.Flush()
		}
		return nil
	}

	// If single file is desired then one file and all sheets are added to this file.
	// Ignoring titles of all subsequent sheets. This is default.
	f, err := os.Create(wb.Worksheets[0].Name + ".csv")
	if err != nil {
		return &Error{"SaveCSV", "create singal file", err}
	}
	cs := csv.NewWriter(f)
	fullFile := [][]string{}

	for sNum, sheet := range wb.Worksheets {
		lenSh := len(sheet.Table)
		if lenSh == 0 {
			continue
		}
		if sNum != 0 && lenSh > 1 {
			sheet.Table = sheet.Table[1:]
		}
		fullFile = append(fullFile, sheet.Table...)
	}

	cs.WriteAll(fullFile)
	cs.Flush()
	return nil
}
