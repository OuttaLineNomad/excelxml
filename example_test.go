package excelxml_test

import (
	"fmt"
	"log"
	"os"

	"github.com/OuttaLineNomad/excelxml"
)

// SliceXML cannot parse xls files, but excel xml files can be named under
// xls when they are not xls.
func ExampleSliceXML() {
	// excel XML files can be named as XLS files, but they are xml.
	f, err := os.Open("test/Test_Excel_Book.xls")
	if err != nil {
		log.Panic(err)
	}
	sheets, err := excelxml.SliceXML(f)
	if err != nil {
		log.Panic(err)
	}
	oneSheet := [][]string{}
	for i, sheet := range sheets.Worksheets {
		table := sheet.Table
		if i != 0 {
			table = table[1:]
		}
		oneSheet = append(oneSheet, table...)
	}
	fmt.Println(oneSheet)
	// Output:[[Title Value] [This is a title! This is a Value!] [Outtaline Nomad is cool YEP!]]
}

func ExampleSliceXML_Second() {
	// excel XML files can be named as XLS files, but they are xml.
	f, err := os.Open("test/Test_Excel_Book.xml")
	if err != nil {
		log.Panic(err)
	}
	sheets, err := excelxml.SliceXML(f)
	if err != nil {
		log.Panic(err)
	}
	oneSheet := [][]string{}
	for i, sheet := range sheets.Worksheets {
		table := sheet.Table
		if i != 0 {
			table = table[1:]
		}
		oneSheet = append(oneSheet, table...)
	}
	fmt.Println(oneSheet)
	// Output:[[Title Value] [This is a title! This is a Value!] [Outtaline Nomad is cool YEP!]]
}
