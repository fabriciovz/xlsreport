package xlsreport

import (
	"fmt"
	"reflect"

	"github.com/xuri/excelize/v2"
)

type xlsReport struct {
	xls           *excelize.File
	sheetName     string
	headersConfig map[string]Header
	report        []interface{}
	startCol      string
	startDataRow  int
	headerRow     int
}

type XLSReport interface {
	GenerateXLSReport() ([]byte, error)
}

func NewXLSReport(report []interface{},
	headersConfig map[string]Header) XLSReport {
	return &xlsReport{
		report:        report,
		headersConfig: headersConfig,
		startCol:      StartColDefault,
		startDataRow:  StartDataRowDefault,
		headerRow:     HeaderRowDefault,
	}
}

func (x *xlsReport) GenerateXLSReport() ([]byte, error) {
	x.xls = excelize.NewFile()
	sheet, _ := x.xls.NewSheet(XLSSheetName)
	x.sheetName = x.xls.GetSheetName(sheet)
	x.addHeaders()
	x.fill()
	x.xls.SetActiveSheet(sheet)
	buffer, err := x.xls.WriteToBuffer()
	return buffer.Bytes(), err
}

func (x *xlsReport) addHeaders() {
	for row := range x.headersConfig {
		_ = x.xls.SetColWidth(x.sheetName, row, row, x.headersConfig[row].Width)
		var headersStyle, _ = x.xls.NewStyle(&excelize.Style{
			Font: &excelize.Font{Size: x.headersConfig[row].FontSize,
				Color:  x.headersConfig[row].FontColor,
				Bold:   true,
				Italic: false,
				Family: x.headersConfig[row].FontFamily,
			},
			Fill: excelize.Fill{Color: []string{x.headersConfig[row].Color},
				Type:    "pattern",
				Pattern: 1},
		})
		cell := fmt.Sprintf("%s%d", x.headersConfig[row].ColName, x.headerRow)
		_ = x.xls.SetCellValue(x.sheetName, cell, x.headersConfig[row].Name)
		_ = x.xls.SetCellStyle(x.sheetName, cell, cell, headersStyle)
	}
}

func (x *xlsReport) fill() {
	actualRow := x.startDataRow
	for i := range x.report {
		rowA := fmt.Sprintf("%s%d", x.startCol, actualRow)
		row := x.getValues(&x.report[i])
		_ = x.xls.SetSheetRow(x.sheetName, rowA, &row)
		actualRow++
	}
}
func (x *xlsReport) getValues(diffsFounded *interface{}) []interface{} {
	y := reflect.ValueOf(*diffsFounded)
	values := make([]interface{}, y.NumField())
	for i := 0; i < y.NumField(); i++ {
		values[i] = y.Field(i).Interface()
	}
	return values
}
