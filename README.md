# XLSReport for Golang

**xlsreport** is a library which implements (https://github.com/qax-os/excelize). The main purpose is to hide the "excelize" details when you have to do a simple xls report.
Instead of coding every time you need to create an xls report, dealing with column numbers and so on, here you just pass the data and the header configuration.

- this library is still construction :).
- does not require any modifications to your source code.
- has a third party dependency (github.com/xuri/excelize/v2)

## Looking for maintainers

I do not have much spare time for this library and willing to transfer the repository ownership
to person or an organization motivated to maintain it. Open up a conversation if you are interested. 

## Install

    go get github.com/fabriciovz/xlsreport

## Documentation and Examples


### How to use the xlsreport lib

``` go
package main

import (
	"fmt"
	"github.com/fabriciovz/xlsreport"
	"os"
)

// Summary report (entity of what you want to export to an XLS file)
// ***Something important to mention is that the fields has to be exported or public
type Summary struct {
	Name    string `json:"name"`
	Surname string `json:"surname"`
	Age     int    `json:"age"`
}

func main() {
	//Adding the header configuration
	var headerConfig = map[string]xlsreport.Header{
		"A": {
			ColName:    "A",
			Name:       "Name",
			Color:      "FFFB04",
			FontSize:   14,
			FontColor:  "#000000",
			FontFamily: "Calibri",
			Width:      40,
		},
		"B": {
			ColName:    "B",
			Name:       "Surname",
			Color:      "FFFB04",
			FontSize:   14,
			FontColor:  "#000000",
			FontFamily: "Calibri",
			Width:      20,
		},
		"C": {
			ColName:    "C",
			Name:       "Age",
			Color:      "FFFB04",
			FontSize:   14,
			FontColor:  "#000000",
			FontFamily: "Calibri",
			Width:      20,
		},
	}
	//Data you want to export in an XLS file
	var response = []*Summary{{
		Name:    "Fabricio",
		Surname: "Bravo Guevara",
		Age:     37,
	},
	}
	//map from your specific response data type to an interface
	data := mapReportToInterface(response)

	xls := xlsreport.NewXLSReport(data, headerConfig)
	report, err := xls.GenerateXLSReport()

	if err != nil {
		panic(err)
	}

	err = os.WriteFile("my_xls_file.xlsx", report, 0644) //create a new file
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("File is created successfully.")
}

// You have to map from  entity to an interface
func mapReportToInterface(report []*Summary) []interface{} {
	y := make([]interface{}, len(report))
	for i, v := range report {
		y[i] = *v
	}
	return y
}

```

## Change Log

- **2024-02-07** - adding the functionality to generate a simple xls report

## Contributions

Feel free to open a pull request. Note, if you wish to contribute an extension to public (exported methods or types) -
please open an issue before, to discuss whether these changes can be accepted. All backward incompatible changes are
and will be treated cautiously

## License

The [three clause BSD license](http://en.wikipedia.org/wiki/BSD_licenses)