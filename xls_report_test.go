package xlsreport

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestXLSReport(t *testing.T) {
	// Summary report (entity of what you want to export to an XLS file)
	// ***Something important to mention is that the fields has to be exported or public
	type Summary struct {
		Name    string `json:"name"`
		Surname string `json:"surname"`
		Age     int    `json:"age"`
	}

	// Adding the header configuration
	headerConfig := map[string]Header{
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

	// Data you want to export in an XLS file
	response := []*Summary{{
		Name:    "Fabricio",
		Surname: "Bravo Guevara",
		Age:     37,
	},
	}

	t.Run("given an empty command, when run the use case process, then it should not return error", func(t *testing.T) {
		xls := NewXLSReport(response, headerConfig)

		_, err := xls.GenerateXLSReport()

		assert.NoError(t, err)
	})
}
