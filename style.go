package excel

import (
	"github.com/xuri/excelize/v2"
)

func CreateDefaultHeader() *excelize.Style {
	style := excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "050505", Style: 1},
			{Type: "top", Color: "050505", Style: 1},
			{Type: "bottom", Color: "050505", Style: 1},
			{Type: "right", Color: "050505", Style: 1},
		},
		Fill: excelize.Fill{Type: "gradient", Color: []string{"#a6a6a6", "#a6a6a6"}, Shading: 1},
		Font: nil,
		Alignment: &excelize.Alignment{
			Horizontal:      "center",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    0,
			RelativeIndent:  1,
			ShrinkToFit:     false,
			TextRotation:    0,
			Vertical:        "",
			WrapText:        false,
		},
		Protection:    nil,
		NumFmt:        0,
		DecimalPlaces: 0,
		CustomNumFmt:  nil,
		Lang:          "",
		NegRed:        false,
	}
	return &style
}
