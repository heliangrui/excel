package excel

import (
	"github.com/xuri/excelize/v2"
	"strconv"
)

func (e *Export[T]) createHead() {
	for r := 0; r < len(e.mod); r++ {
		name, err := excelize.ColumnNumberToName(r + 1)
		if err != nil {
			e.err = err
		}
		s := name + strconv.Itoa(1)
		e.f.SetCellValue(e.sheetName, s, e.mod[r].excelName)
		e.f.SetColWidth(e.sheetName, name, name, float64(e.mod[r].excelColWidth))
	}
}
