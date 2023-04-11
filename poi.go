package excel

import (
	"bufio"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"os"
	"reflect"
	"strconv"
)

type Excel[T any] struct {
	fileName      string
	sheetName     string
	f             *excelize.File
	t             T
	mod           []model
	headRowHeight int
	err           error
}

func (e *Excel[T]) NewExcelExport(sheetName string, t T) *Excel[T] {
	e.t = t
	e.sheetName = sheetName
	e.mod = *getInterfaceExcelModel(t)
	e.f = excelize.NewFile()
	if sheetName != DefaultSheet {
		_ = e.f.DeleteSheet(DefaultSheet)
	}
	// 创建sheet
	index, err := e.f.NewSheet(sheetName)
	if err != nil {
		fmt.Println("表格创建失败")
		e.err = err
	}
	e.f.SetActiveSheet(index)
	e.headRowHeight = 1
	//创建表头
	e.createHead()
	e.SetHeadStyle(CreateDefaultHeader())
	if err != nil {
		e.err = err
	}
	return e
}

func (e *Excel[T]) SetHeadStyle(style *excelize.Style) {
	newStyle, err := e.f.NewStyle(style)
	if err != nil {
		fmt.Println("样式创建失败！")
		e.err = err
	}
	start, _ := excelize.ColumnNumberToName(1)
	start += strconv.Itoa(1)
	end, _ := excelize.ColumnNumberToName(len(e.mod))
	end += strconv.Itoa(1)
	err = e.f.SetCellStyle(e.sheetName, start, end, newStyle)
}

func (e *Excel[T]) ExportSmallExcelByStruct(object []T) *Excel[T] {
	return e.ExportData(object, e.headRowHeight+1)
}

func (e *Excel[T]) ExportData(object []T, start int) *Excel[T] {
	for i := 0; i < len(object); i++ {
		mod := object[i]
		value := reflect.ValueOf(mod)
		for r := 0; r < len(e.mod); r++ {
			fieldName := e.mod[r].fieldName
			nowValue := value.FieldByName(fieldName)
			name, _ := excelize.ColumnNumberToName(r + 1)
			s := name + strconv.Itoa(i+start)

			if e.mod[r].toExcelFormat == "" {
				_ = e.f.SetCellValue(e.sheetName, s, nowValue)
			} else {
				toExcelFun := value.MethodByName(e.mod[r].toExcelFormat)
				call := toExcelFun.Call(nil)
				_ = e.f.SetCellValue(e.sheetName, s, call[0])
			}
		}
	}
	return e
}

func (e *Excel[T]) WriteInWriter(writer io.Writer) {
	err := e.f.Write(writer)
	if err != nil {
		e.err = err
	}
}

func (e *Excel[T]) WriteInFileName(resultFile string) *Excel[T] {
	file, err := os.OpenFile(resultFile, os.O_CREATE|os.O_WRONLY|os.O_TRUNC, os.ModeDevice|os.ModePerm)
	defer file.Close()
	if err == nil {
		writer := bufio.NewWriter(file)
		err = e.f.Write(writer)
	}
	e.err = err
	return e
}

func (e *Excel[T]) Close() {
	err := e.f.Close()
	if err != nil {
		return
	}
}

func (e *Excel[T]) Error() error {
	return e.err
}
