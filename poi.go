package excel

import (
	"bufio"
	"github.com/xuri/excelize/v2"
	"io"
	"os"
)

type Export[T any] struct {
	excelModel[T]
}

type Import[T any] struct {
	excelModel[T]
}

func NewExcelImportFile[T any](fileName string, t T) *Import[T] {
	e := Import[T]{}
	return e.newExcelImportFile(fileName, "", t)
}

func (e *Import[T]) ImportDataToStruct() ([]T, error) {
	return e.importDataToStruct()
}

// NewExcelExport 导出初始化
func NewExcelExport[T any](sheetName string, t T) *Export[T] {
	e := Export[T]{}
	return e.newExcelExport(sheetName, t)
}

func (e *Export[T]) SetHeadStyle(style *excelize.Style) *Export[T] {
	return e.setHeadStyle(style)
}

func (e *Export[T]) ExportSmallExcelByStruct(object []T) *Export[T] {
	return e.exportData(object, e.headRowHeight+1)
}

func (e *Export[T]) ExportData(object []T, start int) *Export[T] {
	return e.exportData(object, start)
}

func (e *excelModel[T]) WriteInWriter(writer io.Writer) {
	err := e.f.Write(writer)
	if err != nil {
		e.err = err
	}
}

func (e *excelModel[T]) WriteInFileName(resultFile string) *excelModel[T] {
	file, err := os.OpenFile(resultFile, os.O_CREATE|os.O_WRONLY|os.O_TRUNC, os.ModeDevice|os.ModePerm)
	defer file.Close()
	if err == nil {
		writer := bufio.NewWriter(file)
		err = e.f.Write(writer)
	}
	e.err = err
	return e
}

func (e *excelModel[T]) Close() {
	err := e.f.Close()
	if err != nil {
		return
	}
}

func (e *excelModel[T]) Error() error {
	return e.err
}
