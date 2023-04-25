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

func NewExcelImportWriter[T any](reader io.Reader, t T) *Import[T] {
	e := Import[T]{}
	return e.newExcelImportWriter(reader, "", t)
}

func NewExcelImportSheetFile[T any](fileName string, sheetName string, t T) *Import[T] {
	e := Import[T]{}
	return e.newExcelImportFile(fileName, sheetName, t)
}

func NewExcelImportSheetWriter[T any](reader io.Reader, sheetName string, t T) *Import[T] {
	e := Import[T]{}
	return e.newExcelImportWriter(reader, sheetName, t)
}

func (e *Import[T]) ImportRead(fu func(row T)) *Import[T] {
	return e.importRead(fu)
}

func (e *Import[T]) ImportDataToStruct(t *[]T) *Import[T] {
	return e.importDataToStruct(t)
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
	return e.exportData(object, 1)
}

// ExportData 指定位置导出 start 默认从1开始 1 数据开始的位置
func (e *Export[T]) ExportData(object []T, start int) *Export[T] {
	return e.exportData(object, start)
}

func (e *excelModel[T]) WriteInWriter(writer io.Writer) *excelModel[T] {
	err := e.f.Write(writer)
	if err != nil {
		e.err = err
	}
	return e
}

func (e *excelModel[T]) WriteInFileName(resultFile string) *excelModel[T] {
	file, err := os.OpenFile(resultFile, os.O_CREATE|os.O_WRONLY|os.O_TRUNC, os.ModeDevice|os.ModePerm)
	defer file.Close()

	if err == nil {
		writer := bufio.NewWriter(file)
		e.WriteInWriter(writer)
		writer.Flush()
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
