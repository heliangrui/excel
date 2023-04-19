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

type excelModel[T any] struct {
	fileName      string
	sheetName     string
	f             *excelize.File
	t             T
	mod           *[]*model
	headRowHeight int
	err           error
}

func (e *Import[T]) newExcelImportFile(fileName string, readSheetName string, t T) *Import[T] {
	open, err := os.Open(fileName)
	if err != nil {
		e.err = err
		return e
	}
	defer open.Close()
	reader := bufio.NewReader(open)
	return e.newExcelImportWriter(reader, readSheetName, t)
}

func (e *Import[T]) newExcelImportWriter(reader io.Reader, readSheetName string, t T) *Import[T] {
	e.mod = getInterfaceExcelModel(t)
	openReader, err := excelize.OpenReader(reader)
	if err != nil {
		e.err = err
		return e
	}
	e.f = openReader
	if readSheetName == "" {
		readSheetName = e.f.GetSheetName(e.f.GetActiveSheetIndex())
	}
	e.sheetName = readSheetName
	return e
}

func (e *Import[T]) importRead(fu func(row T)) *Import[T] {
	rows, err := e.f.Rows(e.sheetName)
	if err != nil {
		e.err = err
		return e
	}
	firstRow := true
	for rows.Next() {
		columns, _ := rows.Columns()
		if firstRow {
			for i := 0; i < len(columns); i++ {
				for _, m := range *e.mod {
					if columns[i] == m.excelName {
						m.fieldIndex = i
						break
					}
				}
			}
			firstRow = false
		} else {
			value := reflect.New(reflect.TypeOf(&e.t).Elem())
			value = value.Elem()

			for _, m := range *e.mod {
				item := ""
				if m.fieldIndex < len(columns) {
					item = columns[m.fieldIndex]
				}

				t := value.FieldByName(m.fieldName)

				if m.toDataFormat != "" {
					byNameFunc := value.MethodByName(m.toDataFormat)
					var param []reflect.Value
					param = append(param, reflect.ValueOf(item))
					call := byNameFunc.Call(param)
					t.Set(call[0])
				} else {
					if t.Type().Kind() == reflect.Bool {
						parseBool, err := strconv.ParseBool(item)
						if err == nil {
							t.SetBool(parseBool)
						}
					} else if t.Type().Kind() == reflect.Int || t.Type().Kind() == reflect.Int8 || t.Type().Kind() == reflect.Int32 || t.Type().Kind() == reflect.Int64 {
						i, err := strconv.ParseInt(item, 10, 64)
						if err != nil {
							t.SetInt(i)
						}
					} else if t.Type().Kind() == reflect.String {
						t.Set(reflect.ValueOf(item))
					} else if t.Type().Kind() == reflect.Float32 || t.Type().Kind() == reflect.Float64 {
						float, err := strconv.ParseFloat(item, 64)
						if err == nil {
							t.SetFloat(float)
						}
					}
				}
			}
			temResult := value.Interface().(T)
			fu(temResult)
		}
	}
	return e

}

func (e *Import[T]) importDataToStruct(t *[]T) *Import[T] {
	e.importRead(func(row T) {
		*t = append(*t, row)
	})
	return e
}

func (e *Export[T]) newExcelExport(sheetName string, t T) *Export[T] {
	e.t = t
	e.sheetName = sheetName
	e.mod = getInterfaceExcelModel(t)
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

func (e *Export[T]) setHeadStyle(style *excelize.Style) *Export[T] {
	newStyle, err := e.f.NewStyle(style)
	if err != nil {
		fmt.Println("样式创建失败！")
		e.err = err
	}
	start, _ := excelize.ColumnNumberToName(1)
	start += strconv.Itoa(1)
	end, _ := excelize.ColumnNumberToName(len(*e.mod))
	end += strconv.Itoa(1)
	err = e.f.SetCellStyle(e.sheetName, start, end, newStyle)
	return e
}

func (e *Export[T]) exportData(object []T, start int) *Export[T] {
	for i := 0; i < len(object); i++ {
		mod := object[i]
		value := reflect.ValueOf(mod)
		for r, m := range *e.mod {
			fieldName := m.fieldName
			nowValue := value.FieldByName(fieldName)
			name, _ := excelize.ColumnNumberToName(r + 1)
			s := name + strconv.Itoa(i+start)

			if m.toExcelFormat == "" {
				_ = e.f.SetCellValue(e.sheetName, s, nowValue)
			} else {
				toExcelFun := value.MethodByName(m.toExcelFormat)
				call := toExcelFun.Call(nil)
				_ = e.f.SetCellValue(e.sheetName, s, call[0])
			}
		}
	}
	return e
}
