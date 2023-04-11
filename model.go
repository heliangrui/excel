package excel

import (
	"reflect"
	"sort"
	"strconv"
)

var DefaultSheet = "sheet1"

/*
*
自定义tag
excelName
excelIndex
//excelFormat  暂未实现
excelColWidth
*/
type model struct {
	excelName     string
	excelIndex    int
	excelFormat   string
	excelColWidth int
	fieldName     string
}

// 根据类型获取打印相关内容
func getInterfaceExcelModel(face interface{}) *[]model {
	m := make([]model, 0)
	field := reflect.TypeOf(face)
	//获取tag 根据excelName 获取输出内容,根据 excelIndex 序号 excelFormat 格式化函数 excelColWidth 单元格宽度
	for i := 0; i < field.NumField(); i++ {
		tag := field.Field(i).Tag
		excelName := tag.Get("excelName")
		if excelName != "" {
			indexString := tag.Get("excelIndex")
			index := i
			if indexString != "" {
				parseInt, err := strconv.ParseInt(indexString, 10, 64)
				if err == nil {
					index = int(parseInt)
				}
			}
			name := field.Field(i).Name
			format := tag.Get("excelFormat")
			widthString := tag.Get("excelColWidth")
			var width int
			if widthString != "" {
				parseInt, err := strconv.ParseInt(widthString, 10, 64)
				if err == nil {
					width = int(parseInt)
				}
			} else {
				width = len(excelName) * 3
			}
			m = append(m, model{excelName, index, format, width, name})
		}
	}
	// 排序
	sort.Slice(m, func(i, j int) bool {
		return m[i].excelIndex < m[j].excelIndex
	})
	return &m
}
