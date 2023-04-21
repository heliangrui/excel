package main

import (
	"fmt"
	"github.com/heliangrui/excel"
	"strconv"
	"time"
)

type ExportDeviceVo struct {
	Id          string `json:"id" example:"ssss" excelName:"设备编码"`
	Area        string `json:"area" example:"ssss" excelName:"设备区域"`
	Name        string `json:"name" example:"name" excelName:"设备名称"`
	ClassId     string `json:"classId" example:"ssss"  excelName:"原型编码" `
	Description string `json:"description" example:"ssss" excelName:"设备描述"`
}

func main() {

	// 测试导出
	now := time.Now()
	testExport()
	dataTime := time.Now().UnixMilli() - now.UnixMilli()
	fmt.Println("总执行时间：", dataTime)
}

func testExport() {

	data := createData(10000000)
	// 测试一次性导出
	name := excel.NewExcelExport("domeSheet", ExportDeviceVo{}).ExportSmallExcelByStruct(data).WriteInFileName("testExport.xlsx")
	defer name.Close()
}

func testExportAsc() {

	name := excel.NewExcelExport("domeSheet", ExportDeviceVo{})

	for i := 0; i < 100; i++ {
		data := createData(100000)
		start := 1
		if i != 0 {
			start = i*len(data) + 1
		}
		go name.ExportData(data, start)
	}

	name.WriteInFileName("testExportAsc.xlsx")
	defer name.Close()

	fmt.Println(name)
}

func createData(num int) []ExportDeviceVo {
	var result []ExportDeviceVo
	for i := 0; i < num; i++ {
		itoa := strconv.Itoa(i)
		result = append(result, ExportDeviceVo{Id: itoa, Area: itoa + itoa, Name: itoa + itoa + itoa, ClassId: itoa + itoa + itoa + itoa})
	}
	return result
}
