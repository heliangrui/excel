## excel工具

#### 标签

- excelName 列名
- excelIndex 列序号
- toExcelFormat 列转excel函数名称
- toDataFormat excel转data函数名称
- excelColWidth 列宽度

结构体示例：
```
type NameStruct struct{
	Name string `excelName:"姓名" excelIndex:"1" excelColWidth:"30"`
	Age string `excelName:"年龄" excelIndex:"3"`
	Sex int `excelName:"性别" excelIndex:"1" toExcelFormat:"ToExcelSexFormat"`
}

func (n NameStruct) ToExcelSexFormat() string{
    if n.Sex == 0 {
		return "女"
    }
	return "男"
}
```

### 导出
```
func main() {
    //创建数据源
	data := createData()
	//创建导出对象
	export := excel.NewExcelExport("test", NameStruct{})
	//销毁对象
	defer export.Close()
	//导出
	err = export.ExportSmallExcelByStruct(data).WriteInFileName("test.xlsx").Error()
	if err != nil {
		fmt.Println("生成失败", err.Error())
	}
}

func createData() []NameStruct {
	var names []NameStruct
	for i := 0; i < 10; i++ {
		names = append(names, NameStruct{name: "hlr" + strconv.Itoa(i), age: strconv.Itoa(i),Sex: i})
	}
	return names
}

```


### 导入
```
func main() {

    
    //接受数据
    var result []NameStruct
    //创建导入对象
	importFile := excel.NewExcelImportFile("111.xlsx", NameStruct{})
	//对象销毁
	defer importFile.Close()
	
	// 方式一
	//数据填充 
	err := importFile.ImportDataToStruct(&result).Error()
    //数据显示
	if err != nil {
		fmt.Println("生成失败", err.Error())
	} else {
		marshal, _ := json.Marshal(result)
		fmt.Println(string(marshal))
	}
    
    // 方式二 逐行遍历
    err := importFile.ImportRead(func(row NameStruct) {
		fmt.Println(row.Name)
	}).Error()
    
}

```