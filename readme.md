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
	e2 := Excel[NameStruct]{}
	//导出
	err := e2.NewExcelExport("hlr", NameStruct{}).ExportSmallExcelByStruct(data).WriteInFileName("111.xlsx").Error()
	//销毁对象
	defer e2.Close()
	
	if err == nil {
		fmt.Println("生成成功")
	} else {
		fmt.Println("生成失败")
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
