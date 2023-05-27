### 注意
本包大量使用反射，介意性能慎用
### Easy to export excel
基于github.com/xuri/excelize/v2的快速导出struct工具
- 基于excelize v2，完全兼容原生函数
- 通过tag设置header，方便快捷
- 多sheet导出
### 安装
```shell
go get github.com/anjude/goexcel
```
### 使用示例
```GO
package main

import (
	"fmt"
	"github.com/Pallinder/go-randomdata"
	"github.com/anjude/goexcel"
)

type Student struct {
	Name string `xlsx:"name"`
	Age  int    `xlsx:"age"`
}

func getData() []Student {
	var datas []Student
	for i := 0; i < 10; i++ {
		datas = append(datas, Student{
			randomdata.RandStringRunes(1),
			i,
		})
	}
	return datas
}

func main() {
	file := goexcel.NewFile()
	err := file.ExportStruct(getData())
	if err != nil {
		fmt.Printf("export err: %v", err)
		return
	}
	err = file.ExportStruct(getData(), goexcel.FileOption{SheetName: "sheet2"})
	if err != nil {
		fmt.Printf("export sheet2 err: %v", err)
		return
	}
	err = file.SaveAs("test_file.xlsx")
	if err != nil {
		fmt.Printf("save err: %v", err)
		return
	}
}

```
### TODO
- [x] 支持多sheet导出 