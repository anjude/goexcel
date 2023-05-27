package goexcel

import (
	"fmt"
	"github.com/Pallinder/go-randomdata"
	"testing"
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
			randomdata.Number(1),
		})
	}
	return datas
}

func TestNewFile(t *testing.T) {
	file := NewFile()
	err := file.ExportStruct(getData())
	if err != nil {
		fmt.Printf("export err: %v", err)
		return
	}
	err = file.ExportStruct(getData(), FileOption{SheetName: "sheet2"})
	if err != nil {
		fmt.Printf("export err: %v", err)
		return
	}
	err = file.SaveAs("test_file.xlsx")
	if err != nil {
		fmt.Printf("save err: %v", err)
		return
	}
}

func Test(t *testing.T) {
	//viper.Unmarshal(nil)
}
