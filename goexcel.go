// Package goexcel provides functionality for exporting Go structs to Excel files.
package goexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
)

type FileOption struct {
	SheetName string
	TagName   string
}

func GetDefaultFileOptions() *FileOption {
	return &FileOption{
		SheetName: defaultSheetName,
		TagName:   defaultTagName,
	}
}

type File struct {
	*excelize.File
	*FileOption
}

// NewFile creates a new Excel file.
func NewFile() *File {
	return &File{excelize.NewFile(), GetDefaultFileOptions()}
}

// ExportStruct writes data from a Go slice of structs to an Excel file.
func (f *File) ExportStruct(datas interface{}, opt ...FileOption) error {
	// Check if data exist
	sliceVal := reflect.ValueOf(datas)
	if datas == nil || sliceVal.Kind() != reflect.Slice || sliceVal.Len() == 0 {
		return ErrNoData
	}

	// 设置options
	f.mergeFileOption(opt)
	// 创建sheet
	if !StringInSlice(f.GetSheetList(), f.SheetName) {
		if _, err := f.NewSheet(f.SheetName); err != nil {
			return err
		}
	}
	// 设置表头
	if err := f.SetHeader(sliceVal.Index(0).Interface()); err != nil {
		return err
	}
	// 写入数据
	if err := f.WriterData(sliceVal); err != nil {
		return err
	}
	return nil
}

// mergeFileOption merges the given file options with the default options.
func (f *File) mergeFileOption(options []FileOption) {
	f.FileOption = GetDefaultFileOptions()
	for _, o := range options {
		if o.SheetName != "" {
			f.SheetName = o.SheetName
		}
		if o.TagName != "" {
			f.TagName = o.TagName
		}
	}
}

// SetHeader writes the header row to the Excel file.
func (f *File) SetHeader(data interface{}) error {
	fieldValue, fieldType := getStructValue(data)

	for i := 0; i < fieldValue.NumField(); i++ {
		field := fieldType.Field(i)
		fieldName := field.Name
		if tag := field.Tag.Get(f.TagName); tag != "" {
			fieldName = tag
		}
		colLetter := getColLetter(i + 1)
		header := fmt.Sprintf("%s%d", colLetter, 1)
		if err := f.SetColWidth(f.SheetName, colLetter, colLetter, 20); err != nil {
			return err
		}
		if err := f.SetCellValue(f.SheetName, header, fieldName); err != nil {
			return err
		}
	}
	return nil
}

// WriterData writes data rows to the Excel file.
func (f *File) WriterData(val reflect.Value) error {
	for i := 0; i < val.Len(); i++ {
		row := val.Index(i)
		for j := 0; j < row.Type().NumField(); j++ {
			colLetter := getColLetter(j + 1)
			cell := fmt.Sprintf("%s%d", colLetter, i+2)
			value := fmt.Sprintf("%v", row.Field(j).Interface())
			if err := f.SetCellValue(f.SheetName, cell, value); err != nil {
				return err
			}
		}
	}
	return nil
}

// getColLetter returns the column letter for the given column index.
func getColLetter(index int) (colLetter string) {
	for index > 0 {
		index--
		colLetter = string(rune('A'+index%26)) + colLetter
		index /= 26
	}
	return colLetter
}

// getStructValue returns the reflect.Value and reflect.Type for the given data.
func getStructValue(data interface{}) (reflect.Value, reflect.Type) {
	// 通过反射获取接口的类型和值
	iv := reflect.ValueOf(data)
	it := reflect.TypeOf(data)

	// 判断接口是否为指针类型，并取出指针所指向的值
	if it.Kind() == reflect.Ptr {
		iv = iv.Elem()
		it = it.Elem()
	}
	return iv, it
}
