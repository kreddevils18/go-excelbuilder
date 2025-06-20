package excelbuilder

import (
	"github.com/xuri/excelize/v2"
)

// ExcelBuilder is the main builder for creating Excel files.
// It acts as the entry point for the entire library.
type ExcelBuilder struct {
	file         *excelize.File
	styleManager *StyleManager
}

// New creates a new ExcelBuilder instance.
// This is the starting point for creating a new Excel file.
func New() *ExcelBuilder {
	file := excelize.NewFile()
	return &ExcelBuilder{
		file:         file,
		styleManager: NewStyleManager(),
	}
}

// NewWorkbook creates a new WorkbookBuilder.
// This is the primary way to start building a workbook.
func (eb *ExcelBuilder) NewWorkbook() *WorkbookBuilder {
	return &WorkbookBuilder{
		excelBuilder: eb,
		file:         eb.file,
	}
}
