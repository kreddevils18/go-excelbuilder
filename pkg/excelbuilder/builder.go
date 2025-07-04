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

// ConvertCSVData converts CSV data to Excel format
func (eb *ExcelBuilder) ConvertCSVData(csvData [][]string) *WorkbookBuilder {
	workbook := eb.NewWorkbook()
	sheet := workbook.AddSheet("Sheet1")

	for _, row := range csvData {
		rowBuilder := sheet.AddRow()
		for _, cell := range row {
			rowBuilder.AddCell(cell)
		}
	}

	return workbook
}

// ConvertJSONToWorkbook converts JSON data to Excel workbook
func (eb *ExcelBuilder) ConvertJSONToWorkbook(jsonData map[string]interface{}) *WorkbookBuilder {
	workbook := eb.NewWorkbook()
	sheet := workbook.AddSheet("Sheet1")

	// Simple JSON to Excel conversion
	if data, ok := jsonData["data"].([]interface{}); ok {
		for _, item := range data {
			if itemMap, ok := item.(map[string]interface{}); ok {
				rowBuilder := sheet.AddRow()
				for _, value := range itemMap {
					rowBuilder.AddCell(value)
				}
			}
		}
	}

	return workbook
}

// WithStreamingMode enables streaming mode for large datasets
func (eb *ExcelBuilder) WithStreamingMode(enabled bool) *ExcelBuilder {
	// For now, just return the same builder
	// In a full implementation, this would configure streaming
	return eb
}

// WithErrorCollection enables error collection mode
func (eb *ExcelBuilder) WithErrorCollection(enabled bool) *ExcelBuilder {
	// For now, just return the same builder
	// In a full implementation, this would enable error collection
	return eb
}

// GetCollectedErrors returns collected errors
func (eb *ExcelBuilder) GetCollectedErrors() []error {
	// For now, return empty slice
	// In a full implementation, this would return actual collected errors
	return []error{}
}

// TransformDataToPivot transforms raw data to pivot format
func (eb *ExcelBuilder) TransformDataToPivot(rawData []map[string]interface{}, config PivotConfig, sheetName string) *WorkbookBuilder {
	// Simple pivot transformation
	workbook := eb.NewWorkbook()
	sheet := workbook.AddSheet(sheetName)

	// Add headers based on pivot configuration
	headerRow := sheet.AddRow()
	for _, field := range config.RowFields {
		headerRow.AddCell(field)
	}
	for _, field := range config.ColumnFields {
		headerRow.AddCell(field)
	}
	for _, field := range config.ValueFields {
		headerRow.AddCell(field)
	}

	// Add data (simplified)
	for _, row := range rawData {
		dataRow := sheet.AddRow()
		for _, field := range config.RowFields {
			if value, ok := row[field]; ok {
				dataRow.AddCell(value)
			}
		}
		for _, field := range config.ColumnFields {
			if value, ok := row[field]; ok {
				dataRow.AddCell(value)
			}
		}
		for _, field := range config.ValueFields {
			if value, ok := row[field]; ok {
				dataRow.AddCell(value)
			}
		}
	}

	return workbook
}
