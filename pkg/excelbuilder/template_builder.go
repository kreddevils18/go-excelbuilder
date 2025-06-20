package excelbuilder

import (
	"fmt"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

// TemplateBuilder provides functionality for creating and processing Excel templates
// with placeholder substitution and dynamic content generation.
type TemplateBuilder struct {
	excelBuilder *ExcelBuilder
	file         *excelize.File
	currentSheet string
	currentRow   int
	currentCol   int
	templateData map[string]interface{}
}

// TemplateCellBuilder handles individual cell operations in template building
type TemplateCellBuilder struct {
	templateBuilder *TemplateBuilder
	cellRef         string
	value           interface{}
}

// TemplateRowBuilder handles row operations in template building
type TemplateRowBuilder struct {
	templateBuilder *TemplateBuilder
	rowIndex        int
}

// TemplateSheetBuilder handles sheet operations in template building
type TemplateSheetBuilder struct {
	templateBuilder *TemplateBuilder
	sheetName       string
}

// NewTemplateBuilder creates a new TemplateBuilder instance
func NewTemplateBuilder() *TemplateBuilder {
	file := excelize.NewFile()
	return &TemplateBuilder{
		excelBuilder: New(),
		file:         file,
		currentRow:   1,
		currentCol:   1,
		templateData: make(map[string]interface{}),
	}
}

// LoadExistingFile loads an existing Excel file for template processing
func LoadExistingFile(filePath string) (*TemplateBuilder, error) {
	file, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file %s: %w", filePath, err)
	}

	return &TemplateBuilder{
		excelBuilder: New(),
		file:         file,
		currentRow:   1,
		currentCol:   1,
		templateData: make(map[string]interface{}),
	}, nil
}

// AddSheet adds a new sheet to the template and returns a TemplateSheetBuilder
func (tb *TemplateBuilder) AddSheet(name string) *TemplateSheetBuilder {
	if tb.currentSheet == "" {
		// Rename the default sheet
		tb.file.SetSheetName("Sheet1", name)
	} else {
		// Create a new sheet
		index, err := tb.file.NewSheet(name)
		if err == nil {
			tb.file.SetActiveSheet(index)
		}
	}
	tb.currentSheet = name
	tb.currentRow = 1
	tb.currentCol = 1
	return &TemplateSheetBuilder{
		templateBuilder: tb,
		sheetName:       name,
	}
}

// AddRow adds a new row to the current sheet and returns a TemplateRowBuilder
func (tsb *TemplateSheetBuilder) AddRow() *TemplateRowBuilder {
	tsb.templateBuilder.currentCol = 1
	return &TemplateRowBuilder{
		templateBuilder: tsb.templateBuilder,
		rowIndex:        tsb.templateBuilder.currentRow,
	}
}

// AddCell adds a new cell to the current row and returns a TemplateCellBuilder
func (trb *TemplateRowBuilder) AddCell() *TemplateCellBuilder {
	cellRef, _ := excelize.CoordinatesToCellName(trb.templateBuilder.currentCol, trb.rowIndex)
	trb.templateBuilder.currentCol++
	return &TemplateCellBuilder{
		templateBuilder: trb.templateBuilder,
		cellRef:         cellRef,
	}
}

// WithValue sets the value for the current cell
func (tcb *TemplateCellBuilder) WithValue(value interface{}) *TemplateCellBuilder {
	tcb.value = value
	tcb.templateBuilder.file.SetCellValue(tcb.templateBuilder.currentSheet, tcb.cellRef, value)
	return tcb
}

// Done completes the cell building and returns the parent TemplateRowBuilder
func (tcb *TemplateCellBuilder) Done() *TemplateRowBuilder {
	return &TemplateRowBuilder{
		templateBuilder: tcb.templateBuilder,
		rowIndex:        tcb.templateBuilder.currentRow,
	}
}

// Done completes the row building and returns the parent TemplateSheetBuilder
func (trb *TemplateRowBuilder) Done() *TemplateSheetBuilder {
	trb.templateBuilder.currentRow++
	return &TemplateSheetBuilder{
		templateBuilder: trb.templateBuilder,
		sheetName:       trb.templateBuilder.currentSheet,
	}
}

// Done completes the sheet building and returns the parent TemplateBuilder
func (tsb *TemplateSheetBuilder) Done() *TemplateBuilder {
	return tsb.templateBuilder
}

// ProcessTemplate processes the template with the provided data
func (tb *TemplateBuilder) ProcessTemplate(data map[string]interface{}) *TemplateBuilder {
	tb.templateData = data
	
	// Get all sheet names
	sheetNames := tb.file.GetSheetList()
	
	for _, sheetName := range sheetNames {
		tb.processSheet(sheetName, data)
	}
	
	return tb
}

// processSheet processes a single sheet with template data
func (tb *TemplateBuilder) processSheet(sheetName string, data map[string]interface{}) {
	rows, err := tb.file.GetRows(sheetName)
	if err != nil {
		return
	}
	
	for rowIndex, row := range rows {
		for colIndex, cellValue := range row {
			if cellValue != "" {
				processedValue := tb.replacePlaceholders(cellValue, data)
				cellRef, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
				tb.file.SetCellValue(sheetName, cellRef, processedValue)
			}
		}
	}
}

// replacePlaceholders replaces template placeholders with actual data
func (tb *TemplateBuilder) replacePlaceholders(template string, data map[string]interface{}) string {
	// Simple placeholder replacement using {{key}} syntax
	re := regexp.MustCompile(`\{\{([^}]+)\}\}`)
	
	return re.ReplaceAllStringFunc(template, func(match string) string {
		// Extract the key from {{key}}
		key := strings.Trim(match, "{}")
		key = strings.TrimSpace(key)
		
		if value, exists := data[key]; exists {
			return fmt.Sprintf("%v", value)
		}
		
		// Return original placeholder if no data found
		return match
	})
}

// Build returns the final Excel file
func (tb *TemplateBuilder) Build() *excelize.File {
	return tb.file
}

// GetCellValue gets the value of a specific cell (for testing purposes)
func (tb *TemplateBuilder) GetCellValue(sheetName, cellRef string) (string, error) {
	return tb.file.GetCellValue(sheetName, cellRef)
}