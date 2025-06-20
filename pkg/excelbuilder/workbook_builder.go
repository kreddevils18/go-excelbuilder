package excelbuilder

import (
	"strings"

	"github.com/xuri/excelize/v2"
)

// WorkbookBuilder handles workbook-level operations
type WorkbookBuilder struct {
	excelBuilder *ExcelBuilder
	file         *excelize.File
}

// SetProperties sets the workbook properties
func (wb *WorkbookBuilder) SetProperties(props WorkbookProperties) *WorkbookBuilder {
	// Set document properties
	if props.Title != "" {
		wb.file.SetDocProps(&excelize.DocProperties{
			Title:       props.Title,
			Creator:     props.Author,
			Subject:     props.Subject,
			Description: props.Description,
			Category:    props.Category,
			Keywords:    props.Keywords,
		})
	}
	return wb
}

// AddSheet creates a new sheet and returns a SheetBuilder
func (wb *WorkbookBuilder) AddSheet(name string) *SheetBuilder {
	if name == "" {
		return nil
	}

	// Check for invalid characters
	invalidChars := []string{"[", "]", "*", "?", "/", "\\"}
	for _, char := range invalidChars {
		if strings.Contains(name, char) {
			return nil
		}
	}

	// Check length (Excel limit is 31 characters)
	if len(name) > 31 {
		return nil
	}

	// Check for reserved names
	reservedNames := []string{"History", "Print_Area", "Print_Titles"}
	for _, reserved := range reservedNames {
		if name == reserved {
			return nil
		}
	}

	// Create the sheet
	index, err := wb.file.NewSheet(name)
	if err != nil {
		return nil
	}

	// Set as active sheet
	wb.file.SetActiveSheet(index)

	return &SheetBuilder{
		workbookBuilder: wb,
		sheetName:       name,
		currentRow:      0,
	}
}

// AddSheetsBatch creates multiple sheets with data in a single call.
func (wb *WorkbookBuilder) AddSheetsBatch(sheets []SheetConfig) *WorkbookBuilder {
	for _, sheetConfig := range sheets {
		sheet := wb.AddSheet(sheetConfig.Name)
		if sheet != nil {
			sheet.AddRows(sheetConfig.Data)
		}
	}
	return wb
}

// Build returns the final excelize.File
func (wb *WorkbookBuilder) Build() *excelize.File {
	return wb.file
}