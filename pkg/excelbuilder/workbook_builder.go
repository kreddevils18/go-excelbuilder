package excelbuilder

import (
	"fmt"
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
		wb.excelBuilder.AddError(fmt.Errorf("sheet name cannot be empty"))
		// Return a placeholder SheetBuilder to avoid nil
		return &SheetBuilder{
			workbookBuilder: wb,
			sheetName:       "Sheet1", // fallback name
			currentRow:      0,
			hasError:        true,
		}
	}

	// Check for invalid characters
	invalidChars := []string{"[", "]", "*", "?", "/", "\\"}
	for _, char := range invalidChars {
		if strings.Contains(name, char) {
			wb.excelBuilder.AddError(fmt.Errorf("sheet name '%s' contains invalid character '%s'", name, char))
			return &SheetBuilder{
				workbookBuilder: wb,
				sheetName:       "Sheet1", // fallback name
				currentRow:      0,
				hasError:        true,
			}
		}
	}

	// Check length (Excel limit is 31 characters)
	if len(name) > 31 {
		wb.excelBuilder.AddError(fmt.Errorf("sheet name '%s' exceeds 31 character limit", name))
		return &SheetBuilder{
			workbookBuilder: wb,
			sheetName:       "Sheet1", // fallback name
			currentRow:      0,
			hasError:        true,
		}
	}

	// Check for reserved names
	reservedNames := []string{"History", "Print_Area", "Print_Titles"}
	for _, reserved := range reservedNames {
		if name == reserved {
			wb.excelBuilder.AddError(fmt.Errorf("sheet name '%s' is reserved", name))
			return &SheetBuilder{
				workbookBuilder: wb,
				sheetName:       "Sheet1", // fallback name
				currentRow:      0,
				hasError:        true,
			}
		}
	}

	// Create the sheet
	index, err := wb.file.NewSheet(name)
	if err != nil {
		wb.excelBuilder.AddError(fmt.Errorf("failed to create sheet '%s': %w", name, err))
		return &SheetBuilder{
			workbookBuilder: wb,
			sheetName:       "Sheet1", // fallback name
			currentRow:      0,
			hasError:        true,
		}
	}

	// Set as active sheet
	wb.file.SetActiveSheet(index)

	return &SheetBuilder{
		workbookBuilder: wb,
		sheetName:       name,
		currentRow:      0,
		hasError:        false,
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

// SetActiveSheet sets the active sheet of the workbook.
func (wb *WorkbookBuilder) SetActiveSheet(name string) *WorkbookBuilder {
	index, err := wb.file.GetSheetIndex(name)
	if err != nil {
		// Sheet not found, do nothing
		return wb
	}
	wb.file.SetActiveSheet(index)
	return wb
}

// Build returns the final excelize.File
func (wb *WorkbookBuilder) Build() *excelize.File {
	return wb.file
}
