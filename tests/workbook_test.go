package tests

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Test Case 2.1: SetProperties Method
func TestWorkbookBuilder_SetProperties(t *testing.T) {
	// Test: Check workbook properties setting
	// Input: WorkbookProperties with title, author, subject
	// Expected:
	// - Properties are set in excelize.File
	// - Method returns WorkbookBuilder (fluent interface)
	// - Can chain with other methods

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	properties := excelbuilder.WorkbookProperties{
		Title:       "Test Workbook",
		Author:      "Test Author",
		Subject:     "Test Subject",
		Description: "Test Description",
		Company:     "Test Company",
	}

	result := workbook.SetProperties(properties)

	// Check fluent interface
	if result == nil {
		t.Fatal("Expected SetProperties to return WorkbookBuilder for fluent interface")
	}

	// Check if we can chain methods
	sheet := result.AddSheet("TestSheet")
	if sheet == nil {
		t.Fatal("Expected to chain AddSheet method after SetProperties")
	}
}

// Test Case 2.2: AddSheet Method
func TestWorkbookBuilder_AddSheet(t *testing.T) {
	// Test: Check adding new sheet
	// Input: Sheet name "TestSheet"
	// Expected:
	// - Sheet is created in excelize.File
	// - Returns SheetBuilder instance
	// - SheetBuilder has reference to sheet name

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	sheetName := "TestSheet"
	sheet := workbook.AddSheet(sheetName)

	if sheet == nil {
		t.Fatal("Expected AddSheet to return SheetBuilder instance")
	}

	// Test that we can continue building the sheet
	row := sheet.AddRow()
	if row == nil {
		t.Fatal("Expected to be able to add row to the sheet")
	}
}

// Test Case 2.3: AddSheet with Invalid Name
func TestWorkbookBuilder_AddSheet_InvalidName(t *testing.T) {
	// Test: Check error handling with invalid sheet name
	// Input: Empty string or special characters
	// Expected:
	// - Returns error with clear message
	// - Does not create sheet in file

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Test empty string
	sheet := workbook.AddSheet("")
	if sheet != nil {
		t.Error("Expected AddSheet with empty name to return nil")
	}

	// Test invalid characters (Excel doesn't allow certain characters in sheet names)
	invalidNames := []string{
		"Sheet[1]",
		"Sheet*1",
		"Sheet?1",
		"Sheet/1",
		"Sheet\\1",
	}

	for _, invalidName := range invalidNames {
		sheet := workbook.AddSheet(invalidName)
		if sheet != nil {
			t.Errorf("Expected AddSheet with invalid name '%s' to return nil", invalidName)
		}
	}
}

// Test Case 2.4: Build Method
func TestWorkbookBuilder_Build(t *testing.T) {
	// Test: Check building final Excel file
	// Expected:
	// - Returns excelize.File instance
	// - File can be saved
	// - No error

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Add some content to make it a valid workbook
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:  "Test Workbook",
		Author: "Test Author",
	})

	sheet := workbook.AddSheet("Sheet1")
	if sheet == nil {
		t.Fatal("Failed to add sheet")
	}

	// Go back to workbook and build
	workbookBuilder := sheet.Done()
	file := workbookBuilder.Build()

	if file == nil {
		t.Fatal("Expected Build to return excelize.File instance")
	}

	// Test that file can be saved (we won't actually save to avoid file system side effects)
	// This is just checking the file structure is valid
	sheets := file.GetSheetList()
	if len(sheets) == 0 {
		t.Error("Expected at least one sheet in the workbook")
	}

	found := false
	for _, sheetName := range sheets {
		if sheetName == "Sheet1" {
			found = true
			break
		}
	}

	if !found {
		t.Error("Expected to find 'Sheet1' in the workbook")
	}
}
