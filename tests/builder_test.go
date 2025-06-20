package tests

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Test Case 1.1: Constructor and Initialization
func TestExcelBuilder_New(t *testing.T) {
	// Test: Check ExcelBuilder instance creation
	// Expected:
	// - ExcelBuilder created successfully
	// - Internal excelize.File initialized
	// - No error
	// - StyleManager initialized

	builder := excelbuilder.New()

	if builder == nil {
		t.Fatal("Expected ExcelBuilder instance, got nil")
	}

	// Check if builder can create workbook
	workbook := builder.NewWorkbook()
	if workbook == nil {
		t.Fatal("Expected WorkbookBuilder instance, got nil")
	}
}

// Test Case 1.2: NewWorkbook Method
func TestExcelBuilder_NewWorkbook(t *testing.T) {
	// Test: Check WorkbookBuilder creation
	// Expected:
	// - Returns WorkbookBuilder instance
	// - WorkbookBuilder has reference to ExcelBuilder
	// - Can chain methods

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	if workbook == nil {
		t.Fatal("Expected WorkbookBuilder instance, got nil")
	}

	// Test fluent interface - can chain methods
	result := workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:  "Test Workbook",
		Author: "Test Author",
	})

	if result == nil {
		t.Fatal("Expected fluent interface to return WorkbookBuilder")
	}
}
