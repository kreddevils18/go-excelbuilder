package tests

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Test Case 3.1: AddRow Method
func TestSheetBuilder_AddRow(t *testing.T) {
	// Test: Check adding new row
	// Expected:
	// - Returns RowBuilder instance
	// - RowBuilder tracks correct row index
	// - Can chain further

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	if sheet == nil {
		t.Fatal("Failed to create sheet")
	}

	row := sheet.AddRow()
	if row == nil {
		t.Fatal("Expected AddRow to return RowBuilder instance")
	}

	// Test that we can add cells to the row
	cell := row.AddCell("Test Value")
	if cell == nil {
		t.Fatal("Expected to be able to add cell to the row")
	}
}

// Test Case 3.2: SetColumnWidth Method
func TestSheetBuilder_SetColumnWidth(t *testing.T) {
	// Test: Check setting column width
	// Input: Column "A", width 20.5
	// Expected:
	// - Column width is set in excelize
	// - Method returns SheetBuilder (fluent)

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	if sheet == nil {
		t.Fatal("Failed to create sheet")
	}

	result := sheet.SetColumnWidth("A", 20.5)
	if result == nil {
		t.Fatal("Expected SetColumnWidth to return SheetBuilder for fluent interface")
	}

	// Test chaining
	row := result.AddRow()
	if row == nil {
		t.Fatal("Expected to chain AddRow after SetColumnWidth")
	}
}

// Test Case 3.3: SetColumnWidth with Multiple Columns
func TestSheetBuilder_SetColumnWidth_MultipleColumns(t *testing.T) {
	// Test: Check setting width for multiple columns
	// Input: Various column references and widths
	// Expected:
	// - All column widths are set correctly
	// - Fluent interface works for chaining

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	if sheet == nil {
		t.Fatal("Failed to create sheet")
	}

	// Set multiple column widths
	result := sheet.
		SetColumnWidth("A", 15.0).
		SetColumnWidth("B", 25.0).
		SetColumnWidth("C", 30.0)

	if result == nil {
		t.Fatal("Expected fluent chaining of SetColumnWidth calls")
	}

	// Verify we can still add rows
	row := result.AddRow()
	if row == nil {
		t.Fatal("Expected to add row after setting column widths")
	}
}

// Test Case 3.4: Done Method
func TestSheetBuilder_Done(t *testing.T) {
	// Test: Check returning to WorkbookBuilder
	// Expected:
	// - Returns WorkbookBuilder instance
	// - Can continue chaining with workbook methods

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	if sheet == nil {
		t.Fatal("Failed to create sheet")
	}

	// Add some content to the sheet
	row := sheet.AddRow()
	if row == nil {
		t.Fatal("Failed to add row")
	}

	cell := row.AddCell("Test")
	if cell == nil {
		t.Fatal("Failed to add cell")
	}

	// Go back to row, then sheet, then workbook
	rowBuilder := cell.Done()
	sheetBuilder := rowBuilder.Done()
	workbookBuilder := sheetBuilder.Done()

	if workbookBuilder == nil {
		t.Fatal("Expected Done to return WorkbookBuilder instance")
	}

	// Test that we can continue with workbook operations
	anotherSheet := workbookBuilder.AddSheet("AnotherSheet")
	if anotherSheet == nil {
		t.Fatal("Expected to be able to add another sheet after Done")
	}
}

// Test Case 3.5: MergeCell Method (if implemented)
func TestSheetBuilder_MergeCell(t *testing.T) {
	// Test: Check cell merging functionality
	// Input: Cell range "A1:C1"
	// Expected:
	// - Cells are merged in excelize
	// - Method returns SheetBuilder (fluent)

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	if sheet == nil {
		t.Fatal("Failed to create sheet")
	}

	// Add some content first
	row := sheet.AddRow()
	if row == nil {
		t.Fatal("Failed to add row")
	}

	cell := row.AddCell("Merged Cell")
	if cell == nil {
		t.Fatal("Failed to add cell")
	}

	// Go back to sheet and merge cells
	sheetBuilder := cell.Done().Done()
	result := sheetBuilder.MergeCell("A1:C1")

	if result == nil {
		t.Fatal("Expected MergeCell to return SheetBuilder for fluent interface")
	}

	// Test chaining
	anotherRow := result.AddRow()
	if anotherRow == nil {
		t.Fatal("Expected to chain AddRow after MergeCell")
	}
}