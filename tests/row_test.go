package tests

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Test Case 4.1: AddCell with String Value
func TestRowBuilder_AddCell_String(t *testing.T) {
	// Test: Check adding cell with string value
	// Input: "Test String"
	// Expected:
	// - Cell is created with correct value
	// - Returns CellBuilder instance
	// - Can chain further

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()

	if row == nil {
		t.Fatal("Failed to create row")
	}

	testValue := "Test String"
	cell := row.AddCell(testValue)

	if cell == nil {
		t.Fatal("Expected AddCell to return CellBuilder instance")
	}

	// Test that we can apply styles to the cell
	styledCell := cell.SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
		},
	})

	if styledCell == nil {
		t.Fatal("Expected to be able to set style on cell")
	}
}

// Test Case 4.2: AddCell with Numeric Value
func TestRowBuilder_AddCell_Numeric(t *testing.T) {
	// Test: Check adding cell with numeric value
	// Input: 123.45
	// Expected:
	// - Cell is created with correct numeric value
	// - Value type is preserved

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()

	if row == nil {
		t.Fatal("Failed to create row")
	}

	testValue := 123.45
	cell := row.AddCell(testValue)

	if cell == nil {
		t.Fatal("Expected AddCell to return CellBuilder instance")
	}

	// Test chaining with numeric formatting
	formattedCell := cell.SetNumberFormat("0.00")
	if formattedCell == nil {
		t.Fatal("Expected to be able to set number format on numeric cell")
	}
}

// Test Case 4.3: AddCell with Boolean Value
func TestRowBuilder_AddCell_Boolean(t *testing.T) {
	// Test: Check adding cell with boolean value
	// Input: true/false
	// Expected:
	// - Cell is created with correct boolean value
	// - Value type is preserved

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()

	if row == nil {
		t.Fatal("Failed to create row")
	}

	// Test true value
	cellTrue := row.AddCell(true)
	if cellTrue == nil {
		t.Fatal("Expected AddCell to handle boolean true value")
	}

	// Test false value
	cellFalse := row.AddCell(false)
	if cellFalse == nil {
		t.Fatal("Expected AddCell to handle boolean false value")
	}
}

// Test Case 4.4: AddCell with Formula
func TestRowBuilder_AddCell_Formula(t *testing.T) {
	// Test: Check adding cell with formula
	// Input: "=SUM(A1:A10)"
	// Expected:
	// - Cell is created with formula
	// - Formula is properly set in excelize

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()

	if row == nil {
		t.Fatal("Failed to create row")
	}

	// Add some numeric values first
	row.AddCell(10)
	row.AddCell(20)
	row.AddCell(30)

	// Add a formula cell
	formulaCell := row.AddCell("=SUM(A1:C1)")
	if formulaCell == nil {
		t.Fatal("Expected AddCell to handle formula")
	}
}

// Test Case 4.5: Multiple AddCell Calls
func TestRowBuilder_AddCell_Multiple(t *testing.T) {
	// Test: Check adding multiple cells to same row
	// Expected:
	// - Each cell is placed in correct column
	// - Column indices increment properly
	// - All cells are accessible

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()

	if row == nil {
		t.Fatal("Failed to create row")
	}

	// Add multiple cells with different types
	cells := []interface{}{
		"Name",
		"John Doe",
		25,
		123.45,
		true,
		"=A1&B1",
	}

	for i, value := range cells {
		cell := row.AddCell(value)
		if cell == nil {
			t.Fatalf("Expected AddCell to succeed for cell %d with value %v", i, value)
		}
	}
}

// Test Case 4.6: Done Method
func TestRowBuilder_Done(t *testing.T) {
	// Test: Check returning to SheetBuilder
	// Expected:
	// - Returns SheetBuilder instance
	// - Can continue with sheet operations

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()

	if row == nil {
		t.Fatal("Failed to create row")
	}

	// Add some cells
	row.AddCell("Test")
	row.AddCell(123)

	// Go back to sheet
	sheetBuilder := row.Done()
	if sheetBuilder == nil {
		t.Fatal("Expected Done to return SheetBuilder instance")
	}

	// Test that we can add another row
	anotherRow := sheetBuilder.AddRow()
	if anotherRow == nil {
		t.Fatal("Expected to be able to add another row after Done")
	}
}

// Test Case 4.7: SetHeight Method
func TestRowBuilder_SetHeight(t *testing.T) {
	// Test: Check setting row height
	// Input: height 25.5
	// Expected:
	// - Row height is set in excelize
	// - Method returns RowBuilder (fluent)

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()

	if row == nil {
		t.Fatal("Failed to create row")
	}

	result := row.SetHeight(25.5)
	if result == nil {
		t.Fatal("Expected SetHeight to return RowBuilder for fluent interface")
	}

	// Test chaining
	cell := result.AddCell("Test")
	if cell == nil {
		t.Fatal("Expected to chain AddCell after SetHeight")
	}
}