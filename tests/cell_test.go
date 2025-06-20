package tests

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Test Case 5.1: SetStyle Method
func TestCellBuilder_SetStyle(t *testing.T) {
	// Test: Check applying style to cell
	// Input: StyleConfig with font, fill, border
	// Expected:
	// - Style is applied to cell in excelize
	// - Method returns CellBuilder (fluent)

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell("Styled Cell")

	if cell == nil {
		t.Fatal("Failed to create cell")
	}

	style := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Italic: false,
			Size:   14,
			Color:  "#FF0000",
			Family: "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#FFFF00",
		},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin"},
			Bottom: excelbuilder.BorderSide{Style: "thin"},
			Left:   excelbuilder.BorderSide{Style: "thin"},
			Right:  excelbuilder.BorderSide{Style: "thin"},
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	}

	result := cell.SetStyle(style)
	if result == nil {
		t.Fatal("Expected SetStyle to return CellBuilder for fluent interface")
	}

	// Test chaining
	formattedCell := result.SetNumberFormat("0.00")
	if formattedCell == nil {
		t.Fatal("Expected to chain SetNumberFormat after SetStyle")
	}
}

// Test Case 5.2: SetNumberFormat Method
func TestCellBuilder_SetNumberFormat(t *testing.T) {
	// Test: Check setting number format
	// Input: Various number formats
	// Expected:
	// - Number format is applied in excelize
	// - Method returns CellBuilder (fluent)

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()

	// Test different number formats
	testCases := []struct {
		value  interface{}
		format string
		desc   string
	}{
		{123.456, "0.00", "decimal format"},
		{0.75, "0.00%", "percentage format"},
		{1234567, "#,##0", "thousands separator"},
		{123.456, "$#,##0.00", "currency format"},
	}

	for _, tc := range testCases {
		cell := row.AddCell(tc.value)
		if cell == nil {
			t.Fatalf("Failed to create cell for %s", tc.desc)
		}

		result := cell.SetNumberFormat(tc.format)
		if result == nil {
			t.Fatalf("Expected SetNumberFormat to return CellBuilder for %s", tc.desc)
		}
	}
}

// Test Case 5.3: SetFormula Method
func TestCellBuilder_SetFormula(t *testing.T) {
	// Test: Check setting formula on existing cell
	// Input: "=SUM(A1:A10)"
	// Expected:
	// - Formula replaces cell value
	// - Method returns CellBuilder (fluent)

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()

	// Create cell with initial value
	cell := row.AddCell("Initial Value")
	if cell == nil {
		t.Fatal("Failed to create cell")
	}

	// Set formula
	formula := "=SUM(A1:A10)"
	result := cell.SetFormula(formula)
	if result == nil {
		t.Fatal("Expected SetFormula to return CellBuilder for fluent interface")
	}

	// Test chaining with style
	styledCell := result.SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
		},
	})

	if styledCell == nil {
		t.Fatal("Expected to chain SetStyle after SetFormula")
	}
}

// Test Case 5.4: SetHyperlink Method
func TestCellBuilder_SetHyperlink(t *testing.T) {
	// Test: Check setting hyperlink on cell
	// Input: URL and display text
	// Expected:
	// - Hyperlink is set in excelize
	// - Method returns CellBuilder (fluent)

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell("Click here")

	if cell == nil {
		t.Fatal("Failed to create cell")
	}

	url := "https://example.com"
	result := cell.SetHyperlink(url)
	if result == nil {
		t.Fatal("Expected SetHyperlink to return CellBuilder for fluent interface")
	}

	// Test chaining with style (hyperlinks often have special styling)
	styledCell := result.SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Color:     "#0000FF",
			Underline: true,
		},
	})

	if styledCell == nil {
		t.Fatal("Expected to chain SetStyle after SetHyperlink")
	}
}

// Test Case 5.5: Done Method
func TestCellBuilder_Done(t *testing.T) {
	// Test: Check returning to RowBuilder
	// Expected:
	// - Returns RowBuilder instance
	// - Can continue with row operations

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell("Test Cell")

	if cell == nil {
		t.Fatal("Failed to create cell")
	}

	// Apply some styling
	cell.SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
		},
	})

	// Go back to row
	rowBuilder := cell.Done()
	if rowBuilder == nil {
		t.Fatal("Expected Done to return RowBuilder instance")
	}

	// Test that we can add another cell
	anotherCell := rowBuilder.AddCell("Another Cell")
	if anotherCell == nil {
		t.Fatal("Expected to be able to add another cell after Done")
	}
}

// Test Case 5.6: Complex Style Combinations
func TestCellBuilder_ComplexStyling(t *testing.T) {
	// Test: Check applying multiple styles in sequence
	// Expected:
	// - All styles are applied correctly
	// - Fluent interface works throughout

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell(123.456)

	if cell == nil {
		t.Fatal("Failed to create cell")
	}

	// Chain multiple styling operations
	result := cell.
		SetNumberFormat("$#,##0.00").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Bold:   true,
				Size:   12,
				Color:  "#000080",
				Family: "Arial",
			},
			Fill: excelbuilder.FillConfig{
				Type:  "pattern",
				Color: "#E0E0E0",
			},
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: "right",
				Vertical:   "middle",
			},
		})

	if result == nil {
		t.Fatal("Expected complex styling chain to work")
	}

	// Verify we can still continue
	rowBuilder := result.Done()
	if rowBuilder == nil {
		t.Fatal("Expected to return to RowBuilder after complex styling")
	}
}

// Test Case 5.7: Error Handling - Invalid Formats
func TestCellBuilder_InvalidFormats(t *testing.T) {
	// Test: Check error handling with invalid formats
	// Expected:
	// - Invalid formats are handled gracefully
	// - Methods still return CellBuilder or handle errors appropriately

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell("Test")

	if cell == nil {
		t.Fatal("Failed to create cell")
	}

	// Test invalid number format (should handle gracefully)
	result := cell.SetNumberFormat("invalid_format")
	// Depending on implementation, this might return nil or handle gracefully
	// The test should verify the expected behavior based on design decisions

	// Test invalid hyperlink
	result2 := cell.SetHyperlink("not_a_valid_url")
	// Again, verify expected behavior

	// The exact assertions here depend on how we decide to handle errors
	// For now, we'll just verify the methods don't panic
	_ = result
	_ = result2
}
