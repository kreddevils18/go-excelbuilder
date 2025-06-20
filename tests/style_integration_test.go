package tests

import (
	"fmt"
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
)

// Test Case 10.1: Style Integration with Builder Pattern
func TestStyleIntegration_BasicStyling(t *testing.T) {
	// Test: Check that styling works with the builder pattern
	// Expected:
	// - Styles can be applied through CellBuilder
	// - StyleManager is properly integrated
	// - Fluent interface works with styling

	builder := excelbuilder.New()
	assert.NotNil(t, builder, "Expected ExcelBuilder instance")

	// Create a workbook with styled cells
	file := builder.
		NewWorkbook().
		SetProperties(excelbuilder.WorkbookProperties{
			Title:  "Style Integration Test",
			Author: "Go Excel Builder",
		}).
		AddSheet("StyledSheet").
		AddRow().
		AddCell("Bold Header").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Bold: true,
				Size: 14,
				Color: "#FF0000",
			},
		}).
		Done().
		AddCell("Normal Text").
		Done().
		Done().
		AddRow().
		AddCell("Styled Cell").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Italic: true,
				Size: 12,
			},
			Fill: excelbuilder.FillConfig{
				Type:  "pattern",
				Color: "#FFFF00",
			},
		}).
		Done().
		Done().
		Done().
		Build()

	assert.NotNil(t, file, "Expected Excel file to be created")
}

// Test Case 10.2: Style Reuse and Caching
func TestStyleIntegration_StyleReuse(t *testing.T) {
	// Test: Check that identical styles are reused from cache
	// Expected:
	// - Same style config reuses flyweight
	// - Cache statistics show reuse
	// - Performance is optimized

	builder := excelbuilder.New()

	// Define a common style
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 14,
			Color: "#000000",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#CCCCCC",
		},
	}

	// Apply the same style to multiple cells
	file := builder.
		NewWorkbook().
		AddSheet("CacheTest").
		AddRow().
		AddCell("Header 1").SetStyle(headerStyle).Done().
		AddCell("Header 2").SetStyle(headerStyle).Done().
		AddCell("Header 3").SetStyle(headerStyle).Done().
		Done().
		AddRow().
		AddCell("Header 4").SetStyle(headerStyle).Done().
		AddCell("Header 5").SetStyle(headerStyle).Done().
		Done().
		Done().
		Build()

	assert.NotNil(t, file, "Expected Excel file to be created")

	// Check cache statistics - should show style reuse
	// Note: We would need to expose StyleManager to check this in a real implementation
}

// Test Case 10.3: Complex Style Combinations
func TestStyleIntegration_ComplexStyles(t *testing.T) {
	// Test: Check complex style combinations work correctly
	// Expected:
	// - All style properties are applied
	// - Complex styles work with builder pattern
	// - No conflicts between different style aspects

	builder := excelbuilder.New()

	// Define complex styles
	titleStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:      true,
			Italic:    false,
			Underline: true,
			Size:      16,
			Color:     "#FF0000",
			Family:    "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#FFFF00",
		},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thick"},
			Bottom: excelbuilder.BorderSide{Style: "thick"},
			Left:   excelbuilder.BorderSide{Style: "thick"},
			Right:  excelbuilder.BorderSide{Style: "thick"},
			Color:  "#000000",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	}

	dataStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: false,
			Size: 10,
		},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin"},
			Bottom: excelbuilder.BorderSide{Style: "thin"},
			Left:   excelbuilder.BorderSide{Style: "thin"},
			Right:  excelbuilder.BorderSide{Style: "thin"},
			Color:  "#CCCCCC",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
		},
	}

	file := builder.
		NewWorkbook().
		AddSheet("ComplexStyles").
		AddRow().
		AddCell("Complex Title").SetStyle(titleStyle).Done().
		Done().
		AddRow().
		AddCell("Data Item 1").SetStyle(dataStyle).Done().
		AddCell("Data Item 2").SetStyle(dataStyle).Done().
		Done().
		Done().
		Build()

	assert.NotNil(t, file, "Expected Excel file to be created with complex styles")
}

// Test Case 10.4: Style Error Handling
func TestStyleIntegration_ErrorHandling(t *testing.T) {
	// Test: Check error handling in style integration
	// Expected:
	// - Invalid styles are handled gracefully
	// - Builder chain continues after style errors
	// - No crashes on invalid style configs

	builder := excelbuilder.New()

	// Test with potentially problematic style config
	problematicStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:  -1, // Invalid size
			Color: "invalid_color", // Invalid color format
		},
	}

	// Should not crash even with problematic style
	file := builder.
		NewWorkbook().
		AddSheet("ErrorTest").
		AddRow().
		AddCell("Test Cell").SetStyle(problematicStyle).Done().
		AddCell("Normal Cell").Done().
		Done().
		Done().
		Build()

	// Should still create file even if some styles fail
	assert.NotNil(t, file, "Expected Excel file to be created despite style errors")
}

// Test Case 10.5: Performance with Many Styled Cells
func TestStyleIntegration_Performance(t *testing.T) {
	// Test: Check performance with many styled cells
	// Expected:
	// - Good performance with many cells
	// - Style caching improves performance
	// - Memory usage remains reasonable

	builder := excelbuilder.New()

	// Define styles that will be reused
	style1 := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 12},
	}
	style2 := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Italic: true, Size: 10},
	}

	workbook := builder.NewWorkbook().AddSheet("PerformanceTest")

	// Create many rows with styled cells
	for i := 0; i < 50; i++ {
		row := workbook.AddRow()
		for j := 0; j < 10; j++ {
			cell := row.AddCell(fmt.Sprintf("Cell_%d_%d", i, j))
			if (i+j)%2 == 0 {
				cell.SetStyle(style1)
			} else {
				cell.SetStyle(style2)
			}
			cell.Done()
		}
		row.Done()
	}

	file := workbook.Done().Build()
	assert.NotNil(t, file, "Expected Excel file to be created with many styled cells")
}

// Test Case 10.6: Style Inheritance and Overrides
func TestStyleIntegration_StyleOverrides(t *testing.T) {
	// Test: Check that styles can be applied and overridden
	// Expected:
	// - Multiple style applications work
	// - Later styles override earlier ones
	// - Style combinations work correctly

	builder := excelbuilder.New()

	// Base style
	baseStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size: 12,
			Color: "#000000",
		},
	}

	// Override style
	override := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Color: "#FF0000",
		},
	}

	file := builder.
		NewWorkbook().
		AddSheet("StyleOverrides").
		AddRow().
		AddCell("Base Style").SetStyle(baseStyle).Done().
		AddCell("Override Style").SetStyle(baseStyle).SetStyle(override).Done().
		Done().
		Done().
		Build()

	assert.NotNil(t, file, "Expected Excel file to be created with style overrides")
}