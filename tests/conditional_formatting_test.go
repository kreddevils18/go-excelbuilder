package tests

// NOTE: Conditional formatting is not yet implemented in the excelbuilder package.
// These tests are commented out until the feature is implemented.

/*
import (
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// TestCellBuilder_AddConditionalFormatting_ValueBased tests value-based conditional formatting
func TestCellBuilder_AddConditionalFormatting_ValueBased(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell()

	// This should work when implemented
	conditionalFormat := excelbuilder.ConditionalFormat{
		Type:     "cell",
		Criteria: ">",
		Value:    "100",
		Format: excelbuilder.StyleConfig{
			Fill: excelbuilder.FillConfig{
				Type:  "pattern",
				Color: "FF0000", // Red background
			},
		},
	}

	result := cell.AddConditionalFormatting(conditionalFormat)

	// Should return the same cell for chaining
	assert.NotNil(t, result)
	assert.Equal(t, cell, result)

	// Build and verify
	file := workbook.Build()
	require.NotNil(t, file)
}

// TestCellBuilder_AddConditionalFormatting_ColorScale tests color scale formatting
func TestCellBuilder_AddConditionalFormatting_ColorScale(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add some test data
	for i := 1; i <= 10; i++ {
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.WithValue(i * 10) // Values 10, 20, 30, ..., 100
	}

	// Apply color scale to range A1:A10
	conditionalFormat := excelbuilder.ConditionalFormat{
		Type:  "colorScale",
		Range: "A1:A10",
		ColorScale: excelbuilder.ColorScale{
			MinColor: "FF0000", // Red for minimum
			MaxColor: "00FF00", // Green for maximum
		},
	}

	result := sheet.AddConditionalFormatting(conditionalFormat)
	assert.NotNil(t, result)
	assert.Equal(t, sheet, result)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestCellBuilder_AddConditionalFormatting_DataBars tests data bars formatting
func TestCellBuilder_AddConditionalFormatting_DataBars(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add some test data
	for i := 1; i <= 10; i++ {
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.WithValue(i * 10)
	}

	conditionalFormat := excelbuilder.ConditionalFormat{
		Type:  "dataBar",
		Range: "A1:A10",
		DataBar: excelbuilder.DataBar{
			Color:    "0066CC", // Blue bars
			ShowValue: true,
		},
	}

	result := sheet.AddConditionalFormatting(conditionalFormat)
	assert.NotNil(t, result)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestCellBuilder_AddConditionalFormatting_IconSets tests icon sets formatting
func TestCellBuilder_AddConditionalFormatting_IconSets(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add some test data
	for i := 1; i <= 10; i++ {
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.WithValue(i * 10)
	}

	conditionalFormat := excelbuilder.ConditionalFormat{
		Type:  "iconSet",
		Range: "A1:A10",
		IconSet: excelbuilder.IconSet{
			Type: "3TrafficLights1", // Traffic light icons
			ShowValue: true,
		},
	}

	result := sheet.AddConditionalFormatting(conditionalFormat)
	assert.NotNil(t, result)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestCellBuilder_AddConditionalFormatting_MultipleRules tests multiple conditional formatting rules
func TestCellBuilder_AddConditionalFormatting_MultipleRules(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell()

	// First rule: Red background for values > 100
	format1 := excelbuilder.ConditionalFormat{
		Type:     "cell",
		Criteria: ">",
		Value:    "100",
		Format: excelbuilder.StyleConfig{
			Fill: excelbuilder.FillConfig{
				Type:  "pattern",
				Color: "FF0000",
			},
		},
		Priority: 1,
	}

	// Second rule: Yellow background for values between 50-100
	format2 := excelbuilder.ConditionalFormat{
		Type:     "cell",
		Criteria: "between",
		Value:    "50",
		Value2:   "100",
		Format: excelbuilder.StyleConfig{
			Fill: excelbuilder.FillConfig{
				Type:  "pattern",
				Color: "FFFF00",
			},
		},
		Priority: 2,
	}

	result1 := cell.AddConditionalFormatting(format1)
	result2 := cell.AddConditionalFormatting(format2)

	assert.NotNil(t, result1)
	assert.NotNil(t, result2)
	assert.Equal(t, cell, result1)
	assert.Equal(t, cell, result2)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestCellBuilder_AddConditionalFormatting_InvalidInput tests invalid conditional formatting input
func TestCellBuilder_AddConditionalFormatting_InvalidInput(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell()

	// Test with empty type
	conditionalFormat := excelbuilder.ConditionalFormat{
		Type: "", // Invalid empty type
	}

	result := cell.AddConditionalFormatting(conditionalFormat)
	// Should handle gracefully
	assert.NotNil(t, result)
}

// TestSheetBuilder_AddConditionalFormatting_Range tests range-based conditional formatting
func TestSheetBuilder_AddConditionalFormatting_Range(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add test data to range
	for i := 1; i <= 5; i++ {
		row := sheet.AddRow()
		for j := 1; j <= 3; j++ {
			cell := row.AddCell()
			cell.WithValue(i * j * 10)
		}
	}

	// Apply conditional formatting to entire range
	conditionalFormat := excelbuilder.ConditionalFormat{
		Type:     "cell",
		Criteria: ">",
		Value:    "50",
		Range:    "A1:C5",
		Format: excelbuilder.StyleConfig{
			Fill: excelbuilder.FillConfig{
				Type:  "pattern",
				Color: "00FF00",
			},
		},
	}

	result := sheet.AddConditionalFormatting(conditionalFormat)
	assert.NotNil(t, result)
	assert.Equal(t, sheet, result)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestCellBuilder_AddConditionalFormatting_FluentAPI tests fluent API chaining
func TestCellBuilder_AddConditionalFormatting_FluentAPI(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()

	conditionalFormat := excelbuilder.ConditionalFormat{
		Type:     "cell",
		Criteria: ">",
		Value:    "100",
		Format: excelbuilder.StyleConfig{
			Fill: excelbuilder.FillConfig{
				Type:  "pattern",
				Color: "FF0000",
			},
		},
	// Test fluent chaining
	result := row.AddCell().
		WithValue("150").
		AddConditionalFormatting(conditionalFormat)

	assert.NotNil(t, result)
}

	file := workbook.Build()
	require.NotNil(t, file)
}
*/