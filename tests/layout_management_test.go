package tests

import (
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// TestSheetBuilder_SetColumnWidth_ValidInput tests setting width for a specific column
func TestSheetBuilder_SetColumnWidth_ValidInput(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// This should work when implemented
	result := sheet.SetColumnWidth("A", 20.5)

	// Should return the same sheet for chaining
	assert.NotNil(t, result)
	assert.Equal(t, sheet, result)

	// Build and verify
	file := workbook.Build()

	// Verify column width was set
	width, err := file.GetColWidth("TestSheet", "A")
	require.NoError(t, err)
	assert.Equal(t, 20.5, width)
}

// TestSheetBuilder_SetColumnWidth_LayoutManagement tests setting width for multiple columns in layout context
func TestSheetBuilder_SetColumnWidth_LayoutManagement(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Set different widths for different columns
	sheet.SetColumnWidth("A", 15.0)
	sheet.SetColumnWidth("B", 25.0)
	sheet.SetColumnWidth("D", 30.0) // Skip C intentionally

	file := workbook.Build()

	// Verify each column width
	widthA, err := file.GetColWidth("TestSheet", "A")
	require.NoError(t, err)
	assert.Equal(t, 15.0, widthA)

	widthB, err := file.GetColWidth("TestSheet", "B")
	require.NoError(t, err)
	assert.Equal(t, 25.0, widthB)

	widthD, err := file.GetColWidth("TestSheet", "D")
	require.NoError(t, err)
	assert.Equal(t, 30.0, widthD)

	// Column C should have default width
	widthC, err := file.GetColWidth("TestSheet", "C")
	require.NoError(t, err)
	// Default width in Excel is typically around 8.43
	assert.True(t, widthC < 15.0, "Column C should have default width")
}

// TestSheetBuilder_SetColumnWidth_InvalidInput tests invalid input handling
func TestSheetBuilder_SetColumnWidth_InvalidInput(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Test negative width - should return error or handle gracefully
	result := sheet.SetColumnWidth("A", -5.0)
	assert.Nil(t, result, "Negative width should return nil")

	// Test extremely large width
	result = sheet.SetColumnWidth("B", 1000.0)
	assert.Nil(t, result, "Extremely large width should return nil")

	// Test invalid column name
	result = sheet.SetColumnWidth("AA1", 20.0)
	assert.Nil(t, result, "Invalid column name should return nil")

	// Test empty column name
	result = sheet.SetColumnWidth("", 20.0)
	assert.Nil(t, result, "Empty column name should return nil")
}

// TestSheetBuilder_SetColumnWidth_FluentAPI tests fluent API chaining
func TestSheetBuilder_SetColumnWidth_FluentAPI(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Test chaining
	result := sheet.SetColumnWidth("A", 20.0).SetColumnWidth("B", 25.0).AddRow()
	assert.NotNil(t, result, "Should be able to chain operations")

	file := workbook.Build()

	// Verify both columns were set
	widthA, err := file.GetColWidth("TestSheet", "A")
	require.NoError(t, err)
	assert.Equal(t, 20.0, widthA)

	widthB, err := file.GetColWidth("TestSheet", "B")
	require.NoError(t, err)
	assert.Equal(t, 25.0, widthB)
}

// TestRowBuilder_SetHeight_ValidInput tests setting height for current row
func TestRowBuilder_SetHeight_ValidInput(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()

	// Set height for current row
	result := row.SetHeight(30.0)
	assert.NotNil(t, result)
	assert.Equal(t, row, result)

	file := workbook.Build()

	// Verify row height was set
	height, err := file.GetRowHeight("TestSheet", 1)
	require.NoError(t, err)
	assert.Equal(t, 30.0, height)
}

// TestSheetBuilder_SetRowHeight_SpecificRow tests setting height for specific row
func TestSheetBuilder_SetRowHeight_SpecificRow(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Set height for specific row
	result := sheet.SetRowHeight(5, 25.5)
	assert.NotNil(t, result)

	file := workbook.Build()

	// Verify row 5 has the specified height
	height, err := file.GetRowHeight("TestSheet", 5)
	require.NoError(t, err)
	assert.Equal(t, 25.5, height)

	// Other rows should have default height
	height1, err := file.GetRowHeight("TestSheet", 1)
	require.NoError(t, err)
	assert.True(t, height1 < 25.5, "Row 1 should have default height")
}

// TestRowBuilder_SetHeight_InvalidInput tests invalid height input
func TestRowBuilder_SetHeight_InvalidInput(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()

	// Test negative height
	result := row.SetHeight(-5.0)
	assert.Nil(t, result, "Negative height should return nil")

	// Test zero height
	result = row.SetHeight(0.0)
	assert.Nil(t, result, "Zero height should return nil")

	// Test extremely large height
	result = row.SetHeight(1000.0)
	assert.Nil(t, result, "Extremely large height should return nil")
}

// TestSheetBuilder_MergeCells_BasicRange tests basic cell merging
func TestSheetBuilder_MergeCells_BasicRange(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add some content first
	sheet.AddRow().AddCell("Merged Content").Done().AddCell("Cell B").Done().AddCell("Cell C")

	// Merge cells A1:C1
	result := sheet.MergeCell("A1:C1")
	assert.NotNil(t, result)

	file := workbook.Build()

	// Verify merge was applied
	mergedCells, err := file.GetMergeCells("TestSheet")
	require.NoError(t, err)
	assert.Len(t, mergedCells, 1)
	assert.Equal(t, "A1:C1", mergedCells[0].GetStartAxis()+":"+mergedCells[0].GetEndAxis())

	// Verify content is in A1
	value, err := file.GetCellValue("TestSheet", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Merged Content", value)
}

// TestSheetBuilder_MergeCells_RectangleRange tests merging rectangular range
func TestSheetBuilder_MergeCells_RectangleRange(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add content in a 3x3 grid
	for i := 0; i < 3; i++ {
		row := sheet.AddRow()
		for j := 0; j < 3; j++ {
			row.AddCell("Content")
		}
	}

	// Merge B2:D4 (3x3 rectangle)
	result := sheet.MergeCell("B2:D4")
	assert.NotNil(t, result)

	file := workbook.Build()

	// Verify merge was applied
	mergedCells, err := file.GetMergeCells("TestSheet")
	require.NoError(t, err)
	assert.Len(t, mergedCells, 1)
}

// TestSheetBuilder_MergeCells_InvalidRange tests invalid merge ranges
func TestSheetBuilder_MergeCells_InvalidRange(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Test single cell (should be no-op or return error)
	result := sheet.MergeCell("A1:A1")
	assert.NotNil(t, result, "Single cell merge should be handled gracefully")

	// Test reverse range
	result = sheet.MergeCell("D1:A1")
	assert.Nil(t, result, "Reverse range should return nil")

	// Test invalid format
	result = sheet.MergeCell("XYZ")
	assert.Nil(t, result, "Invalid format should return nil")

	// Test empty range
	result = sheet.MergeCell("")
	assert.Nil(t, result, "Empty range should return nil")
}
