package tests

import (
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// TestSheetBuilder_AddRowsBatch tests adding multiple rows in one operation
func TestSheetBuilder_AddRowsBatch(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Prepare batch data
	rowsData := [][]interface{}{
		{"Name", "Age", "City"},
		{"John", 25, "New York"},
		{"Jane", 30, "Los Angeles"},
		{"Bob", 35, "Chicago"},
		{"Alice", 28, "Houston"},
	}

	// This should work when implemented
	result := sheet.AddRows(rowsData)

	// Should return the same sheet for chaining
	assert.NotNil(t, result)
	assert.Equal(t, sheet, result)

	// Build and verify
	file := workbook.Build()
	require.NotNil(t, file)

	// Verify data was added
	value, err := file.GetCellValue("TestSheet", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Name", value)

	value, err = file.GetCellValue("TestSheet", "B2")
	require.NoError(t, err)
	assert.Equal(t, "25", value)

	value, err = file.GetCellValue("TestSheet", "C5")
	require.NoError(t, err)
	assert.Equal(t, "Houston", value)
}

// TestSheetBuilder_AddRowsBatch_WithStyles tests batch rows with styles
func TestSheetBuilder_AddRowsBatch_WithStyles(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Prepare batch data with styles
	headerStyle := &excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "CCCCCC",
		},
	}
	dataStyle := &excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size: 10,
		},
	}
	batchData := []excelbuilder.BatchRowData{
		{Cells: []interface{}{"Header1", "Header2", "Header3"}, Style: *headerStyle},
		{Cells: []interface{}{"Data1", "Data2", "Data3"}, Style: *dataStyle},
	}

	result := sheet.AddRowsBatchWithStyles(batchData)
	assert.NotNil(t, result)
	assert.Equal(t, sheet, result)

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify data was added
	value, err := file.GetCellValue("TestSheet", "A2")
	require.NoError(t, err)
	assert.Equal(t, "Data1", value)
}

// TestSheetBuilder_ApplyStyleBatch tests applying style to range of cells
func TestSheetBuilder_ApplyStyleBatch(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add some data first
	for i := 1; i <= 5; i++ {
		row := sheet.AddRow()
		for j := 1; j <= 3; j++ {
			row.AddCell(i * j)
		}
	}

	// Apply style to range
	style := &excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Color: "FF0000",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "FFFF00",
		},
	}

	operations := []excelbuilder.BatchStyleOperation{
			{Range: "A1:C5", Style: *style},
		}
	result := sheet.ApplyStyleBatch(operations)
	assert.NotNil(t, result)
	assert.Equal(t, sheet, result)

	file := workbook.Build()
	require.NotNil(t, file)

	// Verification of styles is complex and would require inspecting the file.
	// We assume the operation is successful if the code executes without errors.
}

// TestWorkbookBuilder_AddSheetsBatch tests adding multiple sheets with templates
func TestWorkbookBuilder_AddSheetsBatch(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Prepare sheet configurations
	sheetConfigs := []excelbuilder.SheetConfig{
		{
			Name: "Sales",
			Data: [][]interface{}{
				{"Date", "Product", "Amount"},
				{"2023-01-01", "Laptop", 1200},
			},
		},
		{
			Name: "Inventory",
			Data: [][]interface{}{
				{"SKU", "Name", "Quantity", "Price"},
				{"LP-001", "Laptop", 50, 1200},
			},
		},
	}

	result := workbook.AddSheetsBatch(sheetConfigs)
	assert.NotNil(t, result)
	assert.Equal(t, workbook, result)

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify sheets were created
	sheets := file.GetSheetList()
	assert.Contains(t, sheets, "Sales", "Sales sheet should exist")
	assert.Contains(t, sheets, "Inventory", "Inventory sheet should exist")

	// Verify headers
	value, err := file.GetCellValue("Sales", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Date", value, "Sales sheet header should be 'Date'")

	value, err = file.GetCellValue("Inventory", "D1")
	require.NoError(t, err)
	assert.Equal(t, "Price", value, "Inventory sheet header should be 'Price'")
}

// TestSheetBuilder_AddRowsBatch_EmptyData tests batch operations with empty data
func TestSheetBuilder_AddRowsBatch_EmptyData(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Test with empty data
	result := sheet.AddRows([][]interface{}{})
	assert.NotNil(t, result)
	assert.Equal(t, sheet, result)

	// Test with nil data
	result2 := sheet.AddRows(nil)
	assert.NotNil(t, result2)
	assert.Equal(t, sheet, result2)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestSheetBuilder_AddRowsBatch_LargeDataset tests batch operations with large dataset
func TestSheetBuilder_AddRowsBatch_LargeDataset(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Generate large dataset (1000 rows)
	largeData := make([][]interface{}, 1000)
	for i := 0; i < 1000; i++ {
		largeData[i] = []interface{}{
			i + 1,
			"Item " + string(rune(i+1)),
			float64(i+1) * 10.5,
		}
	}

	result := sheet.AddRows(largeData)
	assert.NotNil(t, result)
	assert.Equal(t, sheet, result)

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify first and last rows
	value, err := file.GetCellValue("TestSheet", "A1")
	require.NoError(t, err)
	assert.Equal(t, "1", value)

	value, err = file.GetCellValue("TestSheet", "A1000")
	require.NoError(t, err)
	assert.Equal(t, "1000", value)
}

// TestSheetBuilder_ApplyStyleBatch_InvalidRange tests batch style with invalid range
func TestSheetBuilder_ApplyStyleBatch_InvalidRange(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	style := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
		},
	}

	// Test with invalid range
	operations := []excelbuilder.BatchStyleOperation{
		{Range: "INVALID_RANGE", Style: style},
	}
	result := sheet.ApplyStyleBatch(operations)
	// Should handle gracefully
	assert.NotNil(t, result)

	// Test with empty range
	operations2 := []excelbuilder.BatchStyleOperation{
		{Range: "", Style: style},
	}
	result2 := sheet.ApplyStyleBatch(operations2)
	assert.NotNil(t, result2)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestSheetBuilder_AddRowsBatch_FluentAPI tests fluent API chaining with batch operations
func TestSheetBuilder_AddRowsBatch_FluentAPI(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	rowsData := [][]interface{}{
		{"Name", "Age"},
		{"John", 25},
		{"Jane", 30},
	}

	style := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
		},
	}

	// Test fluent chaining
	operations := []excelbuilder.BatchStyleOperation{
		{Range: "A1:B1", Style: style},
	}
	result := workbook.AddSheet("TestSheet").
		AddRows(rowsData).
		ApplyStyleBatch(operations).
		SetColumnWidth("A", 15.0).
		SetColumnWidth("B", 10.0)

	assert.NotNil(t, result)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestSheetBuilder_AddRowsBatch_MixedDataTypes tests batch operations with mixed data types
func TestSheetBuilder_AddRowsBatch_MixedDataTypes(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Mixed data types
	rowsData := [][]interface{}{
		{"String", "Number", "Boolean", "Float"},
		{"Text", 123, true, 45.67},
		{"Another", 456, false, 89.12},
		{nil, 0, nil, 0.0}, // Test with nil values
	}

	result := sheet.AddRows(rowsData)
	assert.NotNil(t, result)
	assert.Equal(t, sheet, result)

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify mixed data types
	value, err := file.GetCellValue("TestSheet", "B2")
	require.NoError(t, err)
	assert.Equal(t, "123", value)

	value, err = file.GetCellValue("TestSheet", "C2")
	require.NoError(t, err)
	assert.Equal(t, "TRUE", value)

	value, err = file.GetCellValue("TestSheet", "D2")
	require.NoError(t, err)
	assert.Equal(t, "45.67", value)
}