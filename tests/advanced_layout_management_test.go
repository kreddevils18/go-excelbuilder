package tests

import (
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// TestAdvancedLayoutManager_ColumnGrouping tests column grouping functionality
func TestAdvancedLayoutManager_ColumnGrouping(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add some data first
	sheet.AddRow().AddCells("A", "B", "C", "D", "E")
	sheet.AddRow().AddCells(1, 2, 3, 4, 5)

	// Group columns B:D
	layoutManager := sheet.GetLayoutManager()
	result := layoutManager.GroupColumns("B:D", 1) // Level 1 grouping
	assert.NotNil(t, result)

	file := workbook.Build()

	// Verify grouping was applied
	// Note: excelize doesn't have direct method to check grouping,
	// but we can verify the structure was created without errors
	assert.NotNil(t, file)
}

// TestAdvancedLayoutManager_RowGrouping tests row grouping functionality
func TestAdvancedLayoutManager_RowGrouping(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add multiple rows
	for i := 1; i <= 10; i++ {
		sheet.AddRow().AddCell(i)
	}

	// Group rows 3:7
	layoutManager := sheet.GetLayoutManager()
	result := layoutManager.GroupRows(3, 7, 1) // Level 1 grouping
	assert.NotNil(t, result)

	file := workbook.Build()
	assert.NotNil(t, file)
}

// TestAdvancedLayoutManager_NestedGrouping tests nested grouping functionality
func TestAdvancedLayoutManager_NestedGrouping(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add data
	for i := 1; i <= 10; i++ {
		row := sheet.AddRow()
		for j := 1; j <= 10; j++ {
			row.AddCell(i*10 + j)
		}
	}

	layoutManager := sheet.GetLayoutManager()

	// Create nested column grouping
	layoutManager.GroupColumns("A:J", 1) // Outer group
	layoutManager.GroupColumns("C:H", 2) // Inner group
	layoutManager.GroupColumns("E:F", 3) // Innermost group

	// Create nested row grouping
	layoutManager.GroupRows(1, 10, 1) // Outer group
	layoutManager.GroupRows(3, 8, 2)  // Inner group
	layoutManager.GroupRows(5, 6, 3)  // Innermost group

	file := workbook.Build()
	assert.NotNil(t, file)
}

// TestAdvancedLayoutManager_FreezePane tests freeze pane functionality
func TestAdvancedLayoutManager_FreezePane(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add headers and data
	sheet.AddRow().AddCells("Header1", "Header2", "Header3", "Header4")
	for i := 1; i <= 20; i++ {
		sheet.AddRow().AddCells(i, i*2, i*3, i*4)
	}

	layoutManager := sheet.GetLayoutManager()

	// Freeze first row and first column
	result := layoutManager.FreezePane("B2")
	assert.NotNil(t, result)

	file := workbook.Build()

	// Verify freeze pane was set
	panes, err := file.GetPanes("TestSheet")
	require.NoError(t, err)
	assert.NotNil(t, panes)
	assert.Equal(t, "B2", panes.TopLeftCell)
}

// TestAdvancedLayoutManager_SplitPane tests split pane functionality
func TestAdvancedLayoutManager_SplitPane(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add data
	for i := 1; i <= 20; i++ {
		row := sheet.AddRow()
		for j := 1; j <= 10; j++ {
			row.AddCell(i*10 + j)
		}
	}

	layoutManager := sheet.GetLayoutManager()

	// Split at column 3, row 5
	result := layoutManager.SplitPane(3, 5)
	assert.NotNil(t, result)

	file := workbook.Build()
	assert.NotNil(t, file)
}

// TestAdvancedLayoutManager_AutoFitColumns tests auto-fit columns functionality
func TestAdvancedLayoutManager_AutoFitColumns(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add data with varying lengths
	sheet.AddRow().AddCells("Short", "Medium Length Text", "Very Long Text Content That Should Expand Column Width")
	sheet.AddRow().AddCells("A", "B", "C")
	sheet.AddRow().AddCells("X", "Y", "Z")

	layoutManager := sheet.GetLayoutManager()

	// Auto-fit specific columns
	result := layoutManager.AutoFitColumns("A:C")
	assert.NotNil(t, result)

	file := workbook.Build()

	// Verify columns were auto-fitted (widths should be different)
	widthA, err := file.GetColWidth("TestSheet", "A")
	require.NoError(t, err)
	widthB, err := file.GetColWidth("TestSheet", "B")
	require.NoError(t, err)
	widthC, err := file.GetColWidth("TestSheet", "C")
	require.NoError(t, err)

	// Column C should be widest due to long text
	assert.True(t, widthC > widthB)
	assert.True(t, widthB > widthA)
}

// TestAdvancedLayoutManager_SetColumnWidthRange tests setting width for column ranges
func TestAdvancedLayoutManager_SetColumnWidthRange(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	layoutManager := sheet.GetLayoutManager()

	// Set width for column range
	result := layoutManager.SetColumnWidthRange("B:E", 25.0)
	assert.NotNil(t, result)

	file := workbook.Build()

	// Verify all columns in range have the same width
	columns := []string{"B", "C", "D", "E"}
	for _, col := range columns {
		width, err := file.GetColWidth("TestSheet", col)
		require.NoError(t, err)
		assert.Equal(t, 25.0, width)
	}

	// Column A should have default width
	widthA, err := file.GetColWidth("TestSheet", "A")
	require.NoError(t, err)
	assert.True(t, widthA < 25.0)
}

// TestAdvancedLayoutManager_SetRowHeightRange tests setting height for row ranges
func TestAdvancedLayoutManager_SetRowHeightRange(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add some rows
	for i := 1; i <= 10; i++ {
		sheet.AddRow().AddCell(i)
	}

	layoutManager := sheet.GetLayoutManager()

	// Set height for row range
	result := layoutManager.SetRowHeightRange(3, 7, 30.0)
	assert.NotNil(t, result)

	file := workbook.Build()

	// Verify rows in range have the specified height
	for row := 3; row <= 7; row++ {
		height, err := file.GetRowHeight("TestSheet", row)
		require.NoError(t, err)
		assert.Equal(t, 30.0, height)
	}

	// Other rows should have default height
	height1, err := file.GetRowHeight("TestSheet", 1)
	require.NoError(t, err)
	assert.True(t, height1 < 30.0)
}

// TestAdvancedLayoutManager_HideColumns tests column hiding functionality
func TestAdvancedLayoutManager_HideColumns(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add data
	sheet.AddRow().AddCells("A", "B", "C", "D", "E")

	layoutManager := sheet.GetLayoutManager()

	// Hide columns B and D
	result := layoutManager.HideColumns("B:B")
	assert.NotNil(t, result)
	result = layoutManager.HideColumns("D:D")
	assert.NotNil(t, result)

	file := workbook.Build()
	assert.NotNil(t, file)
}

// TestAdvancedLayoutManager_HideRows tests row hiding functionality
func TestAdvancedLayoutManager_HideRows(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add rows
	for i := 1; i <= 10; i++ {
		sheet.AddRow().AddCell(i)
	}

	layoutManager := sheet.GetLayoutManager()

	// Hide rows 3, 5, 7
	result := layoutManager.HideRows(3, 3)
	assert.NotNil(t, result)
	result = layoutManager.HideRows(5, 5)
	assert.NotNil(t, result)
	result = layoutManager.HideRows(7, 7)
	assert.NotNil(t, result)

	file := workbook.Build()
	assert.NotNil(t, file)
}

// TestAdvancedLayoutManager_InvalidInput tests error handling for invalid inputs
func TestAdvancedLayoutManager_InvalidInput(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	layoutManager := sheet.GetLayoutManager()

	// Test invalid column range
	result := layoutManager.GroupColumns("Z:A", 1) // Reverse range
	assert.Nil(t, result)

	// Test invalid row range
	result2 := layoutManager.GroupRows(10, 5, 1) // Start > End
	assert.Nil(t, result2)

	// Test invalid grouping level
	result3 := layoutManager.GroupColumns("A:C", 0) // Level 0 is invalid
	assert.Nil(t, result3)

	// Test invalid freeze pane
	result4 := layoutManager.FreezePane("ZZZZZ999") // Invalid cell reference (too many columns)
	assert.Nil(t, result4)
}

// TestAdvancedLayoutManager_FluentAPI tests fluent API chaining
func TestAdvancedLayoutManager_FluentAPI(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	// Add data
	for i := 1; i <= 10; i++ {
		row := sheet.AddRow()
		for j := 1; j <= 10; j++ {
			row.AddCell(i*10 + j)
		}
	}

	// Test fluent API chaining
	layoutManager := sheet.GetLayoutManager()
	result := layoutManager.
		GroupColumns("A:C", 1).
		GroupRows(1, 3, 1).
		FreezePane("B2").
		SetColumnWidthRange("D:F", 20.0).
		SetRowHeightRange(4, 6, 25.0).
		AutoFitColumns("G:J")

	assert.NotNil(t, result)

	file := workbook.Build()
	assert.NotNil(t, file)

	// Verify freeze pane was set
	panes, err := file.GetPanes("TestSheet")
	require.NoError(t, err)
	assert.NotNil(t, panes)
	assert.Equal(t, "B2", panes.TopLeftCell)

	// Verify column widths
	for _, col := range []string{"D", "E", "F"} {
		width, err := file.GetColWidth("TestSheet", col)
		require.NoError(t, err)
		assert.Equal(t, 20.0, width)
	}

	// Verify row heights
	for row := 4; row <= 6; row++ {
		height, err := file.GetRowHeight("TestSheet", row)
		require.NoError(t, err)
		assert.Equal(t, 25.0, height)
	}
}