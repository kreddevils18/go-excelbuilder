package tests

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// Test Case 1: PivotTableBuilder Creation and Basic Configuration

// TestPivotTableBuilder_Creation tests the creation of PivotTableBuilder
func TestPivotTableBuilder_Creation(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("DataSheet")
	
	// Add some sample data for pivot table
	sheet.AddRow().AddCells("Product", "Region", "Sales", "Quantity")
	sheet.AddRow().AddCells("Laptop", "North", 1000, 10)
	sheet.AddRow().AddCells("Mouse", "South", 500, 25)
	sheet.AddRow().AddCells("Laptop", "South", 1200, 12)
	
	// Create pivot table builder
	pivotBuilder := sheet.NewPivotTable("PivotSheet", "A1:D4")
	
	// Should return a valid PivotTableBuilder instance
	assert.NotNil(t, pivotBuilder, "Expected PivotTableBuilder instance, got nil")
	
	// Should be able to chain methods
	result := pivotBuilder.SetName("SalesPivot")
	assert.Equal(t, pivotBuilder, result, "Expected method chaining to return same instance")
}

// TestPivotTableBuilder_DataSourceConfiguration tests data source setup
func TestPivotTableBuilder_DataSourceConfiguration(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("DataSheet")
	
	// Add sample data
	sheet.AddRow().AddCells("Product", "Region", "Sales", "Quantity")
	sheet.AddRow().AddCells("Laptop", "North", 1000, 10)
	sheet.AddRow().AddCells("Mouse", "South", 500, 25)
	sheet.AddRow().AddCells("Keyboard", "East", 300, 15)
	
	pivotBuilder := sheet.NewPivotTable("PivotSheet", "A1:D2")
	
	// Test data source configuration
	result := pivotBuilder.SetDataSource("DataSheet", "A1:D10")
	assert.Equal(t, pivotBuilder, result, "Expected method chaining")
	
	// Test getting data source info
	config := pivotBuilder.GetConfig()
	assert.Equal(t, "DataSheet", config.SourceSheet, "Expected correct source sheet")
	assert.Equal(t, "A1:D10", config.SourceRange, "Expected correct source range")
}

// Test Case 2: Field Configuration

// TestPivotTableBuilder_RowFields tests row field configuration
func TestPivotTableBuilder_RowFields(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("DataSheet")
	
	pivotBuilder := sheet.NewPivotTable("PivotSheet", "A1:D4")
	
	// Test adding row fields
	result := pivotBuilder.AddRowField("Product")
	assert.Equal(t, pivotBuilder, result, "Expected method chaining")
	
	// Test adding multiple row fields
	pivotBuilder.AddRowField("Region")
	
	// Test getting configuration
	config := pivotBuilder.GetConfig()
	assert.Len(t, config.RowFields, 2, "Expected 2 row fields")
	assert.Equal(t, "Product", config.RowFields[0].Name, "Expected first row field to be Product")
	assert.Equal(t, "Region", config.RowFields[1].Name, "Expected second row field to be Region")
}

// TestPivotTableBuilder_ColumnFields tests column field configuration
func TestPivotTableBuilder_ColumnFields(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("DataSheet")
	
	pivotBuilder := sheet.NewPivotTable("PivotSheet", "A1:D4")
	
	// Test adding column fields
	result := pivotBuilder.AddColumnField("Region")
	assert.Equal(t, pivotBuilder, result, "Expected method chaining")
	
	// Test getting configuration
	config := pivotBuilder.GetConfig()
	assert.Len(t, config.ColumnFields, 1, "Expected 1 column field")
	assert.Equal(t, "Region", config.ColumnFields[0].Name, "Expected column field to be Region")
}

// TestPivotTableBuilder_ValueFields tests value field configuration
func TestPivotTableBuilder_ValueFields(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("DataSheet")
	
	pivotBuilder := sheet.NewPivotTable("PivotSheet", "A1:D4")
	
	// Test adding value fields with aggregation
	result := pivotBuilder.AddValueField("Sales", "sum")
	assert.Equal(t, pivotBuilder, result, "Expected method chaining")
	
	// Test adding multiple value fields
	pivotBuilder.AddValueField("Quantity", "count")
	
	// Test getting configuration
	config := pivotBuilder.GetConfig()
	assert.Len(t, config.ValueFields, 2, "Expected 2 value fields")
	assert.Equal(t, "Sales", config.ValueFields[0].Name, "Expected first value field to be Sales")
	assert.Equal(t, "sum", config.ValueFields[0].Function, "Expected aggregation function to be sum")
	assert.Equal(t, "Quantity", config.ValueFields[1].Name, "Expected second value field to be Quantity")
	assert.Equal(t, "count", config.ValueFields[1].Function, "Expected aggregation function to be count")
}

// Test Case 3: Styling and Formatting

// TestPivotTableBuilder_Styling tests pivot table styling options
func TestPivotTableBuilder_Styling(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("DataSheet")
	
	pivotBuilder := sheet.NewPivotTable("PivotSheet", "A1:D4")
	
	// Test setting pivot table style
	result := pivotBuilder.SetStyle("PivotStyleMedium9")
	assert.Equal(t, pivotBuilder, result, "Expected method chaining")
	
	// Test showing/hiding grand totals
	pivotBuilder.ShowRowGrandTotals(true)
	pivotBuilder.ShowColumnGrandTotals(false)
	
	// Test getting configuration
	config := pivotBuilder.GetConfig()
	assert.Equal(t, "PivotStyleMedium9", config.Style, "Expected correct style")
	assert.True(t, config.ShowRowGrandTotals, "Expected row grand totals to be shown")
	assert.False(t, config.ShowColumnGrandTotals, "Expected column grand totals to be hidden")
}

// Test Case 4: Build and Integration

// TestPivotTableBuilder_Build tests building the pivot table
func TestPivotTableBuilder_Build(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("DataSheet")
	
	// Add some sample data for pivot table
	sheet.AddRow().AddCells("Product", "Region", "Sales", "Quantity")
	sheet.AddRow().AddCells("Laptop", "North", 1000, 10)
	sheet.AddRow().AddCells("Mouse", "South", 500, 25)
	
	// Create and configure pivot table
	pivotBuilder := sheet.NewPivotTable("PivotSheet", "A1:D3").
		SetName("SalesPivot").
		AddRowField("Product").
		AddColumnField("Region").
		AddValueField("Sales", "sum").
		SetStyle("PivotStyleMedium9")
	
	// Build the pivot table
	err := pivotBuilder.Build()
	assert.NoError(t, err, "Expected Build to succeed")
	
	// Build the workbook and verify
	file := workbook.Build()
	require.NotNil(t, file, "Expected valid Excel file")
	
	// Verify pivot sheet was created
	sheets := file.GetSheetList()
	assert.Contains(t, sheets, "PivotSheet", "Expected PivotSheet to be created")
}

// TestPivotTableBuilder_ComplexConfiguration tests complex pivot table setup
func TestPivotTableBuilder_ComplexConfiguration(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("SalesData")
	
	// Add comprehensive sample data
	headers := []interface{}{"Date", "Product", "Category", "Region", "Salesperson", "Sales", "Quantity", "Profit"}
	sheet.AddRow().AddCells(headers...)
	
	data := [][]interface{}{
		{"2024-01-01", "Laptop", "Electronics", "North", "John", 1000, 10, 200},
		{"2024-01-02", "Mouse", "Electronics", "South", "Jane", 500, 25, 100},
		{"2024-01-03", "Keyboard", "Electronics", "North", "John", 300, 15, 60},
		{"2024-01-04", "Monitor", "Electronics", "West", "Bob", 800, 8, 160},
	}
	
	for _, row := range data {
		sheet.AddRow().AddCells(row...)
	}
	
	// Create complex pivot table
	pivotBuilder := sheet.NewPivotTable("ComplexPivot", "A1:H5").
		SetName("SalesAnalysis").
		AddRowField("Category").
		AddRowField("Product").
		AddColumnField("Region").
		AddValueField("Sales", "sum").
		AddValueField("Quantity", "sum").
		AddValueField("Profit", "average").
		SetStyle("PivotStyleDark1").
		ShowRowGrandTotals(true).
		ShowColumnGrandTotals(true)
	
	// Build and verify
	err := pivotBuilder.Build()
	assert.NoError(t, err, "Expected Build to succeed")
	
	file := workbook.Build()
	require.NotNil(t, file, "Expected valid Excel file")
	
	// Verify configuration
	config := pivotBuilder.GetConfig()
	assert.Equal(t, "SalesAnalysis", config.Name, "Expected correct pivot table name")
	assert.Len(t, config.RowFields, 2, "Expected 2 row fields")
	assert.Len(t, config.ColumnFields, 1, "Expected 1 column field")
	assert.Len(t, config.ValueFields, 3, "Expected 3 value fields")
}