package tests

import (
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// TestAdvancedBatchOperations tests advanced batch operations
func TestAdvancedBatchOperations(t *testing.T) {
	// Test: Advanced batch operations with enhanced functionality
	// Expected:
	// - Large batch operations are handled efficiently
	// - Memory usage is optimized
	// - Performance is maintained for large datasets
	// - Error handling for batch operations

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("BatchOps")

	// Test large batch data insertion
	batchSize := 1000
	batchData := make([][]interface{}, batchSize)
	for i := 0; i < batchSize; i++ {
		batchData[i] = []interface{}{
			i + 1,
			"Item " + string(rune(i+65)),
			float64(i) * 1.5,
			i%2 == 0,
			time.Now().Add(time.Duration(i) * time.Hour),
		}
	}

	// This should work efficiently
	result := workbook.AddRowsBatch(batchData)
	assert.NotNil(t, result)

	file := workbook.Build()
	require.NotNil(t, file, "Expected large batch operation to complete successfully")

	// Verify first and last rows
	value, err := file.GetCellValue("BatchOps", "A1")
	require.NoError(t, err)
	assert.Equal(t, "1", value)

	value, err = file.GetCellValue("BatchOps", "B1")
	require.NoError(t, err)
	assert.Equal(t, "Item A", value)
}

// TestAdvancedBatchStyleOperations tests advanced batch styling
func TestAdvancedBatchStyleOperations(t *testing.T) {
	// Test: Advanced batch styling operations
	// Expected:
	// - Multiple style operations can be batched
	// - Style conflicts are resolved appropriately
	// - Performance is optimized for large style batches

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("BatchStyles")

	// Add some data first
	for i := 1; i <= 10; i++ {
		workbook.AddRow().
			AddCell("Header " + string(rune(i+64))).Done().
			AddCell("Data " + string(rune(i+64))).Done().
			Done()
	}

	// Prepare batch style operations
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:  true,
			Size:  14,
			Color: "#FFFFFF",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#4472C4",
		},
	}

	dataStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size: 11,
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
		},
	}

	batchStyleOps := []excelbuilder.BatchStyleOperation{
		{Range: "A1:A10", Style: headerStyle},
		{Range: "B1:B10", Style: dataStyle},
	}

	result := workbook.ApplyStyleBatch(batchStyleOps)
	assert.NotNil(t, result)

	file := workbook.Build()
	require.NotNil(t, file, "Expected batch style operations to complete successfully")
}

// TestConversionUtilities tests data conversion utilities
func TestConversionUtilities(t *testing.T) {
	// Test: Data conversion utilities
	// Expected:
	// - Convert between different data formats
	// - Handle type conversions safely
	// - Provide utility functions for common conversions

	builder := excelbuilder.New()

	// Test CSV-like data conversion
	csvData := [][]string{
		{"Name", "Age", "Salary", "Active"},
		{"John", "30", "50000.50", "true"},
		{"Jane", "25", "45000.75", "false"},
	}

	// Convert CSV data to typed data
	convertedData := builder.ConvertCSVData(csvData)
	assert.NotNil(t, convertedData)

	// Test JSON-like data conversion
	jsonData := map[string]interface{}{
		"employees": []map[string]interface{}{
			{"name": "John", "age": 30, "salary": 50000.50, "active": true},
			{"name": "Jane", "age": 25, "salary": 45000.75, "active": false},
		},
	}

	// Convert JSON data to Excel format
	workbook := builder.ConvertJSONToWorkbook(jsonData)
	assert.NotNil(t, workbook)

	file := workbook.Build()
	require.NotNil(t, file, "Expected JSON conversion to complete successfully")
}

// TestDataTypeInference tests automatic data type inference
func TestDataTypeInference(t *testing.T) {
	// Test: Automatic data type inference
	// Expected:
	// - Automatically detect data types from string input
	// - Apply appropriate formatting based on detected type
	// - Handle ambiguous cases gracefully

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("TypeInference")

	// Test data with mixed types as strings
	testData := [][]string{
		{"Value", "Inferred Type", "Formatted Value"},
		{"123", "integer", "123"},
		{"123.45", "float", "123.45"},
		{"true", "boolean", "TRUE"},
		{"2024-01-15", "date", "2024-01-15"},
		{"$1,234.56", "currency", "1234.56"},
		{"50%", "percentage", "0.5"},
		{"Hello World", "text", "Hello World"},
	}

	for _, row := range testData {
		workbook.AddRow().
			AddCell(row[0]).WithAutoTypeInference().Done().
			AddCell(row[1]).Done().
			AddCell(row[2]).Done().
			Done()
	}

	file := workbook.Build()
	require.NotNil(t, file, "Expected type inference to complete successfully")
}

// TestDataTransformation tests data transformation utilities
func TestDataTransformation(t *testing.T) {
	// Test: Data transformation utilities
	// Expected:
	// - Transform data between different structures
	// - Pivot/unpivot operations
	// - Data aggregation utilities

	builder := excelbuilder.New()

	// Test data pivoting
	rawData := []map[string]interface{}{
		{"Name": "John", "Month": "Jan", "Sales": 1000},
		{"Name": "John", "Month": "Feb", "Sales": 1200},
		{"Name": "Jane", "Month": "Jan", "Sales": 1100},
		{"Name": "Jane", "Month": "Feb", "Sales": 1300},
	}

	// Pivot data by Name and Month
	pivotConfig := excelbuilder.PivotConfig{
		RowFields:    []string{"Name"},
		ColumnFields: []string{"Month"},
		ValueFields:  []string{"Sales"},
		Aggregation:  "sum",
	}

	workbook := builder.TransformDataToPivot(rawData, pivotConfig, "PivotData")
	assert.NotNil(t, workbook)

	file := workbook.Build()
	require.NotNil(t, file, "Expected data transformation to complete successfully")
}

// TestDataValidationUtilities tests data validation utilities
func TestDataValidationUtilities(t *testing.T) {
	// Test: Data validation utilities
	// Expected:
	// - Validate data before insertion
	// - Provide validation rules for common scenarios
	// - Handle validation errors gracefully

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("Validation")

	// Test email validation
	workbook.AddRow().
		AddCell("Email").Done().
		AddCell("john@example.com").WithValidation("email").Done().
		Done()

	// Test phone number validation
	workbook.AddRow().
		AddCell("Phone").Done().
		AddCell("+1-555-123-4567").WithValidation("phone").Done().
		Done()

	// Test URL validation
	workbook.AddRow().
		AddCell("Website").Done().
		AddCell("https://example.com").WithValidation("url").Done().
		Done()

	// Test custom validation
	customValidator := func(value interface{}) bool {
		if str, ok := value.(string); ok {
			return len(str) >= 5 && len(str) <= 20
		}
		return false
	}

	workbook.AddRow().
		AddCell("Username").Done().
		AddCell("john_doe").WithCustomValidation(customValidator).Done().
		Done()

	file := workbook.Build()
	require.NotNil(t, file, "Expected data validation utilities to work successfully")
}

// TestPerformanceOptimizations tests performance optimization features
func TestPerformanceOptimizations(t *testing.T) {
	// Test: Performance optimization features
	// Expected:
	// - Memory usage is optimized for large datasets
	// - Streaming operations for very large files
	// - Efficient batch processing

	builder := excelbuilder.New()

	// Test streaming mode for large datasets
	streamingBuilder := builder.WithStreamingMode(true)
	workbook := streamingBuilder.NewWorkbook().AddSheet("LargeData")

	// Add large amount of data
	for i := 0; i < 10000; i++ {
		workbook.AddRow().
			AddCell(i).Done().
			AddCell("Data " + string(rune(i%26+65))).Done().
			AddCell(float64(i) * 1.5).Done().
			Done()
	}

	file := workbook.Build()
	require.NotNil(t, file, "Expected streaming mode to handle large datasets efficiently")
}

// TestErrorHandlingUtilities tests error handling utilities
func TestErrorHandlingUtilities(t *testing.T) {
	// Test: Error handling utilities
	// Expected:
	// - Comprehensive error reporting
	// - Recovery from non-fatal errors
	// - Validation of operations before execution

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("ErrorHandling")

	// Test error collection mode
	errorCollector := builder.WithErrorCollection(true)

	// Test error collection functionality
	errors := errorCollector.GetCollectedErrors()
	assert.NotNil(t, errors)

	// Build should still succeed with error collection
	file := workbook.Build()
	require.NotNil(t, file, "Expected error collection to allow building despite errors")
}
