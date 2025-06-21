package tests

import (
	"math"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// TestComprehensiveDataTypes tests comprehensive data type handling
func TestComprehensiveDataTypes(t *testing.T) {
	// Test: Comprehensive data type handling
	// Expected:
	// - All Go data types are properly handled
	// - Type conversion works correctly
	// - Special values (nil, infinity, NaN) are handled gracefully
	// - Complex data structures are flattened appropriately

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("DataTypes")

	// Test basic types
	workbook.AddRow().
		AddCell("Type").Done().
		AddCell("Value").Done().
		AddCell("Expected").Done().
		Done()

	// String types
	workbook.AddRow().
		AddCell("string").Done().
		AddCell("Hello World").Done().
		AddCell("Hello World").Done().
		Done()

	// Integer types
	workbook.AddRow().
		AddCell("int").Done().
		AddCell(42).Done().
		AddCell("42").Done().
		Done()

	workbook.AddRow().
		AddCell("int8").Done().
		AddCell(int8(127)).Done().
		AddCell("127").Done().
		Done()

	workbook.AddRow().
		AddCell("int16").Done().
		AddCell(int16(32767)).Done().
		AddCell("32767").Done().
		Done()

	workbook.AddRow().
		AddCell("int32").Done().
		AddCell(int32(2147483647)).Done().
		AddCell("2147483647").Done().
		Done()

	workbook.AddRow().
		AddCell("int64").Done().
		AddCell(int64(9223372036854775807)).Done().
		AddCell("9223372036854775807").Done().
		Done()

	// Unsigned integer types
	workbook.AddRow().
		AddCell("uint").Done().
		AddCell(uint(42)).Done().
		AddCell("42").Done().
		Done()

	workbook.AddRow().
		AddCell("uint8").Done().
		AddCell(uint8(255)).Done().
		AddCell("255").Done().
		Done()

	workbook.AddRow().
		AddCell("uint16").Done().
		AddCell(uint16(65535)).Done().
		AddCell("65535").Done().
		Done()

	workbook.AddRow().
		AddCell("uint32").Done().
		AddCell(uint32(4294967295)).Done().
		AddCell("4294967295").Done().
		Done()

	workbook.AddRow().
		AddCell("uint64").Done().
		AddCell(uint64(123456789)).Done().
		AddCell("123456789").Done().
		Done()

	// Float types
	workbook.AddRow().
		AddCell("float32").Done().
		AddCell(float32(3.14159)).Done().
		AddCell("3.14159").Done().
		Done()

	workbook.AddRow().
		AddCell("float64").Done().
		AddCell(3.141592653589793).Done().
		AddCell("3.141592653589793").Done().
		Done()

	// Boolean type
	workbook.AddRow().
		AddCell("bool_true").Done().
		AddCell(true).Done().
		AddCell("true").Done().
		Done()

	workbook.AddRow().
		AddCell("bool_false").Done().
		AddCell(false).Done().
		AddCell("false").Done().
		Done()

	// Time type
	now := time.Now()
	workbook.AddRow().
		AddCell("time").Done().
		AddCell(now).Done().
		AddCell(now.Format("2006-01-02 15:04:05")).Done().
		Done()

	// Nil/null handling
	workbook.AddRow().
		AddCell("nil").Done().
		AddCell(nil).Done().
		AddCell("").Done().
		Done()

	// Special float values
	workbook.AddRow().
		AddCell("infinity").Done().
		AddCell(math.Inf(1)).Done().
		AddCell("+Inf").Done().
		Done()

	workbook.AddRow().
		AddCell("negative_infinity").Done().
		AddCell(math.Inf(-1)).Done().
		AddCell("-Inf").Done().
		Done()

	workbook.AddRow().
		AddCell("nan").Done().
		AddCell(math.NaN()).Done().
		AddCell("NaN").Done().
		Done()

	file := workbook.Build()
	require.NotNil(t, file, "Expected workbook with comprehensive data types to build successfully")

	// Verify some key values
	// Row 2: string
	value, err := file.GetCellValue("DataTypes", "B2")
	require.NoError(t, err)
	assert.Equal(t, "Hello World", value)

	// Row 3: int
	value, err = file.GetCellValue("DataTypes", "B3")
	require.NoError(t, err)
	assert.Equal(t, "42", value)

	// Row 10: float32 (header=1, string=2, int=3, int8=4, int16=5, int32=6, int64=7, uint=8, uint8=9, uint16=10, uint32=11, uint64=12, float32=13)
	value, err = file.GetCellValue("DataTypes", "B13")
	require.NoError(t, err)
	assert.Equal(t, "3.14159", value)

	// Row 15: bool_true (float32=13, float64=14, bool_true=15)
	value, err = file.GetCellValue("DataTypes", "B15")
	require.NoError(t, err)
	assert.Equal(t, "TRUE", value)
}

// TestDataTypeConversion tests data type conversion utilities
func TestDataTypeConversion(t *testing.T) {
	// Test: Data type conversion utilities
	// Expected:
	// - Conversion between different data types works
	// - Error handling for invalid conversions
	// - Precision preservation where possible

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("Conversion")

	// Test conversion from string to numbers
	workbook.AddRow().
		AddCell("String to Int").Done().
		AddCell("123").WithDataType("int").Done().
		Done()

	workbook.AddRow().
		AddCell("String to Float").Done().
		AddCell("123.45").WithDataType("float").Done().
		Done()

	workbook.AddRow().
		AddCell("String to Bool").Done().
		AddCell("true").WithDataType("bool").Done().
		Done()

	// Test conversion from numbers to string
	workbook.AddRow().
		AddCell("Int to String").Done().
		AddCell(123).WithDataType("string").Done().
		Done()

	workbook.AddRow().
		AddCell("Float to String").Done().
		AddCell(123.45).WithDataType("string").Done().
		Done()

	file := workbook.Build()
	require.NotNil(t, file, "Expected workbook with data type conversions to build successfully")
}

// TestComplexDataStructures tests handling of complex data structures
func TestComplexDataStructures(t *testing.T) {
	// Test: Complex data structure handling
	// Expected:
	// - Arrays/slices are flattened or serialized appropriately
	// - Maps are handled correctly
	// - Structs are serialized properly
	// - Nested structures are flattened

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("Complex")

	// Test slice handling
	sliceData := []int{1, 2, 3, 4, 5}
	workbook.AddRow().
		AddCell("Slice").Done().
		AddCell(sliceData).Done().
		Done()

	// Test map handling
	mapData := map[string]interface{}{
		"name": "John",
		"age":  30,
		"city": "New York",
	}
	workbook.AddRow().
		AddCell("Map").Done().
		AddCell(mapData).Done().
		Done()

	// Test struct handling
	type Person struct {
		Name string
		Age  int
		City string
	}
	person := Person{Name: "Jane", Age: 25, City: "Los Angeles"}
	workbook.AddRow().
		AddCell("Struct").Done().
		AddCell(person).Done().
		Done()

	file := workbook.Build()
	require.NotNil(t, file, "Expected workbook with complex data structures to build successfully")
}

// TestDataValidationWithTypes tests data validation with different types
func TestDataValidationWithTypes(t *testing.T) {
	// Test: Data validation with different data types
	// Expected:
	// - Type-specific validation rules work
	// - Error messages are appropriate for data types
	// - Validation preserves data type information

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("Validation")

	// Integer validation
	workbook.AddRow().
		AddCell("Integer (1-100)").Done().
		AddCell(50).WithDataValidation(&excelbuilder.DataValidationConfig{
		Type:     "whole",
		Operator: "between",
		Formula1: []string{"1"},
		Formula2: []string{"100"},
	}).Done().
		Done()

	// Decimal validation
	workbook.AddRow().
		AddCell("Decimal (0.0-1.0)").Done().
		AddCell(0.5).WithDataValidation(&excelbuilder.DataValidationConfig{
		Type:     "decimal",
		Operator: "between",
		Formula1: []string{"0.0"},
		Formula2: []string{"1.0"},
	}).Done().
		Done()

	// Text length validation
	workbook.AddRow().
		AddCell("Text (max 10 chars)").Done().
		AddCell("Hello").WithDataValidation(&excelbuilder.DataValidationConfig{
		Type:     "textLength",
		Operator: "lessThanOrEqual",
		Formula1: []string{"10"},
	}).Done().
		Done()

	file := workbook.Build()
	require.NotNil(t, file, "Expected workbook with data validation to build successfully")
}
