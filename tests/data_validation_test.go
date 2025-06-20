package tests

// NOTE: Data validation is not yet implemented in the excelbuilder package.
// These tests are commented out until the feature is implemented.

import (
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// TestCellBuilder_AddDataValidation_NumberRange tests validation for number range
func TestCellBuilder_AddDataValidation_NumberRange(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell("")

	// This should work when implemented
	validation := excelbuilder.DataValidation{
		Type:             "whole",
		Operator:         "between",
		Formula1:         "1",
		Formula2:         "100",
		ShowErrorMessage: true,
		ErrorTitle:       "Invalid Number",
		ErrorMessage:     "Please enter a number between 1 and 100",
	}

	result := cell.AddDataValidation(validation)

	// Should return the same cell for chaining
	assert.NotNil(t, result)
	assert.Equal(t, cell, result)

	// Build and verify
	file := workbook.Build()
	require.NotNil(t, file)

	// TODO: Add verification when excelize supports reading validation rules
}

// TestCellBuilder_AddDataValidation_TextLength tests validation for text length
func TestCellBuilder_AddDataValidation_TextLength(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell("")

	validation := excelbuilder.DataValidation{
		Type:             "textLength",
		Operator:         "between",
		Formula1:         "5",
		Formula2:         "20",
		ShowErrorMessage: true,
		ErrorTitle:       "Invalid Text Length",
		ErrorMessage:     "Text must be between 5 and 20 characters",
	}

	result := cell.AddDataValidation(validation)
	assert.NotNil(t, result)
	assert.Equal(t, cell, result)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestCellBuilder_AddDataValidation_DropdownList tests dropdown list validation
func TestCellBuilder_AddDataValidation_DropdownList(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell("")

	validation := excelbuilder.DataValidation{
		Type:    "list",
		Formula1: "Option1,Option2,Option3",
		ShowDropDown: true,
		ShowErrorMessage: true,
		ErrorTitle: "Invalid Selection",
		ErrorMessage: "Please select from the dropdown list",
	}

	result := cell.AddDataValidation(validation)
	assert.NotNil(t, result)
	assert.Equal(t, cell, result)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestCellBuilder_AddDataValidation_DateRange tests date validation
func TestCellBuilder_AddDataValidation_DateRange(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell("")

	validation := excelbuilder.DataValidation{
		Type:             "date",
		Operator:         "between",
		Formula1:         "2024-01-01",
		Formula2:         "2024-12-31",
		ShowErrorMessage: true,
		ErrorTitle:       "Invalid Date",
		ErrorMessage:     "Date must be in 2024",
	}

	result := cell.AddDataValidation(validation)
	assert.NotNil(t, result)
	assert.Equal(t, cell, result)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestCellBuilder_AddDataValidation_CustomFormula tests custom validation formula
func TestCellBuilder_AddDataValidation_CustomFormula(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell("")

	validation := excelbuilder.DataValidation{
		Type:             "custom",
		Formula1:         "AND(A1>=0,A1<=100)",
		ShowErrorMessage: true,
		ErrorTitle:       "Custom Validation Failed",
		ErrorMessage:     "Value must satisfy custom formula",
	}

	result := cell.AddDataValidation(validation)
	assert.NotNil(t, result)
	assert.Equal(t, cell, result)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestCellBuilder_AddDataValidation_InvalidInput tests invalid validation input
func TestCellBuilder_AddDataValidation_InvalidInput(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell("")

	// Test with empty type
	validation := excelbuilder.DataValidation{
		Type: "", // Invalid empty type
	}

	result := cell.AddDataValidation(validation)
	// Should handle gracefully - either return nil or same cell
	assert.NotNil(t, result)
}

// TestCellBuilder_AddDataValidation_FluentAPI tests fluent API chaining
func TestCellBuilder_AddDataValidation_FluentAPI(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()

	validation := excelbuilder.DataValidation{
		Type:     "whole",
		Operator: "between",
		Formula1: "1",
		Formula2: "100",
	}

	// Test fluent chaining
	result := row.AddCell("").
		AddDataValidation(validation).
		WithValue("50")

	assert.NotNil(t, result)

	file := workbook.Build()
	require.NotNil(t, file)
}
