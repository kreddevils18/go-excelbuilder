package tests

import (
	"os"
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// TestTemplateBuilder_BasicTemplate tests basic template functionality
func TestTemplateBuilder_BasicTemplate(t *testing.T) {
	// Test: Create a basic template with placeholders
	// Expected:
	// - Template can be created
	// - Placeholders can be replaced
	// - Result is a valid Excel file

	templateBuilder := excelbuilder.NewTemplateBuilder()
	if templateBuilder == nil {
		t.Fatal("Failed to create TemplateBuilder")
	}

	// Create template with placeholders
	template := templateBuilder.
		AddSheet("Template").
		AddRow().
		AddCell().WithValue("{{name}}").Done().
		AddCell().WithValue("{{age}}").Done().
		Done().
		AddRow().
		AddCell().WithValue("{{email}}").Done().
		Done().
		Done()

	assert.NotNil(t, template, "Expected template to be created")

	// Apply data to template
	data := map[string]interface{}{
		"name":  "John Doe",
		"age":   30,
		"email": "john@example.com",
	}

	result := templateBuilder.ProcessTemplate(data)
	assert.NotNil(t, result, "Expected template processing to succeed")

	file := templateBuilder.Build()
	assert.NotNil(t, file, "Expected template to build successfully")
}

// TestTemplateBuilder_DataBinding tests data binding functionality
func TestTemplateBuilder_DataBinding(t *testing.T) {
	// Test: Bind data to template placeholders
	// Expected:
	// - Various data types are handled correctly
	// - Placeholders are replaced with actual values
	// - Formatting is preserved

	templateBuilder := excelbuilder.NewTemplateBuilder()

	// Create template with various data types
	template := templateBuilder.
		AddSheet("DataBinding").
		AddRow().
		AddCell().WithValue("Name: {{name}}").Done().
		AddCell().WithValue("Score: {{score}}").Done().
		AddCell().WithValue("Active: {{active}}").Done().
		AddCell().WithValue("Date: {{date}}").Done().
		Done().
		Done()

	assert.NotNil(t, template)

	// Apply mixed data types
	data := map[string]interface{}{
		"name":   "Alice",
		"score":  95.5,
		"active": true,
		"date":   "2024-01-15",
	}

	result := templateBuilder.ProcessTemplate(data)
	assert.NotNil(t, result)

	file := templateBuilder.Build()
	assert.NotNil(t, file)
}

// TestTemplateBuilder_LoopTemplate tests loop functionality in templates
func TestTemplateBuilder_LoopTemplate(t *testing.T) {
	// Test: Create templates with loops
	// Expected:
	// - Loop syntax is recognized
	// - Multiple rows are generated from array data
	// - Nested data structures are handled

	templateBuilder := excelbuilder.NewTemplateBuilder()

	// Create template with loop
	template := templateBuilder.
		AddSheet("LoopTemplate").
		AddRow().
		AddCell().WithValue("Name").Done().
		AddCell().WithValue("Department").Done().
		AddCell().WithValue("Salary").Done().
		Done().
		// Loop template row
		AddRow().
		AddCell().WithValue("{{#employees}}").Done().
		AddCell().WithValue("{{name}}").Done().
		AddCell().WithValue("{{department}}").Done().
		AddCell().WithValue("{{salary}}").Done().
		AddCell().WithValue("{{/employees}}").Done().
		Done().
		Done()

	assert.NotNil(t, template)

	// Apply array data
	data := map[string]interface{}{
		"employees": []map[string]interface{}{
			{"name": "John", "department": "IT", "salary": 75000},
			{"name": "Jane", "department": "HR", "salary": 65000},
			{"name": "Bob", "department": "Finance", "salary": 70000},
		},
	}

	result := templateBuilder.ProcessTemplate(data)
	assert.NotNil(t, result)

	file := templateBuilder.Build()
	assert.NotNil(t, file)
}

// TestTemplateBuilder_LoadExistingTemplate tests loading existing Excel files as templates
func TestTemplateBuilder_LoadExistingTemplate(t *testing.T) {
	// Test: Load existing Excel file and use as template
	// Expected:
	// - Existing file can be loaded
	// - Template processing works on loaded file
	// - Original file is not modified

	// First create a template file
	builder := excelbuilder.New()
	templateFile := builder.
		NewWorkbook().
		AddSheet("Template").
		AddRow().
		AddCell("Hello {{name}}").Done().
		AddCell("Welcome to {{company}}").Done().
		Done().
		Done().
		Build()

	require.NotNil(t, templateFile)

	// Save template to file
	tempFile := "test_template.xlsx"
	err := templateFile.SaveAs(tempFile)
	require.NoError(t, err)
	defer os.Remove(tempFile)

	// Load existing file as template
	templateBuilder, err := excelbuilder.LoadExistingFile(tempFile)
	require.NoError(t, err)

	// Process template
	data := map[string]interface{}{
		"name":    "Alice",
		"company": "TechCorp",
	}

	result := templateBuilder.ProcessTemplate(data)
	assert.NotNil(t, result)

	// Save processed template as new file
	newFile := "processed_template.xlsx"
	processedFile := templateBuilder.Build()
	err = processedFile.SaveAs(newFile)
	require.NoError(t, err)
	defer os.Remove(newFile)

	// Verify the new file contains processed data
	newFileLoaded, err := excelbuilder.LoadExistingFile(newFile)
	require.NoError(t, err)

	value, err := newFileLoaded.GetCellValue("Template", "A1")
	require.NoError(t, err)
	assert.Contains(t, value, "Alice")
}

// TestTemplateBuilder_ComplexTemplate tests complex template with multiple sheets and data types
func TestTemplateBuilder_ComplexTemplate(t *testing.T) {
	// Create complex template
	templateBuilder := excelbuilder.NewTemplateBuilder()

	// Summary sheet
	summarySheet := templateBuilder.
		AddSheet("Summary").
		AddRow().
		AddCell().WithValue("Report: {{report_title}}").Done().
		Done().
		AddRow().
		AddCell().WithValue("Generated: {{date}}").Done().
		Done().
		AddRow().
		AddCell().WithValue("Total Records: {{total_count}}").Done().
		Done().
		Done()

	// Detail sheet with loops
	detailSheet := templateBuilder.
		AddSheet("Details").
		AddRow().
		AddCell().WithValue("ID").Done().
		AddCell().WithValue("Name").Done().
		AddCell().WithValue("Category").Done().
		AddCell().WithValue("Value").Done().
		AddCell().WithValue("Status").Done().
		Done().
		// Loop for detail records
		AddRow().
		AddCell().WithValue("{{#records}}").Done().
		AddCell().WithValue("{{id}}").Done().
		AddCell().WithValue("{{name}}").Done().
		AddCell().WithValue("{{category}}").Done().
		AddCell().WithValue("{{value}}").Done().
		AddCell().WithValue("{{status}}").Done().
		AddCell().WithValue("{{/records}}").Done().
		Done().
		Done()

	assert.NotNil(t, summarySheet)
	assert.NotNil(t, detailSheet)

	// Complex data structure
	data := map[string]interface{}{
		"report_title": "Monthly Sales Report",
		"date":         "2024-01-31",
		"total_count":  3,
		"records": []map[string]interface{}{
			{
				"id":       1,
				"name":     "Product A",
				"category": "Electronics",
				"value":    1250.50,
				"status":   "Active",
			},
			{
				"id":       2,
				"name":     "Product B",
				"category": "Clothing",
				"value":    750.25,
				"status":   "Active",
			},
			{
				"id":       3,
				"name":     "Product C",
				"category": "Books",
				"value":    125.00,
				"status":   "Discontinued",
			},
		},
	}

	result := templateBuilder.ProcessTemplate(data)
	assert.NotNil(t, result)

	processedFile := templateBuilder.Build()
	assert.NotNil(t, processedFile)

	// Save and verify
	testFile := "complex_template_test.xlsx"
	err := processedFile.SaveAs(testFile)
	require.NoError(t, err)
	defer os.Remove(testFile)

	// Verify file was created
	_, err = os.Stat(testFile)
	assert.NoError(t, err)
}
