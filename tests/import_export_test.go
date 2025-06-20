package tests

// Import/Export functionality implementation using TDD

import (
	"encoding/csv"
	"encoding/json"
	"os"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// TestImportHelper_FromCSV tests importing data from CSV
func TestImportHelper_FromCSV(t *testing.T) {
	// Red Phase: Test should fail initially
	// Create a test CSV file
	csvData := `Name,Age,City
John,25,New York
Jane,30,Los Angeles
Bob,35,Chicago`
	csvFile := "test_data.csv"
	err := os.WriteFile(csvFile, []byte(csvData), 0644)
	require.NoError(t, err)
	defer os.Remove(csvFile)

	// This should work when implemented
	importHelper := excelbuilder.NewImportHelper()
	result := importHelper.FromCSV(csvFile)

	// Should return the same helper for chaining
	assert.NotNil(t, result)
	assert.Equal(t, importHelper, result)

	// Convert to Excel
	excelResult := importHelper.ToExcel()
	assert.NotNil(t, excelResult)

	// Build and verify
	file := excelResult.Build()
	require.NotNil(t, file)

	// Verify imported data
	value, err := file.GetCellValue("ImportedData", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Name", value)

	value, err = file.GetCellValue("ImportedData", "B2")
	require.NoError(t, err)
	assert.Equal(t, "25", value)

	value, err = file.GetCellValue("ImportedData", "C4")
	require.NoError(t, err)
	assert.Equal(t, "Chicago", value)
}

// TestExportHelper_ToCSV tests exporting Excel data to CSV
func TestExportHelper_ToCSV(t *testing.T) {
	// Create Excel file with data
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestData")

	// Add header row
	headerRow := sheet.AddRow()
	headerRow.AddCell("Product")
	headerRow.AddCell("Price")
	headerRow.AddCell("Quantity")

	// Add data rows
	dataRows := [][]interface{}{
		{"Widget A", 25.50, 100},
		{"Widget B", 15.75, 200},
		{"Widget C", 30.00, 150},
	}

	for _, rowData := range dataRows {
		row := sheet.AddRow()
		for _, cellData := range rowData {
			row.AddCell(cellData)
		}
	}

	// Export to CSV
	exportHelper := excelbuilder.NewExportHelper()
	csvFile := "exported_data.csv"
	result := exportHelper.FromExcel(workbook).ToCSV(csvFile)

	assert.NotNil(t, result)
	defer os.Remove(csvFile)

	// Verify CSV file was created
	_, err := os.Stat(csvFile)
	assert.NoError(t, err)

	// Read and verify CSV content
	csvContent, err := os.ReadFile(csvFile)
	require.NoError(t, err)

	reader := csv.NewReader(strings.NewReader(string(csvContent)))
	records, err := reader.ReadAll()
	require.NoError(t, err)

	// Verify header
	require.Greater(t, len(records), 0, "CSV should have at least one record")
	assert.Equal(t, []string{"Product", "Price", "Quantity"}, records[0])

	// Verify data
	require.Greater(t, len(records), 1, "CSV should have at least 2 records (header + data)")
	assert.Equal(t, "Widget A", records[1][0])
	assert.Equal(t, "25.5", records[1][1])
	assert.Equal(t, "100", records[1][2])
}

// TestImportHelper_FromJSON tests importing structured data from JSON
func TestImportHelper_FromJSON(t *testing.T) {
	// Create test JSON data
	jsonData := map[string]interface{}{
		"users": []map[string]interface{}{
			{
				"name":  "John Doe",
				"age":   25,
				"email": "john@example.com",
				"address": map[string]interface{}{
					"street": "123 Main St",
					"city":   "New York",
					"zip":    "10001",
				},
			},
			{
				"name":  "Jane Smith",
				"age":   30,
				"email": "jane@example.com",
				"address": map[string]interface{}{
					"street": "456 Oak Ave",
					"city":   "Los Angeles",
					"zip":    "90210",
				},
			},
		},
	}

	jsonBytes, err := json.Marshal(jsonData)
	require.NoError(t, err)

	jsonFile := "test_data.json"
	err = os.WriteFile(jsonFile, jsonBytes, 0644)
	require.NoError(t, err)
	defer os.Remove(jsonFile)

	// Import JSON data
	importHelper := excelbuilder.NewImportHelper()
	result := importHelper.FromJSON(jsonFile)

	assert.NotNil(t, result)
	assert.Equal(t, importHelper, result)

	// Convert to Excel with flattening
	// Configure flattening options
	flattentOptions := excelbuilder.FlattenOptions{
		Separator: ".",
		MaxDepth:  3,
	}

	excelResult := importHelper.WithFlattenOptions(flattentOptions).ToExcel()
	assert.NotNil(t, excelResult)

	file := excelResult.Build()
	require.NotNil(t, file)

	// Verify flattened data structure
	// Expected headers (alphabetically sorted): address.city, address.street, address.zip, age, email, name
	value, err := file.GetCellValue("JSONData", "A1")
	require.NoError(t, err)
	assert.Equal(t, "address.city", value)

	value, err = file.GetCellValue("JSONData", "B1")
	require.NoError(t, err)
	assert.Equal(t, "address.street", value)
}

// TestExportHelper_ToJSON tests exporting Excel data to JSON
func TestExportHelper_ToJSON(t *testing.T) {
	// Create Excel file with structured data
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("Users")

	// Add headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("name")
	headerRow.AddCell("age")
	headerRow.AddCell("email")
	headerRow.AddCell("city")

	// Add data
	dataRows := [][]interface{}{
		{"John Doe", 25, "john@example.com", "New York"},
		{"Jane Smith", 30, "jane@example.com", "Los Angeles"},
		{"Bob Johnson", 35, "bob@example.com", "Chicago"},
	}

	for _, rowData := range dataRows {
		row := sheet.AddRow()
		for _, cellData := range rowData {
			row.AddCell(cellData)
		}
	}

	// Export to JSON
	exportHelper := excelbuilder.NewExportHelper()
	jsonFile := "exported_data.json"
	result := exportHelper.FromExcel(workbook).ToJSON(jsonFile)

	assert.NotNil(t, result)
	defer os.Remove(jsonFile)

	// Verify JSON file was created
	_, err := os.Stat(jsonFile)
	assert.NoError(t, err)

	// Read and verify JSON content
	jsonContent, err := os.ReadFile(jsonFile)
	require.NoError(t, err)

	var exportedData []map[string]interface{}
	err = json.Unmarshal(jsonContent, &exportedData)
	require.NoError(t, err)

	// Verify structure
	assert.Len(t, exportedData, 3)
	assert.Equal(t, "John Doe", exportedData[0]["name"])
	assert.Equal(t, "25", exportedData[0]["age"]) // Excel data is read as strings
	assert.Equal(t, "john@example.com", exportedData[0]["email"])
}

// TestImportHelper_FromCSV_WithOptions tests CSV import with custom options
func TestImportHelper_FromCSV_WithOptions(t *testing.T) {
	// Create CSV with custom delimiter and quotes
	csvData := `"Name";"Age";"Description"
"John Doe";25;"Software Engineer, Team Lead"
"Jane Smith";30;"Product Manager; Senior Level"`

	csvFile := "custom_csv.csv"
	err := os.WriteFile(csvFile, []byte(csvData), 0644)
	require.NoError(t, err)
	defer os.Remove(csvFile)

	// Import with custom options
	importHelper := excelbuilder.NewImportHelper()
	csvOptions := excelbuilder.CSVOptions{
		Delimiter: ";",
		Quote:     '"',
		SkipRows:  0,
	}

	result := importHelper.FromCSVWithOptions(csvFile, csvOptions)
	assert.NotNil(t, result)

	// Convert to Excel
	workbook := importHelper.ToExcel()
	file := workbook.Build()

	// Verify data with custom delimiter was parsed correctly
	value, err := file.GetCellValue("CustomCSV", "C2")
	require.NoError(t, err)
	assert.Equal(t, "Software Engineer, Team Lead", value)
}

// TestExportHelper_ToCSV_WithOptions tests CSV export with custom options
func TestExportHelper_ToCSV_WithOptions(t *testing.T) {
	// Create Excel data
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("Data")

	headerRow := sheet.AddRow()
	headerRow.AddCell("Name")
	headerRow.AddCell("Description")

	dataRow := sheet.AddRow()
	dataRow.AddCell("John, Jr.")
	dataRow.AddCell("Manager; Senior")

	// Export with custom options
	exportHelper := excelbuilder.NewExportHelper()
	csvOptions := excelbuilder.CSVOptions{
		Delimiter: "|",
		Quote:     '"',
		SkipRows:  0,
	}

	csvFile := "custom_export.csv"
	result := exportHelper.FromExcel(workbook).ToCSVWithOptions(csvFile, csvOptions)
	assert.NotNil(t, result)
	defer os.Remove(csvFile)

	// Verify custom delimiter was used
	csvContent, err := os.ReadFile(csvFile)
	require.NoError(t, err)

	lines := strings.Split(string(csvContent), "\n")
	require.Greater(t, len(lines), 0, "CSV should have at least one line")
	assert.Contains(t, lines[0], "|")
	require.Greater(t, len(lines), 1, "CSV should have at least two lines (header + data)")
	// Check that the custom delimiter is used and data is present
	assert.Contains(t, lines[1], "John, Jr.|Manager; Senior")
}

// TestImportExport_RoundTrip tests round-trip conversion (Excel -> CSV -> Excel)
func TestImportExport_RoundTrip(t *testing.T) {
	// Create original Excel file
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("Original")

	originalData := [][]interface{}{
		{"ID", "Name", "Score", "Active"},
		{1, "Alice", 95.5, true},
		{2, "Bob", 87.2, false},
		{3, "Charlie", 92.8, true},
	}

	for _, rowData := range originalData {
		row := sheet.AddRow()
		for _, cellData := range rowData {
			row.AddCell(cellData)
		}
	}

	// Export to CSV
	exportHelper := excelbuilder.NewExportHelper()
	csvFile := "roundtrip.csv"
	exportHelper.FromExcel(workbook).ToCSV(csvFile)
	defer os.Remove(csvFile)

	// Import back from CSV
	importHelper := excelbuilder.NewImportHelper()
	importHelper.FromCSV(csvFile)

	// Create new Excel file
	newWorkbook := importHelper.ToExcel()
	newFile := newWorkbook.Build()

	// Verify data integrity
	originalFile := workbook.Build()
	originalValue, err := originalFile.GetCellValue("Original", "B2")
	require.NoError(t, err)
	newValue, err := newFile.GetCellValue("ImportedData", "B2")
	require.NoError(t, err)
	assert.Equal(t, originalValue, newValue)

	originalValue, err = originalFile.GetCellValue("Original", "C3")
	require.NoError(t, err)
	newValue, err = newFile.GetCellValue("ImportedData", "C3")
	require.NoError(t, err)
	assert.Equal(t, originalValue, newValue)
}

// TestImportHelper_ErrorHandling tests error handling for import operations
func TestImportHelper_ErrorHandling(t *testing.T) {
	importHelper := excelbuilder.NewImportHelper()

	// Test with non-existent file - should panic
	assert.Panics(t, func() {
		importHelper.FromCSV("non_existent.csv")
	}, "FromCSV should panic when file doesn't exist")

	// Test with invalid JSON
	invalidJSON := `{"invalid": json}`
	jsonFile := "invalid.json"
	err := os.WriteFile(jsonFile, []byte(invalidJSON), 0644)
	require.NoError(t, err)
	defer os.Remove(jsonFile)

	// Test with invalid JSON - should panic
	assert.Panics(t, func() {
		importHelper.FromJSON(jsonFile)
	}, "FromJSON should panic when JSON is invalid")

	// Test converting without data
	emptyResult := importHelper.ToExcel()
	assert.NotNil(t, emptyResult)
}

// TestExportHelper_FluentAPI tests fluent API chaining for export operations
func TestExportHelper_FluentAPI(t *testing.T) {
	// Create Excel data
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("FluentTest")
	row := sheet.AddRow()
	row.AddCell("Test Data")

	// Test fluent chaining
	csvFile := "fluent_export.csv"
	jsonFile := "fluent_export.json"

	result := excelbuilder.NewExportHelper().
		FromExcel(workbook).
		ToCSV(csvFile).
		ToJSON(jsonFile)

	assert.NotNil(t, result)
	defer os.Remove(csvFile)
	defer os.Remove(jsonFile)

	// Verify both files were created
	_, err := os.Stat(csvFile)
	assert.NoError(t, err)

	_, err = os.Stat(jsonFile)
	assert.NoError(t, err)
}
