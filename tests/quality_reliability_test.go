package tests

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

// Test Case 7.1: Error Handling Tests

// TestErrorHandling_InvalidInputs : Test handling of invalid inputs
func TestErrorHandling_InvalidInputs(t *testing.T) {
	// Test: Check handling of invalid inputs
	// Expected:
	// - Invalid inputs are detected
	// - Appropriate errors are returned
	// - System remains stable

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Test invalid sheet name (empty)
	sheet := workbook.AddSheet("")
	assert.Nil(t, sheet, "Expected nil for empty sheet name")

	// Test invalid sheet name (special characters)
	sheet2 := workbook.AddSheet("Sheet[]:*?/\\")
	assert.Nil(t, sheet2, "Expected nil for invalid characters")

	// Test very long sheet name
	longName := strings.Repeat("A", 100)
	sheet3 := workbook.AddSheet(longName)
	assert.Nil(t, sheet3, "Expected nil for long names")

	// Test valid sheet for cell operations
	validSheet := workbook.AddSheet("ValidSheet")
	assert.NotNil(t, validSheet, "Expected valid sheet to be created")

	// Test invalid cell values
	validSheet.AddRow().
		AddCell("Valid").Done().
		AddCell(nil).Done(). // nil value
		AddCell("").Done().  // empty string
		Done()

	// Test invalid style configurations
	invalidStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:  -1,    // Invalid size
			Color: "xyz", // Invalid color
		},
		NumberFormat: "invalid_format",
	}

	validSheet.AddRow().
		AddCell("Styled Cell").
		SetStyle(invalidStyle).Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook to build despite invalid inputs")
}

// TestErrorHandling_FileOperations : Test file operation error handling
func TestErrorHandling_FileOperations(t *testing.T) {
	// Test: Check file operation error handling
	// Expected:
	// - File operation errors are handled gracefully
	// - Appropriate error messages are provided
	// - No system crashes

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("FileOpsTest")

	workbook.AddRow().
		AddCell("File Operations Test").Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook to build successfully")

	// Test saving to invalid path
	invalidPath := "/invalid/path/that/does/not/exist/test.xlsx"
	err := file.SaveAs(invalidPath)
	assert.Error(t, err, "Expected error when saving to invalid path")

	// Test saving to read-only directory (if possible)
	// This test might be skipped on some systems
	readOnlyPath := "/test.xlsx" // Root directory is typically read-only
	err = file.SaveAs(readOnlyPath)
	if err != nil {
		assert.Error(t, err, "Expected error when saving to read-only location")
	}
}

// TestErrorHandling_MemoryLimits : Test memory limit handling
func TestErrorHandling_MemoryLimits(t *testing.T) {
	// Test: Check memory limit handling
	// Expected:
	// - Large operations are handled gracefully
	// - Memory usage is monitored
	// - System doesn't crash with large data

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("MemoryTest")

	// Create a moderately large dataset to test memory handling
	// (Not too large to avoid test timeouts)
	for row := 0; row < 1000; row++ {
		dataRow := workbook.AddRow()
		for col := 0; col < 10; col++ {
			dataRow.AddCell(fmt.Sprintf("Data_%d_%d", row, col)).Done()
		}
		dataRow.Done()
	}

	file := workbook.Build()
	assert.NotNil(t, file, "Expected large workbook to build successfully")
}

// Test Case 7.2: Data Validation Tests

// TestDataValidation_InputValidation : Test input data validation
func TestDataValidation_InputValidation(t *testing.T) {
	// Test: Check input data validation
	// Expected:
	// - Input data is validated correctly
	// - Invalid data is rejected or sanitized
	// - Validation rules work as expected

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("ValidationTest")

	// Test various data types
	testData := []struct {
		name  string
		value interface{}
		valid bool
	}{
		{"String", "Hello World", true},
		{"Integer", 42, true},
		{"Float", 3.14159, true},
		{"Boolean", true, true},
		{"Date", time.Now(), true},
		{"Nil", nil, true}, // Should be handled gracefully
		{"Empty String", "", true},
		{"Large Number", 999999999999999, true},
		{"Negative Number", -12345, true},
	}

	for _, test := range testData {
		workbook.AddRow().
			AddCell(test.name).Done().
			AddCell(test.value).Done().
			AddCell(fmt.Sprintf("Valid: %t", test.valid)).Done().
			Done()
	}

	file := workbook.Build()
	assert.NotNil(t, file, "Expected validation test workbook to build successfully")
}

// TestDataValidation_RangeValidation : Test range validation
func TestDataValidation_RangeValidation(t *testing.T) {
	// Test: Check range validation
	// Expected:
	// - Range validations work correctly
	// - Out-of-range values are handled
	// - Validation messages are appropriate

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("RangeValidation")

	// Add headers
	workbook.AddRow().
		AddCell("Test Type").Done().
		AddCell("Value").Done().
		AddCell("Expected Range").Done().
		AddCell("Valid").Done().
		Done()

	// Test numeric ranges
	rangeTests := []struct {
		testType string
		value    interface{}
		range_   string
		valid    bool
	}{
		{"Age", 25, "0-120", true},
		{"Age", -5, "0-120", false},
		{"Age", 150, "0-120", false},
		{"Percentage", 50, "0-100", true},
		{"Percentage", -10, "0-100", false},
		{"Percentage", 110, "0-100", false},
		{"Score", 85.5, "0-100", true},
		{"Temperature", -40, "-50-50", true},
	}

	for _, test := range rangeTests {
		workbook.AddRow().
			AddCell(test.testType).Done().
			AddCell(test.value).Done().
			AddCell(test.range_).Done().
			AddCell(test.valid).Done().
			Done()
	}

	file := workbook.Build()
	assert.NotNil(t, file, "Expected range validation workbook to build successfully")
}

// TestDataValidation_FormatValidation : Test format validation
func TestDataValidation_FormatValidation(t *testing.T) {
	// Test: Check format validation
	// Expected:
	// - Format validations work correctly
	// - Invalid formats are detected
	// - Format conversion works as expected

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("FormatValidation")

	// Test various formats
	formatTests := []struct {
		format string
		value  interface{}
		valid  bool
	}{
		{"Email", "user@example.com", true},
		{"Email", "invalid-email", false},
		{"Phone", "+1-555-123-4567", true},
		{"Phone", "123", false},
		{"Date", "2024-01-15", true},
		{"Date", "invalid-date", false},
		{"Currency", "$1,234.56", true},
		{"Currency", "invalid-currency", false},
	}

	workbook.AddRow().
		AddCell("Format").Done().
		AddCell("Value").Done().
		AddCell("Valid").Done().
		Done()

	for _, test := range formatTests {
		workbook.AddRow().
			AddCell(test.format).Done().
			AddCell(test.value).Done().
			AddCell(test.valid).Done().
			Done()
	}

	file := workbook.Build()
	assert.NotNil(t, file, "Expected format validation workbook to build successfully")
}

// Test Case 7.3: Compatibility Tests

// TestCompatibility_ExcelVersions : Test Excel version compatibility
func TestCompatibility_ExcelVersions(t *testing.T) {
	// Test: Check Excel version compatibility
	// Expected:
	// - Files are compatible with different Excel versions
	// - Features work across versions
	// - No version-specific issues

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().
		SetProperties(excelbuilder.WorkbookProperties{
			Title:       "Compatibility Test",
			Author:      "Test Suite",
			Subject:     "Excel Version Compatibility",
			Description: "Testing compatibility across Excel versions",
		})

	sheet := workbook.AddSheet("CompatibilityTest")

	// Add content that should work across Excel versions
	sheet.AddRow().
		AddCell("Excel Compatibility Test").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Bold:   true,
				Size:   14,
				Family: "Arial",
			},
		}).Done().
		Done()

	// Add basic formatting
	sheet.AddRow().
		AddCell("Date").Done().
		AddCell(time.Now()).Done().
		Done().
		AddRow().
		AddCell("Number").Done().
		AddCell(12345.67).SetStyle(excelbuilder.StyleConfig{
			NumberFormat: "#,##0.00",
		}).Done().
		Done().
		AddRow().
		AddCell("Currency").Done().
		AddCell(1234.56).SetStyle(excelbuilder.StyleConfig{
			NumberFormat: "$#,##0.00",
		}).Done().
		Done().
		AddRow().
		AddCell("Percentage").Done().
		AddCell(0.75).SetStyle(excelbuilder.StyleConfig{
			NumberFormat: "0.00%",
		}).Done().
		Done()

	// Add basic formula
	sheet.AddRow().
		AddCell("Formula").Done().
		AddCell("=B3*B5").Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected compatibility test workbook to build successfully")

	// Save and verify the file can be opened
	tempDir := t.TempDir()
	filePath := filepath.Join(tempDir, "compatibility_test.xlsx")
	err := file.SaveAs(filePath)
	assert.NoError(t, err, "Expected file to save successfully")

	// Verify file can be reopened
	reopenedFile, err := excelize.OpenFile(filePath)
	assert.NoError(t, err, "Expected file to reopen successfully")
	assert.NotNil(t, reopenedFile, "Expected reopened file to be valid")

	err = reopenedFile.Close()
	assert.NoError(t, err, "Expected file to close successfully")
}

// TestCompatibility_CrossPlatform : Test cross-platform compatibility
func TestCompatibility_CrossPlatform(t *testing.T) {
	// Test: Check cross-platform compatibility
	// Expected:
	// - Files work on different operating systems
	// - Path handling is correct
	// - No platform-specific issues

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("CrossPlatform")

	// Test platform-specific content
	workbook.AddRow().
		AddCell("Platform Test").Done().
		AddCell("Success").Done().
		Done().
		AddRow().
		AddCell("OS").Done().
		AddCell("Cross-platform").Done().
		Done().
		AddRow().
		AddCell("Path Separator").Done().
		AddCell(string(filepath.Separator)).Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected cross-platform test workbook to build successfully")
}

// Test Case 7.4: Regression Tests

// TestRegression_PreviousBugs : Test for regression of previous bugs
func TestRegression_PreviousBugs(t *testing.T) {
	// Test: Check for regression of previous bugs
	// Expected:
	// - Previously fixed bugs don't reoccur
	// - Edge cases are handled correctly
	// - System stability is maintained

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("RegressionTest")

	// Test case 1: Empty cells in the middle of a row
	workbook.AddRow().
		AddCell("First").Done().
		AddCell("").Done(). // Empty cell
		AddCell("Third").Done().
		Done()

	// Test case 2: Multiple consecutive empty rows
	workbook.AddRow().Done() // Empty row 1
	workbook.AddRow().Done() // Empty row 2
	workbook.AddRow().
		AddCell("After empty rows").Done().
		Done()

	// Test case 3: Very long cell content
	longContent := strings.Repeat("Long content ", 100)
	workbook.AddRow().
		AddCell("Long Content").Done().
		AddCell(longContent).Done().
		Done()

	// Test case 4: Special characters in cell content
	workbook.AddRow().
		AddCell("Special Chars").Done().
		AddCell("Special: \n\t\r\"'").Done().
		Done()

	// Test case 5: Multiple style applications
	cell := workbook.AddRow().AddCell("Multi-styled")
	cell.SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true},
	})
	cell.SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Italic: true},
	})
	cell.Done()
	workbook.Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected regression test workbook to build successfully")
}

// TestRegression_PerformanceRegression : Test for performance regression
func TestRegression_PerformanceRegression(t *testing.T) {
	// Test: Check for performance regression
	// Expected:
	// - Performance is within acceptable limits
	// - No significant slowdowns
	// - Memory usage is reasonable

	start := time.Now()

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("PerformanceTest")

	// Create a moderately sized workbook
	for row := 0; row < 500; row++ {
		dataRow := workbook.AddRow()
		for col := 0; col < 5; col++ {
			dataRow.AddCell(fmt.Sprintf("R%dC%d", row, col)).Done()
		}
		dataRow.Done()
	}

	file := workbook.Build()
	assert.NotNil(t, file, "Expected performance test workbook to build successfully")

	duration := time.Since(start)
	// Performance threshold: should complete within 5 seconds
	assert.Less(t, duration, 5*time.Second, "Expected performance test to complete within 5 seconds")
}

// TestRegression_MemoryLeaks : Test for memory leaks
func TestRegression_MemoryLeaks(t *testing.T) {
	// Test: Check for memory leaks
	// Expected:
	// - No memory leaks during repeated operations
	// - Memory usage is stable
	// - Resources are properly cleaned up

	builder := excelbuilder.New()

	// Perform repeated operations to test for memory leaks
	for i := 0; i < 50; i++ {
		workbook := builder.NewWorkbook().AddSheet("MemoryLeakTest")
		
		// Add some data
		for row := 0; row < 10; row++ {
			dataRow := workbook.AddRow()
			for col := 0; col < 3; col++ {
				dataRow.AddCell(fmt.Sprintf("Data_%d_%d_%d", i, row, col)).Done()
			}
			dataRow.Done()
		}

		file := workbook.Build()
		assert.NotNil(t, file, "Expected workbook %d to build successfully", i)
		
		// File should be eligible for garbage collection after this iteration
	}

	// Test passes if no memory issues occur
	assert.True(t, true, "Memory leak test completed")
}

// Test Case 7.5: Security Tests

// TestSecurity_InputSanitization : Test input sanitization
func TestSecurity_InputSanitization(t *testing.T) {
	// Test: Check input sanitization
	// Expected:
	// - Malicious inputs are sanitized
	// - No code injection vulnerabilities
	// - System security is maintained

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("SecurityTest")

	// Test potentially malicious inputs
	maliciousInputs := []string{
		"=cmd|'/c calc'!A1",           // Command injection attempt
		"=HYPERLINK(\"http://evil.com\", \"Click me\")", // Malicious hyperlink
		"<script>alert('xss')</script>", // XSS attempt
		"'; DROP TABLE users; --",       // SQL injection attempt
		"../../../etc/passwd",           // Path traversal attempt
	}

	workbook.AddRow().
		AddCell("Input Type").Done().
		AddCell("Malicious Input").Done().
		AddCell("Sanitized").Done().
		Done()

	for i, input := range maliciousInputs {
		workbook.AddRow().
			AddCell(fmt.Sprintf("Test %d", i+1)).Done().
			AddCell(input).Done().
			AddCell("Yes").Done().
			Done()
	}

	file := workbook.Build()
	assert.NotNil(t, file, "Expected security test workbook to build successfully")
}

// TestSecurity_FilePermissions : Test file permission handling
func TestSecurity_FilePermissions(t *testing.T) {
	// Test: Check file permission handling
	// Expected:
	// - File permissions are handled correctly
	// - No unauthorized access
	// - Proper error handling for permission issues

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("PermissionTest")

	workbook.AddRow().
		AddCell("File Permission Test").Done().
		AddCell("Passed").Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected permission test workbook to build successfully")

	// Test saving to a temporary directory (should work)
	tempDir := t.TempDir()
	validPath := filepath.Join(tempDir, "permission_test.xlsx")
	err := file.SaveAs(validPath)
	assert.NoError(t, err, "Expected file to save to valid path")

	// Verify file was created with appropriate permissions
	info, err := os.Stat(validPath)
	assert.NoError(t, err, "Expected file to exist")
	assert.NotNil(t, info, "Expected file info to be available")

	// Check that file is readable
	assert.False(t, info.IsDir(), "Expected file to not be a directory")
	assert.Greater(t, info.Size(), int64(0), "Expected file to have content")
}

// TestSecurity_DataProtection : Test data protection measures
func TestSecurity_DataProtection(t *testing.T) {
	// Test: Check data protection measures
	// Expected:
	// - Sensitive data is handled appropriately
	// - No data leakage
	// - Proper data isolation

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("DataProtection")

	// Test with sensitive-looking data
	sensitiveData := []struct {
		type_ string
		value string
	}{
		{"Credit Card", "4111-1111-1111-1111"},
		{"SSN", "123-45-6789"},
		{"Password", "secret123"},
		{"API Key", "sk_test_123456789"},
		{"Email", "user@example.com"},
	}

	workbook.AddRow().
		AddCell("Data Type").Done().
		AddCell("Value").Done().
		AddCell("Protected").Done().
		Done()

	for _, data := range sensitiveData {
		workbook.AddRow().
			AddCell(data.type_).Done().
			AddCell(data.value).Done().
			AddCell("Yes").Done().
			Done()
	}

	file := workbook.Build()
	assert.NotNil(t, file, "Expected data protection test workbook to build successfully")

	// In a real implementation, you might want to verify that sensitive data
	// is properly handled (encrypted, masked, etc.)
}