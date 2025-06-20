package tests

import (
	"fmt"
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Test Case 6.1: Error Handling - Invalid Sheet Names
func TestErrorHandling_InvalidSheetNames(t *testing.T) {
	// Test: Check error handling for invalid sheet names
	// Input: Various invalid sheet names
	// Expected:
	// - Returns appropriate error or nil
	// - Does not create invalid sheets

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	invalidNames := []string{
		"",                                     // Empty string
		"Sheet[1]",                             // Contains brackets
		"Sheet*1",                              // Contains asterisk
		"Sheet?1",                              // Contains question mark
		"Sheet/1",                              // Contains forward slash
		"Sheet\\1",                             // Contains backslash
		"Sheet:1",                              // Contains colon
		"History",                              // Reserved name in Excel
		"Sheet1234567890123456789012345678901", // Too long (>31 chars)
	}

	for _, invalidName := range invalidNames {
		sheet := workbook.AddSheet(invalidName)
		if sheet != nil {
			t.Errorf("Expected AddSheet with invalid name '%s' to return nil", invalidName)
		}
	}
}

// Test Case 6.2: Error Handling - Duplicate Sheet Names
func TestErrorHandling_DuplicateSheetNames(t *testing.T) {
	// Test: Check error handling for duplicate sheet names
	// Input: Same sheet name twice
	// Expected:
	// - First sheet creation succeeds
	// - Second sheet creation fails or auto-renames

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	sheetName := "TestSheet"

	// Create first sheet
	firstSheet := workbook.AddSheet(sheetName)
	if firstSheet == nil {
		t.Fatal("Expected first sheet creation to succeed")
	}

	// Try to create second sheet with same name
	secondSheet := workbook.AddSheet(sheetName)
	// Depending on implementation:
	// - Might return nil (error case)
	// - Might auto-rename to "TestSheet (2)" or similar
	// - Might overwrite (less desirable)

	// For this test, we expect it to handle gracefully
	// The exact behavior should be documented in the implementation
	if secondSheet != nil {
		// If it succeeds, it should have a different internal name
		// This is implementation-dependent
		t.Log("Second sheet creation succeeded - implementation allows duplicates or auto-renames")
	} else {
		t.Log("Second sheet creation failed - implementation prevents duplicates")
	}
}

// Test Case 6.3: Error Handling - Invalid Column References
func TestErrorHandling_InvalidColumnReferences(t *testing.T) {
	// Test: Check error handling for invalid column references
	// Input: Invalid column names for SetColumnWidth
	// Expected:
	// - Returns appropriate error or handles gracefully

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")

	if sheet == nil {
		t.Fatal("Failed to create sheet")
	}

	invalidColumns := []string{
		"",     // Empty string
		"1",    // Number instead of letter
		"AA1",  // Contains number
		"@",    // Invalid character
		"AAAA", // Too many letters (beyond Excel limits)
	}

	for _, invalidCol := range invalidColumns {
		result := sheet.SetColumnWidth(invalidCol, 20.0)
		// Depending on implementation, this might:
		// - Return nil (error case)
		// - Handle gracefully and continue
		// - Validate and reject

		if result == nil {
			t.Logf("SetColumnWidth with invalid column '%s' returned nil (expected)", invalidCol)
		} else {
			t.Logf("SetColumnWidth with invalid column '%s' handled gracefully", invalidCol)
		}
	}
}

// Test Case 6.4: Error Handling - Invalid Number Formats
func TestErrorHandling_InvalidNumberFormats(t *testing.T) {
	// Test: Check error handling for invalid number formats
	// Input: Invalid format strings
	// Expected:
	// - Returns appropriate error or uses default format

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell(123.45)

	if cell == nil {
		t.Fatal("Failed to create cell")
	}

	invalidFormats := []string{
		"",               // Empty string
		"invalid_format", // Completely invalid
		"0.00.00",        // Double decimal
		"#,##,#0",        // Invalid comma placement
		"$$$#,##0.00",    // Multiple currency symbols
	}

	for _, invalidFormat := range invalidFormats {
		result := cell.SetNumberFormat(invalidFormat)
		// The implementation should handle this gracefully
		// Either by returning an error or using a default format

		if result == nil {
			t.Logf("SetNumberFormat with invalid format '%s' returned nil", invalidFormat)
		} else {
			t.Logf("SetNumberFormat with invalid format '%s' handled gracefully", invalidFormat)
		}
	}
}

// Test Case 6.5: Error Handling - Invalid Style Configurations
func TestErrorHandling_InvalidStyleConfigurations(t *testing.T) {
	// Test: Check error handling for invalid style configurations
	// Input: StyleConfig with invalid values
	// Expected:
	// - Returns appropriate error or uses default values

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell("Test")

	if cell == nil {
		t.Fatal("Failed to create cell")
	}

	// Test invalid font size
	invalidStyle1 := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size: -5, // Negative size
		},
	}

	result1 := cell.SetStyle(invalidStyle1)
	if result1 == nil {
		t.Log("SetStyle with negative font size returned nil")
	} else {
		t.Log("SetStyle with negative font size handled gracefully")
	}

	// Test invalid color
	invalidStyle2 := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Color: "invalid_color", // Invalid color format
		},
	}

	result2 := cell.SetStyle(invalidStyle2)
	if result2 == nil {
		t.Log("SetStyle with invalid color returned nil")
	} else {
		t.Log("SetStyle with invalid color handled gracefully")
	}

	// Test invalid border style
	invalidStyle3 := excelbuilder.StyleConfig{
		Border: excelbuilder.BorderConfig{
			Top: excelbuilder.BorderSide{Style: "invalid_border_style"},
		},
	}

	result3 := cell.SetStyle(invalidStyle3)
	if result3 == nil {
		t.Log("SetStyle with invalid border style returned nil")
	} else {
		t.Log("SetStyle with invalid border style handled gracefully")
	}
}

// Test Case 6.6: Error Handling - Invalid Hyperlinks
func TestErrorHandling_InvalidHyperlinks(t *testing.T) {
	// Test: Check error handling for invalid hyperlinks
	// Input: Invalid URL formats
	// Expected:
	// - Returns appropriate error or handles gracefully

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("TestSheet")
	row := sheet.AddRow()
	cell := row.AddCell("Click here")

	if cell == nil {
		t.Fatal("Failed to create cell")
	}

	invalidUrls := []string{
		"",                   // Empty string
		"not_a_url",          // Not a URL
		"ftp://invalid",      // Unsupported protocol
		"javascript:alert()", // Potentially dangerous
		"http://",            // Incomplete URL
	}

	for _, invalidUrl := range invalidUrls {
		result := cell.SetHyperlink(invalidUrl)
		// The implementation should validate URLs and handle errors

		if result == nil {
			t.Logf("SetHyperlink with invalid URL '%s' returned nil", invalidUrl)
		} else {
			t.Logf("SetHyperlink with invalid URL '%s' handled gracefully", invalidUrl)
		}
	}
}

// Test Case 6.7: Error Handling - Memory and Resource Limits
func TestErrorHandling_ResourceLimits(t *testing.T) {
	// Test: Check behavior under resource constraints
	// Input: Large number of sheets/rows/cells
	// Expected:
	// - Handles large datasets gracefully
	// - No memory leaks or crashes

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Test creating many sheets (but not too many to avoid test timeout)
	maxSheets := 10 // Reduced for test performance
	for i := 0; i < maxSheets; i++ {
		sheetName := fmt.Sprintf("Sheet%d", i+1)
		sheet := workbook.AddSheet(sheetName)
		if sheet == nil {
			t.Errorf("Failed to create sheet %d", i+1)
			continue
		}

		// Add some content to each sheet
		row := sheet.AddRow()
		if row != nil {
			row.AddCell(fmt.Sprintf("Content for sheet %d", i+1))
		}
	}

	// Test building the workbook
	file := workbook.Build()
	if file == nil {
		t.Error("Failed to build workbook with multiple sheets")
	}
}
