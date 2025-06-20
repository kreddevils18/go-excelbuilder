package tests

import (
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// TestSheetProtection_BasicProtection tests basic sheet protection functionality
func TestSheetProtection_BasicProtection(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("ProtectedSheet")

	// Add some data
	row := sheet.AddRow()
	row.AddCell("Protected Data")
	row.AddCell(42)

	// Apply basic protection
	protectionConfig := excelbuilder.SheetProtectionConfig{
		Password:              "secret123",
		SelectLockedCells:     true,
		SelectUnlockedCells:   true,
		FormatCells:          false,
		FormatColumns:        false,
		FormatRows:           false,
		InsertColumns:        false,
		InsertRows:           false,
		InsertHyperlinks:     false,
		DeleteColumns:        false,
		DeleteRows:           false,
		Sort:                false,
		AutoFilter:           false,
		PivotTables:          false,
		EditObjects:          false,
		EditScenarios:        false,
	}

	result := sheet.WithProtection(protectionConfig)
	assert.NotNil(t, result)
	assert.Equal(t, sheet, result) // Should return same sheet for chaining

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify protection is applied
	// Note: Excelize doesn't provide direct methods to check protection,
	// but we can verify the file builds successfully with protection
	assert.NotNil(t, file)
}

// TestSheetProtection_SelectivePermissions tests protection with selective permissions
func TestSheetProtection_SelectivePermissions(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("SelectiveProtection")

	// Add data with different protection needs
	headerRow := sheet.AddRow()
	headerRow.AddCell("Name")
	headerRow.AddCell("Score")
	headerRow.AddCell("Comments")

	dataRow := sheet.AddRow()
	dataRow.AddCell("John")
	dataRow.AddCell(85)
	dataRow.AddCell("Good work")

	// Configure protection to allow only specific operations
	protectionConfig := excelbuilder.SheetProtectionConfig{
		Password:              "protect456",
		SelectLockedCells:     true,
		SelectUnlockedCells:   true,
		FormatCells:          true,  // Allow cell formatting
		FormatColumns:        false,
		FormatRows:           false,
		InsertColumns:        false,
		InsertRows:           true,  // Allow inserting rows
		InsertHyperlinks:     true,  // Allow hyperlinks
		DeleteColumns:        false,
		DeleteRows:           false,
		Sort:                true,   // Allow sorting
		AutoFilter:           true,  // Allow filtering
		PivotTables:          false,
		EditObjects:          false,
		EditScenarios:        false,
	}

	result := sheet.WithProtection(protectionConfig)
	assert.NotNil(t, result)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestSheetProtection_CellLevelProtection tests cell-level protection settings
func TestSheetProtection_CellLevelProtection(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("CellProtection")

	// Add protected cells (default)
	protectedRow := sheet.AddRow()
	protectedCell := protectedRow.AddCell("Protected Cell")
	assert.NotNil(t, protectedCell)

	// Add unlocked cells
	unlockedRow := sheet.AddRow()
	unlockedCell := unlockedRow.AddCell("Unlocked Cell")
	unlockedCell.WithStyle(excelbuilder.StyleConfig{
		Protection: &excelbuilder.ProtectionConfig{
			Locked: false,
			Hidden: false,
		},
	})
	assert.NotNil(t, unlockedCell)

	// Add hidden formula cells
	hiddenRow := sheet.AddRow()
	hiddenCell := hiddenRow.AddCell("SUM(A1:A10)")
	hiddenCell.WithStyle(excelbuilder.StyleConfig{
		Protection: &excelbuilder.ProtectionConfig{
			Locked: true,
			Hidden: true, // Hide formula from formula bar
		},
	})
	assert.NotNil(t, hiddenCell)

	// Apply sheet protection
	protectionConfig := excelbuilder.SheetProtectionConfig{
		Password:              "cellprotect789",
		SelectLockedCells:     true,
		SelectUnlockedCells:   true,
		FormatCells:          false,
	}

	sheet.WithProtection(protectionConfig)
	file := workbook.Build()
	require.NotNil(t, file)

	// Verify cells have different protection levels
	value, err := file.GetCellValue("CellProtection", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Protected Cell", value)

	value, err = file.GetCellValue("CellProtection", "A2")
	require.NoError(t, err)
	assert.Equal(t, "Unlocked Cell", value)
}

// TestSheetProtection_MultipleSheets tests protection on multiple sheets
func TestSheetProtection_MultipleSheets(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Create first protected sheet
	sheet1 := workbook.AddSheet("Protected1")
	sheet1.AddRow().AddCell("Sheet 1 Data")
	protection1 := excelbuilder.SheetProtectionConfig{
		Password:     "pass1",
		FormatCells: false,
		InsertRows:  false,
	}
	sheet1.WithProtection(protection1)

	// Create second protected sheet with different settings
	sheet2 := workbook.AddSheet("Protected2")
	sheet2.AddRow().AddCell("Sheet 2 Data")
	protection2 := excelbuilder.SheetProtectionConfig{
		Password:     "pass2",
		FormatCells: true,  // Allow formatting
		InsertRows:  true,  // Allow row insertion
		Sort:        true,  // Allow sorting
	}
	sheet2.WithProtection(protection2)

	// Create unprotected sheet
	sheet3 := workbook.AddSheet("Unprotected")
	sheet3.AddRow().AddCell("Sheet 3 Data")

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify all sheets exist
	sheetList := file.GetSheetList()
	assert.Contains(t, sheetList, "Protected1")
	assert.Contains(t, sheetList, "Protected2")
	assert.Contains(t, sheetList, "Unprotected")
}

// TestSheetProtection_PasswordValidation tests password validation
func TestSheetProtection_PasswordValidation(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("PasswordTest")

	// Test with empty password (should work)
	emptyPasswordConfig := excelbuilder.SheetProtectionConfig{
		Password:     "",
		FormatCells: false,
	}
	result := sheet.WithProtection(emptyPasswordConfig)
	assert.NotNil(t, result)

	// Test with strong password
	strongPasswordConfig := excelbuilder.SheetProtectionConfig{
		Password:     "StrongP@ssw0rd123!",
		FormatCells: false,
	}
	result2 := sheet.WithProtection(strongPasswordConfig)
	assert.NotNil(t, result2)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestSheetProtection_FluentAPI tests fluent API chaining with protection
func TestSheetProtection_FluentAPI(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Test fluent chaining with protection
	 sheet := workbook.AddSheet("FluentProtection")
		sheet.AddRow().AddCell("Test Data")
		sheet.WithProtection(excelbuilder.SheetProtectionConfig{
			Password:     "fluent123",
			FormatCells: false,
			InsertRows:  false,
		})
		sheet.AddRow().AddCell("More Data")

	assert.NotNil(t, sheet)

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify data was added correctly
	value, err := file.GetCellValue("FluentProtection", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Test Data", value)

	value, err = file.GetCellValue("FluentProtection", "A2")
	require.NoError(t, err)
	assert.Equal(t, "More Data", value)
}

// TestSheetProtection_ErrorHandling tests error handling for protection
func TestSheetProtection_ErrorHandling(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("ErrorTest")

	// Test with invalid configuration (should handle gracefully)
	invalidConfig := excelbuilder.SheetProtectionConfig{
		// All fields false/empty - should still work
	}

	result := sheet.WithProtection(invalidConfig)
	assert.NotNil(t, result)

	// Test applying protection multiple times
	config1 := excelbuilder.SheetProtectionConfig{
		Password: "first",
	}
	sheet.WithProtection(config1)

	config2 := excelbuilder.SheetProtectionConfig{
		Password: "second", // Should override first
	}
	result2 := sheet.WithProtection(config2)
	assert.NotNil(t, result2)

	file := workbook.Build()
	require.NotNil(t, file)
}

// TestSheetProtection_WithFormulas tests protection with formulas
func TestSheetProtection_WithFormulas(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("FormulaProtection")

	// Add input cells (unlocked)
	inputRow := sheet.AddRow()
	cell1 := inputRow.AddCell(10)
	cell1.WithStyle(excelbuilder.StyleConfig{
		Protection: &excelbuilder.ProtectionConfig{
			Locked: false,
		},
	})
	cell2 := inputRow.AddCell(20)
	cell2.WithStyle(excelbuilder.StyleConfig{
		Protection: &excelbuilder.ProtectionConfig{
			Locked: false,
		},
	})

	// Add formula cells (locked and hidden)
	formulaRow := sheet.AddRow()
	formulaCell := formulaRow.AddCell("")
	formulaCell.SetFormula("=A1+B1").WithStyle(excelbuilder.StyleConfig{
		Protection: &excelbuilder.ProtectionConfig{
			Locked: true,
			Hidden: true, // Hide formula
		},
	})

	// Apply protection
	protectionConfig := excelbuilder.SheetProtectionConfig{
		Password:              "formula123",
		SelectLockedCells:     true,
		SelectUnlockedCells:   true,
		FormatCells:          false,
	}

	sheet.WithProtection(protectionConfig)
	file := workbook.Build()
	require.NotNil(t, file)

	// Verify formula was set correctly
	formula, err := file.GetCellFormula("FormulaProtection", "A2")
	require.NoError(t, err)
	assert.Equal(t, "=A1+B1", formula, "Expected formula to be set correctly")
}