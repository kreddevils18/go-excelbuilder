package tests

import (
	"fmt"
	"os"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// TestTabColors_BasicColors tests basic tab color functionality
func TestTabColors_BasicColors(t *testing.T) {
	// Red Phase: Test should fail initially
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Create sheets with different tab colors
	redSheet := workbook.AddSheet("RedSheet")
	result := redSheet.WithTabColor("#FF0000") // Red
	assert.NotNil(t, result)
	assert.Equal(t, redSheet, result) // Should return same sheet for chaining

	greenSheet := workbook.AddSheet("GreenSheet")
	greenSheet.WithTabColor("#00FF00") // Green

	blueSheet := workbook.AddSheet("BlueSheet")
	blueSheet.WithTabColor("#0000FF") // Blue

	// Add some data to verify sheets work correctly
	redSheet.AddRow().AddCell("Red Sheet Data")
	greenSheet.AddRow().AddCell("Green Sheet Data")
	blueSheet.AddRow().AddCell("Blue Sheet Data")

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify sheets exist and have data
	value, err := file.GetCellValue("RedSheet", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Red Sheet Data", value)

	value, err = file.GetCellValue("GreenSheet", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Green Sheet Data", value)

	value, err = file.GetCellValue("BlueSheet", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Blue Sheet Data", value)
}

// TestTabColors_PredefinedColors tests using predefined color constants
func TestTabColors_PredefinedColors(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Test with predefined color constants
	sheet1 := workbook.AddSheet("PrimaryRed")
	sheet1.WithTabColor(excelbuilder.ColorRed)

	sheet2 := workbook.AddSheet("PrimaryBlue")
	sheet2.WithTabColor(excelbuilder.ColorBlue)

	sheet3 := workbook.AddSheet("PrimaryGreen")
	sheet3.WithTabColor(excelbuilder.ColorGreen)

	sheet4 := workbook.AddSheet("SecondaryYellow")
	sheet4.WithTabColor(excelbuilder.ColorYellow)

	sheet5 := workbook.AddSheet("SecondaryOrange")
	sheet5.WithTabColor(excelbuilder.ColorOrange)

	sheet6 := workbook.AddSheet("SecondaryPurple")
	sheet6.WithTabColor(excelbuilder.ColorPurple)

	// Add data to verify functionality
	sheet1.AddRow().AddCell("Red Tab")
	sheet2.AddRow().AddCell("Blue Tab")
	sheet3.AddRow().AddCell("Green Tab")
	sheet4.AddRow().AddCell("Yellow Tab")
	sheet5.AddRow().AddCell("Orange Tab")
	sheet6.AddRow().AddCell("Purple Tab")

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify all sheets exist
	sheetList := file.GetSheetList()
	assert.Contains(t, sheetList, "PrimaryRed")
	assert.Contains(t, sheetList, "PrimaryBlue")
	assert.Contains(t, sheetList, "PrimaryGreen")
	assert.Contains(t, sheetList, "SecondaryYellow")
	assert.Contains(t, sheetList, "SecondaryOrange")
	assert.Contains(t, sheetList, "SecondaryPurple")
}

// TestTabColors_CustomHexColors tests custom hex color values
func TestTabColors_CustomHexColors(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Test various hex color formats
	sheet1 := workbook.AddSheet("CustomHex1")
	sheet1.WithTabColor("#FF5733") // Orange-red

	sheet2 := workbook.AddSheet("CustomHex2")
	sheet2.WithTabColor("#33FF57") // Lime green

	sheet3 := workbook.AddSheet("CustomHex3")
	sheet3.WithTabColor("#3357FF") // Blue

	sheet4 := workbook.AddSheet("CustomHex4")
	sheet4.WithTabColor("#FF33F5") // Magenta

	sheet5 := workbook.AddSheet("CustomHex5")
	sheet5.WithTabColor("#F5FF33") // Yellow-green

	sheet6 := workbook.AddSheet("CustomHex6")
	sheet6.WithTabColor("#33F5FF") // Cyan

	// Add data to each sheet
	sheets := []*excelbuilder.SheetBuilder{sheet1, sheet2, sheet3, sheet4, sheet5, sheet6}
	for i, sheet := range sheets {
		sheet.AddRow().AddCell(fmt.Sprintf("Custom Color %d", i+1))
	}

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify data in custom colored sheets
	value, err := file.GetCellValue("CustomHex1", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Custom Color 1", value)

	value, err = file.GetCellValue("CustomHex6", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Custom Color 6", value)
}

// TestTabColors_RGBColors tests RGB color specification
func TestTabColors_RGBColors(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Test RGB color specification
	sheet1 := workbook.AddSheet("RGB1")
	sheet1.WithTabColorRGB(excelbuilder.RGBColor{R: 255, G: 100, B: 100}) // Light red

	sheet2 := workbook.AddSheet("RGB2")
	sheet2.WithTabColorRGB(excelbuilder.RGBColor{R: 100, G: 255, B: 100}) // Light green

	sheet3 := workbook.AddSheet("RGB3")
	sheet3.WithTabColorRGB(excelbuilder.RGBColor{R: 100, G: 100, B: 255}) // Light blue

	sheet4 := workbook.AddSheet("RGB4")
	sheet4.WithTabColorRGB(excelbuilder.RGBColor{R: 255, G: 255, B: 100}) // Yellow

	// Add data
	sheet1.AddRow().AddCell("RGB Light Red")
	sheet2.AddRow().AddCell("RGB Light Green")
	sheet3.AddRow().AddCell("RGB Light Blue")
	sheet4.AddRow().AddCell("RGB Yellow")

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify sheets work correctly
	value, err := file.GetCellValue("RGB1", "A1")
	require.NoError(t, err)
	assert.Equal(t, "RGB Light Red", value)
}

// TestTabColors_ThemeColors tests theme-based colors
func TestTabColors_ThemeColors(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Test theme colors
	sheet1 := workbook.AddSheet("ThemeAccent1")
	sheet1.WithTabColorTheme(excelbuilder.ThemeColorIndexAccent1)

	sheet2 := workbook.AddSheet("ThemeAccent2")
	sheet2.WithTabColorTheme(excelbuilder.ThemeColorIndexAccent2)

	sheet3 := workbook.AddSheet("ThemeAccent3")
	sheet3.WithTabColorTheme(excelbuilder.ThemeColorIndexAccent3)

	sheet4 := workbook.AddSheet("ThemeDark1")
	sheet4.WithTabColorTheme(excelbuilder.ThemeColorIndexDark1)

	sheet5 := workbook.AddSheet("ThemeLight1")
	sheet5.WithTabColorTheme(excelbuilder.ThemeColorIndexLight1)

	// Add data
	sheet1.AddRow().AddCell("Theme Accent 1")
	sheet2.AddRow().AddCell("Theme Accent 2")
	sheet3.AddRow().AddCell("Theme Accent 3")
	sheet4.AddRow().AddCell("Theme Dark 1")
	sheet5.AddRow().AddCell("Theme Light 1")

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify theme colored sheets
	value, err := file.GetCellValue("ThemeAccent1", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Theme Accent 1", value)
}

// TestTabColors_FluentAPI tests fluent API chaining with tab colors
func TestTabColors_FluentAPI(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Test fluent API chaining
	sheet := workbook.AddSheet("Fluent").
		WithTabColor("#FF5733")
	
	// Add content to the sheet
	sheet.AddRow().AddCell("Fluent API Test")

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify the sheet was created
	assert.NotNil(t, sheet)

	filePath := "test_fluent_api.xlsx"
	err := file.SaveAs(filePath)
	assert.NoError(t, err)

	// Clean up
	os.Remove(filePath)
}

// TestTabColors_MultipleColorChanges tests changing tab colors multiple times
func TestTabColors_MultipleColorChanges(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("ChangingColors")

	// Change color multiple times (last one should win)
	sheet.WithTabColor("#FF0000") // Red first
	sheet.WithTabColor("#00FF00") // Then green
	sheet.WithTabColor("#0000FF") // Finally blue

	sheet.AddRow().AddCell("Final color should be blue")

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify sheet still works correctly
	value, err := file.GetCellValue("ChangingColors", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Final color should be blue", value)
}

// TestTabColors_InvalidColors tests error handling for invalid colors
func TestTabColors_InvalidColors(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("InvalidColors")

	// Test with invalid hex colors (should handle gracefully)
	result1 := sheet.WithTabColor("invalid") // Invalid format
	assert.NotNil(t, result1)                // Should still return sheet

	result2 := sheet.WithTabColor("#GGGGGG") // Invalid hex characters
	assert.NotNil(t, result2)

	result3 := sheet.WithTabColor("#FF") // Too short
	assert.NotNil(t, result3)

	result4 := sheet.WithTabColor("#FFAABBCC") // Too long
	assert.NotNil(t, result4)

	// Test with empty color
	result5 := sheet.WithTabColor("")
	assert.NotNil(t, result5)

	sheet.AddRow().AddCell("Invalid colors handled")

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify sheet still works despite invalid colors
	value, err := file.GetCellValue("InvalidColors", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Invalid colors handled", value)
}

// TestTabColors_CategorizedSheets tests using colors to categorize sheets
func TestTabColors_CategorizedSheets(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Data sheets (blue theme)
	dataColor := "#4472C4"
	dataSheet1 := workbook.AddSheet("RawData")
	dataSheet1.WithTabColor(dataColor)
	dataSheet1.AddRow().AddCell("Raw data here")

	dataSheet2 := workbook.AddSheet("ProcessedData")
	dataSheet2.WithTabColor(dataColor)
	dataSheet2.AddRow().AddCell("Processed data here")

	// Analysis sheets (green theme)
	analysisColor := "#70AD47"
	analysisSheet1 := workbook.AddSheet("Analysis1")
	analysisSheet1.WithTabColor(analysisColor)
	analysisSheet1.AddRow().AddCell("Analysis results")

	analysisSheet2 := workbook.AddSheet("Analysis2")
	analysisSheet2.WithTabColor(analysisColor)
	analysisSheet2.AddRow().AddCell("More analysis")

	// Report sheets (orange theme)
	reportColor := "#FFC000"
	reportSheet1 := workbook.AddSheet("Summary")
	reportSheet1.WithTabColor(reportColor)
	reportSheet1.AddRow().AddCell("Executive summary")

	reportSheet2 := workbook.AddSheet("DetailedReport")
	reportSheet2.WithTabColor(reportColor)
	reportSheet2.AddRow().AddCell("Detailed findings")

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify all categorized sheets exist and work
	sheetList := file.GetSheetList()
	assert.Contains(t, sheetList, "RawData")
	assert.Contains(t, sheetList, "ProcessedData")
	assert.Contains(t, sheetList, "Analysis1")
	assert.Contains(t, sheetList, "Analysis2")
	assert.Contains(t, sheetList, "Summary")
	assert.Contains(t, sheetList, "DetailedReport")

	// Verify content
	value, err := file.GetCellValue("Summary", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Executive summary", value)
}

// TestTabColors_WithProtection tests tab colors combined with sheet protection
func TestTabColors_WithProtection(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("ColoredAndProtected")

	// Apply both tab color and protection
	sheet.WithTabColor("#E74C3C"). // Red color
					WithProtection(excelbuilder.SheetProtectionConfig{
			Password:    "protect123",
			FormatCells: false,
		})

	// Add some data
	sheet.AddRow().AddCell("Protected and colored sheet")

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify both features work together
	value, err := file.GetCellValue("ColoredAndProtected", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Protected and colored sheet", value)
}

// TestTabColors_GrayscaleColors tests grayscale color variations
func TestTabColors_GrayscaleColors(t *testing.T) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Test various shades of gray
	grayscaleColors := []string{
		"#000000", // Black
		"#333333", // Dark gray
		"#666666", // Medium gray
		"#999999", // Light gray
		"#CCCCCC", // Very light gray
		"#FFFFFF", // White
	}

	for i, color := range grayscaleColors {
		sheetName := fmt.Sprintf("Gray%d", i+1)
		sheet := workbook.AddSheet(sheetName)
		sheet.WithTabColor(color)
		sheet.AddRow().AddCell(fmt.Sprintf("Grayscale %d", i+1))
	}

	file := workbook.Build()
	require.NotNil(t, file)

	// Verify all grayscale sheets
	for i := 1; i <= len(grayscaleColors); i++ {
		sheetName := fmt.Sprintf("Gray%d", i)
		value, err := file.GetCellValue(sheetName, "A1")
		require.NoError(t, err)
		assert.Equal(t, fmt.Sprintf("Grayscale %d", i), value)
	}
}
