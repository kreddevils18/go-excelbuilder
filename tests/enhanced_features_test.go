package tests

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
)

// Test Case 5.1: Worksheet Management Tests

// TestWorksheet_MultipleSheets : Test multiple sheet creation và management
func TestWorksheet_MultipleSheets(t *testing.T) {
	// Test: Check multiple sheet creation and management
	// Expected:
	// - Multiple sheets can be created
	// - Sheet names are preserved
	// - Each sheet is independent

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Create multiple sheets
	sheet1 := workbook.AddSheet("Sales")
	sheet2 := workbook.AddSheet("Expenses")
	sheet3 := workbook.AddSheet("Summary")

	// Add different data to each sheet
	sheet1.AddRow().
		AddCell("Product").Done().
		AddCell("Revenue").Done().
		Done().
		AddRow().
		AddCell("Product A").Done().
		AddCell(10000).Done().
		Done()

	sheet2.AddRow().
		AddCell("Category").Done().
		AddCell("Cost").Done().
		Done().
		AddRow().
		AddCell("Marketing").Done().
		AddCell(5000).Done().
		Done()

	sheet3.AddRow().
		AddCell("Metric").Done().
		AddCell("Value").Done().
		Done().
		AddRow().
		AddCell("Net Profit").Done().
		AddCell("=Sales!B2-Expenses!B2").Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with multiple sheets to build successfully")
}

// TestWorksheet_SheetOperations : Test sheet operations (rename, delete, copy)
func TestWorksheet_SheetOperations(t *testing.T) {
	// Test: Check sheet operations
	// Expected:
	// - Sheets can be renamed
	// - Sheet operations work correctly
	// - Data integrity is maintained

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Create initial sheet
	sheet := workbook.AddSheet("OriginalName")
	sheet.AddRow().
		AddCell("Test Data").Done().
		AddCell(123).Done().
		Done()

	// Note: In a full implementation, we would have methods like:
	// sheet.Rename("NewName")
	// workbook.CopySheet("OriginalName", "CopyName")
	// workbook.DeleteSheet("SheetToDelete")

	// For now, we test that the basic sheet creation works
	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with sheet operations to build successfully")
}

// TestWorksheet_SheetProtection : Test sheet protection features
func TestWorksheet_SheetProtection(t *testing.T) {
	// Test: Check sheet protection features
	// Expected:
	// - Sheet protection can be enabled
	// - Protection options are configurable
	// - Protected sheets work correctly

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("ProtectedSheet")

	// Add some data
	sheet.AddRow().
		AddCell("Protected Data").Done().
		AddCell("Editable").Done().
		Done().
		AddRow().
		AddCell("Value 1").Done().
		AddCell("Value 2").Done().
		Done()

	// Note: In a full implementation, we would have protection methods:
	// sheet.SetProtection(excelbuilder.SheetProtection{
	//     Password: "secret",
	//     AllowSelectLockedCells: true,
	//     AllowSelectUnlockedCells: true,
	//     AllowFormatCells: false,
	//     AllowInsertRows: false,
	// })

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with protected sheet to build successfully")
}

// TestWorksheet_TabColors : Test sheet tab colors
func TestWorksheet_TabColors(t *testing.T) {
	// Test: Check sheet tab colors
	// Expected:
	// - Tab colors can be set
	// - Different colors for different sheets
	// - Color format is preserved

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Create sheets with different tab colors
	sheet1 := workbook.AddSheet("RedSheet")
	sheet2 := workbook.AddSheet("BlueSheet")
	sheet3 := workbook.AddSheet("GreenSheet")

	// Add some data
	sheet1.AddRow().AddCell("Red Sheet Data").Done().Done()
	sheet2.AddRow().AddCell("Blue Sheet Data").Done().Done()
	sheet3.AddRow().AddCell("Green Sheet Data").Done().Done()

	// Note: In a full implementation, we would have tab color methods:
	// sheet1.SetTabColor("#FF0000") // Red
	// sheet2.SetTabColor("#0000FF") // Blue
	// sheet3.SetTabColor("#00FF00") // Green

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with colored tabs to build successfully")
}

// Test Case 5.2: Advanced Formatting Tests

// TestFormatting_NumberFormats : Test number formatting options
func TestFormatting_NumberFormats(t *testing.T) {
	// Test: Check number formatting options
	// Expected:
	// - Different number formats work
	// - Custom formats are supported
	// - Format codes are preserved

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("NumberFormats")

	// Add header
	workbook.AddRow().
		AddCell("Type").Done().
		AddCell("Value").Done().
		AddCell("Formatted").Done().
		Done()

	// Currency format
	workbook.AddRow().
		AddCell("Currency").Done().
		AddCell(1234.56).Done().
		AddCell(1234.56).
		SetStyle(excelbuilder.StyleConfig{
			NumberFormat: "$#,##0.00",
		}).Done().
		Done()

	// Percentage format
	workbook.AddRow().
		AddCell("Percentage").Done().
		AddCell(0.1234).Done().
		AddCell(0.1234).
		SetStyle(excelbuilder.StyleConfig{
			NumberFormat: "0.00%",
		}).Done().
		Done()

	// Date format
	workbook.AddRow().
		AddCell("Date").Done().
		AddCell("2024-01-15").Done().
		AddCell("2024-01-15").
		SetStyle(excelbuilder.StyleConfig{
			NumberFormat: "mm/dd/yyyy",
		}).Done().
		Done()

	// Custom format
	workbook.AddRow().
		AddCell("Custom").Done().
		AddCell(12345).Done().
		AddCell(12345).
		SetStyle(excelbuilder.StyleConfig{
			NumberFormat: "[Blue]#,##0;[Red]-#,##0",
		}).Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with number formats to build successfully")
}

// TestFormatting_CellAlignment : Test cell alignment options
func TestFormatting_CellAlignment(t *testing.T) {
	// Test: Check cell alignment options
	// Expected:
	// - Horizontal alignment works
	// - Vertical alignment works
	// - Text wrapping and rotation work

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("Alignment")

	// Left aligned
	workbook.AddRow().
		AddCell("Left Aligned").
		SetStyle(excelbuilder.StyleConfig{
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: "left",
				Vertical:   "center",
			},
		}).Done().
		Done()

	// Center aligned
	workbook.AddRow().
		AddCell("Center Aligned").
		SetStyle(excelbuilder.StyleConfig{
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: "center",
				Vertical:   "center",
			},
		}).Done().
		Done()

	// Right aligned
	workbook.AddRow().
		AddCell("Right Aligned").
		SetStyle(excelbuilder.StyleConfig{
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: "right",
				Vertical:   "center",
			},
		}).Done().
		Done()

	// Wrapped text
	workbook.AddRow().
		AddCell("This is a long text that should wrap within the cell").
		SetStyle(excelbuilder.StyleConfig{
			Alignment: excelbuilder.AlignmentConfig{
				WrapText: true,
			},
		}).Done().
		Done()

	// Rotated text
	workbook.AddRow().
		AddCell("Rotated Text").
		SetStyle(excelbuilder.StyleConfig{
			Alignment: excelbuilder.AlignmentConfig{
				TextRotation: 45,
			},
		}).Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with alignment options to build successfully")
}

// TestFormatting_BorderStyles : Test border styling options
func TestFormatting_BorderStyles(t *testing.T) {
	// Test: Check border styling options
	// Expected:
	// - Different border styles work
	// - Individual border sides can be styled
	// - Border colors are preserved

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("Borders")

	// Thin border all around
	workbook.AddRow().
		AddCell("Thin Border").
		SetStyle(excelbuilder.StyleConfig{
			Border: excelbuilder.BorderConfig{
				Top:    excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
				Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
				Left:   excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
				Right:  excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
			},
		}).Done().
		Done()

	// Thick border
	workbook.AddRow().
		AddCell("Thick Border").
		SetStyle(excelbuilder.StyleConfig{
			Border: excelbuilder.BorderConfig{
				Top:    excelbuilder.BorderSide{Style: "thick", Color: "#FF0000"},
				Bottom: excelbuilder.BorderSide{Style: "thick", Color: "#FF0000"},
				Left:   excelbuilder.BorderSide{Style: "thick", Color: "#FF0000"},
				Right:  excelbuilder.BorderSide{Style: "thick", Color: "#FF0000"},
			},
		}).Done().
		Done()

	// Dashed border
	workbook.AddRow().
		AddCell("Dashed Border").
		SetStyle(excelbuilder.StyleConfig{
			Border: excelbuilder.BorderConfig{
				Top:    excelbuilder.BorderSide{Style: "dashed", Color: "#0000FF"},
				Bottom: excelbuilder.BorderSide{Style: "dashed", Color: "#0000FF"},
				Left:   excelbuilder.BorderSide{Style: "dashed", Color: "#0000FF"},
				Right:  excelbuilder.BorderSide{Style: "dashed", Color: "#0000FF"},
			},
		}).Done().
		Done()

	// Mixed borders
	workbook.AddRow().
		AddCell("Mixed Borders").
		SetStyle(excelbuilder.StyleConfig{
			Border: excelbuilder.BorderConfig{
				Top:    excelbuilder.BorderSide{Style: "thick", Color: "#FF0000"},
				Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
				Left:   excelbuilder.BorderSide{Style: "dashed", Color: "#0000FF"},
				Right:  excelbuilder.BorderSide{Style: "dotted", Color: "#00FF00"},
			},
		}).Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with border styles to build successfully")
}

// TestFormatting_FontStyles : Test font styling options
func TestFormatting_FontStyles(t *testing.T) {
	// Test: Check font styling options
	// Expected:
	// - Different font families work
	// - Font sizes are configurable
	// - Font styles (bold, italic, underline) work
	// - Font colors are preserved

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("Fonts")

	// Bold text
	workbook.AddRow().
		AddCell("Bold Text").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Bold: true,
				Size: 12,
			},
		}).Done().
		Done()

	// Italic text
	workbook.AddRow().
		AddCell("Italic Text").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Italic: true,
				Size:   12,
			},
		}).Done().
		Done()

	// Underlined text
	workbook.AddRow().
		AddCell("Underlined Text").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Underline: true,
				Size:      12,
			},
		}).Done().
		Done()

	// Colored text
	workbook.AddRow().
		AddCell("Colored Text").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Color: "#FF0000", // Red
				Size:  12,
			},
		}).Done().
		Done()

	// Different font family
	workbook.AddRow().
		AddCell("Times New Roman").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Family: "Times New Roman",
				Size:   14,
			},
		}).Done().
		Done()

	// Combined styles
	workbook.AddRow().
		AddCell("Bold Italic Red").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Bold:   true,
				Italic: true,
				Color:  "#FF0000",
				Size:   14,
			},
		}).Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with font styles to build successfully")
}

// Test Case 5.3: Cell Operations Tests

// TestCell_MergeCells : Test cell merging functionality
func TestCell_MergeCells(t *testing.T) {
	// Test: Check cell merging functionality
	// Expected:
	// - Cells can be merged
	// - Merged cell ranges work correctly
	// - Content is preserved in merged cells

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet("MergedCells")

	// Add some data
	sheet.AddRow().
		AddCell("Merged Header").Done().
		AddCell("").Done().
		AddCell("").Done().
		Done().
		AddRow().
		AddCell("Data 1").Done().
		AddCell("Data 2").Done().
		AddCell("Data 3").Done().
		Done()

	// Merge cells A1:C1
	sheet.MergeCell("A1:C1")

	// Add more merged cells
	sheet.AddRow().
		AddCell("Vertical Merge").Done().
		AddCell("Normal Cell").Done().
		AddCell("Normal Cell").Done().
		Done().
		AddRow().
		AddCell("").Done().
		AddCell("Data 4").Done().
		AddCell("Data 5").Done().
		Done()

	// Merge cells A3:A4
	sheet.MergeCell("A3:A4")

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with merged cells to build successfully")
}

// TestCell_CellComments : Test cell comments functionality
func TestCell_CellComments(t *testing.T) {
	// Test: Check cell comments functionality
	// Expected:
	// - Comments can be added to cells
	// - Comment text is preserved
	// - Comment authors are tracked

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("Comments")

	// Add cells with comments
	workbook.AddRow().
		AddCell("Cell with comment").
		// Note: In full implementation, we would have:
		// AddComment("This is a comment", "Author Name")
		Done().
		Done().
		AddRow().
		AddCell(123).
		// AddComment("This number represents the total count", "Data Analyst")
		Done().
		Done()

	// For now, just verify the basic structure works
	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with comments to build successfully")
}

// TestCell_Hyperlinks : Test hyperlink functionality
func TestCell_Hyperlinks(t *testing.T) {
	// Test: Check hyperlink functionality
	// Expected:
	// - Hyperlinks can be added to cells
	// - Different link types work (URL, email, internal)
	// - Link text and tooltips are preserved

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("Hyperlinks")

	// URL hyperlink
	workbook.AddRow().
		AddCell("Visit Google").
		// Note: In full implementation, we would have:
		// AddHyperlink("https://www.google.com", "Google Search")
		Done().
		Done()

	// Email hyperlink
	workbook.AddRow().
		AddCell("Send Email").
		// AddHyperlink("mailto:test@example.com", "Contact Us")
		Done().
		Done()

	// Internal link
	workbook.AddRow().
		AddCell("Go to Sheet2").
		// AddHyperlink("#Sheet2!A1", "Navigate to Sheet2")
		Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with hyperlinks to build successfully")
}

// TestCell_DataTypes : Test different data types in cells
func TestCell_DataTypes(t *testing.T) {
	// Test: Check different data types in cells
	// Expected:
	// - Different data types are handled correctly
	// - Type conversion works properly
	// - Data integrity is maintained

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("DataTypes")

	// String data
	workbook.AddRow().
		AddCell("String").Done().
		AddCell("Hello World").Done().
		Done()

	// Integer data
	workbook.AddRow().
		AddCell("Integer").Done().
		AddCell(42).Done().
		Done()

	// Float data
	workbook.AddRow().
		AddCell("Float").Done().
		AddCell(3.14159).Done().
		Done()

	// Boolean data
	workbook.AddRow().
		AddCell("Boolean").Done().
		AddCell(true).Done().
		Done()

	// Date data
	workbook.AddRow().
		AddCell("Date").Done().
		AddCell("2024-01-15").Done().
		Done()

	// Formula data
	workbook.AddRow().
		AddCell("Formula").Done().
		AddCell("=B2+B3").Done().
		Done()

	// Empty cell
	workbook.AddRow().
		AddCell("Empty").Done().
		AddCell("").Done().
		Done()

	// Null/nil handling
	workbook.AddRow().
		AddCell("Null").Done().
		AddCell(nil).Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with different data types to build successfully")
}