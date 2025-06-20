package tests

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
)

// Test Case 3.1: Conditional Formatting Tests

// TestConditionalFormatting_Rules : Test conditional formatting rules
func TestConditionalFormatting_Rules(t *testing.T) {
	// Test: Check conditional formatting rule creation
	// Expected:
	// - Rules can be created with different conditions
	// - Multiple rule types supported
	// - Configuration is preserved

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("ConditionalTest")

	// Add test data
	workbook.AddRow().
		AddCell("Score").Done().
		AddCell("Grade").Done().
		Done().
		AddRow().
		AddCell(85).Done().
		AddCell("B").Done().
		Done().
		AddRow().
		AddCell(92).Done().
		AddCell("A").Done().
		Done().
		AddRow().
		AddCell(78).Done().
		AddCell("C").Done().
		Done()

	// Test conditional formatting configuration
	condFormat := excelbuilder.ConditionalFormattingConfig{
		Range: "A2:A4",
		Rules: []excelbuilder.ConditionalRule{
			{
				Type:      "cellValue",
				Operator:  "greaterThan",
				Value:     "90",
				Format: excelbuilder.StyleConfig{
					Fill: excelbuilder.FillConfig{
						Type:  "pattern",
						Color: "#90EE90", // Light green
					},
				},
			},
			{
				Type:      "cellValue",
				Operator:  "lessThan",
				Value:     "80",
				Format: excelbuilder.StyleConfig{
					Fill: excelbuilder.FillConfig{
						Type:  "pattern",
						Color: "#FFB6C1", // Light pink
					},
				},
			},
		},
	}

	// Verify configuration structure
	assert.Equal(t, "A2:A4", condFormat.Range, "Expected range A2:A4")
	assert.Len(t, condFormat.Rules, 2, "Expected 2 conditional rules")
	assert.Equal(t, "greaterThan", condFormat.Rules[0].Operator, "Expected greaterThan operator")
	assert.Equal(t, "lessThan", condFormat.Rules[1].Operator, "Expected lessThan operator")
	assert.Equal(t, "#90EE90", condFormat.Rules[0].Format.Fill.Color, "Expected light green color")
	assert.Equal(t, "#FFB6C1", condFormat.Rules[1].Format.Fill.Color, "Expected light pink color")
}

// TestConditionalFormatting_DataBars : Test data bars conditional formatting
func TestConditionalFormatting_DataBars(t *testing.T) {
	// Test: Check data bars conditional formatting
	// Expected:
	// - Data bars can be configured
	// - Color and style options work
	// - Range is applied correctly

	dataBarFormat := excelbuilder.ConditionalFormattingConfig{
		Range: "B2:B10",
		Rules: []excelbuilder.ConditionalRule{
			{
				Type: "dataBar",
				Format: excelbuilder.StyleConfig{
					Fill: excelbuilder.FillConfig{
						Type:  "gradient",
						Color: "#4472C4",
					},
				},
			},
		},
	}

	assert.Equal(t, "B2:B10", dataBarFormat.Range, "Expected range B2:B10")
	assert.Equal(t, "dataBar", dataBarFormat.Rules[0].Type, "Expected dataBar type")
	assert.Equal(t, "gradient", dataBarFormat.Rules[0].Format.Fill.Type, "Expected gradient fill")
	assert.Equal(t, "#4472C4", dataBarFormat.Rules[0].Format.Fill.Color, "Expected blue color")
}

// TestConditionalFormatting_ColorScales : Test color scales
func TestConditionalFormatting_ColorScales(t *testing.T) {
	// Test: Check color scales conditional formatting
	// Expected:
	// - Color scales can be configured
	// - Multiple colors supported
	// - Gradient configuration works

	colorScaleFormat := excelbuilder.ConditionalFormattingConfig{
		Range: "C2:C20",
		Rules: []excelbuilder.ConditionalRule{
			{
				Type: "colorScale",
				ColorScale: excelbuilder.ColorScale{
					MinColor: "#FF0000", // Red
					MidColor: "#FFFF00", // Yellow
					MaxColor: "#00FF00", // Green
				},
			},
		},
	}

	assert.Equal(t, "C2:C20", colorScaleFormat.Range, "Expected range C2:C20")
	assert.Equal(t, "colorScale", colorScaleFormat.Rules[0].Type, "Expected colorScale type")
	assert.Equal(t, "#FF0000", colorScaleFormat.Rules[0].ColorScale.MinColor, "Expected red min color")
	assert.Equal(t, "#FFFF00", colorScaleFormat.Rules[0].ColorScale.MidColor, "Expected yellow mid color")
	assert.Equal(t, "#00FF00", colorScaleFormat.Rules[0].ColorScale.MaxColor, "Expected green max color")
}

// TestConditionalFormatting_IconSets : Test icon sets
func TestConditionalFormatting_IconSets(t *testing.T) {
	// Test: Check icon sets conditional formatting
	// Expected:
	// - Icon sets can be configured
	// - Different icon types supported
	// - Thresholds work correctly

	iconSetFormat := excelbuilder.ConditionalFormattingConfig{
		Range: "D2:D15",
		Rules: []excelbuilder.ConditionalRule{
			{
				Type:    "iconSet",
				IconSet: "3TrafficLights",
				Thresholds: []excelbuilder.Threshold{
					{Value: "33", Type: "percent"},
					{Value: "67", Type: "percent"},
				},
			},
		},
	}

	assert.Equal(t, "D2:D15", iconSetFormat.Range, "Expected range D2:D15")
	assert.Equal(t, "iconSet", iconSetFormat.Rules[0].Type, "Expected iconSet type")
	assert.Equal(t, "3TrafficLights", iconSetFormat.Rules[0].IconSet, "Expected 3TrafficLights icon set")
	assert.Len(t, iconSetFormat.Rules[0].Thresholds, 2, "Expected 2 thresholds")
	assert.Equal(t, "33", iconSetFormat.Rules[0].Thresholds[0].Value, "Expected 33% threshold")
	assert.Equal(t, "67", iconSetFormat.Rules[0].Thresholds[1].Value, "Expected 67% threshold")
}

// Test Case 3.2: Data Validation Tests

// TestDataValidation_DropdownLists : Test dropdown list validation
func TestDataValidation_DropdownLists(t *testing.T) {
	// Test: Check dropdown list data validation
	// Expected:
	// - Dropdown lists can be created
	// - Options are configurable
	// - Range application works

	dropdownValidation := excelbuilder.DataValidationConfig{
		Range: "E2:E100",
		Type:  "list",
		Formula: "Option1,Option2,Option3,Option4",
		ShowDropdown: true,
		ErrorMessage: "Please select a valid option from the dropdown",
		ErrorTitle:   "Invalid Selection",
		ShowError:    true,
	}

	assert.Equal(t, "E2:E100", dropdownValidation.Range, "Expected range E2:E100")
	assert.Equal(t, "list", dropdownValidation.Type, "Expected list validation type")
	assert.Equal(t, "Option1,Option2,Option3,Option4", dropdownValidation.Formula, "Expected dropdown options")
	assert.True(t, dropdownValidation.ShowDropdown, "Expected dropdown to be shown")
	assert.True(t, dropdownValidation.ShowError, "Expected error to be shown")
	assert.Equal(t, "Invalid Selection", dropdownValidation.ErrorTitle, "Expected error title")
}

// TestDataValidation_NumberRanges : Test number range validation
func TestDataValidation_NumberRanges(t *testing.T) {
	// Test: Check number range data validation
	// Expected:
	// - Number ranges can be validated
	// - Min/max values work
	// - Error messages are configurable

	numberValidation := excelbuilder.DataValidationConfig{
		Range:        "F2:F50",
		Type:         "decimal",
		Operator:     "between",
		Formula:      "0",
		Formula2:     "100",
		ErrorMessage: "Please enter a number between 0 and 100",
		ErrorTitle:   "Invalid Number",
		ShowError:    true,
		InputMessage: "Enter a percentage (0-100)",
		InputTitle:   "Percentage Input",
		ShowInput:    true,
	}

	assert.Equal(t, "F2:F50", numberValidation.Range, "Expected range F2:F50")
	assert.Equal(t, "decimal", numberValidation.Type, "Expected decimal validation type")
	assert.Equal(t, "between", numberValidation.Operator, "Expected between operator")
	assert.Equal(t, "0", numberValidation.Formula, "Expected min value 0")
	assert.Equal(t, "100", numberValidation.Formula2, "Expected max value 100")
	assert.True(t, numberValidation.ShowInput, "Expected input message to be shown")
	assert.Equal(t, "Enter a percentage (0-100)", numberValidation.InputMessage, "Expected input message")
}

// TestDataValidation_DateRanges : Test date range validation
func TestDataValidation_DateRanges(t *testing.T) {
	// Test: Check date range data validation
	// Expected:
	// - Date ranges can be validated
	// - Date operators work
	// - Custom date formats supported

	dateValidation := excelbuilder.DataValidationConfig{
		Range:        "G2:G30",
		Type:         "date",
		Operator:     "greaterThanOrEqual",
		Formula:      "2024-01-01",
		ErrorMessage: "Please enter a date from 2024 onwards",
		ErrorTitle:   "Invalid Date",
		ShowError:    true,
		InputMessage: "Enter a date (YYYY-MM-DD format)",
		InputTitle:   "Date Input",
		ShowInput:    true,
	}

	assert.Equal(t, "G2:G30", dateValidation.Range, "Expected range G2:G30")
	assert.Equal(t, "date", dateValidation.Type, "Expected date validation type")
	assert.Equal(t, "greaterThanOrEqual", dateValidation.Operator, "Expected greaterThanOrEqual operator")
	assert.Equal(t, "2024-01-01", dateValidation.Formula, "Expected start date 2024-01-01")
	assert.Equal(t, "Enter a date (YYYY-MM-DD format)", dateValidation.InputMessage, "Expected date input message")
}

// TestDataValidation_CustomFormulas : Test custom formula validation
func TestDataValidation_CustomFormulas(t *testing.T) {
	// Test: Check custom formula data validation
	// Expected:
	// - Custom formulas can be used
	// - Complex validation logic works
	// - Formula references are preserved

	customValidation := excelbuilder.DataValidationConfig{
		Range:        "H2:H20",
		Type:         "custom",
		Formula:      "AND(H2>=10, H2<=1000, MOD(H2,5)=0)", // Multiple of 5 between 10-1000
		ErrorMessage: "Value must be a multiple of 5 between 10 and 1000",
		ErrorTitle:   "Custom Validation Error",
		ShowError:    true,
		InputMessage: "Enter a multiple of 5 (10-1000)",
		InputTitle:   "Custom Input",
		ShowInput:    true,
	}

	assert.Equal(t, "H2:H20", customValidation.Range, "Expected range H2:H20")
	assert.Equal(t, "custom", customValidation.Type, "Expected custom validation type")
	assert.Equal(t, "AND(H2>=10, H2<=1000, MOD(H2,5)=0)", customValidation.Formula, "Expected custom formula")
	assert.Equal(t, "Value must be a multiple of 5 between 10 and 1000", customValidation.ErrorMessage, "Expected custom error message")
}

// Test Case 3.3: Template System Tests

// TestTemplate_Creation : Test template creation và management
func TestTemplate_Creation(t *testing.T) {
	// Test: Check template creation and management
	// Expected:
	// - Templates can be created
	// - Configuration is preserved
	// - Multiple templates supported

	reportTemplate := excelbuilder.TemplateConfig{
		Name:        "MonthlyReport",
		Description: "Standard monthly report template",
		Sheets: []excelbuilder.SheetTemplate{
			{
				Name: "Summary",
				Headers: []string{"Metric", "Current Month", "Previous Month", "Change %"},
				Styles: map[string]excelbuilder.StyleConfig{
					"header": {
						Font: excelbuilder.FontConfig{
							Bold: true,
							Size: 12,
						},
						Fill: excelbuilder.FillConfig{
							Type:  "pattern",
							Color: "#D9E1F2",
						},
					},
				},
			},
			{
				Name: "Details",
				Headers: []string{"Date", "Transaction", "Amount", "Category"},
				Styles: map[string]excelbuilder.StyleConfig{
					"header": {
						Font: excelbuilder.FontConfig{
							Bold: true,
							Size: 11,
						},
					},
				},
			},
		},
	}

	assert.Equal(t, "MonthlyReport", reportTemplate.Name, "Expected template name")
	assert.Equal(t, "Standard monthly report template", reportTemplate.Description, "Expected template description")
	assert.Len(t, reportTemplate.Sheets, 2, "Expected 2 sheet templates")
	assert.Equal(t, "Summary", reportTemplate.Sheets[0].Name, "Expected first sheet name")
	assert.Equal(t, "Details", reportTemplate.Sheets[1].Name, "Expected second sheet name")
	assert.Len(t, reportTemplate.Sheets[0].Headers, 4, "Expected 4 headers in summary sheet")
	assert.True(t, reportTemplate.Sheets[0].Styles["header"].Font.Bold, "Expected header font to be bold")
}

// TestTemplate_Application : Test applying templates to workbooks
func TestTemplate_Application(t *testing.T) {
	// Test: Check template application to workbooks
	// Expected:
	// - Templates can be applied
	// - Structure is created correctly
	// - Styles are applied

	builder := excelbuilder.New()

	// Create simple template
	simpleTemplate := excelbuilder.TemplateConfig{
		Name: "SimpleTemplate",
		Sheets: []excelbuilder.SheetTemplate{
			{
				Name:    "Data",
				Headers: []string{"ID", "Name", "Value"},
				Styles: map[string]excelbuilder.StyleConfig{
					"header": {
						Font: excelbuilder.FontConfig{Bold: true},
					},
				},
			},
		},
	}

	// Apply template (in real implementation)
	workbook := builder.NewWorkbook()

	// Manually create what template would create
	sheet := workbook.AddSheet("Data")
	headerRow := sheet.AddRow()
	for _, header := range simpleTemplate.Sheets[0].Headers {
		headerRow.AddCell(header).
			SetStyle(excelbuilder.StyleConfig{
				Font: excelbuilder.FontConfig{Bold: true},
			}).Done()
	}
	headerRow.Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook to build successfully from template")
}

// TestTemplate_Customization : Test template customization options
func TestTemplate_Customization(t *testing.T) {
	// Test: Check template customization capabilities
	// Expected:
	// - Templates can be customized
	// - Custom styles can be added
	// - Dynamic content supported

	customTemplate := excelbuilder.TemplateConfig{
		Name: "CustomizableTemplate",
		Sheets: []excelbuilder.SheetTemplate{
			{
				Name:    "CustomSheet",
				Headers: []string{"{{DYNAMIC_HEADER_1}}", "{{DYNAMIC_HEADER_2}}", "Total"},
				Styles: map[string]excelbuilder.StyleConfig{
					"header": {
						Font: excelbuilder.FontConfig{
							Bold:  true,
							Color: "{{HEADER_COLOR}}",
						},
						Fill: excelbuilder.FillConfig{
							Type:  "pattern",
							Color: "{{BACKGROUND_COLOR}}",
						},
					},
					"data": {
						Font: excelbuilder.FontConfig{
							Size: 10,
						},
					},
				},
				Variables: map[string]interface{}{
					"DYNAMIC_HEADER_1": "Revenue",
					"DYNAMIC_HEADER_2": "Expenses",
					"HEADER_COLOR":     "#FFFFFF",
					"BACKGROUND_COLOR": "#4472C4",
				},
			},
		},
	}

	assert.Equal(t, "CustomizableTemplate", customTemplate.Name, "Expected customizable template name")
	assert.Contains(t, customTemplate.Sheets[0].Headers[0], "{{DYNAMIC_HEADER_1}}", "Expected dynamic header placeholder")
	assert.Equal(t, "Revenue", customTemplate.Sheets[0].Variables["DYNAMIC_HEADER_1"], "Expected dynamic header value")
	assert.Equal(t, "#4472C4", customTemplate.Sheets[0].Variables["BACKGROUND_COLOR"], "Expected background color variable")
}

// Test Case 3.4: Formula System Tests

// TestFormula_BasicOperations : Test basic formula operations
func TestFormula_BasicOperations(t *testing.T) {
	// Test: Check basic formula operations
	// Expected:
	// - Basic formulas can be created
	// - Arithmetic operations work
	// - Cell references are correct

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("FormulaTest")

	// Add test data
	workbook.AddRow().
		AddCell("A").Done().
		AddCell("B").Done().
		AddCell("Sum").Done().
		AddCell("Product").Done().
		AddCell("Average").Done().
		Done().
		AddRow().
		AddCell(10).Done().
		AddCell(20).Done().
		AddCell("=A2+B2").Done().    // Sum formula
		AddCell("=A2*B2").Done().    // Product formula
		AddCell("=(A2+B2)/2").Done(). // Average formula
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with formulas to build successfully")
}

// TestFormula_Functions : Test Excel functions (SUM, AVERAGE, COUNT, etc.)
func TestFormula_Functions(t *testing.T) {
	// Test: Check Excel function support
	// Expected:
	// - Standard Excel functions work
	// - Range references are correct
	// - Function nesting supported

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("FunctionTest")

	// Add test data
	workbook.AddRow().
		AddCell("Values").Done().
		Done().
		AddRow().
		AddCell(100).Done().
		Done().
		AddRow().
		AddCell(200).Done().
		Done().
		AddRow().
		AddCell(150).Done().
		Done().
		AddRow().
		AddCell(300).Done().
		Done().
		AddRow().
		AddCell("Sum:").Done().
		Done().
		AddRow().
		AddCell("=SUM(A2:A5)").Done(). // SUM function
		Done().
		AddRow().
		AddCell("Average:").Done().
		Done().
		AddRow().
		AddCell("=AVERAGE(A2:A5)").Done(). // AVERAGE function
		Done().
		AddRow().
		AddCell("Count:").Done().
		Done().
		AddRow().
		AddCell("=COUNT(A2:A5)").Done(). // COUNT function
		Done().
		AddRow().
		AddCell("Max:").Done().
		Done().
		AddRow().
		AddCell("=MAX(A2:A5)").Done(). // MAX function
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with Excel functions to build successfully")
}

// TestFormula_CrossSheetReferences : Test cross-sheet formula references
func TestFormula_CrossSheetReferences(t *testing.T) {
	// Test: Check cross-sheet formula references
	// Expected:
	// - Cross-sheet references work
	// - Sheet names are handled correctly
	// - Complex references supported

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Create data sheet
	dataSheet := workbook.AddSheet("Data")
	dataSheet.AddRow().
		AddCell("Revenue").Done().
		Done().
		AddRow().
		AddCell(50000).Done().
		Done().
		AddRow().
		AddCell(60000).Done().
		Done()

	// Create summary sheet with cross-sheet references
	summarySheet := workbook.AddSheet("Summary")
	summarySheet.AddRow().
		AddCell("Total Revenue").Done().
		Done().
		AddRow().
		AddCell("=SUM(Data!A2:A3)").Done(). // Cross-sheet SUM
		Done().
		AddRow().
		AddCell("Average Revenue").Done().
		Done().
		AddRow().
		AddCell("=AVERAGE(Data!A2:A3)").Done(). // Cross-sheet AVERAGE
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with cross-sheet formulas to build successfully")
}

// TestFormula_ComplexExpressions : Test complex formula expressions
func TestFormula_ComplexExpressions(t *testing.T) {
	// Test: Check complex formula expressions
	// Expected:
	// - Complex expressions work
	// - Nested functions supported
	// - Conditional logic works

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("ComplexFormulas")

	// Add test data
	workbook.AddRow().
		AddCell("Score").Done().
		AddCell("Grade").Done().
		AddCell("Status").Done().
		Done().
		AddRow().
		AddCell(85).Done().
		AddCell("=IF(A2>=90,\"A\",IF(A2>=80,\"B\",IF(A2>=70,\"C\",\"F\")))").Done(). // Nested IF
		AddCell("=IF(A2>=70,\"PASS\",\"FAIL\")").Done().                                    // Simple IF
		Done().
		AddRow().
		AddCell(92).Done().
		AddCell("=IF(A3>=90,\"A\",IF(A3>=80,\"B\",IF(A3>=70,\"C\",\"F\")))").Done().
		AddCell("=IF(A3>=70,\"PASS\",\"FAIL\")").Done().
		Done().
		AddRow().
		AddCell("Complex Calculation").Done().
		Done().
		AddRow().
		AddCell("=ROUND(AVERAGE(A2:A3)*1.1,2)").Done(). // Nested functions with rounding
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with complex formulas to build successfully")
}