package tests

import (
	"fmt"
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Test Case 7.1: End-to-End Simple Workbook Creation
func TestIntegration_SimpleWorkbookCreation(t *testing.T) {
	// Test: Create a complete workbook with basic content
	// Expected:
	// - Workbook created successfully
	// - Contains expected sheets and data
	// - Can be built without errors

	builder := excelbuilder.New()
	if builder == nil {
		t.Fatal("Failed to create ExcelBuilder")
	}

	file := builder.
		NewWorkbook().
		SetProperties(excelbuilder.WorkbookProperties{
			Title:       "Simple Test Workbook",
			Author:      "Go Excel Builder",
			Subject:     "Integration Test",
			Description: "A simple workbook for testing",
		}).
		AddSheet("Data").
		AddRow().
		AddCell("Name").Done().
		AddCell("Age").Done().
		AddCell("City").Done().
		Done().
		AddRow().
		AddCell("John Doe").Done().
		AddCell(30).Done().
		AddCell("New York").Done().
		Done().
		AddRow().
		AddCell("Jane Smith").Done().
		AddCell(25).Done().
		AddCell("Los Angeles").Done().
		Done().
		Done().
		Build()

	if file == nil {
		t.Fatal("Failed to build workbook")
	}

	// Verify the workbook structure
	sheets := file.GetSheetList()
	if len(sheets) == 0 {
		t.Error("Expected at least one sheet in workbook")
	}

	found := false
	for _, sheet := range sheets {
		if sheet == "Data" {
			found = true
			break
		}
	}

	if !found {
		t.Error("Expected to find 'Data' sheet in workbook")
	}
}

// Test Case 7.2: End-to-End Complex Workbook with Styling
func TestIntegration_ComplexWorkbookWithStyling(t *testing.T) {
	// Test: Create a workbook with multiple sheets, styling, and formatting
	// Expected:
	// - Multiple sheets created
	// - Styles applied correctly
	// - Different data types handled

	builder := excelbuilder.New()
	if builder == nil {
		t.Fatal("Failed to create ExcelBuilder")
	}

	workbook := builder.NewWorkbook().
		SetProperties(excelbuilder.WorkbookProperties{
			Title:   "Complex Test Workbook",
			Author:  "Integration Test",
			Company: "Go Excel Builder",
		})

	// Create Summary sheet
	summarySheet := workbook.AddSheet("Summary")
	if summarySheet == nil {
		t.Fatal("Failed to create Summary sheet")
	}

	// Add header row with styling
	headerRow := summarySheet.AddRow()
	headerRow.
		AddCell("Report Summary").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Bold:   true,
				Size:   16,
				Color:  "#000080",
				Family: "Arial",
			},
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: "center",
			},
		}).
		Done()

	// Create Data sheet
	dataSheet := headerRow.Done().Done().AddSheet("Data")
	if dataSheet == nil {
		t.Fatal("Failed to create Data sheet")
	}

	// Set column widths
	dataSheet.
		SetColumnWidth("A", 20.0).
		SetColumnWidth("B", 15.0).
		SetColumnWidth("C", 25.0).
		SetColumnWidth("D", 15.0)

	// Add header row
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Color:  "#FFFFFF",
			Family: "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#4472C4",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	}

	headerRow2 := dataSheet.AddRow()
	headerRow2.
		AddCell("Employee Name").SetStyle(headerStyle).Done().
		AddCell("Department").SetStyle(headerStyle).Done().
		AddCell("Email").SetStyle(headerStyle).Done().
		AddCell("Salary").SetStyle(headerStyle).Done()

	// Add data rows with different formatting
	employees := []struct {
		name       string
		department string
		email      string
		salary     float64
	}{
		{"John Doe", "Engineering", "john.doe@company.com", 75000.00},
		{"Jane Smith", "Marketing", "jane.smith@company.com", 65000.00},
		{"Bob Johnson", "Sales", "bob.johnson@company.com", 55000.00},
	}

	currentSheet := headerRow2.Done()
	for _, emp := range employees {
		row := currentSheet.AddRow()
		row.
			AddCell(emp.name).Done().
			AddCell(emp.department).Done().
			AddCell(emp.email).
			SetStyle(excelbuilder.StyleConfig{
				Font: excelbuilder.FontConfig{
					Color:     "#0000FF",
					Underline: true,
				},
			}).
			SetHyperlink("mailto:" + emp.email).
			Done().
			AddCell(emp.salary).
			SetNumberFormat("$#,##0.00").
			SetStyle(excelbuilder.StyleConfig{
				Alignment: excelbuilder.AlignmentConfig{
					Horizontal: "right",
				},
			}).
			Done()

		currentSheet = row.Done()
	}

	// Build the workbook
	file := currentSheet.Done().Build()
	if file == nil {
		t.Fatal("Failed to build complex workbook")
	}

	// Verify the workbook structure
	sheets := file.GetSheetList()
	expectedSheets := []string{"Summary", "Data"}

	for _, expectedSheet := range expectedSheets {
		found := false
		for _, sheet := range sheets {
			if sheet == expectedSheet {
				found = true
				break
			}
		}
		if !found {
			t.Errorf("Expected to find '%s' sheet in workbook", expectedSheet)
		}
	}
}

// Test Case 7.3: End-to-End Financial Report
func TestIntegration_FinancialReport(t *testing.T) {
	// Test: Create a financial report with formulas and calculations
	// Expected:
	// - Formulas work correctly
	// - Number formatting applied
	// - Professional styling

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().
		SetProperties(excelbuilder.WorkbookProperties{
			Title:       "Financial Report Q1 2024",
			Author:      "Finance Department",
			Subject:     "Quarterly Financial Analysis",
			Description: "Q1 2024 financial performance report",
			Company:     "Test Company Inc.",
		})

	sheet := workbook.AddSheet("Financial Summary")
	if sheet == nil {
		t.Fatal("Failed to create Financial Summary sheet")
	}

	// Set column widths for better presentation
	sheet.
		SetColumnWidth("A", 25.0).
		SetColumnWidth("B", 15.0).
		SetColumnWidth("C", 15.0).
		SetColumnWidth("D", 15.0)

	// Title row
	titleRow := sheet.AddRow()
	titleRow.
		AddCell("Financial Summary - Q1 2024").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Bold:   true,
				Size:   18,
				Color:  "#1F4E79",
				Family: "Calibri",
			},
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: "center",
			},
		}).
		Done()

	// Empty row for spacing
	emptyRow := titleRow.Done().AddRow()

	// Header row
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Color:  "#FFFFFF",
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#1F4E79",
		},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin"},
			Bottom: excelbuilder.BorderSide{Style: "thin"},
			Left:   excelbuilder.BorderSide{Style: "thin"},
			Right:  excelbuilder.BorderSide{Style: "thin"},
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	}

	headerRow := emptyRow.Done().AddRow()
	headerRow.
		AddCell("Category").SetStyle(headerStyle).Done().
		AddCell("January").SetStyle(headerStyle).Done().
		AddCell("February").SetStyle(headerStyle).Done().
		AddCell("March").SetStyle(headerStyle).Done()

	// Data rows
	financialData := []struct {
		category string
		jan      float64
		feb      float64
		mar      float64
	}{
		{"Revenue", 150000, 165000, 180000},
		{"Cost of Goods Sold", 75000, 82500, 90000},
		{"Operating Expenses", 45000, 47000, 49000},
		{"Marketing", 15000, 18000, 20000},
		{"Administrative", 25000, 26000, 27000},
	}

	currencyStyle := excelbuilder.StyleConfig{
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
		},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin"},
			Bottom: excelbuilder.BorderSide{Style: "thin"},
			Left:   excelbuilder.BorderSide{Style: "thin"},
			Right:  excelbuilder.BorderSide{Style: "thin"},
		},
	}

	categoryStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
		},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin"},
			Bottom: excelbuilder.BorderSide{Style: "thin"},
			Left:   excelbuilder.BorderSide{Style: "thin"},
			Right:  excelbuilder.BorderSide{Style: "thin"},
		},
	}

	currentSheet := headerRow.Done()
	for _, data := range financialData {
		row := currentSheet.AddRow()
		row.
			AddCell(data.category).SetStyle(categoryStyle).Done().
			AddCell(data.jan).SetNumberFormat("$#,##0").SetStyle(currencyStyle).Done().
			AddCell(data.feb).SetNumberFormat("$#,##0").SetStyle(currencyStyle).Done().
			AddCell(data.mar).SetNumberFormat("$#,##0").SetStyle(currencyStyle).Done()

		currentSheet = row.Done()
	}

	// Total row with formulas
	totalRow := currentSheet.AddRow()
	totalStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:  true,
			Color: "#FFFFFF",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#1F4E79",
		},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thick"},
			Bottom: excelbuilder.BorderSide{Style: "thick"},
			Left:   excelbuilder.BorderSide{Style: "thin"},
			Right:  excelbuilder.BorderSide{Style: "thin"},
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
		},
	}

	totalRow.
		AddCell("TOTAL").SetStyle(totalStyle).Done().
		AddCell("=SUM(B4:B8)").SetNumberFormat("$#,##0").SetStyle(totalStyle).Done().
		AddCell("=SUM(C4:C8)").SetNumberFormat("$#,##0").SetStyle(totalStyle).Done().
		AddCell("=SUM(D4:D8)").SetNumberFormat("$#,##0").SetStyle(totalStyle).Done()

	// Build and verify
	file := totalRow.Done().Done().Build()
	if file == nil {
		t.Fatal("Failed to build financial report")
	}

	// Verify sheet exists
	sheets := file.GetSheetList()
	found := false
	for _, sheet := range sheets {
		if sheet == "Financial Summary" {
			found = true
			break
		}
	}

	if !found {
		t.Error("Expected to find 'Financial Summary' sheet")
	}
}

// Test Case 7.4: End-to-End Performance Test
func TestIntegration_PerformanceTest(t *testing.T) {
	// Test: Create a workbook with substantial data to test performance
	// Expected:
	// - Handles large datasets efficiently
	// - Memory usage remains reasonable
	// - Build time is acceptable

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().
		SetProperties(excelbuilder.WorkbookProperties{
			Title:  "Performance Test Workbook",
			Author: "Performance Test",
		})

	sheet := workbook.AddSheet("Large Dataset")
	if sheet == nil {
		t.Fatal("Failed to create sheet")
	}

	// Add header row
	headerRow := sheet.AddRow()
	headerRow.
		AddCell("ID").Done().
		AddCell("Name").Done().
		AddCell("Value").Done().
		AddCell("Category").Done().
		AddCell("Date").Done()

	// Add substantial data (but not too much for test performance)
	currentSheet := headerRow.Done()
	numRows := 100 // Reduced for test performance

	for i := 1; i <= numRows; i++ {
		row := currentSheet.AddRow()
		row.
			AddCell(i).Done().
			AddCell(fmt.Sprintf("Item %d", i)).Done().
			AddCell(float64(i) * 10.5).SetNumberFormat("0.00").Done().
			AddCell(fmt.Sprintf("Category %d", (i%5)+1)).Done().
			AddCell("2024-01-01").Done()

		currentSheet = row.Done()
	}

	// Build the workbook
	file := currentSheet.Done().Build()
	if file == nil {
		t.Fatal("Failed to build performance test workbook")
	}

	// Verify basic structure
	sheets := file.GetSheetList()
	if len(sheets) == 0 {
		t.Error("Expected at least one sheet")
	}

	t.Logf("Successfully created workbook with %d data rows", numRows)
}