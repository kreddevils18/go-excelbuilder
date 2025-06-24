package main

import (
	"fmt"
	"log"
	"os"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

func main() {
	// Create output directory if it doesn't exist
	if err := os.MkdirAll("output", 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}

	// Create a new Excel builder instance
	builder := excelbuilder.New()

	// Create a new workbook with comprehensive properties
	workbook := builder.NewWorkbook()
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Enhanced Basic Example",
		Author:      "Go Excel Builder",
		Subject:     "Demonstration of Basic Features",
		Description: "This example showcases enhanced basic operations including multiple sheets, styling, and data types",
		Company:     "Your Company",
		Category:    "Example",
		Keywords:    "go,excel,builder,basic,enhanced",
		Comments:    "Generated on " + time.Now().Format("2006-01-02 15:04:05"),
	})

	// Define reusable styles for consistent formatting
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
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
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
		},
	}

	dataStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
		},
	}

	numberStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "#,##0.00",
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
		},
	}

	// Sheet 1: Employee Data
	employeeSheet := workbook.AddSheet("Employee Data")
	if employeeSheet == nil {
		log.Fatal("Failed to create employee sheet")
	}

	// Set column widths for better presentation
	employeeSheet.SetColumnWidth("A", 15.0) // Employee ID
	employeeSheet.SetColumnWidth("B", 25.0) // Full Name
	employeeSheet.SetColumnWidth("C", 30.0) // Email
	employeeSheet.SetColumnWidth("D", 15.0) // Department
	employeeSheet.SetColumnWidth("E", 12.0) // Age
	employeeSheet.SetColumnWidth("F", 15.0) // Salary
	employeeSheet.SetColumnWidth("G", 15.0) // Start Date

	// Add header row with styling
	headerRow := employeeSheet.AddRow()
	headerRow.AddCell("Employee ID").SetStyle(headerStyle)
	headerRow.AddCell("Full Name").SetStyle(headerStyle)
	headerRow.AddCell("Email").SetStyle(headerStyle)
	headerRow.AddCell("Department").SetStyle(headerStyle)
	headerRow.AddCell("Age").SetStyle(headerStyle)
	headerRow.AddCell("Salary").SetStyle(headerStyle)
	headerRow.AddCell("Start Date").SetStyle(headerStyle)

	// Sample employee data
	employees := [][]interface{}{
		{"EMP001", "John Doe", "john.doe@company.com", "Engineering", 30, 75000.00, "2022-01-15"},
		{"EMP002", "Jane Smith", "jane.smith@company.com", "Marketing", 28, 65000.00, "2022-03-01"},
		{"EMP003", "Bob Johnson", "bob.johnson@company.com", "Sales", 35, 70000.00, "2021-11-10"},
		{"EMP004", "Alice Brown", "alice.brown@company.com", "HR", 32, 68000.00, "2022-02-20"},
		{"EMP005", "Charlie Wilson", "charlie.wilson@company.com", "Engineering", 29, 78000.00, "2022-04-05"},
		{"EMP006", "Diana Davis", "diana.davis@company.com", "Finance", 31, 72000.00, "2021-12-15"},
		{"EMP007", "Eva Martinez", "eva.martinez@company.com", "Marketing", 27, 63000.00, "2022-05-10"},
		{"EMP008", "Frank Garcia", "frank.garcia@company.com", "Sales", 33, 69000.00, "2022-01-30"},
	}

	// Add employee data with appropriate styling
	for _, emp := range employees {
		row := employeeSheet.AddRow()
		row.AddCell(emp[0]).SetStyle(dataStyle)                    // Employee ID
		row.AddCell(emp[1]).SetStyle(dataStyle)                    // Full Name
		row.AddCell(emp[2]).SetStyle(dataStyle)                    // Email
		row.AddCell(emp[3]).SetStyle(dataStyle)                    // Department
		row.AddCell(emp[4]).SetStyle(numberStyle)                  // Age
		row.AddCell(emp[5]).SetStyle(numberStyle)                  // Salary
		row.AddCell(emp[6]).SetStyle(dataStyle)                    // Start Date
	}

	// Sheet 2: Department Summary
	deptSheet := workbook.AddSheet("Department Summary")
	if deptSheet == nil {
		log.Fatal("Failed to create department sheet")
	}

	// Set column widths
	deptSheet.SetColumnWidth("A", 20.0) // Department
	deptSheet.SetColumnWidth("B", 15.0) // Employee Count
	deptSheet.SetColumnWidth("C", 18.0) // Average Salary
	deptSheet.SetColumnWidth("D", 18.0) // Total Salary

	// Add header row
	deptHeaderRow := deptSheet.AddRow()
	deptHeaderRow.AddCell("Department").SetStyle(headerStyle)
	deptHeaderRow.AddCell("Employee Count").SetStyle(headerStyle)
	deptHeaderRow.AddCell("Average Salary").SetStyle(headerStyle)
	deptHeaderRow.AddCell("Total Salary").SetStyle(headerStyle)

	// Department summary data (calculated from employee data)
	departments := [][]interface{}{
		{"Engineering", 2, 76500.00, 153000.00},
		{"Marketing", 2, 64000.00, 128000.00},
		{"Sales", 2, 69500.00, 139000.00},
		{"HR", 1, 68000.00, 68000.00},
		{"Finance", 1, 72000.00, 72000.00},
	}

	// Add department data
	for _, dept := range departments {
		row := deptSheet.AddRow()
		row.AddCell(dept[0]).SetStyle(dataStyle)     // Department
		row.AddCell(dept[1]).SetStyle(numberStyle)   // Employee Count
		row.AddCell(dept[2]).SetStyle(numberStyle)   // Average Salary
		row.AddCell(dept[3]).SetStyle(numberStyle)   // Total Salary
	}

	// Sheet 3: Company Overview
	overviewSheet := workbook.AddSheet("Company Overview")
	if overviewSheet == nil {
		log.Fatal("Failed to create overview sheet")
	}

	// Set column widths
	overviewSheet.SetColumnWidth("A", 25.0)
	overviewSheet.SetColumnWidth("B", 20.0)

	// Company overview data
	overviewData := [][]interface{}{
		{"Company Name", "Tech Solutions Inc."},
		{"Total Employees", 8},
		{"Total Departments", 5},
		{"Average Employee Age", 30.6},
		{"Total Payroll", 560000.00},
		{"Report Generated", time.Now().Format("2006-01-02 15:04:05")},
	}

	// Add overview data with alternating row colors
	for i, data := range overviewData {
		row := overviewSheet.AddRow()
		
		// Create alternating row style
		labelStyle := dataStyle
		valueStyle := dataStyle
		if i%2 == 0 {
			labelStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#F2F2F2"}
			valueStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#F2F2F2"}
		}
		
		labelStyle.Font.Bold = true
		row.AddCell(data[0]).SetStyle(labelStyle)
		row.AddCell(data[1]).SetStyle(valueStyle)
	}

	// Merge cells for title
	titleRow := overviewSheet.AddRow()
	titleRow.AddCell("COMPANY OVERVIEW REPORT").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   16,
			Color:  "#2F5597",
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	})

	// Build the workbook
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build workbook")
	}

	// Save the file
	filename := "output/01-basic-enhanced-example.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save workbook: %v", err)
	}

	fmt.Printf("✅ Enhanced basic example created successfully!\n")
	fmt.Printf("📁 File saved as: %s\n", filename)
	fmt.Printf("📊 Features demonstrated:\n")
	fmt.Printf("   • Workbook properties and metadata\n")
	fmt.Printf("   • Multiple sheets with different purposes\n")
	fmt.Printf("   • Comprehensive styling (fonts, colors, borders)\n")
	fmt.Printf("   • Column width management\n")
	fmt.Printf("   • Different data types (text, numbers, dates)\n")
	fmt.Printf("   • Number formatting\n")
	fmt.Printf("   • Reusable style configurations\n")
	fmt.Printf("   • Professional report layout\n")
	fmt.Printf("\n🎯 Next steps: Try examples/02-data-types/ to explore data handling\n")
}