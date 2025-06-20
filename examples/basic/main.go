package main

import (
	"fmt"
	"log"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

func main() {
	// Create a new Excel builder
	builder := excelbuilder.New()

	// Create a new workbook with properties
	workbook := builder.NewWorkbook()
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Sample Excel File",
		Author:      "Go Excel Builder",
		Subject:     "Demo",
		Description: "A simple example of using Go Excel Builder",
		Category:    "Example",
		Keywords:    "go,excel,builder",
	})

	// Add a sheet
	sheet := workbook.AddSheet("Sample Data")
	if sheet == nil {
		log.Fatal("Failed to create sheet")
	}

	// Set column widths
	sheet.SetColumnWidth("A", 15.0)
	sheet.SetColumnWidth("B", 20.0)
	sheet.SetColumnWidth("C", 12.0)

	// Add header row
	headerRow := sheet.AddRow()
	headerRow.AddCell("Name")
	headerRow.AddCell("Email")
	headerRow.AddCell("Age")

	// Add data rows
	dataRows := [][]interface{}{
		{"John Doe", "john@example.com", 30},
		{"Jane Smith", "jane@example.com", 25},
		{"Bob Johnson", "bob@example.com", 35},
	}

	for _, rowData := range dataRows {
		row := sheet.AddRow()
		for _, cellData := range rowData {
			row.AddCell(cellData)
		}
	}

	// Build the workbook
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build workbook")
	}

	// Save the file
	err := file.SaveAs("sample.xlsx")
	if err != nil {
		log.Fatalf("Failed to save workbook: %v", err)
	}

	fmt.Println("Excel file 'sample.xlsx' created successfully!")
}