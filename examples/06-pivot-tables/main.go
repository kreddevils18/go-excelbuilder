package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

type SalesRecord struct {
	Date    string
	Product string
	Region  string
	Sales   float64
}

func getMockRawData() []SalesRecord {
	return []SalesRecord{
		{"2024-01-15", "Laptop", "North", 15000},
		{"2024-01-20", "Monitor", "South", 8000},
		{"2024-02-10", "Laptop", "North", 18000},
		{"2024-02-18", "Keyboard", "East", 2500},
		{"2024-03-05", "Monitor", "West", 9500},
		{"2024-03-12", "Laptop", "South", 22000},
		{"2024-01-28", "Keyboard", "North", 3000},
		{"2024-02-22", "Monitor", "East", 8500},
	}
}

func main() {
	builder := excelbuilder.New()
	wb := builder.NewWorkbook()
	data := getMockRawData()

	// 1. Create the Raw Data sheet
	rawDataSheet := wb.AddSheet("Raw Data")
	headerStyle := excelbuilder.StyleConfig{Font: excelbuilder.FontConfig{Bold: true}}
	currencyStyle := excelbuilder.StyleConfig{NumberFormat: "#,##0.00"}

	rawDataSheet.AddRow().
		AddCell("Date").WithStyle(headerStyle).Done().
		AddCell("Product").WithStyle(headerStyle).Done().
		AddCell("Region").WithStyle(headerStyle).Done().
		AddCell("Sales").WithStyle(headerStyle)

	for _, record := range data {
		rawDataSheet.AddRow().
			AddCell(record.Date).Done().
			AddCell(record.Product).Done().
			AddCell(record.Region).Done().
			AddCell(record.Sales).WithStyle(currencyStyle)
	}

	// 2. Create the Pivot Table sheet
	pivotSheet := wb.AddSheet("Pivot Report")
	pivotBuilder := pivotSheet.NewPivotTable("Pivot Report", "Raw Data!A1:D9")

	// 3. Configure and build the pivot table
	err := pivotBuilder.
		SetTargetCell("A1").
		WithStyle("PivotStyleLight16").
		AddRowField("Region").
		AddColumnField("Product").
		AddValueField("Sales", "sum").
		Build()

	if err != nil {
		log.Fatalf("Failed to build pivot table: %v", err)
	}

	// Save the file
	file := wb.Build()
	outputDir := "examples/06-pivot-tables/output"
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}
	filePath := filepath.Join(outputDir, "pivot_table.xlsx")

	if err := file.SaveAs(filePath); err != nil {
		log.Fatalf("Failed to save file: %v", err)
	}

	fmt.Printf("Successfully created %s\n", filePath)
}
