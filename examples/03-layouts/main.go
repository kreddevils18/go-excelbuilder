package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

func main() {
	builder := excelbuilder.New()
	sheet := builder.NewWorkbook().AddSheet("Layout Demo")

	// 1. Set Column Widths
	sheet.SetColumnWidth("A", 10).
		SetColumnWidth("B", 30).
		SetColumnWidth("C", 15).
		SetColumnWidth("D", 15)

	// 2. Merge Cells to create a title
	titleStyle := excelbuilder.StyleConfig{
		Font:      excelbuilder.FontConfig{Bold: true, Size: 16},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center", Vertical: "center"},
	}
	sheet.AddRow().
		AddCell("Inventory Report").WithStyle(titleStyle).WithMergeRange("D1")
	sheet.SetRowHeight(1, 40) // Make the merged title row taller
	sheet.AddRow()            // Spacer row

	// 3. Create a header and freeze it
	headerStyle := excelbuilder.StyleConfig{
		Font:      excelbuilder.FontConfig{Bold: true},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	}
	sheet.AddRow().
		AddCell("ID").WithStyle(headerStyle).Done().
		AddCell("Item Name").WithStyle(headerStyle).Done().
		AddCell("In Stock").WithStyle(headerStyle).Done().
		AddCell("Unit Price").WithStyle(headerStyle)

	// Freeze the top 3 rows (title, spacer, header)
	sheet.FreezePanes(0, 3)

	// Add some data rows to demonstrate scrolling with frozen panes
	for i := 1; i <= 50; i++ {
		sheet.AddRow().
			AddCell(fmt.Sprintf("SKU-%04d", i)).Done().
			AddCell(fmt.Sprintf("Product Name %d", i)).Done().
			AddCell(100 + i*2).Done().
			AddCell(19.99 + float64(i))
	}

	// Save the file
	file := sheet.Build()
	outputDir := "examples/03-layouts/output"
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}
	filePath := filepath.Join(outputDir, "layouts.xlsx")

	if err := file.SaveAs(filePath); err != nil {
		log.Fatalf("Failed to save file: %v", err)
	}

	fmt.Printf("Successfully created %s\n", filePath)
}
