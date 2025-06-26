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

	// Define various styles
	headerStyle := excelbuilder.StyleConfig{
		Font:      excelbuilder.FontConfig{Bold: true, Color: "FFFFFF", Size: 12},
		Fill:      excelbuilder.FillConfig{Type: "pattern", Color: "4472C4"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	}

	currencyStyle := excelbuilder.StyleConfig{
		NumberFormat: `"$#,##0.00"`,
		Font:         excelbuilder.FontConfig{Family: "Arial"},
	}

	wrappedTextStyle := excelbuilder.StyleConfig{
		Alignment: excelbuilder.AlignmentConfig{WrapText: true, Vertical: "top"},
	}

	borderedStyle := excelbuilder.StyleConfig{
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "0000FF"},
			Bottom: excelbuilder.BorderSide{Style: "double", Color: "FF0000"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "00FF00"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "00FF00"},
		},
	}

	// Build the sheet
	sheet := builder.NewWorkbook().AddSheet("Styling Demo")
	sheet.SetColumnWidth("A", 20).SetColumnWidth("B", 40).SetColumnWidth("C", 20)

	// Apply styles
	sheet.AddRow().
		AddCell("Header 1").WithStyle(headerStyle).Done().
		AddCell("Header 2").WithStyle(headerStyle).Done().
		AddCell("Header 3").WithStyle(headerStyle)

	sheet.AddRow().
		AddCell(1234.56).WithStyle(currencyStyle).Done().
		AddCell("This is a long text that will be wrapped inside the cell.").WithStyle(wrappedTextStyle).Done().
		AddCell("Bordered").WithStyle(borderedStyle)

	// Save the file
	file := sheet.Build()
	outputDir := "examples/02-cell-styling/output"
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}
	filePath := filepath.Join(outputDir, "cell_styling.xlsx")

	if err := file.SaveAs(filePath); err != nil {
		log.Fatalf("Failed to save file: %v", err)
	}

	fmt.Printf("Successfully created %s\n", filePath)
}
