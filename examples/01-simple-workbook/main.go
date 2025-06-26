package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

func main() {
	// Create a new builder instance
	builder := excelbuilder.New()

	// Use the fluent API to build a workbook
	file := builder.
		NewWorkbook().
		AddSheet("Sheet1").
		AddRow().
		AddCell("Hello").Done().
		AddCell("World!").Done().
		Done().
		Build()

	// Define the output path
	outputDir := "examples/01-simple-workbook/output"
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}
	filePath := filepath.Join(outputDir, "simple_workbook.xlsx")

	// Save the file
	if err := file.SaveAs(filePath); err != nil {
		log.Fatalf("Failed to save file: %v", err)
	}

	fmt.Printf("Successfully created %s\n", filePath)
}
