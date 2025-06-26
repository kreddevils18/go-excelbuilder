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
	sheet := builder.NewWorkbook().AddSheet("Validation Demo")
	sheet.SetColumnWidth("A", 20).SetColumnWidth("B", 30)

	// 1. Add a header
	sheet.AddRow().
		AddCell("Field").WithStyle(excelbuilder.StyleConfig{Font: excelbuilder.FontConfig{Bold: true}}).Done().
		AddCell("Input").WithStyle(excelbuilder.StyleConfig{Font: excelbuilder.FontConfig{Bold: true}})
	sheet.AddRow() // Spacer

	// 2. Define the data validation rule for a dropdown list
	validationConfig := &excelbuilder.DataValidationConfig{
		Type:     "list",
		Formula1: []string{"Admin", "User", "Guest"}, // The items for the dropdown
	}

	// 3. Add a cell and apply the validation rule
	sheet.AddRow().
		AddCell("User Role:").Done().
		AddCell("").WithDataValidation(validationConfig)

	// Save the file
	file := sheet.Build()
	outputDir := "examples/05-data-validation/output"
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}
	filePath := filepath.Join(outputDir, "data_validation.xlsx")

	if err := file.SaveAs(filePath); err != nil {
		log.Fatalf("Failed to save file: %v", err)
	}

	fmt.Printf("Successfully created %s\n", filePath)
}
