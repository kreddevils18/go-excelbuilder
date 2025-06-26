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
	sheet := builder.NewWorkbook().AddSheet("Chart Demo")

	// 1. Add data for the chart
	sheet.AddRow().AddCells("Month", "Revenue", "Profit")
	sheet.AddRow().AddCells("Jan", 15000, 4000)
	sheet.AddRow().AddCells("Feb", 18000, 5000)
	sheet.AddRow().AddCells("Mar", 22000, 6500)
	sheet.AddRow().AddCells("Apr", 21000, 6000)

	// 2. Create the chart object
	chart := sheet.AddChart()

	// 3. Configure the chart
	chart.SetType("col").
		SetTitle("Quarterly Performance").
		SetXAxis(excelbuilder.AxisConfig{Title: "Month"}).
		SetYAxis(excelbuilder.AxisConfig{Title: "Amount (USD)"}).
		SetPosition("E2").
		SetDimensions(480, 320)

	// 4. Add data series to the chart
	chart.AddDataSeries(excelbuilder.DataSeries{
		Name:       "'Chart Demo'!$B$1", // Series name from cell (Revenue)
		Categories: "'Chart Demo'!$A$2:$A$5",
		Values:     "'Chart Demo'!$B$2:$B$5",
	})
	chart.AddDataSeries(excelbuilder.DataSeries{
		Name:       "'Chart Demo'!$C$1", // Series name from cell (Profit)
		Categories: "'Chart Demo'!$A$2:$A$5",
		Values:     "'Chart Demo'!$C$2:$C$5",
	})

	// 5. Build the chart into the sheet
	if err := chart.Build(); err != nil {
		log.Fatalf("Failed to build chart: %v", err)
	}

	// Save the file
	file := sheet.Build()
	outputDir := "examples/04-charts/output"
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}
	filePath := filepath.Join(outputDir, "charts.xlsx")

	if err := file.SaveAs(filePath); err != nil {
		log.Fatalf("Failed to save file: %v", err)
	}

	fmt.Printf("Successfully created %s\n", filePath)
}
