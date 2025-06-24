package main

import (
	"fmt"
	"log"
	"math"
	"os"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

func main() {
	// Create output directory
	if err := os.MkdirAll("output", 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}

	// Create Excel builder
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Data Types Demonstration",
		Author:      "Go Excel Builder",
		Subject:     "Comprehensive Data Type Handling",
		Description: "Demonstrates all supported data types and their proper formatting",
		Keywords:    "go,excel,data-types,formatting,examples",
	})

	// Define styles for different data types
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
			Color:  "#FFFFFF",
			Family: "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#2F5597",
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

	textStyle := excelbuilder.StyleConfig{
		Font:      excelbuilder.FontConfig{Size: 10, Family: "Arial"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "left", Vertical: "middle"},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
		},
	}

	numberStyle := excelbuilder.StyleConfig{
		Font:         excelbuilder.FontConfig{Size: 10, Family: "Arial"},
		Alignment:    excelbuilder.AlignmentConfig{Horizontal: "right", Vertical: "middle"},
		NumberFormat: "#,##0.00",
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
		},
	}

	currencyStyle := excelbuilder.StyleConfig{
		Font:         excelbuilder.FontConfig{Size: 10, Family: "Arial"},
		Alignment:    excelbuilder.AlignmentConfig{Horizontal: "right", Vertical: "middle"},
		NumberFormat: "$#,##0.00",
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
		},
	}

	percentStyle := excelbuilder.StyleConfig{
		Font:         excelbuilder.FontConfig{Size: 10, Family: "Arial"},
		Alignment:    excelbuilder.AlignmentConfig{Horizontal: "right", Vertical: "middle"},
		NumberFormat: "0.00%",
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
		},
	}

	dateStyle := excelbuilder.StyleConfig{
		Font:         excelbuilder.FontConfig{Size: 10, Family: "Arial"},
		Alignment:    excelbuilder.AlignmentConfig{Horizontal: "center", Vertical: "middle"},
		NumberFormat: "yyyy-mm-dd",
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
		},
	}

	datetimeStyle := excelbuilder.StyleConfig{
		Font:         excelbuilder.FontConfig{Size: 10, Family: "Arial"},
		Alignment:    excelbuilder.AlignmentConfig{Horizontal: "center", Vertical: "middle"},
		NumberFormat: "yyyy-mm-dd hh:mm:ss",
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
		},
	}

	booleanStyle := excelbuilder.StyleConfig{
		Font:      excelbuilder.FontConfig{Size: 10, Family: "Arial", Bold: true},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center", Vertical: "middle"},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
		},
	}

	// Sheet 1: Basic Data Types
	basicSheet := workbook.AddSheet("Basic Data Types")
	if basicSheet == nil {
		log.Fatal("Failed to create basic data types sheet")
	}

	// Set column widths
	basicSheet.SetColumnWidth("A", 20.0) // Data Type
	basicSheet.SetColumnWidth("B", 25.0) // Description
	basicSheet.SetColumnWidth("C", 20.0) // Example Value
	basicSheet.SetColumnWidth("D", 15.0) // Go Type
	basicSheet.SetColumnWidth("E", 20.0) // Excel Format

	// Add header
	headerRow := basicSheet.AddRow()
	headerRow.AddCell("Data Type").SetStyle(headerStyle)
	headerRow.AddCell("Description").SetStyle(headerStyle)
	headerRow.AddCell("Example Value").SetStyle(headerStyle)
	headerRow.AddCell("Go Type").SetStyle(headerStyle)
	headerRow.AddCell("Excel Format").SetStyle(headerStyle)

	// String data
	stringRow := basicSheet.AddRow()
	stringRow.AddCell("String").SetStyle(textStyle)
	stringRow.AddCell("Text data, names, descriptions").SetStyle(textStyle)
	stringRow.AddCell("Hello, World!").SetStyle(textStyle)
	stringRow.AddCell("string").SetStyle(textStyle)
	stringRow.AddCell("General").SetStyle(textStyle)

	// Integer data
	intRow := basicSheet.AddRow()
	intRow.AddCell("Integer").SetStyle(textStyle)
	intRow.AddCell("Whole numbers, counts, IDs").SetStyle(textStyle)
	intRow.AddCell(42).SetStyle(numberStyle)
	intRow.AddCell("int").SetStyle(textStyle)
	intRow.AddCell("#,##0").SetStyle(textStyle)

	// Float data
	floatRow := basicSheet.AddRow()
	floatRow.AddCell("Float").SetStyle(textStyle)
	floatRow.AddCell("Decimal numbers, measurements").SetStyle(textStyle)
	floatRow.AddCell(3.14159).SetStyle(numberStyle)
	floatRow.AddCell("float64").SetStyle(textStyle)
	floatRow.AddCell("#,##0.00").SetStyle(textStyle)

	// Currency data
	currencyRow := basicSheet.AddRow()
	currencyRow.AddCell("Currency").SetStyle(textStyle)
	currencyRow.AddCell("Money values, prices, salaries").SetStyle(textStyle)
	currencyRow.AddCell(1234.56).SetStyle(currencyStyle)
	currencyRow.AddCell("float64").SetStyle(textStyle)
	currencyRow.AddCell("$#,##0.00").SetStyle(textStyle)

	// Percentage data
	percentRow := basicSheet.AddRow()
	percentRow.AddCell("Percentage").SetStyle(textStyle)
	percentRow.AddCell("Rates, ratios, completion").SetStyle(textStyle)
	percentRow.AddCell(0.75).SetStyle(percentStyle)
	percentRow.AddCell("float64").SetStyle(textStyle)
	percentRow.AddCell("0.00%").SetStyle(textStyle)

	// Date data
	dateRow := basicSheet.AddRow()
	dateRow.AddCell("Date").SetStyle(textStyle)
	dateRow.AddCell("Dates without time").SetStyle(textStyle)
	dateRow.AddCell(time.Now().Format("2006-01-02")).SetStyle(dateStyle)
	dateRow.AddCell("time.Time").SetStyle(textStyle)
	dateRow.AddCell("yyyy-mm-dd").SetStyle(textStyle)

	// DateTime data
	datetimeRow := basicSheet.AddRow()
	datetimeRow.AddCell("DateTime").SetStyle(textStyle)
	datetimeRow.AddCell("Dates with time").SetStyle(textStyle)
	datetimeRow.AddCell(time.Now().Format("2006-01-02 15:04:05")).SetStyle(datetimeStyle)
	datetimeRow.AddCell("time.Time").SetStyle(textStyle)
	datetimeRow.AddCell("yyyy-mm-dd hh:mm:ss").SetStyle(textStyle)

	// Boolean data
	booleanRow := basicSheet.AddRow()
	booleanRow.AddCell("Boolean").SetStyle(textStyle)
	booleanRow.AddCell("True/false values, flags").SetStyle(textStyle)
	booleanRow.AddCell(true).SetStyle(booleanStyle)
	booleanRow.AddCell("bool").SetStyle(textStyle)
	booleanRow.AddCell("TRUE/FALSE").SetStyle(textStyle)

	// Sheet 2: Numeric Data Examples
	numericSheet := workbook.AddSheet("Numeric Examples")
	if numericSheet == nil {
		log.Fatal("Failed to create numeric sheet")
	}

	// Set column widths
	numericSheet.SetColumnWidth("A", 25.0) // Description
	numericSheet.SetColumnWidth("B", 15.0) // Integer
	numericSheet.SetColumnWidth("C", 15.0) // Float
	numericSheet.SetColumnWidth("D", 15.0) // Currency
	numericSheet.SetColumnWidth("E", 15.0) // Percentage
	numericSheet.SetColumnWidth("F", 20.0) // Scientific

	// Add header
	numericHeaderRow := numericSheet.AddRow()
	numericHeaderRow.AddCell("Description").SetStyle(headerStyle)
	numericHeaderRow.AddCell("Integer").SetStyle(headerStyle)
	numericHeaderRow.AddCell("Float").SetStyle(headerStyle)
	numericHeaderRow.AddCell("Currency").SetStyle(headerStyle)
	numericHeaderRow.AddCell("Percentage").SetStyle(headerStyle)
	numericHeaderRow.AddCell("Scientific").SetStyle(headerStyle)

	// Scientific notation style
	scientificStyle := excelbuilder.StyleConfig{
		Font:         excelbuilder.FontConfig{Size: 10, Family: "Arial"},
		Alignment:    excelbuilder.AlignmentConfig{Horizontal: "right", Vertical: "middle"},
		NumberFormat: "0.00E+00",
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
		},
	}

	// Numeric examples
	numericExamples := []struct {
		description string
		integer     int
		float       float64
		currency    float64
		percentage  float64
		scientific  float64
	}{
		{"Small values", 1, 1.23, 1.23, 0.01, 0.000123},
		{"Medium values", 1000, 1234.56, 1234.56, 0.25, 1234.56},
		{"Large values", 1000000, 1234567.89, 1234567.89, 0.75, 1234567.89},
		{"Very large values", 1000000000, 1.23e9, 1.23e9, 1.0, 1.23e9},
		{"Decimal precision", 0, math.Pi, math.Pi, math.Pi / 10, math.Pi},
		{"Negative values", -500, -123.45, -123.45, -0.15, -123.45},
		{"Zero values", 0, 0.0, 0.0, 0.0, 0.0},
		{"Mathematical constants", 2, math.E, math.E, math.E / 10, math.E},
	}

	for _, example := range numericExamples {
		row := numericSheet.AddRow()
		row.AddCell(example.description).SetStyle(textStyle)
		row.AddCell(example.integer).SetStyle(numberStyle)
		row.AddCell(example.float).SetStyle(numberStyle)
		row.AddCell(example.currency).SetStyle(currencyStyle)
		row.AddCell(example.percentage).SetStyle(percentStyle)
		row.AddCell(example.scientific).SetStyle(scientificStyle)
	}

	// Sheet 3: Date and Time Examples
	dateTimeSheet := workbook.AddSheet("Date & Time Examples")
	if dateTimeSheet == nil {
		log.Fatal("Failed to create date time sheet")
	}

	// Set column widths
	dateTimeSheet.SetColumnWidth("A", 25.0) // Description
	dateTimeSheet.SetColumnWidth("B", 20.0) // Date Only
	dateTimeSheet.SetColumnWidth("C", 25.0) // Date Time
	dateTimeSheet.SetColumnWidth("D", 15.0) // Time Only
	dateTimeSheet.SetColumnWidth("E", 20.0) // Custom Format

	// Add header
	dateTimeHeaderRow := dateTimeSheet.AddRow()
	dateTimeHeaderRow.AddCell("Description").SetStyle(headerStyle)
	dateTimeHeaderRow.AddCell("Date Only").SetStyle(headerStyle)
	dateTimeHeaderRow.AddCell("Date & Time").SetStyle(headerStyle)
	dateTimeHeaderRow.AddCell("Time Only").SetStyle(headerStyle)
	dateTimeHeaderRow.AddCell("Custom Format").SetStyle(headerStyle)

	// Time only style
	timeStyle := excelbuilder.StyleConfig{
		Font:         excelbuilder.FontConfig{Size: 10, Family: "Arial"},
		Alignment:    excelbuilder.AlignmentConfig{Horizontal: "center", Vertical: "middle"},
		NumberFormat: "hh:mm:ss",
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
		},
	}

	// Custom date format style
	customDateStyle := excelbuilder.StyleConfig{
		Font:         excelbuilder.FontConfig{Size: 10, Family: "Arial"},
		Alignment:    excelbuilder.AlignmentConfig{Horizontal: "center", Vertical: "middle"},
		NumberFormat: "dddd, mmmm dd, yyyy",
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#CCCCCC"},
		},
	}

	// Date and time examples
	now := time.Now()
	dateTimeExamples := []struct {
		description string
		date        time.Time
	}{
		{"Current date/time", now},
		{"Start of year", time.Date(now.Year(), 1, 1, 0, 0, 0, 0, time.UTC)},
		{"End of year", time.Date(now.Year(), 12, 31, 23, 59, 59, 0, time.UTC)},
		{"Unix epoch", time.Unix(0, 0)},
		{"Future date", now.AddDate(1, 0, 0)},
		{"Past date", now.AddDate(-1, 0, 0)},
		{"Leap year date", time.Date(2024, 2, 29, 12, 0, 0, 0, time.UTC)},
		{"Midnight", time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, time.UTC)},
		{"Noon", time.Date(now.Year(), now.Month(), now.Day(), 12, 0, 0, 0, time.UTC)},
	}

	for _, example := range dateTimeExamples {
		row := dateTimeSheet.AddRow()
		row.AddCell(example.description).SetStyle(textStyle)
		row.AddCell(example.date.Format("2006-01-02")).SetStyle(dateStyle)
		row.AddCell(example.date.Format("2006-01-02 15:04:05")).SetStyle(datetimeStyle)
		row.AddCell(example.date.Format("15:04:05")).SetStyle(timeStyle)
		row.AddCell(example.date.Format("Monday, January 02, 2006")).SetStyle(customDateStyle)
	}

	// Sheet 4: Complex Data Structures
	complexSheet := workbook.AddSheet("Complex Data")
	if complexSheet == nil {
		log.Fatal("Failed to create complex data sheet")
	}

	// Set column widths
	complexSheet.SetColumnWidth("A", 15.0) // ID
	complexSheet.SetColumnWidth("B", 25.0) // Name
	complexSheet.SetColumnWidth("C", 30.0) // Email
	complexSheet.SetColumnWidth("D", 15.0) // Age
	complexSheet.SetColumnWidth("E", 15.0) // Salary
	complexSheet.SetColumnWidth("F", 15.0) // Bonus %
	complexSheet.SetColumnWidth("G", 20.0) // Start Date
	complexSheet.SetColumnWidth("H", 15.0) // Active
	complexSheet.SetColumnWidth("I", 20.0) // Last Login

	// Add header
	complexHeaderRow := complexSheet.AddRow()
	complexHeaderRow.AddCell("ID").SetStyle(headerStyle)
	complexHeaderRow.AddCell("Name").SetStyle(headerStyle)
	complexHeaderRow.AddCell("Email").SetStyle(headerStyle)
	complexHeaderRow.AddCell("Age").SetStyle(headerStyle)
	complexHeaderRow.AddCell("Salary").SetStyle(headerStyle)
	complexHeaderRow.AddCell("Bonus %").SetStyle(headerStyle)
	complexHeaderRow.AddCell("Start Date").SetStyle(headerStyle)
	complexHeaderRow.AddCell("Active").SetStyle(headerStyle)
	complexHeaderRow.AddCell("Last Login").SetStyle(headerStyle)

	// Complex data examples with mixed types
	complexData := [][]interface{}{
		{1, "John Doe", "john.doe@company.com", 30, 75000.00, 0.15, "2022-01-15", true, time.Now().Add(-2 * time.Hour).Format("2006-01-02 15:04:05")},
		{2, "Jane Smith", "jane.smith@company.com", 28, 68000.00, 0.12, "2022-03-01", true, time.Now().Add(-1 * time.Hour).Format("2006-01-02 15:04:05")},
		{3, "Bob Johnson", "bob.johnson@company.com", 35, 82000.00, 0.18, "2021-11-10", false, time.Now().Add(-24 * time.Hour).Format("2006-01-02 15:04:05")},
		{4, "Alice Brown", "alice.brown@company.com", 32, 71000.00, 0.14, "2022-02-20", true, time.Now().Add(-30 * time.Minute).Format("2006-01-02 15:04:05")},
		{5, "Charlie Wilson", "charlie.wilson@company.com", 29, 79000.00, 0.16, "2022-04-05", true, time.Now().Add(-10 * time.Minute).Format("2006-01-02 15:04:05")},
	}

	for _, data := range complexData {
		row := complexSheet.AddRow()
		row.AddCell(data[0]).SetStyle(numberStyle)   // ID
		row.AddCell(data[1]).SetStyle(textStyle)     // Name
		row.AddCell(data[2]).SetStyle(textStyle)     // Email
		row.AddCell(data[3]).SetStyle(numberStyle)   // Age
		row.AddCell(data[4]).SetStyle(currencyStyle) // Salary
		row.AddCell(data[5]).SetStyle(percentStyle)  // Bonus %
		row.AddCell(data[6]).SetStyle(dateStyle)     // Start Date
		row.AddCell(data[7]).SetStyle(booleanStyle)  // Active
		row.AddCell(data[8]).SetStyle(datetimeStyle) // Last Login
	}

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build workbook")
	}

	filename := "output/02-data-types-comprehensive.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save workbook: %v", err)
	}

	fmt.Printf("✅ Data types example created successfully!\n")
	fmt.Printf("📁 File saved as: %s\n", filename)
	fmt.Printf("📊 Features demonstrated:\n")
	fmt.Printf("   • All basic data types (string, int, float, bool)\n")
	fmt.Printf("   • Formatted numbers (currency, percentage, scientific)\n")
	fmt.Printf("   • Date and time handling with various formats\n")
	fmt.Printf("   • Complex data structures with mixed types\n")
	fmt.Printf("   • Proper type-specific styling and formatting\n")
	fmt.Printf("   • Number format strings for Excel\n")
	fmt.Printf("   • Professional data presentation\n")
	fmt.Printf("\n🎯 Next steps: Try examples/03-styling-advanced/ for advanced styling\n")
}
