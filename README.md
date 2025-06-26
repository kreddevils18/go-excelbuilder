# Go Excel Builder

A fluent, builder-pattern library for creating Excel files in Go, built on top of the excellent [excelize](https://github.com/xuri/excelize) library. This library simplifies Excel creation with a clean, chainable API and optimizes performance for large files using a Flyweight pattern for style management.

## Features

- **Fluent Builder API**: Chain method calls for an intuitive and readable way to build workbooks, sheets, rows, and cells.
- **Efficient Style Management**: Uses the Flyweight design pattern to cache and reuse styles, significantly reducing memory consumption and improving performance when generating large, heavily-styled files.
- **Rich Feature Set**: Comprehensive support for advanced Excel features including:
  - **Charts**: Create column, bar, line, pie, and scatter charts.
  - **Pivot Tables**: Generate complex pivot tables from your data.
  - **Data Validation**: Enforce data integrity with validation rules.
  - **Advanced Layouts**: Control column/row sizing, merge cells, and freeze panes.
  - **Professional Styling**: Full control over fonts, fills, borders, alignments, and number formats.
  - **Import/Export**: Convert data to and from CSV and JSON.
- **Type-Safe**: Designed to maximize type safety and catch errors at compile time.
- **Extensively Tested**: Includes a full suite of unit, integration, and performance tests to ensure reliability.

## Installation

```bash
go get github.com/kreddevils18/go-excelbuilder
```

## Quick Start

Create a simple styled Excel file in just a few lines of code.

```go
package main

import (
	"log"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

func main() {
	builder := excelbuilder.New()

	// Define a reusable style
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Color: "FFFFFF"},
		Fill: excelbuilder.FillConfig{Type: "pattern", Color: "4472C4"},
	}

	// Build the workbook
	file := builder.
		NewWorkbook().
		AddSheet("Sales Report").
		SetColumnWidth("A", 20).
		AddRow().
			AddCell("Product").WithStyle(headerStyle).Done().
			AddCell("Revenue").WithStyle(headerStyle).Done().
		Done().
		AddRow().
			AddCell("Laptop").Done().
			AddCell(3000).Done().
		Done().
		Build()

	// Save the file
	if err := file.SaveAs("SalesReport.xlsx"); err != nil {
		log.Fatal(err)
	}
}
```

## How to Use Advanced Features

### Styling with the Flyweight Pattern

The library's core performance feature is its style management. You define a `StyleConfig` struct and apply it to cells. The `StyleManager` automatically handles caching and reuse in the background.

**Example:** Create and apply multiple styles.

```go
// Define a style for headers
headerStyle := excelbuilder.StyleConfig{
    Font:      excelbuilder.FontConfig{Bold: true, Color: "FFFFFF"},
    Fill:      excelbuilder.FillConfig{Type: "pattern", Color: "4472C4"},
    Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
}

// Define a style for currency data
currencyStyle := excelbuilder.StyleConfig{
    NumberFormat: `"$#,##0.00"`,
}

// Apply styles to cells
sheet.AddRow().
    AddCell("Item").WithStyle(headerStyle).Done().
    AddCell("Price").WithStyle(headerStyle).
Done()

sheet.AddRow().
    AddCell("Keyboard").Done().
    AddCell(75.50).WithStyle(currencyStyle).
Done()
```

### Creating Charts

You can add charts to any sheet to visualize your data.

```go
// ... after adding data to cells A1:B4 ...

// Get the current sheet
sheet := wb.AddSheet("Chart Demo")

// Add data for the chart
sheet.AddRow().AddCells("Category", "Value")
sheet.AddRow().AddCells("A", 100)
sheet.AddRow().AddCells("B", 150)
sheet.AddRow().AddCells("C", 80)

// Create the chart
chart := sheet.AddChart()
chart.SetType("col") // col, bar, pie, line, etc.
chart.SetTitle("Sample Column Chart")
chart.AddDataSeries(excelbuilder.DataSeries{
    Name:       "'Chart Demo'!$B$1", // Series name from cell
    Categories: "'Chart Demo'!$A$2:$A$4",
    Values:     "'Chart Demo'!$B$2:$B$4",
})
chart.SetPosition("D1") // Top-left corner of the chart
chart.Build()
```

### Advanced Layout

Easily control the layout of your worksheet.

```go
sheet := wb.AddSheet("Layout Demo")

// Set column widths
sheet.SetColumnWidth("A", 30) // Set width for a single column
sheet.SetColumnWidth("B", 15)

// Merge cells
sheet.AddRow().AddCell("Merged Header").WithMergeRange("C1") // Merges A1:C1
sheet.AddRow().AddCells("A", "B", "C")
```

### Data Validation

Enforce data integrity with validation rules, such as requiring a number within a range or selecting from a predefined list.

```go
// ... in a sheet builder context ...

// Add a header for the column that will have validation
sheet.AddRow().AddCell("Select a Department:").Done()

// Define the data validation rule for a dropdown list
validation := excelbuilder.DataValidationConfig{
    Type:     "list",
    // The formula for a list must be a string containing comma-separated values, enclosed in quotes.
    Formula1: []string{`"HR,Engineering,Sales,Marketing"`},
}

// Add a cell and apply the validation rule to it.
// The user will now see a dropdown in cell A2.
sheet.AddRow().AddCell("").WithDataValidation(&validation).Done()
```

## Examples

This project includes a suite of examples in the `examples/` directory to demonstrate various features. It is recommended to review them in order to understand how to use the library effectively.

| #   | Example            | Description                                                             |
| --- | ------------------ | ----------------------------------------------------------------------- |
| 01  | `simple-workbook`  | The most basic example, showing how to create a file with a few cells.  |
| 02  | `cell-styling`     | Demonstrates fonts, fills, borders, alignment, and number formats.      |
| 03  | `layouts`          | Shows how to merge cells, set column widths, and freeze panes.          |
| 04  | `charts`           | A focused example on creating a column chart.                           |
| 05  | `data-validation`  | Shows how to create a dropdown list in a cell.                          |
| 06  | `pivot-tables`     | A dedicated example for building a pivot table from raw data.           |
| 07  | `financial-report` | An advanced, multi-sheet report showing many features working together. |

To run an example, navigate to the project root and use the `go run` command:

```bash
go run ./examples/01-simple-workbook/main.go
```

## Testing

The project includes a comprehensive test suite.

**Run all unit and integration tests:**

```bash
go test ./tests
```

**Run performance benchmarks:**

The benchmarks will demonstrate the efficiency of the style management system.

```bash
go test -bench=. -benchmem ./tests
```

You will see results comparing file generation with no styles, reused styles (the ideal case), and unique styles (the worst case).

```
Benchmark_LargeFile_NoStyle-10             ...
Benchmark_LargeFile_WithReusedStyles-10    ...
Benchmark_LargeFile_WithUniqueStyles-10    ...
```

## Project Structure

```
go-excelbuilder/
├── pkg/excelbuilder/     # Core library code
├── tests/                # Unit and integration tests
├── examples/             # Standalone, runnable examples
│   ├── 01-simple-workbook/
│   ├── ...
│   └── 07-financial-report/
└── README.md
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Add tests for new functionality
4. Ensure all tests pass
5. Submit a pull request

## License

MIT License - see LICENSE file for details.

## Dependencies

- [excelize/v2](https://github.com/xuri/excelize) - Excel file manipulation library

## Support

If you encounter any issues or have questions, please open an issue on GitHub.
