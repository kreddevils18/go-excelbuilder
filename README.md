# Go Excel Builder

A fluent, builder-pattern library for creating Excel files in Go, built on top of the excellent [excelize](https://github.com/xuri/excelize) library.

## Features

- **Fluent Builder Pattern**: Chain method calls for intuitive Excel file creation
- **Type Safety**: Strong typing with compile-time error checking
- **Memory Efficient**: Optimized for large file generation
- **Comprehensive Testing**: Full test coverage with TDD approach
- **Easy to Use**: Simple, readable API

## Installation

```bash
go get github.com/kreddevils18/go-excelbuilder
```

## Quick Start

```go
package main

import (
    "fmt"
    "log"
    
    "github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

func main() {
    // Create a new Excel builder
    builder := excelbuilder.New()
    
    // Create workbook with properties
    workbook := builder.NewWorkbook()
    workbook.SetProperties(excelbuilder.WorkbookProperties{
        Title:       "My Excel File",
        Author:      "Your Name",
        Subject:     "Demo",
        Description: "Created with Go Excel Builder",
    })
    
    // Add a sheet
    sheet := workbook.AddSheet("Data")
    
    // Set column widths
    sheet.SetColumnWidth("A", 15.0)
    sheet.SetColumnWidth("B", 20.0)
    
    // Add header row
    headerRow := sheet.AddRow()
    headerRow.AddCell("Name")
    headerRow.AddCell("Email")
    
    // Add data rows
    dataRow := sheet.AddRow()
    dataRow.AddCell("John Doe")
    dataRow.AddCell("john@example.com")
    
    // Build and save
    file := workbook.Build()
    err := file.SaveAs("output.xlsx")
    if err != nil {
        log.Fatal(err)
    }
    
    fmt.Println("Excel file created successfully!")
}
```

## API Reference

### ExcelBuilder

- `New()` - Creates a new ExcelBuilder instance
- `NewWorkbook()` - Creates a new WorkbookBuilder

### WorkbookBuilder

- `SetProperties(props WorkbookProperties)` - Sets workbook metadata
- `AddSheet(name string)` - Adds a new sheet and returns SheetBuilder
- `Build()` - Returns the final excelize.File

### SheetBuilder

- `AddRow()` - Adds a new row and returns RowBuilder
- `SetColumnWidth(column string, width float64)` - Sets column width
- `MergeCell(cellRange string)` - Merges cells in the specified range
- `Done()` - Returns to WorkbookBuilder

### RowBuilder

- `AddCell(value interface{})` - Adds a cell and returns CellBuilder
- `SetRowHeight(height float64)` - Sets row height
- `Done()` - Returns to SheetBuilder

### CellBuilder

- `SetStyle(style StyleConfig)` - Sets cell style
- `SetNumberFormat(format string)` - Sets number format
- `SetFormula(formula string)` - Sets cell formula
- `SetHyperlink(url, display string)` - Sets hyperlink
- `Done()` - Returns to RowBuilder

## Configuration Structures

### WorkbookProperties

```go
type WorkbookProperties struct {
    Title       string
    Author      string
    Subject     string
    Description string
    Category    string
    Keywords    string
}
```

### StyleConfig

```go
type StyleConfig struct {
    Font      FontConfig
    Fill      FillConfig
    Border    BorderConfig
    Alignment AlignmentConfig
}
```

## Examples

Check the `examples/` directory for more comprehensive examples:

- `examples/basic/` - Basic usage
- `examples/advanced/` - Advanced features with styling
- `examples/performance/` - Performance optimization examples

## Testing

Run the test suite:

```bash
go test ./tests/...
```

Run tests with verbose output:

```bash
go test -v ./tests/...
```

## Project Structure

```
go-excelbuilder/
├── pkg/excelbuilder/     # Core library code
│   ├── builder.go        # Main builder implementations
│   └── types.go          # Type definitions
├── tests/                # Test files
├── examples/             # Usage examples
├── docs/                 # Documentation
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

## Roadmap

- [ ] Chart support
- [ ] Conditional formatting
- [ ] Data validation
- [ ] Pivot tables
- [ ] Template support
- [ ] Streaming API for very large files

## Support

If you encounter any issues or have questions, please open an issue on GitHub.