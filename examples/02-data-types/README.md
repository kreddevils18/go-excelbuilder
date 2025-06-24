# Data Types Comprehensive Example

This example demonstrates comprehensive data type handling in `go-excelbuilder`, showcasing how to work with all supported data types and their proper formatting in Excel.

## 🎯 Learning Objectives

After running this example, you'll understand:

- **Data Type Support**: All data types supported by go-excelbuilder
- **Number Formatting**: Currency, percentage, scientific notation, and custom formats
- **Date/Time Handling**: Various date and time formats and presentations
- **Type-Specific Styling**: Appropriate styling for different data types
- **Mixed Data Structures**: Working with complex data containing multiple types
- **Excel Format Strings**: Understanding Excel's number format syntax
- **Professional Data Presentation**: Best practices for data visualization

## 📋 What This Example Creates

The example generates an Excel file with four comprehensive sheets:

### 1. Basic Data Types Sheet
- **Purpose**: Overview of all supported data types
- **Features**:
  - Complete data type reference table
  - Go type mappings
  - Excel format strings
  - Example values for each type
  - Professional documentation layout

### 2. Numeric Examples Sheet
- **Purpose**: Comprehensive numeric data formatting
- **Features**:
  - Integer formatting with thousands separators
  - Float precision handling
  - Currency formatting with symbols
  - Percentage calculations and display
  - Scientific notation for large/small numbers
  - Negative number handling
  - Mathematical constants demonstration

### 3. Date & Time Examples Sheet
- **Purpose**: Date and time formatting variations
- **Features**:
  - Date-only formatting
  - Date and time combinations
  - Time-only display
  - Custom date formats (e.g., "Monday, January 02, 2006")
  - Historical and future dates
  - Special dates (leap year, epoch, etc.)
  - Timezone considerations

### 4. Complex Data Sheet
- **Purpose**: Real-world mixed data structures
- **Features**:
  - Employee records with mixed data types
  - Proper type-specific formatting for each column
  - Boolean value representation
  - Timestamp handling
  - Professional business data layout

## 🚀 Running the Example

```bash
# Navigate to the example directory
cd examples/02-data-types

# Run the example
go run main.go
```

## 📁 Output

The example creates:
- `output/02-data-types-comprehensive.xlsx` - The generated Excel file
- Console output detailing all demonstrated features

## 🔍 Key Code Patterns

### 1. Type-Specific Style Definitions

```go
// Currency formatting
currencyStyle := excelbuilder.StyleConfig{
    Font: excelbuilder.FontConfig{Size: 10, Family: "Arial"},
    Alignment: excelbuilder.AlignmentConfig{Horizontal: "right", Vertical: "middle"},
    NumberFormat: "$#,##0.00",
    // ... border configuration
}

// Percentage formatting
percentStyle := excelbuilder.StyleConfig{
    NumberFormat: "0.00%",
    // ... other styling
}
```

### 2. Date and Time Formatting

```go
// Date only
dateStyle := excelbuilder.StyleConfig{
    NumberFormat: "yyyy-mm-dd",
    Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
}

// Date and time
datetimeStyle := excelbuilder.StyleConfig{
    NumberFormat: "yyyy-mm-dd hh:mm:ss",
}

// Custom date format
customDateStyle := excelbuilder.StyleConfig{
    NumberFormat: "dddd, mmmm dd, yyyy", // "Monday, January 02, 2006"
}
```

### 3. Scientific Notation

```go
scientificStyle := excelbuilder.StyleConfig{
    NumberFormat: "0.00E+00",
    Alignment: excelbuilder.AlignmentConfig{Horizontal: "right"},
}
```

### 4. Mixed Data Type Handling

```go
complexData := [][]interface{}{
    {1, "John Doe", "john@company.com", 30, 75000.00, 0.15, "2022-01-15", true, timestamp},
    // ... more records
}

for _, data := range complexData {
    row := sheet.AddRow()
    row.AddCell(data[0]).SetStyle(numberStyle)     // ID (int)
    row.AddCell(data[1]).SetStyle(textStyle)       // Name (string)
    row.AddCell(data[2]).SetStyle(textStyle)       // Email (string)
    row.AddCell(data[3]).SetStyle(numberStyle)     // Age (int)
    row.AddCell(data[4]).SetStyle(currencyStyle)   // Salary (float64)
    row.AddCell(data[5]).SetStyle(percentStyle)    // Bonus (float64)
    row.AddCell(data[6]).SetStyle(dateStyle)       // Start Date (string)
    row.AddCell(data[7]).SetStyle(booleanStyle)    // Active (bool)
    row.AddCell(data[8]).SetStyle(datetimeStyle)   // Last Login (string)
}
```

## 📊 Supported Data Types

| Go Type | Excel Display | Number Format | Use Cases |
|---------|---------------|---------------|----------|
| `string` | Text | General | Names, descriptions, IDs |
| `int` | Number | `#,##0` | Counts, quantities, ages |
| `float64` | Decimal | `#,##0.00` | Measurements, calculations |
| `float64` (currency) | Currency | `$#,##0.00` | Money, prices, salaries |
| `float64` (percentage) | Percentage | `0.00%` | Rates, ratios, completion |
| `float64` (scientific) | Scientific | `0.00E+00` | Very large/small numbers |
| `bool` | TRUE/FALSE | General | Flags, status indicators |
| `time.Time` | Date | `yyyy-mm-dd` | Dates without time |
| `time.Time` | DateTime | `yyyy-mm-dd hh:mm:ss` | Timestamps |
| `time.Time` | Time | `hh:mm:ss` | Time without date |

## 💡 Excel Number Format Reference

### Common Format Codes
- `#` - Digit placeholder (shows only significant digits)
- `0` - Digit placeholder (shows zeros)
- `,` - Thousands separator
- `.` - Decimal separator
- `%` - Percentage (multiplies by 100)
- `$` - Currency symbol
- `E+` - Scientific notation

### Date Format Codes
- `yyyy` - 4-digit year
- `mm` - 2-digit month
- `dd` - 2-digit day
- `hh` - 2-digit hour (24-hour)
- `mm` - 2-digit minute
- `ss` - 2-digit second
- `dddd` - Full day name
- `mmmm` - Full month name

## 🎨 Styling Best Practices

1. **Alignment by Type**:
   - Text: Left-aligned
   - Numbers: Right-aligned
   - Dates: Center-aligned
   - Booleans: Center-aligned

2. **Number Formatting**:
   - Always use thousands separators for large numbers
   - Show appropriate decimal places
   - Use currency symbols for money
   - Use percentage format for ratios

3. **Date Formatting**:
   - Choose format based on context
   - Use ISO format (yyyy-mm-dd) for data
   - Use readable format for reports

4. **Boolean Display**:
   - Use TRUE/FALSE for data
   - Consider Yes/No for user-facing reports
   - Make boolean values bold for emphasis

## 🔗 Related Examples

- **Previous**: `examples/01-basic-enhanced/` - Enhanced basic operations
- **Next**: `examples/03-styling-advanced/` - Advanced styling techniques
- **Related**: `examples/04-sales-report/` - Real-world data application

## 🚀 Next Steps

1. **Experiment**: Try different number formats and see the results
2. **Customize**: Create your own data types and formatting styles
3. **Extend**: Add validation for data types
4. **Learn More**: Explore advanced styling in the next example

This example provides a comprehensive foundation for understanding how to handle all data types effectively in `go-excelbuilder`, ensuring your Excel files display data professionally and accurately.