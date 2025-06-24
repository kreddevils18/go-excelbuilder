# Enhanced Basic Example

This example demonstrates the core features of `go-excelbuilder` with a comprehensive, real-world approach that goes beyond simple "Hello World" scenarios.

## 🎯 Learning Objectives

After running this example, you'll understand:

- **Workbook Management**: Setting comprehensive workbook properties and metadata
- **Multi-Sheet Operations**: Creating and managing multiple sheets with different purposes
- **Advanced Styling**: Using reusable style configurations for consistent formatting
- **Data Type Handling**: Working with strings, numbers, dates, and mixed data types
- **Layout Management**: Setting column widths and row heights for professional presentation
- **Professional Formatting**: Creating business-ready reports with proper styling

## 📋 What This Example Creates

The example generates an Excel file with three sheets:

### 1. Employee Data Sheet
- **Purpose**: Detailed employee information
- **Features**: 
  - Professional header styling with blue background
  - Proper column widths for readability
  - Mixed data types (text, numbers, emails, dates)
  - Consistent cell borders and alignment
  - Number formatting for salary and age columns

### 2. Department Summary Sheet
- **Purpose**: Aggregated department statistics
- **Features**:
  - Calculated summary data
  - Number formatting for financial data
  - Professional table layout
  - Consistent styling with the main data sheet

### 3. Company Overview Sheet
- **Purpose**: High-level company information
- **Features**:
  - Key-value pair layout
  - Alternating row colors for better readability
  - Bold labels for emphasis
  - Dynamic timestamp generation

## 🚀 Running the Example

```bash
# Navigate to the example directory
cd examples/01-basic-enhanced

# Run the example
go run main.go
```

## 📁 Output

The example creates:
- `output/01-basic-enhanced-example.xlsx` - The generated Excel file
- Comprehensive console output showing what features were demonstrated

## 🔍 Key Code Patterns

### 1. Reusable Style Definitions

```go
headerStyle := excelbuilder.StyleConfig{
    Font: excelbuilder.FontConfig{
        Bold:   true,
        Size:   12,
        Color:  "#FFFFFF",
        Family: "Arial",
    },
    Fill: excelbuilder.FillConfig{
        Type:  "pattern",
        Color: "#4472C4",
    },
    // ... more styling options
}
```

### 2. Professional Column Width Management

```go
employeeSheet.SetColumnWidth("A", 15.0) // Employee ID
employeeSheet.SetColumnWidth("B", 25.0) // Full Name
employeeSheet.SetColumnWidth("C", 30.0) // Email
```

### 3. Comprehensive Workbook Properties

```go
workbook.SetProperties(excelbuilder.WorkbookProperties{
    Title:       "Enhanced Basic Example",
    Author:      "Go Excel Builder",
    Subject:     "Demonstration of Basic Features",
    Description: "This example showcases enhanced basic operations...",
    // ... more metadata
})
```

### 4. Mixed Data Type Handling

```go
employees := [][]interface{}{
    {"EMP001", "John Doe", "john.doe@company.com", "Engineering", 30, 75000.00, "2022-01-15"},
    // ... more employee records
}
```

## 💡 Best Practices Demonstrated

1. **Style Consistency**: Define styles once and reuse them across cells
2. **Data Organization**: Separate different types of information into logical sheets
3. **Professional Formatting**: Use appropriate fonts, colors, and alignment
4. **Column Sizing**: Set appropriate widths based on content type
5. **Error Handling**: Check for nil returns and handle errors gracefully
6. **Documentation**: Provide clear output messages about what was created

## 🔗 Related Examples

- **Previous**: `examples/basic/` - Simple basic usage
- **Next**: `examples/02-data-types/` - Deep dive into data type handling
- **Advanced**: `examples/03-styling-advanced/` - Advanced styling techniques

## 📚 Key Concepts Covered

| Concept | Description | Code Location |
|---------|-------------|---------------|
| Workbook Properties | Setting metadata and document properties | Lines 20-30 |
| Style Configuration | Creating reusable style objects | Lines 32-80 |
| Multi-Sheet Management | Creating and organizing multiple sheets | Lines 82+ |
| Column Width Setting | Professional layout management | Lines 88-94 |
| Data Type Handling | Working with mixed data types | Lines 110-120 |
| Number Formatting | Formatting numeric data appropriately | Lines 65-75 |
| Error Handling | Proper error checking and reporting | Throughout |

## 🎨 Styling Features Used

- **Font Configuration**: Bold, size, color, family
- **Fill Patterns**: Background colors and patterns
- **Border Styling**: Complete border configuration
- **Alignment**: Horizontal and vertical alignment
- **Number Formatting**: Currency and decimal formatting
- **Alternating Row Colors**: Enhanced readability

## 🚀 Next Steps

1. **Experiment**: Modify the employee data and see how it affects the output
2. **Customize**: Change the color scheme and styling to match your brand
3. **Extend**: Add more sheets or data types
4. **Learn More**: Explore the data types example for advanced data handling

This example serves as a solid foundation for understanding how to create professional, well-formatted Excel files using `go-excelbuilder`.