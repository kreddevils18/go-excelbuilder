# Import/Export Data Integration Example

This example demonstrates comprehensive data integration capabilities using `go-excelbuilder`, showcasing how to import data from various sources (JSON, CSV), process it, create integrated Excel reports, and export processed data back to different formats.

## Learning Objectives

- **Multi-Source Data Import**: Import from JSON and CSV formats
- **Data Type Handling**: Proper conversion and validation during import
- **Data Integration**: Combine data from multiple sources into unified reports
- **Cross-Data Analysis**: Calculate metrics across different data sources
- **Export Capabilities**: Export processed data to various formats
- **Round-Trip Integrity**: Verify data consistency through import/export cycles
- **Business Intelligence**: Create comprehensive analytical reports
- **Data Quality Assurance**: Implement validation and quality checks

## What This Example Creates

### Input Data Sources
1. **employees.json** - Employee database in JSON format
2. **sales_data.csv** - Sales transactions in CSV format  
3. **inventory.json** - Inventory data in JSON format

### Excel Workbook Output
**File**: `output/05-import-export-integrated.xlsx`

#### Sheet 1: Employee Data
- Complete employee information imported from JSON
- Professional styling with alternating row colors
- Summary statistics (total employees, active count, payroll analysis)
- Proper data type formatting (dates, currency, boolean values)

#### Sheet 2: Sales Analysis
- Sales transactions imported from CSV
- Calculated unit prices and performance metrics
- Regional and product analysis
- Revenue summaries and deal size analytics

#### Sheet 3: Inventory Report
- Inventory items with calculated total values
- Category-based analysis
- Supplier distribution
- Stock value calculations

#### Sheet 4: Data Summary
- Cross-data source analysis
- Business intelligence metrics
- Data quality assessment
- Integration overview

### Exported Data Files
1. **processed_employees.json** - Enhanced employee data with calculated fields
2. **sales_summary.csv** - Regional sales performance summary
3. **inventory_report.json** - Category-based inventory analysis
4. **round_trip_test.xlsx** - Data integrity verification file

## How to Run

```bash
cd examples/05-import-export
go run main.go
```

## Key Code Patterns

### Data Structure Definition
```go
type Employee struct {
    ID         int       `json:"id"`
    FirstName  string    `json:"first_name"`
    LastName   string    `json:"last_name"`
    Email      string    `json:"email"`
    Department string    `json:"department"`
    Position   string    `json:"position"`
    Salary     float64   `json:"salary"`
    HireDate   time.Time `json:"hire_date"`
    Active     bool      `json:"active"`
}
```

### JSON Import Pattern
```go
func importEmployeesFromJSON(filename string) []Employee {
    data, err := os.ReadFile(filename)
    if err != nil {
        log.Fatalf("Failed to read file: %v", err)
    }
    
    var employees []Employee
    err = json.Unmarshal(data, &employees)
    if err != nil {
        log.Fatalf("Failed to parse JSON: %v", err)
    }
    
    return employees
}
```

### CSV Import Pattern
```go
func importSalesFromCSV(filename string) []SalesData {
    file, err := os.Open(filename)
    if err != nil {
        log.Fatalf("Failed to open file: %v", err)
    }
    defer file.Close()
    
    reader := csv.NewReader(file)
    records, err := reader.ReadAll()
    if err != nil {
        log.Fatalf("Failed to read CSV: %v", err)
    }
    
    var salesData []SalesData
    // Skip header row and process data
    for i := 1; i < len(records); i++ {
        // Convert string data to appropriate types
        quantity, _ := strconv.Atoi(records[i][3])
        revenue, _ := strconv.ParseFloat(records[i][4], 64)
        // ... create SalesData struct
    }
    
    return salesData
}
```

### Cross-Data Analysis
```go
// Calculate business metrics across data sources
totalSalesRevenue := calculateTotalRevenue(salesData)
totalPayroll := calculateTotalPayroll(employees)
totalInventoryValue := calculateInventoryValue(inventory)

// Derive business intelligence metrics
revenuePerEmployee := totalSalesRevenue / float64(len(employees))
inventoryTurnover := totalSalesRevenue / totalInventoryValue
```

### Export Processing
```go
// Transform data for export
processedEmployees := make([]map[string]interface{}, len(employees))
for i, emp := range employees {
    processedEmployees[i] = map[string]interface{}{
        "id":           emp.ID,
        "full_name":    fmt.Sprintf("%s %s", emp.FirstName, emp.LastName),
        "annual_salary": emp.Salary,
        "years_service": calculateYearsOfService(emp.HireDate),
        "status":       formatStatus(emp.Active),
    }
}
```

## Data Integration Features

### Import Capabilities
- **JSON Import**: Structured data with proper type mapping
- **CSV Import**: Delimited data with type conversion
- **Data Validation**: Type checking and error handling
- **Large Dataset Support**: Efficient processing of substantial data volumes

### Processing Features
- **Data Transformation**: Convert and enhance imported data
- **Cross-Reference Analysis**: Link data across different sources
- **Calculated Fields**: Derive new metrics from existing data
- **Data Aggregation**: Summarize and group data for analysis

### Export Capabilities
- **JSON Export**: Structured data with enhanced fields
- **CSV Export**: Summarized data for external analysis
- **Excel Integration**: Rich formatting and multiple sheets
- **Round-Trip Testing**: Verify data integrity through complete cycles

## Best Practices Demonstrated

### Data Handling
- **Type Safety**: Proper Go struct definitions with JSON tags
- **Error Handling**: Comprehensive error checking throughout import/export
- **Memory Efficiency**: Stream processing for large datasets
- **Data Validation**: Verify data integrity during processing

### Excel Integration
- **Multi-Sheet Organization**: Logical separation of different data types
- **Professional Styling**: Consistent formatting across all sheets
- **Data Type Formatting**: Appropriate number formats for different data types
- **Summary Analytics**: Calculated metrics and business intelligence

### File Management
- **Directory Organization**: Separate input and output directories
- **File Naming Conventions**: Clear, descriptive filenames
- **Error Recovery**: Graceful handling of missing or corrupted files
- **Progress Reporting**: User feedback during processing

## Related Examples

- **Previous**: [04-sales-report](../04-sales-report/) - Business reporting foundations
- **Next**: [06-dashboard](../06-dashboard/) - Interactive dashboard creation
- **See Also**: [02-data-types](../02-data-types/) - Data type handling basics

## Key Concepts Covered

### Data Integration
- Multi-source data import
- Data format conversion
- Cross-data analysis
- Business intelligence metrics

### File I/O Operations
- JSON marshaling/unmarshaling
- CSV reading/writing
- File system operations
- Error handling patterns

### Excel Features Used
- Multiple sheet management
- Advanced styling systems
- Data type formatting
- Summary calculations

### Business Applications
- Employee data management
- Sales performance analysis
- Inventory tracking
- Cross-functional reporting

## Advanced Features

### Data Quality Assurance
- Import validation
- Type conversion verification
- Cross-reference integrity checks
- Round-trip testing

### Performance Optimization
- Efficient data structures
- Memory-conscious processing
- Batch operations
- Stream processing capabilities

### Extensibility
- Modular import functions
- Configurable export formats
- Pluggable data transformations
- Scalable architecture patterns

This example provides a comprehensive foundation for building data integration systems that can handle real-world business requirements while maintaining data quality and providing rich analytical capabilities.