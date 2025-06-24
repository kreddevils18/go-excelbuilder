# Templates Example

This example demonstrates a comprehensive template system for creating reusable Excel reports using `go-excelbuilder`. It showcases how to build flexible, configurable templates that can generate multiple types of business reports with consistent styling and structure.

## Learning Objectives

- **Template System Architecture**: Design and implement reusable Excel template systems
- **Dynamic Content Generation**: Create templates that adapt to different data sources and requirements
- **Template Configuration**: Build flexible template configurations with variables and styling options
- **Multi-Report Generation**: Generate multiple report types from a single template system
- **Style Standardization**: Implement consistent styling across different report templates
- **Template Variations**: Create template variations for different use cases and audiences
- **Business Report Templates**: Build professional templates for common business scenarios
- **Template Reusability**: Design templates that can be easily customized and extended

## What This Example Creates

This example generates five comprehensive Excel workbooks demonstrating different template applications:

### 1. Employee Report (`08-employee-report.xlsx`)
- **Employee Overview**: Complete employee listing with personal and professional details
- **Department Analysis**: Department-wise employee distribution and analysis
- **Salary Analysis**: Compensation analysis and salary distribution insights
- **Employee Directory**: Searchable employee contact directory

### 2. Project Dashboard (`08-project-dashboard.xlsx`)
- **Project Dashboard**: Real-time project status and performance overview
- **Project Status**: Detailed project status tracking and milestone analysis
- **Budget Analysis**: Project budget tracking and financial performance

### 3. Financial Report (`08-financial-report.xlsx`)
- **Executive Summary**: High-level financial performance overview
- **Revenue Analysis**: Detailed revenue trends and analysis
- **Expense Analysis**: Comprehensive expense breakdown and cost analysis

### 4. Inventory Report (`08-inventory-report.xlsx`)
- **Inventory Overview**: Complete inventory status and valuation
- **Category Analysis**: Product category performance and distribution
- **Stock Levels**: Stock level monitoring and reorder analysis

### 5. Custom Variations (`08-custom-variations.xlsx`)
- **Template Comparison**: Side-by-side comparison of different template types
- **Style Variations**: Demonstration of different styling approaches
- **Dynamic Content**: Examples of dynamic content generation techniques

## How to Run

```bash
cd examples/08-templates
go run main.go
```

The example will:
1. Create a comprehensive template system with multiple configurations
2. Generate sample data for different business scenarios
3. Apply templates to create five different types of business reports
4. Demonstrate template customization and variation techniques
5. Save all generated reports to the `output/` directory

## Key Code Patterns

### Template Configuration System
```go
type TemplateConfig struct {
    Name        string
    Description string
    Category    string
    Variables   map[string]interface{}
    Styles      map[string]excelbuilder.StyleConfig
}

type ReportTemplate struct {
    ID          string
    Name        string
    Description string
    Sheets      []SheetTemplate
    GlobalVars  map[string]interface{}
}
```

### Dynamic Template Variables
```go
templates["employee_report"] = TemplateConfig{
    Name:        "Employee Management Report",
    Description: "Comprehensive employee data analysis and reporting",
    Category:    "HR",
    Variables: map[string]interface{}{
        "company_name":    "{{COMPANY_NAME}}",
        "report_date":     "{{REPORT_DATE}}",
        "total_employees": "{{TOTAL_EMPLOYEES}}",
        "departments":     "{{DEPARTMENTS}}",
    },
    Styles: styles,
}
```

### Template-Based Report Generation
```go
func generateEmployeeReport(templates map[string]TemplateConfig) {
    builder := excelbuilder.New()
    workbook := builder.NewWorkbook()
    
    template := templates["employee_report"]
    styles := template.Styles
    
    // Generate sample data
    employees := generateSampleEmployees()
    
    // Create sheets using template
    createEmployeeOverviewSheet(overviewSheet, styles, employees)
    createDepartmentAnalysisSheet(deptSheet, styles, employees)
    // ... additional sheets
}
```

### Reusable Style System
```go
func createTemplateStyles() map[string]excelbuilder.StyleConfig {
    styles := make(map[string]excelbuilder.StyleConfig)
    
    // Template title style
    styles["template_title"] = excelbuilder.StyleConfig{
        Font: excelbuilder.FontConfig{
            Bold:   true,
            Size:   20,
            Color:  "#2563EB",
            Family: "Calibri",
        },
        Alignment: excelbuilder.AlignmentConfig{
            Horizontal: "center",
            Vertical:   "middle",
        },
    }
    
    // Additional styles...
    return styles
}
```

## Template System Features

### Comprehensive Template Architecture
- **Template Configuration**: Flexible configuration system for different report types
- **Variable Substitution**: Dynamic content replacement using template variables
- **Style Inheritance**: Consistent styling across all template-generated reports
- **Multi-Sheet Templates**: Support for complex multi-sheet report structures

### Business Report Templates
- **HR Templates**: Employee management and human resources reporting
- **Project Templates**: Project management and tracking dashboards
- **Financial Templates**: Financial reporting and analysis templates
- **Inventory Templates**: Inventory management and stock tracking reports

### Template Customization
- **Style Variations**: Multiple styling options for different audiences
- **Content Adaptation**: Templates that adapt to different data structures
- **Layout Flexibility**: Configurable layouts for different report requirements
- **Brand Customization**: Easy customization for different organizational brands

### Advanced Template Features
- **Dynamic Sections**: Template sections that adapt based on data availability
- **Conditional Styling**: Style application based on data conditions
- **Template Inheritance**: Base templates that can be extended for specific use cases
- **Multi-Format Support**: Templates that can generate different output formats

## Best Practices Demonstrated

### Template Design Principles
- **Modularity**: Reusable template components and sections
- **Flexibility**: Templates that work with varying data structures
- **Consistency**: Standardized styling and formatting across templates
- **Maintainability**: Easy-to-update template configurations

### Code Organization
- **Separation of Concerns**: Clear separation between template logic and data
- **Configuration Management**: Centralized template configuration system
- **Style Management**: Reusable style definitions and inheritance
- **Data Abstraction**: Generic data structures for template flexibility

### Performance Optimization
- **Efficient Rendering**: Optimized template rendering for large datasets
- **Memory Management**: Efficient memory usage during template processing
- **Caching**: Style and configuration caching for improved performance
- **Batch Processing**: Efficient batch generation of multiple reports

## Related Examples

- **Previous**: `07-financial-analysis/` - Advanced financial modeling
- **Next**: `09-advanced-layout/` - Complex layout management
- **See Also**: 
  - `04-sales-report/` - Business reporting fundamentals
  - `06-dashboard/` - Interactive dashboard creation

## Key Concepts Covered

### Template System Design
- Template configuration and management
- Variable substitution and dynamic content
- Style inheritance and customization
- Multi-template report generation
- Template variation and customization

### Business Intelligence Templates
- Employee management and HR analytics
- Project tracking and management dashboards
- Financial reporting and analysis
- Inventory management and tracking
- Executive summary and KPI reporting

### Advanced Excel Features
- Professional template-based formatting
- Dynamic content generation and layout
- Multi-sheet template architecture
- Conditional styling and formatting
- Template-driven data organization

## Excel Features Used

- **Template-Based Styling**: Consistent professional formatting across reports
- **Dynamic Content**: Variable-driven content generation
- **Multi-Sheet Architecture**: Complex report structures with multiple sheets
- **Professional Formatting**: Business-grade presentation standards
- **Conditional Styling**: Data-driven style application

## Advanced Features

### Template Engine Capabilities
- **Variable Interpolation**: Dynamic content replacement using template variables
- **Style Inheritance**: Hierarchical style system with inheritance
- **Template Composition**: Building complex reports from template components
- **Multi-Format Output**: Templates that support different output requirements

### Business Template Library
- **HR Templates**: Complete human resources reporting suite
- **Project Templates**: Comprehensive project management templates
- **Financial Templates**: Professional financial reporting templates
- **Operational Templates**: Operations and inventory management templates

### Customization Framework
- **Brand Customization**: Easy adaptation for different organizational brands
- **Layout Variations**: Multiple layout options for different use cases
- **Style Themes**: Different visual themes for various audiences
- **Content Adaptation**: Templates that adapt to different data sources

## Template Categories

### Human Resources Templates
- Employee management and directory
- Department analysis and reporting
- Salary and compensation analysis
- Performance tracking and evaluation

### Project Management Templates
- Project dashboard and status tracking
- Budget analysis and financial tracking
- Resource allocation and management
- Timeline and milestone tracking

### Financial Reporting Templates
- Executive financial summaries
- Revenue and expense analysis
- Profit and loss reporting
- Budget vs. actual analysis

### Operational Templates
- Inventory management and tracking
- Supply chain analysis
- Quality control reporting
- Performance metrics and KPIs

## Next Steps

After mastering this example, explore:

1. **Advanced Layouts**: Complex multi-section layouts and advanced formatting
2. **Template Automation**: Automated template generation and deployment
3. **Integration**: Template integration with external data sources
4. **Custom Templates**: Building industry-specific template libraries
5. **Template Marketplace**: Creating shareable template ecosystems

This example provides a comprehensive foundation for building professional template systems using `go-excelbuilder`, demonstrating enterprise-grade template architecture and reusable report generation capabilities.