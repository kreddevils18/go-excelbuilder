# Sales Report Generator

A comprehensive business intelligence example demonstrating how to create professional sales reports with `go-excelbuilder`. This example showcases real-world business scenarios with advanced styling, data analysis, and multi-sheet reporting.

## Learning Objectives

- **Business Intelligence Reporting**: Create executive-ready sales reports
- **Advanced Data Visualization**: Use conditional formatting and performance indicators
- **Multi-Sheet Analysis**: Organize complex data across multiple worksheets
- **Professional Styling**: Apply corporate branding and visual hierarchy
- **KPI Dashboard Creation**: Build key performance indicator summaries
- **Data Aggregation**: Calculate and present business metrics
- **Performance Analytics**: Implement growth tracking and variance analysis

## What This Example Creates

Generates `04-sales-report-q4-2024.xlsx` with six comprehensive sheets:

### 📊 Executive Summary
- **Key Performance Indicators**: Revenue, units, growth, and variance analysis
- **Regional Performance Summary**: High-level regional comparisons
- **Status Indicators**: Color-coded performance ratings
- **Key Insights**: Actionable business recommendations

### 📈 Monthly Performance
- **Trend Analysis**: Month-over-month performance tracking
- **Target vs Actual**: Attainment percentages with conditional formatting
- **Growth Metrics**: Period-over-period growth calculations
- **Quarterly Summary**: Aggregated totals and averages

### 🌍 Regional Analysis
- **Geographic Performance**: Revenue and growth by region
- **Sales Team Metrics**: Representatives per region analysis
- **Margin Analysis**: Profitability by geographic area
- **Comparative Rankings**: Performance-based regional scoring

### 👥 Sales Team Performance
- **Individual Metrics**: Revenue, units, and deals per representative
- **Territory Analysis**: Performance by sales territory
- **Deal Size Analytics**: Average transaction values
- **Team Comparisons**: Relative performance indicators

### 📦 Product Analysis
- **Product Performance**: Units sold and revenue by product
- **Margin Analysis**: Gross margin calculations and percentages
- **Category Breakdown**: Performance by product category
- **Profitability Rankings**: Margin-based product scoring

### 📋 Transaction Details
- **Detailed Records**: Individual transaction history
- **Customer Information**: Account and regional data
- **Discount Analysis**: Pricing and discount tracking
- **Data Relationships**: Linked sales rep and product information

## How to Run

```bash
cd examples/04-sales-report
go run main.go
```

**Output**: `output/04-sales-report-q4-2024.xlsx`

## Key Code Patterns

### Corporate Style System
```go
// Define corporate color palette
colors := map[string]string{
    "corporate_blue":   "#1F4E79",
    "corporate_light":  "#5B9BD5",
    "success_green":    "#70AD47",
    "warning_orange":   "#FFC000",
    "danger_red":       "#C5504B",
}

// Create reusable style configurations
styles["kpi_positive"] = excelbuilder.StyleConfig{
    Font: excelbuilder.FontConfig{
        Bold: true, Size: 12, Color: colors["white"],
    },
    Fill: excelbuilder.FillConfig{
        Type: "pattern", Color: colors["success_green"],
    },
    NumberFormat: "$#,##0",
}
```

### Conditional Performance Styling
```go
func getAttainmentStyle(attainment float64, styles map[string]excelbuilder.StyleConfig) excelbuilder.StyleConfig {
    baseStyle := styles["percentage"]
    if attainment >= 1.1 {
        baseStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#70AD47"}
        baseStyle.Font.Color = "#FFFFFF"
        baseStyle.Font.Bold = true
    } else if attainment >= 1.0 {
        baseStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#FFC000"}
    } else {
        baseStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#C5504B"}
        baseStyle.Font.Color = "#FFFFFF"
    }
    return baseStyle
}
```

### Business Data Structures
```go
type SalesRecord struct {
    ID          int
    SalesRepID  int
    ProductID   string
    Quantity    int
    SaleDate    time.Time
    Discount    float64
    Customer    string
    Region      string
}

type MonthlySummary struct {
    Month      string
    Revenue    float64
    Units      int
    AvgDeal    float64
    Growth     float64
    Target     float64
    Attainment float64
}
```

### KPI Calculation and Display
```go
// Calculate key performance indicators
totalRevenue := 0.0
totalUnits := 0
for _, m := range monthly {
    totalRevenue += m.Revenue
    totalUnits += m.Units
}

// Display with conditional formatting
variance := (kpi.actual - kpi.target) / kpi.target
if variance >= 0.1 {
    status = "Excellent"
    statusStyle = styles["performance_excellent"]
} else if variance >= 0 {
    status = "Good"
    statusStyle = styles["performance_good"]
} else {
    status = "Below Target"
    statusStyle = styles["performance_poor"]
}
```

## Business Intelligence Features

### 📊 Performance Indicators
- **Traffic Light System**: Red/Yellow/Green status indicators
- **Variance Analysis**: Target vs actual comparisons
- **Growth Tracking**: Period-over-period calculations
- **Attainment Metrics**: Goal achievement percentages

### 🎨 Professional Styling
- **Corporate Branding**: Consistent color scheme and typography
- **Visual Hierarchy**: Clear section headers and data organization
- **Conditional Formatting**: Performance-based cell styling
- **Alternating Rows**: Enhanced readability with row striping

### 📈 Data Analysis
- **Aggregation Functions**: Sum, average, and percentage calculations
- **Trend Analysis**: Month-over-month growth tracking
- **Comparative Metrics**: Regional and team performance comparisons
- **Margin Analysis**: Profitability calculations and rankings

## Best Practices Demonstrated

### 🏗️ Code Organization
- **Modular Functions**: Separate sheet creation functions
- **Reusable Styles**: Centralized style management
- **Data Structures**: Well-defined business entities
- **Helper Functions**: Utility functions for calculations

### 📊 Report Design
- **Executive Summary**: High-level overview for leadership
- **Detailed Analysis**: Drill-down capabilities for analysts
- **Visual Consistency**: Uniform styling across all sheets
- **Data Validation**: Proper number formatting and alignment

### 💼 Business Context
- **Real-World Scenarios**: Authentic business use cases
- **Actionable Insights**: Meaningful recommendations
- **Stakeholder Focus**: Different views for different audiences
- **Performance Tracking**: Measurable business metrics

## Related Examples

- **Previous**: [`03-styling-advanced`](../03-styling-advanced/) - Advanced styling techniques
- **Next**: [`05-import-export`](../05-import-export/) - Data integration workflows
- **See Also**: [`06-dashboard`](../06-dashboard/) - Interactive dashboard creation

## Key Concepts Covered

### Business Intelligence
- KPI dashboard creation
- Performance variance analysis
- Trend identification
- Executive reporting

### Data Visualization
- Conditional formatting
- Color-coded performance indicators
- Visual hierarchy design
- Professional chart layouts

### Excel Features
- Multi-sheet workbooks
- Advanced number formatting
- Corporate styling
- Data organization

## Styling Features Used

- **Typography**: Multiple font weights, sizes, and colors
- **Color Schemes**: Corporate branding with semantic colors
- **Borders**: Professional table formatting
- **Fills**: Background colors for emphasis and grouping
- **Alignment**: Proper text and number alignment
- **Number Formats**: Currency, percentage, and integer formatting
- **Conditional Styles**: Performance-based visual indicators

## Next Steps

1. **Customize Data**: Replace sample data with your business metrics
2. **Extend Analysis**: Add more KPIs and performance indicators
3. **Automate Updates**: Connect to live data sources
4. **Add Charts**: Integrate visual charts and graphs
5. **Schedule Reports**: Implement automated report generation

---

*This example demonstrates enterprise-grade reporting capabilities with `go-excelbuilder`, showing how to create professional business intelligence reports that provide actionable insights for stakeholders at all levels.*