# Interactive Dashboard Example

This example demonstrates how to create comprehensive business intelligence dashboards using `go-excelbuilder`. It showcases advanced data visualization, KPI tracking, and multi-perspective business reporting.

## Learning Objectives

- **Dashboard Design**: Create professional business intelligence dashboards
- **KPI Visualization**: Display key performance indicators with visual status indicators
- **Multi-Sheet Analysis**: Organize complex data across multiple specialized sheets
- **Interactive Reporting**: Build reports that support decision-making processes
- **Data Storytelling**: Present business metrics in a compelling, actionable format
- **Alert Systems**: Implement visual alert systems for business monitoring
- **Trend Analysis**: Display performance trends and historical data
- **Executive Reporting**: Create executive-level summary dashboards

## What This Example Creates

The example generates three comprehensive dashboard workbooks:

### 1. Executive Dashboard (`06-executive-dashboard.xlsx`)
- **Executive Summary**: High-level KPIs and performance overview
- **KPI Details**: Detailed analysis of all key performance indicators
- **Performance Trends**: 12-month trend analysis with growth calculations
- **Alerts & Actions**: Business alerts with recommended actions

### 2. Operational Dashboard (`06-operational-dashboard.xlsx`)
- **Operations Overview**: Operational efficiency metrics
- **Process Metrics**: Detailed process performance indicators
- **Resource Utilization**: Resource allocation and utilization analysis

### 3. Financial Dashboard (`06-financial-dashboard.xlsx`)
- **Financial Summary**: Key financial metrics and ratios
- **Revenue Analysis**: Revenue breakdown and analysis
- **Profitability Analysis**: Profit margins and profitability metrics

## How to Run

```bash
cd examples/06-dashboard
go run main.go
```

## Key Code Patterns

### Dashboard Data Structures
```go
type KPIMetric struct {
    Name         string
    CurrentValue float64
    PreviousValue float64
    Target       float64
    Unit         string
    Format       string
}

type DashboardData struct {
    Title        string
    GeneratedAt  time.Time
    KPIs         []KPIMetric
    SalesChart   ChartData
    RegionChart  ChartData
    TrendData    []TrendPoint
    Alerts       []Alert
}
```

### Dashboard Style System
```go
func createDashboardStyles() map[string]excelbuilder.StyleConfig {
    styles := make(map[string]excelbuilder.StyleConfig)
    
    // Color palette for consistent branding
    colors := map[string]string{
        "primary":   "#1E3A8A", // Deep blue
        "success":   "#059669", // Green
        "warning":   "#D97706", // Orange
        "danger":    "#DC2626", // Red
        // ... more colors
    }
    
    // KPI value styling
    styles["kpi_value"] = excelbuilder.StyleConfig{
        Font: excelbuilder.FontConfig{
            Bold:   true,
            Size:   16,
            Color:  colors["primary"],
        },
        // ... more styling
    }
}
```

### Conditional Status Styling
```go
func applyStatusStyling(value, target float64) string {
    progress := (value / target) * 100
    
    if progress >= 100 {
        return "alert_success"
    } else if progress >= 90 {
        return "alert_success"
    } else if progress >= 75 {
        return "alert_warning"
    } else {
        return "alert_danger"
    }
}
```

### KPI Card Creation
```go
func createKPICard(sheet *excelbuilder.SheetBuilder, kpi KPIMetric, styles map[string]excelbuilder.StyleConfig) {
    // KPI name
    sheet.AddRow().AddCell(kpi.Name).SetStyle(styles["kpi_label"])
    
    // Current value with formatting
    valueCell := sheet.AddRow().AddCell(formatKPIValue(kpi))
    valueCell.SetStyle(styles["kpi_value"])
    
    // Change indicator
    change := ((kpi.CurrentValue - kpi.PreviousValue) / kpi.PreviousValue) * 100
    changeStyle := "kpi_positive"
    if change < 0 {
        changeStyle = "kpi_negative"
    }
    
    changeCell := sheet.AddRow().AddCell(fmt.Sprintf("%.1f%%", change))
    changeCell.SetStyle(styles[changeStyle])
}
```

## Dashboard Features

### Executive Dashboard Features
- **KPI Overview**: Six key business metrics with trend indicators
- **Performance Comparison**: Current vs. previous period analysis
- **Target Tracking**: Progress toward business targets
- **Sales Breakdown**: Revenue analysis by product category
- **Regional Performance**: Geographic performance distribution
- **Trend Analysis**: 12-month historical performance
- **Alert System**: Business alerts with priority levels

### Visual Elements
- **Color-Coded Status**: Green/yellow/red status indicators
- **Progress Indicators**: Visual progress toward targets
- **Trend Arrows**: Up/down/stable trend indicators
- **Alert Badges**: Priority-based alert styling
- **Professional Layout**: Clean, executive-ready formatting

### Data Analysis Features
- **Growth Calculations**: Period-over-period growth rates
- **Moving Averages**: 3-month moving average calculations
- **Performance Ratios**: Target achievement percentages
- **Variance Analysis**: Actual vs. target variance
- **Trend Detection**: Automatic trend direction detection

## Business Intelligence Best Practices

### Dashboard Design Principles
1. **Clear Hierarchy**: Most important metrics prominently displayed
2. **Consistent Styling**: Uniform color scheme and typography
3. **Actionable Insights**: Data presented with context and recommendations
4. **Executive Focus**: High-level view with drill-down capability
5. **Visual Clarity**: Clean layout with appropriate white space

### KPI Selection
1. **Strategic Alignment**: KPIs aligned with business objectives
2. **Measurable Metrics**: Quantifiable and trackable indicators
3. **Balanced Scorecard**: Mix of financial and operational metrics
4. **Leading Indicators**: Predictive metrics for future performance
5. **Benchmarking**: Comparison against targets and historical data

### Alert System Design
1. **Priority Levels**: Success, info, warning, danger classifications
2. **Actionable Alerts**: Each alert includes recommended actions
3. **Threshold-Based**: Automatic alerts based on performance thresholds
4. **Visual Distinction**: Color-coded alert levels
5. **Management Escalation**: Clear escalation paths for critical alerts

## Related Examples

- **examples/04-sales-report/**: Sales performance reporting
- **examples/07-financial-analysis/**: Advanced financial modeling
- **examples/08-templates/**: Reusable dashboard templates
- **examples/11-enterprise/**: Enterprise-scale reporting

## Key Concepts Covered

### Dashboard Architecture
- Multi-sheet dashboard organization
- Consistent styling across sheets
- Data structure design for dashboards
- Performance metric calculation

### Business Intelligence
- KPI definition and tracking
- Trend analysis and forecasting
- Alert system implementation
- Executive reporting standards

### Data Visualization
- Color-coded status indicators
- Progress visualization
- Trend direction indicators
- Performance comparison charts

### Excel Features Used
- **Conditional Formatting**: Status-based cell styling
- **Number Formats**: Currency, percentage, and custom formats
- **Cell Styling**: Fonts, colors, borders, and fills
- **Layout Management**: Column widths and row heights
- **Multi-Sheet Workbooks**: Organized data presentation

## Advanced Features

### Dynamic Styling
- Performance-based color coding
- Conditional alert styling
- Trend-based indicators
- Target achievement visualization

### Business Logic
- Growth rate calculations
- Moving average computation
- Performance scoring algorithms
- Alert threshold management

### Professional Presentation
- Executive-ready formatting
- Corporate color schemes
- Consistent typography
- Clean, modern layout

## Next Steps

After mastering this example, explore:

1. **Advanced Analytics**: Implement statistical analysis and forecasting
2. **Interactive Elements**: Add data validation and dropdown controls
3. **Automated Reporting**: Schedule dashboard generation
4. **Custom Visualizations**: Create specialized chart types
5. **Integration**: Connect with external data sources

This example provides a solid foundation for creating professional business intelligence dashboards that support data-driven decision making at all organizational levels.