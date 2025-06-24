# Financial Analysis Example

This example demonstrates comprehensive financial modeling and analysis capabilities using `go-excelbuilder`. It creates a complete financial analysis suite including income statements, balance sheets, cash flow statements, ratio analysis, and valuation models.

## Learning Objectives

- **Financial Statement Modeling**: Create professional income statements, balance sheets, and cash flow statements
- **Financial Ratio Analysis**: Calculate and present key financial ratios across multiple categories
- **Valuation Modeling**: Build comprehensive company valuation models using multiple methodologies
- **DCF Analysis**: Implement discounted cash flow models with projections and sensitivity analysis
- **Advanced Financial Styling**: Apply professional financial formatting and presentation standards
- **Multi-Workbook Architecture**: Organize complex financial analysis across multiple specialized workbooks
- **Data-Driven Analysis**: Generate insights from historical data and financial projections
- **Professional Financial Reporting**: Create investment-grade financial analysis reports

## What This Example Creates

This example generates four comprehensive Excel workbooks:

### 1. Financial Statements (`07-financial-statements.xlsx`)
- **Income Statement**: 5-year historical revenue, expenses, and profitability analysis
- **Balance Sheet**: Complete asset, liability, and equity analysis with trend analysis
- **Cash Flow Statement**: Operating, investing, and financing cash flow analysis
- **Financial Summary**: Key metrics and performance indicators dashboard

### 2. Ratio Analysis (`07-ratio-analysis.xlsx`)
- **Ratio Summary**: Comprehensive overview of all financial ratios
- **Liquidity Analysis**: Current ratio, quick ratio, and working capital analysis
- **Profitability Analysis**: Margin analysis, return ratios, and profitability trends
- **Efficiency Analysis**: Asset turnover, inventory turnover, and operational efficiency metrics

### 3. Valuation Model (`07-valuation-model.xlsx`)
- **Valuation Summary**: Multiple valuation methodologies and fair value estimates
- **Trading Multiples**: P/E, P/B, P/S, EV/EBITDA analysis with peer comparisons

### 4. DCF Model (`07-dcf-model.xlsx`)
- **DCF Model**: Complete discounted cash flow analysis with terminal value calculation
- **Sensitivity Analysis**: Scenario analysis with varying growth rates and discount rates

## How to Run

```bash
cd examples/07-financial-analysis
go run main.go
```

The example will:
1. Generate comprehensive financial model data (5 years historical + 5 years projected)
2. Create four specialized financial analysis workbooks
3. Apply professional financial formatting and styling
4. Save all files to the `output/` directory

## Key Code Patterns

### Financial Data Structures
```go
type FinancialStatement struct {
    Period           string
    Revenue          float64
    CostOfGoodsSold  float64
    GrossProfit      float64
    OperatingExpenses float64
    EBITDA           float64
    // ... additional fields
}

type FinancialRatios struct {
    Period              string
    CurrentRatio        float64
    QuickRatio          float64
    DebtToEquity        float64
    ReturnOnAssets      float64
    // ... additional ratios
}
```

### Professional Financial Styling
```go
styles["financial_value"] = excelbuilder.StyleConfig{
    Font: excelbuilder.FontConfig{
        Size:   11,
        Color:  "#374151",
        Family: "Calibri",
    },
    Alignment: excelbuilder.AlignmentConfig{
        Horizontal: "right",
        Vertical:   "middle",
    },
    NumberFormat: "$#,##0",
    Border: createBorder("thin", "#E5E7EB"),
}
```

### Multi-Workbook Financial Analysis
```go
func createFinancialStatementsWorkbook(model FinancialModel) {
    builder := excelbuilder.New()
    workbook := builder.NewWorkbook()
    
    // Create specialized sheets
    incomeSheet := workbook.AddSheet("Income Statement")
    balanceSheet := workbook.AddSheet("Balance Sheet")
    cashFlowSheet := workbook.AddSheet("Cash Flow Statement")
    summarySheet := workbook.AddSheet("Financial Summary")
    
    // Populate with financial data
    createIncomeStatementSheet(incomeSheet, styles, model)
    // ... additional sheet creation
}
```

### Financial Calculations and Projections
```go
// Generate projections with declining growth rates
for i := 0; i < 5; i++ {
    projGrowthRate := 0.12 - float64(i)*0.01 // Declining growth
    revenue := lastRevenue * math.Pow(1+projGrowthRate, float64(i+1))
    
    // Improving margins over time
    cogsRate := 0.58 - float64(i)*0.005 // Improving COGS
    opexRate := 0.23 - float64(i)*0.003 // Improving OpEx
}
```

## Financial Analysis Features

### Comprehensive Financial Statements
- **Income Statement**: Revenue through net income with detailed expense breakdown
- **Balance Sheet**: Complete asset, liability, and equity analysis
- **Cash Flow Statement**: Operating, investing, and financing activities
- **Multi-Year Analysis**: 5-year historical trends and 5-year projections

### Advanced Ratio Analysis
- **Liquidity Ratios**: Current ratio, quick ratio, working capital analysis
- **Profitability Ratios**: Gross margin, operating margin, net margin, ROA, ROE
- **Efficiency Ratios**: Asset turnover, inventory turnover, receivables turnover
- **Leverage Ratios**: Debt-to-equity, interest coverage, debt service coverage

### Valuation Methodologies
- **Trading Multiples**: P/E, P/B, P/S, EV/EBITDA analysis
- **DCF Analysis**: Discounted cash flow with terminal value calculation
- **Sensitivity Analysis**: Multiple scenario analysis with key variable changes
- **Peer Comparison**: Industry benchmark analysis

### Professional Financial Formatting
- **Currency Formatting**: Proper financial number formatting with thousands separators
- **Percentage Display**: Ratio and percentage formatting for financial metrics
- **Conditional Styling**: Positive/negative value highlighting
- **Professional Layout**: Investment-grade presentation standards

## Best Practices Demonstrated

### Financial Modeling Standards
- **Consistent Data Structure**: Standardized financial statement formats
- **Calculation Transparency**: Clear formulas and calculation methodologies
- **Assumption Documentation**: Clear documentation of key assumptions
- **Scenario Analysis**: Multiple case analysis for robust valuation

### Excel Best Practices
- **Professional Formatting**: Investment banking standard formatting
- **Clear Navigation**: Logical sheet organization and naming
- **Data Validation**: Consistent data types and formats
- **Visual Hierarchy**: Clear distinction between headers, data, and totals

### Code Organization
- **Modular Design**: Separate functions for each financial statement
- **Reusable Components**: Common styling and formatting functions
- **Data Separation**: Clear separation between data generation and presentation
- **Error Handling**: Robust error handling for file operations

## Related Examples

- **Previous**: `06-dashboard/` - Interactive business dashboards
- **Next**: `08-templates/` - Reusable financial templates
- **See Also**: 
  - `04-sales-report/` - Business reporting fundamentals
  - `05-import-export/` - Data integration techniques

## Key Concepts Covered

### Financial Analysis
- Income statement analysis and trend identification
- Balance sheet structure and liquidity assessment
- Cash flow analysis and working capital management
- Financial ratio calculation and interpretation
- Company valuation using multiple methodologies
- DCF modeling with terminal value calculation

### Advanced Excel Features
- Professional financial formatting and number formats
- Multi-workbook financial analysis architecture
- Complex data structures for financial modeling
- Conditional formatting for financial indicators
- Professional chart-ready data organization

### Business Intelligence
- Financial performance measurement and KPIs
- Trend analysis and forecasting techniques
- Comparative analysis and benchmarking
- Investment decision support tools
- Risk assessment through sensitivity analysis

## Excel Features Used

- **Advanced Styling**: Professional financial formatting with custom number formats
- **Multi-Sheet Workbooks**: Organized financial analysis across multiple sheets
- **Data Organization**: Structured financial data with proper categorization
- **Professional Presentation**: Investment-grade formatting and layout
- **Complex Calculations**: Financial ratios and valuation metrics

## Advanced Features

### Dynamic Financial Modeling
- **Growth Rate Modeling**: Declining growth rates over projection period
- **Margin Improvement**: Operating leverage and efficiency gains modeling
- **Scenario Analysis**: Multiple case analysis with sensitivity testing
- **Terminal Value**: Perpetual growth and exit multiple methodologies

### Professional Financial Reporting
- **Investment Grade Formatting**: Professional presentation standards
- **Comprehensive Analysis**: Complete financial statement analysis suite
- **Multi-Methodology Valuation**: DCF, multiples, and asset-based approaches
- **Risk Assessment**: Sensitivity analysis and scenario modeling

## Next Steps

After mastering this example, explore:

1. **Template Creation**: Build reusable financial analysis templates
2. **Advanced Modeling**: Monte Carlo simulation and advanced scenario analysis
3. **Industry Analysis**: Sector-specific financial analysis and benchmarking
4. **Integration**: Connect with external data sources and APIs
5. **Automation**: Build automated financial reporting systems

This example provides a comprehensive foundation for professional financial analysis and modeling using `go-excelbuilder`, demonstrating enterprise-grade financial reporting capabilities.