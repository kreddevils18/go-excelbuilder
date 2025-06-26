package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Define mock data structures
type FinancialStatement struct {
	Period string
	Items  []StatementItem
}

type StatementItem struct {
	Category string
	Value    float64
	IsTotal  bool
	Indent   int
}

type KPI struct {
	Name  string
	Value string
	Note  string
}

type SalesRecord struct {
	Date    string
	Product string
	Region  string
	Sales   float64
}

// Generate mock data
func getMockFinancialData() (map[string]FinancialStatement, []KPI) {
	incomeStatementQ1 := FinancialStatement{
		Period: "Q1 2024",
		Items: []StatementItem{
			{Category: "Revenue", Value: 1200000},
			{Category: "Cost of Goods Sold (COGS)", Value: -450000},
			{Category: "Gross Profit", Value: 750000, IsTotal: true},
			{Category: "Operating Expenses", Value: 0, Indent: 0},
			{Category: "  Sales & Marketing", Value: -150000, Indent: 1},
			{Category: "  Research & Development", Value: -200000, Indent: 1},
			{Category: "  General & Administrative", Value: -100000, Indent: 1},
			{Category: "Total Operating Expenses", Value: -450000, IsTotal: true},
			{Category: "Net Income", Value: 300000, IsTotal: true},
		},
	}

	balanceSheetQ1 := FinancialStatement{
		Period: "As of Q1 2024",
		Items: []StatementItem{
			{Category: "Assets", Value: 0},
			{Category: "  Current Assets", Value: 800000, Indent: 1},
			{Category: "  Long-term Assets", Value: 1200000, Indent: 1},
			{Category: "Total Assets", Value: 2000000, IsTotal: true},
			{Category: "Liabilities", Value: 0},
			{Category: "  Current Liabilities", Value: 300000, Indent: 1},
			{Category: "  Long-term Liabilities", Value: 500000, Indent: 1},
			{Category: "Total Liabilities", Value: 800000, IsTotal: true},
			{Category: "Equity", Value: 1200000, IsTotal: true},
		},
	}

	cashFlowQ1 := FinancialStatement{
		Period: "Q1 2024",
		Items: []StatementItem{
			{Category: "Cash Flow from Operating Activities", Value: 400000},
			{Category: "Cash Flow from Investing Activities", Value: -250000},
			{Category: "Cash Flow from Financing Activities", Value: 50000},
			{Category: "Net Change in Cash", Value: 200000, IsTotal: true},
		},
	}

	kpis := []KPI{
		{Name: "Net Profit Margin", Value: "25.0%", Note: "(Net Income / Revenue)"},
		{Name: "Gross Profit Margin", Value: "62.5%", Note: "(Gross Profit / Revenue)"},
		{Name: "Debt-to-Asset Ratio", Value: "0.40", Note: "(Total Liabilities / Total Assets)"},
	}

	statements := map[string]FinancialStatement{
		"Income Statement": incomeStatementQ1,
		"Balance Sheet":    balanceSheetQ1,
		"Cash Flow":        cashFlowQ1,
	}

	return statements, kpis
}

func getMockRawData() []SalesRecord {
	return []SalesRecord{
		{"2024-01-15", "Laptop", "North", 15000},
		{"2024-01-20", "Monitor", "South", 8000},
		{"2024-02-10", "Laptop", "North", 18000},
		{"2024-02-18", "Keyboard", "East", 2500},
		{"2024-03-05", "Monitor", "West", 9500},
		{"2024-03-12", "Laptop", "South", 22000},
		{"2024-01-28", "Keyboard", "North", 3000},
		{"2024-02-22", "Monitor", "East", 8500},
	}
}

func main() {
	fmt.Println("Generating complex financial report with rich features...")

	// Create a new Excel builder
	builder := excelbuilder.New()

	// Get mock data
	statements, kpis := getMockFinancialData()
	rawData := getMockRawData()

	// --- Centralized Styles ---
	styles := map[string]excelbuilder.StyleConfig{
		"title": {
			Font: excelbuilder.FontConfig{Bold: true, Size: 18, Color: "1F4E78"},
		},
		"header": {
			Font:      excelbuilder.FontConfig{Bold: true, Color: "FFFFFF"},
			Fill:      excelbuilder.FillConfig{Type: "pattern", Color: "4472C4"},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
		},
		"kpiName": {
			Font: excelbuilder.FontConfig{Bold: true, Size: 12},
		},
		"kpiValue": {
			Font:      excelbuilder.FontConfig{Size: 12},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "right"},
		},
		"total": {
			Font: excelbuilder.FontConfig{Bold: true},
		},
		"currency": {
			NumberFormat: `_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)`,
		},
		"totalCurrency": {
			Font:         excelbuilder.FontConfig{Bold: true},
			NumberFormat: `_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)`,
			Border:       excelbuilder.BorderConfig{Top: excelbuilder.BorderSide{Style: "thin"}},
		},
	}

	// --- Build Workbook ---
	wb := builder.NewWorkbook()
	createSummarySheet(wb, kpis, statements["Income Statement"], styles)
	createFinancialStatementSheet(wb, "Income Statement", statements["Income Statement"], styles)
	createFinancialStatementSheet(wb, "Balance Sheet", statements["Balance Sheet"], styles)
	createFinancialStatementSheet(wb, "Cash Flow", statements["Cash Flow"], styles)
	createRawDataSheet(wb, rawData, styles["header"])
	createPivotTableSheet(wb)

	// Set the summary sheet as the active one
	wb.SetActiveSheet("Summary")

	// --- Save File ---
	outputDir := filepath.Join("examples", "07-financial-report", "output")
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}
	filePath := filepath.Join(outputDir, "financial_report_rich.xlsx")

	file := wb.Build()
	if err := file.SaveAs(filePath); err != nil {
		fmt.Println(err)
	}

	fmt.Printf("Report saved to %s\n", filePath)
}

func createSummarySheet(wb *excelbuilder.WorkbookBuilder, kpis []KPI, incomeData FinancialStatement, styles map[string]excelbuilder.StyleConfig) {
	sheet := wb.AddSheet("Summary")
	sheet.SetColumnWidth("A", 30).SetColumnWidth("B", 20).SetColumnWidth("C", 30)

	// --- KPIs Section ---
	sheet.AddRow().AddCell("Financial KPIs").WithStyle(styles["title"]).WithMergeRange("B1")
	sheet.AddRow() // Spacer

	for _, kpi := range kpis {
		sheet.AddRow().
			AddCell(kpi.Name).WithStyle(styles["kpiName"]).Done().
			AddCell(kpi.Value).WithStyle(styles["kpiValue"])
	}
	sheet.AddRow() // Spacer

	// --- Chart Section ---
	sheet.AddRow().AddCell("Income Statement Overview").WithStyle(styles["title"]).WithMergeRange("B1")
	sheet.AddRow() // Spacer

	// --- Add data for the chart ---
	chartHeaderRow := sheet.AddRow()
	chartHeaderRow.AddCell("Category").WithStyle(styles["header"])
	chartHeaderRow.AddCell("Amount").WithStyle(styles["header"])

	// We need to know the row numbers for the chart data range.
	// We'll add the data first and get the row numbers.
	chartDataStartRow := sheet.GetCurrentRow() + 1
	sheet.AddRow().AddCells("Revenue", incomeData.Items[0].Value)
	sheet.AddRow().AddCells("Gross Profit", incomeData.Items[2].Value)
	sheet.AddRow().AddCells("Net Income", incomeData.Items[8].Value)
	chartDataEndRow := sheet.GetCurrentRow()

	// --- Create the chart ---
	chart := sheet.AddChart()
	chart.SetType("col")
	chart.SetTitle("Income Statement Highlights")
	chart.AddDataSeries(excelbuilder.DataSeries{
		Name:       fmt.Sprintf("'Summary'!$B$%d", chartDataStartRow-1),
		Categories: fmt.Sprintf("'Summary'!$A$%d:$A$%d", chartDataStartRow, chartDataEndRow),
		Values:     fmt.Sprintf("'Summary'!$B$%d:$B$%d", chartDataStartRow, chartDataEndRow),
	})
	chart.SetLegend(excelbuilder.LegendConfig{Show: false})
	chart.SetPosition("E5")
	chart.Build()

	// --- Pie Chart Section for Operating Expenses ---
	sheet.AddRow()
	sheet.AddRow().AddCell("Operating Expenses Breakdown").WithStyle(styles["title"]).WithMergeRange("B1")
	sheet.AddRow()

	pieChartHeaderRow := sheet.AddRow()
	pieChartHeaderRow.AddCell("Expense Category").WithStyle(styles["header"])
	pieChartHeaderRow.AddCell("Cost").WithStyle(styles["header"])

	pieChartDataStartRow := sheet.GetCurrentRow() + 1
	// Extracting operating expenses, taking absolute value for the chart
	sheet.AddRow().AddCells("Sales & Marketing", -incomeData.Items[4].Value)
	sheet.AddRow().AddCells("Research & Development", -incomeData.Items[5].Value)
	sheet.AddRow().AddCells("General & Administrative", -incomeData.Items[6].Value)
	pieChartDataEndRow := sheet.GetCurrentRow()

	pieChart := sheet.AddChart()
	pieChart.SetType("pie").SetTitle("Operating Expenses").SetPosition("E25")
	pieChart.AddDataSeries(excelbuilder.DataSeries{
		Name:       fmt.Sprintf("'Summary'!$A$%d", pieChartDataStartRow-1),
		Categories: fmt.Sprintf("'Summary'!$A$%d:$A$%d", pieChartDataStartRow, pieChartDataEndRow),
		Values:     fmt.Sprintf("'Summary'!$B$%d:$B$%d", pieChartDataStartRow, pieChartDataEndRow),
	})
	pieChart.Build()

	sheet.AddRow()

	// Data Validation Section
	sheet.AddRow().AddCell("Data Entry Example").WithStyle(styles["title"]).WithMergeRange("B1")
	sheet.AddRow()

	validation := excelbuilder.DataValidationConfig{
		Type:     "list",
		Formula1: []string{`"North,South,East,West"`},
	}
	sheet.AddRow().
		AddCell("Select Region:").Done().
		AddCell("").WithDataValidation(&validation).Done()
}

func createFinancialStatementSheet(wb *excelbuilder.WorkbookBuilder, sheetName string, data FinancialStatement, styles map[string]excelbuilder.StyleConfig) {
	sheet := wb.AddSheet(sheetName)
	sheet.SetColumnWidth("A", 40).SetColumnWidth("B", 20)

	// --- Title ---
	sheet.AddRow().AddCell(sheetName).WithStyle(styles["title"]).WithMergeRange("B1")
	sheet.AddRow().AddCell(data.Period)
	sheet.AddRow() // Spacer

	// --- Header ---
	headerRow := sheet.AddRow()
	headerRow.AddCell("Category").WithStyle(styles["header"])
	headerRow.AddCell("Amount").WithStyle(styles["header"])

	// --- Freeze header rows ---
	sheet.FreezePanes(0, sheet.GetCurrentRow())

	// --- Data Rows ---
	for _, item := range data.Items {
		row := sheet.AddRow()
		cellStyle := styles["currency"]
		if item.IsTotal {
			cellStyle = styles["totalCurrency"]
		}

		// Add indentation for sub-categories
		category := item.Category
		if item.Indent > 0 {
			for i := 0; i < item.Indent; i++ {
				category = "  " + category
			}
		}

		row.AddCell(category)
		if item.Value != 0 {
			row.AddCell(item.Value).WithStyle(cellStyle)
		} else {
			// Don't style zero-value cells which are headers for sub-items
			row.AddCell("").WithStyle(styles["total"])
		}
	}
}

func createRawDataSheet(wb *excelbuilder.WorkbookBuilder, data []SalesRecord, headerStyle excelbuilder.StyleConfig) {
	sheet := wb.AddSheet("Raw Data")
	sheet.SetColumnWidth("A", 15).SetColumnWidth("B", 20).SetColumnWidth("C", 15).SetColumnWidth("D", 20)

	// Header
	headerRow := sheet.AddRow()
	headerRow.AddCell("Date").WithStyle(headerStyle)
	headerRow.AddCell("Product").WithStyle(headerStyle)
	headerRow.AddCell("Region").WithStyle(headerStyle)
	headerRow.AddCell("Sales").WithStyle(headerStyle)

	// Data
	currencyStyle := excelbuilder.StyleConfig{NumberFormat: "#,##0.00"}
	for _, record := range data {
		sheet.AddRow().
			AddCell(record.Date).Done().
			AddCell(record.Product).Done().
			AddCell(record.Region).Done().
			AddCell(record.Sales).WithStyle(currencyStyle)
	}
}

func createPivotTableSheet(wb *excelbuilder.WorkbookBuilder) {
	sheet := wb.AddSheet("Pivot Table Report")

	// NewPivotTable returns a builder. The source data range is specified here.
	pivotBuilder := sheet.NewPivotTable("Pivot Table Report", "Raw Data!A1:D9")

	// Configure and build the pivot table using the fluent API
	err := pivotBuilder.
		SetTargetCell("A1").            // Use SetTargetCell for the top-left corner
		WithStyle("PivotStyleLight16"). // Use WithStyle for consistency
		AddRowField("Region").          // Use AddRowField for row labels
		AddColumnField("Product").      // Use AddColumnField for column labels
		AddValueField("Sales", "sum").  // Use AddValueField for data, specifying function
		Build()

	if err != nil {
		fmt.Printf("could not build pivot table: %v\n", err)
	}
}
