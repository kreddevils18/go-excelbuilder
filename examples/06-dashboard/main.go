package main

import (
	"fmt"
	"log"
	"math"
	"math/rand"
	"os"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Dashboard data structures
type KPIMetric struct {
	Name         string
	CurrentValue float64
	PreviousValue float64
	Target       float64
	Unit         string
	Format       string
}

type ChartData struct {
	Labels []string
	Values []float64
	Colors []string
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

type TrendPoint struct {
	Date  time.Time
	Value float64
	Label string
}

type Alert struct {
	Level   string // "success", "warning", "danger", "info"
	Title   string
	Message string
	Value   float64
}

func main() {
	// Create output directory
	if err := os.MkdirAll("output", 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}

	fmt.Println("📊 Starting Interactive Dashboard Demo...")

	// Step 1: Generate dashboard data
	fmt.Println("\n📈 Step 1: Generating dashboard data...")
	dashboardData := generateDashboardData()

	// Step 2: Create executive dashboard
	fmt.Println("\n🎯 Step 2: Creating executive dashboard...")
	createExecutiveDashboard(dashboardData)

	// Step 3: Create operational dashboard
	fmt.Println("\n⚙️ Step 3: Creating operational dashboard...")
	createOperationalDashboard(dashboardData)

	// Step 4: Create financial dashboard
	fmt.Println("\n💰 Step 4: Creating financial dashboard...")
	createFinancialDashboard(dashboardData)

	fmt.Println("\n✅ Dashboard demonstration completed successfully!")
	fmt.Println("📁 Generated files:")
	fmt.Println("   • output/06-executive-dashboard.xlsx - Executive summary dashboard")
	fmt.Println("   • output/06-operational-dashboard.xlsx - Operational metrics dashboard")
	fmt.Println("   • output/06-financial-dashboard.xlsx - Financial analysis dashboard")
	fmt.Println("\n🎯 Next steps: Try examples/07-financial-analysis/ for advanced financial modeling")
}

func generateDashboardData() DashboardData {
	fmt.Println("   📊 Generating KPI metrics...")
	
	// Generate realistic business KPIs
	kpis := []KPIMetric{
		{
			Name:          "Monthly Revenue",
			CurrentValue:  2847500.00,
			PreviousValue: 2650000.00,
			Target:        3000000.00,
			Unit:          "$",
			Format:        "currency",
		},
		{
			Name:          "Customer Acquisition",
			CurrentValue:  1247,
			PreviousValue: 1180,
			Target:        1300,
			Unit:          "customers",
			Format:        "integer",
		},
		{
			Name:          "Conversion Rate",
			CurrentValue:  0.0847,
			PreviousValue: 0.0792,
			Target:        0.0900,
			Unit:          "%",
			Format:        "percentage",
		},
		{
			Name:          "Customer Satisfaction",
			CurrentValue:  4.7,
			PreviousValue: 4.5,
			Target:        4.8,
			Unit:          "/5",
			Format:        "decimal",
		},
		{
			Name:          "Market Share",
			CurrentValue:  0.1847,
			PreviousValue: 0.1792,
			Target:        0.2000,
			Unit:          "%",
			Format:        "percentage",
		},
		{
			Name:          "Employee Productivity",
			CurrentValue:  127.5,
			PreviousValue: 122.8,
			Target:        130.0,
			Unit:          "index",
			Format:        "decimal",
		},
	}

	fmt.Println("   📈 Generating sales chart data...")
	
	// Sales by product category
	salesChart := ChartData{
		Labels: []string{"Electronics", "Software", "Services", "Hardware", "Accessories"},
		Values: []float64{1247500, 892300, 456700, 678900, 234100},
		Colors: []string{"#2E86AB", "#A23B72", "#F18F01", "#C73E1D", "#6A994E"},
	}

	fmt.Println("   🌍 Generating regional chart data...")
	
	// Sales by region
	regionChart := ChartData{
		Labels: []string{"North America", "Europe", "Asia Pacific", "Latin America", "Middle East"},
		Values: []float64{1456700, 987500, 823400, 345600, 234300},
		Colors: []string{"#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4", "#FFEAA7"},
	}

	fmt.Println("   📊 Generating trend data...")
	
	// Generate 12 months of trend data
	trendData := make([]TrendPoint, 12)
	baseValue := 2000000.0
	for i := 0; i < 12; i++ {
		date := time.Now().AddDate(0, -11+i, 0)
		// Add some realistic variation
		variation := (rand.Float64() - 0.5) * 0.3 // ±15% variation
		seasonality := math.Sin(float64(i)*math.Pi/6) * 0.1 // Seasonal pattern
		growth := float64(i) * 0.02 // 2% monthly growth
		
		value := baseValue * (1 + growth + seasonality + variation)
		
		trendData[i] = TrendPoint{
			Date:  date,
			Value: value,
			Label: date.Format("Jan 2006"),
		}
	}

	fmt.Println("   🚨 Generating alerts...")
	
	// Generate business alerts
	alerts := []Alert{
		{
			Level:   "success",
			Title:   "Revenue Target Exceeded",
			Message: "Q4 revenue exceeded target by 7.5%",
			Value:   107.5,
		},
		{
			Level:   "warning",
			Title:   "Inventory Levels Low",
			Message: "Electronics inventory below reorder point",
			Value:   23.4,
		},
		{
			Level:   "info",
			Title:   "New Market Opportunity",
			Message: "Emerging market showing 45% growth potential",
			Value:   145.0,
		},
		{
			Level:   "danger",
			Title:   "Customer Churn Alert",
			Message: "Churn rate increased to 5.7% (target: 4.0%)",
			Value:   5.7,
		},
	}

	return DashboardData{
		Title:       "Business Intelligence Dashboard",
		GeneratedAt: time.Now(),
		KPIs:        kpis,
		SalesChart:  salesChart,
		RegionChart: regionChart,
		TrendData:   trendData,
		Alerts:      alerts,
	}
}

func createExecutiveDashboard(data DashboardData) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Executive Dashboard",
		Author:      "Business Intelligence System",
		Subject:     "Executive Summary and KPI Tracking",
		Description: "High-level business metrics and performance indicators for executive decision making",
		Keywords:    "dashboard,executive,kpi,business-intelligence,metrics",
		Company:     "TechCorp Solutions",
	})

	// Create dashboard styles
	styles := createDashboardStyles()

	// Sheet 1: Executive Summary
	summarySheet := workbook.AddSheet("Executive Summary")
	if summarySheet == nil {
		log.Fatal("Failed to create executive summary sheet")
	}
	createExecutiveSummarySheet(summarySheet, styles, data)

	// Sheet 2: KPI Details
	kpiSheet := workbook.AddSheet("KPI Details")
	if kpiSheet == nil {
		log.Fatal("Failed to create KPI details sheet")
	}
	createKPIDetailsSheet(kpiSheet, styles, data)

	// Sheet 3: Performance Trends
	trendSheet := workbook.AddSheet("Performance Trends")
	if trendSheet == nil {
		log.Fatal("Failed to create trends sheet")
	}
	createPerformanceTrendsSheet(trendSheet, styles, data)

	// Sheet 4: Alerts & Actions
	alertsSheet := workbook.AddSheet("Alerts & Actions")
	if alertsSheet == nil {
		log.Fatal("Failed to create alerts sheet")
	}
	createAlertsSheet(alertsSheet, styles, data)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build executive dashboard")
	}

	filename := "output/06-executive-dashboard.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save executive dashboard: %v", err)
	}

	fmt.Printf("   ✓ Created executive dashboard: %s\n", filename)
}

func createOperationalDashboard(data DashboardData) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Operational Dashboard",
		Author:      "Operations Management System",
		Subject:     "Operational Metrics and Process Monitoring",
		Description: "Detailed operational metrics for day-to-day business management and process optimization",
		Keywords:    "dashboard,operations,metrics,monitoring,efficiency",
		Company:     "TechCorp Solutions",
	})

	// Create styles
	styles := createDashboardStyles()

	// Sheet 1: Operations Overview
	overviewSheet := workbook.AddSheet("Operations Overview")
	if overviewSheet == nil {
		log.Fatal("Failed to create operations overview sheet")
	}
	createOperationsOverviewSheet(overviewSheet, styles, data)

	// Sheet 2: Process Metrics
	processSheet := workbook.AddSheet("Process Metrics")
	if processSheet == nil {
		log.Fatal("Failed to create process metrics sheet")
	}
	createProcessMetricsSheet(processSheet, styles, data)

	// Sheet 3: Resource Utilization
	resourceSheet := workbook.AddSheet("Resource Utilization")
	if resourceSheet == nil {
		log.Fatal("Failed to create resource utilization sheet")
	}
	createResourceUtilizationSheet(resourceSheet, styles, data)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build operational dashboard")
	}

	filename := "output/06-operational-dashboard.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save operational dashboard: %v", err)
	}

	fmt.Printf("   ✓ Created operational dashboard: %s\n", filename)
}

func createFinancialDashboard(data DashboardData) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Financial Dashboard",
		Author:      "Financial Analysis System",
		Subject:     "Financial Performance and Analysis",
		Description: "Comprehensive financial metrics, analysis, and forecasting for strategic financial management",
		Keywords:    "dashboard,financial,analysis,revenue,profitability",
		Company:     "TechCorp Solutions",
	})

	// Create styles
	styles := createDashboardStyles()

	// Sheet 1: Financial Summary
	financialSheet := workbook.AddSheet("Financial Summary")
	if financialSheet == nil {
		log.Fatal("Failed to create financial summary sheet")
	}
	createFinancialSummarySheet(financialSheet, styles, data)

	// Sheet 2: Revenue Analysis
	revenueSheet := workbook.AddSheet("Revenue Analysis")
	if revenueSheet == nil {
		log.Fatal("Failed to create revenue analysis sheet")
	}
	createRevenueAnalysisSheet(revenueSheet, styles, data)

	// Sheet 3: Profitability Analysis
	profitSheet := workbook.AddSheet("Profitability Analysis")
	if profitSheet == nil {
		log.Fatal("Failed to create profitability analysis sheet")
	}
	createProfitabilityAnalysisSheet(profitSheet, styles, data)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build financial dashboard")
	}

	filename := "output/06-financial-dashboard.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save financial dashboard: %v", err)
	}

	fmt.Printf("   ✓ Created financial dashboard: %s\n", filename)
}

func createDashboardStyles() map[string]excelbuilder.StyleConfig {
	styles := make(map[string]excelbuilder.StyleConfig)

	// Color palette for dashboards
	colors := map[string]string{
		"primary":      "#1E3A8A", // Deep blue
		"secondary":    "#7C3AED", // Purple
		"success":      "#059669", // Green
		"warning":      "#D97706", // Orange
		"danger":       "#DC2626", // Red
		"info":         "#0284C7", // Light blue
		"light":        "#F8FAFC", // Very light gray
		"dark":         "#1F2937", // Dark gray
		"white":        "#FFFFFF",
		"border":       "#E5E7EB", // Light border
		"text":         "#374151", // Text gray
		"accent1":      "#F59E0B", // Amber
		"accent2":      "#8B5CF6", // Violet
	}

	// Dashboard title style
	styles["dashboard_title"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   20,
			Color:  colors["primary"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	}

	// Section headers
	styles["section_header"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   14,
			Color:  colors["white"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["primary"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createBorder("medium", colors["border"]),
	}

	// KPI card styles
	styles["kpi_label"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   11,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["light"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["border"]),
	}

	styles["kpi_value"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   16,
			Color:  colors["primary"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["border"]),
	}

	styles["kpi_positive"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
			Color:  colors["success"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["border"]),
	}

	styles["kpi_negative"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
			Color:  colors["danger"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["border"]),
	}

	// Alert styles
	styles["alert_success"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   10,
			Color:  colors["white"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["success"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["border"]),
	}

	styles["alert_warning"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   10,
			Color:  colors["white"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["warning"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["border"]),
	}

	styles["alert_danger"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   10,
			Color:  colors["white"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["danger"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["border"]),
	}

	styles["alert_info"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   10,
			Color:  colors["white"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["info"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["border"]),
	}

	// Data table styles
	styles["table_header"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   11,
			Color:  colors["white"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["secondary"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["border"]),
	}

	styles["data_normal"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["border"]),
	}

	styles["data_alternate"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["light"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["border"]),
	}

	// Number format styles
	styles["currency"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "$#,##0.00",
		Border:      createBorder("thin", colors["border"]),
	}

	styles["percentage"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "0.00%",
		Border:      createBorder("thin", colors["border"]),
	}

	styles["integer"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "#,##0",
		Border:      createBorder("thin", colors["border"]),
	}

	styles["decimal"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "0.00",
		Border:      createBorder("thin", colors["border"]),
	}

	return styles
}

func createBorder(style, color string) excelbuilder.BorderConfig {
	return excelbuilder.BorderConfig{
		Top:    excelbuilder.BorderSide{Style: style, Color: color},
		Bottom: excelbuilder.BorderSide{Style: style, Color: color},
		Left:   excelbuilder.BorderSide{Style: style, Color: color},
		Right:  excelbuilder.BorderSide{Style: style, Color: color},
	}
}

func createExecutiveSummarySheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, data DashboardData) {
	// Set column widths for dashboard layout
	sheet.SetColumnWidth("A", 3.0)   // Spacer
	sheet.SetColumnWidth("B", 20.0)  // Labels
	sheet.SetColumnWidth("C", 15.0)  // Values
	sheet.SetColumnWidth("D", 12.0)  // Change
	sheet.SetColumnWidth("E", 3.0)   // Spacer
	sheet.SetColumnWidth("F", 20.0)  // Labels
	sheet.SetColumnWidth("G", 15.0)  // Values
	sheet.SetColumnWidth("H", 12.0)  // Change

	// Dashboard title
	titleRow := sheet.AddRow()
	titleRow.AddCell("") // Spacer
	titleRow.AddCell("EXECUTIVE DASHBOARD").SetStyle(styles["dashboard_title"])

	sheet.AddRow() // Empty row
	sheet.AddRow() // Empty row

	// Generated timestamp
	timestampRow := sheet.AddRow()
	timestampRow.AddCell("") // Spacer
	timestampRow.AddCell(fmt.Sprintf("Generated: %s", data.GeneratedAt.Format("January 2, 2006 at 3:04 PM"))).SetStyle(styles["data_normal"])

	sheet.AddRow() // Empty row

	// KPI Section Header
	kpiHeaderRow := sheet.AddRow()
	kpiHeaderRow.AddCell("") // Spacer
	kpiHeaderRow.AddCell("KEY PERFORMANCE INDICATORS").SetStyle(styles["section_header"])

	sheet.AddRow() // Empty row

	// KPI Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("") // Spacer
	headerRow.AddCell("Metric").SetStyle(styles["table_header"])
	headerRow.AddCell("Current Value").SetStyle(styles["table_header"])
	headerRow.AddCell("vs Previous").SetStyle(styles["table_header"])
	headerRow.AddCell("") // Spacer
	headerRow.AddCell("Metric").SetStyle(styles["table_header"])
	headerRow.AddCell("Current Value").SetStyle(styles["table_header"])
	headerRow.AddCell("vs Target").SetStyle(styles["table_header"])

	// Display KPIs in two columns
	for i := 0; i < len(data.KPIs); i += 2 {
		row := sheet.AddRow()
		
		// Left column KPI
		kpi1 := data.KPIs[i]
		row.AddCell("") // Spacer
		row.AddCell(kpi1.Name).SetStyle(styles["kpi_label"])
		
		// Format value based on type
		value1Cell := row.AddCell(formatKPIValue(kpi1))
		value1Cell.SetStyle(styles["kpi_value"])
		
		// Calculate and display change
		change1 := ((kpi1.CurrentValue - kpi1.PreviousValue) / kpi1.PreviousValue) * 100
		change1Cell := row.AddCell(fmt.Sprintf("%.1f%%", change1))
		if change1 >= 0 {
			change1Cell.SetStyle(styles["kpi_positive"])
		} else {
			change1Cell.SetStyle(styles["kpi_negative"])
		}
		
		// Right column KPI (if exists)
		if i+1 < len(data.KPIs) {
			kpi2 := data.KPIs[i+1]
			row.AddCell("") // Spacer
			row.AddCell(kpi2.Name).SetStyle(styles["kpi_label"])
			
			value2Cell := row.AddCell(formatKPIValue(kpi2))
			value2Cell.SetStyle(styles["kpi_value"])
			
			// Calculate progress to target
			progress := (kpi2.CurrentValue / kpi2.Target) * 100
			progressCell := row.AddCell(fmt.Sprintf("%.1f%%", progress))
			if progress >= 90 {
				progressCell.SetStyle(styles["kpi_positive"])
			} else if progress >= 75 {
				progressCell.SetStyle(styles["alert_warning"])
			} else {
				progressCell.SetStyle(styles["kpi_negative"])
			}
		}
	}

	// Sales Performance Section
	sheet.AddRow() // Empty row
	sheet.AddRow() // Empty row

	salesHeaderRow := sheet.AddRow()
	salesHeaderRow.AddCell("") // Spacer
	salesHeaderRow.AddCell("SALES PERFORMANCE BY CATEGORY").SetStyle(styles["section_header"])

	sheet.AddRow() // Empty row

	// Sales chart data headers
	salesChartHeaderRow := sheet.AddRow()
	salesChartHeaderRow.AddCell("") // Spacer
	salesChartHeaderRow.AddCell("Category").SetStyle(styles["table_header"])
	salesChartHeaderRow.AddCell("Revenue").SetStyle(styles["table_header"])
	salesChartHeaderRow.AddCell("% of Total").SetStyle(styles["table_header"])

	// Calculate total for percentages
	totalSales := 0.0
	for _, value := range data.SalesChart.Values {
		totalSales += value
	}

	// Display sales data
	for i, label := range data.SalesChart.Labels {
		row := sheet.AddRow()
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell("") // Spacer
		row.AddCell(label).SetStyle(rowStyle)
		row.AddCell(data.SalesChart.Values[i]).SetStyle(styles["currency"])
		
		percentage := (data.SalesChart.Values[i] / totalSales) * 100
		row.AddCell(percentage / 100).SetStyle(styles["percentage"])
	}

	// Total row
	totalRow := sheet.AddRow()
	totalRow.AddCell("") // Spacer
	totalRow.AddCell("TOTAL").SetStyle(styles["table_header"])
	totalRow.AddCell(totalSales).SetStyle(styles["currency"])
	totalRow.AddCell(1.0).SetStyle(styles["percentage"])
}

func createKPIDetailsSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, data DashboardData) {
	// Set column widths
	sheet.SetColumnWidth("A", 25.0)  // KPI Name
	sheet.SetColumnWidth("B", 15.0)  // Current Value
	sheet.SetColumnWidth("C", 15.0)  // Previous Value
	sheet.SetColumnWidth("D", 15.0)  // Target Value
	sheet.SetColumnWidth("E", 12.0)  // Change %
	sheet.SetColumnWidth("F", 12.0)  // Target %
	sheet.SetColumnWidth("G", 15.0)  // Status
	sheet.SetColumnWidth("H", 30.0)  // Comments

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("KPI DETAILED ANALYSIS").SetStyle(styles["dashboard_title"])

	sheet.AddRow() // Empty row
	sheet.AddRow() // Empty row

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("KPI Metric").SetStyle(styles["table_header"])
	headerRow.AddCell("Current Value").SetStyle(styles["table_header"])
	headerRow.AddCell("Previous Value").SetStyle(styles["table_header"])
	headerRow.AddCell("Target Value").SetStyle(styles["table_header"])
	headerRow.AddCell("Change %").SetStyle(styles["table_header"])
	headerRow.AddCell("Target %").SetStyle(styles["table_header"])
	headerRow.AddCell("Status").SetStyle(styles["table_header"])
	headerRow.AddCell("Performance Notes").SetStyle(styles["table_header"])

	// KPI details
	for i, kpi := range data.KPIs {
		row := sheet.AddRow()
		
		// Alternating row colors
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell(kpi.Name).SetStyle(rowStyle)
		
		// Format values based on type
		currentCell := row.AddCell(kpi.CurrentValue)
		previousCell := row.AddCell(kpi.PreviousValue)
		targetCell := row.AddCell(kpi.Target)
		
		switch kpi.Format {
		case "currency":
			currentCell.SetStyle(styles["currency"])
			previousCell.SetStyle(styles["currency"])
			targetCell.SetStyle(styles["currency"])
		case "percentage":
			currentCell.SetStyle(styles["percentage"])
			previousCell.SetStyle(styles["percentage"])
			targetCell.SetStyle(styles["percentage"])
		case "integer":
			currentCell.SetStyle(styles["integer"])
			previousCell.SetStyle(styles["integer"])
			targetCell.SetStyle(styles["integer"])
		default:
			currentCell.SetStyle(styles["decimal"])
			previousCell.SetStyle(styles["decimal"])
			targetCell.SetStyle(styles["decimal"])
		}
		
		// Calculate changes and progress
		change := ((kpi.CurrentValue - kpi.PreviousValue) / kpi.PreviousValue) * 100
		progress := (kpi.CurrentValue / kpi.Target) * 100
		
		changeCell := row.AddCell(change / 100)
		changeCell.SetStyle(styles["percentage"])
		
		progressCell := row.AddCell(progress / 100)
		progressCell.SetStyle(styles["percentage"])
		
		// Status determination
		var status string
		var statusStyle excelbuilder.StyleConfig
		
		if progress >= 100 {
			status = "Target Achieved"
			statusStyle = styles["alert_success"]
		} else if progress >= 90 {
			status = "On Track"
			statusStyle = styles["alert_success"]
		} else if progress >= 75 {
			status = "Needs Attention"
			statusStyle = styles["alert_warning"]
		} else {
			status = "Below Target"
			statusStyle = styles["alert_danger"]
		}
		
		row.AddCell(status).SetStyle(statusStyle)
		
		// Performance notes
		notes := generatePerformanceNotes(kpi, change, progress)
		row.AddCell(notes).SetStyle(rowStyle)
	}
}

func createPerformanceTrendsSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, data DashboardData) {
	// Set column widths
	sheet.SetColumnWidth("A", 15.0)  // Month
	sheet.SetColumnWidth("B", 15.0)  // Revenue
	sheet.SetColumnWidth("C", 12.0)  // Growth %
	sheet.SetColumnWidth("D", 15.0)  // 3-Month Avg
	sheet.SetColumnWidth("E", 12.0)  // Trend

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("PERFORMANCE TRENDS ANALYSIS").SetStyle(styles["dashboard_title"])

	sheet.AddRow() // Empty row
	sheet.AddRow() // Empty row

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Period").SetStyle(styles["table_header"])
	headerRow.AddCell("Revenue").SetStyle(styles["table_header"])
	headerRow.AddCell("Growth %").SetStyle(styles["table_header"])
	headerRow.AddCell("3-Month Average").SetStyle(styles["table_header"])
	headerRow.AddCell("Trend").SetStyle(styles["table_header"])

	// Trend data
	for i, point := range data.TrendData {
		row := sheet.AddRow()
		
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell(point.Label).SetStyle(rowStyle)
		row.AddCell(point.Value).SetStyle(styles["currency"])
		
		// Calculate growth rate
		if i > 0 {
			growth := ((point.Value - data.TrendData[i-1].Value) / data.TrendData[i-1].Value) * 100
			growthCell := row.AddCell(growth / 100)
			growthCell.SetStyle(styles["percentage"])
		} else {
			row.AddCell("-").SetStyle(rowStyle)
		}
		
		// Calculate 3-month moving average
		if i >= 2 {
			avg := (point.Value + data.TrendData[i-1].Value + data.TrendData[i-2].Value) / 3
			row.AddCell(avg).SetStyle(styles["currency"])
		} else {
			row.AddCell("-").SetStyle(rowStyle)
		}
		
		// Trend indicator
		if i > 0 {
			if point.Value > data.TrendData[i-1].Value {
				row.AddCell("↗ Increasing").SetStyle(styles["kpi_positive"])
			} else if point.Value < data.TrendData[i-1].Value {
				row.AddCell("↘ Decreasing").SetStyle(styles["kpi_negative"])
			} else {
				row.AddCell("→ Stable").SetStyle(rowStyle)
			}
		} else {
			row.AddCell("-").SetStyle(rowStyle)
		}
	}
}

func createAlertsSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, data DashboardData) {
	// Set column widths
	sheet.SetColumnWidth("A", 12.0)  // Priority
	sheet.SetColumnWidth("B", 25.0)  // Title
	sheet.SetColumnWidth("C", 40.0)  // Message
	sheet.SetColumnWidth("D", 15.0)  // Value
	sheet.SetColumnWidth("E", 20.0)  // Recommended Action

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("BUSINESS ALERTS & ACTIONS").SetStyle(styles["dashboard_title"])

	sheet.AddRow() // Empty row
	sheet.AddRow() // Empty row

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Priority").SetStyle(styles["table_header"])
	headerRow.AddCell("Alert Title").SetStyle(styles["table_header"])
	headerRow.AddCell("Description").SetStyle(styles["table_header"])
	headerRow.AddCell("Value").SetStyle(styles["table_header"])
	headerRow.AddCell("Recommended Action").SetStyle(styles["table_header"])

	// Alert data
	for i, alert := range data.Alerts {
		row := sheet.AddRow()
		
		// Priority with appropriate styling
		priorityCell := row.AddCell(strings.ToUpper(alert.Level))
		priorityCell.SetStyle(styles["alert_"+alert.Level])
		
		// Alert details
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell(alert.Title).SetStyle(rowStyle)
		row.AddCell(alert.Message).SetStyle(rowStyle)
		
		// Format value based on alert type
		valueCell := row.AddCell(alert.Value)
		if alert.Level == "success" || alert.Level == "info" {
			valueCell.SetStyle(styles["kpi_positive"])
		} else {
			valueCell.SetStyle(styles["kpi_negative"])
		}
		
		// Recommended actions
		action := generateRecommendedAction(alert)
		row.AddCell(action).SetStyle(rowStyle)
	}
}

func createOperationsOverviewSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, data DashboardData) {
	// This would contain operational metrics like efficiency, throughput, etc.
	// For brevity, implementing a simplified version
	
	titleRow := sheet.AddRow()
	titleRow.AddCell("OPERATIONS OVERVIEW").SetStyle(styles["dashboard_title"])

	sheet.AddRow()
	sheet.AddRow()

	// Operational metrics (simulated)
	operationalMetrics := []struct {
		name  string
		value float64
		unit  string
	}{
		{"Production Efficiency", 94.7, "%"},
		{"Quality Score", 98.2, "%"},
		{"On-Time Delivery", 96.8, "%"},
		{"Resource Utilization", 87.3, "%"},
		{"Customer Response Time", 2.4, "hours"},
		{"Process Automation", 78.5, "%"},
	}

	headerRow := sheet.AddRow()
	headerRow.AddCell("Operational Metric").SetStyle(styles["table_header"])
	headerRow.AddCell("Current Value").SetStyle(styles["table_header"])
	headerRow.AddCell("Performance Status").SetStyle(styles["table_header"])

	for i, metric := range operationalMetrics {
		row := sheet.AddRow()
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell(metric.name).SetStyle(rowStyle)
		row.AddCell(fmt.Sprintf("%.1f %s", metric.value, metric.unit)).SetStyle(styles["kpi_value"])
		
		var status string
		var statusStyle excelbuilder.StyleConfig
		if metric.value >= 95 {
			status = "Excellent"
			statusStyle = styles["alert_success"]
		} else if metric.value >= 85 {
			status = "Good"
			statusStyle = styles["alert_info"]
		} else if metric.value >= 75 {
			status = "Needs Improvement"
			statusStyle = styles["alert_warning"]
		} else {
			status = "Critical"
			statusStyle = styles["alert_danger"]
		}
		
		row.AddCell(status).SetStyle(statusStyle)
	}
}

func createProcessMetricsSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, data DashboardData) {
	// Simplified process metrics implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("PROCESS METRICS").SetStyle(styles["dashboard_title"])
	
	// Add process-specific metrics here
	// This is a placeholder for the full implementation
}

func createResourceUtilizationSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, data DashboardData) {
	// Simplified resource utilization implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("RESOURCE UTILIZATION").SetStyle(styles["dashboard_title"])
	
	// Add resource utilization metrics here
	// This is a placeholder for the full implementation
}

func createFinancialSummarySheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, data DashboardData) {
	// Simplified financial summary implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("FINANCIAL SUMMARY").SetStyle(styles["dashboard_title"])
	
	// Add financial summary metrics here
	// This is a placeholder for the full implementation
}

func createRevenueAnalysisSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, data DashboardData) {
	// Simplified revenue analysis implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("REVENUE ANALYSIS").SetStyle(styles["dashboard_title"])
	
	// Add revenue analysis here
	// This is a placeholder for the full implementation
}

func createProfitabilityAnalysisSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, data DashboardData) {
	// Simplified profitability analysis implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("PROFITABILITY ANALYSIS").SetStyle(styles["dashboard_title"])
	
	// Add profitability analysis here
	// This is a placeholder for the full implementation
}

// Helper functions

func formatKPIValue(kpi KPIMetric) string {
	switch kpi.Format {
	case "currency":
		return fmt.Sprintf("$%.0f", kpi.CurrentValue)
	case "percentage":
		return fmt.Sprintf("%.1f%%", kpi.CurrentValue*100)
	case "integer":
		return fmt.Sprintf("%.0f %s", kpi.CurrentValue, kpi.Unit)
	default:
		return fmt.Sprintf("%.1f %s", kpi.CurrentValue, kpi.Unit)
	}
}

func generatePerformanceNotes(kpi KPIMetric, change, progress float64) string {
	if progress >= 100 {
		return fmt.Sprintf("Exceeded target by %.1f%%. Strong performance.", progress-100)
	} else if progress >= 90 {
		return "On track to meet target. Continue current strategy."
	} else if progress >= 75 {
		return "Below target but recoverable. Review strategy."
	} else {
		return "Significantly below target. Immediate action required."
	}
}

func generateRecommendedAction(alert Alert) string {
	switch alert.Level {
	case "success":
		return "Continue current strategy and document best practices"
	case "info":
		return "Monitor trends and prepare strategic response"
	case "warning":
		return "Review processes and implement corrective measures"
	case "danger":
		return "Immediate intervention required - escalate to management"
	default:
		return "Review and assess situation"
	}
}