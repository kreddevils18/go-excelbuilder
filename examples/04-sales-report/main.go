package main

import (
	"fmt"
	"log"
	"math/rand"
	"os"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Business data structures
type SalesRep struct {
	ID       int
	Name     string
	Region   string
	Territory string
	HireDate time.Time
}

type Product struct {
	ID       string
	Name     string
	Category string
	Price    float64
	Cost     float64
}

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

type RegionalSummary struct {
	Region     string
	Revenue    float64
	Units      int
	Reps       int
	AvgPerRep  float64
	Margin     float64
	Growth     float64
}

func main() {
	// Create output directory
	if err := os.MkdirAll("output", 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}

	// Generate sample business data
	salesReps := generateSalesReps()
	products := generateProducts()
	salesRecords := generateSalesRecords(salesReps, products)
	monthlySummaries := generateMonthlySummaries()
	regionalSummaries := generateRegionalSummaries()

	// Create Excel builder
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Q4 2024 Sales Performance Report",
		Author:      "Sales Analytics Team",
		Subject:     "Quarterly Sales Analysis",
		Description: "Comprehensive sales performance analysis including revenue, growth, and regional breakdowns",
		Keywords:    "sales,revenue,performance,quarterly,analytics,business-intelligence",
		Company:     "TechCorp Solutions",
	})

	// Define professional styling
	styles := createBusinessStyles()

	// Sheet 1: Executive Summary
	execSheet := workbook.AddSheet("Executive Summary")
	if execSheet == nil {
		log.Fatal("Failed to create executive summary sheet")
	}
	createExecutiveSummary(execSheet, styles, monthlySummaries, regionalSummaries)

	// Sheet 2: Monthly Performance
	monthlySheet := workbook.AddSheet("Monthly Performance")
	if monthlySheet == nil {
		log.Fatal("Failed to create monthly performance sheet")
	}
	createMonthlyPerformance(monthlySheet, styles, monthlySummaries)

	// Sheet 3: Regional Analysis
	regionalSheet := workbook.AddSheet("Regional Analysis")
	if regionalSheet == nil {
		log.Fatal("Failed to create regional analysis sheet")
	}
	createRegionalAnalysis(regionalSheet, styles, regionalSummaries)

	// Sheet 4: Sales Team Performance
	teamSheet := workbook.AddSheet("Sales Team")
	if teamSheet == nil {
		log.Fatal("Failed to create sales team sheet")
	}
	createSalesTeamPerformance(teamSheet, styles, salesReps, salesRecords)

	// Sheet 5: Product Performance
	productSheet := workbook.AddSheet("Product Analysis")
	if productSheet == nil {
		log.Fatal("Failed to create product analysis sheet")
	}
	createProductAnalysis(productSheet, styles, products, salesRecords)

	// Sheet 6: Detailed Transactions
	transactionSheet := workbook.AddSheet("Transaction Details")
	if transactionSheet == nil {
		log.Fatal("Failed to create transaction details sheet")
	}
	createTransactionDetails(transactionSheet, styles, salesRecords, salesReps, products)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build workbook")
	}

	filename := "output/04-sales-report-q4-2024.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save workbook: %v", err)
	}

	fmt.Printf("✅ Sales report generated successfully!\n")
	fmt.Printf("📁 File saved as: %s\n", filename)
	fmt.Printf("📊 Report includes:\n")
	fmt.Printf("   • Executive Summary with KPIs\n")
	fmt.Printf("   • Monthly Performance Trends\n")
	fmt.Printf("   • Regional Analysis & Comparisons\n")
	fmt.Printf("   • Sales Team Performance Metrics\n")
	fmt.Printf("   • Product Analysis & Rankings\n")
	fmt.Printf("   • Detailed Transaction Records\n")
	fmt.Printf("\n🎯 Next steps: Try examples/05-import-export/ for data integration\n")
}

func createBusinessStyles() map[string]excelbuilder.StyleConfig {
	styles := make(map[string]excelbuilder.StyleConfig)

	// Corporate colors
	colors := map[string]string{
		"corporate_blue":   "#1F4E79",
		"corporate_light":  "#5B9BD5",
		"success_green":    "#70AD47",
		"warning_orange":   "#FFC000",
		"danger_red":       "#C5504B",
		"neutral_gray":     "#7F7F7F",
		"light_gray":       "#F2F2F2",
		"white":            "#FFFFFF",
		"black":            "#000000",
	}

	// Report title style
	styles["report_title"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   18,
			Color:  colors["corporate_blue"],
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
			Color:  colors["corporate_blue"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: excelbuilder.BorderConfig{
			Bottom: excelbuilder.BorderSide{Style: "medium", Color: colors["corporate_blue"]},
		},
	}

	// Table headers
	styles["table_header"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   11,
			Color:  colors["white"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["corporate_blue"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createFullBorder("thin", colors["white"]),
	}

	// KPI styles
	styles["kpi_positive"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
			Color:  colors["white"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["success_green"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		NumberFormat: "$#,##0",
		Border:      createFullBorder("thin", colors["black"]),
	}

	styles["kpi_negative"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
			Color:  colors["white"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["danger_red"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		NumberFormat: "$#,##0",
		Border:      createFullBorder("thin", colors["black"]),
	}

	styles["kpi_neutral"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
			Color:  colors["black"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["warning_orange"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		NumberFormat: "$#,##0",
		Border:      createFullBorder("thin", colors["black"]),
	}

	// Data styles
	styles["data_normal"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createFullBorder("thin", colors["light_gray"]),
	}

	styles["data_alternate"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["light_gray"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createFullBorder("thin", colors["neutral_gray"]),
	}

	// Number formats
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
		Border:      createFullBorder("thin", colors["light_gray"]),
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
		NumberFormat: "0.0%",
		Border:      createFullBorder("thin", colors["light_gray"]),
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
		Border:      createFullBorder("thin", colors["light_gray"]),
	}

	styles["date"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		NumberFormat: "mm/dd/yyyy",
		Border:      createFullBorder("thin", colors["light_gray"]),
	}

	// Performance indicators
	styles["performance_excellent"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   10,
			Color:  colors["white"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["success_green"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createFullBorder("thin", colors["black"]),
	}

	styles["performance_good"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   10,
			Color:  colors["black"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["warning_orange"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createFullBorder("thin", colors["black"]),
	}

	styles["performance_poor"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   10,
			Color:  colors["white"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["danger_red"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createFullBorder("thin", colors["black"]),
	}

	return styles
}

func createFullBorder(style, color string) excelbuilder.BorderConfig {
	return excelbuilder.BorderConfig{
		Top:    excelbuilder.BorderSide{Style: style, Color: color},
		Bottom: excelbuilder.BorderSide{Style: style, Color: color},
		Left:   excelbuilder.BorderSide{Style: style, Color: color},
		Right:  excelbuilder.BorderSide{Style: style, Color: color},
	}
}

func createExecutiveSummary(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, monthly []MonthlySummary, regional []RegionalSummary) {
	// Set column widths
	sheet.SetColumnWidth("A", 25.0)
	sheet.SetColumnWidth("B", 20.0)
	sheet.SetColumnWidth("C", 20.0)
	sheet.SetColumnWidth("D", 20.0)
	sheet.SetColumnWidth("E", 20.0)

	// Report header
	headerRow := sheet.AddRow()
	headerRow.AddCell("Q4 2024 SALES PERFORMANCE REPORT").SetStyle(styles["report_title"])

	// Report metadata
	metaRow := sheet.AddRow()
	metaRow.AddCell(fmt.Sprintf("Generated: %s | TechCorp Solutions", time.Now().Format("January 2, 2006"))).SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Italic: true,
			Size:   10,
			Color:  "#7F7F7F",
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	})

	// Empty rows for spacing
	sheet.AddRow()
	sheet.AddRow()

	// Key Performance Indicators section
	kpiSectionRow := sheet.AddRow()
	kpiSectionRow.AddCell("KEY PERFORMANCE INDICATORS").SetStyle(styles["section_header"])

	sheet.AddRow() // Empty row

	// Calculate KPIs
	totalRevenue := 0.0
	totalUnits := 0
	for _, m := range monthly {
		totalRevenue += m.Revenue
		totalUnits += m.Units
	}

	avgGrowth := 0.0
	for _, m := range monthly {
		avgGrowth += m.Growth
	}
	avgGrowth /= float64(len(monthly))

	// KPI headers
	kpiHeaderRow := sheet.AddRow()
	kpiHeaderRow.AddCell("Metric").SetStyle(styles["table_header"])
	kpiHeaderRow.AddCell("Q4 2024").SetStyle(styles["table_header"])
	kpiHeaderRow.AddCell("Target").SetStyle(styles["table_header"])
	kpiHeaderRow.AddCell("Variance").SetStyle(styles["table_header"])
	kpiHeaderRow.AddCell("Status").SetStyle(styles["table_header"])

	// KPI data
	kpis := []struct {
		metric   string
		actual   float64
		target   float64
		format   string
	}{
		{"Total Revenue", totalRevenue, 2500000, "currency"},
		{"Units Sold", float64(totalUnits), 15000, "integer"},
		{"Average Growth", avgGrowth, 0.15, "percentage"},
		{"Revenue per Unit", totalRevenue / float64(totalUnits), 180, "currency"},
	}

	for _, kpi := range kpis {
		row := sheet.AddRow()
		row.AddCell(kpi.metric).SetStyle(styles["data_normal"])
		
		// Format actual value
		if kpi.format == "currency" {
			row.AddCell(kpi.actual).SetStyle(styles["currency"])
			row.AddCell(kpi.target).SetStyle(styles["currency"])
		} else if kpi.format == "percentage" {
			row.AddCell(kpi.actual).SetStyle(styles["percentage"])
			row.AddCell(kpi.target).SetStyle(styles["percentage"])
		} else {
			row.AddCell(kpi.actual).SetStyle(styles["integer"])
			row.AddCell(kpi.target).SetStyle(styles["integer"])
		}
		
		// Calculate variance
		variance := (kpi.actual - kpi.target) / kpi.target
		row.AddCell(variance).SetStyle(styles["percentage"])
		
		// Status based on performance
		status := ""
		statusStyle := styles["performance_poor"]
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
		row.AddCell(status).SetStyle(statusStyle)
	}

	// Empty rows
	sheet.AddRow()
	sheet.AddRow()

	// Regional Performance Summary
	regionalSectionRow := sheet.AddRow()
	regionalSectionRow.AddCell("REGIONAL PERFORMANCE SUMMARY").SetStyle(styles["section_header"])

	sheet.AddRow() // Empty row

	// Regional headers
	regionalHeaderRow := sheet.AddRow()
	regionalHeaderRow.AddCell("Region").SetStyle(styles["table_header"])
	regionalHeaderRow.AddCell("Revenue").SetStyle(styles["table_header"])
	regionalHeaderRow.AddCell("Growth").SetStyle(styles["table_header"])
	regionalHeaderRow.AddCell("Margin").SetStyle(styles["table_header"])
	regionalHeaderRow.AddCell("Performance").SetStyle(styles["table_header"])

	// Regional data
	for i, region := range regional {
		row := sheet.AddRow()
		
		// Alternating row colors
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell(region.Region).SetStyle(rowStyle)
		row.AddCell(region.Revenue).SetStyle(styles["currency"])
		row.AddCell(region.Growth).SetStyle(styles["percentage"])
		row.AddCell(region.Margin).SetStyle(styles["percentage"])
		
		// Performance rating
		performance := ""
		performanceStyle := styles["performance_poor"]
		if region.Growth >= 0.2 {
			performance = "Outstanding"
			performanceStyle = styles["performance_excellent"]
		} else if region.Growth >= 0.1 {
			performance = "Strong"
			performanceStyle = styles["performance_good"]
		} else {
			performance = "Needs Focus"
			performanceStyle = styles["performance_poor"]
		}
		row.AddCell(performance).SetStyle(performanceStyle)
	}

	// Key insights
	sheet.AddRow()
	sheet.AddRow()
	insightsRow := sheet.AddRow()
	insightsRow.AddCell("KEY INSIGHTS & RECOMMENDATIONS").SetStyle(styles["section_header"])

	insights := []string{
		"• Q4 revenue exceeded target by 8.2%, driven by strong performance in North America",
		"• Unit sales growth of 12% indicates healthy market expansion",
		"• Asia Pacific region shows highest growth potential with 24% increase",
		"• Recommend increased investment in high-performing regions for Q1 2025",
		"• Product mix optimization could improve overall margins by 2-3%",
	}

	for _, insight := range insights {
		insightRow := sheet.AddRow()
		insightRow.AddCell(insight).SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Size:   10,
				Family: "Calibri",
			},
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: "left",
				Vertical:   "middle",
			},
		})
	}
}

func createMonthlyPerformance(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, monthly []MonthlySummary) {
	// Set column widths
	sheet.SetColumnWidth("A", 15.0)
	sheet.SetColumnWidth("B", 18.0)
	sheet.SetColumnWidth("C", 12.0)
	sheet.SetColumnWidth("D", 15.0)
	sheet.SetColumnWidth("E", 12.0)
	sheet.SetColumnWidth("F", 18.0)
	sheet.SetColumnWidth("G", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("MONTHLY PERFORMANCE ANALYSIS").SetStyle(styles["report_title"])

	sheet.AddRow()
	sheet.AddRow()

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Month").SetStyle(styles["table_header"])
	headerRow.AddCell("Revenue").SetStyle(styles["table_header"])
	headerRow.AddCell("Units").SetStyle(styles["table_header"])
	headerRow.AddCell("Avg Deal Size").SetStyle(styles["table_header"])
	headerRow.AddCell("Growth %").SetStyle(styles["table_header"])
	headerRow.AddCell("Target").SetStyle(styles["table_header"])
	headerRow.AddCell("Attainment %").SetStyle(styles["table_header"])

	// Monthly data
	for i, month := range monthly {
		row := sheet.AddRow()
		
		// Alternating row colors
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell(month.Month).SetStyle(rowStyle)
		row.AddCell(month.Revenue).SetStyle(styles["currency"])
		row.AddCell(month.Units).SetStyle(styles["integer"])
		row.AddCell(month.AvgDeal).SetStyle(styles["currency"])
		row.AddCell(month.Growth).SetStyle(styles["percentage"])
		row.AddCell(month.Target).SetStyle(styles["currency"])
		row.AddCell(month.Attainment).SetStyle(getAttainmentStyle(month.Attainment, styles))
	}

	// Summary statistics
	sheet.AddRow()
	sheet.AddRow()
	summaryHeaderRow := sheet.AddRow()
	summaryHeaderRow.AddCell("QUARTERLY SUMMARY").SetStyle(styles["section_header"])

	sheet.AddRow()

	// Calculate totals
	totalRevenue := 0.0
	totalUnits := 0
	totalTarget := 0.0
	for _, m := range monthly {
		totalRevenue += m.Revenue
		totalUnits += m.Units
		totalTarget += m.Target
	}

	summaryRow := sheet.AddRow()
	summaryRow.AddCell("TOTAL").SetStyle(styles["table_header"])
	summaryRow.AddCell(totalRevenue).SetStyle(styles["kpi_positive"])
	summaryRow.AddCell(totalUnits).SetStyle(styles["kpi_positive"])
	summaryRow.AddCell(totalRevenue / float64(totalUnits)).SetStyle(styles["kpi_positive"])
	summaryRow.AddCell("").SetStyle(styles["data_normal"])
	summaryRow.AddCell(totalTarget).SetStyle(styles["kpi_neutral"])
	summaryRow.AddCell(totalRevenue / totalTarget).SetStyle(getAttainmentStyle(totalRevenue/totalTarget, styles))
}

func getAttainmentStyle(attainment float64, styles map[string]excelbuilder.StyleConfig) excelbuilder.StyleConfig {
	baseStyle := styles["percentage"]
	if attainment >= 1.1 {
		baseStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#70AD47"}
		baseStyle.Font.Color = "#FFFFFF"
		baseStyle.Font.Bold = true
	} else if attainment >= 1.0 {
		baseStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#FFC000"}
		baseStyle.Font.Color = "#000000"
	} else {
		baseStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#C5504B"}
		baseStyle.Font.Color = "#FFFFFF"
		baseStyle.Font.Bold = true
	}
	return baseStyle
}

func createRegionalAnalysis(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, regional []RegionalSummary) {
	// Set column widths
	sheet.SetColumnWidth("A", 20.0)
	sheet.SetColumnWidth("B", 18.0)
	sheet.SetColumnWidth("C", 12.0)
	sheet.SetColumnWidth("D", 12.0)
	sheet.SetColumnWidth("E", 18.0)
	sheet.SetColumnWidth("F", 15.0)
	sheet.SetColumnWidth("G", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("REGIONAL PERFORMANCE ANALYSIS").SetStyle(styles["report_title"])

	sheet.AddRow()
	sheet.AddRow()

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Region").SetStyle(styles["table_header"])
	headerRow.AddCell("Revenue").SetStyle(styles["table_header"])
	headerRow.AddCell("Units").SetStyle(styles["table_header"])
	headerRow.AddCell("Sales Reps").SetStyle(styles["table_header"])
	headerRow.AddCell("Avg per Rep").SetStyle(styles["table_header"])
	headerRow.AddCell("Margin %").SetStyle(styles["table_header"])
	headerRow.AddCell("Growth %").SetStyle(styles["table_header"])

	// Regional data
	for i, region := range regional {
		row := sheet.AddRow()
		
		// Alternating row colors
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell(region.Region).SetStyle(rowStyle)
		row.AddCell(region.Revenue).SetStyle(styles["currency"])
		row.AddCell(region.Units).SetStyle(styles["integer"])
		row.AddCell(region.Reps).SetStyle(styles["integer"])
		row.AddCell(region.AvgPerRep).SetStyle(styles["currency"])
		row.AddCell(region.Margin).SetStyle(styles["percentage"])
		row.AddCell(region.Growth).SetStyle(getGrowthStyle(region.Growth, styles))
	}
}

func getGrowthStyle(growth float64, styles map[string]excelbuilder.StyleConfig) excelbuilder.StyleConfig {
	baseStyle := styles["percentage"]
	if growth >= 0.2 {
		baseStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#70AD47"}
		baseStyle.Font.Color = "#FFFFFF"
		baseStyle.Font.Bold = true
	} else if growth >= 0.1 {
		baseStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#FFC000"}
		baseStyle.Font.Color = "#000000"
	} else if growth >= 0 {
		baseStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#F2F2F2"}
		baseStyle.Font.Color = "#000000"
	} else {
		baseStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#C5504B"}
		baseStyle.Font.Color = "#FFFFFF"
		baseStyle.Font.Bold = true
	}
	return baseStyle
}

func createSalesTeamPerformance(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, reps []SalesRep, records []SalesRecord) {
	// Calculate rep performance
	repPerformance := make(map[int]struct {
		revenue float64
		units   int
		deals   int
	})

	for _, record := range records {
		perf := repPerformance[record.SalesRepID]
		perf.revenue += float64(record.Quantity) * 150.0 // Simplified calculation
		perf.units += record.Quantity
		perf.deals++
		repPerformance[record.SalesRepID] = perf
	}

	// Set column widths
	sheet.SetColumnWidth("A", 20.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 18.0)
	sheet.SetColumnWidth("E", 12.0)
	sheet.SetColumnWidth("F", 12.0)
	sheet.SetColumnWidth("G", 18.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("SALES TEAM PERFORMANCE").SetStyle(styles["report_title"])

	sheet.AddRow()
	sheet.AddRow()

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Sales Rep").SetStyle(styles["table_header"])
	headerRow.AddCell("Region").SetStyle(styles["table_header"])
	headerRow.AddCell("Territory").SetStyle(styles["table_header"])
	headerRow.AddCell("Revenue").SetStyle(styles["table_header"])
	headerRow.AddCell("Units").SetStyle(styles["table_header"])
	headerRow.AddCell("Deals").SetStyle(styles["table_header"])
	headerRow.AddCell("Avg Deal Size").SetStyle(styles["table_header"])

	// Rep data
	for i, rep := range reps {
		row := sheet.AddRow()
		perf := repPerformance[rep.ID]
		
		// Alternating row colors
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell(rep.Name).SetStyle(rowStyle)
		row.AddCell(rep.Region).SetStyle(rowStyle)
		row.AddCell(rep.Territory).SetStyle(rowStyle)
		row.AddCell(perf.revenue).SetStyle(styles["currency"])
		row.AddCell(perf.units).SetStyle(styles["integer"])
		row.AddCell(perf.deals).SetStyle(styles["integer"])
		
		avgDeal := 0.0
		if perf.deals > 0 {
			avgDeal = perf.revenue / float64(perf.deals)
		}
		row.AddCell(avgDeal).SetStyle(styles["currency"])
	}
}

func createProductAnalysis(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, products []Product, records []SalesRecord) {
	// Calculate product performance
	productPerformance := make(map[string]struct {
		units   int
		revenue float64
		margin  float64
	})

	for _, record := range records {
		for _, product := range products {
			if product.ID == record.ProductID {
				perf := productPerformance[product.ID]
				revenue := float64(record.Quantity) * product.Price
				cost := float64(record.Quantity) * product.Cost
				perf.units += record.Quantity
				perf.revenue += revenue
				perf.margin += revenue - cost
				productPerformance[product.ID] = perf
				break
			}
		}
	}

	// Set column widths
	sheet.SetColumnWidth("A", 20.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 12.0)
	sheet.SetColumnWidth("E", 18.0)
	sheet.SetColumnWidth("F", 18.0)
	sheet.SetColumnWidth("G", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("PRODUCT PERFORMANCE ANALYSIS").SetStyle(styles["report_title"])

	sheet.AddRow()
	sheet.AddRow()

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Product").SetStyle(styles["table_header"])
	headerRow.AddCell("Category").SetStyle(styles["table_header"])
	headerRow.AddCell("Price").SetStyle(styles["table_header"])
	headerRow.AddCell("Units Sold").SetStyle(styles["table_header"])
	headerRow.AddCell("Revenue").SetStyle(styles["table_header"])
	headerRow.AddCell("Gross Margin").SetStyle(styles["table_header"])
	headerRow.AddCell("Margin %").SetStyle(styles["table_header"])

	// Product data
	for i, product := range products {
		row := sheet.AddRow()
		perf := productPerformance[product.ID]
		
		// Alternating row colors
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell(product.Name).SetStyle(rowStyle)
		row.AddCell(product.Category).SetStyle(rowStyle)
		row.AddCell(product.Price).SetStyle(styles["currency"])
		row.AddCell(perf.units).SetStyle(styles["integer"])
		row.AddCell(perf.revenue).SetStyle(styles["currency"])
		row.AddCell(perf.margin).SetStyle(styles["currency"])
		
		marginPercent := 0.0
		if perf.revenue > 0 {
			marginPercent = perf.margin / perf.revenue
		}
		row.AddCell(marginPercent).SetStyle(styles["percentage"])
	}
}

func createTransactionDetails(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, records []SalesRecord, reps []SalesRep, products []Product) {
	// Set column widths
	sheet.SetColumnWidth("A", 12.0)
	sheet.SetColumnWidth("B", 18.0)
	sheet.SetColumnWidth("C", 18.0)
	sheet.SetColumnWidth("D", 12.0)
	sheet.SetColumnWidth("E", 15.0)
	sheet.SetColumnWidth("F", 12.0)
	sheet.SetColumnWidth("G", 20.0)
	sheet.SetColumnWidth("H", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("TRANSACTION DETAILS").SetStyle(styles["report_title"])

	sheet.AddRow()
	sheet.AddRow()

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Trans ID").SetStyle(styles["table_header"])
	headerRow.AddCell("Sales Rep").SetStyle(styles["table_header"])
	headerRow.AddCell("Product").SetStyle(styles["table_header"])
	headerRow.AddCell("Quantity").SetStyle(styles["table_header"])
	headerRow.AddCell("Sale Date").SetStyle(styles["table_header"])
	headerRow.AddCell("Discount").SetStyle(styles["table_header"])
	headerRow.AddCell("Customer").SetStyle(styles["table_header"])
	headerRow.AddCell("Region").SetStyle(styles["table_header"])

	// Create lookup maps
	repMap := make(map[int]string)
	for _, rep := range reps {
		repMap[rep.ID] = rep.Name
	}

	productMap := make(map[string]string)
	for _, product := range products {
		productMap[product.ID] = product.Name
	}

	// Transaction data (limit to first 50 for readability)
	limit := len(records)
	if limit > 50 {
		limit = 50
	}

	for i := 0; i < limit; i++ {
		record := records[i]
		row := sheet.AddRow()
		
		// Alternating row colors
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell(record.ID).SetStyle(styles["integer"])
		row.AddCell(repMap[record.SalesRepID]).SetStyle(rowStyle)
		row.AddCell(productMap[record.ProductID]).SetStyle(rowStyle)
		row.AddCell(record.Quantity).SetStyle(styles["integer"])
		row.AddCell(record.SaleDate).SetStyle(styles["date"])
		row.AddCell(record.Discount).SetStyle(styles["percentage"])
		row.AddCell(record.Customer).SetStyle(rowStyle)
		row.AddCell(record.Region).SetStyle(rowStyle)
	}

	// Add note about data limitation
	if len(records) > 50 {
		sheet.AddRow()
		noteRow := sheet.AddRow()
		noteRow.AddCell(fmt.Sprintf("Note: Showing first 50 of %d total transactions", len(records))).SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Italic: true,
				Size:   9,
				Color:  "#7F7F7F",
				Family: "Calibri",
			},
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: "left",
				Vertical:   "middle",
			},
		})
	}
}

// Data generation functions
func generateSalesReps() []SalesRep {
	rand.Seed(time.Now().UnixNano())
	reps := []SalesRep{
		{1, "Alice Johnson", "North America", "West Coast", time.Date(2020, 3, 15, 0, 0, 0, 0, time.UTC)},
		{2, "Bob Smith", "North America", "East Coast", time.Date(2019, 7, 22, 0, 0, 0, 0, time.UTC)},
		{3, "Carol Davis", "Europe", "UK & Ireland", time.Date(2021, 1, 10, 0, 0, 0, 0, time.UTC)},
		{4, "David Wilson", "Europe", "DACH", time.Date(2018, 11, 5, 0, 0, 0, 0, time.UTC)},
		{5, "Eva Brown", "Asia Pacific", "Australia", time.Date(2022, 4, 18, 0, 0, 0, 0, time.UTC)},
		{6, "Frank Miller", "Asia Pacific", "Japan", time.Date(2020, 9, 12, 0, 0, 0, 0, time.UTC)},
		{7, "Grace Lee", "Latin America", "Brazil", time.Date(2021, 6, 8, 0, 0, 0, 0, time.UTC)},
		{8, "Henry Taylor", "Latin America", "Mexico", time.Date(2019, 12, 3, 0, 0, 0, 0, time.UTC)},
	}
	return reps
}

func generateProducts() []Product {
	products := []Product{
		{"PROD001", "TechCorp Pro Suite", "Software", 299.99, 120.00},
		{"PROD002", "TechCorp Analytics", "Software", 199.99, 80.00},
		{"PROD003", "TechCorp Mobile", "Software", 99.99, 40.00},
		{"PROD004", "TechCorp Enterprise", "Software", 599.99, 240.00},
		{"PROD005", "TechCorp Starter", "Software", 49.99, 20.00},
		{"PROD006", "Support Package Basic", "Service", 149.99, 60.00},
		{"PROD007", "Support Package Premium", "Service", 299.99, 120.00},
		{"PROD008", "Training Workshop", "Service", 499.99, 200.00},
	}
	return products
}

func generateSalesRecords(reps []SalesRep, products []Product) []SalesRecord {
	rand.Seed(time.Now().UnixNano())
	var records []SalesRecord
	customers := []string{
		"Acme Corp", "Global Industries", "Tech Solutions Inc", "Innovation Labs",
		"Future Systems", "Digital Dynamics", "Smart Enterprises", "NextGen Solutions",
		"Advanced Analytics", "Cloud Computing Co", "Data Insights Ltd", "AI Innovations",
	}

	for i := 1; i <= 200; i++ {
		rep := reps[rand.Intn(len(reps))]
		product := products[rand.Intn(len(products))]
		
		// Generate sale date in Q4 2024
		startDate := time.Date(2024, 10, 1, 0, 0, 0, 0, time.UTC)
		endDate := time.Date(2024, 12, 31, 0, 0, 0, 0, time.UTC)
		daysDiff := int(endDate.Sub(startDate).Hours() / 24)
		saleDate := startDate.AddDate(0, 0, rand.Intn(daysDiff))
		
		record := SalesRecord{
			ID:         i,
			SalesRepID: rep.ID,
			ProductID:  product.ID,
			Quantity:   rand.Intn(10) + 1,
			SaleDate:   saleDate,
			Discount:   float64(rand.Intn(20)) / 100.0, // 0-20% discount
			Customer:   customers[rand.Intn(len(customers))],
			Region:     rep.Region,
		}
		records = append(records, record)
	}
	return records
}

func generateMonthlySummaries() []MonthlySummary {
	return []MonthlySummary{
		{"October 2024", 625000, 3800, 164.47, 0.12, 600000, 1.042},
		{"November 2024", 680000, 4100, 165.85, 0.15, 650000, 1.046},
		{"December 2024", 725000, 4300, 168.60, 0.18, 700000, 1.036},
	}
}

func generateRegionalSummaries() []RegionalSummary {
	return []RegionalSummary{
		{"North America", 850000, 5200, 3, 283333, 0.42, 0.16},
		{"Europe", 650000, 3900, 2, 325000, 0.38, 0.14},
		{"Asia Pacific", 420000, 2500, 2, 210000, 0.45, 0.24},
		{"Latin America", 110000, 600, 1, 110000, 0.35, 0.08},
	}
}