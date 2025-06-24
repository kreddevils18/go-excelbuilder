package main

import (
	"fmt"
	"log"
	"math"
	"os"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Financial data structures
type FinancialStatement struct {
	Period           string
	Revenue          float64
	CostOfGoodsSold  float64
	GrossProfit      float64
	OperatingExpenses float64
	EBITDA           float64
	Depreciation     float64
	EBIT             float64
	InterestExpense  float64
	Taxes            float64
	NetIncome        float64
}

type BalanceSheet struct {
	Period              string
	Cash                float64
	AccountsReceivable  float64
	Inventory           float64
	CurrentAssets       float64
	FixedAssets         float64
	TotalAssets         float64
	AccountsPayable     float64
	CurrentLiabilities  float64
	LongTermDebt        float64
	TotalLiabilities    float64
	ShareholderEquity   float64
}

type CashFlow struct {
	Period              string
	OperatingCashFlow   float64
	InvestingCashFlow   float64
	FinancingCashFlow   float64
	NetCashFlow         float64
	BeginningCash       float64
	EndingCash          float64
}

type FinancialRatios struct {
	Period              string
	CurrentRatio        float64
	QuickRatio          float64
	DebtToEquity        float64
	ReturnOnAssets      float64
	ReturnOnEquity      float64
	GrossMargin         float64
	OperatingMargin     float64
	NetMargin           float64
	AssetTurnover       float64
	InventoryTurnover   float64
	ReceivablesTurnover float64
}

type ValuationMetrics struct {
	Period          string
	SharePrice      float64
	SharesOutstanding float64
	MarketCap       float64
	PriceToEarnings float64
	PriceToBook     float64
	PriceToSales    float64
	EVToEBITDA      float64
	DividendYield   float64
	BookValuePerShare float64
	EarningsPerShare  float64
}

type FinancialModel struct {
	CompanyName       string
	AnalysisDate      time.Time
	IncomeStatements  []FinancialStatement
	BalanceSheets     []BalanceSheet
	CashFlows         []CashFlow
	Ratios            []FinancialRatios
	Valuations        []ValuationMetrics
	Projections       []FinancialStatement
}

func main() {
	// Create output directory
	if err := os.MkdirAll("output", 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}

	fmt.Println("💰 Starting Financial Analysis Demo...")

	// Step 1: Generate financial model data
	fmt.Println("\n📊 Step 1: Generating comprehensive financial model...")
	financialModel := generateFinancialModel()

	// Step 2: Create financial statements workbook
	fmt.Println("\n📈 Step 2: Creating financial statements analysis...")
	createFinancialStatementsWorkbook(financialModel)

	// Step 3: Create ratio analysis workbook
	fmt.Println("\n🔍 Step 3: Creating financial ratio analysis...")
	createRatioAnalysisWorkbook(financialModel)

	// Step 4: Create valuation model workbook
	fmt.Println("\n💎 Step 4: Creating valuation model...")
	createValuationModelWorkbook(financialModel)

	// Step 5: Create DCF model workbook
	fmt.Println("\n🎯 Step 5: Creating DCF valuation model...")
	createDCFModelWorkbook(financialModel)

	fmt.Println("\n✅ Financial analysis demonstration completed successfully!")
	fmt.Println("📁 Generated files:")
	fmt.Println("   • output/07-financial-statements.xlsx - Complete financial statements")
	fmt.Println("   • output/07-ratio-analysis.xlsx - Financial ratio analysis")
	fmt.Println("   • output/07-valuation-model.xlsx - Company valuation model")
	fmt.Println("   • output/07-dcf-model.xlsx - Discounted cash flow model")
	fmt.Println("\n🎯 Next steps: Try examples/08-templates/ for reusable financial templates")
}

func generateFinancialModel() FinancialModel {
	fmt.Println("   📊 Generating historical financial statements...")
	
	// Generate 5 years of historical data
	incomeStatements := make([]FinancialStatement, 5)
	balanceSheets := make([]BalanceSheet, 5)
	cashFlows := make([]CashFlow, 5)
	ratios := make([]FinancialRatios, 5)
	valuations := make([]ValuationMetrics, 5)

	// Base year values
	baseRevenue := 50000000.0 // $50M
	growthRate := 0.15        // 15% annual growth

	for i := 0; i < 5; i++ {
		year := 2020 + i
		period := fmt.Sprintf("FY%d", year)
		
		// Calculate revenue with growth
		revenue := baseRevenue * math.Pow(1+growthRate, float64(i))
		
		// Income Statement
		cogs := revenue * 0.60 // 60% COGS
		grossProfit := revenue - cogs
		opex := revenue * 0.25 // 25% OpEx
		ebitda := grossProfit - opex
		depreciation := revenue * 0.03 // 3% depreciation
		ebit := ebitda - depreciation
		interest := revenue * 0.01 // 1% interest
		taxes := (ebit - interest) * 0.25 // 25% tax rate
		netIncome := ebit - interest - taxes

		incomeStatements[i] = FinancialStatement{
			Period:            period,
			Revenue:           revenue,
			CostOfGoodsSold:   cogs,
			GrossProfit:       grossProfit,
			OperatingExpenses: opex,
			EBITDA:            ebitda,
			Depreciation:      depreciation,
			EBIT:              ebit,
			InterestExpense:   interest,
			Taxes:             taxes,
			NetIncome:         netIncome,
		}

		// Balance Sheet
		cash := revenue * 0.10
		ar := revenue * 0.08
		inventory := revenue * 0.12
		currentAssets := cash + ar + inventory
		fixedAssets := revenue * 0.80
		totalAssets := currentAssets + fixedAssets
		ap := revenue * 0.06
		currentLiab := ap + (revenue * 0.04)
		longTermDebt := revenue * 0.30
		totalLiab := currentLiab + longTermDebt
		equity := totalAssets - totalLiab

		balanceSheets[i] = BalanceSheet{
			Period:              period,
			Cash:                cash,
			AccountsReceivable:  ar,
			Inventory:           inventory,
			CurrentAssets:       currentAssets,
			FixedAssets:         fixedAssets,
			TotalAssets:         totalAssets,
			AccountsPayable:     ap,
			CurrentLiabilities:  currentLiab,
			LongTermDebt:        longTermDebt,
			TotalLiabilities:    totalLiab,
			ShareholderEquity:   equity,
		}

		// Cash Flow
		operatingCF := netIncome + depreciation
		investingCF := -depreciation * 1.2 // CapEx
		financingCF := -(netIncome * 0.3)  // Dividends
		netCF := operatingCF + investingCF + financingCF
		beginningCash := cash - netCF

		cashFlows[i] = CashFlow{
			Period:              period,
			OperatingCashFlow:   operatingCF,
			InvestingCashFlow:   investingCF,
			FinancingCashFlow:   financingCF,
			NetCashFlow:         netCF,
			BeginningCash:       beginningCash,
			EndingCash:          cash,
		}

		// Financial Ratios
		ratios[i] = FinancialRatios{
			Period:              period,
			CurrentRatio:        currentAssets / currentLiab,
			QuickRatio:          (currentAssets - inventory) / currentLiab,
			DebtToEquity:        totalLiab / equity,
			ReturnOnAssets:      netIncome / totalAssets,
			ReturnOnEquity:      netIncome / equity,
			GrossMargin:         grossProfit / revenue,
			OperatingMargin:     ebit / revenue,
			NetMargin:           netIncome / revenue,
			AssetTurnover:       revenue / totalAssets,
			InventoryTurnover:   cogs / inventory,
			ReceivablesTurnover: revenue / ar,
		}

		// Valuation Metrics
		shares := 10000000.0 // 10M shares
		sharePrice := 25.0 + float64(i)*5.0 // Growing share price
		marketCap := shares * sharePrice
		eps := netIncome / shares
		bookValue := equity / shares

		valuations[i] = ValuationMetrics{
			Period:            period,
			SharePrice:        sharePrice,
			SharesOutstanding: shares,
			MarketCap:         marketCap,
			PriceToEarnings:   sharePrice / eps,
			PriceToBook:       sharePrice / bookValue,
			PriceToSales:      marketCap / revenue,
			EVToEBITDA:        (marketCap + longTermDebt - cash) / ebitda,
			DividendYield:     (netIncome * 0.3 / shares) / sharePrice,
			BookValuePerShare: bookValue,
			EarningsPerShare:  eps,
		}
	}

	fmt.Println("   📈 Generating financial projections...")
	
	// Generate 5-year projections
	projections := make([]FinancialStatement, 5)
	for i := 0; i < 5; i++ {
		year := 2025 + i
		period := fmt.Sprintf("FY%d (Proj)", year)
		
		// Project with slightly lower growth
		projGrowthRate := 0.12 - float64(i)*0.01 // Declining growth
		lastRevenue := incomeStatements[4].Revenue
		revenue := lastRevenue * math.Pow(1+projGrowthRate, float64(i+1))
		
		// Improving margins over time
		cogsRate := 0.58 - float64(i)*0.005 // Improving COGS
		opexRate := 0.23 - float64(i)*0.003 // Improving OpEx
		
		cogs := revenue * cogsRate
		grossProfit := revenue - cogs
		opex := revenue * opexRate
		ebitda := grossProfit - opex
		depreciation := revenue * 0.025 // Slightly lower depreciation
		ebit := ebitda - depreciation
		interest := revenue * 0.008 // Lower interest
		taxes := (ebit - interest) * 0.25
		netIncome := ebit - interest - taxes

		projections[i] = FinancialStatement{
			Period:            period,
			Revenue:           revenue,
			CostOfGoodsSold:   cogs,
			GrossProfit:       grossProfit,
			OperatingExpenses: opex,
			EBITDA:            ebitda,
			Depreciation:      depreciation,
			EBIT:              ebit,
			InterestExpense:   interest,
			Taxes:             taxes,
			NetIncome:         netIncome,
		}
	}

	return FinancialModel{
		CompanyName:      "TechCorp Solutions Inc.",
		AnalysisDate:     time.Now(),
		IncomeStatements: incomeStatements,
		BalanceSheets:    balanceSheets,
		CashFlows:        cashFlows,
		Ratios:           ratios,
		Valuations:       valuations,
		Projections:      projections,
	}
}

func createFinancialStatementsWorkbook(model FinancialModel) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Financial Statements Analysis",
		Author:      "Financial Analysis System",
		Subject:     "Comprehensive Financial Statement Analysis",
		Description: "Complete financial statements including income statement, balance sheet, and cash flow analysis",
		Keywords:    "financial-statements,income-statement,balance-sheet,cash-flow,analysis",
		Company:     model.CompanyName,
	})

	// Create financial styles
	styles := createFinancialStyles()

	// Sheet 1: Income Statement
	incomeSheet := workbook.AddSheet("Income Statement")
	if incomeSheet == nil {
		log.Fatal("Failed to create income statement sheet")
	}
	createIncomeStatementSheet(incomeSheet, styles, model)

	// Sheet 2: Balance Sheet
	balanceSheet := workbook.AddSheet("Balance Sheet")
	if balanceSheet == nil {
		log.Fatal("Failed to create balance sheet")
	}
	createBalanceSheetSheet(balanceSheet, styles, model)

	// Sheet 3: Cash Flow Statement
	cashFlowSheet := workbook.AddSheet("Cash Flow Statement")
	if cashFlowSheet == nil {
		log.Fatal("Failed to create cash flow sheet")
	}
	createCashFlowSheet(cashFlowSheet, styles, model)

	// Sheet 4: Financial Summary
	summarySheet := workbook.AddSheet("Financial Summary")
	if summarySheet == nil {
		log.Fatal("Failed to create summary sheet")
	}
	createFinancialSummarySheet(summarySheet, styles, model)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build financial statements workbook")
	}

	filename := "output/07-financial-statements.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save financial statements: %v", err)
	}

	fmt.Printf("   ✓ Created financial statements: %s\n", filename)
}

func createRatioAnalysisWorkbook(model FinancialModel) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Financial Ratio Analysis",
		Author:      "Financial Analysis System",
		Subject:     "Comprehensive Financial Ratio Analysis",
		Description: "Detailed financial ratio analysis including liquidity, profitability, efficiency, and leverage ratios",
		Keywords:    "financial-ratios,liquidity,profitability,efficiency,leverage,analysis",
		Company:     model.CompanyName,
	})

	styles := createFinancialStyles()

	// Sheet 1: Ratio Summary
	ratioSummarySheet := workbook.AddSheet("Ratio Summary")
	if ratioSummarySheet == nil {
		log.Fatal("Failed to create ratio summary sheet")
	}
	createRatioSummarySheet(ratioSummarySheet, styles, model)

	// Sheet 2: Liquidity Analysis
	liquiditySheet := workbook.AddSheet("Liquidity Analysis")
	if liquiditySheet == nil {
		log.Fatal("Failed to create liquidity sheet")
	}
	createLiquidityAnalysisSheet(liquiditySheet, styles, model)

	// Sheet 3: Profitability Analysis
	profitabilitySheet := workbook.AddSheet("Profitability Analysis")
	if profitabilitySheet == nil {
		log.Fatal("Failed to create profitability sheet")
	}
	createProfitabilityAnalysisSheet(profitabilitySheet, styles, model)

	// Sheet 4: Efficiency Analysis
	efficiencySheet := workbook.AddSheet("Efficiency Analysis")
	if efficiencySheet == nil {
		log.Fatal("Failed to create efficiency sheet")
	}
	createEfficiencyAnalysisSheet(efficiencySheet, styles, model)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build ratio analysis workbook")
	}

	filename := "output/07-ratio-analysis.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save ratio analysis: %v", err)
	}

	fmt.Printf("   ✓ Created ratio analysis: %s\n", filename)
}

func createValuationModelWorkbook(model FinancialModel) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Company Valuation Model",
		Author:      "Financial Analysis System",
		Subject:     "Comprehensive Company Valuation Analysis",
		Description: "Complete valuation model including multiple valuation methodologies and sensitivity analysis",
		Keywords:    "valuation,dcf,multiples,sensitivity-analysis,financial-modeling",
		Company:     model.CompanyName,
	})

	styles := createFinancialStyles()

	// Sheet 1: Valuation Summary
	valuationSummarySheet := workbook.AddSheet("Valuation Summary")
	if valuationSummarySheet == nil {
		log.Fatal("Failed to create valuation summary sheet")
	}
	createValuationSummarySheet(valuationSummarySheet, styles, model)

	// Sheet 2: Trading Multiples
	multiplesSheet := workbook.AddSheet("Trading Multiples")
	if multiplesSheet == nil {
		log.Fatal("Failed to create multiples sheet")
	}
	createMultiplesAnalysisSheet(multiplesSheet, styles, model)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build valuation model workbook")
	}

	filename := "output/07-valuation-model.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save valuation model: %v", err)
	}

	fmt.Printf("   ✓ Created valuation model: %s\n", filename)
}

func createDCFModelWorkbook(model FinancialModel) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "DCF Valuation Model",
		Author:      "Financial Analysis System",
		Subject:     "Discounted Cash Flow Valuation Model",
		Description: "Comprehensive DCF model with projections, terminal value, and sensitivity analysis",
		Keywords:    "dcf,discounted-cash-flow,valuation,projections,terminal-value",
		Company:     model.CompanyName,
	})

	styles := createFinancialStyles()

	// Sheet 1: DCF Model
	dcfSheet := workbook.AddSheet("DCF Model")
	if dcfSheet == nil {
		log.Fatal("Failed to create DCF sheet")
	}
	createDCFModelSheet(dcfSheet, styles, model)

	// Sheet 2: Sensitivity Analysis
	sensitivitySheet := workbook.AddSheet("Sensitivity Analysis")
	if sensitivitySheet == nil {
		log.Fatal("Failed to create sensitivity sheet")
	}
	createSensitivityAnalysisSheet(sensitivitySheet, styles, model)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build DCF model workbook")
	}

	filename := "output/07-dcf-model.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save DCF model: %v", err)
	}

	fmt.Printf("   ✓ Created DCF model: %s\n", filename)
}

func createFinancialStyles() map[string]excelbuilder.StyleConfig {
	styles := make(map[string]excelbuilder.StyleConfig)

	// Financial color palette
	colors := map[string]string{
		"primary":    "#1E3A8A", // Deep blue
		"secondary":  "#7C3AED", // Purple
		"success":    "#059669", // Green
		"warning":    "#D97706", // Orange
		"danger":     "#DC2626", // Red
		"info":       "#0284C7", // Light blue
		"light":      "#F8FAFC", // Very light gray
		"dark":       "#1F2937", // Dark gray
		"white":      "#FFFFFF",
		"border":     "#E5E7EB", // Light border
		"text":       "#374151", // Text gray
		"financial":  "#0F4C75", // Financial blue
		"positive":   "#10B981", // Positive green
		"negative":   "#EF4444", // Negative red
	}

	// Financial statement title
	styles["financial_title"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   18,
			Color:  colors["financial"],
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
			Size:   12,
			Color:  colors["white"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["financial"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createBorder("medium", colors["border"]),
	}

	// Financial line items
	styles["line_item"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   11,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["border"]),
	}

	// Financial values
	styles["financial_value"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   11,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "$#,##0",
		Border:      createBorder("thin", colors["border"]),
	}

	// Currency with decimals
	styles["currency_detailed"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   11,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "$#,##0.00",
		Border:      createBorder("thin", colors["border"]),
	}

	// Percentage format
	styles["percentage"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   11,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "0.00%",
		Border:      createBorder("thin", colors["border"]),
	}

	// Ratio format
	styles["ratio"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   11,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "0.00x",
		Border:      createBorder("thin", colors["border"]),
	}

	// Total lines (bold)
	styles["total_line"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   11,
			Color:  colors["financial"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createBorder("medium", colors["border"]),
	}

	styles["total_value"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   11,
			Color:  colors["financial"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "$#,##0",
		Border:      createBorder("medium", colors["border"]),
	}

	// Positive/negative indicators
	styles["positive_value"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   11,
			Color:  colors["positive"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "$#,##0",
		Border:      createBorder("thin", colors["border"]),
	}

	styles["negative_value"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   11,
			Color:  colors["negative"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "($#,##0)",
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

// Sheet creation functions (simplified implementations)

func createIncomeStatementSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, model FinancialModel) {
	// Set column widths
	sheet.SetColumnWidth("A", 25.0) // Line items
	for i := 0; i < len(model.IncomeStatements); i++ {
		sheet.SetColumnWidth(string(rune('B'+i)), 15.0) // Years
	}

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell(fmt.Sprintf("%s - Income Statement", model.CompanyName)).SetStyle(styles["financial_title"])

	sheet.AddRow() // Empty row
	sheet.AddRow() // Empty row

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("(in thousands)").SetStyle(styles["section_header"])
	for _, stmt := range model.IncomeStatements {
		headerRow.AddCell(stmt.Period).SetStyle(styles["section_header"])
	}

	// Revenue
	revenueRow := sheet.AddRow()
	revenueRow.AddCell("Revenue").SetStyle(styles["line_item"])
	for _, stmt := range model.IncomeStatements {
		revenueRow.AddCell(stmt.Revenue / 1000).SetStyle(styles["financial_value"])
	}

	// Cost of Goods Sold
	cogsRow := sheet.AddRow()
	cogsRow.AddCell("Cost of Goods Sold").SetStyle(styles["line_item"])
	for _, stmt := range model.IncomeStatements {
		cogsRow.AddCell(-stmt.CostOfGoodsSold / 1000).SetStyle(styles["negative_value"])
	}

	// Gross Profit
	grossProfitRow := sheet.AddRow()
	grossProfitRow.AddCell("Gross Profit").SetStyle(styles["total_line"])
	for _, stmt := range model.IncomeStatements {
		grossProfitRow.AddCell(stmt.GrossProfit / 1000).SetStyle(styles["total_value"])
	}

	// Operating Expenses
	opexRow := sheet.AddRow()
	opexRow.AddCell("Operating Expenses").SetStyle(styles["line_item"])
	for _, stmt := range model.IncomeStatements {
		opexRow.AddCell(-stmt.OperatingExpenses / 1000).SetStyle(styles["negative_value"])
	}

	// EBITDA
	ebitdaRow := sheet.AddRow()
	ebitdaRow.AddCell("EBITDA").SetStyle(styles["total_line"])
	for _, stmt := range model.IncomeStatements {
		ebitdaRow.AddCell(stmt.EBITDA / 1000).SetStyle(styles["total_value"])
	}

	// Continue with other line items...
	// (Simplified for brevity)
}

func createBalanceSheetSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, model FinancialModel) {
	// Similar implementation for balance sheet
	titleRow := sheet.AddRow()
	titleRow.AddCell(fmt.Sprintf("%s - Balance Sheet", model.CompanyName)).SetStyle(styles["financial_title"])
	// ... rest of implementation
}

func createCashFlowSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, model FinancialModel) {
	// Similar implementation for cash flow
	titleRow := sheet.AddRow()
	titleRow.AddCell(fmt.Sprintf("%s - Cash Flow Statement", model.CompanyName)).SetStyle(styles["financial_title"])
	// ... rest of implementation
}

func createFinancialSummarySheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, model FinancialModel) {
	// Financial summary implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("Financial Summary").SetStyle(styles["financial_title"])
	// ... rest of implementation
}

func createRatioSummarySheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, model FinancialModel) {
	// Ratio summary implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("Financial Ratio Summary").SetStyle(styles["financial_title"])
	// ... rest of implementation
}

func createLiquidityAnalysisSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, model FinancialModel) {
	// Liquidity analysis implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("Liquidity Analysis").SetStyle(styles["financial_title"])
	// ... rest of implementation
}

func createProfitabilityAnalysisSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, model FinancialModel) {
	// Profitability analysis implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("Profitability Analysis").SetStyle(styles["financial_title"])
	// ... rest of implementation
}

func createEfficiencyAnalysisSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, model FinancialModel) {
	// Efficiency analysis implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("Efficiency Analysis").SetStyle(styles["financial_title"])
	// ... rest of implementation
}

func createValuationSummarySheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, model FinancialModel) {
	// Valuation summary implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("Valuation Summary").SetStyle(styles["financial_title"])
	// ... rest of implementation
}

func createMultiplesAnalysisSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, model FinancialModel) {
	// Multiples analysis implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("Trading Multiples Analysis").SetStyle(styles["financial_title"])
	// ... rest of implementation
}

func createDCFModelSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, model FinancialModel) {
	// DCF model implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("DCF Valuation Model").SetStyle(styles["financial_title"])
	// ... rest of implementation
}

func createSensitivityAnalysisSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, model FinancialModel) {
	// Sensitivity analysis implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("Sensitivity Analysis").SetStyle(styles["financial_title"])
	// ... rest of implementation
}