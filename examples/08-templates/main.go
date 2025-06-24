package main

import (
	"fmt"
	"log"
	"os"
	"strings"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Template data structures
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

type SheetTemplate struct {
	Name     string
	Sections []SectionTemplate
}

type SectionTemplate struct {
	Type     string // "header", "table", "chart", "summary"
	Title    string
	Data     interface{}
	Style    string
	Position Position
}

type Position struct {
	StartRow int
	StartCol int
	EndRow   int
	EndCol   int
}

// Business data structures for templates
type Employee struct {
	ID         int
	Name       string
	Department string
	Position   string
	Salary     float64
	HireDate   time.Time
	Manager    string
	Location   string
}

type Project struct {
	ID          string
	Name        string
	Description string
	StartDate   time.Time
	EndDate     time.Time
	Budget      float64
	Spent       float64
	Status      string
	Manager     string
	TeamSize    int
}

type FinancialData struct {
	Period   string
	Revenue  float64
	Expenses float64
	Profit   float64
	Margin   float64
}

type InventoryItem struct {
	SKU         string
	Name        string
	Category    string
	Quantity    int
	UnitPrice   float64
	TotalValue  float64
	Supplier    string
	LastUpdated time.Time
}

func main() {
	// Create output directory
	if err := os.MkdirAll("output", 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}

	fmt.Println("📋 Starting Template System Demo...")

	// Step 1: Create reusable templates
	fmt.Println("\n🎨 Step 1: Creating reusable template system...")
	templateSystem := createTemplateSystem()

	// Step 2: Generate employee report from template
	fmt.Println("\n👥 Step 2: Generating employee report from template...")
	generateEmployeeReport(templateSystem)

	// Step 3: Generate project dashboard from template
	fmt.Println("\n📊 Step 3: Generating project dashboard from template...")
	generateProjectDashboard(templateSystem)

	// Step 4: Generate financial report from template
	fmt.Println("\n💰 Step 4: Generating financial report from template...")
	generateFinancialReport(templateSystem)

	// Step 5: Generate inventory report from template
	fmt.Println("\n📦 Step 5: Generating inventory report from template...")
	generateInventoryReport(templateSystem)

	// Step 6: Create custom template variations
	fmt.Println("\n🔧 Step 6: Creating custom template variations...")
	createCustomTemplateVariations(templateSystem)

	fmt.Println("\n✅ Template system demonstration completed successfully!")
	fmt.Println("📁 Generated files:")
	fmt.Println("   • output/08-employee-report.xlsx - Employee management report")
	fmt.Println("   • output/08-project-dashboard.xlsx - Project tracking dashboard")
	fmt.Println("   • output/08-financial-report.xlsx - Financial summary report")
	fmt.Println("   • output/08-inventory-report.xlsx - Inventory management report")
	fmt.Println("   • output/08-custom-variations.xlsx - Custom template variations")
	fmt.Println("\n🎯 Next steps: Try examples/09-advanced-layout/ for complex layouts")
}

func createTemplateSystem() map[string]TemplateConfig {
	fmt.Println("   🎨 Setting up template configurations...")
	
	templates := make(map[string]TemplateConfig)
	styles := createTemplateStyles()

	// Employee Report Template
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

	// Project Dashboard Template
	templates["project_dashboard"] = TemplateConfig{
		Name:        "Project Management Dashboard",
		Description: "Real-time project tracking and performance analysis",
		Category:    "Project Management",
		Variables: map[string]interface{}{
			"dashboard_title": "{{DASHBOARD_TITLE}}",
			"update_date":     "{{UPDATE_DATE}}",
			"total_projects":  "{{TOTAL_PROJECTS}}",
			"active_projects": "{{ACTIVE_PROJECTS}}",
		},
		Styles: styles,
	}

	// Financial Report Template
	templates["financial_report"] = TemplateConfig{
		Name:        "Financial Summary Report",
		Description: "Executive financial performance summary",
		Category:    "Finance",
		Variables: map[string]interface{}{
			"period":        "{{PERIOD}}",
			"fiscal_year":   "{{FISCAL_YEAR}}",
			"total_revenue": "{{TOTAL_REVENUE}}",
			"net_profit":    "{{NET_PROFIT}}",
		},
		Styles: styles,
	}

	// Inventory Report Template
	templates["inventory_report"] = TemplateConfig{
		Name:        "Inventory Management Report",
		Description: "Comprehensive inventory analysis and tracking",
		Category:    "Operations",
		Variables: map[string]interface{}{
			"warehouse":     "{{WAREHOUSE}}",
			"report_date":   "{{REPORT_DATE}}",
			"total_items":   "{{TOTAL_ITEMS}}",
			"total_value":   "{{TOTAL_VALUE}}",
		},
		Styles: styles,
	}

	return templates
}

func generateEmployeeReport(templates map[string]TemplateConfig) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Employee Management Report",
		Author:      "Template System",
		Subject:     "HR Analytics and Employee Management",
		Description: "Comprehensive employee report generated from reusable template",
		Keywords:    "employees,hr,analytics,template,management",
		Company:     "TechCorp Solutions",
	})

	template := templates["employee_report"]
	styles := template.Styles

	// Generate sample employee data
	employees := generateSampleEmployees()

	// Sheet 1: Employee Overview
	overviewSheet := workbook.AddSheet("Employee Overview")
	if overviewSheet == nil {
		log.Fatal("Failed to create employee overview sheet")
	}
	createEmployeeOverviewSheet(overviewSheet, styles, employees)

	// Sheet 2: Department Analysis
	deptSheet := workbook.AddSheet("Department Analysis")
	if deptSheet == nil {
		log.Fatal("Failed to create department analysis sheet")
	}
	createDepartmentAnalysisSheet(deptSheet, styles, employees)

	// Sheet 3: Salary Analysis
	salarySheet := workbook.AddSheet("Salary Analysis")
	if salarySheet == nil {
		log.Fatal("Failed to create salary analysis sheet")
	}
	createSalaryAnalysisSheet(salarySheet, styles, employees)

	// Sheet 4: Employee Directory
	directorySheet := workbook.AddSheet("Employee Directory")
	if directorySheet == nil {
		log.Fatal("Failed to create employee directory sheet")
	}
	createEmployeeDirectorySheet(directorySheet, styles, employees)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build employee report workbook")
	}

	filename := "output/08-employee-report.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save employee report: %v", err)
	}

	fmt.Printf("   ✓ Created employee report: %s\n", filename)
}

func generateProjectDashboard(templates map[string]TemplateConfig) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Project Management Dashboard",
		Author:      "Template System",
		Subject:     "Project Tracking and Performance Analysis",
		Description: "Real-time project dashboard generated from reusable template",
		Keywords:    "projects,dashboard,tracking,template,management",
		Company:     "TechCorp Solutions",
	})

	template := templates["project_dashboard"]
	styles := template.Styles

	// Generate sample project data
	projects := generateSampleProjects()

	// Sheet 1: Project Dashboard
	dashboardSheet := workbook.AddSheet("Project Dashboard")
	if dashboardSheet == nil {
		log.Fatal("Failed to create project dashboard sheet")
	}
	createProjectDashboardSheet(dashboardSheet, styles, projects)

	// Sheet 2: Project Status
	statusSheet := workbook.AddSheet("Project Status")
	if statusSheet == nil {
		log.Fatal("Failed to create project status sheet")
	}
	createProjectStatusSheet(statusSheet, styles, projects)

	// Sheet 3: Budget Analysis
	budgetSheet := workbook.AddSheet("Budget Analysis")
	if budgetSheet == nil {
		log.Fatal("Failed to create budget analysis sheet")
	}
	createProjectBudgetSheet(budgetSheet, styles, projects)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build project dashboard workbook")
	}

	filename := "output/08-project-dashboard.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save project dashboard: %v", err)
	}

	fmt.Printf("   ✓ Created project dashboard: %s\n", filename)
}

func generateFinancialReport(templates map[string]TemplateConfig) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Financial Summary Report",
		Author:      "Template System",
		Subject:     "Executive Financial Performance Summary",
		Description: "Financial report generated from reusable template",
		Keywords:    "finance,financial,template,summary,executive",
		Company:     "TechCorp Solutions",
	})

	template := templates["financial_report"]
	styles := template.Styles

	// Generate sample financial data
	financialData := generateSampleFinancialData()

	// Sheet 1: Executive Summary
	execSheet := workbook.AddSheet("Executive Summary")
	if execSheet == nil {
		log.Fatal("Failed to create executive summary sheet")
	}
	createFinancialExecutiveSummarySheet(execSheet, styles, financialData)

	// Sheet 2: Revenue Analysis
	revenueSheet := workbook.AddSheet("Revenue Analysis")
	if revenueSheet == nil {
		log.Fatal("Failed to create revenue analysis sheet")
	}
	createRevenueAnalysisSheet(revenueSheet, styles, financialData)

	// Sheet 3: Expense Analysis
	expenseSheet := workbook.AddSheet("Expense Analysis")
	if expenseSheet == nil {
		log.Fatal("Failed to create expense analysis sheet")
	}
	createExpenseAnalysisSheet(expenseSheet, styles, financialData)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build financial report workbook")
	}

	filename := "output/08-financial-report.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save financial report: %v", err)
	}

	fmt.Printf("   ✓ Created financial report: %s\n", filename)
}

func generateInventoryReport(templates map[string]TemplateConfig) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Inventory Management Report",
		Author:      "Template System",
		Subject:     "Comprehensive Inventory Analysis",
		Description: "Inventory report generated from reusable template",
		Keywords:    "inventory,warehouse,template,management,analysis",
		Company:     "TechCorp Solutions",
	})

	template := templates["inventory_report"]
	styles := template.Styles

	// Generate sample inventory data
	inventory := generateSampleInventory()

	// Sheet 1: Inventory Overview
	overviewSheet := workbook.AddSheet("Inventory Overview")
	if overviewSheet == nil {
		log.Fatal("Failed to create inventory overview sheet")
	}
	createInventoryOverviewSheet(overviewSheet, styles, inventory)

	// Sheet 2: Category Analysis
	categorySheet := workbook.AddSheet("Category Analysis")
	if categorySheet == nil {
		log.Fatal("Failed to create category analysis sheet")
	}
	createInventoryCategorySheet(categorySheet, styles, inventory)

	// Sheet 3: Stock Levels
	stockSheet := workbook.AddSheet("Stock Levels")
	if stockSheet == nil {
		log.Fatal("Failed to create stock levels sheet")
	}
	createStockLevelsSheet(stockSheet, styles, inventory)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build inventory report workbook")
	}

	filename := "output/08-inventory-report.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save inventory report: %v", err)
	}

	fmt.Printf("   ✓ Created inventory report: %s\n", filename)
}

func createCustomTemplateVariations(templates map[string]TemplateConfig) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Custom Template Variations",
		Author:      "Template System",
		Subject:     "Template Customization Examples",
		Description: "Examples of template customization and variation",
		Keywords:    "templates,customization,variations,examples",
		Company:     "TechCorp Solutions",
	})

	styles := createTemplateStyles()

	// Sheet 1: Template Comparison
	comparisonSheet := workbook.AddSheet("Template Comparison")
	if comparisonSheet == nil {
		log.Fatal("Failed to create template comparison sheet")
	}
	createTemplateComparisonSheet(comparisonSheet, styles, templates)

	// Sheet 2: Style Variations
	styleSheet := workbook.AddSheet("Style Variations")
	if styleSheet == nil {
		log.Fatal("Failed to create style variations sheet")
	}
	createStyleVariationsSheet(styleSheet, styles)

	// Sheet 3: Dynamic Content
	dynamicSheet := workbook.AddSheet("Dynamic Content")
	if dynamicSheet == nil {
		log.Fatal("Failed to create dynamic content sheet")
	}
	createDynamicContentSheet(dynamicSheet, styles)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build custom variations workbook")
	}

	filename := "output/08-custom-variations.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save custom variations: %v", err)
	}

	fmt.Printf("   ✓ Created custom variations: %s\n", filename)
}

func createTemplateStyles() map[string]excelbuilder.StyleConfig {
	styles := make(map[string]excelbuilder.StyleConfig)

	// Template color palette
	colors := map[string]string{
		"primary":    "#2563EB", // Blue
		"secondary":  "#7C3AED", // Purple
		"success":    "#10B981", // Green
		"warning":    "#F59E0B", // Amber
		"danger":     "#EF4444", // Red
		"info":       "#06B6D4", // Cyan
		"light":      "#F9FAFB", // Gray 50
		"dark":       "#111827", // Gray 900
		"white":      "#FFFFFF",
		"border":     "#D1D5DB", // Gray 300
		"text":       "#374151", // Gray 700
		"muted":      "#6B7280", // Gray 500
	}

	// Template title style
	styles["template_title"] = excelbuilder.StyleConfig{
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

	// Section header style
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
		Border: createTemplateBorder("medium", colors["border"]),
	}

	// Table header style
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
		Border: createTemplateBorder("thin", colors["border"]),
	}

	// Data cell style
	styles["data_cell"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createTemplateBorder("thin", colors["border"]),
	}

	// Numeric cell style
	styles["numeric_cell"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "#,##0",
		Border:      createTemplateBorder("thin", colors["border"]),
	}

	// Currency cell style
	styles["currency_cell"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "$#,##0.00",
		Border:      createTemplateBorder("thin", colors["border"]),
	}

	// Date cell style
	styles["date_cell"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		NumberFormat: "mm/dd/yyyy",
		Border:      createTemplateBorder("thin", colors["border"]),
	}

	// Percentage cell style
	styles["percentage_cell"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "0.00%",
		Border:      createTemplateBorder("thin", colors["border"]),
	}

	// Status indicators
	styles["status_active"] = excelbuilder.StyleConfig{
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
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createTemplateBorder("thin", colors["border"]),
	}

	styles["status_warning"] = excelbuilder.StyleConfig{
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
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createTemplateBorder("thin", colors["border"]),
	}

	styles["status_danger"] = excelbuilder.StyleConfig{
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
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createTemplateBorder("thin", colors["border"]),
	}

	// Summary styles
	styles["summary_label"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
			Color:  colors["primary"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
	}

	styles["summary_value"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
			Color:  colors["text"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "#,##0",
	}

	return styles
}

func createTemplateBorder(style, color string) excelbuilder.BorderConfig {
	return excelbuilder.BorderConfig{
		Top:    excelbuilder.BorderSide{Style: style, Color: color},
		Bottom: excelbuilder.BorderSide{Style: style, Color: color},
		Left:   excelbuilder.BorderSide{Style: style, Color: color},
		Right:  excelbuilder.BorderSide{Style: style, Color: color},
	}
}

// Data generation functions

func generateSampleEmployees() []Employee {
	employees := []Employee{
		{ID: 1, Name: "John Smith", Department: "Engineering", Position: "Senior Developer", Salary: 95000, HireDate: time.Date(2020, 3, 15, 0, 0, 0, 0, time.UTC), Manager: "Alice Johnson", Location: "New York"},
		{ID: 2, Name: "Sarah Davis", Department: "Marketing", Position: "Marketing Manager", Salary: 78000, HireDate: time.Date(2019, 7, 22, 0, 0, 0, 0, time.UTC), Manager: "Bob Wilson", Location: "San Francisco"},
		{ID: 3, Name: "Mike Johnson", Department: "Engineering", Position: "Lead Developer", Salary: 110000, HireDate: time.Date(2018, 1, 10, 0, 0, 0, 0, time.UTC), Manager: "Alice Johnson", Location: "New York"},
		{ID: 4, Name: "Emily Brown", Department: "Sales", Position: "Sales Representative", Salary: 65000, HireDate: time.Date(2021, 5, 8, 0, 0, 0, 0, time.UTC), Manager: "David Lee", Location: "Chicago"},
		{ID: 5, Name: "Alex Wilson", Department: "HR", Position: "HR Specialist", Salary: 58000, HireDate: time.Date(2020, 9, 12, 0, 0, 0, 0, time.UTC), Manager: "Lisa Chen", Location: "Austin"},
		{ID: 6, Name: "Lisa Chen", Department: "HR", Position: "HR Director", Salary: 85000, HireDate: time.Date(2017, 4, 3, 0, 0, 0, 0, time.UTC), Manager: "CEO", Location: "Austin"},
		{ID: 7, Name: "David Lee", Department: "Sales", Position: "Sales Director", Salary: 98000, HireDate: time.Date(2016, 11, 20, 0, 0, 0, 0, time.UTC), Manager: "CEO", Location: "Chicago"},
		{ID: 8, Name: "Alice Johnson", Department: "Engineering", Position: "Engineering Manager", Salary: 125000, HireDate: time.Date(2015, 8, 14, 0, 0, 0, 0, time.UTC), Manager: "CTO", Location: "New York"},
	}
	return employees
}

func generateSampleProjects() []Project {
	projects := []Project{
		{ID: "PRJ-001", Name: "Mobile App Redesign", Description: "Complete redesign of mobile application", StartDate: time.Date(2024, 1, 15, 0, 0, 0, 0, time.UTC), EndDate: time.Date(2024, 6, 30, 0, 0, 0, 0, time.UTC), Budget: 250000, Spent: 180000, Status: "In Progress", Manager: "Alice Johnson", TeamSize: 8},
		{ID: "PRJ-002", Name: "Data Analytics Platform", Description: "New analytics platform development", StartDate: time.Date(2024, 2, 1, 0, 0, 0, 0, time.UTC), EndDate: time.Date(2024, 8, 15, 0, 0, 0, 0, time.UTC), Budget: 400000, Spent: 120000, Status: "In Progress", Manager: "Mike Johnson", TeamSize: 12},
		{ID: "PRJ-003", Name: "Customer Portal", Description: "Self-service customer portal", StartDate: time.Date(2023, 10, 1, 0, 0, 0, 0, time.UTC), EndDate: time.Date(2024, 3, 31, 0, 0, 0, 0, time.UTC), Budget: 180000, Spent: 175000, Status: "Completed", Manager: "Sarah Davis", TeamSize: 6},
		{ID: "PRJ-004", Name: "Security Upgrade", Description: "Infrastructure security improvements", StartDate: time.Date(2024, 3, 1, 0, 0, 0, 0, time.UTC), EndDate: time.Date(2024, 5, 31, 0, 0, 0, 0, time.UTC), Budget: 150000, Spent: 45000, Status: "Planning", Manager: "John Smith", TeamSize: 4},
		{ID: "PRJ-005", Name: "API Integration", Description: "Third-party API integrations", StartDate: time.Date(2024, 4, 1, 0, 0, 0, 0, time.UTC), EndDate: time.Date(2024, 7, 31, 0, 0, 0, 0, time.UTC), Budget: 120000, Spent: 85000, Status: "At Risk", Manager: "Emily Brown", TeamSize: 5},
	}
	return projects
}

func generateSampleFinancialData() []FinancialData {
	data := []FinancialData{
		{Period: "Q1 2024", Revenue: 2500000, Expenses: 1800000, Profit: 700000, Margin: 0.28},
		{Period: "Q2 2024", Revenue: 2750000, Expenses: 1950000, Profit: 800000, Margin: 0.29},
		{Period: "Q3 2024", Revenue: 2900000, Expenses: 2100000, Profit: 800000, Margin: 0.28},
		{Period: "Q4 2024", Revenue: 3200000, Expenses: 2200000, Profit: 1000000, Margin: 0.31},
	}
	return data
}

func generateSampleInventory() []InventoryItem {
	items := []InventoryItem{
		{SKU: "TECH-001", Name: "Wireless Headphones", Category: "Electronics", Quantity: 150, UnitPrice: 89.99, TotalValue: 13498.50, Supplier: "TechSupply Co", LastUpdated: time.Now().AddDate(0, 0, -2)},
		{SKU: "TECH-002", Name: "Bluetooth Speaker", Category: "Electronics", Quantity: 75, UnitPrice: 129.99, TotalValue: 9749.25, Supplier: "AudioTech Ltd", LastUpdated: time.Now().AddDate(0, 0, -1)},
		{SKU: "OFF-001", Name: "Ergonomic Chair", Category: "Office Furniture", Quantity: 25, UnitPrice: 299.99, TotalValue: 7499.75, Supplier: "OfficeMax Pro", LastUpdated: time.Now().AddDate(0, 0, -3)},
		{SKU: "OFF-002", Name: "Standing Desk", Category: "Office Furniture", Quantity: 12, UnitPrice: 449.99, TotalValue: 5399.88, Supplier: "OfficeMax Pro", LastUpdated: time.Now().AddDate(0, 0, -5)},
		{SKU: "COMP-001", Name: "Laptop Computer", Category: "Computers", Quantity: 30, UnitPrice: 1299.99, TotalValue: 38999.70, Supplier: "CompuWorld", LastUpdated: time.Now().AddDate(0, 0, -1)},
		{SKU: "COMP-002", Name: "Desktop Monitor", Category: "Computers", Quantity: 45, UnitPrice: 249.99, TotalValue: 11249.55, Supplier: "DisplayTech", LastUpdated: time.Now()},
	}
	return items
}

// Sheet creation functions (simplified implementations)

func createEmployeeOverviewSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, employees []Employee) {
	// Set column widths
	sheet.SetColumnWidth("A", 20.0)
	sheet.SetColumnWidth("B", 25.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 20.0)
	sheet.SetColumnWidth("E", 12.0)
	sheet.SetColumnWidth("F", 12.0)
	sheet.SetColumnWidth("G", 15.0)
	sheet.SetColumnWidth("H", 12.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Employee Management Report").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Summary section
	summaryRow := sheet.AddRow()
	summaryRow.AddCell("Total Employees:").SetStyle(styles["summary_label"])
	summaryRow.AddCell(len(employees)).SetStyle(styles["summary_value"])

	sheet.AddRow() // Empty row

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Employee ID").SetStyle(styles["table_header"])
	headerRow.AddCell("Name").SetStyle(styles["table_header"])
	headerRow.AddCell("Department").SetStyle(styles["table_header"])
	headerRow.AddCell("Position").SetStyle(styles["table_header"])
	headerRow.AddCell("Salary").SetStyle(styles["table_header"])
	headerRow.AddCell("Hire Date").SetStyle(styles["table_header"])
	headerRow.AddCell("Manager").SetStyle(styles["table_header"])
	headerRow.AddCell("Location").SetStyle(styles["table_header"])

	// Employee data
	for _, emp := range employees {
		dataRow := sheet.AddRow()
		dataRow.AddCell(emp.ID).SetStyle(styles["numeric_cell"])
		dataRow.AddCell(emp.Name).SetStyle(styles["data_cell"])
		dataRow.AddCell(emp.Department).SetStyle(styles["data_cell"])
		dataRow.AddCell(emp.Position).SetStyle(styles["data_cell"])
		dataRow.AddCell(emp.Salary).SetStyle(styles["currency_cell"])
		dataRow.AddCell(emp.HireDate).SetStyle(styles["date_cell"])
		dataRow.AddCell(emp.Manager).SetStyle(styles["data_cell"])
		dataRow.AddCell(emp.Location).SetStyle(styles["data_cell"])
	}
}

func createDepartmentAnalysisSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, employees []Employee) {
	// Set column widths
	sheet.SetColumnWidth("A", 20.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 15.0)
	sheet.SetColumnWidth("E", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Department Analysis").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Calculate department statistics
	deptStats := make(map[string]struct {
		Count      int
		TotalSalary float64
		AvgSalary   float64
	})

	for _, emp := range employees {
		stats := deptStats[emp.Department]
		stats.Count++
		stats.TotalSalary += emp.Salary
		deptStats[emp.Department] = stats
	}

	// Calculate averages
	for dept, stats := range deptStats {
		stats.AvgSalary = stats.TotalSalary / float64(stats.Count)
		deptStats[dept] = stats
	}

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Department").SetStyle(styles["table_header"])
	headerRow.AddCell("Employee Count").SetStyle(styles["table_header"])
	headerRow.AddCell("Total Salary").SetStyle(styles["table_header"])
	headerRow.AddCell("Average Salary").SetStyle(styles["table_header"])
	headerRow.AddCell("Percentage").SetStyle(styles["table_header"])

	// Department data
	totalEmployees := len(employees)
	for dept, stats := range deptStats {
		dataRow := sheet.AddRow()
		dataRow.AddCell(dept).SetStyle(styles["data_cell"])
		dataRow.AddCell(stats.Count).SetStyle(styles["numeric_cell"])
		dataRow.AddCell(stats.TotalSalary).SetStyle(styles["currency_cell"])
		dataRow.AddCell(stats.AvgSalary).SetStyle(styles["currency_cell"])
		percentage := float64(stats.Count) / float64(totalEmployees)
		dataRow.AddCell(percentage).SetStyle(styles["percentage_cell"])
	}
}

func createSalaryAnalysisSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, employees []Employee) {
	// Set column widths
	sheet.SetColumnWidth("A", 20.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Salary Analysis").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Calculate salary statistics
	var totalSalary, minSalary, maxSalary float64
	minSalary = employees[0].Salary
	maxSalary = employees[0].Salary

	for _, emp := range employees {
		totalSalary += emp.Salary
		if emp.Salary < minSalary {
			minSalary = emp.Salary
		}
		if emp.Salary > maxSalary {
			maxSalary = emp.Salary
		}
	}

	avgSalary := totalSalary / float64(len(employees))

	// Summary statistics
	summaryRow1 := sheet.AddRow()
	summaryRow1.AddCell("Total Employees:").SetStyle(styles["summary_label"])
	summaryRow1.AddCell(len(employees)).SetStyle(styles["summary_value"])

	summaryRow2 := sheet.AddRow()
	summaryRow2.AddCell("Average Salary:").SetStyle(styles["summary_label"])
	summaryRow2.AddCell(avgSalary).SetStyle(styles["currency_cell"])

	summaryRow3 := sheet.AddRow()
	summaryRow3.AddCell("Minimum Salary:").SetStyle(styles["summary_label"])
	summaryRow3.AddCell(minSalary).SetStyle(styles["currency_cell"])

	summaryRow4 := sheet.AddRow()
	summaryRow4.AddCell("Maximum Salary:").SetStyle(styles["summary_label"])
	summaryRow4.AddCell(maxSalary).SetStyle(styles["currency_cell"])

	sheet.AddRow() // Empty row

	// Salary ranges
	headerRow := sheet.AddRow()
	headerRow.AddCell("Salary Range").SetStyle(styles["table_header"])
	headerRow.AddCell("Employee Count").SetStyle(styles["table_header"])
	headerRow.AddCell("Percentage").SetStyle(styles["table_header"])

	// Define salary ranges
	ranges := []struct {
		Label string
		Min   float64
		Max   float64
	}{
		{"Under $60K", 0, 60000},
		{"$60K - $80K", 60000, 80000},
		{"$80K - $100K", 80000, 100000},
		{"$100K - $120K", 100000, 120000},
		{"Over $120K", 120000, 999999},
	}

	// Count employees in each range
	for _, salaryRange := range ranges {
		count := 0
		for _, emp := range employees {
			if emp.Salary >= salaryRange.Min && emp.Salary < salaryRange.Max {
				count++
			}
		}
		percentage := float64(count) / float64(len(employees))

		dataRow := sheet.AddRow()
		dataRow.AddCell(salaryRange.Label).SetStyle(styles["data_cell"])
		dataRow.AddCell(count).SetStyle(styles["numeric_cell"])
		dataRow.AddCell(percentage).SetStyle(styles["percentage_cell"])
	}
}

func createEmployeeDirectorySheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, employees []Employee) {
	// Set column widths
	sheet.SetColumnWidth("A", 25.0)
	sheet.SetColumnWidth("B", 20.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 12.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Employee Directory").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Name").SetStyle(styles["table_header"])
	headerRow.AddCell("Position").SetStyle(styles["table_header"])
	headerRow.AddCell("Department").SetStyle(styles["table_header"])
	headerRow.AddCell("Location").SetStyle(styles["table_header"])

	// Employee directory data (sorted by name)
	for _, emp := range employees {
		dataRow := sheet.AddRow()
		dataRow.AddCell(emp.Name).SetStyle(styles["data_cell"])
		dataRow.AddCell(emp.Position).SetStyle(styles["data_cell"])
		dataRow.AddCell(emp.Department).SetStyle(styles["data_cell"])
		dataRow.AddCell(emp.Location).SetStyle(styles["data_cell"])
	}
}

func createProjectDashboardSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, projects []Project) {
	// Set column widths
	sheet.SetColumnWidth("A", 15.0)
	sheet.SetColumnWidth("B", 25.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 15.0)
	sheet.SetColumnWidth("E", 15.0)
	sheet.SetColumnWidth("F", 15.0)
	sheet.SetColumnWidth("G", 12.0)
	sheet.SetColumnWidth("H", 15.0)
	sheet.SetColumnWidth("I", 10.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Project Management Dashboard").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Calculate project statistics
	var totalBudget, totalSpent float64
	statusCount := make(map[string]int)

	for _, project := range projects {
		totalBudget += project.Budget
		totalSpent += project.Spent
		statusCount[project.Status]++
	}

	// Summary section
	summaryRow1 := sheet.AddRow()
	summaryRow1.AddCell("Total Projects:").SetStyle(styles["summary_label"])
	summaryRow1.AddCell(len(projects)).SetStyle(styles["summary_value"])

	summaryRow2 := sheet.AddRow()
	summaryRow2.AddCell("Total Budget:").SetStyle(styles["summary_label"])
	summaryRow2.AddCell(totalBudget).SetStyle(styles["currency_cell"])

	summaryRow3 := sheet.AddRow()
	summaryRow3.AddCell("Total Spent:").SetStyle(styles["summary_label"])
	summaryRow3.AddCell(totalSpent).SetStyle(styles["currency_cell"])

	budgetUtilization := totalSpent / totalBudget
	summaryRow4 := sheet.AddRow()
	summaryRow4.AddCell("Budget Utilization:").SetStyle(styles["summary_label"])
	summaryRow4.AddCell(budgetUtilization).SetStyle(styles["percentage_cell"])

	sheet.AddRow() // Empty row

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Project ID").SetStyle(styles["table_header"])
	headerRow.AddCell("Project Name").SetStyle(styles["table_header"])
	headerRow.AddCell("Manager").SetStyle(styles["table_header"])
	headerRow.AddCell("Budget").SetStyle(styles["table_header"])
	headerRow.AddCell("Spent").SetStyle(styles["table_header"])
	headerRow.AddCell("Remaining").SetStyle(styles["table_header"])
	headerRow.AddCell("Status").SetStyle(styles["table_header"])
	headerRow.AddCell("End Date").SetStyle(styles["table_header"])
	headerRow.AddCell("Team Size").SetStyle(styles["table_header"])

	// Project data
	for _, project := range projects {
		dataRow := sheet.AddRow()
		dataRow.AddCell(project.ID).SetStyle(styles["data_cell"])
		dataRow.AddCell(project.Name).SetStyle(styles["data_cell"])
		dataRow.AddCell(project.Manager).SetStyle(styles["data_cell"])
		dataRow.AddCell(project.Budget).SetStyle(styles["currency_cell"])
		dataRow.AddCell(project.Spent).SetStyle(styles["currency_cell"])
		remaining := project.Budget - project.Spent
		dataRow.AddCell(remaining).SetStyle(styles["currency_cell"])

		// Status with conditional styling
		var statusStyle string
		switch project.Status {
		case "Completed":
			statusStyle = "status_active"
		case "In Progress":
			statusStyle = "status_warning"
		case "At Risk":
			statusStyle = "status_danger"
		default:
			statusStyle = "data_cell"
		}
		dataRow.AddCell(project.Status).SetStyle(styles[statusStyle])
		dataRow.AddCell(project.EndDate).SetStyle(styles["date_cell"])
		dataRow.AddCell(project.TeamSize).SetStyle(styles["numeric_cell"])
	}
}

func createProjectStatusSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, projects []Project) {
	// Set column widths
	sheet.SetColumnWidth("A", 20.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Project Status Overview").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Calculate status statistics
	statusCount := make(map[string]int)
	for _, project := range projects {
		statusCount[project.Status]++
	}

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Status").SetStyle(styles["table_header"])
	headerRow.AddCell("Project Count").SetStyle(styles["table_header"])
	headerRow.AddCell("Percentage").SetStyle(styles["table_header"])

	// Status data
	totalProjects := len(projects)
	for status, count := range statusCount {
		dataRow := sheet.AddRow()
		var statusStyle string
		switch status {
		case "Completed":
			statusStyle = "status_active"
		case "In Progress":
			statusStyle = "status_warning"
		case "At Risk":
			statusStyle = "status_danger"
		default:
			statusStyle = "data_cell"
		}
		dataRow.AddCell(status).SetStyle(styles[statusStyle])
		dataRow.AddCell(count).SetStyle(styles["numeric_cell"])
		percentage := float64(count) / float64(totalProjects)
		dataRow.AddCell(percentage).SetStyle(styles["percentage_cell"])
	}

	sheet.AddRow() // Empty row

	// Detailed project list by status
	for status := range statusCount {
		sheet.AddRow() // Empty row
		sectionRow := sheet.AddRow()
		sectionRow.AddCell(fmt.Sprintf("%s Projects", status)).SetStyle(styles["section_header"])

		detailHeaderRow := sheet.AddRow()
		detailHeaderRow.AddCell("Project Name").SetStyle(styles["table_header"])
		detailHeaderRow.AddCell("Manager").SetStyle(styles["table_header"])
		detailHeaderRow.AddCell("End Date").SetStyle(styles["table_header"])

		for _, project := range projects {
			if project.Status == status {
				detailRow := sheet.AddRow()
				detailRow.AddCell(project.Name).SetStyle(styles["data_cell"])
				detailRow.AddCell(project.Manager).SetStyle(styles["data_cell"])
				detailRow.AddCell(project.EndDate).SetStyle(styles["date_cell"])
			}
		}
	}
}

func createProjectBudgetSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, projects []Project) {
	// Set column widths
	sheet.SetColumnWidth("A", 25.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 15.0)
	sheet.SetColumnWidth("E", 15.0)
	sheet.SetColumnWidth("F", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Budget Analysis").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Calculate budget totals
	var totalBudget, totalSpent, totalRemaining float64
	for _, project := range projects {
		totalBudget += project.Budget
		totalSpent += project.Spent
		totalRemaining += (project.Budget - project.Spent)
	}

	// Summary section
	summaryRow1 := sheet.AddRow()
	summaryRow1.AddCell("Total Budget Allocated:").SetStyle(styles["summary_label"])
	summaryRow1.AddCell(totalBudget).SetStyle(styles["currency_cell"])

	summaryRow2 := sheet.AddRow()
	summaryRow2.AddCell("Total Amount Spent:").SetStyle(styles["summary_label"])
	summaryRow2.AddCell(totalSpent).SetStyle(styles["currency_cell"])

	summaryRow3 := sheet.AddRow()
	summaryRow3.AddCell("Total Remaining:").SetStyle(styles["summary_label"])
	summaryRow3.AddCell(totalRemaining).SetStyle(styles["currency_cell"])

	overallUtilization := totalSpent / totalBudget
	summaryRow4 := sheet.AddRow()
	summaryRow4.AddCell("Overall Utilization:").SetStyle(styles["summary_label"])
	summaryRow4.AddCell(overallUtilization).SetStyle(styles["percentage_cell"])

	sheet.AddRow() // Empty row

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Project Name").SetStyle(styles["table_header"])
	headerRow.AddCell("Budget").SetStyle(styles["table_header"])
	headerRow.AddCell("Spent").SetStyle(styles["table_header"])
	headerRow.AddCell("Remaining").SetStyle(styles["table_header"])
	headerRow.AddCell("Utilization").SetStyle(styles["table_header"])
	headerRow.AddCell("Status").SetStyle(styles["table_header"])

	// Project budget data
	for _, project := range projects {
		dataRow := sheet.AddRow()
		dataRow.AddCell(project.Name).SetStyle(styles["data_cell"])
		dataRow.AddCell(project.Budget).SetStyle(styles["currency_cell"])
		dataRow.AddCell(project.Spent).SetStyle(styles["currency_cell"])
		remaining := project.Budget - project.Spent
		dataRow.AddCell(remaining).SetStyle(styles["currency_cell"])
		utilization := project.Spent / project.Budget
		dataRow.AddCell(utilization).SetStyle(styles["percentage_cell"])

		// Budget status based on utilization
		var budgetStatus string
		var statusStyle string
		if utilization > 0.9 {
			budgetStatus = "High Utilization"
			statusStyle = "status_danger"
		} else if utilization > 0.7 {
			budgetStatus = "Medium Utilization"
			statusStyle = "status_warning"
		} else {
			budgetStatus = "Low Utilization"
			statusStyle = "status_active"
		}
		dataRow.AddCell(budgetStatus).SetStyle(styles[statusStyle])
	}
}

func createFinancialExecutiveSummarySheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, data []FinancialData) {
	// Set column widths
	sheet.SetColumnWidth("A", 25.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Executive Financial Summary").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Calculate financial totals
	var totalRevenue, totalExpenses, totalProfit float64
	for _, financialData := range data {
		totalRevenue += financialData.Revenue
		totalExpenses += financialData.Expenses
		totalProfit += financialData.Profit
	}

	// Key metrics summary
	summaryRow1 := sheet.AddRow()
	summaryRow1.AddCell("Total Revenue:").SetStyle(styles["summary_label"])
	summaryRow1.AddCell(totalRevenue).SetStyle(styles["currency_cell"])

	summaryRow2 := sheet.AddRow()
	summaryRow2.AddCell("Total Expenses:").SetStyle(styles["summary_label"])
	summaryRow2.AddCell(totalExpenses).SetStyle(styles["currency_cell"])

	summaryRow3 := sheet.AddRow()
	summaryRow3.AddCell("Total Profit:").SetStyle(styles["summary_label"])
	summaryRow3.AddCell(totalProfit).SetStyle(styles["currency_cell"])

	profitMargin := totalProfit / totalRevenue
	summaryRow4 := sheet.AddRow()
	summaryRow4.AddCell("Profit Margin:").SetStyle(styles["summary_label"])
	summaryRow4.AddCell(profitMargin).SetStyle(styles["percentage_cell"])

	sheet.AddRow() // Empty row

	// Period breakdown headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Period").SetStyle(styles["table_header"])
	headerRow.AddCell("Revenue").SetStyle(styles["table_header"])
	headerRow.AddCell("Expenses").SetStyle(styles["table_header"])
	headerRow.AddCell("Profit").SetStyle(styles["table_header"])

	// Period data
	for _, financialData := range data {
		dataRow := sheet.AddRow()
		dataRow.AddCell(financialData.Period).SetStyle(styles["data_cell"])
		dataRow.AddCell(financialData.Revenue).SetStyle(styles["currency_cell"])
		dataRow.AddCell(financialData.Expenses).SetStyle(styles["currency_cell"])
		dataRow.AddCell(financialData.Profit).SetStyle(styles["currency_cell"])
	}
}

func createRevenueAnalysisSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, data []FinancialData) {
	// Set column widths
	sheet.SetColumnWidth("A", 15.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 15.0)
	sheet.SetColumnWidth("E", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Revenue Analysis").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Calculate revenue statistics
	var totalRevenue, maxRevenue, minRevenue float64
	if len(data) > 0 {
		maxRevenue = data[0].Revenue
		minRevenue = data[0].Revenue
	}

	for _, financialData := range data {
		totalRevenue += financialData.Revenue
		if financialData.Revenue > maxRevenue {
			maxRevenue = financialData.Revenue
		}
		if financialData.Revenue < minRevenue {
			minRevenue = financialData.Revenue
		}
	}

	averageRevenue := totalRevenue / float64(len(data))

	// Revenue statistics
	statsRow1 := sheet.AddRow()
	statsRow1.AddCell("Total Revenue:").SetStyle(styles["summary_label"])
	statsRow1.AddCell(totalRevenue).SetStyle(styles["currency_cell"])

	statsRow2 := sheet.AddRow()
	statsRow2.AddCell("Average Revenue:").SetStyle(styles["summary_label"])
	statsRow2.AddCell(averageRevenue).SetStyle(styles["currency_cell"])

	statsRow3 := sheet.AddRow()
	statsRow3.AddCell("Highest Revenue:").SetStyle(styles["summary_label"])
	statsRow3.AddCell(maxRevenue).SetStyle(styles["currency_cell"])

	statsRow4 := sheet.AddRow()
	statsRow4.AddCell("Lowest Revenue:").SetStyle(styles["summary_label"])
	statsRow4.AddCell(minRevenue).SetStyle(styles["currency_cell"])

	sheet.AddRow() // Empty row

	// Period revenue analysis headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Period").SetStyle(styles["table_header"])
	headerRow.AddCell("Revenue").SetStyle(styles["table_header"])
	headerRow.AddCell("vs Average").SetStyle(styles["table_header"])
	headerRow.AddCell("Growth %").SetStyle(styles["table_header"])
	headerRow.AddCell("Performance").SetStyle(styles["table_header"])

	// Period revenue data with analysis
	for i, financialData := range data {
		dataRow := sheet.AddRow()
		dataRow.AddCell(financialData.Period).SetStyle(styles["data_cell"])
		dataRow.AddCell(financialData.Revenue).SetStyle(styles["currency_cell"])

		// Variance from average
		variance := financialData.Revenue - averageRevenue
		dataRow.AddCell(variance).SetStyle(styles["currency_cell"])

		// Growth percentage (compared to previous period)
		var growth float64
		if i > 0 {
			prevRevenue := data[i-1].Revenue
			growth = (financialData.Revenue - prevRevenue) / prevRevenue
		}
		dataRow.AddCell(growth).SetStyle(styles["percentage_cell"])

		// Performance indicator
		var performance string
		var performanceStyle string
		if financialData.Revenue >= averageRevenue*1.1 {
			performance = "Above Average"
			performanceStyle = "status_active"
		} else if financialData.Revenue >= averageRevenue*0.9 {
			performance = "Average"
			performanceStyle = "status_warning"
		} else {
			performance = "Below Average"
			performanceStyle = "status_danger"
		}
		dataRow.AddCell(performance).SetStyle(styles[performanceStyle])
	}
}

func createExpenseAnalysisSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, data []FinancialData) {
	// Set column widths
	sheet.SetColumnWidth("A", 15.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 15.0)
	sheet.SetColumnWidth("E", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Expense Analysis").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Calculate expense statistics
	var totalExpenses, maxExpenses, minExpenses float64
	if len(data) > 0 {
		maxExpenses = data[0].Expenses
		minExpenses = data[0].Expenses
	}

	for _, financialData := range data {
		totalExpenses += financialData.Expenses
		if financialData.Expenses > maxExpenses {
			maxExpenses = financialData.Expenses
		}
		if financialData.Expenses < minExpenses {
			minExpenses = financialData.Expenses
		}
	}

	averageExpenses := totalExpenses / float64(len(data))

	// Expense statistics
	statsRow1 := sheet.AddRow()
	statsRow1.AddCell("Total Expenses:").SetStyle(styles["summary_label"])
	statsRow1.AddCell(totalExpenses).SetStyle(styles["currency_cell"])

	statsRow2 := sheet.AddRow()
	statsRow2.AddCell("Average Expenses:").SetStyle(styles["summary_label"])
	statsRow2.AddCell(averageExpenses).SetStyle(styles["currency_cell"])

	statsRow3 := sheet.AddRow()
	statsRow3.AddCell("Highest Expenses:").SetStyle(styles["summary_label"])
	statsRow3.AddCell(maxExpenses).SetStyle(styles["currency_cell"])

	statsRow4 := sheet.AddRow()
	statsRow4.AddCell("Lowest Expenses:").SetStyle(styles["summary_label"])
	statsRow4.AddCell(minExpenses).SetStyle(styles["currency_cell"])

	sheet.AddRow() // Empty row

	// Period expense analysis headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Period").SetStyle(styles["table_header"])
	headerRow.AddCell("Expenses").SetStyle(styles["table_header"])
	headerRow.AddCell("vs Average").SetStyle(styles["table_header"])
	headerRow.AddCell("Change %").SetStyle(styles["table_header"])
	headerRow.AddCell("Control Status").SetStyle(styles["table_header"])

	// Period expense data with analysis
	for i, financialData := range data {
		dataRow := sheet.AddRow()
		dataRow.AddCell(financialData.Period).SetStyle(styles["data_cell"])
		dataRow.AddCell(financialData.Expenses).SetStyle(styles["currency_cell"])

		// Variance from average
		variance := financialData.Expenses - averageExpenses
		dataRow.AddCell(variance).SetStyle(styles["currency_cell"])

		// Change percentage (compared to previous period)
		var change float64
		if i > 0 {
			prevExpenses := data[i-1].Expenses
			change = (financialData.Expenses - prevExpenses) / prevExpenses
		}
		dataRow.AddCell(change).SetStyle(styles["percentage_cell"])

		// Expense control status
		var controlStatus string
		var controlStyle string
		if financialData.Expenses <= averageExpenses*0.9 {
			controlStatus = "Well Controlled"
			controlStyle = "status_active"
		} else if financialData.Expenses <= averageExpenses*1.1 {
			controlStatus = "Normal Range"
			controlStyle = "status_warning"
		} else {
			controlStatus = "Above Target"
			controlStyle = "status_danger"
		}
		dataRow.AddCell(controlStatus).SetStyle(styles[controlStyle])
	}
}

func createInventoryOverviewSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, inventory []InventoryItem) {
	// Set column widths
	sheet.SetColumnWidth("A", 15.0)
	sheet.SetColumnWidth("B", 25.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 12.0)
	sheet.SetColumnWidth("E", 12.0)
	sheet.SetColumnWidth("F", 15.0)
	sheet.SetColumnWidth("G", 15.0)
	sheet.SetColumnWidth("H", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Inventory Management Overview").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Calculate inventory statistics
	var totalValue, totalQuantity float64
	categoryCount := make(map[string]int)

	for _, item := range inventory {
		totalValue += item.TotalValue
		totalQuantity += float64(item.Quantity)
		categoryCount[item.Category]++
		// Note: StockLevel not available in current struct
	}

	// Summary section
	summaryRow1 := sheet.AddRow()
	summaryRow1.AddCell("Total Items:").SetStyle(styles["summary_label"])
	summaryRow1.AddCell(len(inventory)).SetStyle(styles["summary_value"])

	summaryRow2 := sheet.AddRow()
	summaryRow2.AddCell("Total Quantity:").SetStyle(styles["summary_label"])
	summaryRow2.AddCell(totalQuantity).SetStyle(styles["numeric_cell"])

	summaryRow3 := sheet.AddRow()
	summaryRow3.AddCell("Total Value:").SetStyle(styles["summary_label"])
	summaryRow3.AddCell(totalValue).SetStyle(styles["currency_cell"])

	averageValue := totalValue / float64(len(inventory))
	summaryRow4 := sheet.AddRow()
	summaryRow4.AddCell("Average Item Value:").SetStyle(styles["summary_label"])
	summaryRow4.AddCell(averageValue).SetStyle(styles["currency_cell"])

	sheet.AddRow() // Empty row

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("SKU").SetStyle(styles["table_header"])
	headerRow.AddCell("Item Name").SetStyle(styles["table_header"])
	headerRow.AddCell("Category").SetStyle(styles["table_header"])
	headerRow.AddCell("Quantity").SetStyle(styles["table_header"])
	headerRow.AddCell("Unit Price").SetStyle(styles["table_header"])
	headerRow.AddCell("Total Value").SetStyle(styles["table_header"])
	headerRow.AddCell("Supplier").SetStyle(styles["table_header"])
	headerRow.AddCell("Last Updated").SetStyle(styles["table_header"])

	// Inventory data
	for _, item := range inventory {
		dataRow := sheet.AddRow()
		dataRow.AddCell(item.SKU).SetStyle(styles["data_cell"])
		dataRow.AddCell(item.Name).SetStyle(styles["data_cell"])
		dataRow.AddCell(item.Category).SetStyle(styles["data_cell"])
		dataRow.AddCell(item.Quantity).SetStyle(styles["numeric_cell"])
		dataRow.AddCell(item.UnitPrice).SetStyle(styles["currency_cell"])
		dataRow.AddCell(item.TotalValue).SetStyle(styles["currency_cell"])
		dataRow.AddCell(item.Supplier).SetStyle(styles["data_cell"])
		dataRow.AddCell(item.LastUpdated).SetStyle(styles["date_cell"])
	}
}

func createInventoryCategorySheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, inventory []InventoryItem) {
	// Set column widths
	sheet.SetColumnWidth("A", 20.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 15.0)
	sheet.SetColumnWidth("E", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Category Analysis").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Calculate category statistics
	categoryStats := make(map[string]struct {
		Count    int
		Quantity int
		Value    float64
	})

	for _, item := range inventory {
		stats := categoryStats[item.Category]
		stats.Count++
		stats.Quantity += item.Quantity
		stats.Value += item.TotalValue
		categoryStats[item.Category] = stats
	}

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Category").SetStyle(styles["table_header"])
	headerRow.AddCell("Item Count").SetStyle(styles["table_header"])
	headerRow.AddCell("Total Quantity").SetStyle(styles["table_header"])
	headerRow.AddCell("Total Value").SetStyle(styles["table_header"])
	headerRow.AddCell("Average Value").SetStyle(styles["table_header"])

	// Category data
	for category, stats := range categoryStats {
		dataRow := sheet.AddRow()
		dataRow.AddCell(category).SetStyle(styles["data_cell"])
		dataRow.AddCell(stats.Count).SetStyle(styles["numeric_cell"])
		dataRow.AddCell(stats.Quantity).SetStyle(styles["numeric_cell"])
		dataRow.AddCell(stats.Value).SetStyle(styles["currency_cell"])
		averageValue := stats.Value / float64(stats.Count)
		dataRow.AddCell(averageValue).SetStyle(styles["currency_cell"])
	}

	sheet.AddRow() // Empty row

	// Detailed breakdown by category
	for category := range categoryStats {
		sheet.AddRow() // Empty row
		sectionRow := sheet.AddRow()
		sectionRow.AddCell(fmt.Sprintf("%s Items", category)).SetStyle(styles["section_header"])

		detailHeaderRow := sheet.AddRow()
		detailHeaderRow.AddCell("Item Name").SetStyle(styles["table_header"])
		detailHeaderRow.AddCell("Quantity").SetStyle(styles["table_header"])
		detailHeaderRow.AddCell("Unit Price").SetStyle(styles["table_header"])
		detailHeaderRow.AddCell("Total Value").SetStyle(styles["table_header"])
		detailHeaderRow.AddCell("Supplier").SetStyle(styles["table_header"])

		for _, item := range inventory {
			if item.Category == category {
				detailRow := sheet.AddRow()
				detailRow.AddCell(item.Name).SetStyle(styles["data_cell"])
				detailRow.AddCell(item.Quantity).SetStyle(styles["numeric_cell"])
				detailRow.AddCell(item.UnitPrice).SetStyle(styles["currency_cell"])
				detailRow.AddCell(item.TotalValue).SetStyle(styles["currency_cell"])
				detailRow.AddCell(item.Supplier).SetStyle(styles["data_cell"])
			}
		}
	}
}

func createStockLevelsSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, inventory []InventoryItem) {
	// Set column widths
	sheet.SetColumnWidth("A", 20.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Stock Levels Analysis").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Calculate stock level statistics (simulated based on quantity)
	stockLevelStats := make(map[string]struct {
		Count int
		Value float64
	})

	for _, item := range inventory {
		// Simulate stock level based on quantity
		var stockLevel string
		if item.Quantity > 100 {
			stockLevel = "High"
		} else if item.Quantity > 50 {
			stockLevel = "Medium"
		} else {
			stockLevel = "Low"
		}
		
		stats := stockLevelStats[stockLevel]
		stats.Count++
		stats.Value += item.TotalValue
		stockLevelStats[stockLevel] = stats
	}

	// Stock level summary headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Stock Level").SetStyle(styles["table_header"])
	headerRow.AddCell("Item Count").SetStyle(styles["table_header"])
	headerRow.AddCell("Total Value").SetStyle(styles["table_header"])
	headerRow.AddCell("Percentage").SetStyle(styles["table_header"])

	// Stock level summary data
	totalItems := len(inventory)
	for level, stats := range stockLevelStats {
		dataRow := sheet.AddRow()
		var levelStyle string
		switch level {
		case "High":
			levelStyle = "status_active"
		case "Medium":
			levelStyle = "status_warning"
		case "Low":
			levelStyle = "status_danger"
		default:
			levelStyle = "data_cell"
		}
		dataRow.AddCell(level).SetStyle(styles[levelStyle])
		dataRow.AddCell(stats.Count).SetStyle(styles["numeric_cell"])
		dataRow.AddCell(stats.Value).SetStyle(styles["currency_cell"])
		percentage := float64(stats.Count) / float64(totalItems)
		dataRow.AddCell(percentage).SetStyle(styles["percentage_cell"])
	}

	sheet.AddRow() // Empty row

	// Low stock items alert
	alertRow := sheet.AddRow()
	alertRow.AddCell("LOW STOCK ALERT").SetStyle(styles["section_header"])

	lowStockHeaderRow := sheet.AddRow()
	lowStockHeaderRow.AddCell("Item Name").SetStyle(styles["table_header"])
	lowStockHeaderRow.AddCell("Category").SetStyle(styles["table_header"])
	lowStockHeaderRow.AddCell("Quantity").SetStyle(styles["table_header"])
	lowStockHeaderRow.AddCell("Value").SetStyle(styles["table_header"])

	for _, item := range inventory {
		if item.Quantity <= 50 { // Low stock threshold
			lowStockRow := sheet.AddRow()
			lowStockRow.AddCell(item.Name).SetStyle(styles["data_cell"])
			lowStockRow.AddCell(item.Category).SetStyle(styles["data_cell"])
			lowStockRow.AddCell(item.Quantity).SetStyle(styles["status_danger"])
			lowStockRow.AddCell(item.TotalValue).SetStyle(styles["currency_cell"])
		}
	}
}

func createTemplateComparisonSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, templates map[string]TemplateConfig) {
	// Template comparison implementation
	titleRow := sheet.AddRow()
	titleRow.AddCell("Template System Comparison").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Template Name").SetStyle(styles["table_header"])
	headerRow.AddCell("Category").SetStyle(styles["table_header"])
	headerRow.AddCell("Description").SetStyle(styles["table_header"])
	headerRow.AddCell("Variables").SetStyle(styles["table_header"])

	// Template data
	for _, template := range templates {
		dataRow := sheet.AddRow()
		dataRow.AddCell(template.Name).SetStyle(styles["data_cell"])
		dataRow.AddCell(template.Category).SetStyle(styles["data_cell"])
		dataRow.AddCell(template.Description).SetStyle(styles["data_cell"])
		
		// Convert variables to string
		varList := make([]string, 0, len(template.Variables))
		for key := range template.Variables {
			varList = append(varList, key)
		}
		dataRow.AddCell(strings.Join(varList, ", ")).SetStyle(styles["data_cell"])
	}
}

func createStyleVariationsSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig) {
	// Set column widths
	sheet.SetColumnWidth("A", 25.0)
	sheet.SetColumnWidth("B", 20.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Style Variations Demo").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Style demonstration section
	sectionRow1 := sheet.AddRow()
	sectionRow1.AddCell("Header Styles").SetStyle(styles["section_header"])

	headerDemoRow := sheet.AddRow()
	headerDemoRow.AddCell("Template Title Style").SetStyle(styles["template_title"])
	headerDemoRow.AddCell("Section Header Style").SetStyle(styles["section_header"])
	headerDemoRow.AddCell("Table Header Style").SetStyle(styles["table_header"])

	sheet.AddRow() // Empty row

	// Data cell styles
	sectionRow2 := sheet.AddRow()
	sectionRow2.AddCell("Data Cell Styles").SetStyle(styles["section_header"])

	dataDemoRow := sheet.AddRow()
	dataDemoRow.AddCell("Regular Data Cell").SetStyle(styles["data_cell"])
	dataDemoRow.AddCell("Numeric Cell").SetStyle(styles["numeric_cell"])
	dataDemoRow.AddCell("Currency Cell").SetStyle(styles["currency_cell"])
	dataDemoRow.AddCell("Date Cell").SetStyle(styles["date_cell"])

	percentageDemoRow := sheet.AddRow()
	percentageDemoRow.AddCell("Percentage Cell").SetStyle(styles["percentage_cell"])

	sheet.AddRow() // Empty row

	// Status indicators
	sectionRow3 := sheet.AddRow()
	sectionRow3.AddCell("Status Indicators").SetStyle(styles["section_header"])

	statusDemoRow := sheet.AddRow()
	statusDemoRow.AddCell("Active Status").SetStyle(styles["status_active"])
	statusDemoRow.AddCell("Warning Status").SetStyle(styles["status_warning"])
	statusDemoRow.AddCell("Danger Status").SetStyle(styles["status_danger"])

	sheet.AddRow() // Empty row

	// Summary styles
	sectionRow4 := sheet.AddRow()
	sectionRow4.AddCell("Summary Styles").SetStyle(styles["section_header"])

	summaryDemoRow := sheet.AddRow()
	summaryDemoRow.AddCell("Summary Label").SetStyle(styles["summary_label"])
	summaryDemoRow.AddCell("Summary Value").SetStyle(styles["summary_value"])

	sheet.AddRow() // Empty row

	// Style properties table
	propsHeaderRow := sheet.AddRow()
	propsHeaderRow.AddCell("Style Properties Overview").SetStyle(styles["section_header"])

	propsTableHeaderRow := sheet.AddRow()
	propsTableHeaderRow.AddCell("Style Name").SetStyle(styles["table_header"])
	propsTableHeaderRow.AddCell("Usage").SetStyle(styles["table_header"])
	propsTableHeaderRow.AddCell("Sample Text").SetStyle(styles["table_header"])
	propsTableHeaderRow.AddCell("Notes").SetStyle(styles["table_header"])

	// Style properties data
	styleInfo := []struct {
		Name  string
		Usage string
		Style string
		Notes string
	}{
		{"template_title", "Main report titles", "template_title", "Large, bold, centered"},
		{"section_header", "Section headings", "section_header", "Medium, bold, colored background"},
		{"table_header", "Table column headers", "table_header", "Bold, colored background"},
		{"data_cell", "Regular data", "data_cell", "Standard text formatting"},
		{"currency_cell", "Financial values", "currency_cell", "Currency formatting"},
		{"percentage_cell", "Percentage values", "percentage_cell", "Percentage formatting"},
		{"status_active", "Positive status", "status_active", "Green background"},
		{"status_warning", "Warning status", "status_warning", "Yellow background"},
		{"status_danger", "Critical status", "status_danger", "Red background"},
	}

	for _, info := range styleInfo {
		propsRow := sheet.AddRow()
		propsRow.AddCell(info.Name).SetStyle(styles["data_cell"])
		propsRow.AddCell(info.Usage).SetStyle(styles["data_cell"])
		propsRow.AddCell("Sample").SetStyle(styles[info.Style])
		propsRow.AddCell(info.Notes).SetStyle(styles["data_cell"])
	}
}

func createDynamicContentSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig) {
	// Set column widths
	sheet.SetColumnWidth("A", 25.0)
	sheet.SetColumnWidth("B", 20.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 20.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Dynamic Content Examples").SetStyle(styles["template_title"])

	sheet.AddRow() // Empty row

	// Template variables section
	variablesSection := sheet.AddRow()
	variablesSection.AddCell("Template Variables").SetStyle(styles["section_header"])

	variablesHeaderRow := sheet.AddRow()
	variablesHeaderRow.AddCell("Variable Name").SetStyle(styles["table_header"])
	variablesHeaderRow.AddCell("Description").SetStyle(styles["table_header"])
	variablesHeaderRow.AddCell("Type").SetStyle(styles["table_header"])
	variablesHeaderRow.AddCell("Example Value").SetStyle(styles["table_header"])

	// Sample template variables
	variables := []struct {
		Name        string
		Description string
		Type        string
		Example     string
	}{
		{"{{.CompanyName}}", "Company name for headers", "String", "Acme Corporation"},
		{"{{.ReportDate}}", "Report generation date", "Date", "2024-01-15"},
		{"{{.TotalRevenue}}", "Total revenue amount", "Currency", "$1,250,000"},
		{"{{.EmployeeCount}}", "Number of employees", "Integer", "150"},
		{"{{.GrowthRate}}", "Year-over-year growth", "Percentage", "15.5%"},
		{"{{.Department}}", "Department name", "String", "Engineering"},
		{"{{.ProjectStatus}}", "Current project status", "Enum", "In Progress"},
		{"{{.BudgetUtilization}}", "Budget usage percentage", "Percentage", "87.3%"},
	}

	for _, variable := range variables {
		variableRow := sheet.AddRow()
		variableRow.AddCell(variable.Name).SetStyle(styles["data_cell"])
		variableRow.AddCell(variable.Description).SetStyle(styles["data_cell"])
		variableRow.AddCell(variable.Type).SetStyle(styles["data_cell"])
		variableRow.AddCell(variable.Example).SetStyle(styles["data_cell"])
	}

	sheet.AddRow() // Empty row

	// Dynamic formatting section
	formattingSection := sheet.AddRow()
	formattingSection.AddCell("Dynamic Formatting Examples").SetStyle(styles["section_header"])

	formattingHeaderRow := sheet.AddRow()
	formattingHeaderRow.AddCell("Condition").SetStyle(styles["table_header"])
	formattingHeaderRow.AddCell("Applied Style").SetStyle(styles["table_header"])
	formattingHeaderRow.AddCell("Sample Value").SetStyle(styles["table_header"])
	formattingHeaderRow.AddCell("Result").SetStyle(styles["table_header"])

	// Conditional formatting examples
	formattingExamples := []struct {
		Condition string
		Style     string
		Value     string
		Result    string
	}{
		{"Revenue > Target", "status_active", "$1,250,000", "Above Target"},
		{"Budget 70-90% Used", "status_warning", "87.3%", "Monitor Usage"},
		{"Stock Level Low", "status_danger", "5 units", "Reorder Required"},
		{"Project Completed", "status_active", "100%", "Completed"},
		{"Employee Performance", "status_warning", "Needs Review", "Action Required"},
	}

	for _, example := range formattingExamples {
		formattingRow := sheet.AddRow()
		formattingRow.AddCell(example.Condition).SetStyle(styles["data_cell"])
		formattingRow.AddCell(example.Style).SetStyle(styles["data_cell"])
		formattingRow.AddCell(example.Value).SetStyle(styles["data_cell"])

		// Apply conditional styling based on the example
		var resultStyle string
		switch example.Style {
		case "status_active":
			resultStyle = "status_active"
		case "status_warning":
			resultStyle = "status_warning"
		case "status_danger":
			resultStyle = "status_danger"
		default:
			resultStyle = "data_cell"
		}
		formattingRow.AddCell(example.Result).SetStyle(styles[resultStyle])
	}

	sheet.AddRow() // Empty row

	// Template generation statistics
	statsSection := sheet.AddRow()
	statsSection.AddCell("Template Generation Statistics").SetStyle(styles["section_header"])

	statsRow1 := sheet.AddRow()
	statsRow1.AddCell("Total Templates Available:").SetStyle(styles["summary_label"])
	statsRow1.AddCell("5").SetStyle(styles["summary_value"])

	statsRow2 := sheet.AddRow()
	statsRow2.AddCell("Style Variations:").SetStyle(styles["summary_label"])
	statsRow2.AddCell("12").SetStyle(styles["summary_value"])

	statsRow3 := sheet.AddRow()
	statsRow3.AddCell("Dynamic Variables:").SetStyle(styles["summary_label"])
	statsRow3.AddCell("25+").SetStyle(styles["summary_value"])

	statsRow4 := sheet.AddRow()
	statsRow4.AddCell("Conditional Formats:").SetStyle(styles["summary_label"])
	statsRow4.AddCell("8").SetStyle(styles["summary_value"])
}