package main

import (
	"fmt"
	"math/rand"
	"os"
	"strings"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Template data structures
type TemplateData struct {
	Company     CompanyInfo
	Report      ReportInfo
	Products    []Product
	Customers   []Customer
	Invoices    []Invoice
	Metrics     DashboardMetrics
	Certificate CertificateInfo
	Form        FormData
}

type CompanyInfo struct {
	Name        string
	Address     string
	Phone       string
	Email       string
	Website     string
	Logo        string
	TaxID       string
	Registration string
}

type ReportInfo struct {
	Title       string
	Date        time.Time
	Period      string
	Author      string
	Department  string
	Version     string
	Description string
}

type Product struct {
	ID          string
	Name        string
	Category    string
	Price       float64
	Quantity    int
	Description string
	Supplier    string
	Status      string
}

type Customer struct {
	ID       string
	Name     string
	Email    string
	Phone    string
	Address  string
	City     string
	Country  string
	Segment  string
	Value    float64
	JoinDate time.Time
}

type Invoice struct {
	ID          string
	CustomerID  string
	Date        time.Time
	DueDate     time.Time
	Items       []InvoiceItem
	Subtotal    float64
	Tax         float64
	Discount    float64
	Total       float64
	Status      string
	Notes       string
}

type InvoiceItem struct {
	ProductID   string
	Description string
	Quantity    int
	UnitPrice   float64
	Total       float64
}

type DashboardMetrics struct {
	TotalSales     float64
	TotalOrders    int
	ActiveCustomers int
	ConversionRate float64
	GrowthRate     float64
	TopProducts    []ProductMetric
	SalesTrend     []SalesData
	RegionalData   []RegionData
}

type ProductMetric struct {
	Name     string
	Sales    float64
	Quantity int
	Growth   float64
}

type SalesData struct {
	Date   time.Time
	Amount float64
	Orders int
}

type RegionData struct {
	Region string
	Sales  float64
	Share  float64
}

type CertificateInfo struct {
	RecipientName string
	CourseName    string
	CompletionDate time.Time
	Instructor    string
	Credits       int
	Grade         string
	CertificateID string
	ValidUntil    time.Time
}

type FormData struct {
	Title       string
	Description string
	Fields      []FormField
	Submissions []FormSubmission
}

type FormField struct {
	Name        string
	Type        string
	Label       string
	Required    bool
	Options     []string
	Validation  string
}

type FormSubmission struct {
	ID          string
	SubmittedAt time.Time
	Data        map[string]interface{}
	Status      string
}

// Template engine interface
type TemplateEngine struct {
	templates map[string]*Template
	helpers   map[string]TemplateHelper
}

type Template struct {
	Name     string
	Content  string
	Vars     []string
	Sections []TemplateSection
}

type TemplateSection struct {
	Name      string
	Type      string // "static", "variable", "loop", "conditional"
	Condition string
	Data      interface{}
}

type TemplateHelper func(args ...interface{}) string

func main() {
	fmt.Println("Template System Example")
	fmt.Println("======================")

	// Create output directory
	if err := os.MkdirAll("output", 0755); err != nil {
		fmt.Printf("Error creating output directory: %v\n", err)
		return
	}

	// Initialize template engine
	engine := NewTemplateEngine()
	registerTemplateHelpers(engine)

	// Generate template data
	templateData := generateTemplateData()

	// Generate different template examples
	fmt.Println("Generating Report Templates...")
	if err := generateReportTemplates(engine, templateData); err != nil {
		fmt.Printf("Error generating report templates: %v\n", err)
	} else {
		fmt.Println("✓ Report Templates generated")
	}

	fmt.Println("Generating Invoice Templates...")
	if err := generateInvoiceTemplates(engine, templateData); err != nil {
		fmt.Printf("Error generating invoice templates: %v\n", err)
	} else {
		fmt.Println("✓ Invoice Templates generated")
	}

	fmt.Println("Generating Dashboard Templates...")
	if err := generateDashboardTemplates(engine, templateData); err != nil {
		fmt.Printf("Error generating dashboard templates: %v\n", err)
	} else {
		fmt.Println("✓ Dashboard Templates generated")
	}

	fmt.Println("Generating Form Templates...")
	if err := generateFormTemplates(engine, templateData); err != nil {
		fmt.Printf("Error generating form templates: %v\n", err)
	} else {
		fmt.Println("✓ Form Templates generated")
	}

	fmt.Println("Generating Certificate Templates...")
	if err := generateCertificateTemplates(engine, templateData); err != nil {
		fmt.Printf("Error generating certificate templates: %v\n", err)
	} else {
		fmt.Println("✓ Certificate Templates generated")
	}

	fmt.Println("\nTemplate system examples completed!")
	fmt.Println("Check the output directory for generated files.")
}

// Template Engine Implementation
func NewTemplateEngine() *TemplateEngine {
	return &TemplateEngine{
		templates: make(map[string]*Template),
		helpers:   make(map[string]TemplateHelper),
	}
}

func (te *TemplateEngine) RegisterTemplate(name, content string) {
	te.templates[name] = &Template{
		Name:    name,
		Content: content,
	}
}

func (te *TemplateEngine) RegisterHelper(name string, helper TemplateHelper) {
	te.helpers[name] = helper
}

func (te *TemplateEngine) Render(templateName string, data interface{}) string {
	template, exists := te.templates[templateName]
	if !exists {
		return fmt.Sprintf("Template '%s' not found", templateName)
	}

	// Simple template rendering (in real implementation, use proper template engine)
	result := template.Content
	
	// Replace variables with actual data
	result = te.processVariables(result, data)
	
	return result
}

func (te *TemplateEngine) processVariables(content string, data interface{}) string {
	// Simple variable substitution (in real implementation, use proper parsing)
	result := content
	
	// This is a simplified example - real implementation would use proper template parsing
	if templateData, ok := data.(*TemplateData); ok {
		result = strings.ReplaceAll(result, "{{company_name}}", templateData.Company.Name)
		result = strings.ReplaceAll(result, "{{report_title}}", templateData.Report.Title)
		result = strings.ReplaceAll(result, "{{report_date}}", templateData.Report.Date.Format("2006-01-02"))
		result = strings.ReplaceAll(result, "{{total_sales}}", fmt.Sprintf("%.2f", templateData.Metrics.TotalSales))
	}
	
	return result
}

func registerTemplateHelpers(engine *TemplateEngine) {
	// Date formatting helper
	engine.RegisterHelper("formatDate", func(args ...interface{}) string {
		if len(args) >= 2 {
			if date, ok := args[0].(time.Time); ok {
				if format, ok := args[1].(string); ok {
					return date.Format(format)
				}
			}
		}
		return ""
	})

	// Currency formatting helper
	engine.RegisterHelper("formatCurrency", func(args ...interface{}) string {
		if len(args) >= 2 {
			if amount, ok := args[0].(float64); ok {
				if currency, ok := args[1].(string); ok {
					return fmt.Sprintf("%s %.2f", currency, amount)
				}
			}
		}
		return ""
	})

	// Percentage formatting helper
	engine.RegisterHelper("formatPercent", func(args ...interface{}) string {
		if len(args) >= 1 {
			if value, ok := args[0].(float64); ok {
				return fmt.Sprintf("%.1f%%", value*100)
			}
		}
		return ""
	})

	// Number formatting helper
	engine.RegisterHelper("formatNumber", func(args ...interface{}) string {
		if len(args) >= 1 {
			if value, ok := args[0].(float64); ok {
				return fmt.Sprintf("%.0f", value)
			}
		}
		return ""
	})
}

func generateTemplateData() *TemplateData {
	return &TemplateData{
		Company: CompanyInfo{
			Name:        "TechCorp Solutions",
			Address:     "123 Business Ave, Tech City, TC 12345",
			Phone:       "+1 (555) 123-4567",
			Email:       "info@techcorp.com",
			Website:     "www.techcorp.com",
			Logo:        "logo.png",
			TaxID:       "TAX123456789",
			Registration: "REG987654321",
		},
		Report: ReportInfo{
			Title:       "Quarterly Business Report",
			Date:        time.Now(),
			Period:      "Q4 2024",
			Author:      "John Smith",
			Department:  "Analytics",
			Version:     "1.0",
			Description: "Comprehensive quarterly performance analysis",
		},
		Products:    generateProducts(),
		Customers:   generateCustomers(),
		Invoices:    generateInvoices(),
		Metrics:     generateDashboardMetrics(),
		Certificate: generateCertificateInfo(),
		Form:        generateFormData(),
	}
}

func generateProducts() []Product {
	products := []Product{
		{"P001", "Laptop Pro", "Electronics", 1299.99, 50, "High-performance laptop", "TechSupplier", "Active"},
		{"P002", "Wireless Mouse", "Accessories", 29.99, 200, "Ergonomic wireless mouse", "AccessoryCorp", "Active"},
		{"P003", "Monitor 4K", "Electronics", 399.99, 30, "Ultra HD 4K monitor", "DisplayTech", "Active"},
		{"P004", "Keyboard Mechanical", "Accessories", 89.99, 75, "RGB mechanical keyboard", "KeyboardInc", "Active"},
		{"P005", "Webcam HD", "Electronics", 79.99, 100, "1080p HD webcam", "CameraTech", "Active"},
	}

	// Generate additional random products
	for i := 6; i <= 20; i++ {
		product := Product{
			ID:          fmt.Sprintf("P%03d", i),
			Name:        fmt.Sprintf("Product %d", i),
			Category:    []string{"Electronics", "Accessories", "Software", "Hardware"}[rand.Intn(4)],
			Price:       float64(rand.Intn(1000)) + 9.99,
			Quantity:    rand.Intn(200) + 10,
			Description: fmt.Sprintf("Description for product %d", i),
			Supplier:    fmt.Sprintf("Supplier%d", rand.Intn(5)+1),
			Status:      []string{"Active", "Discontinued", "Coming Soon"}[rand.Intn(3)],
		}
		products = append(products, product)
	}

	return products
}

func generateCustomers() []Customer {
	customers := []Customer{
		{"C001", "Alice Johnson", "alice@email.com", "+1-555-0101", "123 Main St", "New York", "USA", "Premium", 15000.0, time.Now().AddDate(-2, 0, 0)},
		{"C002", "Bob Smith", "bob@email.com", "+1-555-0102", "456 Oak Ave", "Los Angeles", "USA", "Standard", 8500.0, time.Now().AddDate(-1, -6, 0)},
		{"C003", "Carol Davis", "carol@email.com", "+1-555-0103", "789 Pine Rd", "Chicago", "USA", "Premium", 22000.0, time.Now().AddDate(-3, 0, 0)},
		{"C004", "David Wilson", "david@email.com", "+1-555-0104", "321 Elm St", "Houston", "USA", "Basic", 3200.0, time.Now().AddDate(0, -3, 0)},
		{"C005", "Eva Brown", "eva@email.com", "+1-555-0105", "654 Maple Dr", "Phoenix", "USA", "Standard", 12000.0, time.Now().AddDate(-1, 0, 0)},
	}

	// Generate additional customers
	for i := 6; i <= 25; i++ {
		customer := Customer{
			ID:       fmt.Sprintf("C%03d", i),
			Name:     fmt.Sprintf("Customer %d", i),
			Email:    fmt.Sprintf("customer%d@email.com", i),
			Phone:    fmt.Sprintf("+1-555-%04d", rand.Intn(10000)),
			Address:  fmt.Sprintf("%d Random St", rand.Intn(9999)+1),
			City:     []string{"New York", "Los Angeles", "Chicago", "Houston", "Phoenix"}[rand.Intn(5)],
			Country:  "USA",
			Segment:  []string{"Basic", "Standard", "Premium"}[rand.Intn(3)],
			Value:    float64(rand.Intn(25000)) + 1000.0,
			JoinDate: time.Now().AddDate(-rand.Intn(3), -rand.Intn(12), -rand.Intn(30)),
		}
		customers = append(customers, customer)
	}

	return customers
}

func generateInvoices() []Invoice {
	invoices := make([]Invoice, 10)
	for i := 0; i < 10; i++ {
		items := []InvoiceItem{
			{"P001", "Laptop Pro", 1, 1299.99, 1299.99},
			{"P002", "Wireless Mouse", 2, 29.99, 59.98},
		}

		subtotal := 1359.97
		tax := subtotal * 0.08
		discount := 0.0
		if i%3 == 0 {
			discount = subtotal * 0.1 // 10% discount
		}
		total := subtotal + tax - discount

		invoices[i] = Invoice{
			ID:         fmt.Sprintf("INV%04d", i+1),
			CustomerID: fmt.Sprintf("C%03d", (i%5)+1),
			Date:       time.Now().AddDate(0, 0, -rand.Intn(30)),
			DueDate:    time.Now().AddDate(0, 0, rand.Intn(30)+1),
			Items:      items,
			Subtotal:   subtotal,
			Tax:        tax,
			Discount:   discount,
			Total:      total,
			Status:     []string{"Paid", "Pending", "Overdue"}[rand.Intn(3)],
			Notes:      "Thank you for your business!",
		}
	}

	return invoices
}

func generateDashboardMetrics() DashboardMetrics {
	return DashboardMetrics{
		TotalSales:      125000.50,
		TotalOrders:     450,
		ActiveCustomers: 125,
		ConversionRate:  0.035,
		GrowthRate:      0.15,
		TopProducts: []ProductMetric{
			{"Laptop Pro", 45000.0, 35, 0.12},
			{"Monitor 4K", 28000.0, 70, 0.08},
			{"Keyboard Mechanical", 15000.0, 167, 0.25},
		},
		SalesTrend: generateSalesTrend(),
		RegionalData: []RegionData{
			{"North America", 75000.0, 0.60},
			{"Europe", 30000.0, 0.24},
			{"Asia Pacific", 20000.0, 0.16},
		},
	}
}

func generateSalesTrend() []SalesData {
	trend := make([]SalesData, 30)
	baseAmount := 3000.0

	for i := 0; i < 30; i++ {
		variation := (rand.Float64() - 0.5) * 1000 // ±500 variation
		amount := baseAmount + variation
		orders := int(amount/100) + rand.Intn(10)

		trend[i] = SalesData{
			Date:   time.Now().AddDate(0, 0, -29+i),
			Amount: amount,
			Orders: orders,
		}
	}

	return trend
}

func generateCertificateInfo() CertificateInfo {
	return CertificateInfo{
		RecipientName:  "John Doe",
		CourseName:     "Advanced Excel Analytics",
		CompletionDate: time.Now(),
		Instructor:     "Dr. Sarah Johnson",
		Credits:        40,
		Grade:          "A+",
		CertificateID:  "CERT-2024-001",
		ValidUntil:     time.Now().AddDate(2, 0, 0),
	}
}

func generateFormData() FormData {
	return FormData{
		Title:       "Customer Feedback Form",
		Description: "Please provide your feedback about our products and services",
		Fields: []FormField{
			{"name", "text", "Full Name", true, nil, "required"},
			{"email", "email", "Email Address", true, nil, "email"},
			{"rating", "select", "Overall Rating", true, []string{"Excellent", "Good", "Average", "Poor"}, "required"},
			{"comments", "textarea", "Additional Comments", false, nil, ""},
			{"recommend", "radio", "Would you recommend us?", true, []string{"Yes", "No", "Maybe"}, "required"},
		},
		Submissions: generateFormSubmissions(),
	}
}

func generateFormSubmissions() []FormSubmission {
	submissions := make([]FormSubmission, 5)
	for i := 0; i < 5; i++ {
		submissions[i] = FormSubmission{
			ID:          fmt.Sprintf("SUB%03d", i+1),
			SubmittedAt: time.Now().AddDate(0, 0, -rand.Intn(30)),
			Data: map[string]interface{}{
				"name":      fmt.Sprintf("User %d", i+1),
				"email":     fmt.Sprintf("user%d@email.com", i+1),
				"rating":    []string{"Excellent", "Good", "Average"}[rand.Intn(3)],
				"comments":  "Great service and products!",
				"recommend": []string{"Yes", "No", "Maybe"}[rand.Intn(3)],
			},
			Status: "Completed",
		}
	}
	return submissions
}

// Template generation functions
func generateReportTemplates(engine *TemplateEngine, data *TemplateData) error {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Sales Report Template
	salesSheet := workbook.AddSheet("Sales Report")
	createReportTemplate(salesSheet, "SALES PERFORMANCE REPORT", data)

	// Product Report Template
	productSheet := workbook.AddSheet("Product Report")
	createProductReportTemplate(productSheet, data)

	// Customer Report Template
	customerSheet := workbook.AddSheet("Customer Report")
	createCustomerReportTemplate(customerSheet, data)

	// Executive Summary Template
	execSheet := workbook.AddSheet("Executive Summary")
	createExecutiveSummaryTemplate(execSheet, data)

	file := workbook.Build()
	if file == nil {
		return fmt.Errorf("failed to build workbook")
	}
	return file.SaveAs("output/08-template-reports.xlsx")
}

func generateInvoiceTemplates(engine *TemplateEngine, data *TemplateData) error {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Standard Invoice Template
	standardSheet := workbook.AddSheet("Standard Invoice")
	createStandardInvoiceTemplate(standardSheet, data)

	// Professional Invoice Template
	professionalSheet := workbook.AddSheet("Professional Invoice")
	createProfessionalInvoiceTemplate(professionalSheet, data)

	// Service Invoice Template
	serviceSheet := workbook.AddSheet("Service Invoice")
	createServiceInvoiceTemplate(serviceSheet, data)

	file := workbook.Build()
	if file == nil {
		return fmt.Errorf("failed to build workbook")
	}
	return file.SaveAs("output/08-template-invoices.xlsx")
}

func generateDashboardTemplates(engine *TemplateEngine, data *TemplateData) error {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// KPI Dashboard Template
	kpiSheet := workbook.AddSheet("KPI Dashboard")
	createKPIDashboardTemplate(kpiSheet, data)

	// Sales Dashboard Template
	salesSheet := workbook.AddSheet("Sales Dashboard")
	createSalesDashboardTemplate(salesSheet, data)

	// Analytics Dashboard Template
	analyticsSheet := workbook.AddSheet("Analytics Dashboard")
	createAnalyticsDashboardTemplate(analyticsSheet, data)

	file := workbook.Build()
	if file == nil {
		return fmt.Errorf("failed to build workbook")
	}
	return file.SaveAs("output/08-template-dashboards.xlsx")
}

func generateFormTemplates(engine *TemplateEngine, data *TemplateData) error {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Survey Form Template
	surveySheet := workbook.AddSheet("Survey Form")
	createSurveyFormTemplate(surveySheet, data)

	// Registration Form Template
	registrationSheet := workbook.AddSheet("Registration Form")
	createRegistrationFormTemplate(registrationSheet, data)

	// Feedback Form Template
	feedbackSheet := workbook.AddSheet("Feedback Form")
	createFeedbackFormTemplate(feedbackSheet, data)

	file := workbook.Build()
	if file == nil {
		return fmt.Errorf("failed to build workbook")
	}
	return file.SaveAs("output/08-template-forms.xlsx")
}

func generateCertificateTemplates(engine *TemplateEngine, data *TemplateData) error {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Achievement Certificate Template
	achievementSheet := workbook.AddSheet("Achievement Certificate")
	createAchievementCertificateTemplate(achievementSheet, data)

	// Course Completion Template
	courseSheet := workbook.AddSheet("Course Completion")
	createCourseCompletionTemplate(courseSheet, data)

	// Award Certificate Template
	awardSheet := workbook.AddSheet("Award Certificate")
	createAwardCertificateTemplate(awardSheet, data)

	file := workbook.Build()
	if file == nil {
		return fmt.Errorf("failed to build workbook")
	}
	return file.SaveAs("output/08-template-certificates.xlsx")
}

// Template creation functions
func createReportTemplate(sheet *excelbuilder.SheetBuilder, title string, data *TemplateData) {
	// Header
	sheet.SetCell("A1", title).SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 18, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#2E75B6"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:F1")

	// Company info template variables
	sheet.SetCell("A3", "Company:").SetStyle(getBoldStyle())
	sheet.SetCell("B3", "{{company_name}}")
	sheet.SetCell("A4", "Report Date:").SetStyle(getBoldStyle())
	sheet.SetCell("B4", "{{report_date}}")
	sheet.SetCell("A5", "Period:").SetStyle(getBoldStyle())
	sheet.SetCell("B5", "{{report_period}}")

	// Metrics section with template variables
	sheet.SetCell("A7", "KEY METRICS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A7:F7")

	metricsData := [][]interface{}{
		{"Metric", "Value", "Previous", "Change", "Target", "Status"},
		{"Total Sales", "{{total_sales}}", "{{prev_sales}}", "{{sales_change}}", "{{sales_target}}", "{{sales_status}}"},
		{"Total Orders", "{{total_orders}}", "{{prev_orders}}", "{{orders_change}}", "{{orders_target}}", "{{orders_status}}"},
		{"Active Customers", "{{active_customers}}", "{{prev_customers}}", "{{customers_change}}", "{{customers_target}}", "{{customers_status}}"},
		{"Conversion Rate", "{{conversion_rate}}", "{{prev_conversion}}", "{{conversion_change}}", "{{conversion_target}}", "{{conversion_status}}"},
	}

	for i, row := range metricsData {
		rowNum := 9 + i
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}

	// Template instructions
	sheet.SetCell("A15", "TEMPLATE INSTRUCTIONS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A15:F15")
	sheet.SetCell("A17", "This template uses the following variables:")
	sheet.SetCell("A18", "• {{company_name}} - Company name")
	sheet.SetCell("A19", "• {{report_date}} - Report generation date")
	sheet.SetCell("A20", "• {{total_sales}} - Total sales amount")
	sheet.SetCell("A21", "• {{total_orders}} - Number of orders")
	sheet.SetCell("A22", "• {{active_customers}} - Active customer count")
}

// Helper functions for styling
func getBoldStyle() excelbuilder.StyleConfig {
	return excelbuilder.StyleConfig{
        Font: excelbuilder.FontConfig{Bold: true},
    }
}

func getSectionHeaderStyle() excelbuilder.StyleConfig {
	return excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 12, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#4472C4"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	}
}

func getTableHeaderStyle() excelbuilder.StyleConfig {
	return excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#70AD47"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	}
}

func getDataCellStyle() excelbuilder.StyleConfig {
	return excelbuilder.StyleConfig{
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "left"},
	}
}

func createKPIDashboardTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	// KPI cards with enhanced styling
	kpiData := [][]interface{}{
		{"Revenue", fmt.Sprintf("$%.2f", data.Metrics.TotalSales), "+15.2%", "On Track"},
		{"Orders", data.Metrics.TotalOrders, "+8.5%", "Excellent"},
		{"Customers", data.Metrics.ActiveCustomers, "+12.1%", "Good"},
		{"Conversion", fmt.Sprintf("%.1f%%", data.Metrics.ConversionRate*100), "+2.3%", "Improving"},
	}

	for i, kpi := range kpiData {
		col := string(rune('A' + i*2))
		sheet.SetCell(fmt.Sprintf("%s3", col), kpi[0]).SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{Bold: true, Size: 12},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#F2F2F2"},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
		})
		sheet.SetCell(fmt.Sprintf("%s4", col), kpi[1]).SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{Bold: true, Size: 18, Color: "#2E75B6"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
		})
		sheet.SetCell(fmt.Sprintf("%s5", col), kpi[2]).SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{Color: "#70AD47"},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
		})
		sheet.SetCell(fmt.Sprintf("%s6", col), kpi[3]).SetStyle(excelbuilder.StyleConfig{
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
		})
	}
}

func createSalesDashboardTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	sheet.SetCell("A1", "SALES PERFORMANCE DASHBOARD").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 18, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#70AD47"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:H1")

	// Sales metrics
	sheet.SetCell("A3", "SALES METRICS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A3:H3")

	salesData := [][]interface{}{
		{"Period", "Revenue", "Orders", "Avg Order", "Growth", "Target", "Achievement", "Status"},
		{"This Month", data.Metrics.TotalSales, data.Metrics.TotalOrders, data.Metrics.TotalSales/float64(data.Metrics.TotalOrders), "15.2%", "$120,000", "104.2%", "Exceeded"},
		{"Last Month", 108500.0, 425, 255.29, "12.8%", "$105,000", "103.3%", "Exceeded"},
		{"YTD", 485000.0, 1850, 262.16, "18.5%", "$450,000", "107.8%", "Exceeded"},
	}

	for i, row := range salesData {
		rowNum := 5 + i
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createAnalyticsDashboardTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	sheet.SetCell("A1", "ANALYTICS DASHBOARD").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 18, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#FFC000"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:H1")

	// Analytics data
	sheet.SetCell("A3", "PERFORMANCE ANALYTICS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A3:H3")

	analyticsData := [][]interface{}{
		{"Metric", "Current", "Previous", "Change", "Trend", "Forecast", "Confidence", "Action"},
		{"Conversion Rate", fmt.Sprintf("%.1f%%", data.Metrics.ConversionRate*100), "3.1%", "+0.4%", "↗", "3.8%", "High", "Optimize"},
		{"Customer LTV", "$2,450", "$2,280", "+7.5%", "↗", "$2,650", "Medium", "Retain"},
		{"Churn Rate", "2.1%", "2.8%", "-0.7%", "↘", "1.8%", "High", "Monitor"},
		{"Market Share", "12.5%", "11.8%", "+0.7%", "↗", "13.2%", "Medium", "Expand"},
	}

	for i, row := range analyticsData {
		rowNum := 5 + i
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createSurveyFormTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	sheet.SetCell("A1", "CUSTOMER SURVEY FORM").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 16, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#C5504B"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:F1")

	// Form fields
	formFields := [][]interface{}{
		{"Field", "Type", "Required", "Options", "Validation", "Response"},
		{"Customer Name", "Text", "Yes", "-", "Required", "[Text Input]"},
		{"Email Address", "Email", "Yes", "-", "Email Format", "[Email Input]"},
		{"Satisfaction Rating", "Scale", "Yes", "1-10", "Numeric", "[Rating]"},
		{"Product Category", "Dropdown", "Yes", "Electronics,Accessories,Software", "Selection", "[Dropdown]"},
		{"Comments", "Textarea", "No", "-", "Max 500 chars", "[Text Area]"},
	}

	for i, row := range formFields {
		rowNum := 3 + i
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createRegistrationFormTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	sheet.SetCell("A1", "EVENT REGISTRATION FORM").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 16, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#4472C4"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:F1")

	// Registration fields
	registrationFields := []string{
		"Full Name:", "Email Address:", "Phone Number:", "Company:",
		"Job Title:", "Dietary Restrictions:", "Emergency Contact:", "T-Shirt Size:",
	}

	for i, field := range registrationFields {
		row := 3 + i
		sheet.SetCell(fmt.Sprintf("A%d", row), field).SetStyle(getBoldStyle())
		sheet.SetCell(fmt.Sprintf("B%d", row), "[Input Field]")
	}

	// Agreement section
	sheet.SetCell("A12", "AGREEMENT").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A12:F12")
	sheet.SetCell("A14", "☐ I agree to the terms and conditions")
	sheet.SetCell("A15", "☐ I consent to receive event updates")
	sheet.SetCell("A16", "☐ I allow photography during the event")

	// Signature section
	sheet.SetCell("A18", "Signature: ________________________")
	sheet.SetCell("D18", "Date: ____________")
}

func createFeedbackFormTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	sheet.SetCell("A1", data.Form.Title).SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 16, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#70AD47"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:F1")

	sheet.SetCell("A3", data.Form.Description)

	// Dynamic form fields from data
	for i, field := range data.Form.Fields {
		row := 5 + i*2
		label := field.Label
		if field.Required {
			label += " *"
		}
		sheet.SetCell(fmt.Sprintf("A%d", row), label).SetStyle(getBoldStyle())
		
		switch field.Type {
		case "select", "radio":
			sheet.SetCell(fmt.Sprintf("B%d", row), fmt.Sprintf("[%s: %v]", field.Type, field.Options))
		case "textarea":
			sheet.SetCell(fmt.Sprintf("B%d", row), "[Large Text Area]")
		default:
			sheet.SetCell(fmt.Sprintf("B%d", row), fmt.Sprintf("[%s Input]", field.Type))
		}
	}
}

func createAchievementCertificateTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	// Certificate border
	sheet.SetCell("A1", "CERTIFICATE OF ACHIEVEMENT").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 24, Color: "#C5504B"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:H3")

	// Certificate content
	sheet.SetCell("A6", "This is to certify that").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Size: 14},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A6:H6")

	sheet.SetCell("A8", data.Certificate.RecipientName).SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 20, Color: "#2E75B6"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A8:H8")

	sheet.SetCell("A10", "has successfully completed").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Size: 14},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A10:H10")

	sheet.SetCell("A12", data.Certificate.CourseName).SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 16, Color: "#70AD47"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A12:H12")

	// Certificate details
	sheet.SetCell("A15", fmt.Sprintf("Completion Date: %s", data.Certificate.CompletionDate.Format("January 2, 2006")))
	sheet.SetCell("A16", fmt.Sprintf("Credits Earned: %d", data.Certificate.Credits))
	sheet.SetCell("A17", fmt.Sprintf("Grade: %s", data.Certificate.Grade))
	sheet.SetCell("A18", fmt.Sprintf("Certificate ID: %s", data.Certificate.CertificateID))

	// Signature section
	sheet.SetCell("A21", "Instructor: " + data.Certificate.Instructor)
	sheet.SetCell("F21", "Date: " + data.Certificate.CompletionDate.Format("2006-01-02"))
	sheet.SetCell("A22", "Signature: _________________________")
}

func createCourseCompletionTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	sheet.SetCell("A1", "COURSE COMPLETION CERTIFICATE").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 20, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#4472C4"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:H2")

	// Course details
	sheet.SetCell("A5", "Course Information").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A5:H5")

	courseInfo := [][]interface{}{
		{"Course Name:", data.Certificate.CourseName},
		{"Student Name:", data.Certificate.RecipientName},
		{"Completion Date:", data.Certificate.CompletionDate.Format("January 2, 2006")},
		{"Instructor:", data.Certificate.Instructor},
		{"Credits:", fmt.Sprintf("%d hours", data.Certificate.Credits)},
		{"Final Grade:", data.Certificate.Grade},
		{"Valid Until:", data.Certificate.ValidUntil.Format("January 2, 2006")},
	}

	for i, info := range courseInfo {
		row := 7 + i
		sheet.SetCell(fmt.Sprintf("A%d", row), info[0]).SetStyle(getBoldStyle())
		sheet.SetCell(fmt.Sprintf("C%d", row), info[1])
	}

	// Verification section
	sheet.SetCell("A16", "VERIFICATION").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A16:H16")
	sheet.SetCell("A18", "This certificate can be verified at: www.techcorp.com/verify")
	sheet.SetCell("A19", fmt.Sprintf("Verification Code: %s", data.Certificate.CertificateID))
}

func createAwardCertificateTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	// Decorative award certificate
	sheet.SetCell("A1", "🏆 AWARD CERTIFICATE 🏆").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 22, Color: "#FFC000"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:H2")

	// Award content
	sheet.SetCell("A5", "In recognition of outstanding achievement").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Size: 14, Italic: true},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A5:H5")

	sheet.SetCell("A7", "This award is presented to").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Size: 12},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A7:H7")

	sheet.SetCell("A9", data.Certificate.RecipientName).SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 20, Color: "#C5504B"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A9:H9")

	sheet.SetCell("A11", "For excellence in").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Size: 12},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A11:H11")

	sheet.SetCell("A13", data.Certificate.CourseName).SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 16, Color: "#FFFFFF"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A13:H13")

	// Award details
	sheet.SetCell("A16", "Presented on " + data.Certificate.CompletionDate.Format("January 2, 2006"))
	sheet.SetCell("A17", "By " + data.Certificate.Instructor)
	sheet.SetCell("A18", "Grade Achieved: " + data.Certificate.Grade)

	// Seal/signature area
	sheet.SetCell("A21", "[OFFICIAL SEAL]")
	sheet.SetCell("F21", "Authorized Signature")
	sheet.SetCell("F22", "_________________________")
}

func createProductReportTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	sheet.SetCell("A1", "PRODUCT PERFORMANCE REPORT").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 16, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#70AD47"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:G1")

	// Template section for product loop
	sheet.SetCell("A3", "TOP PRODUCTS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A3:G3")

	headers := []string{"Product ID", "Name", "Category", "Price", "Quantity", "Sales", "Status"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s5", col), header).SetStyle(getTableHeaderStyle())
	}

	// Template loop indicator
	sheet.SetCell("A7", "{{#each products}}")
	sheet.SetCell("A8", "{{product_id}}")
	sheet.SetCell("B8", "{{product_name}}")
	sheet.SetCell("C8", "{{product_category}}")
	sheet.SetCell("D8", "{{product_price}}")
	sheet.SetCell("E8", "{{product_quantity}}")
	sheet.SetCell("F8", "{{product_sales}}")
	sheet.SetCell("G8", "{{product_status}}")
	sheet.SetCell("A9", "{{/each}}")

	// Sample data for demonstration
	for i, product := range data.Products[:5] {
		row := 11 + i
		sheet.SetCell(fmt.Sprintf("A%d", row), product.ID)
		sheet.SetCell(fmt.Sprintf("B%d", row), product.Name)
		sheet.SetCell(fmt.Sprintf("C%d", row), product.Category)
		sheet.SetCell(fmt.Sprintf("D%d", row), product.Price).SetStyle(excelbuilder.StyleConfig{NumberFormat: "$#,##0.00"})
		sheet.SetCell(fmt.Sprintf("E%d", row), product.Quantity)
		sheet.SetCell(fmt.Sprintf("F%d", row), product.Price*float64(product.Quantity)).SetStyle(excelbuilder.StyleConfig{NumberFormat: "$#,##0.00"})
		sheet.SetCell(fmt.Sprintf("G%d", row), product.Status)
	}
}

func createCustomerReportTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	sheet.SetCell("A1", "CUSTOMER ANALYSIS REPORT").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 16, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#FFC000"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:H1")

	// Customer segmentation template
	sheet.SetCell("A3", "CUSTOMER SEGMENTATION").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A3:H3")

	segmentData := [][]interface{}{
		{"Segment", "Count", "Percentage", "Avg Value", "Total Value", "Growth", "Retention", "Status"},
		{"Premium", "{{premium_count}}", "{{premium_percent}}", "{{premium_avg}}", "{{premium_total}}", "{{premium_growth}}", "{{premium_retention}}", "{{premium_status}}"},
		{"Standard", "{{standard_count}}", "{{standard_percent}}", "{{standard_avg}}", "{{standard_total}}", "{{standard_growth}}", "{{standard_retention}}", "{{standard_status}}"},
		{"Basic", "{{basic_count}}", "{{basic_percent}}", "{{basic_avg}}", "{{basic_total}}", "{{basic_growth}}", "{{basic_retention}}", "{{basic_status}}"},
	}

	for i, row := range segmentData {
		rowNum := 5 + i
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}

	// Top customers template
	sheet.SetCell("A11", "TOP CUSTOMERS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A11:H11")

	customerHeaders := []string{"Customer ID", "Name", "Email", "Segment", "Value", "Orders", "Join Date", "Status"}
	for i, header := range customerHeaders {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s13", col), header).SetStyle(getTableHeaderStyle())
	}

	// Sample customer data
	for i, customer := range data.Customers[:5] {
		row := 14 + i
		sheet.SetCell(fmt.Sprintf("A%d", row), customer.ID)
		sheet.SetCell(fmt.Sprintf("B%d", row), customer.Name)
		sheet.SetCell(fmt.Sprintf("C%d", row), customer.Email)
		sheet.SetCell(fmt.Sprintf("D%d", row), customer.Segment)
		sheet.SetCell(fmt.Sprintf("E%d", row), customer.Value).SetStyle(excelbuilder.StyleConfig{NumberFormat: "$#,##0.00"})
		sheet.SetCell(fmt.Sprintf("F%d", row), rand.Intn(20)+1) // Random order count
		sheet.SetCell(fmt.Sprintf("G%d", row), customer.JoinDate.Format("2006-01-02"))
		sheet.SetCell(fmt.Sprintf("H%d", row), "Active")
	}
}

func createExecutiveSummaryTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	sheet.SetCell("A1", "EXECUTIVE SUMMARY").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 18, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#C5504B"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:F1")

	// Executive summary content with template variables
	sheet.SetCell("A3", "OVERVIEW").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A3:F3")

	sheet.SetCell("A5", "This {{report_period}} report shows {{company_name}} achieved:")
	sheet.SetCell("A6", "• Total Revenue: {{total_revenue}}")
	sheet.SetCell("A7", "• Growth Rate: {{growth_rate}}")
	sheet.SetCell("A8", "• Customer Acquisition: {{new_customers}} new customers")
	sheet.SetCell("A9", "• Market Share: {{market_share}}")

	sheet.SetCell("A11", "KEY ACHIEVEMENTS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A11:F11")

	sheet.SetCell("A13", "{{#if achievements}}")
	sheet.SetCell("A14", "{{#each achievements}}")
	sheet.SetCell("A15", "• {{achievement_text}}")
	sheet.SetCell("A16", "{{/each}}")
	sheet.SetCell("A17", "{{else}}")
	sheet.SetCell("A18", "No specific achievements to highlight this period.")
	sheet.SetCell("A19", "{{/if}}")

	sheet.SetCell("A21", "NEXT STEPS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A21:F21")

	sheet.SetCell("A23", "• {{next_step_1}}")
	sheet.SetCell("A24", "• {{next_step_2}}")
	sheet.SetCell("A25", "• {{next_step_3}}")
}

// Invoice template functions
func createStandardInvoiceTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	// Invoice header
	sheet.SetCell("A1", "INVOICE").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 24, Color: "#2E75B6"},
	})

	// Company info template
	sheet.SetCell("A3", "{{company_name}}").SetStyle(getBoldStyle())
	sheet.SetCell("A4", "{{company_address}}")
	sheet.SetCell("A5", "{{company_phone}}")
	sheet.SetCell("A6", "{{company_email}}")

	// Invoice details
	sheet.SetCell("E1", "Invoice #:").SetStyle(getBoldStyle())
	sheet.SetCell("F1", "{{invoice_number}}")
	sheet.SetCell("E2", "Date:").SetStyle(getBoldStyle())
	sheet.SetCell("F2", "{{invoice_date}}")
	sheet.SetCell("E3", "Due Date:").SetStyle(getBoldStyle())
	sheet.SetCell("F3", "{{due_date}}")

	// Bill to section
	sheet.SetCell("A8", "BILL TO:").SetStyle(getBoldStyle())
	sheet.SetCell("A9", "{{customer_name}}")
	sheet.SetCell("A10", "{{customer_address}}")
	sheet.SetCell("A11", "{{customer_city}}, {{customer_state}} {{customer_zip}}")

	// Items table
	sheet.SetCell("A13", "DESCRIPTION").SetStyle(getTableHeaderStyle())
	sheet.SetCell("D13", "QTY").SetStyle(getTableHeaderStyle())
	sheet.SetCell("E13", "RATE").SetStyle(getTableHeaderStyle())
	sheet.SetCell("F13", "AMOUNT").SetStyle(getTableHeaderStyle())

	// Template loop for items
	sheet.SetCell("A15", "{{#each invoice_items}}")
	sheet.SetCell("A16", "{{item_description}}")
	sheet.SetCell("D16", "{{item_quantity}}")
	sheet.SetCell("E16", "{{item_rate}}")
	sheet.SetCell("F16", "{{item_amount}}")
	sheet.SetCell("A17", "{{/each}}")

	// Totals
	sheet.SetCell("E20", "Subtotal:").SetStyle(getBoldStyle())
	sheet.SetCell("F20", "{{subtotal}}")
	sheet.SetCell("E21", "Tax:").SetStyle(getBoldStyle())
	sheet.SetCell("F21", "{{tax_amount}}")
	sheet.SetCell("E22", "Total:").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 14},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#E7E6E6"},
	})
	sheet.SetCell("F22", "{{total_amount}}").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 14},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#E7E6E6"},
		NumberFormat: "$#,##0.00",
	})

	// Payment terms
	sheet.SetCell("A25", "Payment Terms: {{payment_terms}}")
	sheet.SetCell("A26", "Notes: {{invoice_notes}}")
}

func createProfessionalInvoiceTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	// Professional invoice with enhanced styling
	sheet.SetCell("A1", "PROFESSIONAL INVOICE").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 20, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#4472C4"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:F1")

	// Enhanced company branding section
	sheet.SetCell("A3", "{{company_logo}}") // Placeholder for logo
	sheet.SetCell("B3", "{{company_name}}").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 16, Color: "#4472C4"},
	})
	sheet.SetCell("B4", "{{company_tagline}}")
	sheet.SetCell("B5", "{{company_address}}")
	sheet.SetCell("B6", "Phone: {{company_phone}} | Email: {{company_email}}")
	sheet.SetCell("B7", "Website: {{company_website}}")

	// Professional invoice details with styling
	sheet.SetCell("E3", "Invoice Details").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Color: "#4472C4"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#F2F2F2"},
	})
	sheet.SetCell("E4", "Invoice #: {{invoice_number}}")
	sheet.SetCell("E5", "Date: {{invoice_date}}")
	sheet.SetCell("E6", "Due: {{due_date}}")
	sheet.SetCell("E7", "Terms: {{payment_terms}}")

	// Client information with professional styling
	sheet.SetCell("A10", "Bill To").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Color: "#4472C4"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#F2F2F2"},
	})
	sheet.SetCell("A11", "{{client_company}}")
	sheet.SetCell("A12", "{{client_contact}}")
	sheet.SetCell("A13", "{{client_address}}")
	sheet.SetCell("A14", "{{client_email}}")

	// Professional items table
	headers := []string{"Description", "Hours/Qty", "Rate", "Amount"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s16", col), header).SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{Bold: true, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#4472C4"},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
		})
	}

	// Sample professional service items
	serviceItems := [][]interface{}{
		{"Consulting Services - Strategy Development", 40, 150.00, 6000.00},
		{"Project Management - Implementation Phase", 60, 125.00, 7500.00},
		{"Training and Documentation", 20, 100.00, 2000.00},
	}

	for i, item := range serviceItems {
		row := 17 + i
		for j, cell := range item {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if j >= 2 { // Rate and Amount columns
				style.NumberFormat = "$#,##0.00"
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, row), cell).SetStyle(style)
		}
	}

	// Professional totals section
	totalsRow := 22
	sheet.SetCell(fmt.Sprintf("C%d", totalsRow), "Subtotal:").SetStyle(getBoldStyle())
	sheet.SetCell(fmt.Sprintf("D%d", totalsRow), 15500.00).SetStyle(excelbuilder.StyleConfig{NumberFormat: "$#,##0.00"})
	sheet.SetCell(fmt.Sprintf("C%d", totalsRow+1), "Tax (8%):").SetStyle(getBoldStyle())
	sheet.SetCell(fmt.Sprintf("D%d", totalsRow+1), 1240.00).SetStyle(excelbuilder.StyleConfig{NumberFormat: "$#,##0.00"})
	sheet.SetCell(fmt.Sprintf("C%d", totalsRow+2), "Total:").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 14, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#4472C4"},
	})
	sheet.SetCell(fmt.Sprintf("D%d", totalsRow+2), 16740.00).SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 14, Color: "#FFFFFF"},
		Fill: excelbuilder.FillConfig{Type: "solid", Color: "#4472C4"},
		NumberFormat: "$#,##0.00",
	})
}

func createServiceInvoiceTemplate(sheet *excelbuilder.SheetBuilder, data *TemplateData) {
	// Service-specific invoice template
	sheet.SetCell("A1", "SERVICE INVOICE").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 18, Color: "#FFFFFF"},
        Fill: excelbuilder.FillConfig{Type: "solid", Color: "#70AD47"},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "center"},
	})
	sheet.MergeRange("A1:F1")

	// Service provider info
	sheet.SetCell("A3", "Service Provider").SetStyle(getBoldStyle())
	sheet.SetCell("A4", "{{provider_name}}")
	sheet.SetCell("A5", "License #: {{license_number}}")
	sheet.SetCell("A6", "{{provider_address}}")
	sheet.SetCell("A7", "{{provider_contact}}")

	// Service details
	sheet.SetCell("D3", "Service Details").SetStyle(getBoldStyle())
	sheet.SetCell("D4", "Service Date: {{service_date}}")
	sheet.SetCell("D5", "Completion: {{completion_date}}")
	sheet.SetCell("D6", "Technician: {{technician_name}}")
	sheet.SetCell("D7", "Work Order: {{work_order_number}}")

	// Services performed
	sheet.SetCell("A10", "SERVICES PERFORMED").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A10:F10")

	serviceHeaders := []string{"Service Description", "Hours", "Rate", "Parts", "Labor", "Total"}
	for i, header := range serviceHeaders {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s12", col), header).SetStyle(getTableHeaderStyle())
	}

	// Template for service items
	sheet.SetCell("A14", "{{#each services}}")
	sheet.SetCell("A15", "{{service_description}}")
	sheet.SetCell("B15", "{{service_hours}}")
	sheet.SetCell("C15", "{{hourly_rate}}")
	sheet.SetCell("D15", "{{parts_cost}}")
	sheet.SetCell("E15", "{{labor_cost}}")
	sheet.SetCell("F15", "{{service_total}}")
	sheet.SetCell("A16", "{{/each}}")

	// Service summary
	sheet.SetCell("A20", "SERVICE SUMMARY").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A20:F20")

	sheet.SetCell("A22", "Work performed: {{work_summary}}")
	sheet.SetCell("A23", "Parts used: {{parts_summary}}")
	sheet.SetCell("A24", "Warranty: {{warranty_info}}")
	sheet.SetCell("A25", "Next service: {{next_service_date}}")
}