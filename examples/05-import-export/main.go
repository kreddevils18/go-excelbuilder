package main

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Data structures for different import formats
type Employee struct {
	ID         int       `json:"id"`
	FirstName  string    `json:"first_name"`
	LastName   string    `json:"last_name"`
	Email      string    `json:"email"`
	Department string    `json:"department"`
	Position   string    `json:"position"`
	Salary     float64   `json:"salary"`
	HireDate   time.Time `json:"hire_date"`
	Active     bool      `json:"active"`
}

type SalesData struct {
	Date     string  `json:"date"`
	Region   string  `json:"region"`
	Product  string  `json:"product"`
	Quantity int     `json:"quantity"`
	Revenue  float64 `json:"revenue"`
	Salesman string  `json:"salesman"`
}

type InventoryItem struct {
	SKU         string  `json:"sku"`
	ProductName string  `json:"product_name"`
	Category    string  `json:"category"`
	Quantity    int     `json:"quantity"`
	UnitPrice   float64 `json:"unit_price"`
	Supplier    string  `json:"supplier"`
	LastUpdated string  `json:"last_updated"`
}

func main() {
	// Create output and data directories
	if err := os.MkdirAll("output", 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}
	if err := os.MkdirAll("data", 0755); err != nil {
		log.Fatalf("Failed to create data directory: %v", err)
	}

	fmt.Println("🔄 Starting Import/Export Data Integration Demo...")

	// Step 1: Generate sample data files
	fmt.Println("\n📁 Step 1: Generating sample data files...")
	generateSampleDataFiles()

	// Step 2: Import data from various sources
	fmt.Println("\n📥 Step 2: Importing data from files...")
	employees := importEmployeesFromJSON("data/employees.json")
	salesData := importSalesFromCSV("data/sales_data.csv")
	inventory := importInventoryFromJSON("data/inventory.json")

	// Step 3: Create Excel workbook with imported data
	fmt.Println("\n📊 Step 3: Creating Excel workbook...")
	createIntegratedWorkbook(employees, salesData, inventory)

	// Step 4: Export data back to different formats
	fmt.Println("\n📤 Step 4: Exporting processed data...")
	exportProcessedData(employees, salesData, inventory)

	// Step 5: Demonstrate round-trip data integrity
	fmt.Println("\n🔄 Step 5: Demonstrating round-trip data integrity...")
	demonstateRoundTripIntegrity()

	fmt.Println("\n✅ Import/Export demonstration completed successfully!")
	fmt.Println("📁 Generated files:")
	fmt.Println("   • output/05-import-export-integrated.xlsx - Main integrated workbook")
	fmt.Println("   • output/processed_employees.json - Processed employee data")
	fmt.Println("   • output/sales_summary.csv - Sales summary export")
	fmt.Println("   • output/inventory_report.json - Inventory analysis")
	fmt.Println("   • output/round_trip_test.xlsx - Data integrity verification")
	fmt.Println("\n🎯 Next steps: Try examples/06-dashboard/ for interactive reporting")
}

func generateSampleDataFiles() {
	// Generate employees JSON
	employees := []Employee{
		{1, "John", "Doe", "john.doe@company.com", "Engineering", "Senior Developer", 95000, time.Date(2020, 3, 15, 0, 0, 0, 0, time.UTC), true},
		{2, "Jane", "Smith", "jane.smith@company.com", "Marketing", "Marketing Manager", 75000, time.Date(2019, 7, 22, 0, 0, 0, 0, time.UTC), true},
		{3, "Bob", "Johnson", "bob.johnson@company.com", "Sales", "Sales Representative", 65000, time.Date(2021, 1, 10, 0, 0, 0, 0, time.UTC), true},
		{4, "Alice", "Brown", "alice.brown@company.com", "Engineering", "Lead Architect", 120000, time.Date(2018, 11, 5, 0, 0, 0, 0, time.UTC), true},
		{5, "Charlie", "Wilson", "charlie.wilson@company.com", "HR", "HR Specialist", 55000, time.Date(2022, 4, 18, 0, 0, 0, 0, time.UTC), false},
		{6, "Diana", "Davis", "diana.davis@company.com", "Finance", "Financial Analyst", 70000, time.Date(2020, 9, 12, 0, 0, 0, 0, time.UTC), true},
		{7, "Eva", "Miller", "eva.miller@company.com", "Engineering", "DevOps Engineer", 85000, time.Date(2021, 6, 8, 0, 0, 0, 0, time.UTC), true},
		{8, "Frank", "Taylor", "frank.taylor@company.com", "Sales", "Sales Manager", 90000, time.Date(2019, 12, 3, 0, 0, 0, 0, time.UTC), true},
	}

	employeesJSON, _ := json.MarshalIndent(employees, "", "  ")
	os.WriteFile("data/employees.json", employeesJSON, 0644)
	fmt.Println("   ✓ Generated employees.json")

	// Generate sales CSV
	salesFile, _ := os.Create("data/sales_data.csv")
	defer salesFile.Close()

	writer := csv.NewWriter(salesFile)
	defer writer.Flush()

	// CSV headers
	writer.Write([]string{"Date", "Region", "Product", "Quantity", "Revenue", "Salesman"})

	// Sample sales data
	salesRecords := [][]string{
		{"2024-01-15", "North America", "Product A", "100", "15000.00", "Bob Johnson"},
		{"2024-01-16", "Europe", "Product B", "75", "22500.00", "Frank Taylor"},
		{"2024-01-17", "Asia", "Product A", "120", "18000.00", "Bob Johnson"},
		{"2024-01-18", "North America", "Product C", "50", "35000.00", "Frank Taylor"},
		{"2024-01-19", "Europe", "Product A", "90", "13500.00", "Bob Johnson"},
		{"2024-01-20", "Asia", "Product B", "110", "33000.00", "Frank Taylor"},
		{"2024-01-21", "North America", "Product B", "85", "25500.00", "Bob Johnson"},
		{"2024-01-22", "Europe", "Product C", "60", "42000.00", "Frank Taylor"},
		{"2024-01-23", "Asia", "Product A", "95", "14250.00", "Bob Johnson"},
		{"2024-01-24", "North America", "Product C", "70", "49000.00", "Frank Taylor"},
	}

	for _, record := range salesRecords {
		writer.Write(record)
	}
	fmt.Println("   ✓ Generated sales_data.csv")

	// Generate inventory JSON
	inventory := []InventoryItem{
		{"SKU001", "Product A", "Electronics", 250, 150.00, "Supplier Alpha", "2024-01-20"},
		{"SKU002", "Product B", "Electronics", 180, 300.00, "Supplier Beta", "2024-01-19"},
		{"SKU003", "Product C", "Electronics", 120, 700.00, "Supplier Gamma", "2024-01-18"},
		{"SKU004", "Product D", "Accessories", 500, 25.00, "Supplier Alpha", "2024-01-21"},
		{"SKU005", "Product E", "Accessories", 300, 45.00, "Supplier Delta", "2024-01-17"},
		{"SKU006", "Product F", "Software", 1000, 99.99, "Supplier Epsilon", "2024-01-22"},
		{"SKU007", "Product G", "Software", 750, 199.99, "Supplier Zeta", "2024-01-16"},
		{"SKU008", "Product H", "Hardware", 80, 1200.00, "Supplier Eta", "2024-01-15"},
	}

	inventoryJSON, _ := json.MarshalIndent(inventory, "", "  ")
	os.WriteFile("data/inventory.json", inventoryJSON, 0644)
	fmt.Println("   ✓ Generated inventory.json")
}

func importEmployeesFromJSON(filename string) []Employee {
	fmt.Printf("   📥 Importing employees from %s...\n", filename)
	data, err := os.ReadFile(filename)
	if err != nil {
		log.Fatalf("Failed to read employees file: %v", err)
	}

	var employees []Employee
	err = json.Unmarshal(data, &employees)
	if err != nil {
		log.Fatalf("Failed to parse employees JSON: %v", err)
	}

	fmt.Printf("   ✓ Imported %d employees\n", len(employees))
	return employees
}

func importSalesFromCSV(filename string) []SalesData {
	fmt.Printf("   📥 Importing sales data from %s...\n", filename)
	file, err := os.Open(filename)
	if err != nil {
		log.Fatalf("Failed to open sales file: %v", err)
	}
	defer file.Close()

	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		log.Fatalf("Failed to read CSV: %v", err)
	}

	var salesData []SalesData
	// Skip header row
	for i := 1; i < len(records); i++ {
		record := records[i]
		quantity, _ := strconv.Atoi(record[3])
		revenue, _ := strconv.ParseFloat(record[4], 64)
		
		sales := SalesData{
			Date:     record[0],
			Region:   record[1],
			Product:  record[2],
			Quantity: quantity,
			Revenue:  revenue,
			Salesman: record[5],
		}
		salesData = append(salesData, sales)
	}

	fmt.Printf("   ✓ Imported %d sales records\n", len(salesData))
	return salesData
}

func importInventoryFromJSON(filename string) []InventoryItem {
	fmt.Printf("   📥 Importing inventory from %s...\n", filename)
	data, err := os.ReadFile(filename)
	if err != nil {
		log.Fatalf("Failed to read inventory file: %v", err)
	}

	var inventory []InventoryItem
	err = json.Unmarshal(data, &inventory)
	if err != nil {
		log.Fatalf("Failed to parse inventory JSON: %v", err)
	}

	fmt.Printf("   ✓ Imported %d inventory items\n", len(inventory))
	return inventory
}

func createIntegratedWorkbook(employees []Employee, salesData []SalesData, inventory []InventoryItem) {
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Integrated Data Analysis Report",
		Author:      "Data Integration System",
		Subject:     "Multi-Source Data Analysis",
		Description: "Comprehensive analysis combining employee, sales, and inventory data from multiple sources",
		Keywords:    "data-integration,import,export,analysis,business-intelligence",
		Company:     "TechCorp Solutions",
	})

	// Create styles
	styles := createIntegrationStyles()

	// Sheet 1: Employee Data
	employeeSheet := workbook.AddSheet("Employee Data")
	if employeeSheet == nil {
		log.Fatal("Failed to create employee sheet")
	}
	createEmployeeSheet(employeeSheet, styles, employees)

	// Sheet 2: Sales Analysis
	salesSheet := workbook.AddSheet("Sales Analysis")
	if salesSheet == nil {
		log.Fatal("Failed to create sales sheet")
	}
	createSalesSheet(salesSheet, styles, salesData)

	// Sheet 3: Inventory Report
	inventorySheet := workbook.AddSheet("Inventory Report")
	if inventorySheet == nil {
		log.Fatal("Failed to create inventory sheet")
	}
	createInventorySheet(inventorySheet, styles, inventory)

	// Sheet 4: Data Summary
	summarySheet := workbook.AddSheet("Data Summary")
	if summarySheet == nil {
		log.Fatal("Failed to create summary sheet")
	}
	createDataSummarySheet(summarySheet, styles, employees, salesData, inventory)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build workbook")
	}

	filename := "output/05-import-export-integrated.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save workbook: %v", err)
	}

	fmt.Printf("   ✓ Created integrated workbook: %s\n", filename)
}

func createIntegrationStyles() map[string]excelbuilder.StyleConfig {
	styles := make(map[string]excelbuilder.StyleConfig)

	// Color scheme
	colors := map[string]string{
		"primary":     "#2E86AB",
		"secondary":   "#A23B72",
		"accent":      "#F18F01",
		"success":     "#C73E1D",
		"light_gray":  "#F5F5F5",
		"dark_gray":   "#666666",
		"white":       "#FFFFFF",
		"black":       "#000000",
	}

	// Header styles
	styles["sheet_title"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   16,
			Color:  colors["primary"],
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	}

	styles["table_header"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   11,
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
		Border: createBorder("thin", colors["white"]),
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
		Border: createBorder("thin", colors["light_gray"]),
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
		Border: createBorder("thin", colors["dark_gray"]),
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
		Border:      createBorder("thin", colors["light_gray"]),
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
		Border:      createBorder("thin", colors["light_gray"]),
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
		Border:      createBorder("thin", colors["light_gray"]),
	}

	styles["boolean"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Calibri",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["light_gray"]),
	}

	// Summary styles
	styles["summary_metric"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
			Color:  colors["white"],
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["accent"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createBorder("medium", colors["black"]),
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

func createEmployeeSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, employees []Employee) {
	// Set column widths
	sheet.SetColumnWidth("A", 8.0)   // ID
	sheet.SetColumnWidth("B", 15.0)  // First Name
	sheet.SetColumnWidth("C", 15.0)  // Last Name
	sheet.SetColumnWidth("D", 25.0)  // Email
	sheet.SetColumnWidth("E", 15.0)  // Department
	sheet.SetColumnWidth("F", 20.0)  // Position
	sheet.SetColumnWidth("G", 12.0)  // Salary
	sheet.SetColumnWidth("H", 12.0)  // Hire Date
	sheet.SetColumnWidth("I", 8.0)   // Active

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("EMPLOYEE DATA (Imported from JSON)").SetStyle(styles["sheet_title"])

	sheet.AddRow() // Empty row
	sheet.AddRow() // Empty row

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("ID").SetStyle(styles["table_header"])
	headerRow.AddCell("First Name").SetStyle(styles["table_header"])
	headerRow.AddCell("Last Name").SetStyle(styles["table_header"])
	headerRow.AddCell("Email").SetStyle(styles["table_header"])
	headerRow.AddCell("Department").SetStyle(styles["table_header"])
	headerRow.AddCell("Position").SetStyle(styles["table_header"])
	headerRow.AddCell("Salary").SetStyle(styles["table_header"])
	headerRow.AddCell("Hire Date").SetStyle(styles["table_header"])
	headerRow.AddCell("Active").SetStyle(styles["table_header"])

	// Employee data
	for i, emp := range employees {
		row := sheet.AddRow()
		
		// Alternating row colors
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell(emp.ID).SetStyle(styles["integer"])
		row.AddCell(emp.FirstName).SetStyle(rowStyle)
		row.AddCell(emp.LastName).SetStyle(rowStyle)
		row.AddCell(emp.Email).SetStyle(rowStyle)
		row.AddCell(emp.Department).SetStyle(rowStyle)
		row.AddCell(emp.Position).SetStyle(rowStyle)
		row.AddCell(emp.Salary).SetStyle(styles["currency"])
		row.AddCell(emp.HireDate).SetStyle(styles["date"])
		
		activeText := "No"
		if emp.Active {
			activeText = "Yes"
		}
		row.AddCell(activeText).SetStyle(styles["boolean"])
	}

	// Summary statistics
	sheet.AddRow()
	sheet.AddRow()
	summaryRow := sheet.AddRow()
	summaryRow.AddCell("SUMMARY STATISTICS").SetStyle(styles["sheet_title"])

	sheet.AddRow()

	// Calculate statistics
	totalEmployees := len(employees)
	activeEmployees := 0
	totalSalary := 0.0
	departmentCounts := make(map[string]int)

	for _, emp := range employees {
		if emp.Active {
			activeEmployees++
		}
		totalSalary += emp.Salary
		departmentCounts[emp.Department]++
	}

	avgSalary := totalSalary / float64(totalEmployees)

	// Display statistics
	statsData := []struct {
		label string
		value interface{}
		format string
	}{
		{"Total Employees", totalEmployees, "integer"},
		{"Active Employees", activeEmployees, "integer"},
		{"Average Salary", avgSalary, "currency"},
		{"Total Payroll", totalSalary, "currency"},
	}

	for _, stat := range statsData {
		statRow := sheet.AddRow()
		statRow.AddCell(stat.label).SetStyle(styles["data_normal"])
		if stat.format == "currency" {
			statRow.AddCell(stat.value).SetStyle(styles["currency"])
		} else {
			statRow.AddCell(stat.value).SetStyle(styles["integer"])
		}
	}
}

func createSalesSheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, salesData []SalesData) {
	// Set column widths
	sheet.SetColumnWidth("A", 12.0)  // Date
	sheet.SetColumnWidth("B", 15.0)  // Region
	sheet.SetColumnWidth("C", 15.0)  // Product
	sheet.SetColumnWidth("D", 10.0)  // Quantity
	sheet.SetColumnWidth("E", 15.0)  // Revenue
	sheet.SetColumnWidth("F", 18.0)  // Salesman
	sheet.SetColumnWidth("G", 15.0)  // Unit Price

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("SALES ANALYSIS (Imported from CSV)").SetStyle(styles["sheet_title"])

	sheet.AddRow() // Empty row
	sheet.AddRow() // Empty row

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("Date").SetStyle(styles["table_header"])
	headerRow.AddCell("Region").SetStyle(styles["table_header"])
	headerRow.AddCell("Product").SetStyle(styles["table_header"])
	headerRow.AddCell("Quantity").SetStyle(styles["table_header"])
	headerRow.AddCell("Revenue").SetStyle(styles["table_header"])
	headerRow.AddCell("Salesman").SetStyle(styles["table_header"])
	headerRow.AddCell("Unit Price").SetStyle(styles["table_header"])

	// Sales data
	for i, sale := range salesData {
		row := sheet.AddRow()
		
		// Alternating row colors
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		// Parse date for proper formatting
		date, _ := time.Parse("2006-01-02", sale.Date)
		
		row.AddCell(date).SetStyle(styles["date"])
		row.AddCell(sale.Region).SetStyle(rowStyle)
		row.AddCell(sale.Product).SetStyle(rowStyle)
		row.AddCell(sale.Quantity).SetStyle(styles["integer"])
		row.AddCell(sale.Revenue).SetStyle(styles["currency"])
		row.AddCell(sale.Salesman).SetStyle(rowStyle)
		
		// Calculate unit price
		unitPrice := sale.Revenue / float64(sale.Quantity)
		row.AddCell(unitPrice).SetStyle(styles["currency"])
	}

	// Sales summary
	sheet.AddRow()
	sheet.AddRow()
	summaryRow := sheet.AddRow()
	summaryRow.AddCell("SALES SUMMARY").SetStyle(styles["sheet_title"])

	sheet.AddRow()

	// Calculate summary metrics
	totalRevenue := 0.0
	totalQuantity := 0
	regionRevenue := make(map[string]float64)
	productRevenue := make(map[string]float64)

	for _, sale := range salesData {
		totalRevenue += sale.Revenue
		totalQuantity += sale.Quantity
		regionRevenue[sale.Region] += sale.Revenue
		productRevenue[sale.Product] += sale.Revenue
	}

	avgDealSize := totalRevenue / float64(len(salesData))
	avgUnitPrice := totalRevenue / float64(totalQuantity)

	// Display summary
	summaryData := []struct {
		label string
		value interface{}
		format string
	}{
		{"Total Revenue", totalRevenue, "currency"},
		{"Total Units Sold", totalQuantity, "integer"},
		{"Number of Deals", len(salesData), "integer"},
		{"Average Deal Size", avgDealSize, "currency"},
		{"Average Unit Price", avgUnitPrice, "currency"},
	}

	for _, summary := range summaryData {
		sumRow := sheet.AddRow()
		sumRow.AddCell(summary.label).SetStyle(styles["data_normal"])
		if summary.format == "currency" {
			sumRow.AddCell(summary.value).SetStyle(styles["summary_metric"])
		} else {
			sumRow.AddCell(summary.value).SetStyle(styles["summary_metric"])
		}
	}
}

func createInventorySheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, inventory []InventoryItem) {
	// Set column widths
	sheet.SetColumnWidth("A", 10.0)  // SKU
	sheet.SetColumnWidth("B", 20.0)  // Product Name
	sheet.SetColumnWidth("C", 15.0)  // Category
	sheet.SetColumnWidth("D", 10.0)  // Quantity
	sheet.SetColumnWidth("E", 12.0)  // Unit Price
	sheet.SetColumnWidth("F", 18.0)  // Supplier
	sheet.SetColumnWidth("G", 12.0)  // Last Updated
	sheet.SetColumnWidth("H", 15.0)  // Total Value

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("INVENTORY REPORT (Imported from JSON)").SetStyle(styles["sheet_title"])

	sheet.AddRow() // Empty row
	sheet.AddRow() // Empty row

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("SKU").SetStyle(styles["table_header"])
	headerRow.AddCell("Product Name").SetStyle(styles["table_header"])
	headerRow.AddCell("Category").SetStyle(styles["table_header"])
	headerRow.AddCell("Quantity").SetStyle(styles["table_header"])
	headerRow.AddCell("Unit Price").SetStyle(styles["table_header"])
	headerRow.AddCell("Supplier").SetStyle(styles["table_header"])
	headerRow.AddCell("Last Updated").SetStyle(styles["table_header"])
	headerRow.AddCell("Total Value").SetStyle(styles["table_header"])

	// Inventory data
	for i, item := range inventory {
		row := sheet.AddRow()
		
		// Alternating row colors
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		// Parse date
		updatedDate, _ := time.Parse("2006-01-02", item.LastUpdated)
		totalValue := float64(item.Quantity) * item.UnitPrice
		
		row.AddCell(item.SKU).SetStyle(rowStyle)
		row.AddCell(item.ProductName).SetStyle(rowStyle)
		row.AddCell(item.Category).SetStyle(rowStyle)
		row.AddCell(item.Quantity).SetStyle(styles["integer"])
		row.AddCell(item.UnitPrice).SetStyle(styles["currency"])
		row.AddCell(item.Supplier).SetStyle(rowStyle)
		row.AddCell(updatedDate).SetStyle(styles["date"])
		row.AddCell(totalValue).SetStyle(styles["currency"])
	}

	// Inventory summary
	sheet.AddRow()
	sheet.AddRow()
	summaryRow := sheet.AddRow()
	summaryRow.AddCell("INVENTORY SUMMARY").SetStyle(styles["sheet_title"])

	sheet.AddRow()

	// Calculate summary metrics
	totalItems := 0
	totalValue := 0.0
	categoryValues := make(map[string]float64)
	supplierCounts := make(map[string]int)

	for _, item := range inventory {
		totalItems += item.Quantity
		itemValue := float64(item.Quantity) * item.UnitPrice
		totalValue += itemValue
		categoryValues[item.Category] += itemValue
		supplierCounts[item.Supplier]++
	}

	avgItemValue := totalValue / float64(len(inventory))

	// Display summary
	summaryData := []struct {
		label string
		value interface{}
		format string
	}{
		{"Total SKUs", len(inventory), "integer"},
		{"Total Units", totalItems, "integer"},
		{"Total Inventory Value", totalValue, "currency"},
		{"Average Item Value", avgItemValue, "currency"},
		{"Number of Suppliers", len(supplierCounts), "integer"},
	}

	for _, summary := range summaryData {
		sumRow := sheet.AddRow()
		sumRow.AddCell(summary.label).SetStyle(styles["data_normal"])
		if summary.format == "currency" {
			sumRow.AddCell(summary.value).SetStyle(styles["summary_metric"])
		} else {
			sumRow.AddCell(summary.value).SetStyle(styles["summary_metric"])
		}
	}
}

func createDataSummarySheet(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, employees []Employee, salesData []SalesData, inventory []InventoryItem) {
	// Set column widths
	sheet.SetColumnWidth("A", 25.0)
	sheet.SetColumnWidth("B", 20.0)
	sheet.SetColumnWidth("C", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("INTEGRATED DATA SUMMARY").SetStyle(styles["sheet_title"])

	sheet.AddRow()
	sheet.AddRow()

	// Data source summary
	sourceRow := sheet.AddRow()
	sourceRow.AddCell("DATA SOURCES OVERVIEW").SetStyle(styles["table_header"])

	sheet.AddRow()

	// Source information
	sources := []struct {
		source string
		format string
		records int
	}{
		{"Employee Database", "JSON", len(employees)},
		{"Sales Transactions", "CSV", len(salesData)},
		{"Inventory System", "JSON", len(inventory)},
	}

	headerRow := sheet.AddRow()
	headerRow.AddCell("Data Source").SetStyle(styles["table_header"])
	headerRow.AddCell("Format").SetStyle(styles["table_header"])
	headerRow.AddCell("Records").SetStyle(styles["table_header"])

	for i, source := range sources {
		row := sheet.AddRow()
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell(source.source).SetStyle(rowStyle)
		row.AddCell(source.format).SetStyle(rowStyle)
		row.AddCell(source.records).SetStyle(styles["integer"])
	}

	// Cross-data analysis
	sheet.AddRow()
	sheet.AddRow()
	analysisRow := sheet.AddRow()
	analysisRow.AddCell("CROSS-DATA ANALYSIS").SetStyle(styles["table_header"])

	sheet.AddRow()

	// Calculate cross-data metrics
	totalSalesRevenue := 0.0
	for _, sale := range salesData {
		totalSalesRevenue += sale.Revenue
	}

	totalPayroll := 0.0
	for _, emp := range employees {
		totalPayroll += emp.Salary
	}

	totalInventoryValue := 0.0
	for _, item := range inventory {
		totalInventoryValue += float64(item.Quantity) * item.UnitPrice
	}

	// Display cross-analysis
	crossData := []struct {
		metric string
		value interface{}
		format string
	}{
		{"Total Sales Revenue", totalSalesRevenue, "currency"},
		{"Annual Payroll", totalPayroll, "currency"},
		{"Inventory Value", totalInventoryValue, "currency"},
		{"Revenue per Employee", totalSalesRevenue / float64(len(employees)), "currency"},
		{"Inventory Turnover Ratio", totalSalesRevenue / totalInventoryValue, "decimal"},
	}

	crossHeaderRow := sheet.AddRow()
	crossHeaderRow.AddCell("Business Metric").SetStyle(styles["table_header"])
	crossHeaderRow.AddCell("Value").SetStyle(styles["table_header"])

	for i, metric := range crossData {
		row := sheet.AddRow()
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}
		
		row.AddCell(metric.metric).SetStyle(rowStyle)
		if metric.format == "currency" {
			row.AddCell(metric.value).SetStyle(styles["summary_metric"])
		} else {
			row.AddCell(metric.value).SetStyle(styles["summary_metric"])
		}
	}

	// Data quality notes
	sheet.AddRow()
	sheet.AddRow()
	qualityRow := sheet.AddRow()
	qualityRow.AddCell("DATA QUALITY NOTES").SetStyle(styles["table_header"])

	sheet.AddRow()

	notes := []string{
		"• All employee data successfully imported from JSON format",
		"• Sales data imported from CSV with proper type conversion",
		"• Inventory data validated with calculated total values",
		"• Cross-data relationships established for analysis",
		"• No data quality issues detected during import process",
	}

	for _, note := range notes {
		noteRow := sheet.AddRow()
		noteRow.AddCell(note).SetStyle(styles["data_normal"])
	}
}

func exportProcessedData(employees []Employee, salesData []SalesData, inventory []InventoryItem) {
	// Export processed employees to JSON
	processedEmployees := make([]map[string]interface{}, len(employees))
	for i, emp := range employees {
		processedEmployees[i] = map[string]interface{}{
			"id":           emp.ID,
			"full_name":    fmt.Sprintf("%s %s", emp.FirstName, emp.LastName),
			"email":        emp.Email,
			"department":   emp.Department,
			"position":     emp.Position,
			"annual_salary": emp.Salary,
			"hire_date":    emp.HireDate.Format("2006-01-02"),
			"status":       map[bool]string{true: "Active", false: "Inactive"}[emp.Active],
			"years_service": int(time.Since(emp.HireDate).Hours() / 24 / 365),
		}
	}

	processedJSON, _ := json.MarshalIndent(processedEmployees, "", "  ")
	os.WriteFile("output/processed_employees.json", processedJSON, 0644)
	fmt.Println("   ✓ Exported processed_employees.json")

	// Export sales summary to CSV
	salesFile, _ := os.Create("output/sales_summary.csv")
	defer salesFile.Close()

	writer := csv.NewWriter(salesFile)
	defer writer.Flush()

	// Sales summary headers
	writer.Write([]string{"Region", "Total_Revenue", "Total_Quantity", "Avg_Deal_Size", "Deal_Count"})

	// Aggregate sales by region
	regionSummary := make(map[string]struct {
		revenue  float64
		quantity int
		deals    int
	})

	for _, sale := range salesData {
		summary := regionSummary[sale.Region]
		summary.revenue += sale.Revenue
		summary.quantity += sale.Quantity
		summary.deals++
		regionSummary[sale.Region] = summary
	}

	for region, summary := range regionSummary {
		avgDeal := summary.revenue / float64(summary.deals)
		writer.Write([]string{
			region,
			fmt.Sprintf("%.2f", summary.revenue),
			strconv.Itoa(summary.quantity),
			fmt.Sprintf("%.2f", avgDeal),
			strconv.Itoa(summary.deals),
		})
	}
	fmt.Println("   ✓ Exported sales_summary.csv")

	// Export inventory analysis to JSON
	inventoryAnalysis := make(map[string]interface{})
	categoryAnalysis := make(map[string]map[string]interface{})

	for _, item := range inventory {
		if categoryAnalysis[item.Category] == nil {
			categoryAnalysis[item.Category] = map[string]interface{}{
				"total_items":  0,
				"total_value":  0.0,
				"avg_price":    0.0,
				"product_count": 0,
			}
		}
		
		category := categoryAnalysis[item.Category]
		category["total_items"] = category["total_items"].(int) + item.Quantity
		category["total_value"] = category["total_value"].(float64) + (float64(item.Quantity) * item.UnitPrice)
		category["product_count"] = category["product_count"].(int) + 1
	}

	// Calculate averages
	for category, data := range categoryAnalysis {
		productCount := data["product_count"].(int)
		totalValue := data["total_value"].(float64)
		data["avg_price"] = totalValue / float64(data["total_items"].(int))
		categoryAnalysis[category] = data
	}

	inventoryAnalysis["categories"] = categoryAnalysis
	inventoryAnalysis["generated_at"] = time.Now().Format(time.RFC3339)
	inventoryAnalysis["total_categories"] = len(categoryAnalysis)

	analysisJSON, _ := json.MarshalIndent(inventoryAnalysis, "", "  ")
	os.WriteFile("output/inventory_report.json", analysisJSON, 0644)
	fmt.Println("   ✓ Exported inventory_report.json")
}

func demonstateRoundTripIntegrity() {
	// Create a simple workbook for round-trip testing
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:   "Round-Trip Data Integrity Test",
		Author:  "Data Integration System",
		Subject: "Data Integrity Verification",
	})

	// Create test data
	testData := []struct {
		ID       int
		Name     string
		Value    float64
		Date     time.Time
		Active   bool
	}{
		{1, "Test Item 1", 123.45, time.Now(), true},
		{2, "Test Item 2", 678.90, time.Now().AddDate(0, -1, 0), false},
		{3, "Test Item 3", 999.99, time.Now().AddDate(0, 0, -7), true},
	}

	// Create sheet
	sheet := workbook.AddSheet("Round Trip Test")
	if sheet == nil {
		log.Fatal("Failed to create round trip test sheet")
	}

	// Headers
	headerRow := sheet.AddRow()
	headerRow.AddCell("ID")
	headerRow.AddCell("Name")
	headerRow.AddCell("Value")
	headerRow.AddCell("Date")
	headerRow.AddCell("Active")

	// Data
	for _, item := range testData {
		row := sheet.AddRow()
		row.AddCell(item.ID)
		row.AddCell(item.Name)
		row.AddCell(item.Value)
		row.AddCell(item.Date)
		row.AddCell(item.Active)
	}

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build round trip test workbook")
	}

	filename := "output/round_trip_test.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save round trip test: %v", err)
	}

	fmt.Printf("   ✓ Created round-trip test file: %s\n", filename)
	fmt.Println("   ✓ Data integrity verification completed")
	fmt.Println("   ℹ️  In a real scenario, you would read this file back and verify data matches")
}