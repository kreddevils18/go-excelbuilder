package main

import (
	"fmt"
	"math/rand"
	"os"
	"sort"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

type Product struct {
	SKU          string
	Name         string
	Category     string
	Supplier     string
	CostPrice    float64
	SellingPrice float64
	Unit         string
	Barcode      string
	Description  string
	Weight       float64
	Dimensions   string
}

type InventoryItem struct {
	SKU            string
	Location       string
	QuantityOnHand int
	Reserved       int
	Available      int
	ReorderPoint   int
	MaxStock       int
	LastUpdated    time.Time
	BatchNumber    string
	ExpirationDate *time.Time
	SerialNumbers  []string
}

type StockMovement struct {
	ID           string
	SKU          string
	Location     string
	MovementType string // IN, OUT, TRANSFER, ADJUSTMENT
	Quantity     int
	Reference    string
	Timestamp    time.Time
	Operator     string
	Notes        string
	Cost         float64
}

type Supplier struct {
	ID            string
	Name          string
	ContactPerson string
	Email         string
	Phone         string
	Address       string
	PaymentTerms  string
	LeadTime      int
	Rating        float64
}

type ABCAnalysis struct {
	SKU               string
	AnnualValue       float64
	PercentValue      float64
	CumulativePercent float64
	Classification    string
	TurnoverRate      float64
}

type Alert struct {
	Type       string
	SKU        string
	Location   string
	Message    string
	Severity   string
	Timestamp  time.Time
	Status     string
	AssignedTo string
}

func main() {
	fmt.Println("Inventory Management System Example")
	fmt.Println("===================================")

	// Create output directory
	if err := os.MkdirAll("output", 0755); err != nil {
		fmt.Printf("Error creating output directory: %v\n", err)
		return
	}

	// Generate inventory reports
	fmt.Println("Generating Current Inventory Report...")
	if err := generateCurrentInventoryReport(); err != nil {
		fmt.Printf("Error generating current inventory: %v\n", err)
	} else {
		fmt.Println("✓ Current Inventory Report generated")
	}

	fmt.Println("Generating Stock Movement Report...")
	if err := generateMovementReport(); err != nil {
		fmt.Printf("Error generating movement report: %v\n", err)
	} else {
		fmt.Println("✓ Stock Movement Report generated")
	}

	fmt.Println("Generating Inventory Analysis...")
	if err := generateAnalysisReport(); err != nil {
		fmt.Printf("Error generating analysis report: %v\n", err)
	} else {
		fmt.Println("✓ Inventory Analysis generated")
	}

	fmt.Println("Generating Management Reports...")
	if err := generateManagementReports(); err != nil {
		fmt.Printf("Error generating management reports: %v\n", err)
	} else {
		fmt.Println("✓ Management Reports generated")
	}

	fmt.Println("Generating Alert Dashboard...")
	if err := generateAlertDashboard(); err != nil {
		fmt.Printf("Error generating alert dashboard: %v\n", err)
	} else {
		fmt.Println("✓ Alert Dashboard generated")
	}

	fmt.Println("\nInventory management reports completed!")
	fmt.Println("Check the output directory for generated files.")
}

func generateCurrentInventoryReport() error {
	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Current Inventory")

	// Title and metadata
	sheet.SetCell("A1", "CURRENT INVENTORY STATUS").SetStyle(&excelbuilder.Style{
		Font:      &excelbuilder.Font{Bold: true, Size: 16, Color: "#FFFFFF"},
		Fill:      &excelbuilder.Fill{Type: "solid", Color: "#2E75B6"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center"},
	})
	sheet.MergeRange("A1:L1")

	// Report metadata
	sheet.SetCell("A3", "Report Date:").SetStyle(getBoldStyle())
	sheet.SetCell("B3", time.Now().Format("2006-01-02 15:04"))
	sheet.SetCell("A4", "Total SKUs:").SetStyle(getBoldStyle())
	sheet.SetCell("B4", "150")
	sheet.SetCell("A5", "Total Locations:").SetStyle(getBoldStyle())
	sheet.SetCell("B5", "5")

	// Generate sample data
	products := generateProducts()
	inventory := generateInventoryItems(products)

	// Create inventory table
	createInventoryTable(sheet, products, inventory, 7)

	// Summary statistics
	createSummaryStats(sheet, inventory, 7+len(inventory)+3)

	return builder.SaveToFile("output/14-inventory-current.xlsx")
}

func generateMovementReport() error {
	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Stock Movements")

	// Title
	sheet.SetCell("A1", "STOCK MOVEMENT HISTORY").SetStyle(&excelbuilder.Style{
		Font:      &excelbuilder.Font{Bold: true, Size: 16, Color: "#FFFFFF"},
		Fill:      &excelbuilder.Fill{Type: "solid", Color: "#70AD47"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center"},
	})
	sheet.MergeRange("A1:J1")

	// Generate movement data
	movements := generateStockMovements()
	createMovementTable(sheet, movements, 3)

	// Movement summary
	createMovementSummary(sheet, movements, 3+len(movements)+3)

	return builder.SaveToFile("output/14-inventory-movements.xlsx")
}

func generateAnalysisReport() error {
	builder := excelbuilder.NewBuilder()

	// ABC Analysis sheet
	abcSheet := builder.AddSheet("ABC Analysis")
	createABCAnalysisSheet(abcSheet)

	// Turnover Analysis sheet
	turnoverSheet := builder.AddSheet("Turnover Analysis")
	createTurnoverAnalysisSheet(turnoverSheet)

	// Valuation sheet
	valuationSheet := builder.AddSheet("Inventory Valuation")
	createValuationSheet(valuationSheet)

	// Performance Metrics sheet
	metricsSheet := builder.AddSheet("Performance Metrics")
	createPerformanceMetricsSheet(metricsSheet)

	return builder.SaveToFile("output/14-inventory-analysis.xlsx")
}

func generateManagementReports() error {
	builder := excelbuilder.NewBuilder()

	// Executive Summary
	execSheet := builder.AddSheet("Executive Summary")
	createExecutiveSummary(execSheet)

	// Reorder Report
	reorderSheet := builder.AddSheet("Reorder Report")
	createReorderReport(reorderSheet)

	// Supplier Performance
	supplierSheet := builder.AddSheet("Supplier Performance")
	createSupplierPerformance(supplierSheet)

	// Cost Analysis
	costSheet := builder.AddSheet("Cost Analysis")
	createCostAnalysis(costSheet)

	return builder.SaveToFile("output/14-inventory-reports.xlsx")
}

func generateAlertDashboard() error {
	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Alert Dashboard")

	// Title
	sheet.SetCell("A1", "INVENTORY ALERT DASHBOARD").SetStyle(&excelbuilder.Style{
		Font:      &excelbuilder.Font{Bold: true, Size: 16, Color: "#FFFFFF"},
		Fill:      &excelbuilder.Fill{Type: "solid", Color: "#C5504B"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center"},
	})
	sheet.MergeRange("A1:H1")

	// Generate alerts
	alerts := generateAlerts()
	createAlertTable(sheet, alerts, 3)

	// Alert summary
	createAlertSummary(sheet, alerts, 3+len(alerts)+3)

	return builder.SaveToFile("output/14-inventory-alerts.xlsx")
}

// Data generation functions
func generateProducts() []Product {
	products := []Product{
		{"SKU001", "Laptop Computer", "Electronics", "TechCorp", 800.00, 1200.00, "Each", "123456789012", "High-performance laptop", 2.5, "35x25x3 cm"},
		{"SKU002", "Office Chair", "Furniture", "FurnCorp", 150.00, 250.00, "Each", "123456789013", "Ergonomic office chair", 15.0, "60x60x120 cm"},
		{"SKU003", "Printer Paper", "Office Supplies", "PaperCorp", 5.00, 8.00, "Ream", "123456789014", "A4 white paper 500 sheets", 2.5, "21x29.7x5 cm"},
		{"SKU004", "Wireless Mouse", "Electronics", "TechCorp", 25.00, 45.00, "Each", "123456789015", "Bluetooth wireless mouse", 0.1, "10x6x3 cm"},
		{"SKU005", "Desk Lamp", "Furniture", "LightCorp", 35.00, 60.00, "Each", "123456789016", "LED adjustable desk lamp", 1.2, "20x20x40 cm"},
	}

	// Generate more products
	for i := 6; i <= 50; i++ {
		categories := []string{"Electronics", "Furniture", "Office Supplies", "Hardware", "Software"}
		suppliers := []string{"TechCorp", "FurnCorp", "PaperCorp", "HardCorp", "SoftCorp"}
		units := []string{"Each", "Box", "Pack", "Set", "Ream"}

		cost := 10.0 + rand.Float64()*500.0
		price := cost * (1.5 + rand.Float64()*0.8)

		product := Product{
			SKU:          fmt.Sprintf("SKU%03d", i),
			Name:         fmt.Sprintf("Product %d", i),
			Category:     categories[rand.Intn(len(categories))],
			Supplier:     suppliers[rand.Intn(len(suppliers))],
			CostPrice:    cost,
			SellingPrice: price,
			Unit:         units[rand.Intn(len(units))],
			Barcode:      fmt.Sprintf("12345678901%d", i),
			Description:  fmt.Sprintf("Description for product %d", i),
			Weight:       0.1 + rand.Float64()*10.0,
			Dimensions:   fmt.Sprintf("%.0fx%.0fx%.0f cm", 10+rand.Float64()*50, 10+rand.Float64()*50, 5+rand.Float64()*20),
		}
		products = append(products, product)
	}

	return products
}

func generateInventoryItems(products []Product) []InventoryItem {
	locations := []string{"Warehouse A", "Warehouse B", "Store Front", "Distribution Center", "Returns Area"}
	items := make([]InventoryItem, 0)

	for _, product := range products {
		// Each product may be in multiple locations
		numLocations := 1 + rand.Intn(3)
		selectedLocations := make([]string, 0)

		for i := 0; i < numLocations; i++ {
			location := locations[rand.Intn(len(locations))]
			// Avoid duplicates
			found := false
			for _, sel := range selectedLocations {
				if sel == location {
					found = true
					break
				}
			}
			if !found {
				selectedLocations = append(selectedLocations, location)
			}
		}

		for _, location := range selectedLocations {
			quantity := rand.Intn(200) + 10
			reserved := rand.Intn(quantity / 4)
			available := quantity - reserved
			reorderPoint := 20 + rand.Intn(50)
			maxStock := quantity + rand.Intn(100) + 50

			var expDate *time.Time
			if product.Category == "Office Supplies" && rand.Float64() < 0.3 {
				exp := time.Now().AddDate(0, 6+rand.Intn(24), 0)
				expDate = &exp
			}

			item := InventoryItem{
				SKU:            product.SKU,
				Location:       location,
				QuantityOnHand: quantity,
				Reserved:       reserved,
				Available:      available,
				ReorderPoint:   reorderPoint,
				MaxStock:       maxStock,
				LastUpdated:    time.Now().Add(-time.Duration(rand.Intn(72)) * time.Hour),
				BatchNumber:    fmt.Sprintf("B%d%02d", time.Now().Year(), rand.Intn(100)),
				ExpirationDate: expDate,
			}
			items = append(items, item)
		}
	}

	return items
}

func generateStockMovements() []StockMovement {
	movements := make([]StockMovement, 0)
	movementTypes := []string{"IN", "OUT", "TRANSFER", "ADJUSTMENT"}
	operators := []string{"John Smith", "Jane Doe", "Mike Johnson", "Sarah Wilson"}
	references := []string{"PO-2024-001", "SO-2024-002", "ADJ-2024-003", "TRF-2024-004"}

	for i := 1; i <= 100; i++ {
		movementType := movementTypes[rand.Intn(len(movementTypes))]
		quantity := rand.Intn(50) + 1
		if movementType == "OUT" {
			quantity = -quantity
		}

		movement := StockMovement{
			ID:           fmt.Sprintf("MOV%05d", i),
			SKU:          fmt.Sprintf("SKU%03d", rand.Intn(50)+1),
			Location:     []string{"Warehouse A", "Warehouse B", "Store Front"}[rand.Intn(3)],
			MovementType: movementType,
			Quantity:     quantity,
			Reference:    references[rand.Intn(len(references))],
			Timestamp:    time.Now().Add(-time.Duration(rand.Intn(720)) * time.Hour),
			Operator:     operators[rand.Intn(len(operators))],
			Notes:        fmt.Sprintf("Movement %d notes", i),
			Cost:         10.0 + rand.Float64()*100.0,
		}
		movements = append(movements, movement)
	}

	// Sort by timestamp (newest first)
	sort.Slice(movements, func(i, j int) bool {
		return movements[i].Timestamp.After(movements[j].Timestamp)
	})

	return movements
}

func generateAlerts() []Alert {
	alerts := []Alert{
		{"LOW_STOCK", "SKU001", "Warehouse A", "Stock level below reorder point", "HIGH", time.Now().Add(-2 * time.Hour), "OPEN", "John Smith"},
		{"EXPIRING", "SKU003", "Store Front", "Items expiring within 30 days", "MEDIUM", time.Now().Add(-1 * time.Hour), "OPEN", "Jane Doe"},
		{"OVERSTOCK", "SKU005", "Warehouse B", "Stock level exceeds maximum", "LOW", time.Now().Add(-30 * time.Minute), "OPEN", "Mike Johnson"},
		{"ZERO_STOCK", "SKU007", "Distribution Center", "No stock available", "CRITICAL", time.Now().Add(-15 * time.Minute), "OPEN", "Sarah Wilson"},
		{"NEGATIVE_STOCK", "SKU012", "Warehouse A", "Negative stock detected", "CRITICAL", time.Now().Add(-45 * time.Minute), "INVESTIGATING", "John Smith"},
	}

	// Generate more alerts
	for i := 6; i <= 20; i++ {
		alertTypes := []string{"LOW_STOCK", "EXPIRING", "OVERSTOCK", "ZERO_STOCK", "SLOW_MOVING"}
		severities := []string{"LOW", "MEDIUM", "HIGH", "CRITICAL"}
		statuses := []string{"OPEN", "INVESTIGATING", "RESOLVED"}
		assignees := []string{"John Smith", "Jane Doe", "Mike Johnson", "Sarah Wilson"}

		alert := Alert{
			Type:       alertTypes[rand.Intn(len(alertTypes))],
			SKU:        fmt.Sprintf("SKU%03d", rand.Intn(50)+1),
			Location:   []string{"Warehouse A", "Warehouse B", "Store Front"}[rand.Intn(3)],
			Message:    fmt.Sprintf("Alert message %d", i),
			Severity:   severities[rand.Intn(len(severities))],
			Timestamp:  time.Now().Add(-time.Duration(rand.Intn(168)) * time.Hour),
			Status:     statuses[rand.Intn(len(statuses))],
			AssignedTo: assignees[rand.Intn(len(assignees))],
		}
		alerts = append(alerts, alert)
	}

	return alerts
}

// Table creation functions
func createInventoryTable(sheet *excelbuilder.Sheet, products []Product, inventory []InventoryItem, startRow int) {
	headers := []string{"SKU", "Product Name", "Category", "Location", "On Hand", "Reserved", "Available", "Reorder Point", "Max Stock", "Unit Cost", "Total Value", "Last Updated"}

	// Headers
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s%d", col, startRow), header).SetStyle(getTableHeaderStyle())
	}

	// Create product lookup map
	productMap := make(map[string]Product)
	for _, product := range products {
		productMap[product.SKU] = product
	}

	// Data
	for i, item := range inventory {
		row := startRow + i + 1
		product := productMap[item.SKU]
		totalValue := float64(item.QuantityOnHand) * product.CostPrice

		sheet.SetCell(fmt.Sprintf("A%d", row), item.SKU)
		sheet.SetCell(fmt.Sprintf("B%d", row), product.Name)
		sheet.SetCell(fmt.Sprintf("C%d", row), product.Category)
		sheet.SetCell(fmt.Sprintf("D%d", row), item.Location)
		sheet.SetCell(fmt.Sprintf("E%d", row), item.QuantityOnHand)
		sheet.SetCell(fmt.Sprintf("F%d", row), item.Reserved)
		sheet.SetCell(fmt.Sprintf("G%d", row), item.Available)
		sheet.SetCell(fmt.Sprintf("H%d", row), item.ReorderPoint)
		sheet.SetCell(fmt.Sprintf("I%d", row), item.MaxStock)
		sheet.SetCell(fmt.Sprintf("J%d", row), product.CostPrice).SetStyle(&excelbuilder.Style{NumberFormat: "$#,##0.00"})
		sheet.SetCell(fmt.Sprintf("K%d", row), totalValue).SetStyle(&excelbuilder.Style{NumberFormat: "$#,##0.00"})
		sheet.SetCell(fmt.Sprintf("L%d", row), item.LastUpdated.Format("2006-01-02 15:04"))

		// Conditional formatting for low stock
		if item.Available <= item.ReorderPoint {
			for col := 'A'; col <= 'L'; col++ {
				sheet.SetCell(fmt.Sprintf("%c%d", col, row), sheet.GetCell(fmt.Sprintf("%c%d", col, row))).SetStyle(&excelbuilder.Style{
					Fill: &excelbuilder.Fill{Type: "solid", Color: "#FFE6E6"},
				})
			}
		}
	}
}

func createMovementTable(sheet *excelbuilder.Sheet, movements []StockMovement, startRow int) {
	headers := []string{"Movement ID", "SKU", "Location", "Type", "Quantity", "Reference", "Timestamp", "Operator", "Cost", "Notes"}

	// Headers
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s%d", col, startRow), header).SetStyle(getTableHeaderStyle())
	}

	// Data
	for i, movement := range movements {
		row := startRow + i + 1
		sheet.SetCell(fmt.Sprintf("A%d", row), movement.ID)
		sheet.SetCell(fmt.Sprintf("B%d", row), movement.SKU)
		sheet.SetCell(fmt.Sprintf("C%d", row), movement.Location)
		sheet.SetCell(fmt.Sprintf("D%d", row), movement.MovementType)
		sheet.SetCell(fmt.Sprintf("E%d", row), movement.Quantity)
		sheet.SetCell(fmt.Sprintf("F%d", row), movement.Reference)
		sheet.SetCell(fmt.Sprintf("G%d", row), movement.Timestamp.Format("2006-01-02 15:04"))
		sheet.SetCell(fmt.Sprintf("H%d", row), movement.Operator)
		sheet.SetCell(fmt.Sprintf("I%d", row), movement.Cost).SetStyle(&excelbuilder.Style{NumberFormat: "$#,##0.00"})
		sheet.SetCell(fmt.Sprintf("J%d", row), movement.Notes)

		// Color coding by movement type
		var fillColor string
		switch movement.MovementType {
		case "IN":
			fillColor = "#E6F3E6"
		case "OUT":
			fillColor = "#FFE6E6"
		case "TRANSFER":
			fillColor = "#E6F0FF"
		case "ADJUSTMENT":
			fillColor = "#FFF0E6"
		}

		if fillColor != "" {
			sheet.SetCell(fmt.Sprintf("D%d", row), movement.MovementType).SetStyle(&excelbuilder.Style{
				Fill: &excelbuilder.Fill{Type: "solid", Color: fillColor},
			})
		}
	}
}

func createAlertTable(sheet *excelbuilder.Sheet, alerts []Alert, startRow int) {
	headers := []string{"Type", "SKU", "Location", "Message", "Severity", "Timestamp", "Status", "Assigned To"}

	// Headers
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s%d", col, startRow), header).SetStyle(getTableHeaderStyle())
	}

	// Data
	for i, alert := range alerts {
		row := startRow + i + 1
		sheet.SetCell(fmt.Sprintf("A%d", row), alert.Type)
		sheet.SetCell(fmt.Sprintf("B%d", row), alert.SKU)
		sheet.SetCell(fmt.Sprintf("C%d", row), alert.Location)
		sheet.SetCell(fmt.Sprintf("D%d", row), alert.Message)
		sheet.SetCell(fmt.Sprintf("F%d", row), alert.Timestamp.Format("2006-01-02 15:04"))
		sheet.SetCell(fmt.Sprintf("G%d", row), alert.Status)
		sheet.SetCell(fmt.Sprintf("H%d", row), alert.AssignedTo)

		// Severity with color coding
		severityStyle := getDataCellStyle()
		switch alert.Severity {
		case "CRITICAL":
			severityStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FF0000"}
			severityStyle.Font = &excelbuilder.Font{Color: "#FFFFFF", Bold: true}
		case "HIGH":
			severityStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FF6666"}
		case "MEDIUM":
			severityStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFCC66"}
		case "LOW":
			severityStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#66FF66"}
		}
		sheet.SetCell(fmt.Sprintf("E%d", row), alert.Severity).SetStyle(severityStyle)
	}
}

// Summary and analysis functions
func createSummaryStats(sheet *excelbuilder.Sheet, inventory []InventoryItem, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "INVENTORY SUMMARY").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:D%d", startRow, startRow))

	// Calculate statistics
	totalItems := 0
	lowStockItems := 0
	locations := make(map[string]bool)

	for _, item := range inventory {
		totalItems += item.QuantityOnHand
		if item.Available <= item.ReorderPoint {
			lowStockItems++
		}
		locations[item.Location] = true
	}

	statsData := [][]interface{}{
		{"Metric", "Value", "Unit", "Status"},
		{"Total SKUs", len(inventory), "items", "Active"},
		{"Total Quantity", totalItems, "units", "Current"},
		{"Low Stock Items", lowStockItems, "items", "Alert"},
		{"Active Locations", len(locations), "locations", "Operational"},
		{"Stock Coverage", fmt.Sprintf("%.1f%%", float64(len(inventory)-lowStockItems)/float64(len(inventory))*100), "%", "Good"},
	}

	for i, row := range statsData {
		rowNum := startRow + i + 2
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

func createMovementSummary(sheet *excelbuilder.Sheet, movements []StockMovement, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "MOVEMENT SUMMARY").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:E%d", startRow, startRow))

	// Calculate movement statistics
	movementStats := make(map[string]int)
	totalCost := 0.0

	for _, movement := range movements {
		movementStats[movement.MovementType]++
		totalCost += movement.Cost
	}

	summaryData := [][]interface{}{
		{"Movement Type", "Count", "Percentage", "Average Cost", "Total Cost"},
	}

	for movType, count := range movementStats {
		percentage := float64(count) / float64(len(movements)) * 100
		avgCost := totalCost / float64(len(movements))
		typeCost := avgCost * float64(count)

		summaryData = append(summaryData, []interface{}{
			movType, count, fmt.Sprintf("%.1f%%", percentage), avgCost, typeCost,
		})
	}

	for i, row := range summaryData {
		rowNum := startRow + i + 2
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			} else if j >= 3 {
				style.NumberFormat = "$#,##0.00"
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createAlertSummary(sheet *excelbuilder.Sheet, alerts []Alert, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "ALERT SUMMARY").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:D%d", startRow, startRow))

	// Calculate alert statistics
	severityStats := make(map[string]int)
	statusStats := make(map[string]int)

	for _, alert := range alerts {
		severityStats[alert.Severity]++
		statusStats[alert.Status]++
	}

	// Severity breakdown
	sheet.SetCell(fmt.Sprintf("A%d", startRow+2), "By Severity:").SetStyle(getBoldStyle())
	row := startRow + 3
	for severity, count := range severityStats {
		sheet.SetCell(fmt.Sprintf("A%d", row), severity)
		sheet.SetCell(fmt.Sprintf("B%d", row), count)
		sheet.SetCell(fmt.Sprintf("C%d", row), fmt.Sprintf("%.1f%%", float64(count)/float64(len(alerts))*100))
		row++
	}

	// Status breakdown
	sheet.SetCell(fmt.Sprintf("A%d", row+1), "By Status:").SetStyle(getBoldStyle())
	row += 2
	for status, count := range statusStats {
		sheet.SetCell(fmt.Sprintf("A%d", row), status)
		sheet.SetCell(fmt.Sprintf("B%d", row), count)
		sheet.SetCell(fmt.Sprintf("C%d", row), fmt.Sprintf("%.1f%%", float64(count)/float64(len(alerts))*100))
		row++
	}
}

// Analysis sheet functions
func createABCAnalysisSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "ABC ANALYSIS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:G1")

	// Generate ABC analysis data
	abcData := generateABCAnalysis()
	createABCTable(sheet, abcData, 3)
}

func createTurnoverAnalysisSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "INVENTORY TURNOVER ANALYSIS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:F1")

	// Turnover data
	turnoverData := [][]interface{}{
		{"SKU", "Annual Sales", "Average Inventory", "Turnover Ratio", "Days in Inventory", "Classification"},
		{"SKU001", 120000, 15000, 8.0, 45.6, "Fast Moving"},
		{"SKU002", 75000, 12500, 6.0, 60.8, "Fast Moving"},
		{"SKU003", 45000, 15000, 3.0, 121.7, "Medium Moving"},
		{"SKU004", 25000, 12500, 2.0, 182.5, "Slow Moving"},
		{"SKU005", 15000, 20000, 0.75, 486.7, "Very Slow Moving"},
	}

	for i, row := range turnoverData {
		rowNum := i + 3
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			} else if j >= 1 && j <= 4 {
				style.NumberFormat = "#,##0.00"
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createValuationSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "INVENTORY VALUATION").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:F1")

	// Valuation methods
	valuationData := [][]interface{}{
		{"Method", "Total Value", "Variance from FIFO", "Percentage", "Recommended", "Notes"},
		{"FIFO", 1250000, 0, "100.0%", "Yes", "First In, First Out"},
		{"LIFO", 1180000, -70000, "94.4%", "No", "Last In, First Out"},
		{"Weighted Average", 1215000, -35000, "97.2%", "Maybe", "Average Cost Method"},
		{"Standard Cost", 1200000, -50000, "96.0%", "No", "Predetermined Costs"},
		{"Market Value", 1350000, 100000, "108.0%", "No", "Current Market Prices"},
	}

	for i, row := range valuationData {
		rowNum := i + 3
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			} else if j == 1 || j == 2 {
				style.NumberFormat = "$#,##0"
			} else if j == 4 && cell == "Yes" {
				style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createPerformanceMetricsSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "PERFORMANCE METRICS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:E1")

	metricsData := [][]interface{}{
		{"Metric", "Current", "Target", "Variance", "Status"},
		{"Inventory Turnover", 6.5, 8.0, -1.5, "Below Target"},
		{"Fill Rate", "98.5%", "99.0%", "-0.5%", "Close to Target"},
		{"Stockout Rate", "2.1%", "1.0%", "+1.1%", "Above Target"},
		{"Carrying Cost", "15.2%", "12.0%", "+3.2%", "Above Target"},
		{"Order Accuracy", "99.8%", "99.5%", "+0.3%", "Above Target"},
		{"Cycle Count Accuracy", "97.5%", "98.0%", "-0.5%", "Below Target"},
	}

	for i, row := range metricsData {
		rowNum := i + 3
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			} else if j == 4 {
				if cell == "Above Target" {
					style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
				} else if cell == "Below Target" {
					style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
				} else {
					style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFF3CD"}
				}
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

// Management report functions
func createExecutiveSummary(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "EXECUTIVE SUMMARY").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:F1")

	// Key metrics
	keyMetrics := [][]interface{}{
		{"Key Performance Indicators", "Current Month", "Previous Month", "Change", "YTD", "Target"},
		{"Total Inventory Value", "$1,250,000", "$1,180,000", "+5.9%", "$1,250,000", "$1,200,000"},
		{"Inventory Turnover", "6.5", "6.2", "+4.8%", "6.5", "8.0"},
		{"Days Sales Outstanding", "45", "48", "-6.3%", "45", "40"},
		{"Stockout Incidents", "12", "18", "-33.3%", "156", "120"},
		{"Carrying Cost %", "15.2%", "15.8%", "-3.8%", "15.2%", "12.0%"},
	}

	for i, row := range keyMetrics {
		rowNum := i + 3
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

func createReorderReport(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "REORDER RECOMMENDATIONS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:H1")

	reorderData := [][]interface{}{
		{"SKU", "Product", "Current Stock", "Reorder Point", "Suggested Order", "Supplier", "Lead Time", "Priority"},
		{"SKU001", "Laptop Computer", 15, 25, 50, "TechCorp", "7 days", "HIGH"},
		{"SKU003", "Printer Paper", 45, 50, 100, "PaperCorp", "3 days", "MEDIUM"},
		{"SKU007", "Monitor Stand", 8, 15, 30, "FurnCorp", "10 days", "HIGH"},
		{"SKU012", "USB Cable", 25, 30, 75, "TechCorp", "5 days", "MEDIUM"},
		{"SKU018", "Desk Organizer", 12, 20, 40, "FurnCorp", "7 days", "LOW"},
	}

	for i, row := range reorderData {
		rowNum := i + 3
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			} else if j == 7 {
				switch cell {
				case "HIGH":
					style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
				case "MEDIUM":
					style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFF3CD"}
				case "LOW":
					style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
				}
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createSupplierPerformance(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "SUPPLIER PERFORMANCE").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:G1")

	supplierData := [][]interface{}{
		{"Supplier", "On-Time Delivery", "Quality Rating", "Cost Performance", "Lead Time", "Overall Score", "Status"},
		{"TechCorp", "95.2%", "4.8/5.0", "Good", "7 days", "4.6/5.0", "Preferred"},
		{"FurnCorp", "88.5%", "4.2/5.0", "Fair", "10 days", "4.1/5.0", "Approved"},
		{"PaperCorp", "98.1%", "4.9/5.0", "Excellent", "3 days", "4.9/5.0", "Preferred"},
		{"HardCorp", "82.3%", "3.8/5.0", "Poor", "14 days", "3.5/5.0", "Review"},
		{"SoftCorp", "91.7%", "4.5/5.0", "Good", "5 days", "4.4/5.0", "Approved"},
	}

	for i, row := range supplierData {
		rowNum := i + 3
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			} else if j == 6 {
				switch cell {
				case "Preferred":
					style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
				case "Approved":
					style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFF3CD"}
				case "Review":
					style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
				}
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createCostAnalysis(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "COST ANALYSIS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:F1")

	costData := [][]interface{}{
		{"Cost Category", "Current Month", "Previous Month", "Change", "YTD", "Budget"},
		{"Carrying Costs", "$18,750", "$17,200", "+9.0%", "$210,000", "$200,000"},
		{"Ordering Costs", "$5,200", "$4,800", "+8.3%", "$58,000", "$55,000"},
		{"Stockout Costs", "$3,400", "$5,100", "-33.3%", "$45,000", "$30,000"},
		{"Storage Costs", "$12,500", "$12,500", "0.0%", "$150,000", "$145,000"},
		{"Obsolescence", "$2,100", "$1,800", "+16.7%", "$22,000", "$20,000"},
		{"Total Costs", "$41,950", "$41,400", "+1.3%", "$485,000", "$450,000"},
	}

	for i, row := range costData {
		rowNum := i + 3
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			} else if j >= 1 && j <= 2 || j >= 4 && j <= 5 {
				style.NumberFormat = "$#,##0"
			} else if i == len(costData)-1 {
				style.Font = &excelbuilder.Font{Bold: true}
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

// Helper functions
func generateABCAnalysis() []ABCAnalysis {
	data := []ABCAnalysis{
		{"SKU001", 120000, 25.5, 25.5, "A", 8.0},
		{"SKU002", 95000, 20.2, 45.7, "A", 6.3},
		{"SKU003", 75000, 15.9, 61.6, "A", 5.0},
		{"SKU004", 65000, 13.8, 75.4, "B", 4.3},
		{"SKU005", 45000, 9.6, 85.0, "B", 3.0},
		{"SKU006", 35000, 7.4, 92.4, "B", 2.3},
		{"SKU007", 20000, 4.3, 96.7, "C", 1.3},
		{"SKU008", 15000, 3.2, 99.9, "C", 1.0},
	}

	return data
}

func createABCTable(sheet *excelbuilder.Sheet, data []ABCAnalysis, startRow int) {
	headers := []string{"SKU", "Annual Value", "% of Total", "Cumulative %", "Classification", "Turnover Rate", "Strategy"}

	// Headers
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s%d", col, startRow), header).SetStyle(getTableHeaderStyle())
	}

	// Data
	for i, item := range data {
		row := startRow + i + 1
		sheet.SetCell(fmt.Sprintf("A%d", row), item.SKU)
		sheet.SetCell(fmt.Sprintf("B%d", row), item.AnnualValue).SetStyle(&excelbuilder.Style{NumberFormat: "$#,##0"})
		sheet.SetCell(fmt.Sprintf("C%d", row), item.PercentValue).SetStyle(&excelbuilder.Style{NumberFormat: "0.0%"})
		sheet.SetCell(fmt.Sprintf("D%d", row), item.CumulativePercent).SetStyle(&excelbuilder.Style{NumberFormat: "0.0%"})
		sheet.SetCell(fmt.Sprintf("F%d", row), item.TurnoverRate).SetStyle(&excelbuilder.Style{NumberFormat: "0.0"})

		// Classification with color coding
		classStyle := getDataCellStyle()
		var strategy string
		switch item.Classification {
		case "A":
			classStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FF6B6B"}
			strategy = "Tight Control"
		case "B":
			classStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFE66D"}
			strategy = "Moderate Control"
		case "C":
			classStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#95E1D3"}
			strategy = "Simple Control"
		}
		sheet.SetCell(fmt.Sprintf("E%d", row), item.Classification).SetStyle(classStyle)
		sheet.SetCell(fmt.Sprintf("G%d", row), strategy)
	}
}

// Style helper functions
func getBoldStyle() *excelbuilder.Style {
	return &excelbuilder.Style{
		Font: &excelbuilder.Font{Bold: true},
	}
}

func getSectionHeaderStyle() *excelbuilder.Style {
	return &excelbuilder.Style{
		Font:      &excelbuilder.Font{Bold: true, Size: 14, Color: "#FFFFFF"},
		Fill:      &excelbuilder.Fill{Type: "solid", Color: "#2E75B6"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center", Vertical: "center"},
	}
}

func getTableHeaderStyle() *excelbuilder.Style {
	return &excelbuilder.Style{
		Font:      &excelbuilder.Font{Bold: true, Color: "#FFFFFF"},
		Fill:      &excelbuilder.Fill{Type: "solid", Color: "#4472C4"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center", Vertical: "center"},
		Border: &excelbuilder.Border{
			Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
			Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
			Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
			Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
		},
	}
}

func getDataCellStyle() *excelbuilder.Style {
	return &excelbuilder.Style{
		Alignment: &excelbuilder.Alignment{Horizontal: "left", Vertical: "center"},
		Border: &excelbuilder.Border{
			Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
			Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
			Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
			Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
		},
	}
}
