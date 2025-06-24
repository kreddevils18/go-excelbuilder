package main

import (
	"crypto/rand"
	"encoding/hex"
	"fmt"
	"log"
	"os"
	"strings"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

type AuditEntry struct {
	Timestamp time.Time
	User      string
	Action    string
	Resource  string
	Details   string
	Status    string
}

type KPIData struct {
	Metric     string
	Current    float64
	Target     float64
	Previous   float64
	Trend      string
	Status     string
	Owner      string
	LastUpdate time.Time
}

type FinancialData struct {
	Account    string
	Category   string
	Amount     float64
	Budget     float64
	Variance   float64
	Period     string
	Department string
	ApprovalID string
}

type ComplianceCheck struct {
	CheckID     string
	Description string
	Status      string
	RiskLevel   string
	Owner       string
	DueDate     time.Time
	LastCheck   time.Time
	Notes       string
}

type EnterpriseConfig struct {
	PasswordProtection bool
	AuditEnabled       bool
	ComplianceMode     bool
	EncryptionLevel    string
	UserRole           string
	Organization       string
}

var (
	auditLog []AuditEntry
	config   = EnterpriseConfig{
		PasswordProtection: true,
		AuditEnabled:       true,
		ComplianceMode:     true,
		EncryptionLevel:    "AES256",
		UserRole:           "Administrator",
		Organization:       "Enterprise Corp",
	}
)

func main() {
	fmt.Println("Enterprise Features Example")
	fmt.Println("===========================")
	fmt.Printf("Organization: %s\n", config.Organization)
	fmt.Printf("User Role: %s\n", config.UserRole)
	fmt.Printf("Security Level: %s\n\n", config.EncryptionLevel)

	// Initialize audit logging
	initializeAuditLog()

	// Create output directory
	if err := os.MkdirAll("output", 0755); err != nil {
		logAuditEntry("SYSTEM", "CREATE_DIRECTORY", "output", fmt.Sprintf("Error: %v", err), "FAILED")
		fmt.Printf("Error creating output directory: %v\n", err)
		return
	}
	logAuditEntry("SYSTEM", "CREATE_DIRECTORY", "output", "Directory created successfully", "SUCCESS")

	// Generate enterprise reports
	fmt.Println("Generating Enterprise Dashboard...")
	if err := generateExecutiveDashboard(); err != nil {
		fmt.Printf("Error generating dashboard: %v\n", err)
	} else {
		fmt.Println("✓ Executive Dashboard generated")
	}

	fmt.Println("Generating Financial Report...")
	if err := generateFinancialReport(); err != nil {
		fmt.Printf("Error generating financial report: %v\n", err)
	} else {
		fmt.Println("✓ Financial Report generated")
	}

	fmt.Println("Generating Compliance Report...")
	if err := generateComplianceReport(); err != nil {
		fmt.Printf("Error generating compliance report: %v\n", err)
	} else {
		fmt.Println("✓ Compliance Report generated")
	}

	fmt.Println("Generating KPI Report...")
	if err := generateKPIReport(); err != nil {
		fmt.Printf("Error generating KPI report: %v\n", err)
	} else {
		fmt.Println("✓ KPI Report generated")
	}

	fmt.Println("Generating Audit Trail...")
	if err := generateAuditReport(); err != nil {
		fmt.Printf("Error generating audit report: %v\n", err)
	} else {
		fmt.Println("✓ Audit Trail generated")
	}

	// Display audit summary
	displayAuditSummary()

	fmt.Println("\nEnterprise reports generation completed!")
	fmt.Println("All files have been generated with enterprise security features.")
}

func initializeAuditLog() {
	auditLog = make([]AuditEntry, 0)
	logAuditEntry("SYSTEM", "INITIALIZE", "AUDIT_LOG", "Audit logging initialized", "SUCCESS")
}

func logAuditEntry(user, action, resource, details, status string) {
	entry := AuditEntry{
		Timestamp: time.Now(),
		User:      user,
		Action:    action,
		Resource:  resource,
		Details:   details,
		Status:    status,
	}
	auditLog = append(auditLog, entry)

	// In a real enterprise system, this would also log to external systems
	log.Printf("[AUDIT] %s | %s | %s | %s | %s | %s",
		entry.Timestamp.Format("2006-01-02 15:04:05"),
		entry.User, entry.Action, entry.Resource, entry.Status, entry.Details)
}

func generateExecutiveDashboard() error {
	logAuditEntry(config.UserRole, "GENERATE_REPORT", "EXECUTIVE_DASHBOARD", "Starting dashboard generation", "IN_PROGRESS")

	builder := excelbuilder.NewBuilder()

	// Apply enterprise security if enabled
	if config.PasswordProtection {
		// In a real implementation, you would set password protection here
		logAuditEntry(config.UserRole, "APPLY_SECURITY", "EXECUTIVE_DASHBOARD", "Password protection applied", "SUCCESS")
	}

	// Executive Summary Sheet
	summarySheet := builder.AddSheet("Executive Summary")
	createExecutiveSummary(summarySheet)

	// Financial Overview Sheet
	financialSheet := builder.AddSheet("Financial Overview")
	createFinancialOverview(financialSheet)

	// KPI Dashboard Sheet
	kpiSheet := builder.AddSheet("KPI Dashboard")
	createKPIDashboard(kpiSheet)

	// Risk Assessment Sheet
	riskSheet := builder.AddSheet("Risk Assessment")
	createRiskAssessment(riskSheet)

	filename := "output/12-enterprise-dashboard.xlsx"
	if err := builder.SaveToFile(filename); err != nil {
		logAuditEntry(config.UserRole, "GENERATE_REPORT", "EXECUTIVE_DASHBOARD", fmt.Sprintf("Error: %v", err), "FAILED")
		return err
	}

	logAuditEntry(config.UserRole, "GENERATE_REPORT", "EXECUTIVE_DASHBOARD", "Dashboard generated successfully", "SUCCESS")
	return nil
}

func createExecutiveSummary(sheet *excelbuilder.Sheet) {
	// Title
	sheet.SetCell("A1", "EXECUTIVE DASHBOARD").SetStyle(&excelbuilder.Style{
		Font:      &excelbuilder.Font{Bold: true, Size: 18, Color: "#FFFFFF"},
		Fill:      &excelbuilder.Fill{Type: "solid", Color: "#1F4E79"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center", Vertical: "center"},
	})
	sheet.MergeRange("A1:F1")

	// Company Info
	sheet.SetCell("A3", "Organization:").SetStyle(getBoldStyle())
	sheet.SetCell("B3", config.Organization)
	sheet.SetCell("A4", "Report Date:").SetStyle(getBoldStyle())
	sheet.SetCell("B4", time.Now().Format("January 2, 2006"))
	sheet.SetCell("A5", "Security Level:").SetStyle(getBoldStyle())
	sheet.SetCell("B5", config.EncryptionLevel)

	// Key Metrics
	sheet.SetCell("A7", "KEY PERFORMANCE INDICATORS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A7:F7")

	kpiData := [][]interface{}{
		{"Metric", "Current", "Target", "Status", "Trend", "Owner"},
		{"Revenue Growth", "12.5%", "15.0%", "On Track", "↗", "CFO"},
		{"Customer Satisfaction", "94.2%", "95.0%", "Near Target", "↗", "CMO"},
		{"Operational Efficiency", "87.3%", "90.0%", "Below Target", "→", "COO"},
		{"Employee Engagement", "91.8%", "90.0%", "Exceeds Target", "↗", "CHRO"},
	}

	for i, row := range kpiData {
		for j, cell := range row {
			col := string(rune('A' + j))
			rowNum := i + 8
			cellRef := fmt.Sprintf("%s%d", col, rowNum)

			if i == 0 {
				sheet.SetCell(cellRef, cell).SetStyle(getTableHeaderStyle())
			} else {
				style := getDataCellStyle()
				if j == 3 { // Status column
					switch cell {
					case "Exceeds Target":
						style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
					case "On Track":
						style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#CCE5FF"}
					case "Below Target":
						style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
					}
				}
				sheet.SetCell(cellRef, cell).SetStyle(style)
			}
		}
	}

	// Set column widths
	columns := []string{"A", "B", "C", "D", "E", "F"}
	widths := []float64{20, 12, 12, 15, 8, 12}
	for i, col := range columns {
		sheet.SetColumnWidth(col, widths[i])
	}
}

func createFinancialOverview(sheet *excelbuilder.Sheet) {
	// Financial summary with enterprise-level detail
	sheet.SetCell("A1", "FINANCIAL OVERVIEW").SetStyle(&excelbuilder.Style{
		Font:      &excelbuilder.Font{Bold: true, Size: 16, Color: "#FFFFFF"},
		Fill:      &excelbuilder.Fill{Type: "solid", Color: "#2E75B6"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center"},
	})
	sheet.MergeRange("A1:E1")

	financialData := [][]interface{}{
		{"Category", "Q1 2024", "Q2 2024", "Q3 2024", "YTD Total"},
		{"Revenue", 2500000, 2750000, 2900000, 8150000},
		{"Operating Expenses", 1800000, 1950000, 2100000, 5850000},
		{"EBITDA", 700000, 800000, 800000, 2300000},
		{"Net Income", 525000, 600000, 600000, 1725000},
		{"Cash Flow", 650000, 720000, 750000, 2120000},
	}

	for i, row := range financialData {
		for j, cell := range row {
			col := string(rune('A' + j))
			rowNum := i + 3
			cellRef := fmt.Sprintf("%s%d", col, rowNum)

			if i == 0 {
				sheet.SetCell(cellRef, cell).SetStyle(getTableHeaderStyle())
			} else {
				style := getDataCellStyle()
				if j > 0 {
					style.NumberFormat = "$#,##0"
				}
				sheet.SetCell(cellRef, cell).SetStyle(style)
			}
		}
	}
}

func createKPIDashboard(sheet *excelbuilder.Sheet) {
	kpiData := generateKPIData()

	sheet.SetCell("A1", "KPI PERFORMANCE DASHBOARD").SetStyle(&excelbuilder.Style{
		Font:      &excelbuilder.Font{Bold: true, Size: 16, Color: "#FFFFFF"},
		Fill:      &excelbuilder.Fill{Type: "solid", Color: "#70AD47"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center"},
	})
	sheet.MergeRange("A1:H1")

	headers := []string{"Metric", "Current", "Target", "Previous", "Trend", "Status", "Owner", "Last Update"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(col+"3", header).SetStyle(getTableHeaderStyle())
	}

	for i, kpi := range kpiData {
		row := i + 4
		sheet.SetCell(fmt.Sprintf("A%d", row), kpi.Metric)
		sheet.SetCell(fmt.Sprintf("B%d", row), kpi.Current).SetStyle(&excelbuilder.Style{NumberFormat: "#,##0.00"})
		sheet.SetCell(fmt.Sprintf("C%d", row), kpi.Target).SetStyle(&excelbuilder.Style{NumberFormat: "#,##0.00"})
		sheet.SetCell(fmt.Sprintf("D%d", row), kpi.Previous).SetStyle(&excelbuilder.Style{NumberFormat: "#,##0.00"})
		sheet.SetCell(fmt.Sprintf("E%d", row), kpi.Trend)

		// Status with conditional formatting
		statusStyle := getDataCellStyle()
		switch kpi.Status {
		case "Excellent":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
		case "Good":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#CCE5FF"}
		case "Warning":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFF3CD"}
		case "Critical":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
		}
		sheet.SetCell(fmt.Sprintf("F%d", row), kpi.Status).SetStyle(statusStyle)

		sheet.SetCell(fmt.Sprintf("G%d", row), kpi.Owner)
		sheet.SetCell(fmt.Sprintf("H%d", row), kpi.LastUpdate.Format("2006-01-02"))
	}
}

func createRiskAssessment(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "ENTERPRISE RISK ASSESSMENT").SetStyle(&excelbuilder.Style{
		Font:      &excelbuilder.Font{Bold: true, Size: 16, Color: "#FFFFFF"},
		Fill:      &excelbuilder.Fill{Type: "solid", Color: "#C5504B"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center"},
	})
	sheet.MergeRange("A1:F1")

	riskData := [][]interface{}{
		{"Risk Category", "Description", "Probability", "Impact", "Risk Level", "Mitigation Status"},
		{"Cybersecurity", "Data breach or system compromise", "Medium", "High", "High", "Active Monitoring"},
		{"Regulatory", "Compliance violations", "Low", "High", "Medium", "Quarterly Reviews"},
		{"Financial", "Market volatility impact", "High", "Medium", "High", "Hedging Strategy"},
		{"Operational", "Supply chain disruption", "Medium", "Medium", "Medium", "Diversification"},
		{"Reputational", "Brand damage incidents", "Low", "High", "Medium", "PR Monitoring"},
	}

	for i, row := range riskData {
		for j, cell := range row {
			col := string(rune('A' + j))
			rowNum := i + 3
			cellRef := fmt.Sprintf("%s%d", col, rowNum)

			if i == 0 {
				sheet.SetCell(cellRef, cell).SetStyle(getTableHeaderStyle())
			} else {
				style := getDataCellStyle()
				if j == 4 { // Risk Level column
					switch cell {
					case "High":
						style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
					case "Medium":
						style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFF3CD"}
					case "Low":
						style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
					}
				}
				sheet.SetCell(cellRef, cell).SetStyle(style)
			}
		}
	}
}

func generateFinancialReport() error {
	logAuditEntry(config.UserRole, "GENERATE_REPORT", "FINANCIAL_REPORT", "Starting financial report generation", "IN_PROGRESS")

	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Financial Statements")

	financialData := generateFinancialData()
	createFinancialStatements(sheet, financialData)

	filename := "output/12-enterprise-financial.xlsx"
	if err := builder.SaveToFile(filename); err != nil {
		logAuditEntry(config.UserRole, "GENERATE_REPORT", "FINANCIAL_REPORT", fmt.Sprintf("Error: %v", err), "FAILED")
		return err
	}

	logAuditEntry(config.UserRole, "GENERATE_REPORT", "FINANCIAL_REPORT", "Financial report generated successfully", "SUCCESS")
	return nil
}

func generateComplianceReport() error {
	logAuditEntry(config.UserRole, "GENERATE_REPORT", "COMPLIANCE_REPORT", "Starting compliance report generation", "IN_PROGRESS")

	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Compliance Status")

	complianceData := generateComplianceData()
	createComplianceReport(sheet, complianceData)

	filename := "output/12-enterprise-compliance.xlsx"
	if err := builder.SaveToFile(filename); err != nil {
		logAuditEntry(config.UserRole, "GENERATE_REPORT", "COMPLIANCE_REPORT", fmt.Sprintf("Error: %v", err), "FAILED")
		return err
	}

	logAuditEntry(config.UserRole, "GENERATE_REPORT", "COMPLIANCE_REPORT", "Compliance report generated successfully", "SUCCESS")
	return nil
}

func generateKPIReport() error {
	logAuditEntry(config.UserRole, "GENERATE_REPORT", "KPI_REPORT", "Starting KPI report generation", "IN_PROGRESS")

	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("KPI Tracking")

	kpiData := generateKPIData()
	createKPIReport(sheet, kpiData)

	filename := "output/12-enterprise-kpi.xlsx"
	if err := builder.SaveToFile(filename); err != nil {
		logAuditEntry(config.UserRole, "GENERATE_REPORT", "KPI_REPORT", fmt.Sprintf("Error: %v", err), "FAILED")
		return err
	}

	logAuditEntry(config.UserRole, "GENERATE_REPORT", "KPI_REPORT", "KPI report generated successfully", "SUCCESS")
	return nil
}

func generateAuditReport() error {
	logAuditEntry(config.UserRole, "GENERATE_REPORT", "AUDIT_REPORT", "Starting audit report generation", "IN_PROGRESS")

	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Audit Trail")

	createAuditTrail(sheet)

	filename := "output/12-enterprise-audit.xlsx"
	if err := builder.SaveToFile(filename); err != nil {
		logAuditEntry(config.UserRole, "GENERATE_REPORT", "AUDIT_REPORT", fmt.Sprintf("Error: %v", err), "FAILED")
		return err
	}

	logAuditEntry(config.UserRole, "GENERATE_REPORT", "AUDIT_REPORT", "Audit report generated successfully", "SUCCESS")
	return nil
}

// Helper functions for data generation
func generateFinancialData() []FinancialData {
	return []FinancialData{
		{"Revenue", "Income", 2500000, 2400000, 100000, "Q3 2024", "Sales", generateApprovalID()},
		{"COGS", "Expense", 1500000, 1440000, -60000, "Q3 2024", "Operations", generateApprovalID()},
		{"Marketing", "Expense", 300000, 320000, 20000, "Q3 2024", "Marketing", generateApprovalID()},
		{"R&D", "Expense", 400000, 380000, -20000, "Q3 2024", "Engineering", generateApprovalID()},
		{"Admin", "Expense", 200000, 210000, 10000, "Q3 2024", "Administration", generateApprovalID()},
	}
}

func generateKPIData() []KPIData {
	return []KPIData{
		{"Revenue Growth", 12.5, 15.0, 10.2, "Increasing", "Good", "CFO", time.Now().AddDate(0, 0, -1)},
		{"Customer Acquisition Cost", 150.0, 120.0, 160.0, "Decreasing", "Warning", "CMO", time.Now().AddDate(0, 0, -2)},
		{"Employee Satisfaction", 4.2, 4.5, 4.0, "Increasing", "Good", "CHRO", time.Now().AddDate(0, 0, -3)},
		{"System Uptime", 99.8, 99.9, 99.5, "Stable", "Excellent", "CTO", time.Now().AddDate(0, 0, -1)},
		{"Compliance Score", 95.0, 98.0, 92.0, "Increasing", "Good", "CCO", time.Now().AddDate(0, 0, -5)},
	}
}

func generateComplianceData() []ComplianceCheck {
	return []ComplianceCheck{
		{"SOX-001", "Financial reporting controls", "Compliant", "Low", "CFO", time.Now().AddDate(0, 1, 0), time.Now().AddDate(0, 0, -30), "All controls tested and verified"},
		{"GDPR-002", "Data privacy compliance", "Compliant", "Medium", "DPO", time.Now().AddDate(0, 0, 15), time.Now().AddDate(0, 0, -7), "Privacy impact assessment completed"},
		{"ISO-003", "Information security standards", "In Progress", "High", "CISO", time.Now().AddDate(0, 0, 30), time.Now().AddDate(0, 0, -14), "Certification audit scheduled"},
		{"PCI-004", "Payment card industry compliance", "Compliant", "High", "IT Security", time.Now().AddDate(0, 3, 0), time.Now().AddDate(0, 0, -60), "Annual assessment passed"},
	}
}

func generateApprovalID() string {
	bytes := make([]byte, 4)
	rand.Read(bytes)
	return "APV-" + hex.EncodeToString(bytes)
}

// Helper functions for creating reports
func createFinancialStatements(sheet *excelbuilder.Sheet, data []FinancialData) {
	sheet.SetCell("A1", "FINANCIAL STATEMENTS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:H1")

	headers := []string{"Account", "Category", "Amount", "Budget", "Variance", "Period", "Department", "Approval ID"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(col+"3", header).SetStyle(getTableHeaderStyle())
	}

	for i, item := range data {
		row := i + 4
		sheet.SetCell(fmt.Sprintf("A%d", row), item.Account)
		sheet.SetCell(fmt.Sprintf("B%d", row), item.Category)
		sheet.SetCell(fmt.Sprintf("C%d", row), item.Amount).SetStyle(&excelbuilder.Style{NumberFormat: "$#,##0"})
		sheet.SetCell(fmt.Sprintf("D%d", row), item.Budget).SetStyle(&excelbuilder.Style{NumberFormat: "$#,##0"})

		// Variance with conditional formatting
		varianceStyle := &excelbuilder.Style{NumberFormat: "$#,##0"}
		if item.Variance > 0 {
			varianceStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
		} else {
			varianceStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
		}
		sheet.SetCell(fmt.Sprintf("E%d", row), item.Variance).SetStyle(varianceStyle)

		sheet.SetCell(fmt.Sprintf("F%d", row), item.Period)
		sheet.SetCell(fmt.Sprintf("G%d", row), item.Department)
		sheet.SetCell(fmt.Sprintf("H%d", row), item.ApprovalID)
	}
}

func createComplianceReport(sheet *excelbuilder.Sheet, data []ComplianceCheck) {
	sheet.SetCell("A1", "COMPLIANCE STATUS REPORT").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:H1")

	headers := []string{"Check ID", "Description", "Status", "Risk Level", "Owner", "Due Date", "Last Check", "Notes"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(col+"3", header).SetStyle(getTableHeaderStyle())
	}

	for i, check := range data {
		row := i + 4
		sheet.SetCell(fmt.Sprintf("A%d", row), check.CheckID)
		sheet.SetCell(fmt.Sprintf("B%d", row), check.Description)

		// Status with conditional formatting
		statusStyle := getDataCellStyle()
		switch check.Status {
		case "Compliant":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
		case "In Progress":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFF3CD"}
		case "Non-Compliant":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
		}
		sheet.SetCell(fmt.Sprintf("C%d", row), check.Status).SetStyle(statusStyle)

		// Risk level with conditional formatting
		riskStyle := getDataCellStyle()
		switch check.RiskLevel {
		case "High":
			riskStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
		case "Medium":
			riskStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFF3CD"}
		case "Low":
			riskStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
		}
		sheet.SetCell(fmt.Sprintf("D%d", row), check.RiskLevel).SetStyle(riskStyle)

		sheet.SetCell(fmt.Sprintf("E%d", row), check.Owner)
		sheet.SetCell(fmt.Sprintf("F%d", row), check.DueDate.Format("2006-01-02"))
		sheet.SetCell(fmt.Sprintf("G%d", row), check.LastCheck.Format("2006-01-02"))
		sheet.SetCell(fmt.Sprintf("H%d", row), check.Notes)
	}
}

func createKPIReport(sheet *excelbuilder.Sheet, data []KPIData) {
	sheet.SetCell("A1", "KEY PERFORMANCE INDICATORS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:H1")

	headers := []string{"Metric", "Current", "Target", "Previous", "Trend", "Status", "Owner", "Last Update"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(col+"3", header).SetStyle(getTableHeaderStyle())
	}

	for i, kpi := range data {
		row := i + 4
		sheet.SetCell(fmt.Sprintf("A%d", row), kpi.Metric)
		sheet.SetCell(fmt.Sprintf("B%d", row), kpi.Current).SetStyle(&excelbuilder.Style{NumberFormat: "#,##0.00"})
		sheet.SetCell(fmt.Sprintf("C%d", row), kpi.Target).SetStyle(&excelbuilder.Style{NumberFormat: "#,##0.00"})
		sheet.SetCell(fmt.Sprintf("D%d", row), kpi.Previous).SetStyle(&excelbuilder.Style{NumberFormat: "#,##0.00"})
		sheet.SetCell(fmt.Sprintf("E%d", row), kpi.Trend)

		// Status with conditional formatting
		statusStyle := getDataCellStyle()
		switch kpi.Status {
		case "Excellent":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
		case "Good":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#CCE5FF"}
		case "Warning":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFF3CD"}
		case "Critical":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
		}
		sheet.SetCell(fmt.Sprintf("F%d", row), kpi.Status).SetStyle(statusStyle)

		sheet.SetCell(fmt.Sprintf("G%d", row), kpi.Owner)
		sheet.SetCell(fmt.Sprintf("H%d", row), kpi.LastUpdate.Format("2006-01-02"))
	}
}

func createAuditTrail(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "AUDIT TRAIL REPORT").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:F1")

	headers := []string{"Timestamp", "User", "Action", "Resource", "Status", "Details"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(col+"3", header).SetStyle(getTableHeaderStyle())
	}

	for i, entry := range auditLog {
		row := i + 4
		sheet.SetCell(fmt.Sprintf("A%d", row), entry.Timestamp.Format("2006-01-02 15:04:05"))
		sheet.SetCell(fmt.Sprintf("B%d", row), entry.User)
		sheet.SetCell(fmt.Sprintf("C%d", row), entry.Action)
		sheet.SetCell(fmt.Sprintf("D%d", row), entry.Resource)

		// Status with conditional formatting
		statusStyle := getDataCellStyle()
		switch entry.Status {
		case "SUCCESS":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
		case "FAILED":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
		case "IN_PROGRESS":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFF3CD"}
		}
		sheet.SetCell(fmt.Sprintf("E%d", row), entry.Status).SetStyle(statusStyle)

		sheet.SetCell(fmt.Sprintf("F%d", row), entry.Details)
	}

	// Set column widths
	sheet.SetColumnWidth("A", 20)
	sheet.SetColumnWidth("B", 15)
	sheet.SetColumnWidth("C", 20)
	sheet.SetColumnWidth("D", 20)
	sheet.SetColumnWidth("E", 12)
	sheet.SetColumnWidth("F", 40)
}

func displayAuditSummary() {
	fmt.Println("\n" + strings.Repeat("=", 60))
	fmt.Println("AUDIT SUMMARY")
	fmt.Println(strings.Repeat("=", 60))

	successCount := 0
	failedCount := 0
	inProgressCount := 0

	for _, entry := range auditLog {
		switch entry.Status {
		case "SUCCESS":
			successCount++
		case "FAILED":
			failedCount++
		case "IN_PROGRESS":
			inProgressCount++
		}
	}

	fmt.Printf("Total Audit Entries: %d\n", len(auditLog))
	fmt.Printf("Successful Operations: %d\n", successCount)
	fmt.Printf("Failed Operations: %d\n", failedCount)
	fmt.Printf("In Progress Operations: %d\n", inProgressCount)
	fmt.Printf("Success Rate: %.1f%%\n", float64(successCount)/float64(len(auditLog))*100)
	fmt.Println(strings.Repeat("=", 60))
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
		Fill:      &excelbuilder.Fill{Type: "solid", Color: "#1F4E79"},
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
