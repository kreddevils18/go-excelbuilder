package main

import (
	"fmt"
	"log"
	"math/rand"
	"os"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

func main() {
	// Create output directory
	if err := os.MkdirAll("output", 0755); err != nil {
		log.Fatalf("Failed to create output directory: %v", err)
	}

	// Create Excel builder
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Set workbook properties
	workbook.SetProperties(excelbuilder.WorkbookProperties{
		Title:       "Advanced Styling Demonstration",
		Author:      "Go Excel Builder",
		Subject:     "Advanced Styling Techniques",
		Description: "Comprehensive demonstration of advanced styling capabilities including themes, gradients, and conditional formatting",
		Keywords:    "go,excel,styling,advanced,themes,conditional-formatting",
	})

	// Define color palette for consistent theming
	colorPalette := map[string]string{
		"primary":   "#2F5597",
		"secondary": "#4472C4",
		"accent":    "#70AD47",
		"warning":   "#FFC000",
		"danger":    "#C5504B",
		"success":   "#70AD47",
		"info":      "#5B9BD5",
		"light":     "#F2F2F2",
		"dark":      "#404040",
		"white":     "#FFFFFF",
		"black":     "#000000",
	}

	// Define advanced style configurations
	styles := createAdvancedStyles(colorPalette)

	// Sheet 1: Typography and Font Styling
	typographySheet := workbook.AddSheet("Typography")
	if typographySheet == nil {
		log.Fatal("Failed to create typography sheet")
	}

	createTypographyDemo(typographySheet, styles, colorPalette)

	// Sheet 2: Color Schemes and Themes
	colorSheet := workbook.AddSheet("Color Schemes")
	if colorSheet == nil {
		log.Fatal("Failed to create color sheet")
	}

	createColorSchemeDemo(colorSheet, styles, colorPalette)

	// Sheet 3: Border and Fill Patterns
	borderSheet := workbook.AddSheet("Borders & Fills")
	if borderSheet == nil {
		log.Fatal("Failed to create border sheet")
	}

	createBorderFillDemo(borderSheet, styles, colorPalette)

	// Sheet 4: Data Visualization Styles
	dataVizSheet := workbook.AddSheet("Data Visualization")
	if dataVizSheet == nil {
		log.Fatal("Failed to create data visualization sheet")
	}

	createDataVisualizationDemo(dataVizSheet, styles, colorPalette)

	// Sheet 5: Conditional Formatting Simulation
	conditionalSheet := workbook.AddSheet("Conditional Styling")
	if conditionalSheet == nil {
		log.Fatal("Failed to create conditional sheet")
	}

	createConditionalFormattingDemo(conditionalSheet, styles, colorPalette)

	// Sheet 6: Professional Report Layout
	reportSheet := workbook.AddSheet("Professional Report")
	if reportSheet == nil {
		log.Fatal("Failed to create report sheet")
	}

	createProfessionalReportDemo(reportSheet, styles, colorPalette)

	// Build and save
	file := workbook.Build()
	if file == nil {
		log.Fatal("Failed to build workbook")
	}

	filename := "output/03-styling-advanced-demo.xlsx"
	err := file.SaveAs(filename)
	if err != nil {
		log.Fatalf("Failed to save workbook: %v", err)
	}

	fmt.Printf("✅ Advanced styling example created successfully!\n")
	fmt.Printf("📁 File saved as: %s\n", filename)
	fmt.Printf("🎨 Features demonstrated:\n")
	fmt.Printf("   • Advanced typography and font styling\n")
	fmt.Printf("   • Professional color schemes and themes\n")
	fmt.Printf("   • Comprehensive border and fill patterns\n")
	fmt.Printf("   • Data visualization styling techniques\n")
	fmt.Printf("   • Conditional formatting simulation\n")
	fmt.Printf("   • Professional report layouts\n")
	fmt.Printf("   • Consistent theming across sheets\n")
	fmt.Printf("\n🎯 Next steps: Try examples/04-sales-report/ for real-world application\n")
}

func createAdvancedStyles(colors map[string]string) map[string]excelbuilder.StyleConfig {
	styles := make(map[string]excelbuilder.StyleConfig)

	// Title styles
	styles["title"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   18,
			Color:  colors["primary"],
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	}

	styles["subtitle"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   14,
			Color:  colors["secondary"],
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	}

	// Header styles with different themes
	styles["header_primary"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
			Color:  colors["white"],
			Family: "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["primary"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["dark"]),
	}

	styles["header_success"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
			Color:  colors["white"],
			Family: "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["success"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["dark"]),
	}

	styles["header_warning"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
			Color:  colors["dark"],
			Family: "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["warning"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["dark"]),
	}

	styles["header_danger"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   12,
			Color:  colors["white"],
			Family: "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["danger"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["dark"]),
	}

	// Data styles
	styles["data_normal"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["light"]),
	}

	styles["data_alternate"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#F8F9FA",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "left",
			Vertical:   "middle",
		},
		Border: createBorder("thin", colors["light"]),
	}

	// Number styles with different formatting
	styles["number_currency"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "$#,##0.00",
		Border:       createBorder("thin", colors["light"]),
	}

	styles["number_percentage"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "0.00%",
		Border:       createBorder("thin", colors["light"]),
	}

	// Conditional formatting styles
	styles["high_value"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   10,
			Color:  colors["white"],
			Family: "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["success"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "$#,##0.00",
		Border:       createBorder("thin", colors["dark"]),
	}

	styles["medium_value"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   10,
			Color:  colors["dark"],
			Family: "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["warning"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "$#,##0.00",
		Border:       createBorder("thin", colors["dark"]),
	}

	styles["low_value"] = excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   10,
			Color:  colors["white"],
			Family: "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: colors["danger"],
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "right",
			Vertical:   "middle",
		},
		NumberFormat: "$#,##0.00",
		Border:       createBorder("thin", colors["dark"]),
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

func createTypographyDemo(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, colors map[string]string) {
	// Set column widths
	sheet.SetColumnWidth("A", 25.0)
	sheet.SetColumnWidth("B", 30.0)
	sheet.SetColumnWidth("C", 20.0)
	sheet.SetColumnWidth("D", 15.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Typography Demonstration").SetStyle(styles["title"])

	// Empty row for spacing
	sheet.AddRow()

	// Header
	headerRow := sheet.AddRow()
	headerRow.AddCell("Font Style").SetStyle(styles["header_primary"])
	headerRow.AddCell("Sample Text").SetStyle(styles["header_primary"])
	headerRow.AddCell("Use Case").SetStyle(styles["header_primary"])
	headerRow.AddCell("Size").SetStyle(styles["header_primary"])

	// Typography examples
	typographyExamples := []struct {
		style   string
		text    string
		useCase string
		size    string
		custom  *excelbuilder.StyleConfig
	}{
		{"Normal Text", "The quick brown fox jumps over the lazy dog", "Body text, descriptions", "10pt", nil},
		{"Bold Text", "Important Information", "Emphasis, labels", "10pt", &excelbuilder.StyleConfig{
			Font:      excelbuilder.FontConfig{Bold: true, Size: 10, Family: "Arial"},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "left", Vertical: "middle"},
			Border:    createBorder("thin", colors["light"]),
		}},
		{"Italic Text", "Subtle emphasis or notes", "Comments, annotations", "10pt", &excelbuilder.StyleConfig{
			Font:      excelbuilder.FontConfig{Italic: true, Size: 10, Family: "Arial"},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "left", Vertical: "middle"},
			Border:    createBorder("thin", colors["light"]),
		}},
		{"Large Header", "Section Title", "Major headings", "14pt", &excelbuilder.StyleConfig{
			Font:      excelbuilder.FontConfig{Bold: true, Size: 14, Family: "Arial", Color: colors["primary"]},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "left", Vertical: "middle"},
			Border:    createBorder("thin", colors["light"]),
		}},
		{"Small Text", "Fine print or metadata", "Footnotes, details", "8pt", &excelbuilder.StyleConfig{
			Font:      excelbuilder.FontConfig{Size: 8, Family: "Arial", Color: colors["dark"]},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "left", Vertical: "middle"},
			Border:    createBorder("thin", colors["light"]),
		}},
		{"Colored Text", "Attention-grabbing content", "Warnings, highlights", "10pt", &excelbuilder.StyleConfig{
			Font:      excelbuilder.FontConfig{Bold: true, Size: 10, Family: "Arial", Color: colors["danger"]},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "left", Vertical: "middle"},
			Border:    createBorder("thin", colors["light"]),
		}},
	}

	for i, example := range typographyExamples {
		row := sheet.AddRow()
		row.AddCell(example.style).SetStyle(styles["data_normal"])

		if example.custom != nil {
			row.AddCell(example.text).SetStyle(*example.custom)
		} else {
			if i%2 == 0 {
				row.AddCell(example.text).SetStyle(styles["data_normal"])
			} else {
				row.AddCell(example.text).SetStyle(styles["data_alternate"])
			}
		}

		row.AddCell(example.useCase).SetStyle(styles["data_normal"])
		row.AddCell(example.size).SetStyle(styles["data_normal"])
	}
}

func createColorSchemeDemo(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, colors map[string]string) {
	// Set column widths
	sheet.SetColumnWidth("A", 20.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 25.0)
	sheet.SetColumnWidth("D", 20.0)
	sheet.SetColumnWidth("E", 20.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Color Schemes & Themes").SetStyle(styles["title"])

	// Empty row
	sheet.AddRow()

	// Header
	headerRow := sheet.AddRow()
	headerRow.AddCell("Color Name").SetStyle(styles["header_primary"])
	headerRow.AddCell("Hex Code").SetStyle(styles["header_primary"])
	headerRow.AddCell("Sample Text").SetStyle(styles["header_primary"])
	headerRow.AddCell("Background Sample").SetStyle(styles["header_primary"])
	headerRow.AddCell("Use Case").SetStyle(styles["header_primary"])

	// Color examples
	colorExamples := []struct {
		name    string
		hex     string
		useCase string
	}{
		{"Primary", colors["primary"], "Main brand color, headers"},
		{"Secondary", colors["secondary"], "Supporting elements"},
		{"Success", colors["success"], "Positive values, completion"},
		{"Warning", colors["warning"], "Caution, pending items"},
		{"Danger", colors["danger"], "Errors, negative values"},
		{"Info", colors["info"], "Information, neutral data"},
		{"Light", colors["light"], "Backgrounds, subtle elements"},
		{"Dark", colors["dark"], "Text, borders"},
	}

	for _, example := range colorExamples {
		row := sheet.AddRow()
		row.AddCell(example.name).SetStyle(styles["data_normal"])
		row.AddCell(example.hex).SetStyle(styles["data_normal"])

		// Sample text with color
		textStyle := excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Bold:   true,
				Size:   10,
				Color:  example.hex,
				Family: "Arial",
			},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "center", Vertical: "middle"},
			Border:    createBorder("thin", colors["light"]),
		}
		row.AddCell("Sample Text").SetStyle(textStyle)

		// Background sample
		bgStyle := excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Bold:   true,
				Size:   10,
				Color:  colors["white"],
				Family: "Arial",
			},
			Fill: excelbuilder.FillConfig{
				Type:  "pattern",
				Color: example.hex,
			},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "center", Vertical: "middle"},
			Border:    createBorder("thin", colors["dark"]),
		}
		row.AddCell("Background").SetStyle(bgStyle)

		row.AddCell(example.useCase).SetStyle(styles["data_normal"])
	}
}

func createBorderFillDemo(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, colors map[string]string) {
	// Set column widths
	sheet.SetColumnWidth("A", 20.0)
	sheet.SetColumnWidth("B", 25.0)
	sheet.SetColumnWidth("C", 20.0)
	sheet.SetColumnWidth("D", 25.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Borders & Fill Patterns").SetStyle(styles["title"])

	// Empty row
	sheet.AddRow()

	// Border examples section
	subtitleRow := sheet.AddRow()
	subtitleRow.AddCell("Border Styles").SetStyle(styles["subtitle"])

	headerRow := sheet.AddRow()
	headerRow.AddCell("Style Name").SetStyle(styles["header_primary"])
	headerRow.AddCell("Sample").SetStyle(styles["header_primary"])
	headerRow.AddCell("Description").SetStyle(styles["header_primary"])
	headerRow.AddCell("Use Case").SetStyle(styles["header_primary"])

	// Border style examples
	borderStyles := []struct {
		name        string
		description string
		useCase     string
		style       string
	}{
		{"Thin", "Light border for data tables", "Regular data, subtle separation", "thin"},
		{"Medium", "Standard border for emphasis", "Headers, important sections", "medium"},
		{"Thick", "Heavy border for strong emphasis", "Titles, major divisions", "thick"},
		{"Double", "Double line for special emphasis", "Totals, summary sections", "double"},
	}

	for _, border := range borderStyles {
		row := sheet.AddRow()
		row.AddCell(border.name).SetStyle(styles["data_normal"])

		// Sample with specific border
		sampleStyle := excelbuilder.StyleConfig{
			Font:      excelbuilder.FontConfig{Size: 10, Family: "Arial"},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "center", Vertical: "middle"},
			Border:    createBorder(border.style, colors["dark"]),
		}
		row.AddCell("Sample").SetStyle(sampleStyle)

		row.AddCell(border.description).SetStyle(styles["data_normal"])
		row.AddCell(border.useCase).SetStyle(styles["data_normal"])
	}

	// Empty row
	sheet.AddRow()

	// Fill patterns section
	fillSubtitleRow := sheet.AddRow()
	fillSubtitleRow.AddCell("Fill Patterns").SetStyle(styles["subtitle"])

	fillHeaderRow := sheet.AddRow()
	fillHeaderRow.AddCell("Pattern Type").SetStyle(styles["header_success"])
	fillHeaderRow.AddCell("Sample").SetStyle(styles["header_success"])
	fillHeaderRow.AddCell("Description").SetStyle(styles["header_success"])
	fillHeaderRow.AddCell("Best Use").SetStyle(styles["header_success"])

	// Fill pattern examples
	fillPatterns := []struct {
		name        string
		description string
		bestUse     string
		color       string
	}{
		{"Solid Fill", "Solid background color", "Headers, emphasis", colors["primary"]},
		{"Light Fill", "Subtle background tint", "Alternating rows", "#F8F9FA"},
		{"Accent Fill", "Branded background", "Special sections", colors["accent"]},
		{"Warning Fill", "Attention background", "Warnings, alerts", colors["warning"]},
		{"Success Fill", "Positive background", "Achievements, success", colors["success"]},
	}

	for _, fill := range fillPatterns {
		row := sheet.AddRow()
		row.AddCell(fill.name).SetStyle(styles["data_normal"])

		// Sample with fill
		fillStyle := excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Bold:   true,
				Size:   10,
				Color:  colors["white"],
				Family: "Arial",
			},
			Fill: excelbuilder.FillConfig{
				Type:  "pattern",
				Color: fill.color,
			},
			Alignment: excelbuilder.AlignmentConfig{Horizontal: "center", Vertical: "middle"},
			Border:    createBorder("thin", colors["dark"]),
		}
		row.AddCell("Sample").SetStyle(fillStyle)

		row.AddCell(fill.description).SetStyle(styles["data_normal"])
		row.AddCell(fill.bestUse).SetStyle(styles["data_normal"])
	}
}

func createDataVisualizationDemo(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, colors map[string]string) {
	// Set column widths
	sheet.SetColumnWidth("A", 20.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 15.0)
	sheet.SetColumnWidth("E", 15.0)
	sheet.SetColumnWidth("F", 20.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Data Visualization Styles").SetStyle(styles["title"])

	// Empty row
	sheet.AddRow()

	// Header
	headerRow := sheet.AddRow()
	headerRow.AddCell("Product").SetStyle(styles["header_primary"])
	headerRow.AddCell("Q1 Sales").SetStyle(styles["header_primary"])
	headerRow.AddCell("Q2 Sales").SetStyle(styles["header_primary"])
	headerRow.AddCell("Q3 Sales").SetStyle(styles["header_primary"])
	headerRow.AddCell("Q4 Sales").SetStyle(styles["header_primary"])
	headerRow.AddCell("Performance").SetStyle(styles["header_primary"])

	// Sample data for visualization
	rand.Seed(time.Now().UnixNano())
	products := []string{"Product A", "Product B", "Product C", "Product D", "Product E"}

	for i, product := range products {
		row := sheet.AddRow()

		// Product name with alternating row colors
		if i%2 == 0 {
			row.AddCell(product).SetStyle(styles["data_normal"])
		} else {
			row.AddCell(product).SetStyle(styles["data_alternate"])
		}

		// Generate sample sales data
		q1 := float64(rand.Intn(50000) + 10000)
		q2 := float64(rand.Intn(50000) + 10000)
		q3 := float64(rand.Intn(50000) + 10000)
		q4 := float64(rand.Intn(50000) + 10000)

		// Style numbers based on value ranges
		row.AddCell(q1).SetStyle(getValueStyle(q1, styles))
		row.AddCell(q2).SetStyle(getValueStyle(q2, styles))
		row.AddCell(q3).SetStyle(getValueStyle(q3, styles))
		row.AddCell(q4).SetStyle(getValueStyle(q4, styles))

		// Performance indicator
		total := q1 + q2 + q3 + q4
		performance := ""
		performanceStyle := styles["data_normal"]

		if total > 200000 {
			performance = "Excellent"
			performanceStyle = styles["high_value"]
		} else if total > 150000 {
			performance = "Good"
			performanceStyle = styles["medium_value"]
		} else {
			performance = "Needs Improvement"
			performanceStyle = styles["low_value"]
		}

		row.AddCell(performance).SetStyle(performanceStyle)
	}

	// Add summary row
	sheet.AddRow() // Empty row
	summaryRow := sheet.AddRow()
	summaryRow.AddCell("Summary Statistics").SetStyle(styles["subtitle"])

	// Add totals with special formatting
	totalsHeaderRow := sheet.AddRow()
	totalsHeaderRow.AddCell("Metric").SetStyle(styles["header_warning"])
	totalsHeaderRow.AddCell("Q1 Total").SetStyle(styles["header_warning"])
	totalsHeaderRow.AddCell("Q2 Total").SetStyle(styles["header_warning"])
	totalsHeaderRow.AddCell("Q3 Total").SetStyle(styles["header_warning"])
	totalsHeaderRow.AddCell("Q4 Total").SetStyle(styles["header_warning"])
	totalsHeaderRow.AddCell("Grand Total").SetStyle(styles["header_warning"])

	totalsRow := sheet.AddRow()
	totalsRow.AddCell("All Products").SetStyle(styles["data_normal"])
	totalsRow.AddCell(150000.0).SetStyle(styles["number_currency"])
	totalsRow.AddCell(175000.0).SetStyle(styles["number_currency"])
	totalsRow.AddCell(165000.0).SetStyle(styles["number_currency"])
	totalsRow.AddCell(180000.0).SetStyle(styles["number_currency"])
	totalsRow.AddCell(670000.0).SetStyle(styles["high_value"])
}

func getValueStyle(value float64, styles map[string]excelbuilder.StyleConfig) excelbuilder.StyleConfig {
	if value > 40000 {
		return styles["high_value"]
	} else if value > 25000 {
		return styles["medium_value"]
	}
	return styles["low_value"]
}

func createConditionalFormattingDemo(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, colors map[string]string) {
	// Set column widths
	sheet.SetColumnWidth("A", 20.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 20.0)
	sheet.SetColumnWidth("E", 20.0)

	// Title
	titleRow := sheet.AddRow()
	titleRow.AddCell("Conditional Formatting Simulation").SetStyle(styles["title"])

	// Empty row
	sheet.AddRow()

	// Description
	descRow := sheet.AddRow()
	descRow.AddCell("This demonstrates how to simulate conditional formatting using go-excelbuilder").SetStyle(excelbuilder.StyleConfig{
		Font:      excelbuilder.FontConfig{Italic: true, Size: 10, Family: "Arial", Color: colors["dark"]},
		Alignment: excelbuilder.AlignmentConfig{Horizontal: "left", Vertical: "middle"},
	})

	// Empty row
	sheet.AddRow()

	// Header
	headerRow := sheet.AddRow()
	headerRow.AddCell("Employee").SetStyle(styles["header_primary"])
	headerRow.AddCell("Score").SetStyle(styles["header_primary"])
	headerRow.AddCell("Percentage").SetStyle(styles["header_primary"])
	headerRow.AddCell("Grade").SetStyle(styles["header_primary"])
	headerRow.AddCell("Status").SetStyle(styles["header_primary"])

	// Sample employee performance data
	employees := []struct {
		name       string
		score      int
		percentage float64
	}{
		{"Alice Johnson", 95, 0.95},
		{"Bob Smith", 87, 0.87},
		{"Carol Davis", 76, 0.76},
		{"David Wilson", 92, 0.92},
		{"Eva Brown", 68, 0.68},
		{"Frank Miller", 84, 0.84},
		{"Grace Lee", 91, 0.91},
		{"Henry Taylor", 73, 0.73},
		{"Ivy Chen", 89, 0.89},
		{"Jack Anderson", 65, 0.65},
	}

	for _, emp := range employees {
		row := sheet.AddRow()
		row.AddCell(emp.name).SetStyle(styles["data_normal"])

		// Score with conditional styling
		scoreStyle := getScoreStyle(emp.score, styles)
		row.AddCell(emp.score).SetStyle(scoreStyle)

		// Percentage with conditional styling
		percentStyle := getPercentageStyle(emp.percentage, styles)
		row.AddCell(emp.percentage).SetStyle(percentStyle)

		// Grade based on score
		grade, gradeStyle := getGrade(emp.score, styles)
		row.AddCell(grade).SetStyle(gradeStyle)

		// Status based on performance
		status, statusStyle := getStatus(emp.score, styles)
		row.AddCell(status).SetStyle(statusStyle)
	}

	// Add legend
	sheet.AddRow() // Empty row
	legendRow := sheet.AddRow()
	legendRow.AddCell("Legend:").SetStyle(styles["subtitle"])

	legendItems := []struct {
		label string
		style string
	}{
		{"Excellent (90+)", "high_value"},
		{"Good (75-89)", "medium_value"},
		{"Needs Improvement (<75)", "low_value"},
	}

	for _, item := range legendItems {
		legendItemRow := sheet.AddRow()
		legendItemRow.AddCell(item.label).SetStyle(styles[item.style])
	}
}

func getScoreStyle(score int, styles map[string]excelbuilder.StyleConfig) excelbuilder.StyleConfig {
	if score >= 90 {
		return styles["high_value"]
	} else if score >= 75 {
		return styles["medium_value"]
	}
	return styles["low_value"]
}

func getPercentageStyle(percentage float64, styles map[string]excelbuilder.StyleConfig) excelbuilder.StyleConfig {
	baseStyle := styles["number_percentage"]
	if percentage >= 0.90 {
		baseStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#70AD47"}
		baseStyle.Font.Color = "#FFFFFF"
		baseStyle.Font.Bold = true
	} else if percentage >= 0.75 {
		baseStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#FFC000"}
		baseStyle.Font.Color = "#404040"
	} else {
		baseStyle.Fill = excelbuilder.FillConfig{Type: "pattern", Color: "#C5504B"}
		baseStyle.Font.Color = "#FFFFFF"
		baseStyle.Font.Bold = true
	}
	return baseStyle
}

func getGrade(score int, styles map[string]excelbuilder.StyleConfig) (string, excelbuilder.StyleConfig) {
	if score >= 90 {
		return "A", styles["high_value"]
	} else if score >= 80 {
		return "B", styles["medium_value"]
	} else if score >= 70 {
		return "C", styles["medium_value"]
	} else if score >= 60 {
		return "D", styles["low_value"]
	}
	return "F", styles["low_value"]
}

func getStatus(score int, styles map[string]excelbuilder.StyleConfig) (string, excelbuilder.StyleConfig) {
	if score >= 90 {
		return "Outstanding", styles["high_value"]
	} else if score >= 75 {
		return "Satisfactory", styles["medium_value"]
	}
	return "Needs Improvement", styles["low_value"]
}

func createProfessionalReportDemo(sheet *excelbuilder.SheetBuilder, styles map[string]excelbuilder.StyleConfig, colors map[string]string) {
	// Set column widths for professional layout
	sheet.SetColumnWidth("A", 25.0)
	sheet.SetColumnWidth("B", 15.0)
	sheet.SetColumnWidth("C", 15.0)
	sheet.SetColumnWidth("D", 15.0)
	sheet.SetColumnWidth("E", 15.0)
	sheet.SetColumnWidth("F", 20.0)

	// Report header
	reportTitleRow := sheet.AddRow()
	reportTitleRow.AddCell("QUARTERLY SALES REPORT").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Size:   16,
			Color:  colors["primary"],
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	})

	// Report metadata
	metaRow := sheet.AddRow()
	metaRow.AddCell(fmt.Sprintf("Generated: %s", time.Now().Format("January 2, 2006"))).SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Italic: true,
			Size:   10,
			Color:  colors["dark"],
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	})

	// Empty rows for spacing
	sheet.AddRow()
	sheet.AddRow()

	// Section header
	sectionRow := sheet.AddRow()
	sectionRow.AddCell("Sales Performance by Region").SetStyle(styles["subtitle"])

	// Table header
	headerRow := sheet.AddRow()
	headerRow.AddCell("Region").SetStyle(styles["header_primary"])
	headerRow.AddCell("Q1 2024").SetStyle(styles["header_primary"])
	headerRow.AddCell("Q2 2024").SetStyle(styles["header_primary"])
	headerRow.AddCell("Q3 2024").SetStyle(styles["header_primary"])
	headerRow.AddCell("Q4 2024").SetStyle(styles["header_primary"])
	headerRow.AddCell("Total").SetStyle(styles["header_primary"])

	// Sample regional data
	regionalData := []struct {
		region         string
		q1, q2, q3, q4 float64
	}{
		{"North America", 125000, 135000, 142000, 158000},
		{"Europe", 98000, 105000, 112000, 118000},
		{"Asia Pacific", 87000, 94000, 101000, 109000},
		{"Latin America", 45000, 48000, 52000, 55000},
		{"Middle East & Africa", 32000, 35000, 38000, 41000},
	}

	var grandTotal float64
	for i, data := range regionalData {
		row := sheet.AddRow()

		// Alternating row colors
		rowStyle := styles["data_normal"]
		if i%2 == 1 {
			rowStyle = styles["data_alternate"]
		}

		row.AddCell(data.region).SetStyle(rowStyle)
		row.AddCell(data.q1).SetStyle(styles["number_currency"])
		row.AddCell(data.q2).SetStyle(styles["number_currency"])
		row.AddCell(data.q3).SetStyle(styles["number_currency"])
		row.AddCell(data.q4).SetStyle(styles["number_currency"])

		total := data.q1 + data.q2 + data.q3 + data.q4
		grandTotal += total
		row.AddCell(total).SetStyle(styles["high_value"])
	}

	// Totals row
	totalsRow := sheet.AddRow()
	totalsRow.AddCell("TOTAL").SetStyle(styles["header_warning"])
	totalsRow.AddCell(387000.0).SetStyle(styles["header_warning"])
	totalsRow.AddCell(417000.0).SetStyle(styles["header_warning"])
	totalsRow.AddCell(445000.0).SetStyle(styles["header_warning"])
	totalsRow.AddCell(481000.0).SetStyle(styles["header_warning"])
	totalsRow.AddCell(grandTotal).SetStyle(styles["header_danger"])

	// Empty rows
	sheet.AddRow()
	sheet.AddRow()

	// Key insights section
	insightsRow := sheet.AddRow()
	insightsRow.AddCell("Key Insights").SetStyle(styles["subtitle"])

	insights := []string{
		"• North America remains the strongest performing region",
		"• Consistent growth across all quarters",
		"• Asia Pacific shows promising growth trajectory",
		"• Total revenue increased by 24% year-over-year",
	}

	for _, insight := range insights {
		insightRow := sheet.AddRow()
		insightRow.AddCell(insight).SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Size:   10,
				Family: "Arial",
				Color:  colors["dark"],
			},
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: "left",
				Vertical:   "middle",
			},
		})
	}

	// Footer
	sheet.AddRow()
	sheet.AddRow()
	footerRow := sheet.AddRow()
	footerRow.AddCell("Report prepared by Sales Analytics Team").SetStyle(excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Italic: true,
			Size:   9,
			Color:  colors["dark"],
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	})
}
