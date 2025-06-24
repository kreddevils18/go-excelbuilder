package main

import (
	"fmt"
	"log"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

func main() {
	fmt.Println("Creating Advanced Layout Example...")

	// Create a new workbook
	wb := excelbuilder.NewWorkbook()

	// Create different sheets to demonstrate advanced layout features
	createComplexMergeSheet(wb)
	createDashboardLayoutSheet(wb)
	createMatrixReportSheet(wb)
	createHierarchicalDataSheet(wb)
	createResponsiveLayoutSheet(wb)

	// Save the workbook
	filename := "09-advanced-layout.xlsx"
	if err := wb.SaveAs(filename); err != nil {
		log.Fatalf("Failed to save workbook: %v", err)
	}

	fmt.Printf("Advanced layout example saved as %s\n", filename)
}

// createComplexMergeSheet demonstrates advanced cell merging patterns
func createComplexMergeSheet(wb *excelbuilder.Workbook) {
	sheet := wb.NewSheet("Complex Merge")

	// Main title spanning entire width
	sheet.SetCell("A1", "Advanced Cell Merging Patterns").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{
			Size:   18,
			Bold:   true,
			Family: "Arial",
			Color:  "#FFFFFF",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#2E75B6",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	sheet.MergeRange("A1:H1")
	sheet.SetRowHeight(1, 40)

	// Section 1: L-shaped merge
	sheet.SetCell("A3", "L-Shaped Merge Pattern").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{
			Bold:   true,
			Size:   14,
			Color:  "#2E75B6",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#F2F2F2",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: &excelbuilder.Border{
			Top:    &excelbuilder.BorderStyle{Style: "thick", Color: "#2E75B6"},
			Bottom: &excelbuilder.BorderStyle{Style: "thick", Color: "#2E75B6"},
			Left:   &excelbuilder.BorderStyle{Style: "thick", Color: "#2E75B6"},
			Right:  &excelbuilder.BorderStyle{Style: "thick", Color: "#2E75B6"},
		},
	})
	sheet.MergeRange("A3:D5")

	// Data cells for L-shape
	sheet.SetCell("E3", "Data 1").SetStyle(getDataCellStyle())
	sheet.SetCell("F3", "Data 2").SetStyle(getDataCellStyle())
	sheet.SetCell("E4", "Data 3").SetStyle(getDataCellStyle())
	sheet.SetCell("F4", "Data 4").SetStyle(getDataCellStyle())
	sheet.SetCell("E5", "Data 5").SetStyle(getDataCellStyle())
	sheet.SetCell("F5", "Data 6").SetStyle(getDataCellStyle())

	// Section 2: Cross merge pattern
	sheet.SetCell("A7", "Cross Merge Pattern").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{
			Bold:   true,
			Size:   12,
			Color:  "#70AD47",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#E2EFDA",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	sheet.MergeRange("A7:H7")

	// Vertical header
	sheet.SetCell("A8", "Vertical Header").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{
			Bold:   true,
			Color:  "#FFFFFF",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#70AD47",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
			Vertical:   "center",
			TextRotation: 90,
		},
	})
	sheet.MergeRange("A8:A12")

	// Horizontal headers
	horizontalHeaders := []string{"Q1", "Q2", "Q3", "Q4"}
	for i, header := range horizontalHeaders {
		col := string(rune('B' + i))
		sheet.SetCell(col+"8", header).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{
				Bold:   true,
				Color:  "#FFFFFF",
			},
			Fill: &excelbuilder.Fill{
				Type:  "solid",
				Color: "#70AD47",
			},
			Alignment: &excelbuilder.Alignment{
				Horizontal: "center",
				Vertical:   "center",
			},
		})
	}

	// Data grid
	data := [][]interface{}{
		{100, 120, 110, 130},
		{90, 95, 105, 115},
		{80, 85, 90, 95},
		{70, 75, 80, 85},
	}

	for i, row := range data {
		for j, value := range row {
			col := string(rune('B' + j))
			rowNum := fmt.Sprintf("%d", i+9)
			sheet.SetCell(col+rowNum, value).SetStyle(getDataCellStyle())
		}
	}

	// Section 3: Nested merge pattern
	sheet.SetCell("A14", "Nested Merge Structure").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{
			Bold:   true,
			Size:   12,
			Color:  "#C5504B",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#FCE4D6",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
		},
	})
	sheet.MergeRange("A14:H14")

	// Outer container
	sheet.SetCell("A15", "Main Category").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{
			Bold:   true,
			Color:  "#FFFFFF",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#C5504B",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	sheet.MergeRange("A15:H17")

	// Sub-categories
	subCategories := []string{"Sub A", "Sub B", "Sub C"}
	for i, subCat := range subCategories {
		col := string(rune('A' + i*2))
		nextCol := string(rune('A' + i*2 + 1))
		sheet.SetCell(col+"18", subCat).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{
				Bold:   true,
				Color:  "#C5504B",
			},
			Fill: &excelbuilder.Fill{
				Type:  "solid",
				Color: "#F2F2F2",
			},
			Alignment: &excelbuilder.Alignment{
				Horizontal: "center",
			},
		})
		sheet.MergeRange(col+"18:"+nextCol+"18")
	}

	// Set column widths
	for i := 0; i < 8; i++ {
		col := string(rune('A' + i))
		sheet.SetColumnWidth(col, 12)
	}
}

// createDashboardLayoutSheet creates a multi-section dashboard layout
func createDashboardLayoutSheet(wb *excelbuilder.Workbook) {
	sheet := wb.NewSheet("Dashboard Layout")

	// Dashboard title
	sheet.SetCell("A1", "Executive Dashboard").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{
			Size:   20,
			Bold:   true,
			Family: "Arial",
			Color:  "#FFFFFF",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#1F4E79",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	sheet.MergeRange("A1:L1")
	sheet.SetRowHeight(1, 50)

	// Date and time info
	sheet.SetCell("A2", fmt.Sprintf("Generated: %s", time.Now().Format("January 2, 2006 15:04"))).SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{
			Italic: true,
			Color:  "#666666",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "right",
		},
	})
	sheet.MergeRange("A2:L2")

	// KPI Section (Top Row)
	kpis := []struct {
		title string
		value string
		change string
		color string
	}{
		{"Revenue", "$2.4M", "+12%", "#70AD47"},
		{"Orders", "1,234", "+8%", "#70AD47"},
		{"Customers", "856", "+15%", "#70AD47"},
		{"Conversion", "3.2%", "-2%", "#C5504B"},
	}

	for i, kpi := range kpis {
		startCol := string(rune('A' + i*3))
		endCol := string(rune('A' + i*3 + 2))
		
		// KPI Title
		sheet.SetCell(startCol+"4", kpi.title).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{
				Bold:   true,
				Size:   12,
				Color:  "#333333",
			},
			Fill: &excelbuilder.Fill{
				Type:  "solid",
				Color: "#F8F9FA",
			},
			Alignment: &excelbuilder.Alignment{
				Horizontal: "center",
			},
			Border: &excelbuilder.Border{
				Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
				Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
				Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
			},
		})
		sheet.MergeRange(startCol+"4:"+endCol+"4")
		
		// KPI Value
		sheet.SetCell(startCol+"5", kpi.value).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{
				Bold:   true,
				Size:   18,
				Color:  "#1F4E79",
			},
			Alignment: &excelbuilder.Alignment{
				Horizontal: "center",
			},
			Border: &excelbuilder.Border{
				Left:  &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
				Right: &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
			},
		})
		sheet.MergeRange(startCol+"5:"+endCol+"5")
		
		// KPI Change
		sheet.SetCell(startCol+"6", kpi.change).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{
				Bold:   true,
				Color:  kpi.color,
			},
			Alignment: &excelbuilder.Alignment{
				Horizontal: "center",
			},
			Border: &excelbuilder.Border{
				Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
				Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
				Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
			},
		})
		sheet.MergeRange(startCol+"6:"+endCol+"6")
	}

	// Left Panel - Sales by Region
	sheet.SetCell("A8", "Sales by Region").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A8:F8")

	regionData := [][]interface{}{
		{"Region", "Q1", "Q2", "Q3", "Q4", "Total"},
		{"North", 250000, 280000, 320000, 350000, 1200000},
		{"South", 180000, 200000, 220000, 240000, 840000},
		{"East", 300000, 320000, 340000, 380000, 1340000},
		{"West", 220000, 240000, 260000, 280000, 1000000},
	}

	for i, row := range regionData {
		for j, value := range row {
			col := string(rune('A' + j))
			rowNum := fmt.Sprintf("%d", i+9)
			
			var style *excelbuilder.Style
			if i == 0 {
				style = getTableHeaderStyle()
			} else {
				style = getDataCellStyle()
			}
			
			sheet.SetCell(col+rowNum, value).SetStyle(style)
		}
	}

	// Right Panel - Top Products
	sheet.SetCell("H8", "Top Products").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("H8:L8")

	productData := [][]interface{}{
		{"Product", "Units", "Revenue", "Margin", "Trend"},
		{"Product A", 1200, 240000, "20%", "↑"},
		{"Product B", 980, 196000, "18%", "↑"},
		{"Product C", 750, 150000, "22%", "→"},
		{"Product D", 650, 130000, "15%", "↓"},
		{"Product E", 500, 100000, "25%", "↑"},
	}

	for i, row := range productData {
		for j, value := range row {
			col := string(rune('H' + j))
			rowNum := fmt.Sprintf("%d", i+9)
			
			var style *excelbuilder.Style
			if i == 0 {
				style = getTableHeaderStyle()
			} else {
				style = getDataCellStyle()
			}
			
			sheet.SetCell(col+rowNum, value).SetStyle(style)
		}
	}

	// Set column widths
	colWidths := []int{12, 10, 10, 10, 10, 12, 2, 12, 8, 12, 8, 8}
	for i, width := range colWidths {
		col := string(rune('A' + i))
		sheet.SetColumnWidth(col, width)
	}
}

// createMatrixReportSheet creates a cross-tabulation matrix layout
func createMatrixReportSheet(wb *excelbuilder.Workbook) {
	sheet := wb.NewSheet("Matrix Report")

	// Title
	sheet.SetCell("A1", "Sales Performance Matrix").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{
			Size:   16,
			Bold:   true,
			Family: "Arial",
			Color:  "#FFFFFF",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#2E75B6",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	sheet.MergeRange("A1:H1")
	sheet.SetRowHeight(1, 40)

	// Matrix structure with merged headers
	sheet.SetCell("A3", "Product\\Region").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{
			Bold:   true,
			Color:  "#FFFFFF",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#4472C4",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: &excelbuilder.Border{
			Top:    &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
			Bottom: &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
			Left:   &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
			Right:  &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
		},
	})

	// Region headers
	regions := []string{"North", "South", "East", "West", "Total"}
	for i, region := range regions {
		col := string(rune('B' + i))
		sheet.SetCell(col+"3", region).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{
				Bold:   true,
				Color:  "#FFFFFF",
			},
			Fill: &excelbuilder.Fill{
				Type:  "solid",
				Color: "#4472C4",
			},
			Alignment: &excelbuilder.Alignment{
				Horizontal: "center",
			},
			Border: &excelbuilder.Border{
				Top:    &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
				Bottom: &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
				Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
				Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
			},
		})
	}

	// Product data matrix
	products := []string{"Product A", "Product B", "Product C", "Product D", "Total"}
	matrixData := [][]int{
		{120, 95, 140, 85, 440},
		{80, 110, 90, 120, 400},
		{100, 75, 130, 95, 400},
		{90, 85, 110, 105, 390},
		{390, 365, 470, 405, 1630},
	}

	for i, product := range products {
		row := fmt.Sprintf("%d", i+4)
		
		// Product name
		var productStyle *excelbuilder.Style
		if i == len(products)-1 {
			// Total row
			productStyle = &excelbuilder.Style{
				Font: &excelbuilder.Font{
					Bold:   true,
					Color:  "#FFFFFF",
				},
				Fill: &excelbuilder.Fill{
					Type:  "solid",
					Color: "#4472C4",
				},
				Alignment: &excelbuilder.Alignment{
					Horizontal: "center",
				},
				Border: &excelbuilder.Border{
					Top:    &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
					Bottom: &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
					Left:   &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
					Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
				},
			}
		} else {
			productStyle = &excelbuilder.Style{
				Font: &excelbuilder.Font{
					Bold:   true,
					Color:  "#FFFFFF",
				},
				Fill: &excelbuilder.Fill{
					Type:  "solid",
					Color: "#70AD47",
				},
				Alignment: &excelbuilder.Alignment{
					Horizontal: "center",
				},
				Border: &excelbuilder.Border{
					Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
					Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
					Left:   &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
					Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
				},
			}
		}
		
		sheet.SetCell("A"+row, product).SetStyle(productStyle)
		
		// Data values
		for j, value := range matrixData[i] {
			col := string(rune('B' + j))
			
			var cellStyle *excelbuilder.Style
			if i == len(products)-1 || j == len(matrixData[i])-1 {
				// Total row or column
				cellStyle = &excelbuilder.Style{
					Font: &excelbuilder.Font{
						Bold:   true,
						Color:  "#FFFFFF",
					},
					Fill: &excelbuilder.Fill{
						Type:  "solid",
						Color: "#4472C4",
					},
					Alignment: &excelbuilder.Alignment{
						Horizontal: "center",
					},
					Border: &excelbuilder.Border{
						Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
						Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
						Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
						Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
					},
				}
			} else {
				style = getDataCellStyle()
			}
			
			sheet.SetCell(col+fmt.Sprintf("%d", row), value).SetStyle(style)
			sheet.MergeRange(col + fmt.Sprintf("%d", row) + ":" + nextCol + fmt.Sprintf("%d", row))
		}
	}

	// Set column widths for responsive layout
	for i := 0; i < 12; i++ {
		col := string(rune('A' + i))
		sheet.SetColumnWidth(col, 8)
	}
}

// Helper functions for consistent styling
func getDataCellStyle() *excelbuilder.Style {
	return &excelbuilder.Style{
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: &excelbuilder.Border{
			Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
			Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
			Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
			Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
		},
	}
}

func getSectionHeaderStyle() *excelbuilder.Style {
	return &excelbuilder.Style{
		Font: &excelbuilder.Font{
			Bold:   true,
			Size:   14,
			Color:  "#FFFFFF",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#4472C4",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	}
}

func getTableHeaderStyle() *excelbuilder.Style {
	return &excelbuilder.Style{
		Font: &excelbuilder.Font{
			Bold:   true,
			Color:  "#FFFFFF",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#2E75B6",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: &excelbuilder.Border{
			Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
			Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
			Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
			Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
		},
	}
}
						Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
						Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
						Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
					},
				}
			} else {
				// Regular data cell with conditional formatting
				var bgColor string
				if value >= 120 {
					bgColor = "#C6EFCE" // Light green
				} else if value >= 100 {
					bgColor = "#FFEB9C" // Light yellow
				} else {
					bgColor = "#FFC7CE" // Light red
				}
				
				cellStyle = &excelbuilder.Style{
					Fill: &excelbuilder.Fill{
						Type:  "solid",
						Color: bgColor,
					},
					Alignment: &excelbuilder.Alignment{
						Horizontal: "center",
					},
					Border: &excelbuilder.Border{
						Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
						Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
						Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
						Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
					},
				}
			}
			
			sheet.SetCell(col+row, value).SetStyle(cellStyle)
		}
	}

	// Legend
	sheet.SetCell("A10", "Performance Legend:").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{
			Bold: true,
		},
	})
	
	legendItems := []struct {
		text  string
		color string
	}{
		{"Excellent (≥120)", "#C6EFCE"},
		{"Good (100-119)", "#FFEB9C"},
		{"Needs Improvement (<100)", "#FFC7CE"},
	}
	
	for i, item := range legendItems {
		row := fmt.Sprintf("%d", 11+i)
		sheet.SetCell("A"+row, item.text).SetStyle(&excelbuilder.Style{
			Fill: &excelbuilder.Fill{
				Type:  "solid",
				Color: item.color,
			},
			Border: &excelbuilder.Border{
				Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
				Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
				Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
				Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
			},
		})
	}

	// Set column widths
	for i := 0; i < 7; i++ {
		col := string(rune('A' + i))
		sheet.SetColumnWidth(col, 15)
	}
}

// createHierarchicalDataSheet demonstrates tree-like data structures
func createHierarchicalDataSheet(wb *excelbuilder.Workbook) {
	sheet := wb.NewSheet("Hierarchical Data")

	// Title
	sheet.SetCell("A1", "Organizational Hierarchy").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{
			Size:   16,
			Bold:   true,
			Family: "Arial",
			Color:  "#FFFFFF",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#2E75B6",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	sheet.MergeRange("A1:F1")
	sheet.SetRowHeight(1, 40)

	// Hierarchical structure
	hierarchy := []struct {
		level       int
		name        string
		position    string
		employees   int
		budget      string
		description string
	}{
		{0, "CEO", "Chief Executive Officer", 250, "$50M", "Overall company leadership"},
		{1, "CTO", "Chief Technology Officer", 80, "$15M", "Technology strategy and development"},
		{2, "VP Engineering", "Vice President of Engineering", 45, "$8M", "Engineering teams management"},
		{3, "Backend Team", "Backend Development Team", 15, "$2.5M", "Server-side development"},
		{3, "Frontend Team", "Frontend Development Team", 12, "$2M", "User interface development"},
		{3, "DevOps Team", "DevOps and Infrastructure", 8, "$1.5M", "Infrastructure and deployment"},
		{2, "VP Product", "Vice President of Product", 25, "$4M", "Product strategy and management"},
		{3, "Product Managers", "Product Management Team", 8, "$1.2M", "Product planning and execution"},
		{3, "UX/UI Team", "User Experience Design", 10, "$1.5M", "User experience and design"},
		{3, "QA Team", "Quality Assurance", 7, "$1M", "Testing and quality control"},
		{1, "CFO", "Chief Financial Officer", 35, "$8M", "Financial planning and analysis"},
		{2, "Accounting", "Accounting Department", 15, "$3M", "Financial reporting and compliance"},
		{2, "Finance", "Finance Department", 12, "$2.5M", "Financial planning and budgeting"},
		{1, "CMO", "Chief Marketing Officer", 40, "$12M", "Marketing and customer acquisition"},
		{2, "Digital Marketing", "Digital Marketing Team", 20, "$6M", "Online marketing and campaigns"},
		{2, "Sales", "Sales Department", 15, "$4M", "Customer acquisition and retention"},
	}

	// Headers
	headers := []string{"Position", "Title", "Employees", "Budget", "Description"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(col+"3", header).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{
				Bold:   true,
				Color:  "#FFFFFF",
			},
			Fill: &excelbuilder.Fill{
				Type:  "solid",
				Color: "#4472C4",
			},
			Alignment: &excelbuilder.Alignment{
				Horizontal: "center",
			},
			Border: &excelbuilder.Border{
				Top:    &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
				Bottom: &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
				Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
				Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
			},
		})
	}

	// Hierarchical data with indentation and styling
	for i, item := range hierarchy {
		row := fmt.Sprintf("%d", i+4)
		
		// Determine styling based on hierarchy level
		var bgColor, textColor string
		var fontSize int
		var bold bool
		
		switch item.level {
		case 0: // CEO level
			bgColor = "#1F4E79"
			textColor = "#FFFFFF"
			fontSize = 14
			bold = true
		case 1: // C-level
			bgColor = "#2E75B6"
			textColor = "#FFFFFF"
			fontSize = 12
			bold = true
		case 2: // VP level
			bgColor = "#5B9BD5"
			textColor = "#FFFFFF"
			fontSize = 11
			bold = true
		case 3: // Team level
			bgColor = "#BDD7EE"
			textColor = "#000000"
			fontSize = 10
			bold = false
		}
		
		// Create indented name with hierarchy indicators
		indent := ""
		for j := 0; j < item.level; j++ {
			indent += "  "
		}
		
		hierarchyIndicator := ""
		if item.level > 0 {
			hierarchyIndicator = "└─ "
		}
		
		indentedName := indent + hierarchyIndicator + item.name
		
		// Position name with indentation
		sheet.SetCell("A"+row, indentedName).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{
				Bold:   bold,
				Size:   fontSize,
				Color:  textColor,
				Family: "Courier New", // Monospace for better alignment
			},
			Fill: &excelbuilder.Fill{
				Type:  "solid",
				Color: bgColor,
			},
			Border: &excelbuilder.Border{
				Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
				Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
				Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
				Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
			},
		})
		
		// Other columns
		values := []interface{}{item.position, item.employees, item.budget, item.description}
		for j, value := range values {
			col := string(rune('B' + j))
			sheet.SetCell(col+row, value).SetStyle(&excelbuilder.Style{
				Font: &excelbuilder.Font{
					Bold:   bold,
					Size:   fontSize,
					Color:  textColor,
				},
				Fill: &excelbuilder.Fill{
					Type:  "solid",
					Color: bgColor,
				},
				Border: &excelbuilder.Border{
					Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
					Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
					Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
					Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
				},
			})
		}
	}

	// Set column widths
	colWidths := []int{25, 25, 12, 12, 30}
	for i, width := range colWidths {
		col := string(rune('A' + i))
		sheet.SetColumnWidth(col, width)
	}
}

// createResponsiveLayoutSheet demonstrates adaptive layout patterns
func createResponsiveLayoutSheet(wb *excelbuilder.Workbook) {
	sheet := wb.NewSheet("Responsive Layout")

	// Title
	sheet.SetCell("A1", "Responsive Layout Patterns").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{
			Size:   16,
			Bold:   true,
			Family: "Arial",
			Color:  "#FFFFFF",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#2E75B6",
		},
		Alignment: &excelbuilder.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	sheet.MergeRange("A1:L1")
	sheet.SetRowHeight(1, 40)

	// Flexible grid system demonstration
	sheet.SetCell("A3", "Flexible Grid System").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A3:L3")

	// 12-column grid demonstration
	gridSections := []struct {
		name     string
		colSpan  int
		startCol int
		color    string
	}{
		{"Header (12 cols)", 12, 0, "#4472C4"},
		{"Sidebar (3 cols)", 3, 0, "#70AD47"},
		{"Main Content (6 cols)", 6, 3, "#FFC000"},
		{"Aside (3 cols)", 3, 9, "#C5504B"},
		{"Footer (12 cols)", 12, 0, "#7030A0"},
	}

	row := 5
	for _, section := range gridSections {
		startCol := string(rune('A' + section.startCol))
		endCol := string(rune('A' + section.startCol + section.colSpan - 1))
		
		sheet.SetCell(startCol+fmt.Sprintf("%d", row), section.name).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{
				Bold:   true,
				Color:  "#FFFFFF",
			},
			Fill: &excelbuilder.Fill{
				Type:  "solid",
				Color: section.color,
			},
			Alignment: &excelbuilder.Alignment{
				Horizontal: "center",
				Vertical:   "center",
			},
			Border: &excelbuilder.Border{
				Top:    &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
				Bottom: &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
				Left:   &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
				Right:  &excelbuilder.BorderStyle{Style: "thick", Color: "#000000"},
			},
		})
		
		if section.colSpan > 1 {
			sheet.MergeRange(startCol + fmt.Sprintf("%d", row) + ":" + endCol + fmt.Sprintf("%d", row))
		}
		
		sheet.SetRowHeight(row, 30)
		
		// Move to next row for different sections
		if section.name == "Header (12 cols)" || section.name == "Aside (3 cols)" {
			row++
		}
	}

	// Breakpoint demonstration
	sheet.SetCell("A9", "Responsive Breakpoints").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A9:L9")

	breakpoints := []struct {
		size        string
		description string
		layout      string
		color       string
	}{
		{"Mobile (1-4 cols)", "Small screens, stacked layout", "Single column", "#FF6B6B"},
		{"Tablet (5-8 cols)", "Medium screens, 2-column layout", "Two columns", "#4ECDC4"},
		{"Desktop (9-12 cols)", "Large screens, multi-column", "Three+ columns", "#45B7D1"},
	}

	for i, bp := range breakpoints {
		row := 11 + i
		
		// Size indicator
		sheet.SetCell("A"+fmt.Sprintf("%d", row), bp.size).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{
				Bold:   true,
				Color:  "#FFFFFF",
			},
			Fill: &excelbuilder.Fill{
				Type:  "solid",
				Color: bp.color,
			},
			Alignment: &excelbuilder.Alignment{
				Horizontal: "center",
			},
		})
		sheet.MergeRange("A" + fmt.Sprintf("%d", row) + ":C" + fmt.Sprintf("%d", row))
		
		// Description
		sheet.SetCell("D"+fmt.Sprintf("%d", row), bp.description).SetStyle(getDataCellStyle())
		sheet.MergeRange("D" + fmt.Sprintf("%d", row) + ":H" + fmt.Sprintf("%d", row))
		
		// Layout
		sheet.SetCell("I"+fmt.Sprintf("%d", row), bp.layout).SetStyle(getDataCellStyle())
		sheet.MergeRange("I" + fmt.Sprintf("%d", row) + ":L" + fmt.Sprintf("%d", row))
	}

	// Adaptive content demonstration
	sheet.SetCell("A15", "Adaptive Content Areas").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A15:L15")

	// Content priority demonstration
	contentAreas := []struct {
		priority string
		content  string
		mobile   string
		tablet   string
		desktop  string
	}{
		{"High", "Primary Navigation", "Hamburger Menu", "Horizontal Menu", "Full Navigation Bar"},
		{"High", "Main Content", "Full Width", "2/3 Width", "60% Width"},
		{"Medium", "Secondary Nav", "Hidden", "Sidebar", "Top Bar"},
		{"Low", "Advertisements", "Hidden", "Bottom", "Sidebar"},
		{"Low", "Related Links", "Hidden", "Hidden", "Footer"},
	}

	// Headers for content table
	contentHeaders := []string{"Priority", "Content", "Mobile", "Tablet", "Desktop"}
	for i, header := range contentHeaders {
		col := string(rune('A' + i*2))
		nextCol := string(rune('A' + i*2 + 1))
		sheet.SetCell(col+"17", header).SetStyle(getTableHeaderStyle())
		sheet.MergeRange(col + "17:" + nextCol + "17")
	}

	for i, area := range contentAreas {
		row := 18 + i
		
		values := []string{area.priority, area.content, area.mobile, area.tablet, area.desktop}
		for j, value := range values {
			col := string(rune('A' + j*2))
			nextCol := string(rune('A' + j*2 + 1))
			
			var style *excelbuilder.Style
			if j == 0 {
				// Priority column with color coding
				var bgColor string
				switch area.priority {
				case "High":
					bgColor = "#C6EFCE"
				case "Medium":
					bgColor = "#FFEB9C"
				case "Low":
					bgColor = "#FFC7CE"
				}
				
				style = &excelbuilder.Style{
					Font: &excelbuilder.Font{
						Bold: true,
					},
					Fill: &excelbuilder.Fill{
						Type:  "solid",
						Color: bgColor,
					},
					Alignment: &excelbuilder.Alignment{
						Horizontal: "center",
					},
					Border: &excelbuilder.Border{
						Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},