package tests

import (
	"fmt"
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

// Benchmark Case 11.1: Style Creation Performance
func BenchmarkStyleCreation_WithFlyweight(b *testing.B) {
	// Benchmark: Style creation with Flyweight pattern
	// Expected: Efficient style creation and reuse

	builder := excelbuilder.New()
	styleConfig := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:  true,
			Size:  12,
			Color: "#FF0000",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#FFFF00",
		},
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		workbook := builder.NewWorkbook().AddSheet("BenchSheet")
		row := workbook.AddRow()
		cell := row.AddCell(fmt.Sprintf("Cell_%d", i))
		cell.SetStyle(styleConfig)
		cell.Done().Done().Done().Build()
	}
}

// Benchmark Case 11.2: Style Reuse Performance
func BenchmarkStyleReuse_SameStyle(b *testing.B) {
	// Benchmark: Performance when reusing the same style
	// Expected: Better performance due to caching

	builder := excelbuilder.New()
	styleConfig := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 14,
		},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin"},
			Bottom: excelbuilder.BorderSide{Style: "thin"},
			Left:   excelbuilder.BorderSide{Style: "thin"},
			Right:  excelbuilder.BorderSide{Style: "thin"},
			Color:  "#000000",
		},
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		workbook := builder.NewWorkbook().AddSheet("ReuseSheet")
		for j := 0; j < 10; j++ {
			row := workbook.AddRow()
			for k := 0; k < 5; k++ {
				cell := row.AddCell(fmt.Sprintf("Cell_%d_%d", j, k))
				cell.SetStyle(styleConfig) // Same style reused
				cell.Done()
			}
			row.Done()
		}
		workbook.Done().Build()
	}
}

// Benchmark Case 11.3: Different Styles Performance
func BenchmarkStyleCreation_DifferentStyles(b *testing.B) {
	// Benchmark: Performance with many different styles
	// Expected: Shows overhead of creating unique styles

	builder := excelbuilder.New()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		workbook := builder.NewWorkbook().AddSheet("DiffSheet")
		row := workbook.AddRow()

		// Create unique style for each iteration
		styleConfig := excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Size:  10 + (i % 10),                    // Different sizes
				Color: fmt.Sprintf("#%06X", i%0xFFFFFF), // Different colors
			},
		}

		cell := row.AddCell(fmt.Sprintf("Cell_%d", i))
		cell.SetStyle(styleConfig)
		cell.Done().Done().Done().Build()
	}
}

// Benchmark Case 11.4: Complex Style Performance
func BenchmarkStyleCreation_ComplexStyles(b *testing.B) {
	// Benchmark: Performance with complex style configurations
	// Expected: Shows performance with full style objects

	builder := excelbuilder.New()
	complexStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:      true,
			Italic:    true,
			Underline: true,
			Size:      14,
			Color:     "#FF0000",
			Family:    "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#FFFF00",
		},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thick"},
			Bottom: excelbuilder.BorderSide{Style: "thick"},
			Left:   excelbuilder.BorderSide{Style: "thick"},
			Right:  excelbuilder.BorderSide{Style: "thick"},
			Color:  "#000000",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
			WrapText:   true,
		},
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		workbook := builder.NewWorkbook().AddSheet("ComplexSheet")
		row := workbook.AddRow()
		cell := row.AddCell(fmt.Sprintf("Complex_%d", i))
		cell.SetStyle(complexStyle)
		cell.Done().Done().Done().Build()
	}
}

// Benchmark Case 11.5: Memory Usage Comparison
func BenchmarkMemoryUsage_StyleCaching(b *testing.B) {
	// Benchmark: Memory usage with style caching
	// Expected: Lower memory usage due to flyweight pattern

	builder := excelbuilder.New()

	// Define a few styles that will be reused
	styles := []excelbuilder.StyleConfig{
		{
			Font: excelbuilder.FontConfig{Bold: true, Size: 12},
		},
		{
			Font: excelbuilder.FontConfig{Italic: true, Size: 10},
		},
		{
			Font: excelbuilder.FontConfig{Size: 14, Color: "#FF0000"},
		},
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		workbook := builder.NewWorkbook().AddSheet("MemorySheet")

		// Create multiple rows with reused styles
		for j := 0; j < 20; j++ {
			row := workbook.AddRow()
			for k := 0; k < 10; k++ {
				cell := row.AddCell(fmt.Sprintf("Cell_%d_%d", j, k))
				// Cycle through styles - this should reuse flyweights
				cell.SetStyle(styles[(j+k)%len(styles)])
				cell.Done()
			}
			row.Done()
		}
		workbook.Done().Build()
	}
}

// Benchmark Case 11.6: Large Workbook Performance
func BenchmarkLargeWorkbook_WithStyles(b *testing.B) {
	// Benchmark: Performance with large workbooks containing styled cells
	// Expected: Good performance even with many styled cells

	builder := excelbuilder.New()
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:  true,
			Size:  14,
			Color: "#FFFFFF",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#4472C4",
		},
	}

	dataStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size: 10,
		},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin"},
			Bottom: excelbuilder.BorderSide{Style: "thin"},
			Left:   excelbuilder.BorderSide{Style: "thin"},
			Right:  excelbuilder.BorderSide{Style: "thin"},
			Color:  "#D0D0D0",
		},
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		workbook := builder.NewWorkbook().AddSheet("LargeSheet")

		// Create header row
		headerRow := workbook.AddRow()
		for col := 0; col < 10; col++ {
			cell := headerRow.AddCell(fmt.Sprintf("Header_%d", col))
			cell.SetStyle(headerStyle)
			cell.Done()
		}
		headerRow.Done()

		// Create data rows
		for row := 0; row < 100; row++ {
			dataRow := workbook.AddRow()
			for col := 0; col < 10; col++ {
				cell := dataRow.AddCell(fmt.Sprintf("Data_%d_%d", row, col))
				cell.SetStyle(dataStyle)
				cell.Done()
			}
			dataRow.Done()
		}
		workbook.Done().Build()
	}
}
