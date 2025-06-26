package excelbuilder_test

import (
	"fmt"
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

const (
	// Using a smaller number for benchmark runs to keep them reasonably fast.
	// 100k rows is the target, but 10k is a good proxy for benchmarks.
	benchmarkRows = 10000
	benchmarkCols = 10
)

// Benchmark_LargeFile_NoStyle measures the baseline performance of creating a large
// number of rows and cells without applying any styles.
func Benchmark_LargeFile_NoStyle(b *testing.B) {
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		builder := excelbuilder.New()
		sheet := builder.NewWorkbook().AddSheet("PerfTest")

		for r := 0; r < benchmarkRows; r++ {
			row := sheet.AddRow()
			for c := 0; c < benchmarkCols; c++ {
				row.AddCell(fmt.Sprintf("Cell %d-%d", r, c))
			}
		}
		_ = sheet.Build()
	}
}

// Benchmark_LargeFile_WithReusedStyles measures the performance of creating a large
// file where a small set of styles are reused across thousands of cells. This is
// the optimal case for the Flyweight pattern.
func Benchmark_LargeFile_WithReusedStyles(b *testing.B) {
	styles := []excelbuilder.StyleConfig{
		{Font: excelbuilder.FontConfig{Bold: true}},
		{Font: excelbuilder.FontConfig{Italic: true}},
		{Fill: excelbuilder.FillConfig{Type: "pattern", Color: "FFFF00"}},
		{Font: excelbuilder.FontConfig{Size: 14, Color: "00FF00"}},
		{Border: excelbuilder.BorderConfig{Bottom: excelbuilder.BorderSide{Style: "thin"}}},
	}

	b.ReportAllocs()
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		builder := excelbuilder.New()
		sheet := builder.NewWorkbook().AddSheet("PerfTest")

		for r := 0; r < benchmarkRows; r++ {
			row := sheet.AddRow()
			for c := 0; c < benchmarkCols; c++ {
				// Cycle through the predefined styles
				style := styles[r%len(styles)]
				row.AddCell(fmt.Sprintf("Cell %d-%d", r, c)).WithStyle(style)
			}
		}
		_ = sheet.Build()
	}
}

// Benchmark_LargeFile_WithUniqueStyles measures the performance of creating a large
// file where every cell receives a unique style. This is the worst-case scenario
// for the StyleManager, as it results in all cache misses.
func Benchmark_LargeFile_WithUniqueStyles(b *testing.B) {
	b.ReportAllocs()
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		builder := excelbuilder.New()
		sheet := builder.NewWorkbook().AddSheet("PerfTest")

		for r := 0; r < benchmarkRows; r++ {
			row := sheet.AddRow()
			for c := 0; c < benchmarkCols; c++ {
				// Create a unique style for each cell
				// Note: This is an anti-pattern in practice, used here for benchmarking.
				style := excelbuilder.StyleConfig{
					Font: excelbuilder.FontConfig{
						Size: 10 + (r % 5), // Cycle through a few sizes
					},
					Fill: excelbuilder.FillConfig{
						Type:  "pattern",
						Color: fmt.Sprintf("FF%02x%02x", r%256, c%256), // Unique color
					},
				}
				row.AddCell(fmt.Sprintf("Cell %d-%d", r, c)).WithStyle(style)
			}
		}
		_ = sheet.Build()
	}
}
