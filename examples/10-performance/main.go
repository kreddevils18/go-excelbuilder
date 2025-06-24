package main

import (
	"fmt"
	"math/rand"
	"os"
	"runtime"
	"strings"
	"sync"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

type PerformanceData struct {
	ID          int
	Name        string
	Value       float64
	Category    string
	Date        time.Time
	Description string
}

type PerformanceMetrics struct {
	Approach     string
	RowCount     int
	Duration     time.Duration
	MemoryUsed   uint64
	RowsPerSec   float64
	FileSizeKB   int64
}

func main() {
	fmt.Println("Performance Optimization Example")
	fmt.Println("================================")

	// Create output directory
	if err := os.MkdirAll("output", 0755); err != nil {
		fmt.Printf("Error creating output directory: %v\n", err)
		return
	}

	// Generate test data
	rowCount := 10000
	data := generateTestData(rowCount)
	fmt.Printf("Generated %d rows of test data\n\n", len(data))

	// Test different approaches
	var metrics []PerformanceMetrics

	// 1. Basic approach (slower)
	fmt.Println("Testing Basic Approach...")
	metric1 := testBasicApproach(data)
	metrics = append(metrics, metric1)

	// 2. Optimized approach (faster)
	fmt.Println("Testing Optimized Approach...")
	metric2 := testOptimizedApproach(data)
	metrics = append(metrics, metric2)

	// 3. Bulk operations approach
	fmt.Println("Testing Bulk Operations Approach...")
	metric3 := testBulkApproach(data)
	metrics = append(metrics, metric3)

	// 4. Streaming approach
	fmt.Println("Testing Streaming Approach...")
	metric4 := testStreamingApproach(data)
	metrics = append(metrics, metric4)

	// Display performance comparison
	displayPerformanceComparison(metrics)

	fmt.Println("\nPerformance optimization examples completed!")
	fmt.Println("Check the output directory for generated files.")
}

func generateTestData(count int) []PerformanceData {
	data := make([]PerformanceData, count)
	categories := []string{"Sales", "Marketing", "Operations", "Finance", "HR"}
	names := []string{"Product A", "Product B", "Product C", "Service X", "Service Y"}

	for i := 0; i < count; i++ {
		data[i] = PerformanceData{
			ID:          i + 1,
			Name:        fmt.Sprintf("%s %d", names[rand.Intn(len(names))], i+1),
			Value:       rand.Float64() * 10000,
			Category:    categories[rand.Intn(len(categories))],
			Date:        time.Now().AddDate(0, 0, -rand.Intn(365)),
			Description: fmt.Sprintf("Description for item %d with detailed information", i+1),
		}
	}

	return data
}

func testBasicApproach(data []PerformanceData) PerformanceMetrics {
	start := time.Now()
	memStart := getMemUsage()

	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Basic Approach")

	// Add headers with individual styling
	headers := []string{"ID", "Name", "Value", "Category", "Date", "Description"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(col+"1", header).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{
				Bold:  true,
				Color: "#FFFFFF",
			},
			Fill: &excelbuilder.Fill{
				Type:  "solid",
				Color: "#4472C4",
			},
		})
	}

	// Add data with individual styling (inefficient)
	for i, row := range data {
		rowNum := i + 2
		sheet.SetCell(fmt.Sprintf("A%d", rowNum), row.ID)
		sheet.SetCell(fmt.Sprintf("B%d", rowNum), row.Name)
		sheet.SetCell(fmt.Sprintf("C%d", rowNum), row.Value).SetStyle(&excelbuilder.Style{
			NumberFormat: "#,##0.00",
		})
		sheet.SetCell(fmt.Sprintf("D%d", rowNum), row.Category)
		sheet.SetCell(fmt.Sprintf("E%d", rowNum), row.Date.Format("2006-01-02"))
		sheet.SetCell(fmt.Sprintf("F%d", rowNum), row.Description)
	}

	filename := "output/10-performance-basic.xlsx"
	if err := builder.SaveToFile(filename); err != nil {
		fmt.Printf("Error saving basic file: %v\n", err)
	}

	duration := time.Since(start)
	memEnd := getMemUsage()
	fileSize := getFileSize(filename)

	return PerformanceMetrics{
		Approach:   "Basic",
		RowCount:   len(data),
		Duration:   duration,
		MemoryUsed: memEnd - memStart,
		RowsPerSec: float64(len(data)) / duration.Seconds(),
		FileSizeKB: fileSize,
	}
}

func testOptimizedApproach(data []PerformanceData) PerformanceMetrics {
	start := time.Now()
	memStart := getMemUsage()

	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Optimized Approach")

	// Pre-define styles (reuse)
	headerStyle := &excelbuilder.Style{
		Font: &excelbuilder.Font{
			Bold:  true,
			Color: "#FFFFFF",
		},
		Fill: &excelbuilder.Fill{
			Type:  "solid",
			Color: "#4472C4",
		},
	}

	numberStyle := &excelbuilder.Style{
		NumberFormat: "#,##0.00",
	}

	// Add headers
	headers := []string{"ID", "Name", "Value", "Category", "Date", "Description"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(col+"1", header).SetStyle(headerStyle)
	}

	// Batch process data in chunks
	chunkSize := 1000
	for i := 0; i < len(data); i += chunkSize {
		end := i + chunkSize
		if end > len(data) {
			end = len(data)
		}

		// Process chunk
		for j := i; j < end; j++ {
			row := data[j]
			rowNum := j + 2

			sheet.SetCell(fmt.Sprintf("A%d", rowNum), row.ID)
			sheet.SetCell(fmt.Sprintf("B%d", rowNum), row.Name)
			sheet.SetCell(fmt.Sprintf("C%d", rowNum), row.Value).SetStyle(numberStyle)
			sheet.SetCell(fmt.Sprintf("D%d", rowNum), row.Category)
			sheet.SetCell(fmt.Sprintf("E%d", rowNum), row.Date.Format("2006-01-02"))
			sheet.SetCell(fmt.Sprintf("F%d", rowNum), row.Description)
		}

		// Force garbage collection periodically
		if i%5000 == 0 {
			runtime.GC()
		}
	}

	filename := "output/10-performance-optimized.xlsx"
	if err := builder.SaveToFile(filename); err != nil {
		fmt.Printf("Error saving optimized file: %v\n", err)
	}

	duration := time.Since(start)
	memEnd := getMemUsage()
	fileSize := getFileSize(filename)

	return PerformanceMetrics{
		Approach:   "Optimized",
		RowCount:   len(data),
		Duration:   duration,
		MemoryUsed: memEnd - memStart,
		RowsPerSec: float64(len(data)) / duration.Seconds(),
		FileSizeKB: fileSize,
	}
}

func testBulkApproach(data []PerformanceData) PerformanceMetrics {
	start := time.Now()
	memStart := getMemUsage()

	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Bulk Operations")

	// Prepare bulk data
	bulkData := make([][]interface{}, len(data)+1)

	// Headers
	bulkData[0] = []interface{}{"ID", "Name", "Value", "Category", "Date", "Description"}

	// Data rows
	for i, row := range data {
		bulkData[i+1] = []interface{}{
			row.ID,
			row.Name,
			row.Value,
			row.Category,
			row.Date.Format("2006-01-02"),
			row.Description,
		}
	}

	// Bulk insert (if supported by the library)
	for i, rowData := range bulkData {
		for j, cellData := range rowData {
			col := string(rune('A' + j))
			row := i + 1
			cell := sheet.SetCell(fmt.Sprintf("%s%d", col, row), cellData)

			// Apply styles efficiently
			if i == 0 { // Header row
				cell.SetStyle(&excelbuilder.Style{
					Font: &excelbuilder.Font{Bold: true, Color: "#FFFFFF"},
					Fill: &excelbuilder.Fill{Type: "solid", Color: "#4472C4"},
				})
			} else if j == 2 { // Value column
				cell.SetStyle(&excelbuilder.Style{NumberFormat: "#,##0.00"})
			}
		}
	}

	filename := "output/10-performance-bulk.xlsx"
	if err := builder.SaveToFile(filename); err != nil {
		fmt.Printf("Error saving bulk file: %v\n", err)
	}

	duration := time.Since(start)
	memEnd := getMemUsage()
	fileSize := getFileSize(filename)

	return PerformanceMetrics{
		Approach:   "Bulk Operations",
		RowCount:   len(data),
		Duration:   duration,
		MemoryUsed: memEnd - memStart,
		RowsPerSec: float64(len(data)) / duration.Seconds(),
		FileSizeKB: fileSize,
	}
}

func testStreamingApproach(data []PerformanceData) PerformanceMetrics {
	start := time.Now()
	memStart := getMemUsage()

	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Streaming")

	// Process data in streaming fashion with goroutines
	const numWorkers = 4
	const batchSize = 500

	// Channel for work distribution
	workChan := make(chan []PerformanceData, numWorkers)
	doneChan := make(chan bool, numWorkers)

	// Start workers
	var wg sync.WaitGroup
	for i := 0; i < numWorkers; i++ {
		wg.Add(1)
		go func(workerID int) {
			defer wg.Done()
			for batch := range workChan {
				processBatch(sheet, batch, workerID)
			}
		}(i)
	}

	// Add headers first
	headers := []string{"ID", "Name", "Value", "Category", "Date", "Description"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(col+"1", header).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{Bold: true, Color: "#FFFFFF"},
			Fill: &excelbuilder.Fill{Type: "solid", Color: "#4472C4"},
		})
	}

	// Send work in batches
	go func() {
		defer close(workChan)
		for i := 0; i < len(data); i += batchSize {
			end := i + batchSize
			if end > len(data) {
				end = len(data)
			}
			batch := data[i:end]
			workChan <- batch
		}
	}()

	// Wait for all workers to complete
	wg.Wait()

	filename := "output/10-performance-streaming.xlsx"
	if err := builder.SaveToFile(filename); err != nil {
		fmt.Printf("Error saving streaming file: %v\n", err)
	}

	duration := time.Since(start)
	memEnd := getMemUsage()
	fileSize := getFileSize(filename)

	return PerformanceMetrics{
		Approach:   "Streaming",
		RowCount:   len(data),
		Duration:   duration,
		MemoryUsed: memEnd - memStart,
		RowsPerSec: float64(len(data)) / duration.Seconds(),
		FileSizeKB: fileSize,
	}
}

func processBatch(sheet *excelbuilder.Sheet, batch []PerformanceData, workerID int) {
	numberStyle := &excelbuilder.Style{NumberFormat: "#,##0.00"}

	for _, row := range batch {
		rowNum := row.ID + 1 // +1 for header

		sheet.SetCell(fmt.Sprintf("A%d", rowNum), row.ID)
		sheet.SetCell(fmt.Sprintf("B%d", rowNum), row.Name)
		sheet.SetCell(fmt.Sprintf("C%d", rowNum), row.Value).SetStyle(numberStyle)
		sheet.SetCell(fmt.Sprintf("D%d", rowNum), row.Category)
		sheet.SetCell(fmt.Sprintf("E%d", rowNum), row.Date.Format("2006-01-02"))
		sheet.SetCell(fmt.Sprintf("F%d", rowNum), row.Description)
	}
}

func getMemUsage() uint64 {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	return m.Alloc
}

func getFileSize(filename string) int64 {
	if info, err := os.Stat(filename); err == nil {
		return info.Size() / 1024 // Convert to KB
	}
	return 0
}

func displayPerformanceComparison(metrics []PerformanceMetrics) {
	fmt.Println("\n" + strings.Repeat("=", 80))
	fmt.Println("PERFORMANCE COMPARISON RESULTS")
	fmt.Println(strings.Repeat("=", 80))
	fmt.Printf("%-20s %-10s %-12s %-12s %-12s %-10s\n",
		"Approach", "Rows", "Duration", "Memory(KB)", "Rows/Sec", "Size(KB)")
	fmt.Println(strings.Repeat("-", 80))

	for _, m := range metrics {
		memoryKB := m.MemoryUsed / 1024
		fmt.Printf("%-20s %-10d %-12s %-12d %-12.0f %-10d\n",
			m.Approach,
			m.RowCount,
			m.Duration.Round(time.Millisecond),
			memoryKB,
			m.RowsPerSec,
			m.FileSizeKB)
	}

	fmt.Println(strings.Repeat("-", 80))

	// Find best performing approach
	bestSpeed := metrics[0]
	bestMemory := metrics[0]

	for _, m := range metrics[1:] {
		if m.RowsPerSec > bestSpeed.RowsPerSec {
			bestSpeed = m
		}
		if m.MemoryUsed < bestMemory.MemoryUsed {
			bestMemory = m
		}
	}

	fmt.Printf("\nBest Speed: %s (%.0f rows/sec)\n", bestSpeed.Approach, bestSpeed.RowsPerSec)
	fmt.Printf("Best Memory: %s (%d KB)\n", bestMemory.Approach, bestMemory.MemoryUsed/1024)
	fmt.Println(strings.Repeat("=", 80))
}