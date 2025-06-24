package main

import (
	"fmt"
	"math/rand"
	"os"
	"runtime"
	"strings"
	"sync"
	"sync/atomic"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

type ReportData struct {
	ID       int
	Name     string
	Value    float64
	Category string
	Date     time.Time
}

type ReportJob struct {
	ID       int
	Name     string
	Data     []ReportData
	Filename string
}

type ProcessingResult struct {
	JobID    int
	Filename string
	Duration time.Duration
	RowCount int
	Error    error
}

type ConcurrencyMetrics struct {
	SequentialTime time.Duration
	ConcurrentTime time.Duration
	Speedup        float64
	Goroutines     int
	MemoryUsed     uint64
}

func main() {
	fmt.Println("Concurrent Processing Example")
	fmt.Println("=============================")

	// Create output directory
	if err := os.MkdirAll("output", 0755); err != nil {
		fmt.Printf("Error creating output directory: %v\n", err)
		return
	}

	// Generate test data
	jobs := generateReportJobs(5, 2000)
	fmt.Printf("Generated %d report jobs\n\n", len(jobs))

	// Test sequential processing
	fmt.Println("Testing Sequential Processing...")
	seqStart := time.Now()
	sequentialResults := processReportsSequentially(jobs)
	sequentialTime := time.Since(seqStart)

	// Test concurrent processing
	fmt.Println("Testing Concurrent Processing...")
	concStart := time.Now()
	concurrentResults := processReportsConcurrently(jobs, 3)
	concurrentTime := time.Since(concStart)

	// Test pipeline processing
	fmt.Println("Testing Pipeline Processing...")
	pipelineStart := time.Now()
	pipelineResult := processPipeline(generateLargeDataset(5000))
	pipelineTime := time.Since(pipelineStart)

	// Test producer-consumer pattern
	fmt.Println("Testing Producer-Consumer Pattern...")
	producerStart := time.Now()
	producerResult := processProducerConsumer(1000)
	producerTime := time.Since(producerStart)

	// Display results
	displayResults(sequentialResults, concurrentResults)
	displayMetrics(ConcurrencyMetrics{
		SequentialTime: sequentialTime,
		ConcurrentTime: concurrentTime,
		Speedup:        sequentialTime.Seconds() / concurrentTime.Seconds(),
		Goroutines:     3,
		MemoryUsed:     getMemUsage(),
	})

	fmt.Printf("\nPipeline Processing: %v (rows: %d)\n", pipelineTime, pipelineResult)
	fmt.Printf("Producer-Consumer Processing: %v (rows: %d)\n", producerTime, producerResult)

	fmt.Println("\nConcurrent processing examples completed!")
	fmt.Println("Check the output directory for generated files.")
}

func generateReportJobs(jobCount, rowsPerJob int) []ReportJob {
	jobs := make([]ReportJob, jobCount)
	reportTypes := []string{"Sales", "Inventory", "Financial", "Marketing", "Operations"}

	for i := 0; i < jobCount; i++ {
		data := make([]ReportData, rowsPerJob)
		for j := 0; j < rowsPerJob; j++ {
			data[j] = ReportData{
				ID:       j + 1,
				Name:     fmt.Sprintf("Item %d-%d", i+1, j+1),
				Value:    rand.Float64() * 1000,
				Category: reportTypes[rand.Intn(len(reportTypes))],
				Date:     time.Now().AddDate(0, 0, -rand.Intn(30)),
			}
		}

		jobs[i] = ReportJob{
			ID:       i + 1,
			Name:     fmt.Sprintf("%s Report %d", reportTypes[i%len(reportTypes)], i+1),
			Data:     data,
			Filename: fmt.Sprintf("output/11-concurrent-report-%d.xlsx", i+1),
		}
	}

	return jobs
}

func processReportsSequentially(jobs []ReportJob) []ProcessingResult {
	results := make([]ProcessingResult, len(jobs))

	for i, job := range jobs {
		start := time.Now()
		err := generateReport(job)
		duration := time.Since(start)

		results[i] = ProcessingResult{
			JobID:    job.ID,
			Filename: job.Filename,
			Duration: duration,
			RowCount: len(job.Data),
			Error:    err,
		}

		fmt.Printf("Sequential: Completed job %d in %v\n", job.ID, duration)
	}

	return results
}

func processReportsConcurrently(jobs []ReportJob, numWorkers int) []ProcessingResult {
	jobChan := make(chan ReportJob, len(jobs))
	resultChan := make(chan ProcessingResult, len(jobs))

	// Start workers
	var wg sync.WaitGroup
	for i := 0; i < numWorkers; i++ {
		wg.Add(1)
		go func(workerID int) {
			defer wg.Done()
			for job := range jobChan {
				start := time.Now()
				err := generateReport(job)
				duration := time.Since(start)

				result := ProcessingResult{
					JobID:    job.ID,
					Filename: job.Filename,
					Duration: duration,
					RowCount: len(job.Data),
					Error:    err,
				}

				resultChan <- result
				fmt.Printf("Worker %d: Completed job %d in %v\n", workerID, job.ID, duration)
			}
		}(i)
	}

	// Send jobs
	go func() {
		defer close(jobChan)
		for _, job := range jobs {
			jobChan <- job
		}
	}()

	// Collect results
	go func() {
		wg.Wait()
		close(resultChan)
	}()

	var results []ProcessingResult
	for result := range resultChan {
		results = append(results, result)
	}

	return results
}

func generateReport(job ReportJob) error {
	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet(job.Name)

	// Header style
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

	// Add headers
	headers := []string{"ID", "Name", "Value", "Category", "Date"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(col+"1", header).SetStyle(headerStyle)
	}

	// Add data
	for i, row := range job.Data {
		rowNum := i + 2
		sheet.SetCell(fmt.Sprintf("A%d", rowNum), row.ID)
		sheet.SetCell(fmt.Sprintf("B%d", rowNum), row.Name)
		sheet.SetCell(fmt.Sprintf("C%d", rowNum), row.Value).SetStyle(&excelbuilder.Style{
			NumberFormat: "#,##0.00",
		})
		sheet.SetCell(fmt.Sprintf("D%d", rowNum), row.Category)
		sheet.SetCell(fmt.Sprintf("E%d", rowNum), row.Date.Format("2006-01-02"))
	}

	// Set column widths
	sheet.SetColumnWidth("A", 8)
	sheet.SetColumnWidth("B", 20)
	sheet.SetColumnWidth("C", 12)
	sheet.SetColumnWidth("D", 15)
	sheet.SetColumnWidth("E", 12)

	return builder.SaveToFile(job.Filename)
}

func processPipeline(data []ReportData) int {
	// Pipeline stages
	stage1 := make(chan ReportData, 100)
	stage2 := make(chan ReportData, 100)
	stage3 := make(chan ReportData, 100)

	var processedCount int64

	// Stage 1: Data validation
	go func() {
		defer close(stage1)
		for _, item := range data {
			// Simulate validation processing
			if item.Value > 0 {
				stage1 <- item
			}
		}
	}()

	// Stage 2: Data transformation
	go func() {
		defer close(stage2)
		for item := range stage1 {
			// Simulate transformation
			item.Value = item.Value * 1.1 // Apply 10% markup
			stage2 <- item
		}
	}()

	// Stage 3: Data aggregation
	go func() {
		defer close(stage3)
		for item := range stage2 {
			// Simulate aggregation
			stage3 <- item
		}
	}()

	// Final stage: Excel generation
	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Pipeline Result")

	// Headers
	headers := []string{"ID", "Name", "Processed Value", "Category", "Date"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(col+"1", header).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{Bold: true, Color: "#FFFFFF"},
			Fill: &excelbuilder.Fill{Type: "solid", Color: "#4472C4"},
		})
	}

	row := 2
	for item := range stage3 {
		sheet.SetCell(fmt.Sprintf("A%d", row), item.ID)
		sheet.SetCell(fmt.Sprintf("B%d", row), item.Name)
		sheet.SetCell(fmt.Sprintf("C%d", row), item.Value).SetStyle(&excelbuilder.Style{
			NumberFormat: "#,##0.00",
		})
		sheet.SetCell(fmt.Sprintf("D%d", row), item.Category)
		sheet.SetCell(fmt.Sprintf("E%d", row), item.Date.Format("2006-01-02"))
		row++
		atomic.AddInt64(&processedCount, 1)
	}

	builder.SaveToFile("output/11-concurrent-pipeline.xlsx")
	return int(processedCount)
}

func processProducerConsumer(dataSize int) int {
	dataChan := make(chan ReportData, 50)
	resultChan := make(chan ReportData, 50)
	var wg sync.WaitGroup

	// Producer
	wg.Add(1)
	go func() {
		defer wg.Done()
		defer close(dataChan)

		for i := 0; i < dataSize; i++ {
			data := ReportData{
				ID:       i + 1,
				Name:     fmt.Sprintf("Producer Item %d", i+1),
				Value:    rand.Float64() * 1000,
				Category: "Generated",
				Date:     time.Now(),
			}
			dataChan <- data
		}
	}()

	// Consumers
	numConsumers := 3
	for i := 0; i < numConsumers; i++ {
		wg.Add(1)
		go func(consumerID int) {
			defer wg.Done()
			for data := range dataChan {
				// Simulate processing
				time.Sleep(time.Microsecond * 10)
				data.Value = data.Value * 1.05 // Apply 5% processing fee
				resultChan <- data
			}
		}(i)
	}

	// Close result channel when all consumers are done
	go func() {
		wg.Wait()
		close(resultChan)
	}()

	// Collect results and generate Excel
	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Producer-Consumer")

	// Headers
	headers := []string{"ID", "Name", "Processed Value", "Category", "Date"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(col+"1", header).SetStyle(&excelbuilder.Style{
			Font: &excelbuilder.Font{Bold: true, Color: "#FFFFFF"},
			Fill: &excelbuilder.Fill{Type: "solid", Color: "#4472C4"},
		})
	}

	row := 2
	processedCount := 0
	for result := range resultChan {
		sheet.SetCell(fmt.Sprintf("A%d", row), result.ID)
		sheet.SetCell(fmt.Sprintf("B%d", row), result.Name)
		sheet.SetCell(fmt.Sprintf("C%d", row), result.Value).SetStyle(&excelbuilder.Style{
			NumberFormat: "#,##0.00",
		})
		sheet.SetCell(fmt.Sprintf("D%d", row), result.Category)
		sheet.SetCell(fmt.Sprintf("E%d", row), result.Date.Format("2006-01-02"))
		row++
		processedCount++
	}

	builder.SaveToFile("output/11-concurrent-producer-consumer.xlsx")
	return processedCount
}

func generateLargeDataset(size int) []ReportData {
	data := make([]ReportData, size)
	categories := []string{"A", "B", "C", "D", "E"}

	for i := 0; i < size; i++ {
		data[i] = ReportData{
			ID:       i + 1,
			Name:     fmt.Sprintf("Dataset Item %d", i+1),
			Value:    rand.Float64() * 1000,
			Category: categories[rand.Intn(len(categories))],
			Date:     time.Now().AddDate(0, 0, -rand.Intn(365)),
		}
	}

	return data
}

func displayResults(sequential, concurrent []ProcessingResult) {
	fmt.Println("\n" + strings.Repeat("=", 70))
	fmt.Println("PROCESSING RESULTS COMPARISON")
	fmt.Println(strings.Repeat("=", 70))

	fmt.Println("Sequential Processing:")
	var seqTotal time.Duration
	for _, result := range sequential {
		fmt.Printf("  Job %d: %v (%d rows)\n", result.JobID, result.Duration, result.RowCount)
		seqTotal += result.Duration
	}
	fmt.Printf("  Total: %v\n\n", seqTotal)

	fmt.Println("Concurrent Processing:")
	var concTotal time.Duration
	for _, result := range concurrent {
		fmt.Printf("  Job %d: %v (%d rows)\n", result.JobID, result.Duration, result.RowCount)
		if result.Duration > concTotal {
			concTotal = result.Duration // Max duration (parallel execution)
		}
	}
	fmt.Printf("  Total: %v\n", concTotal)
}

func displayMetrics(metrics ConcurrencyMetrics) {
	fmt.Println("\n" + strings.Repeat("-", 70))
	fmt.Println("CONCURRENCY METRICS")
	fmt.Println(strings.Repeat("-", 70))
	fmt.Printf("Sequential Time: %v\n", metrics.SequentialTime)
	fmt.Printf("Concurrent Time: %v\n", metrics.ConcurrentTime)
	fmt.Printf("Speedup: %.2fx\n", metrics.Speedup)
	fmt.Printf("Goroutines Used: %d\n", metrics.Goroutines)
	fmt.Printf("Memory Used: %d KB\n", metrics.MemoryUsed/1024)
	fmt.Printf("Efficiency: %.1f%%\n", (metrics.Speedup/float64(metrics.Goroutines))*100)
	fmt.Println(strings.Repeat("-", 70))
}

func getMemUsage() uint64 {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	return m.Alloc
}