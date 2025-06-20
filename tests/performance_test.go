package tests

import (
	"runtime"
	"sync"
	"testing"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

// Test Case 4.1: Large Dataset Performance Tests

// TestPerformance_LargeDataset : Test performance với large datasets
func TestPerformance_LargeDataset(t *testing.T) {
	// Test: Check performance with large datasets
	// Expected:
	// - Large datasets can be processed
	// - Performance is acceptable
	// - Memory usage is reasonable

	start := time.Now()
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("LargeDataset")

	// Add header
	headerRow := workbook.AddRow()
	for i := 0; i < 50; i++ {
		headerRow.AddCell("Column" + string(rune('A'+i))).Done()
	}
	headerRow.Done()

	// Add 10,000 rows of data
	for row := 0; row < 10000; row++ {
		dataRow := workbook.AddRow()
		for col := 0; col < 50; col++ {
			dataRow.AddCell(row*50 + col).Done()
		}
		dataRow.Done()
	}

	file := workbook.Build()
	elapsed := time.Since(start)

	assert.NotNil(t, file, "Expected workbook to build successfully")
	assert.Less(t, elapsed, 30*time.Second, "Expected large dataset processing to complete within 30 seconds")

	t.Logf("Large dataset (10,000 rows x 50 columns) processed in %v", elapsed)
}

// TestPerformance_MemoryUsage : Monitor memory usage during large operations
func TestPerformance_MemoryUsage(t *testing.T) {
	// Test: Check memory usage during large operations
	// Expected:
	// - Memory usage is reasonable
	// - No significant memory leaks
	// - Garbage collection works effectively

	var m1, m2 runtime.MemStats
	runtime.GC()
	runtime.ReadMemStats(&m1)

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("MemoryTest")

	// Create moderate dataset
	for row := 0; row < 5000; row++ {
		dataRow := workbook.AddRow()
		for col := 0; col < 20; col++ {
			dataRow.AddCell("Data_" + string(rune('A'+col)) + "_" + string(rune('0'+row%10))).Done()
		}
		dataRow.Done()
	}

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook to build successfully")

	runtime.GC()
	runtime.ReadMemStats(&m2)

	// Use current allocation as a reasonable check
	currentAlloc := m2.Alloc
	t.Logf("Current memory allocation: %d bytes (%.2f MB)", currentAlloc, float64(currentAlloc)/(1024*1024))

	// Memory usage should be reasonable (less than 100MB for this test)
	assert.Less(t, currentAlloc, uint64(100*1024*1024), "Expected current memory allocation to be less than 100MB")
}

// TestPerformance_StyleCaching : Verify style caching effectiveness
func TestPerformance_StyleCaching(t *testing.T) {
	// Test: Check style caching performance
	// Expected:
	// - Style caching improves performance
	// - Repeated styles are reused
	// - Cache hit rate is high

	styleManager := excelbuilder.NewStyleManager()
	file := excelize.NewFile()

	// Define common styles
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#D9E1F2",
		},
	}

	dataStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size: 10,
		},
	}

	start := time.Now()

	// Create many flyweights with same styles (should hit cache)
	for i := 0; i < 1000; i++ {
		headerFlyweight := styleManager.GetStyleFlyweight(headerStyle, file)
		dataFlyweight := styleManager.GetStyleFlyweight(dataStyle, file)

		assert.NotNil(t, headerFlyweight, "Expected header flyweight")
		assert.NotNil(t, dataFlyweight, "Expected data flyweight")
	}

	elapsed := time.Since(start)
	t.Logf("1000 style flyweight creations completed in %v", elapsed)

	// Should be very fast due to caching
	assert.Less(t, elapsed, 100*time.Millisecond, "Expected style caching to be very fast")

	// Verify cache statistics
	stats := styleManager.GetCacheStats()
	assert.Greater(t, stats.HitRate(), 0.9, "Expected cache hit rate > 90%")
	t.Logf("Cache hit rate: %.2f%%, Total requests: %d", stats.HitRate()*100, stats.TotalRequests())
}

// TestPerformance_ConcurrentAccess : Test concurrent access performance
func TestPerformance_ConcurrentAccess(t *testing.T) {
	// Test: Check concurrent access performance
	// Expected:
	// - Concurrent operations work correctly
	// - No race conditions
	// - Performance scales reasonably

	styleManager := excelbuilder.NewStyleManager()
	file := excelize.NewFile()
	var wg sync.WaitGroup
	workerCount := 10
	operationsPerWorker := 100

	style := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 11,
		},
	}

	start := time.Now()

	// Launch concurrent workers
	for worker := 0; worker < workerCount; worker++ {
		wg.Add(1)
		go func(workerID int) {
			defer wg.Done()
			for op := 0; op < operationsPerWorker; op++ {
				flyweight := styleManager.GetStyleFlyweight(style, file)
				assert.NotNil(t, flyweight, "Expected flyweight from worker %d, operation %d", workerID, op)
			}
		}(worker)
	}

	wg.Wait()
	elapsed := time.Since(start)

	t.Logf("Concurrent access test (%d workers, %d ops each) completed in %v", workerCount, operationsPerWorker, elapsed)
	assert.Less(t, elapsed, 5*time.Second, "Expected concurrent operations to complete within 5 seconds")

	// Verify final state
	stats := styleManager.GetCacheStats()
	assert.Equal(t, workerCount*operationsPerWorker, stats.TotalRequests(), "Expected total requests to match operations")
}

// Test Case 4.2: Memory Optimization Tests

// TestMemory_StyleFlyweightSharing : Verify flyweight pattern reduces memory
func TestMemory_StyleFlyweightSharing(t *testing.T) {
	// Test: Check flyweight pattern memory efficiency
	// Expected:
	// - Flyweight pattern reduces memory usage
	// - Same styles share flyweight instances
	// - Memory scales with unique styles, not total usage

	styleManager := excelbuilder.NewStyleManager()

	// Create same style multiple times
	style := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
		},
	}

	// Get multiple flyweights for same style
	file := excelize.NewFile()
	flyweights := make([]*excelbuilder.StyleFlyweight, 1000)
	for i := 0; i < 1000; i++ {
		flyweights[i] = styleManager.GetStyleFlyweight(style, file)
	}

	// All flyweights should be the same instance (memory sharing)
	for i := 1; i < 1000; i++ {
		assert.Same(t, flyweights[0], flyweights[i], "Expected same flyweight instance for identical styles")
	}

	// Verify cache contains only one entry
	stats := styleManager.GetCacheStats()
	assert.Equal(t, uint64(1), stats.UniqueStyles, "Expected only 1 unique style in cache")
	assert.Equal(t, 1000, stats.TotalRequests(), "Expected 1000 total requests")
	assert.Greater(t, stats.HitRate(), 0.99, "Expected very high cache hit rate")
}

// TestMemory_GarbageCollection : Test garbage collection effectiveness
func TestMemory_GarbageCollection(t *testing.T) {
	// Test: Check garbage collection effectiveness
	// Expected:
	// - Unused objects are garbage collected
	// - Memory is reclaimed properly
	// - No memory leaks

	var m1, m2, m3 runtime.MemStats

	// Initial memory state
	runtime.GC()
	runtime.ReadMemStats(&m1)

	// Create and discard many objects
	func() {
		builder := excelbuilder.New()
		for i := 0; i < 100; i++ {
			workbook := builder.NewWorkbook().AddSheet("TempSheet")
			for row := 0; row < 100; row++ {
				dataRow := workbook.AddRow()
				for col := 0; col < 10; col++ {
					dataRow.AddCell("temp_data").Done()
				}
				dataRow.Done()
			}
			// workbook goes out of scope here
		}
	}()

	// Memory after creating objects
	runtime.ReadMemStats(&m2)

	// Force garbage collection
	runtime.GC()
	runtime.GC() // Run twice to ensure cleanup
	runtime.ReadMemStats(&m3)

	memoryBeforeGC := m2.Alloc - m1.Alloc
	memoryAfterGC := m3.Alloc - m1.Alloc
	memoryReclaimed := memoryBeforeGC - memoryAfterGC

	t.Logf("Memory before GC: %d bytes", memoryBeforeGC)
	t.Logf("Memory after GC: %d bytes", memoryAfterGC)
	t.Logf("Memory reclaimed: %d bytes (%.1f%%)", memoryReclaimed, float64(memoryReclaimed)/float64(memoryBeforeGC)*100)

	// Should reclaim significant memory
	assert.Greater(t, float64(memoryReclaimed)/float64(memoryBeforeGC), 0.5, "Expected to reclaim at least 50% of memory")
}

// TestMemory_ResourceCleanup : Test proper resource cleanup
func TestMemory_ResourceCleanup(t *testing.T) {
	// Test: Check proper resource cleanup
	// Expected:
	// - Resources are cleaned up properly
	// - No resource leaks
	// - Finalizers work correctly

	var initialGoroutines int
	initialGoroutines = runtime.NumGoroutine()

	// Create and use resources
	func() {
		styleManager := excelbuilder.NewStyleManager()
		file := excelize.NewFile()
		builder := excelbuilder.New()

		// Create multiple workbooks with styles
		for i := 0; i < 10; i++ {
			workbook := builder.NewWorkbook().AddSheet("ResourceTest")
			style := excelbuilder.StyleConfig{
				Font: excelbuilder.FontConfig{
					Size: 10 + i,
				},
			}
			flyweight := styleManager.GetStyleFlyweight(style, file)
			assert.NotNil(t, flyweight, "Expected flyweight")

			workbook.AddRow().AddCell("test").SetStyle(style).Done().Done()
			builtFile := workbook.Build()
			assert.NotNil(t, builtFile, "Expected file")
		}
		// All resources should be cleaned up when function exits
	}()

	// Allow time for cleanup
	runtime.GC()
	time.Sleep(100 * time.Millisecond)

	finalGoroutines := runtime.NumGoroutine()
	t.Logf("Goroutines: initial=%d, final=%d", initialGoroutines, finalGoroutines)

	// Should not have significant goroutine leaks
	assert.LessOrEqual(t, finalGoroutines-initialGoroutines, 2, "Expected minimal goroutine increase")
}

// Test Case 4.3: Scalability Tests

// TestScalability_MultipleWorkbooks : Test handling multiple workbooks
func TestScalability_MultipleWorkbooks(t *testing.T) {
	// Test: Check scalability with multiple workbooks
	// Expected:
	// - Multiple workbooks can be handled
	// - Performance scales reasonably
	// - No interference between workbooks

	start := time.Now()
	builder := excelbuilder.New()
	workbookCount := 50
	workbooks := make([]*excelbuilder.WorkbookBuilder, workbookCount)
	sheets := make([]*excelbuilder.SheetBuilder, workbookCount)

	// Create multiple workbooks
	for i := 0; i < workbookCount; i++ {
		workbooks[i] = builder.NewWorkbook()
		sheets[i] = workbooks[i].AddSheet("Sheet" + string(rune('A'+i%26)))

		// Add some data to each workbook
		for row := 0; row < 100; row++ {
			dataRow := sheets[i].AddRow()
			for col := 0; col < 5; col++ {
				dataRow.AddCell(i*100 + row*5 + col).Done()
			}
			dataRow.Done()
		}
	}

	// Build all workbooks
	files := make([]*excelize.File, workbookCount)
	for i := 0; i < workbookCount; i++ {
		files[i] = workbooks[i].Build()
		assert.NotNil(t, files[i], "Expected workbook %d to build successfully", i)
	}

	elapsed := time.Since(start)
	t.Logf("Created %d workbooks (100 rows x 5 cols each) in %v", workbookCount, elapsed)

	// Should complete within reasonable time
	assert.Less(t, elapsed, 30*time.Second, "Expected multiple workbook creation to complete within 30 seconds")
}

// TestScalability_ConcurrentWorkbooks : Test concurrent workbook creation
func TestScalability_ConcurrentWorkbooks(t *testing.T) {
	// Test: Check concurrent workbook creation
	// Expected:
	// - Concurrent workbook creation works
	// - No race conditions
	// - Performance benefits from concurrency

	var wg sync.WaitGroup
	workerCount := 5
	workbooksPerWorker := 10
	results := make([][]*excelize.File, workerCount)

	start := time.Now()

	// Launch concurrent workers
	for worker := 0; worker < workerCount; worker++ {
		wg.Add(1)
		go func(workerID int) {
			defer wg.Done()
			builder := excelbuilder.New()
			results[workerID] = make([]*excelize.File, workbooksPerWorker)

			for wb := 0; wb < workbooksPerWorker; wb++ {
				workbookBuilder := builder.NewWorkbook()
				sheet := workbookBuilder.AddSheet("ConcurrentSheet")

				// Add test data
				for row := 0; row < 50; row++ {
					dataRow := sheet.AddRow()
					for col := 0; col < 3; col++ {
						dataRow.AddCell(workerID*1000 + wb*100 + row*3 + col).Done()
					}
					dataRow.Done()
				}

				results[workerID][wb] = workbookBuilder.Build()
			}
		}(worker)
	}

	wg.Wait()
	elapsed := time.Since(start)

	// Verify all workbooks were created successfully
	for worker := 0; worker < workerCount; worker++ {
		for wb := 0; wb < workbooksPerWorker; wb++ {
			assert.NotNil(t, results[worker][wb], "Expected workbook from worker %d, wb %d", worker, wb)
		}
	}

	totalWorkbooks := workerCount * workbooksPerWorker
	t.Logf("Created %d workbooks concurrently (%d workers) in %v", totalWorkbooks, workerCount, elapsed)

	// Should complete within reasonable time
	assert.Less(t, elapsed, 20*time.Second, "Expected concurrent workbook creation to complete within 20 seconds")
}

// TestScalability_LargeSheets : Test very large sheets
func TestScalability_LargeSheets(t *testing.T) {
	// Test: Check handling of very large sheets
	// Expected:
	// - Large sheets can be created
	// - Memory usage is reasonable
	// - Performance is acceptable

	if testing.Short() {
		t.Skip("Skipping large sheet test in short mode")
	}

	start := time.Now()
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("LargeSheet")

	// Create a very large sheet (20,000 rows x 20 columns)
	rowCount := 20000
	colCount := 20

	// Add header
	headerRow := workbook.AddRow()
	for col := 0; col < colCount; col++ {
		headerRow.AddCell("Col" + string(rune('A'+col%26))).Done()
	}
	headerRow.Done()

	// Add data rows
	for row := 0; row < rowCount; row++ {
		dataRow := workbook.AddRow()
		for col := 0; col < colCount; col++ {
			dataRow.AddCell(row*colCount + col).Done()
		}
		dataRow.Done()

		// Log progress every 5000 rows
		if (row+1)%5000 == 0 {
			t.Logf("Processed %d/%d rows", row+1, rowCount)
		}
	}

	file := workbook.Build()
	elapsed := time.Since(start)

	assert.NotNil(t, file, "Expected large sheet to build successfully")
	t.Logf("Large sheet (%d rows x %d columns) created in %v", rowCount, colCount, elapsed)

	// Should complete within reasonable time (2 minutes)
	assert.Less(t, elapsed, 2*time.Minute, "Expected large sheet creation to complete within 2 minutes")
}
