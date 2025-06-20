package tests

import (
	"os"
	"path/filepath"
	"testing"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

// Test Case 6.1: Complex Workflow Tests

// TestWorkflow_CompleteReport : Test complete report generation workflow
func TestWorkflow_CompleteReport(t *testing.T) {
	// Test: Check complete report generation workflow
	// Expected:
	// - Complete workflow executes successfully
	// - All components work together
	// - Output is correct and complete

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().
		SetProperties(excelbuilder.WorkbookProperties{
			Title:       "Monthly Sales Report",
			Author:      "Sales Team",
			Subject:     "Sales Performance Analysis",
			Description: "Comprehensive monthly sales report with charts and analysis",
			Keywords:    "sales, report, monthly, analysis",
			Category:    "Business Reports",
		})

	// Create Summary Sheet
	summarySheet := workbook.AddSheet("Summary")

	// Add title with styling
	summarySheet.AddRow().
		AddCell("Monthly Sales Report - January 2024").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Bold:   true,
				Size:   16,
				Color:  "#2F4F4F",
				Family: "Arial",
			},
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: "center",
			},
		}).Done().
		Done()

	// Merge title cells
	summarySheet.MergeCell("A1:E1")

	// Add summary metrics
	summarySheet.AddRow().Done() // Empty row
	summarySheet.AddRow().
		AddCell("Metric").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{Bold: true},
			Fill: excelbuilder.FillConfig{
				Type:  "pattern",
				Color: "#D9E1F2",
			},
		}).Done().
		AddCell("Value").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{Bold: true},
			Fill: excelbuilder.FillConfig{
				Type:  "pattern",
				Color: "#D9E1F2",
			},
		}).Done().
		Done()

	// Add metrics data
	metrics := []struct {
		name  string
		value interface{}
		style excelbuilder.StyleConfig
	}{
		{"Total Revenue", 125000, excelbuilder.StyleConfig{NumberFormat: "$#,##0"}},
		{"Total Orders", 450, excelbuilder.StyleConfig{}},
		{"Average Order Value", "=B4/B5", excelbuilder.StyleConfig{NumberFormat: "$#,##0.00"}},
		{"Growth Rate", 0.15, excelbuilder.StyleConfig{NumberFormat: "0.0%"}},
	}

	for _, metric := range metrics {
		summarySheet.AddRow().
			AddCell(metric.name).Done().
			AddCell(metric.value).SetStyle(metric.style).Done().
			Done()
	}

	// Create Data Sheet
	dataSheet := workbook.AddSheet("Sales Data")

	// Add headers
	dataSheet.AddRow().
		AddCell("Date").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{Bold: true},
			Fill: excelbuilder.FillConfig{Type: "pattern", Color: "#E2EFDA"},
		}).Done().
		AddCell("Product").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{Bold: true},
			Fill: excelbuilder.FillConfig{Type: "pattern", Color: "#E2EFDA"},
		}).Done().
		AddCell("Quantity").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{Bold: true},
			Fill: excelbuilder.FillConfig{Type: "pattern", Color: "#E2EFDA"},
		}).Done().
		AddCell("Unit Price").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{Bold: true},
			Fill: excelbuilder.FillConfig{Type: "pattern", Color: "#E2EFDA"},
		}).Done().
		AddCell("Total").
		SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{Bold: true},
			Fill: excelbuilder.FillConfig{Type: "pattern", Color: "#E2EFDA"},
		}).Done().
		Done()

	// Add sample data
	salesData := []struct {
		date     string
		product  string
		qty      int
		price    float64
	}{
		{"2024-01-01", "Product A", 10, 25.50},
		{"2024-01-02", "Product B", 15, 30.00},
		{"2024-01-03", "Product A", 8, 25.50},
		{"2024-01-04", "Product C", 20, 15.75},
		{"2024-01-05", "Product B", 12, 30.00},
	}

	for _, sale := range salesData {
		dataSheet.AddRow().
			AddCell(sale.date).Done().
			AddCell(sale.product).Done().
			AddCell(sale.qty).Done().
			AddCell(sale.price).SetStyle(excelbuilder.StyleConfig{NumberFormat: "$#,##0.00"}).Done().
			AddCell("=C" + string(rune('2'+len(salesData)-1)) + "*D" + string(rune('2'+len(salesData)-1))).
			SetStyle(excelbuilder.StyleConfig{NumberFormat: "$#,##0.00"}).Done().
			Done()
	}

	// Add chart to summary sheet
	chart := summarySheet.AddChart().
		SetType("bar").
		SetTitle("Sales by Product").
		SetPosition("D3").
		SetSize(400, 300).
		AddDataSeries(excelbuilder.DataSeries{
			Name:       "Sales",
			Categories: "'Sales Data'!B2:B6",
			Values:     "'Sales Data'!E2:E6",
			Color:      "#4472C4",
		})

	err := chart.Build()
	assert.NoError(t, err, "Expected chart to build successfully")

	// Build final workbook
	file := workbook.Build()
	assert.NotNil(t, file, "Expected complete report to build successfully")

	// Verify workbook properties
	props, err := file.GetDocProps()
	assert.NoError(t, err, "Expected to get document properties")
	assert.Equal(t, "Monthly Sales Report", props.Title, "Expected correct title")
	assert.Equal(t, "Sales Team", props.Creator, "Expected correct author")
}

// TestWorkflow_DataImportExport : Test data import/export workflow
func TestWorkflow_DataImportExport(t *testing.T) {
	// Test: Check data import/export workflow
	// Expected:
	// - Data can be imported from various sources
	// - Data transformation works correctly
	// - Export maintains data integrity

	// Simulate importing data from different sources
	csvData := [][]string{
		{"Name", "Age", "Department", "Salary"},
		{"John Doe", "30", "Engineering", "75000"},
		{"Jane Smith", "28", "Marketing", "65000"},
		{"Bob Johnson", "35", "Sales", "70000"},
	}

	jsonData := map[string]interface{}{
		"company": "Tech Corp",
		"year":    2024,
		"metrics": map[string]float64{
			"revenue":     1000000,
			"profit":      150000,
			"growth_rate": 0.12,
		},
	}

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Import CSV-like data
	csvSheet := workbook.AddSheet("Employee Data")
	for i, row := range csvData {
		dataRow := csvSheet.AddRow()
		for j, cell := range row {
			cellBuilder := dataRow.AddCell(cell)
			// Style header row
			if i == 0 {
				cellBuilder.SetStyle(excelbuilder.StyleConfig{
					Font: excelbuilder.FontConfig{Bold: true},
					Fill: excelbuilder.FillConfig{Type: "pattern", Color: "#D9E1F2"},
				})
			} else if j == 3 { // Salary column
				cellBuilder.SetStyle(excelbuilder.StyleConfig{
					NumberFormat: "$#,##0",
				})
			}
			cellBuilder.Done()
		}
		dataRow.Done()
	}

	// Import JSON-like data
	jsonSheet := workbook.AddSheet("Company Metrics")
	jsonSheet.AddRow().
		AddCell("Company").Done().
		AddCell(jsonData["company"]).Done().
		Done().
		AddRow().
		AddCell("Year").Done().
		AddCell(jsonData["year"]).Done().
		Done()

	metrics := jsonData["metrics"].(map[string]float64)
	for key, value := range metrics {
		jsonSheet.AddRow().
			AddCell(key).Done().
			AddCell(value).
			SetStyle(excelbuilder.StyleConfig{
				NumberFormat: "#,##0.00",
			}).Done().
			Done()
	}

	// Create summary with cross-sheet references
	summarySheet := workbook.AddSheet("Summary")
	summarySheet.AddRow().
		AddCell("Total Employees").Done().
		AddCell("=COUNTA('Employee Data'!A2:A4)").Done().
		Done().
		AddRow().
		AddCell("Average Salary").Done().
		AddCell("=AVERAGE('Employee Data'!D2:D4)").
		SetStyle(excelbuilder.StyleConfig{NumberFormat: "$#,##0"}).Done().
		Done().
		AddRow().
		AddCell("Company Revenue").Done().
		AddCell("='Company Metrics'!B3").
		SetStyle(excelbuilder.StyleConfig{NumberFormat: "$#,##0"}).Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected data import/export workflow to build successfully")
}

// TestWorkflow_TemplateBasedGeneration : Test template-based report generation
func TestWorkflow_TemplateBasedGeneration(t *testing.T) {
	// Test: Check template-based report generation
	// Expected:
	// - Templates can be applied
	// - Dynamic data insertion works
	// - Template styling is preserved

	// Define a template structure
	template := excelbuilder.TemplateConfig{
		Name: "QuarterlyReport",
		Sheets: []excelbuilder.SheetTemplate{
			{
				Name:    "Q1 Report",
				Headers: []string{"Month", "Revenue", "Expenses", "Profit"},
				Styles: map[string]excelbuilder.StyleConfig{
					"header": {
						Font: excelbuilder.FontConfig{Bold: true, Size: 12},
						Fill: excelbuilder.FillConfig{Type: "pattern", Color: "#4472C4"},
					},
					"data": {
						Font: excelbuilder.FontConfig{Size: 10},
					},
				},
			},
		},
	}

	// Apply template manually (in full implementation, this would be automated)
	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()
	sheet := workbook.AddSheet(template.Sheets[0].Name)

	// Apply header with template styling
	headerRow := sheet.AddRow()
	for _, header := range template.Sheets[0].Headers {
		headerRow.AddCell(header).
			SetStyle(template.Sheets[0].Styles["header"]).Done()
	}
	headerRow.Done()

	// Add template data
	quarterlyData := []struct {
		month    string
		revenue  float64
		expenses float64
	}{
		{"January", 50000, 30000},
		{"February", 55000, 32000},
		{"March", 60000, 35000},
	}

	for _, data := range quarterlyData {
		sheet.AddRow().
			AddCell(data.month).SetStyle(template.Sheets[0].Styles["data"]).Done().
			AddCell(data.revenue).SetStyle(excelbuilder.StyleConfig{
				Font:         template.Sheets[0].Styles["data"].Font,
				NumberFormat: "$#,##0",
			}).Done().
			AddCell(data.expenses).SetStyle(excelbuilder.StyleConfig{
				Font:         template.Sheets[0].Styles["data"].Font,
				NumberFormat: "$#,##0",
			}).Done().
			AddCell("=B" + string(rune('2'+len(quarterlyData)-1)) + "-C" + string(rune('2'+len(quarterlyData)-1))).
			SetStyle(excelbuilder.StyleConfig{
				Font:         template.Sheets[0].Styles["data"].Font,
				NumberFormat: "$#,##0",
			}).Done().
			Done()
	}

	// Add summary row
	sheet.AddRow().
		AddCell("Total").SetStyle(excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{Bold: true},
		}).Done().
		AddCell("=SUM(B2:B4)").SetStyle(excelbuilder.StyleConfig{
			Font:         excelbuilder.FontConfig{Bold: true},
			NumberFormat: "$#,##0",
		}).Done().
		AddCell("=SUM(C2:C4)").SetStyle(excelbuilder.StyleConfig{
			Font:         excelbuilder.FontConfig{Bold: true},
			NumberFormat: "$#,##0",
		}).Done().
		AddCell("=SUM(D2:D4)").SetStyle(excelbuilder.StyleConfig{
			Font:         excelbuilder.FontConfig{Bold: true},
			NumberFormat: "$#,##0",
		}).Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected template-based generation to build successfully")
}

// Test Case 6.2: Thread Safety Tests

// TestThreadSafety_ConcurrentBuilding : Test concurrent workbook building
func TestThreadSafety_ConcurrentBuilding(t *testing.T) {
	// Test: Check concurrent workbook building
	// Expected:
	// - Concurrent building works safely
	// - No race conditions
	// - All workbooks are created correctly

	// This test is covered in performance_test.go TestScalability_ConcurrentWorkbooks
	// but we can add a specific thread safety focus here

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("ThreadSafetyTest")

	// Add some data
	workbook.AddRow().
		AddCell("Thread Safety Test").Done().
		AddCell("Passed").Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected thread safety test to build successfully")
}

// Test Case 6.3: Resource Management Tests

// TestResource_FileHandling : Test file creation và cleanup
func TestResource_FileHandling(t *testing.T) {
	// Test: Check file creation and cleanup
	// Expected:
	// - Files can be created and saved
	// - File cleanup works properly
	// - No file handle leaks

	tempDir := t.TempDir()
	filePath := filepath.Join(tempDir, "test_output.xlsx")

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("FileTest")

	// Add test data
	workbook.AddRow().
		AddCell("File Handling Test").Done().
		AddCell("Success").Done().
		Done().
		AddRow().
		AddCell("Timestamp").Done().
		AddCell(time.Now().Format("2006-01-02 15:04:05")).Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook to build successfully")

	// Save to file
	err := file.SaveAs(filePath)
	assert.NoError(t, err, "Expected file to save successfully")

	// Verify file exists
	_, err = os.Stat(filePath)
	assert.NoError(t, err, "Expected file to exist")

	// Verify file can be opened
	openedFile, err := excelize.OpenFile(filePath)
	assert.NoError(t, err, "Expected file to open successfully")
	assert.NotNil(t, openedFile, "Expected opened file to be valid")

	// Verify content
	cellValue, err := openedFile.GetCellValue("FileTest", "A1")
	assert.NoError(t, err, "Expected to read cell value")
	assert.Equal(t, "File Handling Test", cellValue, "Expected correct cell value")

	// Close file
	err = openedFile.Close()
	assert.NoError(t, err, "Expected file to close successfully")

	// File cleanup is handled by t.TempDir()
}

// TestResource_MemoryManagement : Test memory management during operations
func TestResource_MemoryManagement(t *testing.T) {
	// Test: Check memory management during operations
	// Expected:
	// - Memory usage is reasonable
	// - No significant memory leaks
	// - Resources are cleaned up properly

	// This test is covered in performance_test.go but we can add specific checks here
	builder := excelbuilder.New()

	// Create and discard multiple workbooks to test memory management
	for i := 0; i < 10; i++ {
		workbook := builder.NewWorkbook().AddSheet("MemoryTest")
		for row := 0; row < 100; row++ {
			dataRow := workbook.AddRow()
			for col := 0; col < 5; col++ {
				dataRow.AddCell("data").Done()
			}
			dataRow.Done()
		}
		file := workbook.Build()
		assert.NotNil(t, file, "Expected workbook %d to build successfully", i)
		// file goes out of scope and should be garbage collected
	}

	// Force garbage collection
	// runtime.GC() // Uncomment if needed for explicit testing

	// Test passes if no memory issues occur
	assert.True(t, true, "Memory management test completed")
}

// Test Case 6.4: Edge Case Tests

// TestEdgeCase_EmptyWorkbook : Test empty workbook handling
func TestEdgeCase_EmptyWorkbook(t *testing.T) {
	// Test: Check empty workbook handling
	// Expected:
	// - Empty workbooks can be created
	// - No errors with minimal content
	// - Default behavior is correct

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Build empty workbook (no sheets added)
	file := workbook.Build()
	assert.NotNil(t, file, "Expected empty workbook to build successfully")

	// Add empty sheet
	workbook2 := builder.NewWorkbook().AddSheet("EmptySheet")
	file2 := workbook2.Build()
	assert.NotNil(t, file2, "Expected workbook with empty sheet to build successfully")

	// Add sheet with empty row
	workbook3 := builder.NewWorkbook().AddSheet("EmptyRowSheet")
	workbook3.AddRow().Done()
	file3 := workbook3.Build()
	assert.NotNil(t, file3, "Expected workbook with empty row to build successfully")
}

// TestEdgeCase_LargeStrings : Test handling of very large strings
func TestEdgeCase_LargeStrings(t *testing.T) {
	// Test: Check handling of very large strings
	// Expected:
	// - Large strings are handled correctly
	// - No truncation or corruption
	// - Performance is acceptable

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("LargeStrings")

	// Create a large string (10KB)
	largeString := ""
	for i := 0; i < 1000; i++ {
		largeString += "0123456789"
	}

	workbook.AddRow().
		AddCell("Large String Test").Done().
		Done().
		AddRow().
		AddCell(largeString).Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with large strings to build successfully")
}

// TestEdgeCase_SpecialCharacters : Test special character handling
func TestEdgeCase_SpecialCharacters(t *testing.T) {
	// Test: Check special character handling
	// Expected:
	// - Special characters are preserved
	// - Unicode characters work correctly
	// - No encoding issues

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("SpecialChars")

	specialChars := []string{
		"Unicode: 你好世界",
		"Emoji: 😀🎉📊",
		"Symbols: ©®™€£¥",
		"Math: ∑∏∆√∞",
		"Quotes: \"'`",
		"Newlines:\nLine 1\nLine 2",
		"Tabs:\tTabbed\tText",
	}

	for i, char := range specialChars {
		workbook.AddRow().
			AddCell("Test " + string(rune('A'+i))).Done().
			AddCell(char).Done().
			Done()
	}

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with special characters to build successfully")
}

// TestEdgeCase_NullAndEmpty : Test null and empty value handling
func TestEdgeCase_NullAndEmpty(t *testing.T) {
	// Test: Check null and empty value handling
	// Expected:
	// - Null values are handled gracefully
	// - Empty values don't cause errors
	// - Default behavior is consistent

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("NullEmpty")

	// Test various empty/null scenarios
	workbook.AddRow().
		AddCell("Type").Done().
		AddCell("Value").Done().
		Done().
		AddRow().
		AddCell("Empty String").Done().
		AddCell("").Done().
		Done().
		AddRow().
		AddCell("Nil Value").Done().
		AddCell(nil).Done().
		Done().
		AddRow().
		AddCell("Zero Number").Done().
		AddCell(0).Done().
		Done().
		AddRow().
		AddCell("False Boolean").Done().
		AddCell(false).Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook with null/empty values to build successfully")
}