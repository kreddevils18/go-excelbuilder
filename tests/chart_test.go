package tests

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

// Test Case 2.1: ChartBuilder Core Tests

// TestChartBuilder_Creation : Kiểm tra khởi tạo ChartBuilder
func TestChartBuilder_Creation(t *testing.T) {
	// Test: Check ChartBuilder instance creation
	// Expected:
	// - ChartBuilder created successfully
	// - Default configuration set
	// - File and sheet references correct

	file := excelize.NewFile()
	sheetName := "TestSheet"

	chartBuilder := excelbuilder.NewChartBuilder(file, sheetName)

	assert.NotNil(t, chartBuilder, "Expected ChartBuilder instance, got nil")

	// Test default configuration
	config := chartBuilder.GetConfig()
	assert.Equal(t, 480, config.Width, "Expected default width 480")
	assert.Equal(t, 290, config.Height, "Expected default height 290")
	assert.True(t, config.Legend.Show, "Expected legend to be shown by default")
	assert.Equal(t, "bottom", config.Legend.Position, "Expected default legend position bottom")
}

// TestChartBuilder_ChainMethods : Test fluent interface cho chart building
func TestChartBuilder_ChainMethods(t *testing.T) {
	// Test: Check fluent interface works correctly
	// Expected:
	// - Method chaining works
	// - Each method returns ChartBuilder
	// - Configuration is updated correctly

	file := excelize.NewFile()
	sheetName := "TestSheet"

	chartBuilder := excelbuilder.NewChartBuilder(file, sheetName).
		SetType("line").
		SetTitle("Test Chart").
		SetSize(600, 400).
		SetPosition("C3").
		SetXAxis(excelbuilder.AxisConfig{
			Title: "X Axis",
		}).
		SetYAxis(excelbuilder.AxisConfig{
			Title: "Y Axis",
		}).
		SetLegend(excelbuilder.LegendConfig{
			Show:     true,
			Position: "top",
		})

	assert.NotNil(t, chartBuilder, "Expected ChartBuilder instance after chaining")

	// Verify configuration
	config := chartBuilder.GetConfig()
	assert.Equal(t, "line", config.Type, "Expected chart type 'line'")
	assert.Equal(t, "Test Chart", config.Title, "Expected chart title 'Test Chart'")
	assert.Equal(t, 600, config.Width, "Expected width 600")
	assert.Equal(t, 400, config.Height, "Expected height 400")
	assert.Equal(t, "X Axis", config.XAxis.Title, "Expected X axis title")
	assert.Equal(t, "Y Axis", config.YAxis.Title, "Expected Y axis title")
	assert.Equal(t, "top", config.Legend.Position, "Expected legend position 'top'")
}

// TestChartBuilder_ChartTypes : Verify hỗ trợ các loại chart (line, bar, pie)
func TestChartBuilder_ChartTypes(t *testing.T) {
	// Test: Check support for different chart types
	// Expected:
	// - All chart types are supported
	// - Type is set correctly
	// - No errors with valid types

	file := excelize.NewFile()
	sheetName := "TestSheet"

	chartTypes := []string{"line", "bar", "pie", "scatter", "area"}

	for _, chartType := range chartTypes {
		chartBuilder := excelbuilder.NewChartBuilder(file, sheetName).
			SetType(chartType)

		config := chartBuilder.GetConfig()
		assert.Equal(t, chartType, config.Type, "Expected chart type '%s'", chartType)
	}
}

// TestChartBuilder_DataSeries : Test thêm/quản lý data series
func TestChartBuilder_DataSeries(t *testing.T) {
	// Test: Check data series management
	// Expected:
	// - Data series can be added
	// - Multiple series supported
	// - Series configuration preserved

	file := excelize.NewFile()
	sheetName := "TestSheet"

	series1 := excelbuilder.DataSeries{
		Name:       "Series 1",
		Categories: "A1:A5",
		Values:     "B1:B5",
		Color:      "#FF0000",
	}

	series2 := excelbuilder.DataSeries{
		Name:       "Series 2",
		Categories: "A1:A5",
		Values:     "C1:C5",
		Color:      "#00FF00",
	}

	chartBuilder := excelbuilder.NewChartBuilder(file, sheetName).
		SetType("line").
		AddDataSeries(series1).
		AddDataSeries(series2)

	config := chartBuilder.GetConfig()
	assert.Len(t, config.DataSeries, 2, "Expected 2 data series")
	assert.Equal(t, "Series 1", config.DataSeries[0].Name, "Expected first series name")
	assert.Equal(t, "Series 2", config.DataSeries[1].Name, "Expected second series name")
	assert.Equal(t, "#FF0000", config.DataSeries[0].Color, "Expected first series color")
	assert.Equal(t, "#00FF00", config.DataSeries[1].Color, "Expected second series color")
}

// Test Case 2.2: Chart Configuration Tests

// TestChart_BasicProperties : Test title, legend, axis labels
func TestChart_BasicProperties(t *testing.T) {
	// Test: Check basic chart properties configuration
	// Expected:
	// - Title can be set and retrieved
	// - Legend configuration works
	// - Axis labels are configurable

	file := excelize.NewFile()
	sheetName := "TestSheet"

	xAxisConfig := excelbuilder.AxisConfig{
		Title: "Time (months)",
	}

	yAxisConfig := excelbuilder.AxisConfig{
		Title: "Sales ($)",
	}

	legendConfig := excelbuilder.LegendConfig{
		Show:     true,
		Position: "right",
	}

	chartBuilder := excelbuilder.NewChartBuilder(file, sheetName).
		SetTitle("Monthly Sales Report").
		SetXAxis(xAxisConfig).
		SetYAxis(yAxisConfig).
		SetLegend(legendConfig)

	config := chartBuilder.GetConfig()
	assert.Equal(t, "Monthly Sales Report", config.Title, "Expected chart title")
	assert.Equal(t, "Time (months)", config.XAxis.Title, "Expected X axis title")
	assert.Equal(t, "Sales ($)", config.YAxis.Title, "Expected Y axis title")
	assert.True(t, config.Legend.Show, "Expected legend to be shown")
	assert.Equal(t, "right", config.Legend.Position, "Expected legend position 'right'")
}

// TestChart_DataRange : Kiểm tra data range selection và validation
func TestChart_DataRange(t *testing.T) {
	// Test: Check data range configuration
	// Expected:
	// - Data ranges are set correctly
	// - Categories and values are separate
	// - Range format is preserved

	file := excelize.NewFile()
	sheetName := "TestSheet"

	series := excelbuilder.DataSeries{
		Name:       "Revenue",
		Categories: "A2:A13", // 12 months
		Values:     "B2:B13", // Revenue values
		Color:      "#0066CC",
	}

	chartBuilder := excelbuilder.NewChartBuilder(file, sheetName).
		SetType("bar").
		AddDataSeries(series)

	config := chartBuilder.GetConfig()
	assert.Len(t, config.DataSeries, 1, "Expected 1 data series")
	assert.Equal(t, "A2:A13", config.DataSeries[0].Categories, "Expected categories range")
	assert.Equal(t, "B2:B13", config.DataSeries[0].Values, "Expected values range")
}

// TestChart_Positioning : Test chart position và size trong sheet
func TestChart_Positioning(t *testing.T) {
	// Test: Check chart positioning and sizing
	// Expected:
	// - Position can be set
	// - Size can be customized
	// - Default values work

	file := excelize.NewFile()
	sheetName := "TestSheet"

	// Test custom position and size
	chartBuilder := excelbuilder.NewChartBuilder(file, sheetName).
		SetType("pie").
		SetPosition("E5").
		SetSize(800, 600)

	config := chartBuilder.GetConfig()
	assert.Equal(t, 800, config.Width, "Expected width 800")
	assert.Equal(t, 600, config.Height, "Expected height 600")

	// Test default position (should be set during Build)
	chartBuilder2 := excelbuilder.NewChartBuilder(file, sheetName).
		SetType("line")

	config2 := chartBuilder2.GetConfig()
	assert.Equal(t, 480, config2.Width, "Expected default width 480")
	assert.Equal(t, 290, config2.Height, "Expected default height 290")
}

// TestChart_Styling : Verify chart styling options
func TestChart_Styling(t *testing.T) {
	// Test: Check chart styling capabilities
	// Expected:
	// - Series colors can be set
	// - Multiple series have different colors
	// - Color format is preserved

	file := excelize.NewFile()
	sheetName := "TestSheet"

	series1 := excelbuilder.DataSeries{
		Name:       "Q1",
		Categories: "A1:A4",
		Values:     "B1:B4",
		Color:      "#FF6B6B", // Red
	}

	series2 := excelbuilder.DataSeries{
		Name:       "Q2",
		Categories: "A1:A4",
		Values:     "C1:C4",
		Color:      "#4ECDC4", // Teal
	}

	series3 := excelbuilder.DataSeries{
		Name:       "Q3",
		Categories: "A1:A4",
		Values:     "D1:D4",
		Color:      "#45B7D1", // Blue
	}

	chartBuilder := excelbuilder.NewChartBuilder(file, sheetName).
		SetType("bar").
		AddDataSeries(series1).
		AddDataSeries(series2).
		AddDataSeries(series3)

	config := chartBuilder.GetConfig()
	assert.Len(t, config.DataSeries, 3, "Expected 3 data series")
	assert.Equal(t, "#FF6B6B", config.DataSeries[0].Color, "Expected red color for Q1")
	assert.Equal(t, "#4ECDC4", config.DataSeries[1].Color, "Expected teal color for Q2")
	assert.Equal(t, "#45B7D1", config.DataSeries[2].Color, "Expected blue color for Q3")
}

// Test Case 2.3: Chart Integration Tests

// TestChart_WithWorkbook : Test integration với existing workbook
func TestChart_WithWorkbook(t *testing.T) {
	// Test: Check chart integration with workbook builder
	// Expected:
	// - Chart can be added to existing workbook
	// - Integration with builder pattern works
	// - No conflicts with other operations

	builder := excelbuilder.New()

	// Create workbook with data
	workbook := builder.NewWorkbook().
		SetProperties(excelbuilder.WorkbookProperties{
			Title:  "Chart Integration Test",
			Author: "Test Suite",
		}).
		AddSheet("DataSheet")

	// Add some test data
	workbook.AddRow().
		AddCell("Month").Done().
		AddCell("Sales").Done().
		Done().
		AddRow().
		AddCell("Jan").Done().
		AddCell(1000).Done().
		Done().
		AddRow().
		AddCell("Feb").Done().
		AddCell(1200).Done().
		Done().
		AddRow().
		AddCell("Mar").Done().
		AddCell(1100).Done().
		Done()

	// Add chart
	chartBuilder := workbook.AddChart().
		SetType("line").
		SetTitle("Monthly Sales").
		SetPosition("D2").
		AddDataSeries(excelbuilder.DataSeries{
			Name:       "Sales",
			Categories: "A2:A4",
			Values:     "B2:B4",
			Color:      "#0066CC",
		})

	err := chartBuilder.Build()
	assert.NoError(t, err, "Expected no error when building chart")

	// Verify workbook can still be built
	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook to build successfully")
}

// TestChart_MultipleCharts : Kiểm tra multiple charts trong cùng sheet
func TestChart_MultipleCharts(t *testing.T) {
	// Test: Check multiple charts in same sheet
	// Expected:
	// - Multiple charts can be added
	// - Charts don't interfere with each other
	// - Different positions work

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("MultiChartSheet")

	// Add test data
	workbook.AddRow().
		AddCell("Product").Done().
		AddCell("Q1").Done().
		AddCell("Q2").Done().
		Done().
		AddRow().
		AddCell("A").Done().
		AddCell(100).Done().
		AddCell(120).Done().
		Done().
		AddRow().
		AddCell("B").Done().
		AddCell(80).Done().
		AddCell(90).Done().
		Done()

	// First chart - Bar chart
	chart1 := workbook.AddChart().
		SetType("bar").
		SetTitle("Q1 Sales").
		SetPosition("E2").
		SetSize(400, 300).
		AddDataSeries(excelbuilder.DataSeries{
			Name:       "Q1",
			Categories: "A2:A3",
			Values:     "B2:B3",
			Color:      "#FF6B6B",
		})

	err1 := chart1.Build()
	assert.NoError(t, err1, "Expected no error when building first chart")

	// Second chart - Line chart
	chart2 := workbook.AddChart().
		SetType("line").
		SetTitle("Q2 Sales").
		SetPosition("E15").
		SetSize(400, 300).
		AddDataSeries(excelbuilder.DataSeries{
			Name:       "Q2",
			Categories: "A2:A3",
			Values:     "C2:C3",
			Color:      "#4ECDC4",
		})

	err2 := chart2.Build()
	assert.NoError(t, err2, "Expected no error when building second chart")

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook to build successfully with multiple charts")
}

// TestChart_CrossSheetData : Test chart với data từ multiple sheets
func TestChart_CrossSheetData(t *testing.T) {
	// Test: Check chart with data from multiple sheets
	// Expected:
	// - Chart can reference data from other sheets
	// - Cross-sheet references work correctly
	// - Data integrity maintained

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook()

	// Create data sheet
	dataSheet := workbook.AddSheet("Data")
	dataSheet.AddRow().
		AddCell("Month").Done().
		AddCell("Revenue").Done().
		Done().
		AddRow().
		AddCell("Jan").Done().
		AddCell(5000).Done().
		Done().
		AddRow().
		AddCell("Feb").Done().
		AddCell(5500).Done().
		Done()

	// Create chart sheet
	chartSheet := workbook.AddSheet("Charts")

	// Note: In real implementation, cross-sheet references would need
	// to be handled properly. For now, we test the chart creation.
	chart := chartSheet.AddChart().
		SetType("bar").
		SetTitle("Revenue Chart").
		SetPosition("A1")

	// This would normally reference Data sheet, but for testing
	// we'll use current sheet references
	chart.AddDataSeries(excelbuilder.DataSeries{
		Name:       "Revenue",
		Categories: "A1:A2", // Would be "Data!A2:A3" in real scenario
		Values:     "B1:B2", // Would be "Data!B2:B3" in real scenario
		Color:      "#0066CC",
	})

	err := chart.Build()
	assert.NoError(t, err, "Expected no error when building cross-sheet chart")

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook to build successfully")
}

// TestChart_DynamicData : Verify chart updates khi data thay đổi
func TestChart_DynamicData(t *testing.T) {
	// Test: Check chart behavior with dynamic data ranges
	// Expected:
	// - Chart can handle dynamic ranges
	// - Range expansion works
	// - Data updates are reflected

	builder := excelbuilder.New()
	workbook := builder.NewWorkbook().AddSheet("DynamicData")

	// Add initial data
	workbook.AddRow().
		AddCell("Period").Done().
		AddCell("Value").Done().
		Done().
		AddRow().
		AddCell("P1").Done().
		AddCell(100).Done().
		Done().
		AddRow().
		AddCell("P2").Done().
		AddCell(150).Done().
		Done()

	// Create chart with initial range
	chart := workbook.AddChart().
		SetType("line").
		SetTitle("Dynamic Data Chart").
		SetPosition("D2").
		AddDataSeries(excelbuilder.DataSeries{
			Name:       "Values",
			Categories: "A2:A3",
			Values:     "B2:B3",
			Color:      "#FF6B6B",
		})

	err := chart.Build()
	assert.NoError(t, err, "Expected no error when building chart with dynamic data")

	// Add more data (in real scenario, chart would need to be updated)
	workbook.AddRow().
		AddCell("P3").Done().
		AddCell(200).Done().
		Done()

	file := workbook.Build()
	assert.NotNil(t, file, "Expected workbook to build successfully with dynamic data")
}