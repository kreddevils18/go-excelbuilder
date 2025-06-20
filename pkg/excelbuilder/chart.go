package excelbuilder

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// ChartBuilder provides a fluent interface for creating charts
type ChartBuilder struct {
	file   *excelize.File
	sheet  string
	config ChartConfig
	cell   string // Position where chart will be placed
}

// NewChartBuilder creates a new ChartBuilder instance
func NewChartBuilder(file *excelize.File, sheet string) *ChartBuilder {
	return &ChartBuilder{
		file:  file,
		sheet: sheet,
		config: ChartConfig{
			Width:  480,
			Height: 290,
			Legend: LegendConfig{
				Show:     true,
				Position: "bottom",
			},
		},
	}
}

// SetType sets the chart type
func (cb *ChartBuilder) SetType(chartType string) *ChartBuilder {
	cb.config.Type = chartType
	return cb
}

// SetTitle sets the chart title
func (cb *ChartBuilder) SetTitle(title string) *ChartBuilder {
	cb.config.Title = title
	return cb
}

// SetSize sets the chart dimensions
func (cb *ChartBuilder) SetSize(width, height int) *ChartBuilder {
	cb.config.Width = width
	cb.config.Height = height
	return cb
}

// SetPosition sets where the chart will be placed
func (cb *ChartBuilder) SetPosition(cell string) *ChartBuilder {
	cb.cell = cell
	return cb
}

// SetXAxis configures the X axis
func (cb *ChartBuilder) SetXAxis(config AxisConfig) *ChartBuilder {
	cb.config.XAxis = config
	return cb
}

// SetYAxis configures the Y axis
func (cb *ChartBuilder) SetYAxis(config AxisConfig) *ChartBuilder {
	cb.config.YAxis = config
	return cb
}

// SetLegend configures the chart legend
func (cb *ChartBuilder) SetLegend(config LegendConfig) *ChartBuilder {
	cb.config.Legend = config
	return cb
}

// AddDataSeries adds a data series to the chart
func (cb *ChartBuilder) AddDataSeries(series DataSeries) *ChartBuilder {
	cb.config.DataSeries = append(cb.config.DataSeries, series)
	return cb
}

// Build creates the chart in the Excel file
func (cb *ChartBuilder) Build() error {
	if cb.file == nil {
		return fmt.Errorf("file is nil")
	}

	if cb.config.Type == "" {
		return fmt.Errorf("chart type is required")
	}

	if len(cb.config.DataSeries) == 0 {
		return fmt.Errorf("at least one data series is required")
	}

	if cb.cell == "" {
		cb.cell = "A1" // Default position
	}

	// Convert our config to excelize chart format
	excelizeChart := cb.convertToExcelizeChart()

	// Add chart to the file
	err := cb.file.AddChart(cb.sheet, cb.cell, excelizeChart)
	if err != nil {
		return fmt.Errorf("failed to add chart: %w", err)
	}

	return nil
}

// convertToExcelizeChart converts our ChartConfig to excelize chart format
func (cb *ChartBuilder) convertToExcelizeChart() *excelize.Chart {
	chart := &excelize.Chart{
		Type: cb.mapChartType(cb.config.Type),
		Dimension: excelize.ChartDimension{
			Width:  uint(cb.config.Width),
			Height: uint(cb.config.Height),
		},
		Title: []excelize.RichTextRun{
			{
				Text: cb.config.Title,
			},
		},
		Legend: excelize.ChartLegend{
			Position:      cb.config.Legend.Position,
			ShowLegendKey: cb.config.Legend.Show,
		},
	}

	// Add data series
	for _, series := range cb.config.DataSeries {
		excelizeSeries := excelize.ChartSeries{
			Name:       series.Name,
			Categories: fmt.Sprintf("%s!%s", cb.sheet, series.Categories),
			Values:     fmt.Sprintf("%s!%s", cb.sheet, series.Values),
		}

		if series.Color != "" {
			excelizeSeries.Fill = excelize.Fill{
				Type:  "pattern",
				Color: []string{series.Color},
			}
		}

		chart.Series = append(chart.Series, excelizeSeries)
	}

	// Configure axes if specified
	if cb.config.XAxis.Title != "" {
		chart.XAxis.Title = []excelize.RichTextRun{
			{
				Text: cb.config.XAxis.Title,
			},
		}
	}

	if cb.config.YAxis.Title != "" {
		chart.YAxis.Title = []excelize.RichTextRun{
			{
				Text: cb.config.YAxis.Title,
			},
		}
	}

	return chart
}

// mapChartType maps our chart types to excelize chart types
func (cb *ChartBuilder) mapChartType(chartType string) excelize.ChartType {
	switch chartType {
	case "line":
		return excelize.Line
	case "bar":
		return excelize.Col
	case "pie":
		return excelize.Pie
	case "scatter":
		return excelize.Scatter
	case "area":
		return excelize.Area
	default:
		return excelize.Col // Default to column chart
	}
}

// GetConfig returns the current chart configuration
func (cb *ChartBuilder) GetConfig() ChartConfig {
	return cb.config
}