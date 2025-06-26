package excelbuilder

import (
	"github.com/xuri/excelize/v2"
)

// ChartBuilder handles chart creation and configuration
type ChartBuilder struct {
	file      *excelize.File
	sheetName string
	config    ChartConfig
	cell      string // Position where chart will be placed
}

// NewChartBuilder creates a new ChartBuilder instance
func NewChartBuilder(file *excelize.File, sheetName string) *ChartBuilder {
	return &ChartBuilder{
		file:      file,
		sheetName: sheetName,
		config: ChartConfig{
			Width:  480, // Default width
			Height: 290, // Default height
			Legend: LegendConfig{
				Position: "bottom",
				Show:     true,
			},
		},
	}
}

// SetType sets the chart type (e.g., "col", "bar", "pie")
func (cb *ChartBuilder) SetType(chartType string) *ChartBuilder {
	cb.config.Type = chartType
	return cb
}

// SetTitle sets the chart title
func (cb *ChartBuilder) SetTitle(title string) *ChartBuilder {
	cb.config.Title = title
	return cb
}

// SetDimensions sets the width and height of the chart.
func (cb *ChartBuilder) SetDimensions(width, height int) *ChartBuilder {
	cb.config.Width = width
	cb.config.Height = height
	return cb
}

// SetLegend configures the chart legend
func (cb *ChartBuilder) SetLegend(config LegendConfig) *ChartBuilder {
	cb.config.Legend = config
	return cb
}

// SetXAxis configures the X-axis of the chart
func (cb *ChartBuilder) SetXAxis(config AxisConfig) *ChartBuilder {
	cb.config.XAxis = config
	return cb
}

// SetYAxis configures the Y-axis of the chart
func (cb *ChartBuilder) SetYAxis(config AxisConfig) *ChartBuilder {
	cb.config.YAxis = config
	return cb
}

// AddDataSeries adds a data series to the chart
func (cb *ChartBuilder) AddDataSeries(series DataSeries) *ChartBuilder {
	cb.config.DataSeries = append(cb.config.DataSeries, series)
	return cb
}

// SetPosition sets the top-left cell where the chart will be placed
func (cb *ChartBuilder) SetPosition(cell string) *ChartBuilder {
	cb.cell = cell
	return cb
}

// Build creates the chart and adds it to the sheet.
func (cb *ChartBuilder) Build() error {
	var series []excelize.ChartSeries
	for _, s := range cb.config.DataSeries {
		series = append(series, excelize.ChartSeries{
			Name:       s.Name,
			Categories: s.Categories,
			Values:     s.Values,
			Fill:       excelize.Fill{Color: []string{s.Color}},
		})
	}

	chartOptions := &excelize.Chart{
		Type: mapChartType(cb.config.Type),
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
		Series: series,
	}

	return cb.file.AddChart(cb.sheetName, cb.cell, chartOptions)
}

// mapChartType converts a string representation of a chart type to the excelize constant.
func mapChartType(chartType string) excelize.ChartType {
	switch chartType {
	case "col":
		return excelize.Col
	case "bar":
		return excelize.Bar
	case "line":
		return excelize.Line
	case "pie":
		return excelize.Pie
	case "scatter":
		return excelize.Scatter
	default:
		return excelize.Col // Default to column chart
	}
}

// GetConfig returns the current chart configuration
func (cb *ChartBuilder) GetConfig() ChartConfig {
	return cb.config
}
