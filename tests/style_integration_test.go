package excelbuilder_test

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
)

func TestBuilder_EndToEnd_WithStyles(t *testing.T) {
	// 1. Define styles
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 14},
		Fill: excelbuilder.FillConfig{Type: "pattern", Color: "FFFF00"},
	}
	dataStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Size: 12},
	}

	// 2. Build the Excel file with styles
	builder := excelbuilder.New()
	file := builder.
		NewWorkbook().
		AddSheet("Styled Report").
		AddRow().
		AddCell("Product A").WithStyle(headerStyle).Done().
		AddCell("1000").WithStyle(headerStyle).Done().
		Done().
		AddRow().
		AddCell("Product B").WithStyle(dataStyle).Done().
		AddCell("2500").WithStyle(dataStyle).Done().
		Done().
		AddRow().
		// Reuse the header style for a summary row
		AddCell("Total").WithStyle(headerStyle).Done().
		AddCell("3500").WithStyle(headerStyle).Done().
		Done().
		Build()

	assert.NotNil(t, file)

	// 3. Verify the styles were applied correctly
	// Get the style ID for the first header cell
	headerStyleID1, err := file.GetCellStyle("Styled Report", "A1")
	assert.NoError(t, err)

	// Get the style ID for the second header cell
	headerStyleID2, err := file.GetCellStyle("Styled Report", "B1")
	assert.NoError(t, err)

	// Get the style ID for a data cell
	dataStyleID1, err := file.GetCellStyle("Styled Report", "A2")
	assert.NoError(t, err)

	// Get the style ID for the reused header style
	headerStyleID3, err := file.GetCellStyle("Styled Report", "A3")
	assert.NoError(t, err)

	// Assert that styles that should be the same are the same
	assert.Equal(t, headerStyleID1, headerStyleID2, "Header cells should share the same style ID")
	assert.Equal(t, headerStyleID1, headerStyleID3, "Reused header style should have the same ID as the original")

	// Assert that styles that should be different are different
	assert.NotEqual(t, headerStyleID1, dataStyleID1, "Header style and data style should have different IDs")
}
