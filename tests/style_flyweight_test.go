package tests

import (
	"fmt"
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

// Test Case 9.1: StyleFlyweight Creation
func TestStyleFlyweight_Creation(t *testing.T) {
	// Test: Check StyleFlyweight instance creation
	// Expected:
	// - StyleFlyweight created successfully
	// - Immutable properties
	// - Contains style configuration

	styleConfig := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 14,
			Color: "#FF0000",
			Family: "Arial",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#FFFF00",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	}

	flyweight := excelbuilder.NewStyleFlyweight(styleConfig, 0)

	assert.NotNil(t, flyweight, "Expected StyleFlyweight instance, got nil")

	// Test that flyweight contains the configuration
	config := flyweight.GetConfig()
	assert.Equal(t, styleConfig.Font.Bold, config.Font.Bold)
	assert.Equal(t, styleConfig.Font.Size, config.Font.Size)
	assert.Equal(t, styleConfig.Font.Color, config.Font.Color)
	assert.Equal(t, styleConfig.Fill.Type, config.Fill.Type)
	assert.Equal(t, styleConfig.Fill.Color, config.Fill.Color)
}

// Test Case 9.2: StyleFlyweight Apply Method
func TestStyleFlyweight_Apply(t *testing.T) {
	// Test: Check that StyleFlyweight can apply styles to cells
	// Expected:
	// - Apply method works correctly
	// - Styles are applied to excelize file
	// - No errors during application

	styleConfig := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
			Color: "#000000",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#FFFFFF",
		},
	}

	flyweight := excelbuilder.NewStyleFlyweight(styleConfig, 0)
	file := excelize.NewFile()

	// Apply style to a cell
	err := flyweight.Apply(file, "A1")
	assert.NoError(t, err, "Expected no error when applying style")

	// Verify that style was applied (check if cell has style)
	styleID, err := file.GetCellStyle("Sheet1", "A1")
	assert.NoError(t, err, "Expected no error when getting cell style")
	assert.Greater(t, styleID, 0, "Expected cell to have a style applied")
}

// Test Case 9.3: StyleFlyweight Immutability
func TestStyleFlyweight_Immutability(t *testing.T) {
	// Test: Check that StyleFlyweight is immutable
	// Expected:
	// - Configuration cannot be modified after creation
	// - GetConfig returns copy, not reference
	// - Flyweight remains consistent

	originalConfig := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
		},
	}

	flyweight := excelbuilder.NewStyleFlyweight(originalConfig, 0)

	// Get config and try to modify it
	config := flyweight.GetConfig()
	config.Font.Bold = false
	config.Font.Size = 20

	// Original flyweight should remain unchanged
	originalConfigFromFlyweight := flyweight.GetConfig()
	assert.True(t, originalConfigFromFlyweight.Font.Bold, "Expected flyweight config to remain unchanged")
	assert.Equal(t, 12, originalConfigFromFlyweight.Font.Size, "Expected flyweight config to remain unchanged")
}

// Test Case 9.4: Complex Style Configuration
func TestStyleFlyweight_ComplexStyle(t *testing.T) {
	// Test: Check StyleFlyweight with complex style configuration
	// Expected:
	// - All style properties are preserved
	// - Complex styles apply correctly
	// - No data loss

	complexConfig := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:      true,
			Italic:    true,
			Underline: true,
			Size:      16,
			Color:     "#FF0000",
			Family:    "Times New Roman",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#FFFF00",
		},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin"},
			Bottom: excelbuilder.BorderSide{Style: "thick"},
			Left:   excelbuilder.BorderSide{Style: "medium"},
			Right:  excelbuilder.BorderSide{Style: "double"},
			Color:  "#000000",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
			Vertical:   "middle",
		},
	}

	flyweight := excelbuilder.NewStyleFlyweight(complexConfig, 0)
	config := flyweight.GetConfig()

	// Verify all properties are preserved
	assert.Equal(t, complexConfig.Font.Bold, config.Font.Bold)
	assert.Equal(t, complexConfig.Font.Italic, config.Font.Italic)
	assert.Equal(t, complexConfig.Font.Underline, config.Font.Underline)
	assert.Equal(t, complexConfig.Font.Size, config.Font.Size)
	assert.Equal(t, complexConfig.Font.Color, config.Font.Color)
	assert.Equal(t, complexConfig.Font.Family, config.Font.Family)
	assert.Equal(t, complexConfig.Fill.Type, config.Fill.Type)
	assert.Equal(t, complexConfig.Fill.Color, config.Fill.Color)
	assert.Equal(t, complexConfig.Border.Top, config.Border.Top)
	assert.Equal(t, complexConfig.Border.Bottom, config.Border.Bottom)
	assert.Equal(t, complexConfig.Border.Left, config.Border.Left)
	assert.Equal(t, complexConfig.Border.Right, config.Border.Right)
	assert.Equal(t, complexConfig.Border.Color, config.Border.Color)
	assert.Equal(t, complexConfig.Alignment.Horizontal, config.Alignment.Horizontal)
	assert.Equal(t, complexConfig.Alignment.Vertical, config.Alignment.Vertical)
}

// Test Case 9.5: StyleFlyweight Performance
func TestStyleFlyweight_Performance(t *testing.T) {
	// Test: Check StyleFlyweight performance with multiple applications
	// Expected:
	// - Fast application to multiple cells
	// - No performance degradation
	// - Efficient style reuse

	styleConfig := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
		},
	}

	flyweight := excelbuilder.NewStyleFlyweight(styleConfig, 0)
	file := excelize.NewFile()

	// Apply style to many cells
	const numCells = 100
	cells := make([]string, numCells)
	for i := 0; i < numCells; i++ {
		row := (i / 10) + 1
		col := (i % 10) + 1
		// Generate cell reference manually
		colLetter := string(rune('A' + col - 1))
		cells[i] = fmt.Sprintf("%s%d", colLetter, row)
	}

	// Apply style to all cells
	for _, cell := range cells {
		err := flyweight.Apply(file, cell)
		assert.NoError(t, err, "Expected no error when applying style to cell %s", cell)
	}

	// Verify all cells have styles applied
	for _, cell := range cells {
		styleID, err := file.GetCellStyle("Sheet1", cell)
		assert.NoError(t, err, "Expected no error when getting style for cell %s", cell)
		assert.Greater(t, styleID, 0, "Expected cell %s to have a style applied", cell)
	}
}

// Test Case 9.6: StyleFlyweight Error Handling
func TestStyleFlyweight_ErrorHandling(t *testing.T) {
	// Test: Check StyleFlyweight error handling
	// Expected:
	// - Handles invalid cell references gracefully
	// - Returns appropriate errors
	// - Doesn't crash on invalid input

	styleConfig := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
		},
	}

	flyweight := excelbuilder.NewStyleFlyweight(styleConfig, 0)
	file := excelize.NewFile()

	// Test with invalid cell reference
	err := flyweight.Apply(file, "INVALID_CELL")
	assert.Error(t, err, "Expected error when applying style to invalid cell reference")

	// Test with empty cell reference
	err = flyweight.Apply(file, "")
	assert.Error(t, err, "Expected error when applying style to empty cell reference")

	// Test with nil file
	err = flyweight.Apply(nil, "A1")
	assert.Error(t, err, "Expected error when applying style to nil file")
}

// Test Case 9.7: StyleFlyweight Hash and Equality
func TestStyleFlyweight_HashAndEquality(t *testing.T) {
	// Test: Check StyleFlyweight hash and equality methods
	// Expected:
	// - Same configs produce same hash
	// - Different configs produce different hash
	// - Equality works correctly

	styleConfig1 := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
		},
	}

	styleConfig2 := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
		},
	}

	styleConfig3 := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: false,
			Size: 14,
		},
	}

	flyweight1 := excelbuilder.NewStyleFlyweight(styleConfig1, 0)
	flyweight2 := excelbuilder.NewStyleFlyweight(styleConfig2, 0)
	flyweight3 := excelbuilder.NewStyleFlyweight(styleConfig3, 0)

	// Same configs should produce same hash
	hash1 := flyweight1.Hash()
	hash2 := flyweight2.Hash()
	hash3 := flyweight3.Hash()

	assert.Equal(t, hash1, hash2, "Expected same hash for identical configs")
	assert.NotEqual(t, hash1, hash3, "Expected different hash for different configs")

	// Test equality
	assert.True(t, flyweight1.Equals(flyweight2), "Expected flyweights with same config to be equal")
	assert.False(t, flyweight1.Equals(flyweight3), "Expected flyweights with different config to be unequal")
}