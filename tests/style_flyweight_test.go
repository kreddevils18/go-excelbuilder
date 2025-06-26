package excelbuilder_test

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

func TestStyleFlyweight_Apply(t *testing.T) {
	file := excelize.NewFile()
	// Create a dummy sheet because excelize requires it to apply styles
	_, err := file.NewSheet("Sheet1")
	assert.NoError(t, err)

	config := excelbuilder.StyleConfig{Font: excelbuilder.FontConfig{Bold: true}}

	// Create a style directly to get a valid ID
	styleID, err := file.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true}})
	assert.NoError(t, err)

	flyweight := excelbuilder.NewStyleFlyweight(config, styleID)

	// Apply the style
	err = flyweight.Apply(file, "Sheet1", "A1")
	assert.NoError(t, err)

	// Verify the style was applied
	appliedStyleID, err := file.GetCellStyle("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, styleID, appliedStyleID)
}

func TestStyleFlyweight_Apply_LazyCreation(t *testing.T) {
	file := excelize.NewFile()
	_, err := file.NewSheet("Sheet1")
	assert.NoError(t, err)

	config := excelbuilder.StyleConfig{Font: excelbuilder.FontConfig{Italic: true}}

	// Create flyweight with ID 0 to trigger lazy creation
	flyweight := excelbuilder.NewStyleFlyweight(config, 0)
	assert.Equal(t, 0, flyweight.GetID(), "Initial style ID should be 0")

	// Apply the style, which should create it
	err = flyweight.Apply(file, "Sheet1", "B2")
	assert.NoError(t, err)

	// Verify the flyweight now has a valid ID
	newID := flyweight.GetID()
	assert.NotEqual(t, 0, newID, "Style ID should be updated after lazy creation")

	// Verify the style was applied correctly in the file
	appliedStyleID, err := file.GetCellStyle("Sheet1", "B2")
	assert.NoError(t, err)
	assert.Equal(t, newID, appliedStyleID)
}

func TestStyleFlyweight_Immutability(t *testing.T) {
	config := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Size: 12},
	}
	flyweight := excelbuilder.NewStyleFlyweight(config, 1)

	// Get the config and try to modify it
	retrievedConfig := flyweight.GetConfig()
	retrievedConfig.Font.Size = 14

	// Get the config again and assert it hasn't changed
	finalConfig := flyweight.GetConfig()
	assert.Equal(t, 12, finalConfig.Font.Size, "Internal flyweight config should be immutable")
}

func TestStyleFlyweight_Equals(t *testing.T) {
	config1 := excelbuilder.StyleConfig{Font: excelbuilder.FontConfig{Bold: true}}
	config2 := excelbuilder.StyleConfig{Font: excelbuilder.FontConfig{Bold: true}}
	config3 := excelbuilder.StyleConfig{Font: excelbuilder.FontConfig{Bold: false}}

	flyweight1 := excelbuilder.NewStyleFlyweight(config1, 1)
	flyweight2 := excelbuilder.NewStyleFlyweight(config2, 2) // Different ID, same config
	flyweight3 := excelbuilder.NewStyleFlyweight(config3, 3)

	assert.True(t, flyweight1.Equals(flyweight2), "Flyweights with same config should be equal")
	assert.False(t, flyweight1.Equals(flyweight3), "Flyweights with different config should not be equal")
	assert.False(t, flyweight1.Equals(nil), "Flyweight should not be equal to nil")
}
