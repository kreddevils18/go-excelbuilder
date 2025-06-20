package excelbuilder

import (
	"crypto/md5"
	"encoding/json"
	"errors"
	"fmt"

	"github.com/xuri/excelize/v2"
)

// StyleFlyweight represents an immutable style object that can be shared
// Implements the Flyweight pattern for memory-efficient style management
type StyleFlyweight struct {
	config StyleConfig
	hash   string
	id     int
}

// NewStyleFlyweight creates a new StyleFlyweight instance
func NewStyleFlyweight(config StyleConfig, id int) *StyleFlyweight {
	// Create a deep copy to ensure immutability
	configCopy := copyStyleConfig(config)

	// Generate hash for the configuration
	hash := generateStyleHash(configCopy)

	return &StyleFlyweight{
		config: configCopy,
		hash:   hash,
		id:     id,
	}
}

// GetConfig returns a copy of the style configuration
// Returns a copy to maintain immutability
func (sf *StyleFlyweight) GetConfig() StyleConfig {
	return copyStyleConfig(sf.config)
}

// GetID returns the style ID.
func (sf *StyleFlyweight) GetID() int {
	return sf.id
}

// Apply applies the style to a given cell.
func (sf *StyleFlyweight) Apply(f *excelize.File, cellRef string) error {
	if f == nil {
		return errors.New("file cannot be nil")
	}

	if cellRef == "" {
		return errors.New("cell reference cannot be empty")
	}

	// Validate cell reference format
	if !isValidCellRef(cellRef) {
		return fmt.Errorf("invalid cell reference: %s", cellRef)
	}

	// If style ID is 0, create the style in the file first
	styleID := sf.id
	if styleID == 0 {
		// Convert StyleConfig to excelize.Style
		style := convertToExcelizeStyle(sf.config)
		
		// Create style in the file
		newStyleID, err := f.NewStyle(&style)
		if err != nil {
			return fmt.Errorf("failed to create style: %w", err)
		}
		styleID = newStyleID
		// Update the flyweight's ID for future use
		sf.id = newStyleID
	}

	// Apply style to cell
	err := f.SetCellStyle("Sheet1", cellRef, cellRef, styleID)
	if err != nil {
		return fmt.Errorf("failed to apply style to cell %s: %w", cellRef, err)
	}

	return nil
}

// Hash returns the hash of this flyweight
func (sf *StyleFlyweight) Hash() string {
	return sf.hash
}

// Equals checks if this flyweight is equal to another
func (sf *StyleFlyweight) Equals(other *StyleFlyweight) bool {
	if other == nil {
		return false
	}
	return sf.hash == other.hash
}

// Helper functions

// copyStyleConfig creates a deep copy of StyleConfig
func copyStyleConfig(config StyleConfig) StyleConfig {
	return StyleConfig{
		Font: FontConfig{
			Bold:      config.Font.Bold,
			Italic:    config.Font.Italic,
			Underline: config.Font.Underline,
			Size:      config.Font.Size,
			Color:     config.Font.Color,
			Family:    config.Font.Family,
		},
		Fill: FillConfig{
			Type:  config.Fill.Type,
			Color: config.Fill.Color,
		},
		Border: BorderConfig{
			Top:    config.Border.Top,
			Bottom: config.Border.Bottom,
			Left:   config.Border.Left,
			Right:  config.Border.Right,
			Color:  config.Border.Color,
		},
		Alignment: AlignmentConfig{
			Horizontal: config.Alignment.Horizontal,
			Vertical:   config.Alignment.Vertical,
		},
	}
}

// generateStyleHash generates a hash for the style configuration
func generateStyleHash(config StyleConfig) string {
	jsonData, err := json.Marshal(config)
	if err != nil {
		// Fallback to reflection-based hash if JSON fails
		return fmt.Sprintf("%x", md5.Sum([]byte(fmt.Sprintf("%+v", config))))
	}

	hash := md5.Sum(jsonData)
	return fmt.Sprintf("%x", hash)
}

// isValidCellRef validates if a cell reference is valid
func isValidCellRef(cellRef string) bool {
	if len(cellRef) < 2 {
		return false
	}

	// Simple validation - should start with letter(s) followed by number(s)
	hasLetter := false
	hasNumber := false

	for _, char := range cellRef {
		if char >= 'A' && char <= 'Z' {
			if hasNumber {
				return false // Letters after numbers
			}
			hasLetter = true
		} else if char >= '0' && char <= '9' {
			if !hasLetter {
				return false // Numbers before letters
			}
			hasNumber = true
		} else {
			return false // Invalid character
		}
	}

	return hasLetter && hasNumber
}