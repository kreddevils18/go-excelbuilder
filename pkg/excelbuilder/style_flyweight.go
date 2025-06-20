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
}

// NewStyleFlyweight creates a new StyleFlyweight instance
func NewStyleFlyweight(config StyleConfig) *StyleFlyweight {
	// Create a deep copy to ensure immutability
	configCopy := copyStyleConfig(config)
	
	// Generate hash for the configuration
	hash := generateStyleHash(configCopy)
	
	return &StyleFlyweight{
		config: configCopy,
		hash:   hash,
	}
}

// GetConfig returns a copy of the style configuration
// Returns a copy to maintain immutability
func (sf *StyleFlyweight) GetConfig() StyleConfig {
	return copyStyleConfig(sf.config)
}

// Apply applies the style to a cell in the given Excel file
func (sf *StyleFlyweight) Apply(file *excelize.File, cellRef string) error {
	if file == nil {
		return errors.New("file cannot be nil")
	}
	
	if cellRef == "" {
		return errors.New("cell reference cannot be empty")
	}
	
	// Validate cell reference format
	if !isValidCellRef(cellRef) {
		return fmt.Errorf("invalid cell reference: %s", cellRef)
	}
	
	// Convert StyleConfig to excelize.Style
	excelizeStyle := sf.convertToExcelizeStyle()
	
	// Create style in excelize
	styleID, err := file.NewStyle(&excelizeStyle)
	if err != nil {
		return fmt.Errorf("failed to create style: %w", err)
	}
	
	// Apply style to cell
	err = file.SetCellStyle("Sheet1", cellRef, cellRef, styleID)
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

// convertToExcelizeStyle converts StyleConfig to excelize.Style
func (sf *StyleFlyweight) convertToExcelizeStyle() excelize.Style {
	style := excelize.Style{}
	
	// Font configuration
	if sf.config.Font.Size > 0 || sf.config.Font.Bold || sf.config.Font.Italic || 
	   sf.config.Font.Underline || sf.config.Font.Color != "" || sf.config.Font.Family != "" {
		font := &excelize.Font{}
		
		if sf.config.Font.Size > 0 {
			font.Size = float64(sf.config.Font.Size)
		}
		if sf.config.Font.Bold {
			font.Bold = true
		}
		if sf.config.Font.Italic {
			font.Italic = true
		}
		if sf.config.Font.Underline {
			font.Underline = "single"
		}
		if sf.config.Font.Color != "" {
			font.Color = sf.config.Font.Color
		}
		if sf.config.Font.Family != "" {
			font.Family = sf.config.Font.Family
		}
		
		style.Font = font
	}
	
	// Fill configuration
	if sf.config.Fill.Type != "" || sf.config.Fill.Color != "" {
		fill := &excelize.Fill{}
		
		if sf.config.Fill.Type == "pattern" && sf.config.Fill.Color != "" {
			fill.Type = "pattern"
			fill.Pattern = 1 // Solid fill
			fill.Color = []string{sf.config.Fill.Color}
		}
		
		style.Fill = *fill
	}
	
	// Border configuration
	if sf.config.Border.Top.Style != "" || sf.config.Border.Bottom.Style != "" ||
	   sf.config.Border.Left.Style != "" || sf.config.Border.Right.Style != "" {
		border := []excelize.Border{}
		
		if sf.config.Border.Top.Style != "" {
			border = append(border, excelize.Border{
				Type:  "top",
				Style: getBorderStyle(sf.config.Border.Top.Style),
				Color: getColorOrDefault(sf.config.Border.Top.Color, sf.config.Border.Color),
			})
		}
		if sf.config.Border.Bottom.Style != "" {
			border = append(border, excelize.Border{
				Type:  "bottom",
				Style: getBorderStyle(sf.config.Border.Bottom.Style),
				Color: getColorOrDefault(sf.config.Border.Bottom.Color, sf.config.Border.Color),
			})
		}
		if sf.config.Border.Left.Style != "" {
			border = append(border, excelize.Border{
				Type:  "left",
				Style: getBorderStyle(sf.config.Border.Left.Style),
				Color: getColorOrDefault(sf.config.Border.Left.Color, sf.config.Border.Color),
			})
		}
		if sf.config.Border.Right.Style != "" {
			border = append(border, excelize.Border{
				Type:  "right",
				Style: getBorderStyle(sf.config.Border.Right.Style),
				Color: getColorOrDefault(sf.config.Border.Right.Color, sf.config.Border.Color),
			})
		}
		
		style.Border = border
	}
	
	// Alignment configuration
	if sf.config.Alignment.Horizontal != "" || sf.config.Alignment.Vertical != "" {
		alignment := &excelize.Alignment{}
		
		if sf.config.Alignment.Horizontal != "" {
			alignment.Horizontal = sf.config.Alignment.Horizontal
		}
		if sf.config.Alignment.Vertical != "" {
			alignment.Vertical = sf.config.Alignment.Vertical
		}
		
		style.Alignment = alignment
	}
	
	return style
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

// getColorOrDefault returns the primary color if not empty, otherwise returns the default color
func getColorOrDefault(primary, defaultColor string) string {
	if primary != "" {
		return primary
	}
	return defaultColor
}