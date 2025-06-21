package excelbuilder

import "github.com/xuri/excelize/v2"

// WorkbookProperties defines the metadata for a workbook
type WorkbookProperties struct {
	Title       string
	Author      string
	Subject     string
	Description string
	Company     string
	Category    string
	Keywords    string
	Comments    string
}

// ProtectionConfig defines cell protection settings.
type ProtectionConfig struct {
	Locked bool
	Hidden bool
}

// StyleConfig defines the styling configuration for cells
type StyleConfig struct {
	Font         FontConfig
	Fill         FillConfig
	Border       BorderConfig
	Alignment    AlignmentConfig
	NumberFormat string
	Protection   *ProtectionConfig
}

// FontConfig defines font styling options
type FontConfig struct {
	Bold      bool
	Italic    bool
	Underline bool
	Size      int
	Color     string
	Family    string
}

// FillConfig defines cell fill/background options
type FillConfig struct {
	Type  string // "pattern", "gradient", etc.
	Color string
}

// BorderSide defines a single border side configuration
type BorderSide struct {
	Style string
	Color string
}

// BorderConfig defines cell border options
type BorderConfig struct {
	Top    BorderSide
	Bottom BorderSide
	Left   BorderSide
	Right  BorderSide
	Color  string // For backward compatibility
}

// AlignmentConfig defines cell alignment options
type AlignmentConfig struct {
	Horizontal   string // "left", "center", "right"
	Vertical     string // "top", "middle", "bottom"
	WrapText     bool
	TextRotation int
}

// ChartConfig defines chart configuration
type ChartConfig struct {
	Type       string // "line", "bar", "pie", "scatter", "area"
	Title      string
	Width      int
	Height     int
	XAxis      AxisConfig
	YAxis      AxisConfig
	Legend     LegendConfig
	DataSeries []DataSeries
}

// AxisConfig defines chart axis configuration
type AxisConfig struct {
	Title     string
	Min       *float64
	Max       *float64
	MajorUnit *float64
	MinorUnit *float64
}

// LegendConfig defines chart legend configuration
type LegendConfig struct {
	Show     bool
	Position string // "top", "bottom", "left", "right"
}

// DataSeries defines a data series for charts
type DataSeries struct {
	Name       string
	Categories string // Cell range for categories (e.g., "A1:A10")
	Values     string // Cell range for values (e.g., "B1:B10")
	Color      string
}

// ColorScale defines color scale configuration for conditional formatting
type ColorScale struct {
	MinColor string
	MidColor string
	MaxColor string
}

// Threshold defines threshold configuration for icon sets
type Threshold struct {
	Value string
	Type  string // "percent", "number", "formula"
}

// ConditionalRule defines a single conditional formatting rule
type ConditionalRule struct {
	Type       string // "cell", "average", "duplicateValues", "uniqueValues", "top", "bottom", "blanks", "noBlanks", "errors", "noErrors", "timePeriod", "2_color_scale", "3_color_scale", "data_bar", "icon_set"
	Operator   string // "equal", "notEqual", "greaterThan", "lessThan", "between", "notBetween"
	Value      interface{}
	Style      StyleConfig
	ColorScale struct {
		MinType  string
		MinValue string
		MinColor string
		MidType  string
		MidValue string
		MidColor string
		MaxType  string
		MaxValue string
		MaxColor string
	}
	DataBar struct {
		Color       string
		MinLength   uint8
		MaxLength   uint8
		BorderColor string
		Direction   string
		BarOnly     bool
	}
	IconSet struct {
		Style     string
		Reverse   bool
		IconsOnly bool
	}
}

// ConditionalFormattingConfig defines conditional formatting rules
type ConditionalFormattingConfig struct {
	Range string
	Rules []ConditionalRule
}

// DataValidation defines data validation rules for cells
type DataValidation struct {
	Type             string // "whole", "decimal", "list", "date", "time", "textLength", "custom"
	Operator         string // "between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"
	Formula1         string
	Formula2         string
	AllowBlank       bool
	ShowDropDown     bool
	ShowInputMessage bool
	ShowErrorMessage bool
	ErrorTitle       string
	ErrorMessage     string
	ErrorStyle       string // "stop", "warning", "information"
	PromptTitle      string
	PromptMessage    string
}

// DataValidationConfig defines data validation rules
type DataValidationConfig struct {
	Type             string // "whole", "decimal", "list", "date", "time", "textLength", "custom"
	Operator         string // "between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"
	AllowBlank       bool
	ShowInputMessage bool
	ShowErrorMessage bool
	ErrorTitle       string
	ErrorBody        string
	ErrorStyle       string // "stop", "warning", "information"
	PromptTitle      string
	PromptBody       string
	Formula1         []string
	Formula2         []string
}

// SheetTemplate defines a sheet template configuration
type SheetTemplate struct {
	Name      string
	Headers   []string
	Styles    map[string]StyleConfig
	Variables map[string]interface{}
}

// TemplateConfig defines template configuration
type TemplateConfig struct {
	Name        string
	Description string
	Path        string
	Sheets      []SheetTemplate
	Variables   map[string]interface{}
}

// FormulaConfig defines advanced formula configuration
type FormulaConfig struct {
	Expression string
	IsArray    bool
	Range      string
}

// BatchRowData defines a row of data for batch insertion
type BatchRowData struct {
	Cells []interface{}
	Style StyleConfig
}

// BatchStyleOperation defines a batch styling operation
type BatchStyleOperation struct {
	Range string
	Style StyleConfig
}

// SheetConfig defines the configuration for a sheet in batch creation
type SheetConfig struct {
	Name string
	Data [][]interface{}
}

// SheetProtectionConfig defines the protection settings for a worksheet.
type SheetProtectionConfig struct {
	Password            string
	SelectLockedCells   bool
	SelectUnlockedCells bool
	FormatCells         bool
	FormatColumns       bool
	FormatRows          bool
	InsertColumns       bool
	InsertRows          bool
	InsertHyperlinks    bool
	DeleteColumns       bool
	DeleteRows          bool
	Sort                bool
	AutoFilter          bool
	PivotTables         bool
	EditObjects         bool
	EditScenarios       bool
}

// RGBColor defines a color using RGB values.
type RGBColor struct {
	R int
	G int
	B int
}

// ThemeColor defines a theme color with an optional tint.
type ThemeColor struct {
	Theme string
	Tint  float64
}

// Theme color constants (string values for compatibility)
const (
	ThemeColorDark1             = "dark1"
	ThemeColorLight1            = "light1"
	ThemeColorDark2             = "dark2"
	ThemeColorLight2            = "light2"
	ThemeColorAccent1           = "accent1"
	ThemeColorAccent2           = "accent2"
	ThemeColorAccent3           = "accent3"
	ThemeColorAccent4           = "accent4"
	ThemeColorAccent5           = "accent5"
	ThemeColorAccent6           = "accent6"
	ThemeColorHyperlink         = "hyperlink"
	ThemeColorFollowedHyperlink = "followedHyperlink"
)

// Theme color index constants (integer values for excelize API)
const (
	ThemeColorIndexDark1             = 1
	ThemeColorIndexLight1            = 2
	ThemeColorIndexDark2             = 3
	ThemeColorIndexLight2            = 4
	ThemeColorIndexAccent1           = 5
	ThemeColorIndexAccent2           = 6
	ThemeColorIndexAccent3           = 7
	ThemeColorIndexAccent4           = 8
	ThemeColorIndexAccent5           = 9
	ThemeColorIndexAccent6           = 10
	ThemeColorIndexHyperlink         = 11
	ThemeColorIndexFollowedHyperlink = 12
)

// Predefined standard colors
const (
	ColorBlack  = "#000000"
	ColorWhite  = "#FFFFFF"
	ColorRed    = "#FF0000"
	ColorGreen  = "#00FF00"
	ColorBlue   = "#0000FF"
	ColorYellow = "#FFFF00"
	ColorOrange = "#FFA500"
	ColorPurple = "#800080"
)

// Import/Export types

// CSVOptions defines options for CSV import/export
type CSVOptions struct {
	Delimiter string
	Quote     rune
	Comment   rune
	SkipRows  int
}

// FlattenOptions defines options for flattening nested JSON structures
type FlattenOptions struct {
	Separator string
	MaxDepth  int
}

// ImportHelper provides functionality to import data from various formats
type ImportHelper struct {
	data      interface{}
	sheetName string
}

// ExportHelper provides functionality to export Excel data to various formats
type ExportHelper struct {
	file *excelize.File
}

// Pivot Table types

// PivotField defines a field in a pivot table
type PivotField struct {
	Name     string
	Function string // For value fields: "sum", "count", "average", "max", "min", "product", "countNums", "stdDev", "stdDevp", "var", "varp"
}

// PivotTableConfig defines the configuration for a pivot table
type PivotTableConfig struct {
	Name                  string
	SourceSheet          string
	SourceRange          string
	TargetSheet          string
	TargetCell           string
	RowFields            []PivotField
	ColumnFields         []PivotField
	ValueFields          []PivotField
	FilterFields         []PivotField
	Style                string
	ShowRowGrandTotals   bool
	ShowColumnGrandTotals bool
	Compact              bool
	Outline              bool
	Subtotals            bool
}

// Advanced Layout Management types

// GroupingConfig defines configuration for column/row grouping
type GroupingConfig struct {
	Level     int
	Collapsed bool
	Summary   bool
}

// PaneConfig defines configuration for freeze/split panes
type PaneConfig struct {
	Type         string // "freeze" or "split"
	TopLeftCell  string // For freeze panes
	XSplit       int    // For split panes - horizontal split position
	YSplit       int    // For split panes - vertical split position
	ActivePane   string // Which pane is active
}

// LayoutRange defines a range for layout operations
type LayoutRange struct {
	StartColumn string
	EndColumn   string
	StartRow    int
	EndRow      int
}

// AutoFitConfig defines configuration for auto-fitting columns
type AutoFitConfig struct {
	MinWidth    float64
	MaxWidth    float64
	Padding     float64
	IgnoreEmpty bool
}
