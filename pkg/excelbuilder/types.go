package excelbuilder

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

// StyleConfig defines the styling configuration for cells
type StyleConfig struct {
	Font         FontConfig
	Fill         FillConfig
	Border       BorderConfig
	Alignment    AlignmentConfig
	NumberFormat string
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
	Title    string
	Min      *float64
	Max      *float64
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
	Type       string // "cellValue", "expression", "colorScale", "dataBar", "iconSet"
	Operator   string // "equal", "notEqual", "greaterThan", "lessThan", "between", "notBetween"
	Value      interface{}
	Value2     interface{}
	Format     StyleConfig
	ColorScale ColorScale
	IconSet    string
	Thresholds []Threshold
}

// ConditionalFormattingConfig defines conditional formatting rules
type ConditionalFormattingConfig struct {
	Range string
	Rules []ConditionalRule
}

// DataValidationConfig defines data validation rules
type DataValidationConfig struct {
	Range        string
	Type         string // "whole", "decimal", "list", "date", "time", "textLength", "custom"
	Operator     string // "between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"
	Formula      string
	Formula1     string
	Formula2     string
	ShowDropDown bool
	ShowDropdown bool
	ShowError    bool
	ShowInput    bool
	ErrorTitle   string
	ErrorMessage string
	InputTitle   string
	InputMessage string
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
