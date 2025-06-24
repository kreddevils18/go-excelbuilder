# Advanced Styling Example

This example demonstrates comprehensive styling capabilities of `go-excelbuilder`, including advanced typography, color schemes, border patterns, and professional report layouts.

## Learning Objectives

- **Advanced Typography**: Font families, sizes, weights, and text styling
- **Color Theory**: Professional color palettes and consistent theming
- **Border Patterns**: Various border styles and their appropriate usage
- **Fill Techniques**: Background colors, patterns, and visual hierarchy
- **Conditional Styling**: Simulating conditional formatting with dynamic styles
- **Professional Layouts**: Creating publication-ready reports and dashboards
- **Visual Hierarchy**: Using styling to guide reader attention
- **Consistency**: Maintaining design standards across multiple sheets

## What This Example Creates

The example generates an Excel file with six comprehensive sheets:

### 1. Typography Sheet
- Font style demonstrations (normal, bold, italic, large headers, small text)
- Color variations and their use cases
- Size hierarchy examples
- Professional typography best practices

### 2. Color Schemes Sheet
- Complete color palette with hex codes
- Text color examples
- Background color samples
- Use case recommendations for each color

### 3. Borders & Fills Sheet
- Border style variations (thin, medium, thick, double)
- Fill pattern demonstrations
- Practical usage examples
- Visual impact comparisons

### 4. Data Visualization Sheet
- Sales performance data with conditional styling
- Value-based color coding
- Performance indicators
- Summary statistics with special formatting

### 5. Conditional Styling Sheet
- Employee performance simulation
- Score-based styling rules
- Grade assignments with visual indicators
- Legend and documentation

### 6. Professional Report Sheet
- Complete quarterly sales report
- Corporate header and metadata
- Regional performance analysis
- Key insights section
- Professional footer

## How to Run

```bash
cd examples/03-styling-advanced
go run main.go
```

The example will create `output/03-styling-advanced-demo.xlsx` with all styling demonstrations.

## Key Code Patterns

### Color Palette Management

```go
// Define consistent color palette
colorPalette := map[string]string{
    "primary":   "#2F5597",
    "secondary": "#4472C4",
    "accent":    "#70AD47",
    "warning":   "#FFC000",
    "danger":    "#C5504B",
    "success":   "#70AD47",
}
```

### Advanced Style Configuration

```go
// Create comprehensive style with all properties
style := excelbuilder.StyleConfig{
    Font: excelbuilder.FontConfig{
        Bold:   true,
        Size:   12,
        Color:  colors["white"],
        Family: "Arial",
    },
    Fill: excelbuilder.FillConfig{
        Type:  "pattern",
        Color: colors["primary"],
    },
    Alignment: excelbuilder.AlignmentConfig{
        Horizontal: "center",
        Vertical:   "middle",
    },
    Border: createBorder("thin", colors["dark"]),
    NumberFormat: "$#,##0.00",
}
```

### Conditional Styling Logic

```go
// Apply styles based on data values
func getValueStyle(value float64, styles map[string]excelbuilder.StyleConfig) excelbuilder.StyleConfig {
    if value > 40000 {
        return styles["high_value"]
    } else if value > 25000 {
        return styles["medium_value"]
    }
    return styles["low_value"]
}
```

### Border Creation Helper

```go
// Reusable border creation function
func createBorder(style, color string) excelbuilder.BorderConfig {
    return excelbuilder.BorderConfig{
        Top:    excelbuilder.BorderSide{Style: style, Color: color},
        Bottom: excelbuilder.BorderSide{Style: style, Color: color},
        Left:   excelbuilder.BorderSide{Style: style, Color: color},
        Right:  excelbuilder.BorderSide{Style: style, Color: color},
    }
}
```

## Styling Best Practices Demonstrated

### 1. Color Consistency
- Use a defined color palette throughout
- Maintain contrast ratios for readability
- Apply semantic colors (success = green, danger = red)

### 2. Typography Hierarchy
- Use size and weight to create visual hierarchy
- Maintain consistent font families
- Apply appropriate text colors for different contexts

### 3. Border Usage
- Thin borders for data separation
- Medium borders for section emphasis
- Thick borders for major divisions
- Double borders for totals and summaries

### 4. Fill Patterns
- Solid fills for headers and emphasis
- Light fills for alternating rows
- Colored fills for status indicators
- Consistent fill usage across sheets

### 5. Professional Layout
- Clear section divisions
- Consistent spacing and alignment
- Appropriate use of white space
- Logical information flow

## Advanced Features Showcased

### Dynamic Style Application
- Value-based conditional styling
- Performance-driven color coding
- Automatic grade assignment
- Status indicator generation

### Professional Report Elements
- Corporate headers with metadata
- Summary statistics sections
- Key insights documentation
- Professional footers

### Visual Data Communication
- Color-coded performance metrics
- Hierarchical information presentation
- Clear data categorization
- Intuitive visual cues

## Related Examples

- **Previous**: `02-data-types/` - Data type handling and basic formatting
- **Next**: `04-sales-report/` - Real-world business report implementation
- **See Also**: `05-import-export/` - Data import/export with styling preservation

## Key Concepts Covered

### Styling Architecture
- Centralized style management
- Reusable style configurations
- Consistent theming approach
- Modular style functions

### Visual Design Principles
- Color theory application
- Typography best practices
- Layout and spacing
- Visual hierarchy creation

### Data Presentation
- Conditional formatting simulation
- Performance visualization
- Status indication systems
- Professional report layouts

## Styling Features Used

### Font Properties
- **Bold/Italic**: Text emphasis and hierarchy
- **Size**: Visual importance and readability
- **Color**: Semantic meaning and branding
- **Family**: Professional appearance

### Fill Properties
- **Pattern**: Solid background colors
- **Color**: Brand colors and status indicators
- **Type**: Consistent fill application

### Border Properties
- **Style**: Visual separation and emphasis
- **Color**: Subtle or prominent borders
- **Sides**: Complete or selective borders

### Alignment Properties
- **Horizontal**: Text positioning and layout
- **Vertical**: Cell content alignment
- **Consistency**: Professional appearance

### Number Formatting
- **Currency**: Financial data presentation
- **Percentage**: Ratio and performance metrics
- **Custom**: Specialized formatting needs

## Next Steps

1. **Experiment** with different color combinations
2. **Modify** the conditional styling logic
3. **Add** new styling patterns and themes
4. **Create** your own professional report layouts
5. **Explore** the sales report example for real-world application

This example provides a comprehensive foundation for creating visually appealing and professionally styled Excel documents using `go-excelbuilder`.