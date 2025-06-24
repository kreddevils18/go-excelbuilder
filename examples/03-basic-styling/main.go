package main

import (
	"fmt"
	"log"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

func main() {
	fmt.Println("Creating Basic Styling Example...")

	// Create a new workbook
	builder := excelbuilder.New()
	wb := builder.NewWorkbook()

	// Create different sheets to demonstrate various styling features
	createFontStylesSheet(wb)
	createColorsSheet(wb)
	createBordersSheet(wb)
	createAlignmentSheet(wb)
	createNumberFormatsSheet(wb)

	// Build and save the workbook
	file := wb.Build()
	if file == nil {
		log.Fatal("Failed to build workbook")
	}

	filename := "03-basic-styling.xlsx"
	if err := file.SaveAs(filename); err != nil {
		log.Fatalf("Failed to save workbook: %v", err)
	}

	fmt.Printf("Basic styling example saved as %s\n", filename)
}

// createFontStylesSheet demonstrates various font styling options
func createFontStylesSheet(wb *excelbuilder.WorkbookBuilder) {
	sheet := wb.AddSheet("Font Styles")

	// Title
	titleStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   16,
			Bold:   true,
			Family: "Arial",
			Color:  "#2E75B6",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
		},
	}
	sheet.AddRow().AddCell("Font Styling Examples").SetStyle(titleStyle)

	// Empty row
	sheet.AddRow()

	// Headers
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Family: "Calibri",
			Size:   11,
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#D9E2F3",
		},
		Border: excelbuilder.BorderConfig{
			Top:    excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
			Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
			Left:   excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
			Right:  excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
		},
	}

	headers := []string{"Font Family", "Size", "Style", "Color", "Example Text", "Description"}
	headerRow := sheet.AddRow()
	for _, header := range headers {
		headerRow.AddCell(header).SetStyle(headerStyle)
	}

	// Font examples
	fontExamples := []struct {
		family      string
		size        int
		style       string
		color       string
		example     string
		description string
		fontConfig  excelbuilder.FontConfig
	}{
		{"Arial", 12, "Regular", "Black", "Sample Text", "Standard Arial font", excelbuilder.FontConfig{Family: "Arial", Size: 12}},
		{"Arial", 14, "Bold", "Blue", "Bold Text", "Bold Arial in blue", excelbuilder.FontConfig{Family: "Arial", Size: 14, Bold: true, Color: "#0070C0"}},
		{"Calibri", 11, "Italic", "Green", "Italic Text", "Italic Calibri in green", excelbuilder.FontConfig{Family: "Calibri", Size: 11, Italic: true, Color: "#70AD47"}},
		{"Times New Roman", 13, "Underline", "Red", "Underlined Text", "Underlined Times New Roman", excelbuilder.FontConfig{Family: "Times New Roman", Size: 13, Underline: true, Color: "#C5504B"}},
		{"Verdana", 10, "Bold+Italic", "Purple", "Bold Italic", "Combined bold and italic", excelbuilder.FontConfig{Family: "Verdana", Size: 10, Bold: true, Italic: true, Color: "#7030A0"}},
		{"Georgia", 15, "Large", "Orange", "Large Text", "Large Georgia font", excelbuilder.FontConfig{Family: "Georgia", Size: 15, Color: "#FF8C00"}},
	}

	for _, example := range fontExamples {
		row := sheet.AddRow()

		// Font family
		row.AddCell(example.family)

		// Size
		row.AddCell(example.size)

		// Style
		row.AddCell(example.style)

		// Color
		row.AddCell(example.color)

		// Example text with styling
		exampleStyle := excelbuilder.StyleConfig{
			Font: example.fontConfig,
		}
		row.AddCell(example.example).SetStyle(exampleStyle)

		// Description
		row.AddCell(example.description)
	}

	// Set column widths
	sheet.SetColumnWidth("A", 15)
	sheet.SetColumnWidth("B", 8)
	sheet.SetColumnWidth("C", 12)
	sheet.SetColumnWidth("D", 10)
	sheet.SetColumnWidth("E", 15)
	sheet.SetColumnWidth("F", 25)
}

// createColorsSheet demonstrates background and text colors
func createColorsSheet(wb *excelbuilder.WorkbookBuilder) {
	sheet := wb.AddSheet("Colors")

	// Title
	titleStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   16,
			Bold:   true,
			Family: "Arial",
			Color:  "#FFFFFF",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#2E75B6",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
		},
	}
	sheet.AddRow().AddCell("Color Examples").SetStyle(titleStyle)

	// Empty row
	sheet.AddRow()

	// Color palette examples
	colors := []struct {
		name        string
		hex         string
		textColor   string
		description string
	}{
		{"Primary Blue", "#2E75B6", "#FFFFFF", "Corporate primary color"},
		{"Light Blue", "#D9E2F3", "#000000", "Light background color"},
		{"Success Green", "#70AD47", "#FFFFFF", "Success indicator"},
		{"Warning Orange", "#FF8C00", "#FFFFFF", "Warning indicator"},
		{"Danger Red", "#C5504B", "#FFFFFF", "Error indicator"},
		{"Purple", "#7030A0", "#FFFFFF", "Accent color"},
		{"Gray", "#808080", "#FFFFFF", "Neutral color"},
		{"Light Gray", "#F2F2F2", "#000000", "Background color"},
	}

	// Headers
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#E7E6E6",
		},
	}

	headers := []string{"Color Name", "Hex Code", "Sample", "Description"}
	headerRow := sheet.AddRow()
	for _, header := range headers {
		headerRow.AddCell(header).SetStyle(headerStyle)
	}

	// Color examples
	for _, color := range colors {
		row := sheet.AddRow()

		// Color name
		row.AddCell(color.name)

		// Hex code
		row.AddCell(color.hex)

		// Sample with background color
		sampleStyle := excelbuilder.StyleConfig{
			Font: excelbuilder.FontConfig{
				Color: color.textColor,
				Bold:  true,
			},
			Fill: excelbuilder.FillConfig{
				Type:  "pattern",
				Color: color.hex,
			},
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: "center",
			},
		}
		row.AddCell("Sample Text").SetStyle(sampleStyle)

		// Description
		row.AddCell(color.description)
	}

	// Set column widths
	sheet.SetColumnWidth("A", 15)
	sheet.SetColumnWidth("B", 12)
	sheet.SetColumnWidth("C", 15)
	sheet.SetColumnWidth("D", 25)
}

// createBordersSheet demonstrates various border styles
func createBordersSheet(wb *excelbuilder.WorkbookBuilder) {
	sheet := wb.AddSheet("Borders")

	// Title
	titleStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   16,
			Bold:   true,
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
		},
	}
	sheet.AddRow().AddCell("Border Styles").SetStyle(titleStyle)

	// Empty row
	sheet.AddRow()

	// Border examples
	borderExamples := []struct {
		name        string
		description string
		border      excelbuilder.BorderConfig
	}{
		{
			"Thin Border",
			"Standard thin border all around",
			excelbuilder.BorderConfig{
				Top:    excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
				Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
				Left:   excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
				Right:  excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
			},
		},
		{
			"Thick Border",
			"Thick border for emphasis",
			excelbuilder.BorderConfig{
				Top:    excelbuilder.BorderSide{Style: "thick", Color: "#2E75B6"},
				Bottom: excelbuilder.BorderSide{Style: "thick", Color: "#2E75B6"},
				Left:   excelbuilder.BorderSide{Style: "thick", Color: "#2E75B6"},
				Right:  excelbuilder.BorderSide{Style: "thick", Color: "#2E75B6"},
			},
		},
		{
			"Dashed Border",
			"Dashed border style",
			excelbuilder.BorderConfig{
				Top:    excelbuilder.BorderSide{Style: "dashed", Color: "#70AD47"},
				Bottom: excelbuilder.BorderSide{Style: "dashed", Color: "#70AD47"},
				Left:   excelbuilder.BorderSide{Style: "dashed", Color: "#70AD47"},
				Right:  excelbuilder.BorderSide{Style: "dashed", Color: "#70AD47"},
			},
		},
		{
			"Double Border",
			"Double line border",
			excelbuilder.BorderConfig{
				Top:    excelbuilder.BorderSide{Style: "double", Color: "#C5504B"},
				Bottom: excelbuilder.BorderSide{Style: "double", Color: "#C5504B"},
				Left:   excelbuilder.BorderSide{Style: "double", Color: "#C5504B"},
				Right:  excelbuilder.BorderSide{Style: "double", Color: "#C5504B"},
			},
		},
		{
			"Bottom Only",
			"Border only at bottom",
			excelbuilder.BorderConfig{
				Bottom: excelbuilder.BorderSide{Style: "thick", Color: "#FF8C00"},
			},
		},
	}

	// Headers
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#D9E2F3",
		},
	}

	headers := []string{"Style Name", "Example", "Description"}
	headerRow := sheet.AddRow()
	for _, header := range headers {
		headerRow.AddCell(header).SetStyle(headerStyle)
	}

	// Border examples
	for _, example := range borderExamples {
		row := sheet.AddRow()

		// Style name
		row.AddCell(example.name)

		// Example with border
		exampleStyle := excelbuilder.StyleConfig{
			Border: example.border,
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: "center",
			},
		}
		row.AddCell("Sample").SetStyle(exampleStyle)

		// Description
		row.AddCell(example.description)
	}

	// Set column widths
	sheet.SetColumnWidth("A", 15)
	sheet.SetColumnWidth("B", 15)
	sheet.SetColumnWidth("C", 30)
}

// createAlignmentSheet demonstrates text alignment options
func createAlignmentSheet(wb *excelbuilder.WorkbookBuilder) {
	sheet := wb.AddSheet("Alignment")

	// Title
	titleStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   16,
			Bold:   true,
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
		},
	}
	sheet.AddRow().AddCell("Text Alignment Examples").SetStyle(titleStyle)

	// Empty row
	sheet.AddRow()

	// Alignment examples
	alignmentExamples := []struct {
		name        string
		horizontal  string
		vertical    string
		description string
	}{
		{"Left Top", "left", "top", "Text aligned to left and top"},
		{"Center Middle", "center", "center", "Text centered both ways"},
		{"Right Bottom", "right", "bottom", "Text aligned to right and bottom"},
		{"Center Top", "center", "top", "Centered horizontally, top vertically"},
		{"Left Center", "left", "center", "Left horizontally, centered vertically"},
		{"Justify", "justify", "center", "Justified text alignment"},
	}

	// Headers
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#D9E2F3",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
		},
	}

	headers := []string{"Alignment Type", "Example", "H-Align", "V-Align"}
	headerRow := sheet.AddRow()
	for _, header := range headers {
		headerRow.AddCell(header).SetStyle(headerStyle)
	}

	// Set row height for better visibility
	for i := 0; i < len(alignmentExamples); i++ {
		row := sheet.AddRow()
		row.SetHeight(30)

		example := alignmentExamples[i]

		// Alignment name
		row.AddCell(example.name)

		// Example with alignment
		exampleStyle := excelbuilder.StyleConfig{
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: example.horizontal,
				Vertical:   example.vertical,
				WrapText:   true,
			},
			Border: excelbuilder.BorderConfig{
				Top:    excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
				Bottom: excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
				Left:   excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
				Right:  excelbuilder.BorderSide{Style: "thin", Color: "#000000"},
			},
		}
		row.AddCell(example.description).SetStyle(exampleStyle)

		// Horizontal alignment
		row.AddCell(example.horizontal)

		// Vertical alignment
		row.AddCell(example.vertical)
	}

	// Set column widths
	sheet.SetColumnWidth("A", 15)
	sheet.SetColumnWidth("B", 30)
	sheet.SetColumnWidth("C", 12)
	sheet.SetColumnWidth("D", 12)
}

// createNumberFormatsSheet demonstrates number formatting
func createNumberFormatsSheet(wb *excelbuilder.WorkbookBuilder) {
	sheet := wb.AddSheet("Number Formats")

	// Title
	titleStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Size:   16,
			Bold:   true,
			Family: "Arial",
		},
		Alignment: excelbuilder.AlignmentConfig{
			Horizontal: "center",
		},
	}
	sheet.AddRow().AddCell("Number Formatting Examples").SetStyle(titleStyle)

	// Empty row
	sheet.AddRow()

	// Number format examples
	numberExamples := []struct {
		name        string
		value       interface{}
		format      string
		description string
	}{
		{"Integer", 1234, "0", "Whole numbers"},
		{"Decimal", 1234.56, "0.00", "Two decimal places"},
		{"Currency", 1234.56, "$#,##0.00", "Currency format"},
		{"Percentage", 0.1234, "0.00%", "Percentage format"},
		{"Date", time.Now(), "mm/dd/yyyy", "Date format"},
		{"Time", time.Now(), "hh:mm:ss", "Time format"},
		{"Scientific", 1234567, "0.00E+00", "Scientific notation"},
		{"Thousands", 1234567, "#,##0", "Thousands separator"},
	}

	// Headers
	headerStyle := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold:   true,
			Family: "Calibri",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#D9E2F3",
		},
	}

	headers := []string{"Format Type", "Raw Value", "Formatted", "Description"}
	headerRow := sheet.AddRow()
	for _, header := range headers {
		headerRow.AddCell(header).SetStyle(headerStyle)
	}

	// Number format examples
	for _, example := range numberExamples {
		row := sheet.AddRow()

		// Format name
		row.AddCell(example.name)

		// Raw value
		row.AddCell(example.value)

		// Formatted value
		formattedStyle := excelbuilder.StyleConfig{
			NumberFormat: example.format,
			Alignment: excelbuilder.AlignmentConfig{
				Horizontal: "right",
			},
		}
		row.AddCell(example.value).SetNumberFormat(example.format).SetStyle(formattedStyle)

		// Description
		row.AddCell(example.description)
	}

	// Set column widths
	sheet.SetColumnWidth("A", 15)
	sheet.SetColumnWidth("B", 15)
	sheet.SetColumnWidth("C", 15)
	sheet.SetColumnWidth("D", 25)
}
