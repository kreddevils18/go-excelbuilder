package excelbuilder

import (
	"strings"

	"github.com/xuri/excelize/v2"
)

// ExcelBuilder is the main builder for creating Excel files
type ExcelBuilder struct {
	file         *excelize.File
	styleManager *StyleManager
}

// WorkbookBuilder handles workbook-level operations
type WorkbookBuilder struct {
	excelBuilder *ExcelBuilder
	file         *excelize.File
}

// SheetBuilder handles sheet-level operations
type SheetBuilder struct {
	workbookBuilder *WorkbookBuilder
	sheetName       string
	currentRow      int
}

// RowBuilder handles row-level operations
type RowBuilder struct {
	sheetBuilder *SheetBuilder
	rowIndex     int
	currentCol   int
}

// CellBuilder handles cell-level operations
type CellBuilder struct {
	rowBuilder *RowBuilder
	cellRef    string
}

// New creates a new ExcelBuilder instance
func New() *ExcelBuilder {
	file := excelize.NewFile()
	return &ExcelBuilder{
		file:         file,
		styleManager: NewStyleManager(),
	}
}

// NewWorkbook creates a new WorkbookBuilder
func (eb *ExcelBuilder) NewWorkbook() *WorkbookBuilder {
	return &WorkbookBuilder{
		excelBuilder: eb,
		file:         eb.file,
	}
}

// SetProperties sets the workbook properties
func (wb *WorkbookBuilder) SetProperties(props WorkbookProperties) *WorkbookBuilder {
	// Set document properties
	if props.Title != "" {
		wb.file.SetDocProps(&excelize.DocProperties{
			Title:       props.Title,
			Creator:     props.Author,
			Subject:     props.Subject,
			Description: props.Description,
			Category:    props.Category,
			Keywords:    props.Keywords,
		})
	}
	return wb
}

// AddSheet creates a new sheet and returns a SheetBuilder
func (wb *WorkbookBuilder) AddSheet(name string) *SheetBuilder {
	if name == "" {
		return nil
	}

	// Check for invalid characters
	invalidChars := []string{"[", "]", "*", "?", "/", "\\"}
	for _, char := range invalidChars {
		if contains(name, char) {
			return nil
		}
	}

	// Check length (Excel limit is 31 characters)
	if len(name) > 31 {
		return nil
	}

	// Check for reserved names
	reservedNames := []string{"History", "Print_Area", "Print_Titles"}
	for _, reserved := range reservedNames {
		if name == reserved {
			return nil
		}
	}

	// Create the sheet
	index, err := wb.file.NewSheet(name)
	if err != nil {
		return nil
	}

	// Set as active sheet
	wb.file.SetActiveSheet(index)

	return &SheetBuilder{
		workbookBuilder: wb,
		sheetName:       name,
		currentRow:      0,
	}
}

// Build returns the final excelize.File
func (wb *WorkbookBuilder) Build() *excelize.File {
	return wb.file
}

// AddRow creates a new row and returns a RowBuilder
func (sb *SheetBuilder) AddRow() *RowBuilder {
	sb.currentRow++
	return &RowBuilder{
		sheetBuilder: sb,
		rowIndex:     sb.currentRow,
		currentCol:   0,
	}
}

// SetColumnWidth sets the width of a column
func (sb *SheetBuilder) SetColumnWidth(col string, width float64) *SheetBuilder {
	if col == "" {
		return nil
	}

	err := sb.workbookBuilder.file.SetColWidth(sb.sheetName, col, col, width)
	if err != nil {
		return nil
	}

	return sb
}

// MergeCell merges cells in the specified range
func (sb *SheetBuilder) MergeCell(cellRange string) *SheetBuilder {
	// Parse the range to get start and end cells
	// For simplicity, assume cellRange is in format "A1:C1"
	parts := strings.Split(cellRange, ":")
	if len(parts) != 2 {
		return nil
	}
	err := sb.workbookBuilder.file.MergeCell(sb.sheetName, parts[0], parts[1])
	if err != nil {
		return nil
	}
	return sb
}

// AddChart creates a new chart and returns a ChartBuilder
func (sb *SheetBuilder) AddChart() *ChartBuilder {
	return NewChartBuilder(sb.workbookBuilder.file, sb.sheetName)
}

// Build returns the final excelize.File
func (sb *SheetBuilder) Build() *excelize.File {
	return sb.workbookBuilder.Build()
}

// Done returns to the WorkbookBuilder
func (sb *SheetBuilder) Done() *WorkbookBuilder {
	return sb.workbookBuilder
}

// AddCell adds a cell with the given value and returns a CellBuilder
func (rb *RowBuilder) AddCell(value interface{}) *CellBuilder {
	rb.currentCol++
	cellRef, err := excelize.CoordinatesToCellName(rb.currentCol, rb.rowIndex)
	if err != nil {
		return nil
	}

	err = rb.sheetBuilder.workbookBuilder.file.SetCellValue(rb.sheetBuilder.sheetName, cellRef, value)
	if err != nil {
		return nil
	}

	return &CellBuilder{
		rowBuilder: rb,
		cellRef:    cellRef,
	}
}

// SetRowHeight sets the height of the current row
func (rb *RowBuilder) SetRowHeight(height float64) *RowBuilder {
	err := rb.sheetBuilder.workbookBuilder.file.SetRowHeight(rb.sheetBuilder.sheetName, rb.rowIndex, height)
	if err != nil {
		return nil
	}
	return rb
}

// Done returns to the SheetBuilder
func (rb *RowBuilder) Done() *SheetBuilder {
	return rb.sheetBuilder
}



// SetNumberFormat sets the number format for the cell
func (cb *CellBuilder) SetNumberFormat(format string) *CellBuilder {
	if format == "" {
		return cb
	}

	style := &excelize.Style{
		CustomNumFmt: &format,
	}

	styleID, err := cb.rowBuilder.sheetBuilder.workbookBuilder.file.NewStyle(style)
	if err != nil {
		return nil
	}

	err = cb.rowBuilder.sheetBuilder.workbookBuilder.file.SetCellStyle(
		cb.rowBuilder.sheetBuilder.sheetName,
		cb.cellRef,
		cb.cellRef,
		styleID,
	)
	if err != nil {
		return nil
	}

	return cb
}

// SetStyle applies a style configuration to the cell using StyleManager
func (cb *CellBuilder) SetStyle(config StyleConfig) *CellBuilder {
	// Get style flyweight from StyleManager
	styleFlyweight := cb.rowBuilder.sheetBuilder.workbookBuilder.excelBuilder.styleManager.GetStyle(config)
	
	// Apply style to the cell
	err := styleFlyweight.Apply(
		cb.rowBuilder.sheetBuilder.workbookBuilder.file,
		cb.cellRef,
	)
	if err != nil {
		return nil
	}
	
	return cb
}

// SetFormula sets a formula for the cell
func (cb *CellBuilder) SetFormula(formula string) *CellBuilder {
	err := cb.rowBuilder.sheetBuilder.workbookBuilder.file.SetCellFormula(
		cb.rowBuilder.sheetBuilder.sheetName,
		cb.cellRef,
		formula,
	)
	if err != nil {
		return nil
	}
	return cb
}

// SetHyperlink sets a hyperlink for the cell
func (cb *CellBuilder) SetHyperlink(url string) *CellBuilder {
	if url == "" {
		return cb
	}

	err := cb.rowBuilder.sheetBuilder.workbookBuilder.file.SetCellHyperLink(
		cb.rowBuilder.sheetBuilder.sheetName,
		cb.cellRef,
		url,
		"External",
	)
	if err != nil {
		return nil
	}
	return cb
}

// Done returns to the RowBuilder
func (cb *CellBuilder) Done() *RowBuilder {
	return cb.rowBuilder
}

// Helper function to check if string contains substring
func contains(s, substr string) bool {
	for i := 0; i <= len(s)-len(substr); i++ {
		if s[i:i+len(substr)] == substr {
			return true
		}
	}
	return false
}

// Helper function to convert border style string to excelize border style
func getBorderStyle(style string) int {
	switch style {
	case "thin":
		return 1
	case "medium":
		return 2
	case "thick":
		return 5
	case "double":
		return 6
	case "dotted":
		return 3
	case "dashed":
		return 4
	default:
		return 1 // Default to thin
	}
}
