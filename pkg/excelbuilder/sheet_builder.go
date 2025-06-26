package excelbuilder

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// SheetBuilder handles sheet-level operations
type SheetBuilder struct {
	workbookBuilder *WorkbookBuilder
	sheetName       string
	currentRow      int
}

// GetCurrentRow returns the current row number (1-indexed).
func (sb *SheetBuilder) GetCurrentRow() int {
	return sb.currentRow
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

// AddRows adds multiple rows of data to the sheet.
func (sb *SheetBuilder) AddRows(data [][]interface{}) *SheetBuilder {
	for _, rowData := range data {
		row := sb.AddRow()
		for _, cellData := range rowData {
			row.AddCell(cellData)
		}
	}
	return sb
}

// SetColumnWidth sets the width of a column.
func (sb *SheetBuilder) SetColumnWidth(col string, width float64) *SheetBuilder {
	// Validate column name
	if col == "" {
		return nil
	}

	// Validate width (negative or extremely large values)
	if width < 0 || width > 255 {
		return nil
	}

	// Validate column format (should be A-Z, AA-ZZ, etc., not contain numbers)
	if !isValidColumnName(col) {
		return nil
	}

	err := sb.workbookBuilder.file.SetColWidth(sb.sheetName, col, col, width)
	if err != nil {
		// Return nil for any errors during setting column width
		return nil
	}

	return sb
}

// AutoSizeColumn adjusts the column width to fit the content.
func (sb *SheetBuilder) AutoSizeColumn(col string) *SheetBuilder {
	if col == "" {
		return sb
	}

	rows, err := sb.workbookBuilder.file.GetRows(sb.sheetName)
	if err != nil {
		return sb
	}

	colIndex, err := excelize.ColumnNameToNumber(col)
	if err != nil {
		return sb
	}

	var maxWidth float64
	for _, row := range rows {
		if len(row) >= colIndex {
			cellValue := row[colIndex-1]
			cellWidth := float64(len(cellValue)) // Simple width calculation
			if cellWidth > maxWidth {
				maxWidth = cellWidth
			}
		}
	}

	// Add some padding
	maxWidth += 2

	sb.SetColumnWidth(col, maxWidth)

	return sb
}

// AutoSizeColumns adjusts all column widths to fit the content.
func (sb *SheetBuilder) AutoSizeColumns() *SheetBuilder {
	cols, err := sb.workbookBuilder.file.GetCols(sb.sheetName)
	if err != nil {
		return sb
	}

	for i := range cols {
		colName, _ := excelize.ColumnNumberToName(i + 1)
		sb.AutoSizeColumn(colName)
	}

	return sb
}

// SetCell sets the value of a specific cell and returns a CellBuilder
func (sb *SheetBuilder) SetCell(cellRef string, value interface{}) *CellBuilder {
	// Validate input
	if cellRef == "" {
		return nil
	}

	err := sb.workbookBuilder.file.SetCellValue(sb.sheetName, cellRef, value)
	if err != nil {
		return nil
	}

	return &CellBuilder{
		rowBuilder:   nil, // This will be nil since we're setting directly by reference
		cellRef:      cellRef,
		sheetBuilder: sb, // Add reference to sheet builder
	}
}

// MergeRange merges cells in the specified range (alias for MergeCell)
func (sb *SheetBuilder) MergeRange(cellRange string) *SheetBuilder {
	return sb.MergeCell(cellRange)
}

// MergeCell merges cells in the specified range
func (sb *SheetBuilder) MergeCell(cellRange string) *SheetBuilder {
	// Validate input
	if cellRange == "" {
		return nil
	}

	// Parse the range to get start and end cells
	// For simplicity, assume cellRange is in format "A1:C1"
	parts := strings.Split(cellRange, ":")
	if len(parts) != 2 {
		return nil
	}

	// Validate that start and end cells are different
	if parts[0] == parts[1] {
		// Single cell merge is allowed but should be handled gracefully
		return sb
	}

	// Check for reverse range (e.g., "D1:A1")
	if isReverseRange(parts[0], parts[1]) {
		return nil
	}

	err := sb.workbookBuilder.file.MergeCell(sb.sheetName, parts[0], parts[1])
	if err != nil {
		return nil
	}
	return sb
}

// FreezePanes provides a way to freeze rows and columns, making them always visible during scrolling.
//
// Parameters:
//   - cols: The number of columns to freeze from the left side of the sheet (e.g., 1 to freeze column A).
//   - rows: The number of rows to freeze from the top of the sheet (e.g., 1 to freeze row 1).
//
// Example:
//
//	// Freeze the first column and the top two rows
//	sheet.FreezePanes(1, 2)
func (sb *SheetBuilder) FreezePanes(cols, rows int) *SheetBuilder {
	if cols < 0 || rows < 0 {
		return sb // Invalid input
	}
	if cols == 0 && rows == 0 {
		return sb // Nothing to freeze
	}

	// excelize requires the top-left cell of the unfrozen pane to be specified.
	// For example, to freeze 1 column, the split is after column A, and the first unfrozen cell is B1.
	// To freeze 1 row, the split is after row 1, and the first unfrozen cell is A2.
	colName, err := excelize.ColumnNumberToName(cols + 1)
	if err != nil {
		fmt.Printf("could not convert column number to name for freeze panes: %v\n", err)
		return sb
	}
	cell := fmt.Sprintf("%s%d", colName, rows+1)

	panes := &excelize.Panes{
		Freeze:      true,
		XSplit:      cols,
		YSplit:      rows,
		TopLeftCell: cell,
	}

	if err := sb.workbookBuilder.file.SetPanes(sb.sheetName, panes); err != nil {
		// In a real library, this error should be handled more gracefully,
		// perhaps by accumulating errors in the builder.
		fmt.Printf("could not set freeze panes for sheet %s: %v\n", sb.sheetName, err)
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

// SetConditionalFormatting applies conditional formatting to a range of cells.
func (sb *SheetBuilder) SetConditionalFormatting(config ConditionalFormattingConfig) *SheetBuilder {
	var formats []excelize.ConditionalFormatOptions
	for _, rule := range config.Rules {
		style := sb.workbookBuilder.excelBuilder.styleManager.GetStyle(rule.Style, sb.workbookBuilder.file)

		// excelize requires a pointer to the style ID.
		styleIDPtr := style.GetID()

		format := excelize.ConditionalFormatOptions{
			Type:     rule.Type,
			Criteria: rule.Operator,
			Format:   &styleIDPtr,
		}

		switch rule.Type {
		case "cell", "average", "duplicateValues", "uniqueValues", "top", "bottom", "blanks", "noBlanks", "errors", "noErrors", "timePeriod":
			if rule.Value != nil {
				format.Value = rule.Value.(string)
			}
		case "2_color_scale", "3_color_scale":
			format.MinType = rule.ColorScale.MinType
			format.MinValue = rule.ColorScale.MinValue
			format.MinColor = rule.ColorScale.MinColor
			format.MaxType = rule.ColorScale.MaxType
			format.MaxValue = rule.ColorScale.MaxValue
			format.MaxColor = rule.ColorScale.MaxColor
			if rule.Type == "3_color_scale" {
				format.MidType = rule.ColorScale.MidType
				format.MidValue = rule.ColorScale.MidValue
				format.MidColor = rule.ColorScale.MidColor
			}
		case "data_bar":
			format.BarColor = rule.DataBar.Color
			format.MinValue = strconv.Itoa(int(rule.DataBar.MinLength))
			format.MaxValue = strconv.Itoa(int(rule.DataBar.MaxLength))
			format.BarBorderColor = rule.DataBar.BorderColor
			format.BarDirection = rule.DataBar.Direction
			format.BarOnly = rule.DataBar.BarOnly
		case "icon_set":
			format.IconStyle = rule.IconSet.Style
			format.ReverseIcons = rule.IconSet.Reverse
			format.IconsOnly = rule.IconSet.IconsOnly
		}

		formats = append(formats, format)
	}

	// The range is a single string like "A1:A10"
	// excelize SetConditionalFormat takes a single range and a slice of formats
	// We need to get the range from the config.
	// The range needs to be split into individual cells for some reason in some versions.
	// Let's assume a single range string is fine.

	// The format string needs to be created from the slice of options.
	// This is a bit tricky as the library doesn't directly take the struct.
	// Let's try to pass the struct directly.
	// After re-reading the docs, it seems we need to pass a JSON string.
	// But we should use the structs to build it, not manual string concatenation.

	// The issue is that we need to pass a slice of format options to SetConditionalFormat.
	// Let's try to do that.

	// The `SetConditionalFormat` function takes a range and a slice of `ConditionalFormatOptions`.
	// The previous code was trying to build a JSON string manually, which is wrong.
	// The correct way is to build the slice of structs and pass it.

	// The range is a single string like "A1:A10"
	// The function signature is: `func (f *File) SetConditionalFormat(sheet, rangeRef string, opts []ConditionalFormatOptions) error`
	// So we need to pass the slice of formats we built.

	// The range is a single string like "A1:A10"
	// The function signature is: `func (f *File) SetConditionalFormat(sheet, rangeRef string, opts []ConditionalFormatOptions) error`
	// So we need to pass the slice of formats we built.

	// The range is a single string like "A1:A10"
	err := sb.workbookBuilder.file.SetConditionalFormat(sb.sheetName, config.Range, formats)
	if err != nil {
		// Handle error
	}

	return sb
}

// AddRowsBatch adds multiple rows with data in a single operation
func (sb *SheetBuilder) AddRowsBatch(batchData [][]interface{}) *SheetBuilder {
	for _, rowData := range batchData {
		row := sb.AddRow()
		for _, cellData := range rowData {
			row.AddCell(cellData).Done()
		}
		row.Done()
	}
	return sb
}

// AddRowsBatchWithStyles adds multiple rows with individual styles.
func (sb *SheetBuilder) AddRowsBatchWithStyles(batchData []BatchRowData) *SheetBuilder {
	for _, rowData := range batchData {
		row := sb.AddRow()
		for _, cellData := range rowData.Cells {
			row.AddCell(cellData).WithStyle(rowData.Style)
		}
	}
	return sb
}

// ApplyStyleBatch applies a style to multiple ranges.
func (sb *SheetBuilder) ApplyStyleBatch(operations []BatchStyleOperation) *SheetBuilder {
	for _, op := range operations {
		style := sb.workbookBuilder.excelBuilder.styleManager.GetStyle(op.Style, sb.workbookBuilder.file)
		// Apply style to range. This is a simplified example.
		// In a real implementation, you would iterate over cells in the range.
		sb.workbookBuilder.file.SetCellStyle(sb.sheetName, op.Range, op.Range, style.GetID())
	}
	return sb
}

// SetRowHeight sets the height for a specific row.
func (sb *SheetBuilder) SetRowHeight(row int, height float64) *SheetBuilder {
	if row <= 0 || height <= 0 || height > 409 {
		return nil // Invalid input
	}
	err := sb.workbookBuilder.file.SetRowHeight(sb.sheetName, row, height)
	if err != nil {
		return nil
	}
	return sb
}

// WithProtection applies protection settings to the sheet.
func (sb *SheetBuilder) WithProtection(config SheetProtectionConfig) *SheetBuilder {
	opts := &excelize.SheetProtectionOptions{
		Password:            config.Password,
		SelectLockedCells:   config.SelectLockedCells,
		SelectUnlockedCells: config.SelectUnlockedCells,
		FormatCells:         config.FormatCells,
		FormatColumns:       config.FormatColumns,
		FormatRows:          config.FormatRows,
		InsertColumns:       config.InsertColumns,
		InsertRows:          config.InsertRows,
		InsertHyperlinks:    config.InsertHyperlinks,
		DeleteColumns:       config.DeleteColumns,
		DeleteRows:          config.DeleteRows,
		Sort:                config.Sort,
		AutoFilter:          config.AutoFilter,
		PivotTables:         config.PivotTables,
		EditObjects:         config.EditObjects,
		EditScenarios:       config.EditScenarios,
	}

	err := sb.workbookBuilder.file.ProtectSheet(sb.sheetName, opts)
	if err != nil {
		// Handle error
	}

	return sb
}

// WithTabColor sets the tab color of the sheet using RGB hex color.
func (sb *SheetBuilder) WithTabColor(hexColor string) *SheetBuilder {
	if sb.workbookBuilder == nil || sb.workbookBuilder.file == nil {
		return sb // Or handle error appropriately
	}

	// Basic validation for hex color
	if len(hexColor) > 0 && hexColor[0] == '#' {
		hexColor = hexColor[1:]
	}
	if len(hexColor) != 6 {
		return sb // Invalid hex code, handle error or ignore
	}

	opts := &excelize.SheetPropsOptions{
		TabColorRGB: &hexColor,
	}

	err := sb.workbookBuilder.file.SetSheetProps(sb.sheetName, opts)
	if err != nil {
		// Handle error, e.g., log it
	}

	return sb
}

// isValidColumnName validates Excel column names
func isValidColumnName(columnName string) bool {
	if columnName == "" {
		return false
	}

	// Check if all characters are uppercase letters
	for _, char := range columnName {
		if char < 'A' || char > 'Z' {
			return false
		}
	}

	// Check length (Excel supports up to 3 characters: A-XFD)
	if len(columnName) > 3 {
		return false
	}

	// Additional validation for maximum column (XFD = 16384)
	if len(columnName) == 3 {
		if columnName > "XFD" {
			return false
		}
	}

	return true
}

// isReverseRange checks if the range is reversed (e.g., "D1:A1")
func isReverseRange(startCell, endCell string) bool {
	// Parse cell coordinates
	startCol, startRow, err1 := excelize.CellNameToCoordinates(startCell)
	endCol, endRow, err2 := excelize.CellNameToCoordinates(endCell)

	if err1 != nil || err2 != nil {
		return true // Invalid cell references are considered reverse
	}

	// Check if start is after end (reverse range)
	return startCol > endCol || startRow > endRow
}

// WithTabColorRGB sets the tab color of the sheet using RGB values.
func (sb *SheetBuilder) WithTabColorRGB(color RGBColor) *SheetBuilder {
	hexColor := rgbToHex(color.R, color.G, color.B)
	return sb.WithTabColor(hexColor)
}

// WithTabColorTheme sets the tab color of the sheet using a theme color.
func (sb *SheetBuilder) WithTabColorTheme(theme int, tint ...float64) *SheetBuilder {
	if sb.workbookBuilder == nil || sb.workbookBuilder.file == nil {
		return sb
	}

	opts := &excelize.SheetPropsOptions{
		TabColorTheme: &theme,
	}

	if len(tint) > 0 {
		opts.TabColorTint = &tint[0]
	}

	err := sb.workbookBuilder.file.SetSheetProps(sb.sheetName, opts)
	if err != nil {
		// Handle error, e.g., log it
	}

	return sb
}

// WithTabColorIndexed sets the tab color of the sheet using an indexed color.
func (sb *SheetBuilder) WithTabColorIndexed(index int) *SheetBuilder {
	if sb.workbookBuilder == nil || sb.workbookBuilder.file == nil {
		return sb
	}

	opts := &excelize.SheetPropsOptions{
		TabColorIndexed: &index,
	}

	err := sb.workbookBuilder.file.SetSheetProps(sb.sheetName, opts)
	if err != nil {
		// Handle error, e.g., log it
	}

	return sb
}

// NewPivotTable creates a new PivotTableBuilder for this sheet
func (sb *SheetBuilder) NewPivotTable(targetSheet, sourceRange string) *PivotTableBuilder {
	return NewPivotTableBuilder(sb, targetSheet, sourceRange)
}

// GetLayoutManager returns an AdvancedLayoutManager for this sheet
func (sb *SheetBuilder) GetLayoutManager() *AdvancedLayoutManager {
	return NewAdvancedLayoutManager(sb)
}

// Done returns to the WorkbookBuilder
func (sb *SheetBuilder) Done() *WorkbookBuilder {
	return sb.workbookBuilder
}

// rgbToHex converts RGB values to a hex color string.
func rgbToHex(r, g, b int) string {
	return fmt.Sprintf("%02X%02X%02X", r, g, b)
}
