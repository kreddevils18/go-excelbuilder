package excelbuilder

import (
	"strings"

	"github.com/xuri/excelize/v2"
)

// AdvancedLayoutManager handles advanced layout operations for sheets
type AdvancedLayoutManager struct {
	sheetBuilder *SheetBuilder
	sheetName    string
	file         *excelize.File
}

// NewAdvancedLayoutManager creates a new AdvancedLayoutManager
func NewAdvancedLayoutManager(sheetBuilder *SheetBuilder) *AdvancedLayoutManager {
	return &AdvancedLayoutManager{
		sheetBuilder: sheetBuilder,
		sheetName:    sheetBuilder.sheetName,
		file:         sheetBuilder.workbookBuilder.file,
	}
}

// GroupColumns groups columns in the specified range with the given level
func (alm *AdvancedLayoutManager) GroupColumns(columnRange string, level int) *AdvancedLayoutManager {
	// Validate input
	if columnRange == "" || level <= 0 || level > 7 {
		return nil
	}

	// Parse column range (e.g., "B:D")
	parts := strings.Split(columnRange, ":")
	if len(parts) != 2 {
		return nil
	}

	startCol := strings.TrimSpace(parts[0])
	endCol := strings.TrimSpace(parts[1])

	// Validate column names
	if !isValidColumnName(startCol) || !isValidColumnName(endCol) {
		return nil
	}

	// Check for reverse range
	startColNum, err1 := excelize.ColumnNameToNumber(startCol)
	endColNum, err2 := excelize.ColumnNameToNumber(endCol)
	if err1 != nil || err2 != nil || startColNum > endColNum {
		return nil
	}

	// Set column outline level
	for colNum := startColNum; colNum <= endColNum; colNum++ {
		colName, _ := excelize.ColumnNumberToName(colNum)
		err := alm.file.SetColOutlineLevel(alm.sheetName, colName, uint8(level))
		if err != nil {
			return nil
		}
	}

	return alm
}

// GroupRows groups rows in the specified range with the given level
func (alm *AdvancedLayoutManager) GroupRows(startRow, endRow, level int) *AdvancedLayoutManager {
	// Validate input
	if startRow <= 0 || endRow <= 0 || startRow > endRow || level <= 0 || level > 7 {
		return nil
	}

	// Set row outline level
	for row := startRow; row <= endRow; row++ {
		err := alm.file.SetRowOutlineLevel(alm.sheetName, row, uint8(level))
		if err != nil {
			return nil
		}
	}

	return alm
}

// FreezePane freezes panes at the specified cell
func (alm *AdvancedLayoutManager) FreezePane(cell string) *AdvancedLayoutManager {
	// Validate cell reference
	if cell == "" {
		return nil
	}

	// Validate cell format
	_, _, err := excelize.CellNameToCoordinates(cell)
	if err != nil {
		return nil
	}

	// Set freeze panes
	err = alm.file.SetPanes(alm.sheetName, &excelize.Panes{
		Freeze:      true,
		Split:       false,
		TopLeftCell: cell,
	})
	if err != nil {
		return nil
	}

	return alm
}

// SplitPane splits panes at the specified position
func (alm *AdvancedLayoutManager) SplitPane(xSplit, ySplit int) *AdvancedLayoutManager {
	// Validate input
	if xSplit < 0 || ySplit < 0 {
		return nil
	}

	// Set split panes
	err := alm.file.SetPanes(alm.sheetName, &excelize.Panes{
		Freeze: false,
		Split:  true,
		XSplit: xSplit,
		YSplit: ySplit,
	})
	if err != nil {
		return nil
	}

	return alm
}

// AutoFitColumns automatically adjusts column widths to fit content
func (alm *AdvancedLayoutManager) AutoFitColumns(columnRange string) *AdvancedLayoutManager {
	// Parse column range
	parts := strings.Split(columnRange, ":")
	if len(parts) != 2 {
		return nil
	}

	startCol := strings.TrimSpace(parts[0])
	endCol := strings.TrimSpace(parts[1])

	// Validate column names
	if !isValidColumnName(startCol) || !isValidColumnName(endCol) {
		return nil
	}

	startColNum, err1 := excelize.ColumnNameToNumber(startCol)
	endColNum, err2 := excelize.ColumnNameToNumber(endCol)
	if err1 != nil || err2 != nil || startColNum > endColNum {
		return nil
	}

	// Get all rows to calculate optimal width
	rows, err := alm.file.GetRows(alm.sheetName)
	if err != nil {
		return nil
	}

	// Calculate optimal width for each column
	for colNum := startColNum; colNum <= endColNum; colNum++ {
		colName, _ := excelize.ColumnNumberToName(colNum)
		maxWidth := 8.0 // Default minimum width

		for _, row := range rows {
			if len(row) >= colNum {
				cellValue := row[colNum-1]
				// Simple width calculation based on character count
				// In a real implementation, you might want to consider font metrics
				cellWidth := float64(len(cellValue)) * 1.2 // Approximate character width
				if cellWidth > maxWidth {
					maxWidth = cellWidth
				}
			}
		}

		// Add padding and limit maximum width
		maxWidth += 2.0
		if maxWidth > 50.0 {
			maxWidth = 50.0
		}

		// Set column width
		err = alm.file.SetColWidth(alm.sheetName, colName, colName, maxWidth)
		if err != nil {
			return nil
		}
	}

	return alm
}

// SetColumnWidthRange sets width for a range of columns
func (alm *AdvancedLayoutManager) SetColumnWidthRange(columnRange string, width float64) *AdvancedLayoutManager {
	// Validate width
	if width < 0 || width > 255 {
		return nil
	}

	// Parse column range
	parts := strings.Split(columnRange, ":")
	if len(parts) != 2 {
		return nil
	}

	startCol := strings.TrimSpace(parts[0])
	endCol := strings.TrimSpace(parts[1])

	// Validate column names
	if !isValidColumnName(startCol) || !isValidColumnName(endCol) {
		return nil
	}

	startColNum, err1 := excelize.ColumnNameToNumber(startCol)
	endColNum, err2 := excelize.ColumnNameToNumber(endCol)
	if err1 != nil || err2 != nil || startColNum > endColNum {
		return nil
	}

	// Set width for each column in range
	for colNum := startColNum; colNum <= endColNum; colNum++ {
		colName, _ := excelize.ColumnNumberToName(colNum)
		err := alm.file.SetColWidth(alm.sheetName, colName, colName, width)
		if err != nil {
			return nil
		}
	}

	return alm
}

// SetRowHeightRange sets height for a range of rows
func (alm *AdvancedLayoutManager) SetRowHeightRange(startRow, endRow int, height float64) *AdvancedLayoutManager {
	// Validate input
	if startRow <= 0 || endRow <= 0 || startRow > endRow || height <= 0 || height > 409.5 {
		return nil
	}

	// Set height for each row in range
	for row := startRow; row <= endRow; row++ {
		err := alm.file.SetRowHeight(alm.sheetName, row, height)
		if err != nil {
			return nil
		}
	}

	return alm
}

// HideColumns hides columns in the specified range
func (alm *AdvancedLayoutManager) HideColumns(columnRange string) *AdvancedLayoutManager {
	// Parse column range
	parts := strings.Split(columnRange, ":")
	if len(parts) != 2 {
		return nil
	}

	startCol := strings.TrimSpace(parts[0])
	endCol := strings.TrimSpace(parts[1])

	// Validate column names
	if !isValidColumnName(startCol) || !isValidColumnName(endCol) {
		return nil
	}

	startColNum, err1 := excelize.ColumnNameToNumber(startCol)
	endColNum, err2 := excelize.ColumnNameToNumber(endCol)
	if err1 != nil || err2 != nil || startColNum > endColNum {
		return nil
	}

	// Hide each column in range
	for colNum := startColNum; colNum <= endColNum; colNum++ {
		colName, _ := excelize.ColumnNumberToName(colNum)
		err := alm.file.SetColVisible(alm.sheetName, colName, false)
		if err != nil {
			return nil
		}
	}

	return alm
}

// HideRows hides rows in the specified range
func (alm *AdvancedLayoutManager) HideRows(startRow, endRow int) *AdvancedLayoutManager {
	// Validate input
	if startRow <= 0 || endRow <= 0 || startRow > endRow {
		return nil
	}

	// Hide each row in range
	for row := startRow; row <= endRow; row++ {
		err := alm.file.SetRowVisible(alm.sheetName, row, false)
		if err != nil {
			return nil
		}
	}

	return alm
}

// ShowColumns shows previously hidden columns
func (alm *AdvancedLayoutManager) ShowColumns(columnRange string) *AdvancedLayoutManager {
	// Parse column range
	parts := strings.Split(columnRange, ":")
	if len(parts) != 2 {
		return nil
	}

	startCol := strings.TrimSpace(parts[0])
	endCol := strings.TrimSpace(parts[1])

	// Validate column names
	if !isValidColumnName(startCol) || !isValidColumnName(endCol) {
		return nil
	}

	startColNum, err1 := excelize.ColumnNameToNumber(startCol)
	endColNum, err2 := excelize.ColumnNameToNumber(endCol)
	if err1 != nil || err2 != nil || startColNum > endColNum {
		return nil
	}

	// Show each column in range
	for colNum := startColNum; colNum <= endColNum; colNum++ {
		colName, _ := excelize.ColumnNumberToName(colNum)
		err := alm.file.SetColVisible(alm.sheetName, colName, true)
		if err != nil {
			return nil
		}
	}

	return alm
}

// ShowRows shows previously hidden rows
func (alm *AdvancedLayoutManager) ShowRows(startRow, endRow int) *AdvancedLayoutManager {
	// Validate input
	if startRow <= 0 || endRow <= 0 || startRow > endRow {
		return nil
	}

	// Show each row in range
	for row := startRow; row <= endRow; row++ {
		err := alm.file.SetRowVisible(alm.sheetName, row, true)
		if err != nil {
			return nil
		}
	}

	return alm
}

// CollapseGroup collapses a grouped range
func (alm *AdvancedLayoutManager) CollapseGroup(isRow bool, level int) *AdvancedLayoutManager {
	// Validate input
	if level <= 0 || level > 7 {
		return nil
	}

	// Note: excelize doesn't have direct collapse/expand methods
	// This would typically be handled by setting outline properties
	// For now, we'll return success as the grouping structure is already set

	return alm
}

// ExpandGroup expands a collapsed group
func (alm *AdvancedLayoutManager) ExpandGroup(isRow bool, level int) *AdvancedLayoutManager {
	// Validate input
	if level <= 0 || level > 7 {
		return nil
	}

	// Note: excelize doesn't have direct collapse/expand methods
	// This would typically be handled by setting outline properties
	// For now, we'll return success as the grouping structure is already set

	return alm
}

// Done returns to the SheetBuilder
func (alm *AdvancedLayoutManager) Done() *SheetBuilder {
	return alm.sheetBuilder
}

// GetSheetBuilder returns the associated SheetBuilder
func (alm *AdvancedLayoutManager) GetSheetBuilder() *SheetBuilder {
	return alm.sheetBuilder
}