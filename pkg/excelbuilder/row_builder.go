package excelbuilder

import "github.com/xuri/excelize/v2"

// RowBuilder handles row-level operations
type RowBuilder struct {
	sheetBuilder *SheetBuilder
	rowIndex     int
	currentCol   int
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

// AddCells adds multiple cells with the given values and returns the RowBuilder
func (rb *RowBuilder) AddCells(values ...interface{}) *RowBuilder {
	for _, value := range values {
		rb.AddCell(value)
	}
	return rb
}

// SetHeight sets the height of the current row
func (rb *RowBuilder) SetHeight(height float64) *RowBuilder {
	// Validate height input
	if height <= 0 || height > 409.5 {
		return nil
	}
	
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