package excelbuilder

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// RowBuilder handles row-level operations
type RowBuilder struct {
	sheetBuilder *SheetBuilder
	rowIndex     int
	currentCol   int
	hasError     bool
}

// AddCell adds a cell with the given value and returns a CellBuilder
func (rb *RowBuilder) AddCell(value interface{}) *CellBuilder {
	rb.currentCol++
	cellRef, err := excelize.CoordinatesToCellName(rb.currentCol, rb.rowIndex)
	if err != nil {
		rb.sheetBuilder.workbookBuilder.excelBuilder.AddError(fmt.Errorf("failed to convert coordinates to cell name: %w", err))
		rb.hasError = true
		return &CellBuilder{
			rowBuilder:   rb,
			sheetBuilder: rb.sheetBuilder,
			cellRef:      "A1", // fallback
			hasError:     true,
		}
	}

	err = rb.sheetBuilder.workbookBuilder.file.SetCellValue(rb.sheetBuilder.sheetName, cellRef, value)
	if err != nil {
		rb.sheetBuilder.workbookBuilder.excelBuilder.AddError(fmt.Errorf("failed to set cell value at %s: %w", cellRef, err))
		rb.hasError = true
		return &CellBuilder{
			rowBuilder:   rb,
			sheetBuilder: rb.sheetBuilder,
			cellRef:      cellRef,
			hasError:     true,
		}
	}

	return &CellBuilder{
		rowBuilder:   rb,
		sheetBuilder: rb.sheetBuilder,
		cellRef:      cellRef,
		hasError:     false,
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
	// Validate height input. Max row height in Excel is 409.
	if height <= 0 || height > 409 {
		rb.sheetBuilder.workbookBuilder.excelBuilder.AddError(fmt.Errorf("invalid row height %f (must be between 0 and 409)", height))
		rb.hasError = true
		return rb
	}

	err := rb.sheetBuilder.workbookBuilder.file.SetRowHeight(rb.sheetBuilder.sheetName, rb.rowIndex, height)
	if err != nil {
		rb.sheetBuilder.workbookBuilder.excelBuilder.AddError(fmt.Errorf("failed to set row height for row %d: %w", rb.rowIndex, err))
		rb.hasError = true
		return rb
	}
	return rb
}

// Done returns to the SheetBuilder
func (rb *RowBuilder) Done() *SheetBuilder {
	return rb.sheetBuilder
}
