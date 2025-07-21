package excelbuilder

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// CellBuilder handles cell-level operations
type CellBuilder struct {
	rowBuilder   *RowBuilder
	sheetBuilder *SheetBuilder
	cellRef      string
	hasError     bool
}

// WithValue sets the value of the cell.
func (cb *CellBuilder) WithValue(value interface{}) *CellBuilder {
	err := cb.sheetBuilder.workbookBuilder.file.SetCellValue(
		cb.sheetBuilder.sheetName,
		cb.cellRef,
		value,
	)
	if err != nil {
		cb.sheetBuilder.workbookBuilder.excelBuilder.AddError(fmt.Errorf("failed to set cell value at %s: %w", cb.cellRef, err))
		cb.hasError = true
		return cb
	}
	return cb
}

// WithStyle applies a style configuration to the cell using StyleManager
func (cb *CellBuilder) WithStyle(config StyleConfig) *CellBuilder {
	// Get style flyweight from StyleManager
	styleFlyweight := cb.sheetBuilder.workbookBuilder.excelBuilder.styleManager.GetStyle(config, cb.sheetBuilder.workbookBuilder.file)

	// Apply style to the cell
	err := styleFlyweight.Apply(
		cb.sheetBuilder.workbookBuilder.file,
		cb.sheetBuilder.sheetName,
		cb.cellRef,
	)
	if err != nil {
		cb.sheetBuilder.workbookBuilder.excelBuilder.AddError(fmt.Errorf("failed to apply style to cell %s: %w", cb.cellRef, err))
		cb.hasError = true
		return cb
	}

	return cb
}

// WithDataValidation adds a data validation rule to the cell.
func (cb *CellBuilder) WithDataValidation(dvc *DataValidationConfig) *CellBuilder {
	if dvc == nil {
		return cb
	}

	dv := excelize.NewDataValidation(true)
	dv.SetSqref(cb.cellRef)
	dv.AllowBlank = dvc.AllowBlank
	dv.ShowInputMessage = dvc.ShowInputMessage
	dv.ShowErrorMessage = dvc.ShowErrorMessage
	dv.ErrorTitle = &dvc.ErrorTitle
	dv.Error = &dvc.ErrorBody
	if dvc.ErrorStyle != "" {
		dv.ErrorStyle = &dvc.ErrorStyle
	}
	dv.PromptTitle = &dvc.PromptTitle
	dv.Prompt = &dvc.PromptBody

	switch dvc.Type {
	case "list":
		_ = dv.SetDropList(dvc.Formula1)
	case "whole", "decimal", "date", "time", "text_length":
		var f1, f2 interface{}
		if len(dvc.Formula1) > 0 {
			f1 = dvc.Formula1[0]
		}
		if len(dvc.Formula2) > 0 {
			f2 = dvc.Formula2[0]
		}
		_ = dv.SetRange(f1, f2, getValidationType(dvc.Type), getOperatorType(dvc.Operator))
	case "custom":
		var f1, f2 string
		if len(dvc.Formula1) > 0 {
			f1 = dvc.Formula1[0]
		}
		if len(dvc.Formula2) > 0 {
			f2 = dvc.Formula2[0]
		}
		dv.Formula1 = f1
		dv.Formula2 = f2
	}

	err := cb.sheetBuilder.workbookBuilder.file.AddDataValidation(cb.sheetBuilder.sheetName, dv)
	if err != nil {
		cb.sheetBuilder.workbookBuilder.excelBuilder.AddError(fmt.Errorf("failed to add data validation to cell %s: %w", cb.cellRef, err))
		cb.hasError = true
		return cb
	}
	return cb
}

// WithNumberFormat sets the number format for the cell
func (cb *CellBuilder) WithNumberFormat(format string) *CellBuilder {
	if format == "" {
		return cb
	}

	style := &excelize.Style{
		CustomNumFmt: &format,
	}

	styleID, err := cb.sheetBuilder.workbookBuilder.file.NewStyle(style)
	if err != nil {
		cb.sheetBuilder.workbookBuilder.excelBuilder.AddError(fmt.Errorf("failed to create style for cell %s: %w", cb.cellRef, err))
		cb.hasError = true
		return cb
	}

	err = cb.sheetBuilder.workbookBuilder.file.SetCellStyle(
		cb.sheetBuilder.sheetName,
		cb.cellRef,
		cb.cellRef,
		styleID,
	)
	if err != nil {
		cb.sheetBuilder.workbookBuilder.excelBuilder.AddError(fmt.Errorf("failed to set cell style for cell %s: %w", cb.cellRef, err))
		cb.hasError = true
		return cb
	}

	return cb
}

// WithFormula sets a formula for the cell
func (cb *CellBuilder) WithFormula(formula string) *CellBuilder {
	err := cb.sheetBuilder.workbookBuilder.file.SetCellFormula(
		cb.sheetBuilder.sheetName,
		cb.cellRef,
		formula,
	)
	if err != nil {
		cb.sheetBuilder.workbookBuilder.excelBuilder.AddError(fmt.Errorf("failed to set formula for cell %s: %w", cb.cellRef, err))
		cb.hasError = true
		return cb
	}
	return cb
}

// WithHyperlink sets a hyperlink for the cell
func (cb *CellBuilder) WithHyperlink(url string) *CellBuilder {
	if url == "" {
		return cb
	}

	err := cb.sheetBuilder.workbookBuilder.file.SetCellHyperLink(
		cb.sheetBuilder.sheetName,
		cb.cellRef,
		url,
		"External",
	)
	if err != nil {
		cb.sheetBuilder.workbookBuilder.excelBuilder.AddError(fmt.Errorf("failed to set hyperlink for cell %s: %w", cb.cellRef, err))
		cb.hasError = true
		return cb
	}
	return cb
}

// WithMergeRange merges this cell with a given range.
// The cell this is called on will be the top-left cell of the merged range.
func (cb *CellBuilder) WithMergeRange(endCellRef string) *CellBuilder {
	err := cb.sheetBuilder.workbookBuilder.file.MergeCell(cb.sheetBuilder.sheetName, cb.cellRef, endCellRef)
	if err != nil {
		cb.sheetBuilder.workbookBuilder.excelBuilder.AddError(fmt.Errorf("failed to merge cells %s to %s: %w", cb.cellRef, endCellRef, err))
		cb.hasError = true
		return cb
	}
	return cb
}

// Done returns to the RowBuilder to continue building the row
func (cb *CellBuilder) Done() *RowBuilder {
	// If rowBuilder is nil, it means the cell was created directly on the sheet.
	// In this case, there's no row to return to, so we can return nil or handle it differently.
	// For now, we return the (potentially nil) rowBuilder.
	return cb.rowBuilder
}

// Helper functions for data validation
func getErrorStyle(style string) excelize.DataValidationErrorStyle {
	switch style {
	case "stop":
		return excelize.DataValidationErrorStyleStop
	case "warning":
		return excelize.DataValidationErrorStyleWarning
	case "information":
		return excelize.DataValidationErrorStyleInformation
	default:
		return excelize.DataValidationErrorStyleStop
	}
}

func getValidationType(t string) excelize.DataValidationType {
	switch t {
	case "whole":
		return excelize.DataValidationTypeWhole
	case "decimal":
		return excelize.DataValidationTypeDecimal
	case "date":
		return excelize.DataValidationTypeDate
	case "time":
		return excelize.DataValidationTypeTime
	case "text_length":
		return excelize.DataValidationTypeTextLength
	case "list":
		return excelize.DataValidationTypeList
	default:
		return excelize.DataValidationTypeNone
	}
}

func getOperatorType(op string) excelize.DataValidationOperator {
	switch op {
	case "between":
		return excelize.DataValidationOperatorBetween
	case "not_between":
		return excelize.DataValidationOperatorNotBetween
	case "equal":
		return excelize.DataValidationOperatorEqual
	case "not_equal":
		return excelize.DataValidationOperatorNotEqual
	case "greater_than":
		return excelize.DataValidationOperatorGreaterThan
	case "less_than":
		return excelize.DataValidationOperatorLessThan
	case "greater_than_or_equal_to":
		return excelize.DataValidationOperatorGreaterThanOrEqual
	case "less_than_or_equal_to":
		return excelize.DataValidationOperatorLessThanOrEqual
	default:
		return excelize.DataValidationOperatorGreaterThan
	}
}
