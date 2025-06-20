package excelbuilder

import (
	"strings"
	"github.com/xuri/excelize/v2"
)

// CellBuilder handles cell-level operations
type CellBuilder struct {
	rowBuilder *RowBuilder
	cellRef    string
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
	styleFlyweight := cb.rowBuilder.sheetBuilder.workbookBuilder.excelBuilder.styleManager.GetStyle(config, cb.rowBuilder.sheetBuilder.workbookBuilder.file)

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

// SetDataValidation adds a data validation rule to the cell.
func (cb *CellBuilder) SetDataValidation(dvc DataValidationConfig) *CellBuilder {
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
		// Ignoring error for now
		_ = dv.SetDropList(dvc.Formula1)
	case "whole", "decimal", "date", "time", "text_length":
		var f1, f2 interface{}
		if len(dvc.Formula1) > 0 {
			f1 = dvc.Formula1[0]
		}
		if len(dvc.Formula2) > 0 {
			f2 = dvc.Formula2[0]
		}
		// Ignoring error for now
		_ = dv.SetRange(f1, f2, getValidationType(dvc.Type), getOperatorType(dvc.Operator))
	case "custom":
		var f1, f2 string
		if len(dvc.Formula1) > 0 {
			f1 = dvc.Formula1[0]
		}
		if len(dvc.Formula2) > 0 {
			f2 = dvc.Formula2[0]
		}
		// Ignoring error for now
		dv.Formula1 = f1
		dv.Formula2 = f2
	}

	err := cb.rowBuilder.sheetBuilder.workbookBuilder.file.AddDataValidation(cb.rowBuilder.sheetBuilder.sheetName, dv)
	if err != nil {
		// Handle error, maybe log it or return it
	}
	return cb
}

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

// AddDataValidation adds data validation to the cell
func (cb *CellBuilder) AddDataValidation(validation DataValidation) *CellBuilder {
	if validation.Type == "" {
		return cb
	}

	// Create excelize data validation
	dv := excelize.NewDataValidation(true)
	dv.SetSqref(cb.cellRef)

	// Set validation type and operator
	dv.SetDropList([]string{}) // Initialize for list type
	if validation.Type == "list" {
		if validation.Formula1 != "" {
			// Split comma-separated values for dropdown
			options := strings.Split(validation.Formula1, ",")
			// Trim spaces from each option
			for i, opt := range options {
				options[i] = strings.TrimSpace(opt)
			}
			dv.SetDropList(options)
		}
	} else {
		// Set validation type and operator for non-list types
		dv.SetRange(validation.Formula1, validation.Formula2, getValidationType(validation.Type), getOperatorType(validation.Operator))
	}

	// Set options
	dv.AllowBlank = validation.AllowBlank
	dv.ShowDropDown = validation.ShowDropDown
	dv.ShowInputMessage = validation.ShowInputMessage
	dv.ShowErrorMessage = validation.ShowErrorMessage

	// Set messages with pointers
	if validation.ErrorTitle != "" {
		dv.ErrorTitle = &validation.ErrorTitle
	}
	if validation.ErrorMessage != "" {
		dv.Error = &validation.ErrorMessage
	}
	if validation.PromptTitle != "" {
		dv.PromptTitle = &validation.PromptTitle
	}
	if validation.PromptMessage != "" {
		dv.Prompt = &validation.PromptMessage
	}

	// Add validation to sheet
	err := cb.rowBuilder.sheetBuilder.workbookBuilder.file.AddDataValidation(
		cb.rowBuilder.sheetBuilder.sheetName,
		dv,
	)
	if err != nil {
		return nil
	}

	return cb
}

// WithValue sets the value of the cell
func (cb *CellBuilder) WithValue(value interface{}) *CellBuilder {
	err := cb.rowBuilder.sheetBuilder.workbookBuilder.file.SetCellValue(
		cb.rowBuilder.sheetBuilder.sheetName,
		cb.cellRef,
		value,
	)
	if err != nil {
		return nil
	}
	return cb
}

// WithStyle is an alias for SetStyle.
func (cb *CellBuilder) WithStyle(config StyleConfig) *CellBuilder {
	return cb.SetStyle(config)
}

// Done returns to the RowBuilder
func (cb *CellBuilder) Done() *RowBuilder {
	return cb.rowBuilder
}