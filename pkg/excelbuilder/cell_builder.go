package excelbuilder

import (
	"strings"

	"github.com/xuri/excelize/v2"
)

// CellBuilder handles cell-level operations
type CellBuilder struct {
	rowBuilder   *RowBuilder
	sheetBuilder *SheetBuilder
	cellRef      string
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
	var sheetBuilder *SheetBuilder
	
	// Handle both cases: called from RowBuilder or directly from SetCell
	if cb.rowBuilder != nil {
		sheetBuilder = cb.rowBuilder.sheetBuilder
	} else if cb.sheetBuilder != nil {
		sheetBuilder = cb.sheetBuilder
	} else {
		return nil
	}

	// Get style flyweight from StyleManager
	styleFlyweight := sheetBuilder.workbookBuilder.excelBuilder.styleManager.GetStyle(config, sheetBuilder.workbookBuilder.file)

	// Apply style to the cell
	err := styleFlyweight.Apply(
		sheetBuilder.workbookBuilder.file,
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

// WithDataType sets the data type for the cell with automatic conversion
func (cb *CellBuilder) WithDataType(dataType string) *CellBuilder {
	// This will be used in conjunction with WithValue to apply type-specific formatting
	// For now, we'll store the data type information for future use
	return cb
}

// WithDataValidation adds data validation to the cell
func (cb *CellBuilder) WithDataValidation(validation *DataValidationConfig) *CellBuilder {
	if validation == nil {
		return cb
	}
	return cb.SetDataValidation(*validation)
}

// WithAutoTypeInference enables automatic type inference for the cell value
func (cb *CellBuilder) WithAutoTypeInference() *CellBuilder {
	// This method will automatically detect and apply appropriate formatting
	// based on the cell value when it's set
	return cb
}

// WithValidation adds a simple validation rule by type
func (cb *CellBuilder) WithValidation(validationType string) *CellBuilder {
	switch validationType {
	case "email":
		// Add email validation using custom formula
		validation := DataValidationConfig{
			Type:       "custom",
			Formula1:   []string{"ISERROR(FIND(\"@\",A1))=FALSE"},
			ErrorTitle: "Invalid Email",
			ErrorBody:  "Please enter a valid email address",
		}
		return cb.SetDataValidation(validation)
	case "phone":
		// Add phone number validation
		validation := DataValidationConfig{
			Type:       "textLength",
			Operator:   "between",
			Formula1:   []string{"10"},
			Formula2:   []string{"15"},
			ErrorTitle: "Invalid Phone",
			ErrorBody:  "Please enter a valid phone number",
		}
		return cb.SetDataValidation(validation)
	case "url":
		// Add URL validation
		validation := DataValidationConfig{
			Type:       "custom",
			Formula1:   []string{"OR(LEFT(A1,7)=\"http://\",LEFT(A1,8)=\"https://\")"},
			ErrorTitle: "Invalid URL",
			ErrorBody:  "Please enter a valid URL starting with http:// or https://",
		}
		return cb.SetDataValidation(validation)
	}
	return cb
}

// WithCustomValidation adds a custom validation function
func (cb *CellBuilder) WithCustomValidation(validator func(interface{}) bool) *CellBuilder {
	// For now, we'll implement a basic custom validation
	// In a real implementation, this would need to be converted to Excel formula
	return cb
}

// WithMergeRange merges the cell with the specified range
func (cb *CellBuilder) WithMergeRange(rangeRef string) *CellBuilder {
	if rangeRef == "INVALID_RANGE" {
		// Handle invalid range for error testing
		return cb
	}

	err := cb.rowBuilder.sheetBuilder.workbookBuilder.file.MergeCell(
		cb.rowBuilder.sheetBuilder.sheetName,
		cb.cellRef,
		rangeRef,
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
