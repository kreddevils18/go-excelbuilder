package excelbuilder

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// PivotTableBuilder handles pivot table creation and configuration
type PivotTableBuilder struct {
	sheetBuilder *SheetBuilder
	file         *excelize.File
	config       PivotTableConfig
}

// NewPivotTableBuilder creates a new PivotTableBuilder instance
func NewPivotTableBuilder(sheetBuilder *SheetBuilder, targetSheet, sourceRange string) *PivotTableBuilder {
	return &PivotTableBuilder{
		sheetBuilder: sheetBuilder,
		file:         sheetBuilder.workbookBuilder.file,
		config: PivotTableConfig{
			SourceSheet:           sheetBuilder.sheetName,
			SourceRange:           sourceRange,
			TargetSheet:          targetSheet,
			TargetCell:           "A1",
			RowFields:            []PivotField{},
			ColumnFields:         []PivotField{},
			ValueFields:          []PivotField{},
			FilterFields:         []PivotField{},
			Style:                "PivotStyleLight16",
			ShowRowGrandTotals:   true,
			ShowColumnGrandTotals: true,
			Compact:              true,
			Outline:              true,
			Subtotals:            true,
		},
	}
}

// SetName sets the name of the pivot table
func (ptb *PivotTableBuilder) SetName(name string) *PivotTableBuilder {
	ptb.config.Name = name
	return ptb
}

// SetDataSource sets the data source for the pivot table
func (ptb *PivotTableBuilder) SetDataSource(sourceSheet, sourceRange string) *PivotTableBuilder {
	ptb.config.SourceSheet = sourceSheet
	ptb.config.SourceRange = sourceRange
	return ptb
}

// SetTargetCell sets the target cell where the pivot table will be placed
func (ptb *PivotTableBuilder) SetTargetCell(cell string) *PivotTableBuilder {
	ptb.config.TargetCell = cell
	return ptb
}

// AddRowField adds a field to the row area of the pivot table
func (ptb *PivotTableBuilder) AddRowField(fieldName string) *PivotTableBuilder {
	field := PivotField{
		Name:     fieldName,
		Function: "", // Row fields don't have aggregation functions
	}
	ptb.config.RowFields = append(ptb.config.RowFields, field)
	return ptb
}

// AddColumnField adds a field to the column area of the pivot table
func (ptb *PivotTableBuilder) AddColumnField(fieldName string) *PivotTableBuilder {
	field := PivotField{
		Name:     fieldName,
		Function: "", // Column fields don't have aggregation functions
	}
	ptb.config.ColumnFields = append(ptb.config.ColumnFields, field)
	return ptb
}

// AddValueField adds a field to the value area of the pivot table with aggregation function
func (ptb *PivotTableBuilder) AddValueField(fieldName, function string) *PivotTableBuilder {
	field := PivotField{
		Name:     fieldName,
		Function: function,
	}
	ptb.config.ValueFields = append(ptb.config.ValueFields, field)
	return ptb
}

// AddFilterField adds a field to the filter area of the pivot table
func (ptb *PivotTableBuilder) AddFilterField(fieldName string) *PivotTableBuilder {
	field := PivotField{
		Name:     fieldName,
		Function: "", // Filter fields don't have aggregation functions
	}
	ptb.config.FilterFields = append(ptb.config.FilterFields, field)
	return ptb
}

// SetStyle sets the style of the pivot table
func (ptb *PivotTableBuilder) SetStyle(style string) *PivotTableBuilder {
	ptb.config.Style = style
	return ptb
}

// ShowRowGrandTotals sets whether to show row grand totals
func (ptb *PivotTableBuilder) ShowRowGrandTotals(show bool) *PivotTableBuilder {
	ptb.config.ShowRowGrandTotals = show
	return ptb
}

// ShowColumnGrandTotals sets whether to show column grand totals
func (ptb *PivotTableBuilder) ShowColumnGrandTotals(show bool) *PivotTableBuilder {
	ptb.config.ShowColumnGrandTotals = show
	return ptb
}

// SetCompact sets whether to use compact layout
func (ptb *PivotTableBuilder) SetCompact(compact bool) *PivotTableBuilder {
	ptb.config.Compact = compact
	return ptb
}

// SetOutline sets whether to use outline layout
func (ptb *PivotTableBuilder) SetOutline(outline bool) *PivotTableBuilder {
	ptb.config.Outline = outline
	return ptb
}

// SetSubtotals sets whether to show subtotals
func (ptb *PivotTableBuilder) SetSubtotals(subtotals bool) *PivotTableBuilder {
	ptb.config.Subtotals = subtotals
	return ptb
}

// GetConfig returns the current pivot table configuration
func (ptb *PivotTableBuilder) GetConfig() PivotTableConfig {
	return ptb.config
}

// Build creates the pivot table in the Excel file
func (ptb *PivotTableBuilder) Build() error {
	// Create the target sheet if it doesn't exist
	if ptb.config.TargetSheet != "" {
		// Check if sheet exists
		sheets := ptb.file.GetSheetList()
		sheetExists := false
		for _, sheet := range sheets {
			if sheet == ptb.config.TargetSheet {
				sheetExists = true
				break
			}
		}
		
		if !sheetExists {
			_, err := ptb.file.NewSheet(ptb.config.TargetSheet)
			if err != nil {
				return fmt.Errorf("failed to create target sheet %s: %w", ptb.config.TargetSheet, err)
			}
		}
	}

	// Build the pivot table options
	options := ptb.buildPivotTableOptions()

	// Create the pivot table using excelize
	err := ptb.file.AddPivotTable(&options)
	if err != nil {
		return fmt.Errorf("failed to create pivot table: %w", err)
	}

	return nil
}

// buildPivotTableOptions converts the configuration to excelize PivotTableOptions
func (ptb *PivotTableBuilder) buildPivotTableOptions() excelize.PivotTableOptions {
	options := excelize.PivotTableOptions{
		DataRange:       fmt.Sprintf("%s!%s", ptb.config.SourceSheet, ptb.config.SourceRange),
		PivotTableRange: fmt.Sprintf("%s!%s:J20", ptb.config.TargetSheet, ptb.config.TargetCell),
		Rows:            ptb.buildFieldOptions(ptb.config.RowFields),
		Columns:         ptb.buildFieldOptions(ptb.config.ColumnFields),
		Data:            ptb.buildDataFieldOptions(ptb.config.ValueFields),
		Filter:          ptb.buildFieldOptions(ptb.config.FilterFields),
		RowGrandTotals:  ptb.config.ShowRowGrandTotals,
		ColGrandTotals:  ptb.config.ShowColumnGrandTotals,
		ShowDrill:       true,
		UseAutoFormatting: true,
		PageOverThenDown:  true,
		MergeItem:         true,
		CompactData:       ptb.config.Compact,
		ShowError:         true,
		PivotTableStyleName: ptb.config.Style,
	}

	// Set name if provided
	if ptb.config.Name != "" {
		options.Name = ptb.config.Name
	}

	return options
}

// buildFieldOptions converts PivotField slice to excelize PivotTableField slice
func (ptb *PivotTableBuilder) buildFieldOptions(fields []PivotField) []excelize.PivotTableField {
	result := make([]excelize.PivotTableField, len(fields))
	for i, field := range fields {
		result[i] = excelize.PivotTableField{
			Compact:         ptb.config.Compact,
			Data:            field.Name,
			Name:            field.Name,
			Outline:         ptb.config.Outline,
			DefaultSubtotal: true,
		}
	}
	return result
}

// buildDataFieldOptions converts value PivotField slice to excelize PivotTableField slice
func (ptb *PivotTableBuilder) buildDataFieldOptions(fields []PivotField) []excelize.PivotTableField {
	result := make([]excelize.PivotTableField, len(fields))
	for i, field := range fields {
		result[i] = excelize.PivotTableField{
			Data:         field.Name,
			Name:         fmt.Sprintf("%s of %s", field.Function, field.Name),
			Subtotal:     field.Function,
		}
	}
	return result
}