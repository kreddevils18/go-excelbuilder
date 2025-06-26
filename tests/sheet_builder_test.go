package excelbuilder_test

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
)

func TestSheetBuilder_AddSheet_Validation(t *testing.T) {
	builder := excelbuilder.New()
	wb := builder.NewWorkbook()

	testCases := []struct {
		name        string
		sheetName   string
		shouldBeNil bool
		message     string
	}{
		{"Valid Name", "Report", false, "Should allow valid sheet name"},
		{"Empty Name", "", true, "Should not allow empty sheet name"},
		{"Too Long Name", "ThisIsAVeryLongSheetNameThatExceedsTheThirtyOneCharacterLimit", true, "Should not allow sheet name > 31 chars"},
		{"Invalid Char [", "Sheet[1]", true, "Should not allow '[' in sheet name"},
		{"Invalid Char ]", "Sheet]1", true, "Should not allow ']' in sheet name"},
		{"Invalid Char *", "Sheet*1", true, "Should not allow '*' in sheet name"},
		{"Invalid Char ?", "Sheet?1", true, "Should not allow '?' in sheet name"},
		{"Invalid Char /", "Sheet/1", true, "Should not allow '/' in sheet name"},
		{"Invalid Char \\", "Sheet\\1", true, "Should not allow '\\' in sheet name"},
		{"Reserved Name History", "History", true, "Should not allow reserved name 'History'"},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			sheetBuilder := wb.AddSheet(tc.sheetName)
			if tc.shouldBeNil {
				assert.Nil(t, sheetBuilder, tc.message)
			} else {
				assert.NotNil(t, sheetBuilder, tc.message)
			}
		})
	}
}

func TestSheetBuilder_SetColumnWidth_Validation(t *testing.T) {
	sheetBuilder := excelbuilder.New().NewWorkbook().AddSheet("Sheet1")
	assert.NotNil(t, sheetBuilder)

	testCases := []struct {
		name        string
		col         string
		width       float64
		shouldBeNil bool
		message     string
	}{
		{"Valid Width", "A", 15.0, false, "Should allow valid column width"},
		{"Zero Width", "B", 0, false, "Should allow zero width to hide column"},
		{"Negative Width", "C", -10.0, true, "Should not allow negative column width"},
		{"Too Large Width", "D", 300.0, true, "Should not allow width > 255"},
		{"Invalid Column Name", "1A", 15.0, true, "Should not allow invalid column name"},
		{"Empty Column Name", "", 15.0, true, "Should not allow empty column name"},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			result := sheetBuilder.SetColumnWidth(tc.col, tc.width)
			if tc.shouldBeNil {
				assert.Nil(t, result, tc.message)
			} else {
				assert.NotNil(t, result, tc.message)
				assert.Same(t, sheetBuilder, result, "Should return the same builder instance on success")
			}
		})
	}
}

func TestSheetBuilder_MergeCell_Validation(t *testing.T) {
	sheetBuilder := excelbuilder.New().NewWorkbook().AddSheet("Sheet1")
	assert.NotNil(t, sheetBuilder)

	testCases := []struct {
		name        string
		cellRange   string
		shouldBeNil bool
		message     string
	}{
		{"Valid Merge", "A1:C1", false, "Should allow valid merge range"},
		{"Single Cell Merge", "A1:A1", false, "Should handle single cell merge gracefully"},
		{"Invalid Range", "A1C1", true, "Should not allow invalid range format"},
		{"Empty Range", "", true, "Should not allow empty range"},
		{"Reverse Range", "C1:A1", true, "Should not allow reverse merge range"},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			result := sheetBuilder.MergeCell(tc.cellRange)
			if tc.shouldBeNil {
				assert.Nil(t, result, tc.message)
			} else {
				assert.NotNil(t, result, tc.message)
				assert.Same(t, sheetBuilder, result, "Should return the same builder instance on success")
			}
		})
	}
}

func TestSheetBuilder_FreezePanes(t *testing.T) {
	// Setup
	builder := excelbuilder.New()
	wb := builder.NewWorkbook()
	sheet := wb.AddSheet("FreezeSheet")

	// Action
	sheet.FreezePanes(1, 2) // Freeze 1 column and 2 rows

	// Verification
	file := wb.Build()
	panes, err := file.GetPanes("FreezeSheet")
	assert.NoError(t, err)

	assert.True(t, panes.Freeze, "Panes should be set to freeze")
	assert.Equal(t, 1, panes.XSplit, "XSplit should be 1")
	assert.Equal(t, 2, panes.YSplit, "YSplit should be 2")
	assert.Equal(t, "B3", panes.TopLeftCell, "TopLeftCell should be calculated correctly")
}

func TestSheetBuilder_MergeCell_Integration(t *testing.T) {
	// Setup
	builder := excelbuilder.New()
	wb := builder.NewWorkbook()
	sheet := wb.AddSheet("MergeSheet")

	// Action
	sheet.AddRow().AddCell("Merged Header").WithMergeRange("C1")
	sheet.MergeCell("A2:B3")

	// Verification
	file := wb.Build()
	mergedCells, err := file.GetMergeCells("MergeSheet")
	assert.NoError(t, err)

	assert.Len(t, mergedCells, 2, "Expected two merged regions")

	// Check the first merge from WithMergeRange
	firstMergeStartAxis := mergedCells[0].GetStartAxis()
	assert.Equal(t, "A1", firstMergeStartAxis, "First merge start axis should be A1")
	assert.Equal(t, "C1", mergedCells[0].GetEndAxis(), "First merge end axis should be C1")
	val, err := file.GetCellValue("MergeSheet", firstMergeStartAxis)
	assert.NoError(t, err)
	assert.Equal(t, "Merged Header", val, "Cell value for first merge should be set correctly")

	// Check the second merge from MergeCell
	assert.Equal(t, "A2", mergedCells[1].GetStartAxis(), "Second merge start axis should be A2")
	assert.Equal(t, "B3", mergedCells[1].GetEndAxis(), "Second merge end axis should be B3")
}
