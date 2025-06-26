package excelbuilder_test

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
)

func TestRowBuilder_SetHeight_Validation(t *testing.T) {
	rowBuilder := excelbuilder.New().NewWorkbook().AddSheet("Sheet1").AddRow()
	assert.NotNil(t, rowBuilder)

	testCases := []struct {
		name        string
		height      float64
		shouldBeNil bool
		message     string
	}{
		{"Valid Height", 50.0, false, "Should allow valid row height"},
		{"Maximum Height", 409, false, "Should allow maximum row height"},
		{"Zero Height", 0, true, "Should not allow zero row height"},
		{"Negative Height", -20.0, true, "Should not allow negative row height"},
		{"Too Large Height", 409.1, true, "Should not allow height > 409"},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			result := rowBuilder.SetHeight(tc.height)
			if tc.shouldBeNil {
				assert.Nil(t, result, tc.message)
			} else {
				assert.NotNil(t, result, tc.message)
				assert.Same(t, rowBuilder, result, "Should return the same builder instance on success")
			}
		})
	}
}
