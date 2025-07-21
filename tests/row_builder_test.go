package excelbuilder_test

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
)

func TestRowBuilder_SetHeight_Validation(t *testing.T) {
	testCases := []struct {
		name        string
		height      float64
		shouldError bool
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
			builder := excelbuilder.New().WithErrorCollection(true)
			rowBuilder := builder.NewWorkbook().AddSheet("Sheet1").AddRow()
			
			result := rowBuilder.SetHeight(tc.height)
			
			// Should never return nil - that was the dangerous behavior we fixed
			assert.NotNil(t, result, "Builder should never return nil")
			assert.Same(t, rowBuilder, result, "Should return the same builder instance")
			
			if tc.shouldError {
				assert.True(t, builder.HasErrors(), tc.message)
				errors := builder.GetCollectedErrors()
				assert.Greater(t, len(errors), 0, "Should have collected errors for invalid input")
			} else {
				assert.False(t, builder.HasErrors(), tc.message)
			}
		})
	}
}
