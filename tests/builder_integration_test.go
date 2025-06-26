package excelbuilder_test

import (
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
)

func TestBuilder_EndToEnd_Simple(t *testing.T) {
	// 1. Build the Excel file using the fluent API
	builder := excelbuilder.New()
	file := builder.
		NewWorkbook().
		AddSheet("Sheet1").
		AddRow().AddCells("Name", "Age").
		Done().
		AddRow().AddCells("Alice", 30).
		Done().
		AddRow().AddCells("Bob", 25).
		Done().
		Build()

	assert.NotNil(t, file)

	// 2. Read the data back from the generated file to verify correctness
	rows, err := file.GetRows("Sheet1")
	assert.NoError(t, err)

	// 3. Assert the contents are as expected
	assert.Len(t, rows, 3, "Expected 3 rows in Sheet1")

	// Verify header row
	assert.Equal(t, []string{"Name", "Age"}, rows[0])

	// Verify data rows
	assert.Equal(t, []string{"Alice", "30"}, rows[1])
	assert.Equal(t, []string{"Bob", "25"}, rows[2])
}
