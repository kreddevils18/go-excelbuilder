package excelbuilder

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"os"
	"sort"
	"strconv"
)

// NewImportHelper creates a new ImportHelper instance
func NewImportHelper() *ImportHelper {
	return &ImportHelper{
		sheetName: "ImportedData",
	}
}

// NewExportHelper creates a new ExportHelper instance
func NewExportHelper() *ExportHelper {
	return &ExportHelper{}
}

// FromCSV imports data from a CSV file
func (ih *ImportHelper) FromCSV(filename string) *ImportHelper {
	file, err := os.Open(filename)
	if err != nil {
		panic(fmt.Sprintf("failed to open CSV file: %v", err))
	}
	defer file.Close()

	// Set sheet name for CSV import
	ih.sheetName = "ImportedData"

	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		panic(fmt.Sprintf("failed to read CSV file: %v", err))
	}

	ih.data = records
	return ih
}

// FromCSVWithOptions imports data from a CSV file with custom options
func (ih *ImportHelper) FromCSVWithOptions(filename string, options CSVOptions) *ImportHelper {
	file, err := os.Open(filename)
	if err != nil {
		panic(fmt.Sprintf("failed to open CSV file: %v", err))
	}
	defer file.Close()

	// Set sheet name for CSV with options
	ih.sheetName = "CustomCSV"

	reader := csv.NewReader(file)
	if options.Delimiter != "" {
		reader.Comma = rune(options.Delimiter[0])
	}
	// Note: csv.Reader doesn't have Quote field in standard library
	if options.Comment != 0 {
		reader.Comment = options.Comment
	}

	records, err := reader.ReadAll()
	if err != nil {
		panic(fmt.Sprintf("failed to read CSV file: %v", err))
	}

	// Skip rows if specified
	if options.SkipRows > 0 && len(records) > options.SkipRows {
		records = records[options.SkipRows:]
	}

	ih.data = records
	return ih
}

// FromJSON imports data from a JSON file
func (ih *ImportHelper) FromJSON(filename string) *ImportHelper {
	file, err := os.Open(filename)
	if err != nil {
		panic(fmt.Sprintf("failed to open JSON file: %v", err))
	}
	defer file.Close()

	var data interface{}
	decoder := json.NewDecoder(file)
	if err := decoder.Decode(&data); err != nil {
		panic(fmt.Sprintf("failed to decode JSON: %v", err))
	}

	// Set sheet name for JSON imports
	ih.sheetName = "JSONData"
	ih.data = data
	return ih
}

// WithFlattenOptions applies flattening options to JSON data
func (ih *ImportHelper) WithFlattenOptions(options FlattenOptions) *ImportHelper {
	if ih.data == nil {
		return ih
	}

	// If data is a wrapper object with an array, extract the array first
	if wrapper, ok := ih.data.(map[string]interface{}); ok {
		for _, value := range wrapper {
			if arr, ok := value.([]interface{}); ok {
				// Found an array, use it as the main data for flattening
				ih.data = arr
				break
			}
		}
	}

	// Convert to flattened structure
	flattened := ih.flattenData(ih.data, options)
	ih.data = flattened
	return ih
}

// ToExcel converts imported data to Excel format
func (ih *ImportHelper) ToExcel() *WorkbookBuilder {
	eb := New()
	workbook := eb.NewWorkbook()
	sheet := workbook.AddSheet(ih.sheetName)

	switch data := ih.data.(type) {
	case [][]string:
		// CSV data
		for _, row := range data {
			excelRow := sheet.AddRow()
			for _, cell := range row {
				excelRow.AddCell(cell)
			}
		}
	case []interface{}:
		// JSON array
		if len(data) > 0 {
			// Create headers from first object
			if firstObj, ok := data[0].(map[string]interface{}); ok {
				headerRow := sheet.AddRow()
				var keys []string
				for key := range firstObj {
					keys = append(keys, key)
				}
				sort.Strings(keys)
				for _, key := range keys {
					headerRow.AddCell(key)
				}

				// Add data rows
				for _, item := range data {
					if obj, ok := item.(map[string]interface{}); ok {
						dataRow := sheet.AddRow()
						for _, key := range keys {
							value := obj[key]
							dataRow.AddCell(value)
						}
					}
				}
			}
		}
	case map[string]interface{}:
		// Check if this is a wrapper object with an array inside
		var arrayFound bool
		for _, value := range data {
			if arr, ok := value.([]interface{}); ok {
				// Found an array, use it as the main data
				if len(arr) > 0 {
					if firstObj, ok := arr[0].(map[string]interface{}); ok {
						headerRow := sheet.AddRow()
						var keys []string
						for key := range firstObj {
							keys = append(keys, key)
						}
						sort.Strings(keys)
						for _, key := range keys {
							headerRow.AddCell(key)
						}

						// Add data rows
						for _, item := range arr {
							if obj, ok := item.(map[string]interface{}); ok {
								dataRow := sheet.AddRow()
								for _, key := range keys {
									value := obj[key]
									dataRow.AddCell(value)
								}
							}
						}
					}
					arrayFound = true
					break
				}
			}
		}
		
		if !arrayFound {
			// Single JSON object
			headerRow := sheet.AddRow()
			dataRow := sheet.AddRow()
			for key, value := range data {
				headerRow.AddCell(key)
				dataRow.AddCell(value)
			}
		}
	}

	return workbook
}

// FromExcel sets the Excel file for export
func (eh *ExportHelper) FromExcel(workbook *WorkbookBuilder) *ExportHelper {
	eh.file = workbook.Build()
	return eh
}

// ToCSV exports Excel data to CSV format
func (eh *ExportHelper) ToCSV(filename string) *ExportHelper {
	if eh.file == nil {
		panic("no Excel file set for export")
	}

	outFile, err := os.Create(filename)
	if err != nil {
		panic(fmt.Sprintf("failed to create CSV file: %v", err))
	}
	defer outFile.Close()

	writer := csv.NewWriter(outFile)
	defer writer.Flush()

	// Get sheet names and find the one with data
	sheetNames := eh.file.GetSheetList()
	if len(sheetNames) > 0 {
		// Try to find a sheet with data, skip default empty sheets
		var targetSheet string
		for _, sheetName := range sheetNames {
			if sheetName != "Sheet1" { // Skip default empty sheet
				targetSheet = sheetName
				break
			}
		}
		// If no non-default sheet found, use the first one
		if targetSheet == "" {
			targetSheet = sheetNames[0]
		}
		
		rows, err := eh.file.GetRows(targetSheet)
		if err == nil {
			for _, row := range rows {
				writer.Write(row)
			}
		}
	}

	return eh
}

// ToCSVWithOptions exports Excel data to CSV with custom options
func (eh *ExportHelper) ToCSVWithOptions(filename string, options CSVOptions) *ExportHelper {
	if eh.file == nil {
		panic("no Excel file set for export")
	}

	outFile, err := os.Create(filename)
	if err != nil {
		panic(fmt.Sprintf("failed to create CSV file: %v", err))
	}
	defer outFile.Close()

	writer := csv.NewWriter(outFile)
	if options.Delimiter != "" {
		writer.Comma = rune(options.Delimiter[0])
	}
	defer writer.Flush()

	// Get sheet names and find the best sheet to export
	sheetNames := eh.file.GetSheetList()
	var selectedSheet string
	if len(sheetNames) > 0 {
		// Prefer sheets other than "Sheet1" (which is often empty)
		for _, name := range sheetNames {
			if name != "Sheet1" {
				selectedSheet = name
				break
			}
		}
		// If no other sheet found, use the first one
		if selectedSheet == "" {
			selectedSheet = sheetNames[0]
		}

		rows, err := eh.file.GetRows(selectedSheet)
		if err == nil {
			for _, row := range rows {
				writer.Write(row)
			}
		}
	}

	return eh
}

// ToJSON exports Excel data to JSON format
func (eh *ExportHelper) ToJSON(filename string) *ExportHelper {
	if eh.file == nil {
		panic("no Excel file set for export")
	}

	outFile, err := os.Create(filename)
	if err != nil {
		panic(fmt.Sprintf("failed to create JSON file: %v", err))
	}
	defer outFile.Close()

	var data []map[string]interface{}

	// Get sheet names and find the best sheet to export
	sheetNames := eh.file.GetSheetList()
	var selectedSheet string
	if len(sheetNames) > 0 {
		// Prefer sheets other than "Sheet1" (which is often empty)
		for _, name := range sheetNames {
			if name != "Sheet1" {
				selectedSheet = name
				break
			}
		}
		// If no other sheet found, use the first one
		if selectedSheet == "" {
			selectedSheet = sheetNames[0]
		}

		rows, err := eh.file.GetRows(selectedSheet)
		if err == nil && len(rows) > 0 {
			// Use first row as headers
			headers := rows[0]

			// Process data rows
			for i := 1; i < len(rows); i++ {
				row := rows[i]
				rowData := make(map[string]interface{})
				for j, cellValue := range row {
					if j < len(headers) {
						rowData[headers[j]] = cellValue
					}
				}
				data = append(data, rowData)
			}
		}
	}

	encoder := json.NewEncoder(outFile)
	encoder.SetIndent("", "  ")
	if err := encoder.Encode(data); err != nil {
		panic(fmt.Sprintf("failed to encode JSON: %v", err))
	}

	return eh
}

// flattenData flattens nested data structures
func (ih *ImportHelper) flattenData(data interface{}, options FlattenOptions) interface{} {
	separator := options.Separator
	if separator == "" {
		separator = "."
	}

	switch v := data.(type) {
	case []interface{}:
		var result []interface{}
		for _, item := range v {
			result = append(result, ih.flattenObject(item, "", separator, options.MaxDepth, 0))
		}
		return result
	case map[string]interface{}:
		return ih.flattenObject(v, "", separator, options.MaxDepth, 0)
	default:
		return data
	}
}

// flattenObject flattens a single object
func (ih *ImportHelper) flattenObject(obj interface{}, prefix, separator string, maxDepth, currentDepth int) interface{} {
	if maxDepth > 0 && currentDepth >= maxDepth {
		return obj
	}

	switch v := obj.(type) {
	case map[string]interface{}:
		result := make(map[string]interface{})
		for key, value := range v {
			newKey := key
			if prefix != "" {
				newKey = prefix + separator + key
			}

			switch nested := value.(type) {
			case map[string]interface{}, []interface{}:
				flattened := ih.flattenObject(nested, newKey, separator, maxDepth, currentDepth+1)
				if flatMap, ok := flattened.(map[string]interface{}); ok {
					for k, v := range flatMap {
						result[k] = v
					}
				} else {
					result[newKey] = flattened
				}
			default:
				result[newKey] = value
			}
		}
		return result
	case []interface{}:
		var result []interface{}
		for i, item := range v {
			newPrefix := prefix + separator + strconv.Itoa(i)
			if prefix == "" {
				newPrefix = strconv.Itoa(i)
			}
			result = append(result, ih.flattenObject(item, newPrefix, separator, maxDepth, currentDepth+1))
		}
		return result
	default:
		return obj
	}
}
