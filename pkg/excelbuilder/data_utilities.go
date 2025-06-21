package excelbuilder

import (
	"fmt"
	"math"
	"reflect"
	"strconv"
	"strings"
	"time"
)

// DataTypeHandler provides comprehensive data type handling utilities
type DataTypeHandler struct{}

// NewDataTypeHandler creates a new data type handler
func NewDataTypeHandler() *DataTypeHandler {
	return &DataTypeHandler{}
}

// ConvertToExcelValue converts any Go value to an Excel-compatible value
func (dth *DataTypeHandler) ConvertToExcelValue(value interface{}) interface{} {
	if value == nil {
		return ""
	}

	v := reflect.ValueOf(value)
	switch v.Kind() {
	case reflect.String:
		return value
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return v.Int()
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		return v.Uint()
	case reflect.Float32, reflect.Float64:
		f := v.Float()
		if math.IsInf(f, 1) {
			return "+Inf"
		}
		if math.IsInf(f, -1) {
			return "-Inf"
		}
		if math.IsNaN(f) {
			return "NaN"
		}
		return f
	case reflect.Bool:
		if v.Bool() {
			return "true"
		}
		return "false"
	case reflect.Slice, reflect.Array:
		return dth.flattenSlice(value)
	case reflect.Map:
		return dth.flattenMap(value)
	case reflect.Struct:
		return dth.flattenStruct(value)
	default:
		// Handle time.Time specially
		if t, ok := value.(time.Time); ok {
			return t.Format("2006-01-02 15:04:05")
		}
		return fmt.Sprintf("%v", value)
	}
}

// flattenSlice converts a slice to a string representation
func (dth *DataTypeHandler) flattenSlice(value interface{}) string {
	v := reflect.ValueOf(value)
	if v.Kind() != reflect.Slice && v.Kind() != reflect.Array {
		return fmt.Sprintf("%v", value)
	}

	var parts []string
	for i := 0; i < v.Len(); i++ {
		elem := v.Index(i).Interface()
		parts = append(parts, fmt.Sprintf("%v", dth.ConvertToExcelValue(elem)))
	}
	return "[" + strings.Join(parts, ", ") + "]"
}

// flattenMap converts a map to a string representation
func (dth *DataTypeHandler) flattenMap(value interface{}) string {
	v := reflect.ValueOf(value)
	if v.Kind() != reflect.Map {
		return fmt.Sprintf("%v", value)
	}

	var parts []string
	for _, key := range v.MapKeys() {
		mapValue := v.MapIndex(key).Interface()
		parts = append(parts, fmt.Sprintf("%v: %v", key, dth.ConvertToExcelValue(mapValue)))
	}
	return "{" + strings.Join(parts, ", ") + "}"
}

// flattenStruct converts a struct to a string representation
func (dth *DataTypeHandler) flattenStruct(value interface{}) string {
	v := reflect.ValueOf(value)
	t := reflect.TypeOf(value)
	if v.Kind() != reflect.Struct {
		return fmt.Sprintf("%v", value)
	}

	var parts []string
	for i := 0; i < v.NumField(); i++ {
		field := t.Field(i)
		fieldValue := v.Field(i).Interface()
		parts = append(parts, fmt.Sprintf("%s: %v", field.Name, dth.ConvertToExcelValue(fieldValue)))
	}
	return "{" + strings.Join(parts, ", ") + "}"
}

// InferDataType attempts to infer the data type from a string value
func (dth *DataTypeHandler) InferDataType(value string) (interface{}, string) {
	value = strings.TrimSpace(value)

	if value == "" {
		return "", "string"
	}

	// Try boolean
	if strings.ToLower(value) == "true" {
		return true, "boolean"
	}
	if strings.ToLower(value) == "false" {
		return false, "boolean"
	}

	// Try integer
	if intVal, err := strconv.ParseInt(value, 10, 64); err == nil {
		return intVal, "integer"
	}

	// Try float
	if floatVal, err := strconv.ParseFloat(value, 64); err == nil {
		return floatVal, "float"
	}

	// Try date
	if dateVal, err := time.Parse("2006-01-02", value); err == nil {
		return dateVal, "date"
	}
	if dateVal, err := time.Parse("2006-01-02 15:04:05", value); err == nil {
		return dateVal, "datetime"
	}

	// Try currency
	if strings.HasPrefix(value, "$") {
		currencyStr := strings.ReplaceAll(strings.TrimPrefix(value, "$"), ",", "")
		if currencyVal, err := strconv.ParseFloat(currencyStr, 64); err == nil {
			return currencyVal, "currency"
		}
	}

	// Try percentage
	if strings.HasSuffix(value, "%") {
		percentStr := strings.TrimSuffix(value, "%")
		if percentVal, err := strconv.ParseFloat(percentStr, 64); err == nil {
			return percentVal / 100, "percentage"
		}
	}

	// Default to text
	return value, "text"
}

// ValidateDataType validates if a value matches the expected data type
func (dth *DataTypeHandler) ValidateDataType(value interface{}, expectedType string) bool {
	if value == nil {
		return true // nil is valid for any type
	}

	v := reflect.ValueOf(value)
	switch expectedType {
	case "string", "text":
		return v.Kind() == reflect.String
	case "int", "integer":
		return v.Kind() >= reflect.Int && v.Kind() <= reflect.Int64
	case "uint":
		return v.Kind() >= reflect.Uint && v.Kind() <= reflect.Uint64
	case "float":
		return v.Kind() == reflect.Float32 || v.Kind() == reflect.Float64
	case "bool", "boolean":
		return v.Kind() == reflect.Bool
	case "time", "date", "datetime":
		_, ok := value.(time.Time)
		return ok
	default:
		return true // Unknown type, assume valid
	}
}

// ConversionUtilities provides data conversion utilities
type ConversionUtilities struct {
	dataHandler *DataTypeHandler
}

// NewConversionUtilities creates a new conversion utilities instance
func NewConversionUtilities() *ConversionUtilities {
	return &ConversionUtilities{
		dataHandler: NewDataTypeHandler(),
	}
}

// ConvertStringToType converts a string value to the specified type
func (cu *ConversionUtilities) ConvertStringToType(value string, targetType string) (interface{}, error) {
	switch targetType {
	case "int", "integer":
		return strconv.ParseInt(value, 10, 64)
	case "float":
		return strconv.ParseFloat(value, 64)
	case "bool", "boolean":
		return strconv.ParseBool(value)
	case "string", "text":
		return value, nil
	default:
		return value, nil
	}
}

// ConvertToString converts any value to string
func (cu *ConversionUtilities) ConvertToString(value interface{}) string {
	if value == nil {
		return ""
	}
	return fmt.Sprintf("%v", value)
}

// BatchDataProcessor provides batch data processing utilities
type BatchDataProcessor struct {
	dataHandler *DataTypeHandler
}

// NewBatchDataProcessor creates a new batch data processor
func NewBatchDataProcessor() *BatchDataProcessor {
	return &BatchDataProcessor{
		dataHandler: NewDataTypeHandler(),
	}
}

// ProcessBatchData processes a batch of data with type conversion
func (bdp *BatchDataProcessor) ProcessBatchData(data [][]interface{}) [][]interface{} {
	result := make([][]interface{}, len(data))
	for i, row := range data {
		result[i] = make([]interface{}, len(row))
		for j, cell := range row {
			result[i][j] = bdp.dataHandler.ConvertToExcelValue(cell)
		}
	}
	return result
}

// OptimizeBatchData optimizes batch data for memory efficiency
func (bdp *BatchDataProcessor) OptimizeBatchData(data [][]interface{}) [][]interface{} {
	// For now, just return the processed data
	// In a full implementation, this would include memory optimization
	return bdp.ProcessBatchData(data)
}

// DataTransformer provides data transformation utilities
type DataTransformer struct{}

// NewDataTransformer creates a new data transformer
func NewDataTransformer() *DataTransformer {
	return &DataTransformer{}
}

// TransformJSONToRows converts JSON data to row format
func (dt *DataTransformer) TransformJSONToRows(jsonData map[string]interface{}) ([][]interface{}, error) {
	// Simple JSON to rows transformation
	var rows [][]interface{}

	// Extract headers and data
	if data, ok := jsonData["data"].([]interface{}); ok {
		for _, item := range data {
			if itemMap, ok := item.(map[string]interface{}); ok {
				var row []interface{}
				for _, value := range itemMap {
					row = append(row, value)
				}
				rows = append(rows, row)
			}
		}
	}

	return rows, nil
}

// PivotData transforms data into pivot format
func (dt *DataTransformer) PivotData(data []map[string]interface{}, config PivotConfig) ([][]interface{}, error) {
	// Simple pivot implementation
	var result [][]interface{}

	// Add header row
	header := make([]interface{}, 0)
	for _, field := range config.RowFields {
		header = append(header, field)
	}
	for _, field := range config.ColumnFields {
		header = append(header, field)
	}
	for _, field := range config.ValueFields {
		header = append(header, field)
	}
	result = append(result, header)

	// Add data rows
	for _, row := range data {
		dataRow := make([]interface{}, 0)
		for _, field := range config.RowFields {
			if value, ok := row[field]; ok {
				dataRow = append(dataRow, value)
			} else {
				dataRow = append(dataRow, "")
			}
		}
		for _, field := range config.ColumnFields {
			if value, ok := row[field]; ok {
				dataRow = append(dataRow, value)
			} else {
				dataRow = append(dataRow, "")
			}
		}
		for _, field := range config.ValueFields {
			if value, ok := row[field]; ok {
				dataRow = append(dataRow, value)
			} else {
				dataRow = append(dataRow, "")
			}
		}
		result = append(result, dataRow)
	}

	return result, nil
}

// FlattenNestedData flattens nested data structures
func (dt *DataTransformer) FlattenNestedData(data interface{}, separator string) map[string]interface{} {
	result := make(map[string]interface{})
	dt.flattenRecursive(data, "", separator, result)
	return result
}

func (dt *DataTransformer) flattenRecursive(data interface{}, prefix string, separator string, result map[string]interface{}) {
	v := reflect.ValueOf(data)
	if !v.IsValid() {
		return
	}

	switch v.Kind() {
	case reflect.Map:
		for _, key := range v.MapKeys() {
			keyStr := fmt.Sprintf("%v", key.Interface())
			newPrefix := keyStr
			if prefix != "" {
				newPrefix = prefix + separator + keyStr
			}
			dt.flattenRecursive(v.MapIndex(key).Interface(), newPrefix, separator, result)
		}
	case reflect.Slice, reflect.Array:
		for i := 0; i < v.Len(); i++ {
			newPrefix := fmt.Sprintf("%d", i)
			if prefix != "" {
				newPrefix = prefix + separator + newPrefix
			}
			dt.flattenRecursive(v.Index(i).Interface(), newPrefix, separator, result)
		}
	default:
		if prefix != "" {
			result[prefix] = data
		}
	}
}
