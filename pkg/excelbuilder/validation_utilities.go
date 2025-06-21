package excelbuilder

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// ValidationRule represents a validation rule
type ValidationRule struct {
	Type       string
	Message    string
	Validator  func(interface{}) bool
	Parameters map[string]interface{}
}

// ValidationResult represents the result of a validation
type ValidationResult struct {
	Valid   bool
	Message string
	Value   interface{}
}

// DataValidator provides comprehensive data validation utilities
type DataValidator struct {
	rules map[string]*ValidationRule
}

// NewDataValidator creates a new data validator
func NewDataValidator() *DataValidator {
	return &DataValidator{
		rules: make(map[string]*ValidationRule),
	}
}

// AddRule adds a validation rule
func (dv *DataValidator) AddRule(name string, rule *ValidationRule) {
	dv.rules[name] = rule
}

// ValidateValue validates a value against a specific rule
func (dv *DataValidator) ValidateValue(value interface{}, ruleName string) ValidationResult {
	rule, exists := dv.rules[ruleName]
	if !exists {
		return ValidationResult{
			Valid:   false,
			Message: fmt.Sprintf("Validation rule '%s' not found", ruleName),
			Value:   value,
		}
	}

	if rule.Validator(value) {
		return ValidationResult{
			Valid:   true,
			Message: "Valid",
			Value:   value,
		}
	}

	return ValidationResult{
		Valid:   false,
		Message: rule.Message,
		Value:   value,
	}
}

// ValidateEmail validates email format
func (dv *DataValidator) ValidateEmail(value interface{}) bool {
	str, ok := value.(string)
	if !ok {
		return false
	}

	emailRegex := regexp.MustCompile(`^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$`)
	return emailRegex.MatchString(str)
}

// ValidatePhone validates phone number format
func (dv *DataValidator) ValidatePhone(value interface{}) bool {
	str, ok := value.(string)
	if !ok {
		return false
	}

	// Remove common phone number separators
	cleanPhone := strings.ReplaceAll(str, "-", "")
	cleanPhone = strings.ReplaceAll(cleanPhone, " ", "")
	cleanPhone = strings.ReplaceAll(cleanPhone, "(", "")
	cleanPhone = strings.ReplaceAll(cleanPhone, ")", "")
	cleanPhone = strings.ReplaceAll(cleanPhone, "+", "")

	// Check if remaining characters are digits and length is reasonable
	if len(cleanPhone) < 10 || len(cleanPhone) > 15 {
		return false
	}

	for _, char := range cleanPhone {
		if char < '0' || char > '9' {
			return false
		}
	}

	return true
}

// ValidateURL validates URL format
func (dv *DataValidator) ValidateURL(value interface{}) bool {
	str, ok := value.(string)
	if !ok {
		return false
	}

	urlRegex := regexp.MustCompile(`^https?://[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}(/.*)?$`)
	return urlRegex.MatchString(str)
}

// ValidateRange validates if a numeric value is within a specified range
func (dv *DataValidator) ValidateRange(value interface{}, min, max float64) bool {
	var numValue float64
	var ok bool

	switch v := value.(type) {
	case int:
		numValue = float64(v)
		ok = true
	case int64:
		numValue = float64(v)
		ok = true
	case float32:
		numValue = float64(v)
		ok = true
	case float64:
		numValue = v
		ok = true
	case string:
		if parsed, err := strconv.ParseFloat(v, 64); err == nil {
			numValue = parsed
			ok = true
		}
	}

	if !ok {
		return false
	}

	return numValue >= min && numValue <= max
}

// ValidateLength validates string length
func (dv *DataValidator) ValidateLength(value interface{}, minLen, maxLen int) bool {
	str, ok := value.(string)
	if !ok {
		return false
	}

	length := len(str)
	return length >= minLen && length <= maxLen
}

// ValidateRequired validates that a value is not empty
func (dv *DataValidator) ValidateRequired(value interface{}) bool {
	if value == nil {
		return false
	}

	switch v := value.(type) {
	case string:
		return strings.TrimSpace(v) != ""
	case int, int64, float32, float64:
		return true
	case bool:
		return true
	default:
		return true
	}
}

// ValidatePattern validates against a regex pattern
func (dv *DataValidator) ValidatePattern(value interface{}, pattern string) bool {
	str, ok := value.(string)
	if !ok {
		return false
	}

	regex, err := regexp.Compile(pattern)
	if err != nil {
		return false
	}

	return regex.MatchString(str)
}

// CreateEmailRule creates an email validation rule
func (dv *DataValidator) CreateEmailRule() *ValidationRule {
	return &ValidationRule{
		Type:      "email",
		Message:   "Invalid email format",
		Validator: dv.ValidateEmail,
	}
}

// CreatePhoneRule creates a phone validation rule
func (dv *DataValidator) CreatePhoneRule() *ValidationRule {
	return &ValidationRule{
		Type:      "phone",
		Message:   "Invalid phone number format",
		Validator: dv.ValidatePhone,
	}
}

// CreateURLRule creates a URL validation rule
func (dv *DataValidator) CreateURLRule() *ValidationRule {
	return &ValidationRule{
		Type:      "url",
		Message:   "Invalid URL format",
		Validator: dv.ValidateURL,
	}
}

// CreateRangeRule creates a range validation rule
func (dv *DataValidator) CreateRangeRule(min, max float64) *ValidationRule {
	return &ValidationRule{
		Type:    "range",
		Message: fmt.Sprintf("Value must be between %.2f and %.2f", min, max),
		Validator: func(value interface{}) bool {
			return dv.ValidateRange(value, min, max)
		},
		Parameters: map[string]interface{}{
			"min": min,
			"max": max,
		},
	}
}

// CreateLengthRule creates a length validation rule
func (dv *DataValidator) CreateLengthRule(minLen, maxLen int) *ValidationRule {
	return &ValidationRule{
		Type:    "length",
		Message: fmt.Sprintf("Length must be between %d and %d characters", minLen, maxLen),
		Validator: func(value interface{}) bool {
			return dv.ValidateLength(value, minLen, maxLen)
		},
		Parameters: map[string]interface{}{
			"minLen": minLen,
			"maxLen": maxLen,
		},
	}
}

// CreateRequiredRule creates a required validation rule
func (dv *DataValidator) CreateRequiredRule() *ValidationRule {
	return &ValidationRule{
		Type:      "required",
		Message:   "This field is required",
		Validator: dv.ValidateRequired,
	}
}

// CreatePatternRule creates a pattern validation rule
func (dv *DataValidator) CreatePatternRule(pattern, message string) *ValidationRule {
	return &ValidationRule{
		Type:    "pattern",
		Message: message,
		Validator: func(value interface{}) bool {
			return dv.ValidatePattern(value, pattern)
		},
		Parameters: map[string]interface{}{
			"pattern": pattern,
		},
	}
}

// BatchValidator provides batch validation utilities
type BatchValidator struct {
	validator *DataValidator
	errors    []ValidationResult
}

// NewBatchValidator creates a new batch validator
func NewBatchValidator() *BatchValidator {
	return &BatchValidator{
		validator: NewDataValidator(),
		errors:    make([]ValidationResult, 0),
	}
}

// ValidateBatch validates a batch of data
func (bv *BatchValidator) ValidateBatch(data [][]interface{}, rules map[int]string) []ValidationResult {
	var results []ValidationResult

	for rowIndex, row := range data {
		for colIndex, cell := range row {
			if ruleName, exists := rules[colIndex]; exists {
				result := bv.validator.ValidateValue(cell, ruleName)
				if !result.Valid {
					result.Message = fmt.Sprintf("Row %d, Col %d: %s", rowIndex+1, colIndex+1, result.Message)
					results = append(results, result)
				}
			}
		}
	}

	return results
}

// GetErrors returns collected validation errors
func (bv *BatchValidator) GetErrors() []ValidationResult {
	return bv.errors
}

// ClearErrors clears collected validation errors
func (bv *BatchValidator) ClearErrors() {
	bv.errors = make([]ValidationResult, 0)
}

// AddError adds a validation error
func (bv *BatchValidator) AddError(result ValidationResult) {
	bv.errors = append(bv.errors, result)
}
