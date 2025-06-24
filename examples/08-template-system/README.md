# Template System Example

This example demonstrates the advanced template system capabilities of the go-excelbuilder library, showing how to create reusable Excel templates with dynamic content generation.

## Features Demonstrated

### Template Management
- **Template Creation**: Define reusable Excel templates
- **Template Loading**: Load templates from files or embedded resources
- **Template Validation**: Ensure template integrity and structure
- **Template Versioning**: Manage different template versions

### Dynamic Content Generation
- **Variable Substitution**: Replace placeholders with dynamic data
- **Conditional Sections**: Show/hide sections based on data conditions
- **Loop Structures**: Repeat sections for collections of data
- **Nested Templates**: Include sub-templates within main templates

### Advanced Template Features
- **Custom Functions**: Define custom template functions
- **Data Binding**: Bind complex data structures to templates
- **Template Inheritance**: Extend base templates with specific layouts
- **Multi-language Support**: Generate templates in different languages

### Template Types
- **Report Templates**: Standard business reports
- **Invoice Templates**: Professional invoice layouts
- **Dashboard Templates**: Interactive dashboard layouts
- **Form Templates**: Data entry and collection forms
- **Certificate Templates**: Awards and certification documents

## Expected Output

The example generates several Excel files demonstrating different template scenarios:

1. **08-template-reports.xlsx**: Various business report templates
2. **08-template-invoices.xlsx**: Professional invoice templates
3. **08-template-dashboards.xlsx**: Dashboard and analytics templates
4. **08-template-forms.xlsx**: Data collection form templates
5. **08-template-certificates.xlsx**: Certificate and award templates

## Usage

```bash
cd examples/08-template-system
go run main.go
```

## Template Structure

Templates use a combination of:
- **Placeholders**: `{{variable_name}}` for simple substitution
- **Conditionals**: `{{#if condition}}...{{/if}}` for conditional content
- **Loops**: `{{#each items}}...{{/each}}` for repeating sections
- **Includes**: `{{>template_name}}` for sub-templates
- **Helpers**: `{{helper_name args}}` for custom functions

## Key Learning Points

1. **Template Design**: Best practices for creating maintainable templates
2. **Data Modeling**: Structuring data for optimal template rendering
3. **Performance**: Efficient template compilation and rendering
4. **Reusability**: Creating modular and extensible template systems
5. **Localization**: Supporting multiple languages and regions
6. **Error Handling**: Robust template validation and error reporting

## Template Examples

### Simple Variable Substitution
```
Company: {{company_name}}
Date: {{report_date}}
Total Sales: {{total_sales}}
```

### Conditional Content
```
{{#if has_discount}}
Discount Applied: {{discount_percentage}}%
{{/if}}
```

### Loop Structures
```
{{#each products}}
Product: {{name}} - Price: {{price}}
{{/each}}
```

### Custom Functions
```
Formatted Date: {{formatDate date "YYYY-MM-DD"}}
Currency: {{formatCurrency amount "USD"}}
```

This example showcases the power and flexibility of template-based Excel generation, enabling rapid development of consistent, professional documents.