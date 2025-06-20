# Go Excel Builder - Product Requirements Document (PRD)

## 1. Tổng quan dự án

### 1.1 Mục tiêu

Tạo một package Golang đơn giản và dễ sử dụng để xây dựng file Excel, sử dụng thư viện excelize v2 làm nền tảng và áp dụng các design pattern Builder và Flyweight để tối ưu hiệu suất và trải nghiệm developer.

### 1.2 Vấn đề cần giải quyết

- Việc tạo file Excel phức tạp với excelize v2 thường yêu cầu nhiều code boilerplate
- Khó khăn trong việc quản lý style và format cho nhiều cell
- Thiếu tính nhất quán trong cách tạo và format Excel
- Cần một API đơn giản, dễ đọc và dễ maintain

### 1.3 Giá trị mang lại

- API đơn giản, fluent interface dễ sử dụng
- Tối ưu memory thông qua Flyweight pattern cho style management
- Flexible và extensible architecture
- Type-safe operations
- Giảm thiểu boilerplate code

## 2. Kiến trúc hệ thống

### 2.1 Design Patterns

#### 2.1.1 Builder Pattern

```
ExcelBuilder
├── WorkbookBuilder
│   ├── AddSheet(name) -> SheetBuilder
│   ├── SetProperties(props)
│   └── Build() -> *excelize.File
├── SheetBuilder
│   ├── AddRow() -> RowBuilder
│   ├── SetColumnWidth(col, width)
│   ├── MergeCell(range)
│   └── Done() -> WorkbookBuilder
└── RowBuilder
    ├── AddCell(value) -> CellBuilder
    ├── SetHeight(height)
    └── Done() -> SheetBuilder
```

#### 2.1.2 Flyweight Pattern

```
StyleManager (Flyweight Factory)
├── GetStyle(styleConfig) -> StyleFlyweight
├── CreateStyle(config) -> StyleFlyweight
└── Cache: map[string]StyleFlyweight

StyleFlyweight (Flyweight)
├── StyleID: int
├── Config: StyleConfig
└── Apply(file, cell)
```

### 2.2 Core Components

#### 2.2.1 ExcelBuilder (Director)

```go
type ExcelBuilder struct {
    file        *excelize.File
    styleManager *StyleManager
    currentSheet string
}
```

#### 2.2.2 WorkbookBuilder (ConcreteBuilder)

```go
type WorkbookBuilder struct {
    builder *ExcelBuilder
    properties WorkbookProperties
}
```

#### 2.2.3 SheetBuilder (ConcreteBuilder)

```go
type SheetBuilder struct {
    builder   *ExcelBuilder
    sheetName string
    currentRow int
}
```

#### 2.2.4 RowBuilder (ConcreteBuilder)

```go
type RowBuilder struct {
    sheetBuilder *SheetBuilder
    rowIndex     int
    currentCol   int
}
```

#### 2.2.5 CellBuilder (ConcreteBuilder)

```go
type CellBuilder struct {
    rowBuilder *RowBuilder
    cellRef    string
    value      interface{}
}
```

#### 2.2.6 StyleManager (Flyweight Factory)

```go
type StyleManager struct {
    file       *excelize.File
    styleCache map[string]*StyleFlyweight
    mutex      sync.RWMutex
}
```

#### 2.2.7 StyleFlyweight (Flyweight)

```go
type StyleFlyweight struct {
    styleID int
    config  StyleConfig
}
```

## 3. API Design

### 3.1 Fluent Interface Examples

#### 3.1.1 Basic Usage

```go
builder := excelbuilder.New()
file, err := builder.
    NewWorkbook().
        SetProperties(excelbuilder.WorkbookProperties{
            Title:   "Sales Report",
            Author:  "John Doe",
            Subject: "Monthly Sales Data",
        }).
        AddSheet("Sales").
            SetColumnWidth("A", 15).
            SetColumnWidth("B", 20).
            AddRow().
                AddCell("Product").WithStyle(headerStyle).
                AddCell("Revenue").WithStyle(headerStyle).
                Done().
            AddRow().
                AddCell("iPhone").
                AddCell(1000000).WithFormat("currency").
                Done().
            Done().
        Build()
```

#### 3.1.2 Advanced Usage với Custom Styles

```go
headerStyle := excelbuilder.StyleConfig{
    Font: excelbuilder.FontConfig{
        Bold:   true,
        Size:   12,
        Color:  "#FFFFFF",
        Family: "Arial",
    },
    Fill: excelbuilder.FillConfig{
        Type:  "pattern",
        Color: "#4472C4",
    },
    Alignment: excelbuilder.AlignmentConfig{
        Horizontal: "center",
        Vertical:   "center",
    },
    Border: excelbuilder.BorderConfig{
        Top:    excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
        Bottom: excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
        Left:   excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
        Right:  excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
    },
}

builder := excelbuilder.New()
file, err := builder.
    NewWorkbook().
        AddSheet("Data").
            AddRow().
                AddCell("Header 1").WithStyle(headerStyle).
                AddCell("Header 2").WithStyle(headerStyle).
                Done().
            AddRowWithData([]interface{}{"Value 1", "Value 2"}).
            MergeCell("A3:B3").
            AddRow().
                AddCell("Merged Cell").WithStyle(headerStyle).
                Done().
            Done().
        Build()
```

### 3.2 Configuration Structures

#### 3.2.1 StyleConfig

```go
type StyleConfig struct {
    Font      FontConfig      `json:"font,omitempty"`
    Fill      FillConfig      `json:"fill,omitempty"`
    Alignment AlignmentConfig `json:"alignment,omitempty"`
    Border    BorderConfig    `json:"border,omitempty"`
    Format    string          `json:"format,omitempty"`
}
```

#### 3.2.2 FontConfig

```go
type FontConfig struct {
    Bold      bool    `json:"bold,omitempty"`
    Italic    bool    `json:"italic,omitempty"`
    Underline string  `json:"underline,omitempty"`
    Size      float64 `json:"size,omitempty"`
    Color     string  `json:"color,omitempty"`
    Family    string  `json:"family,omitempty"`
}
```

#### 3.2.3 WorkbookProperties

```go
type WorkbookProperties struct {
    Title       string `json:"title,omitempty"`
    Subject     string `json:"subject,omitempty"`
    Author      string `json:"author,omitempty"`
    Manager     string `json:"manager,omitempty"`
    Company     string `json:"company,omitempty"`
    Category    string `json:"category,omitempty"`
    Keywords    string `json:"keywords,omitempty"`
    Description string `json:"description,omitempty"`
}
```

## 4. Cấu trúc thư mục

```
go-excelbuilder/
├── README.md
├── go.mod
├── go.sum
├── .gitignore
├── .cursor/
│   └── rules/
│       └── prd.md
├── pkg/
│   └── excelbuilder/
│       ├── builder.go          # ExcelBuilder, WorkbookBuilder
│       ├── sheet.go            # SheetBuilder
│       ├── row.go              # RowBuilder
│       ├── cell.go             # CellBuilder
│       ├── style.go            # StyleManager, StyleFlyweight
│       ├── config.go           # Configuration structs
│       ├── errors.go           # Custom errors
│       └── utils.go            # Utility functions
├── examples/
│   ├── basic/
│   │   └── main.go
│   ├── advanced/
│   │   └── main.go
│   └── performance/
│       └── main.go
├── tests/
│   ├── builder_test.go
│   ├── style_test.go
│   ├── integration_test.go
│   └── benchmark_test.go
└── docs/
    ├── api.md
    ├── examples.md
    └── performance.md
```

## 5. Technical Specifications

### 5.1 Dependencies

```go
module github.com/username/go-excelbuilder

go 1.21

require (
    github.com/xuri/excelize/v2 v2.8.0
)

require (
    github.com/mohae/deepcopy v0.0.0-20170929034955-c48cc78d4826 // indirect
    github.com/richardlehane/mscfb v1.0.4 // indirect
    github.com/richardlehane/msoleps v1.0.3 // indirect
    github.com/xuri/efp v0.0.0-20230802181842-ad255f2331ca // indirect
    github.com/xuri/nfp v0.0.0-20230819163627-dc951e3ffe1a // indirect
    golang.org/x/crypto v0.12.0 // indirect
    golang.org/x/net v0.14.0 // indirect
    golang.org/x/text v0.12.0 // indirect
)
```

### 5.2 Performance Requirements

- Hỗ trợ tạo file Excel với ít nhất 100,000 rows
- Memory usage tối ưu thông qua Flyweight pattern
- Style caching để tránh tạo duplicate styles
- Concurrent-safe operations

### 5.3 Error Handling

```go
type ExcelBuilderError struct {
    Op   string // Operation that failed
    Err  error  // Underlying error
    Code ErrorCode
}

type ErrorCode int

const (
    ErrInvalidSheetName ErrorCode = iota
    ErrInvalidCellReference
    ErrStyleNotFound
    ErrWorkbookNotInitialized
    ErrSheetNotFound
)
```

## 6. Implementation Phases

### Phase 1: Core Builder Implementation

- [ ] ExcelBuilder và WorkbookBuilder
- [ ] SheetBuilder với basic operations
- [ ] RowBuilder và CellBuilder
- [ ] Basic error handling
- [ ] Unit tests cho core functionality

### Phase 2: Style Management với Flyweight

- [ ] StyleManager implementation
- [ ] StyleFlyweight với caching
- [ ] Style configuration structures
- [ ] Integration với builders
- [ ] Performance tests

### Phase 3: Advanced Features

- [ ] Column width và row height management
- [ ] Cell merging functionality
- [ ] Data validation
- [ ] Conditional formatting
- [ ] Chart support (future)

### Phase 4: Documentation và Examples

- [ ] Comprehensive API documentation
- [ ] Usage examples
- [ ] Performance benchmarks
- [ ] Best practices guide

## 7. Testing Strategy

### 7.1 Unit Tests

- Builder pattern functionality
- Style management và caching
- Error handling
- Configuration validation

### 7.2 Integration Tests

- End-to-end Excel file creation
- Complex scenarios với multiple sheets
- Style application verification

### 7.3 Performance Tests

- Memory usage với large datasets
- Style caching effectiveness
- Concurrent operations

### 7.4 Benchmark Tests

```go
func BenchmarkExcelBuilder_LargeFile(b *testing.B) {
    for i := 0; i < b.N; i++ {
        builder := excelbuilder.New()
        // Create file with 10,000 rows
    }
}

func BenchmarkStyleManager_CacheEfficiency(b *testing.B) {
    // Test style caching performance
}
```

## 8. Future Enhancements

### 8.1 Version 2.0 Features

- Chart creation support
- Pivot table functionality
- Advanced conditional formatting
- Template-based generation
- Streaming support cho very large files

### 8.2 Integration Possibilities

- CLI tool cho Excel generation
- Web service wrapper
- Database integration helpers
- JSON to Excel converter

## 9. Success Metrics

### 9.1 Technical Metrics

- Code coverage > 90%
- Memory usage < 50MB cho 100K rows
- API response time < 100ms cho basic operations
- Zero memory leaks

### 9.2 Developer Experience Metrics

- Reduced boilerplate code by 70%
- API learning curve < 30 minutes
- Clear error messages
- Comprehensive documentation

## 10. Risk Assessment

### 10.1 Technical Risks

- **Risk**: Excelize v2 API changes

  - **Mitigation**: Version pinning và compatibility layer

- **Risk**: Memory usage với large files

  - **Mitigation**: Streaming approach và memory profiling

- **Risk**: Performance degradation với complex styles
  - **Mitigation**: Flyweight pattern và benchmarking

### 10.2 Adoption Risks

- **Risk**: Learning curve cho new API

  - **Mitigation**: Comprehensive examples và documentation

- **Risk**: Migration từ existing solutions
  - **Mitigation**: Migration guides và compatibility helpers

This PRD provides a comprehensive roadmap for building a robust, efficient, and developer-friendly Excel builder package in Go.
