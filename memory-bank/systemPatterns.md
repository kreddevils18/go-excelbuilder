# System Patterns - Go Excel Builder

## Architecture Overview

Go Excel Builder sử dụng layered architecture với hai design patterns chính:
- **Builder Pattern**: Để tạo fluent interface và quản lý construction process
- **Flyweight Pattern**: Để tối ưu memory usage cho style management

## Design Patterns Implementation

### 1. Builder Pattern

#### Pattern Structure
```
Director (ExcelBuilder)
├── ConcreteBuilder (WorkbookBuilder)
│   ├── Product: Workbook
│   └── BuildParts: AddSheet, SetProperties
├── ConcreteBuilder (SheetBuilder)
│   ├── Product: Sheet
│   └── BuildParts: AddRow, SetColumnWidth, MergeCell
├── ConcreteBuilder (RowBuilder)
│   ├── Product: Row
│   └── BuildParts: AddCell, SetHeight
└── ConcreteBuilder (CellBuilder)
    ├── Product: Cell
    └── BuildParts: WithStyle, WithFormat
```

#### Key Components

**ExcelBuilder (Director)**
- Orchestrates the building process
- Manages excelize.File instance
- Coordinates with StyleManager
- Provides entry point for fluent API

**WorkbookBuilder (ConcreteBuilder)**
- Builds workbook-level configurations
- Manages workbook properties
- Creates and manages sheets
- Returns final excelize.File

**SheetBuilder (ConcreteBuilder)**
- Builds sheet-level configurations
- Manages rows and columns
- Handles sheet-specific operations
- Tracks current position

**RowBuilder (ConcreteBuilder)**
- Builds row-level configurations
- Manages cells within row
- Handles row-specific styling
- Tracks column position

**CellBuilder (ConcreteBuilder)**
- Builds individual cell configurations
- Applies values and styles
- Handles cell-specific formatting
- Integrates with StyleManager

#### Benefits
- **Fluent Interface**: Natural, readable API
- **Step-by-step Construction**: Complex objects built incrementally
- **Flexibility**: Different representations of same data
- **Validation**: Each step can validate inputs

### 2. Flyweight Pattern

#### Pattern Structure
```
FlyweightFactory (StyleManager)
├── GetFlyweight(key) -> StyleFlyweight
├── CreateFlyweight(config) -> StyleFlyweight
└── flyweights: map[string]*StyleFlyweight

Flyweight (StyleFlyweight)
├── IntrinsicState: StyleConfig
├── ExtrinsicState: CellReference
└── Operation: Apply(file, cellRef)

Context (CellBuilder)
├── flyweight: *StyleFlyweight
├── extrinsicState: cellReference
└── Operation: WithStyle(config)
```

#### Key Components

**StyleManager (FlyweightFactory)**
- Creates and manages StyleFlyweight instances
- Maintains cache of created styles
- Generates unique keys for style configurations
- Thread-safe operations with mutex

**StyleFlyweight (Flyweight)**
- Stores intrinsic state (style configuration)
- Provides Apply method for extrinsic state
- Immutable once created
- Shared across multiple cells

**StyleConfig (IntrinsicState)**
- Font, Fill, Alignment, Border configurations
- Serializable for cache key generation
- Immutable data structure
- JSON-compatible for future persistence

#### Benefits
- **Memory Efficiency**: Shared style instances
- **Performance**: Reduced object creation overhead
- **Consistency**: Centralized style management
- **Scalability**: Handles large numbers of styled cells

## Component Relationships

### Data Flow
```
User API Call
    ↓
ExcelBuilder (validates, coordinates)
    ↓
ConcreteBuilder (builds specific component)
    ↓
StyleManager (manages styles if needed)
    ↓
Excelize Operations (actual Excel manipulation)
    ↓
Result (updated Excel file)
```

### Dependency Graph
```
ExcelBuilder
├── depends on: excelize.File
├── depends on: StyleManager
└── creates: WorkbookBuilder

WorkbookBuilder
├── depends on: ExcelBuilder
└── creates: SheetBuilder

SheetBuilder
├── depends on: ExcelBuilder
└── creates: RowBuilder

RowBuilder
├── depends on: SheetBuilder
└── creates: CellBuilder

CellBuilder
├── depends on: RowBuilder
├── depends on: StyleManager
└── uses: StyleFlyweight

StyleManager
├── depends on: excelize.File
├── creates: StyleFlyweight
└── manages: style cache
```

## Key Technical Decisions

### 1. Fluent Interface Design
**Decision**: Each builder returns itself or next builder in chain
**Rationale**: Enables method chaining for readable code
**Trade-offs**: Slightly more complex error handling

### 2. Style Caching Strategy
**Decision**: Hash-based caching with JSON serialization
**Rationale**: Efficient lookup and collision avoidance
**Trade-offs**: Serialization overhead for cache key generation

### 3. Thread Safety Approach
**Decision**: RWMutex for StyleManager, immutable flyweights
**Rationale**: Balance between safety and performance
**Trade-offs**: Some locking overhead for concurrent access

### 4. Error Handling Strategy
**Decision**: Custom error types with operation context
**Rationale**: Better debugging and error recovery
**Trade-offs**: More complex error handling code

### 5. Memory Management
**Decision**: Flyweight for styles, builder cleanup
**Rationale**: Optimize for large files with many styled cells
**Trade-offs**: Additional complexity in style management

## Performance Characteristics

### Time Complexity
- **Style Creation**: O(1) amortized (with caching)
- **Cell Creation**: O(1) per cell
- **Row Creation**: O(n) where n = number of cells
- **Sheet Creation**: O(m) where m = number of rows

### Space Complexity
- **Style Storage**: O(k) where k = unique styles
- **Builder State**: O(1) per active builder
- **Excel File**: O(n*m) where n*m = total cells

### Scalability Factors
- **Style Reuse**: Higher reuse = better performance
- **Memory Pressure**: Flyweight pattern reduces pressure
- **Concurrent Access**: RWMutex allows concurrent reads

## Extension Points

### 1. New Builder Types
- ChartBuilder for chart creation
- PivotBuilder for pivot tables
- ValidationBuilder for data validation

### 2. Style Extensions
- ConditionalStyleFlyweight for conditional formatting
- ThemeStyleFlyweight for theme-based styling
- CustomStyleFlyweight for user-defined styles

### 3. Output Formats
- StreamingBuilder for large file streaming
- TemplateBuilder for template-based generation
- BatchBuilder for multiple file generation

## Quality Attributes

### Maintainability
- Clear separation of concerns
- Single responsibility principle
- Dependency injection ready
- Comprehensive test coverage

### Extensibility
- Plugin-ready architecture
- Interface-based design
- Factory patterns for new components
- Configuration-driven behavior

### Performance
- Memory-efficient style management
- Lazy evaluation where possible
- Optimized for common use cases
- Benchmarked critical paths

### Reliability
- Immutable data structures
- Thread-safe operations
- Comprehensive error handling
- Input validation at boundaries

## Implementation Status

### Current State: Core Implementation ✅
- [x] **Architecture Implemented**: Complete pattern implementation
- [x] **Builder Pattern**: Full fluent interface working
- [x] **Interface Design**: All builder interfaces implemented
- [x] **Error Handling**: Comprehensive validation system
- [x] **Testing Strategy**: TDD with 32 passing tests
- [x] **Basic Style Management**: Font, alignment, borders, fills
- [x] **Excel Integration**: Working excelize.File generation

### Next Phase: Advanced Features 🚧
- [ ] **Flyweight Pattern**: Style caching optimization
- [ ] **Chart Support**: Chart creation framework
- [ ] **Advanced Styling**: Complex style combinations
- [ ] **Performance**: Memory and speed optimization

#### Implementation Notes
```go
// Successfully implemented structure
type ExcelBuilder struct {
    file         *excelize.File
    styleManager *StyleManager
    currentSheet string
    config       *Config
    // Additional fields for state management
    currentRow   int
    currentCol   string
}

// Working fluent interface example:
builder := excelbuilder.New()
workbook := builder.NewWorkbook()
sheet := workbook.AddSheet("Sales")
row := sheet.AddRow()
row.AddCell("Product").AddCell("Price")
```