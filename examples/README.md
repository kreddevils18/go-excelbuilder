# Go Excel Builder Examples

This directory contains comprehensive examples demonstrating the Go Excel Builder package capabilities, from basic usage to advanced enterprise scenarios.

## 📚 Learning Path

### 🟢 **Beginner Level (Start Here)**

| Example | Description | Key Features |
|---------|-------------|-------------|
| [01-basic-enhanced](./01-basic-enhanced/) | Enhanced basic operations | Workbook properties, multiple sheets, basic styling |
| [02-data-types](./02-data-types/) | All supported data types | String, numbers, dates, booleans, formulas |
| [03-basic-styling](./03-basic-styling/) | Fundamental styling | Fonts, colors, borders, alignment, number formats |

### 🟡 **Intermediate Level (Business Use Cases)**

| Example | Description | Key Features |
|---------|-------------|-------------|
| [04-sales-report](./04-sales-report/) | Business report generation | Multi-sheet reports, calculations, summaries |
| [05-import-export](./05-import-export/) | Data import/export workflows | CSV/JSON processing, data transformation |
| [06-dashboard](./06-dashboard/) | Interactive Excel dashboards | Charts, pivot tables, conditional formatting |

### 🟠 **Advanced Level (Complex Scenarios)**

| Example | Description | Key Features |
|---------|-------------|-------------|
| [07-financial-analysis](./07-financial-analysis/) | Financial modeling | Complex formulas, cross-sheet references |
| [08-template-system](./08-template-system/) | Template-based generation | Template loading, data binding |
| [09-advanced-layout](./09-advanced-layout/) | Advanced Excel features | Grouping, panes, protection, auto-fit |

### 🔴 **Performance & Enterprise Level**

| Example | Description | Key Features |
|---------|-------------|-------------|
| [10-performance](./10-performance/) | High-performance processing | Large datasets, memory optimization |
| [11-concurrent](./11-concurrent/) | Concurrent operations | Goroutines, thread safety |
| [12-enterprise](./12-enterprise/) | Enterprise integration | Database, REST APIs, error handling |

### 🟣 **Specialized Use Cases**

| Example | Description | Key Features |
|---------|-------------|-------------|
| [13-scientific](./13-scientific/) | Scientific data visualization | Statistical charts, research data |
| [14-inventory](./14-inventory/) | Inventory management | Stock tracking, automated calculations |
| [15-gradebook](./15-gradebook/) | Educational applications | Student data, grade calculations |

## 🚀 Quick Start

1. **Install the package**:
   ```bash
   go get github.com/kreddevils18/go-excelbuilder
   ```

2. **Run any example**:
   ```bash
   cd examples/01-basic-enhanced
   go run main.go
   ```

3. **Check the generated Excel files** in each example's `output/` directory

## 📁 Directory Structure

```
examples/
├── README.md                    # This file
├── 01-basic-enhanced/           # Enhanced basic operations
├── 02-data-types/               # Data type demonstrations
├── 03-basic-styling/            # Styling fundamentals
├── 04-sales-report/             # Business report example
├── 05-import-export/            # Data import/export
├── 06-dashboard/                # Interactive dashboards
├── 07-financial-analysis/       # Financial modeling
├── 08-template-system/          # Template-based generation
├── 09-advanced-layout/          # Advanced layout features
├── 10-performance/              # Performance optimization
├── 11-concurrent/               # Concurrent operations
├── 12-enterprise/               # Enterprise integration
├── 13-scientific/               # Scientific applications
├── 14-inventory/                # Inventory management
├── 15-gradebook/                # Educational applications
├── shared/
│   ├── data/                    # Sample data files
│   ├── templates/               # Excel templates
│   └── utils/                   # Common utilities
└── benchmarks/                  # Performance benchmarks
```

## 🎯 Learning Recommendations

### For Go Beginners
- Start with examples 01-03 to understand the fluent API
- Focus on understanding method chaining patterns
- Practice with different data types and basic styling

### For Business Users
- Jump to examples 04-06 for practical applications
- Learn data import/export workflows
- Master chart and pivot table creation

### For Advanced Developers
- Explore examples 07-09 for complex scenarios
- Study template systems and advanced layouts
- Understand performance optimization techniques

### For Enterprise Applications
- Focus on examples 10-12 for scalable solutions
- Learn concurrent processing patterns
- Study integration with databases and APIs

## 🔧 Common Utilities

The `shared/` directory contains:
- **Sample data files** for testing and learning
- **Excel templates** for template-based examples
- **Common utilities** used across multiple examples

## 📊 Performance Benchmarks

The `benchmarks/` directory includes:
- Performance comparison with raw excelize
- Memory usage analysis
- Scalability testing results

## 🤝 Contributing

To add new examples:
1. Create a new directory following the naming convention
2. Include a `main.go` file with comprehensive comments
3. Add a `README.md` explaining the example's purpose
4. Create an `output/` directory for generated files
5. Update this main README with the new example

## 📖 Additional Resources

- [Package Documentation](../README.md)
- [API Reference](../pkg/excelbuilder/)
- [Test Examples](../tests/)
- [Memory Bank Documentation](../memory-bank/)

---

**Happy Excel Building! 🎉**

Each example is designed to be self-contained and runnable. Start with the basic examples and progressively work your way up to more complex scenarios based on your needs.