# 10-performance - Performance Optimization Example

This example demonstrates performance optimization techniques for handling large datasets and improving Excel generation speed.

## Features Demonstrated

### Large Dataset Handling
- Efficient processing of 10,000+ rows
- Memory-optimized data structures
- Batch processing techniques
- Streaming data processing

### Performance Optimizations
- Style caching and reuse
- Bulk operations
- Memory management
- Concurrent processing

### Benchmarking
- Performance measurement
- Memory usage tracking
- Execution time analysis
- Optimization comparisons

### Best Practices
- Efficient data structures
- Resource management
- Error handling for large datasets
- Progress tracking

## Output

The example generates multiple files demonstrating different optimization approaches:

1. **10-performance-basic.xlsx** - Basic approach (slower)
2. **10-performance-optimized.xlsx** - Optimized approach (faster)
3. **10-performance-bulk.xlsx** - Bulk operations approach
4. **10-performance-streaming.xlsx** - Streaming approach

## Usage

```bash
cd examples/10-performance
go run main.go
```

The example will output performance metrics and generate optimized Excel files.

## Performance Metrics

The example tracks and displays:
- Execution time for each approach
- Memory usage patterns
- Rows processed per second
- File size comparisons

## Key Learning Points

- How to handle large datasets efficiently
- Memory optimization techniques
- Performance measurement and profiling
- Best practices for Excel generation at scale
- Trade-offs between speed and memory usage