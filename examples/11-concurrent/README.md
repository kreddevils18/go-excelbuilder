# 11-concurrent - Concurrent Processing Example

This example demonstrates concurrent processing techniques for generating multiple Excel files simultaneously and handling parallel data processing.

## Features Demonstrated

### Concurrent File Generation
- Multiple Excel files generated in parallel
- Worker pool pattern implementation
- Goroutine-based processing
- Channel communication

### Parallel Data Processing
- Concurrent data transformation
- Parallel sheet creation
- Synchronized access to shared resources
- Error handling in concurrent operations

### Advanced Concurrency Patterns
- Producer-consumer pattern
- Fan-out/fan-in pattern
- Pipeline processing
- Rate limiting

### Thread Safety
- Safe concurrent access
- Mutex usage for shared resources
- Atomic operations
- Deadlock prevention

## Output

The example generates multiple files concurrently:

1. **11-concurrent-report-1.xlsx** - Sales report (generated concurrently)
2. **11-concurrent-report-2.xlsx** - Inventory report (generated concurrently)
3. **11-concurrent-report-3.xlsx** - Financial report (generated concurrently)
4. **11-concurrent-summary.xlsx** - Combined summary report
5. **11-concurrent-pipeline.xlsx** - Pipeline processing result

## Usage

```bash
cd examples/11-concurrent
go run main.go
```

The example will demonstrate concurrent processing with timing information and progress tracking.

## Concurrency Metrics

The example tracks and displays:
- Execution time for sequential vs concurrent approaches
- Number of goroutines used
- Memory usage patterns
- Throughput improvements
- Error rates and handling

## Key Learning Points

- How to implement concurrent Excel generation
- Worker pool patterns for file processing
- Safe handling of shared resources
- Performance benefits of concurrency
- Best practices for error handling in concurrent operations
- When to use concurrency vs sequential processing