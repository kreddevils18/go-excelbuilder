.PHONY: all examples run-examples

all: test test-cover test-race test-all test-clean

test:
	go test -v ./tests/...

test-cover:
	go test -v -coverprofile=coverage.out ./tests/...
	go tool cover -html=coverage.out

test-race:
	go test -v -race ./tests/...

test-all:
	go test -v -coverprofile=coverage.out -race ./tests/...
	go tool cover -html=coverage.out

test-clean:
	rm -f coverage.out

# Run all examples
examples: run-examples

run-examples:
	@echo "Running all examples..."
	@for dir in examples/*/; do \
		if [ -f "$$dir/main.go" ]; then \
			echo "\n=== Running $$dir ==="; \
			cd "$$dir" && go run main.go && cd ../..; \
		fi; \
	done
	@echo "\nAll examples completed!"

# Run specific example
run-example:
	@if [ -z "$(EXAMPLE)" ]; then \
		echo "Usage: make run-example EXAMPLE=<example-name>"; \
		echo "Available examples:"; \
		ls -1 examples/ | grep -E '^[0-9]' | sort; \
	else \
		if [ -f "examples/$(EXAMPLE)/main.go" ]; then \
			echo "Running example: $(EXAMPLE)"; \
			cd "examples/$(EXAMPLE)" && go run main.go; \
		else \
			echo "Example '$(EXAMPLE)' not found or missing main.go"; \
		fi; \
	fi

# List all available examples
list-examples:
	@echo "Available examples:"
	@ls -1 examples/ | grep -E '^[0-9]' | sort

# Clean generated Excel files from examples
clean-examples:
	@echo "Cleaning generated Excel files..."
	@find examples/ -name "*.xlsx" -type f -delete
	@echo "Excel files cleaned!"

help:
	@echo "Usage: make <target>"
	@echo "Targets:"
	@echo "  all           - Run all tests and generate coverage report"
	@echo "  test          - Run tests"
	@echo "  test-cover    - Run tests with coverage"
	@echo "  test-race     - Run tests with race detector"
	@echo "  test-all      - Run all tests (with coverage and race detector)"
	@echo "  examples      - Run all examples"
	@echo "  run-examples  - Run all examples (alias for examples)"
	@echo "  run-example   - Run specific example (Usage: make run-example EXAMPLE=01-basic-enhanced)"
	@echo "  list-examples - List all available examples"
	@echo "  clean-examples- Clean generated Excel files from examples"
	@echo "  test-clean - Remove coverage profile file"