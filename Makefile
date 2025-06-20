.PHONY: all

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

help:
	@echo "Usage: make <target>"
	@echo "Targets:"
	@echo "  all      - Run all tests and generate coverage report"
	@echo "  test     - Run tests"
	@echo "  test-cover - Run tests with coverage"
	@echo "  test-race  - Run tests with race detector"
	@echo "  test-all   - Run all tests (with coverage and race detector)"
	@echo "  test-clean - Remove coverage profile file"