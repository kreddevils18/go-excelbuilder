package tests

import (
	"sync"
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
)

// Test Case 8.1: StyleManager Creation and Initialization
func TestStyleManager_New(t *testing.T) {
	// Test: Check StyleManager instance creation
	// Expected:
	// - StyleManager created successfully
	// - Cache initialized
	// - Thread-safe operations
	// - No error

	manager := excelbuilder.NewStyleManager()

	assert.NotNil(t, manager, "Expected StyleManager instance, got nil")

	// Test that manager can handle basic operations
	styleConfig := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
			Color: "#000000",
		},
	}

	style := manager.GetStyle(styleConfig)
	assert.NotNil(t, style, "Expected StyleFlyweight instance, got nil")
}

// Test Case 8.2: Style Caching Mechanism
func TestStyleManager_StyleCaching(t *testing.T) {
	// Test: Check that identical styles are cached and reused
	// Expected:
	// - Same style config returns same flyweight instance
	// - Cache hit improves performance
	// - Memory usage optimized

	manager := excelbuilder.NewStyleManager()

	styleConfig := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 14,
			Color: "#FF0000",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#FFFF00",
		},
	}

	// Get style twice with same config
	style1 := manager.GetStyle(styleConfig)
	style2 := manager.GetStyle(styleConfig)

	// Should return the same instance (pointer equality)
	assert.Same(t, style1, style2, "Expected same StyleFlyweight instance for identical configs")

	// Verify cache statistics
	stats := manager.GetCacheStats()
	assert.Equal(t, 1, stats.TotalStyles, "Expected 1 unique style in cache")
	assert.Equal(t, 2, stats.CacheHits+stats.CacheMisses, "Expected 2 total requests")
	assert.Equal(t, 1, stats.CacheHits, "Expected 1 cache hit")
}

// Test Case 8.3: Thread Safety
func TestStyleManager_ThreadSafety(t *testing.T) {
	// Test: Check thread-safe operations
	// Expected:
	// - Concurrent access works correctly
	// - No race conditions
	// - Cache consistency maintained

	manager := excelbuilder.NewStyleManager()
	var wg sync.WaitGroup
	const numGoroutines = 100
	const numOperations = 10

	// Create multiple goroutines accessing the same style
	styleConfig := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
		},
	}

	results := make([]*excelbuilder.StyleFlyweight, numGoroutines*numOperations)
	resultIndex := 0
	var mu sync.Mutex

	for i := 0; i < numGoroutines; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for j := 0; j < numOperations; j++ {
				style := manager.GetStyle(styleConfig)
				mu.Lock()
				results[resultIndex] = style
				resultIndex++
				mu.Unlock()
			}
		}()
	}

	wg.Wait()

	// All results should be the same instance
	firstStyle := results[0]
	for i, style := range results {
		assert.Same(t, firstStyle, style, "Expected same StyleFlyweight instance at index %d", i)
	}

	// Should have only one unique style in cache
	stats := manager.GetCacheStats()
	assert.Equal(t, 1, stats.TotalStyles, "Expected 1 unique style in cache after concurrent access")
}

// Test Case 8.4: Different Styles Create Different Flyweights
func TestStyleManager_DifferentStyles(t *testing.T) {
	// Test: Check that different style configs create different flyweights
	// Expected:
	// - Different configs return different instances
	// - Cache stores multiple styles
	// - Each style is unique

	manager := excelbuilder.NewStyleManager()

	style1Config := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
		},
	}

	style2Config := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: false,
			Size: 14,
		},
	}

	style3Config := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
			Color: "#FF0000",
		},
	}

	style1 := manager.GetStyle(style1Config)
	style2 := manager.GetStyle(style2Config)
	style3 := manager.GetStyle(style3Config)

	// All should be different instances
	assert.NotSame(t, style1, style2, "Expected different StyleFlyweight instances")
	assert.NotSame(t, style1, style3, "Expected different StyleFlyweight instances")
	assert.NotSame(t, style2, style3, "Expected different StyleFlyweight instances")

	// Cache should contain 3 unique styles
	stats := manager.GetCacheStats()
	assert.Equal(t, 3, stats.TotalStyles, "Expected 3 unique styles in cache")
}

// Test Case 8.5: Cache Key Generation
func TestStyleManager_CacheKeyGeneration(t *testing.T) {
	// Test: Check that cache key generation is consistent and unique
	// Expected:
	// - Same config generates same key
	// - Different configs generate different keys
	// - Keys are deterministic

	manager := excelbuilder.NewStyleManager()

	styleConfig := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
			Color: "#000000",
		},
		Fill: excelbuilder.FillConfig{
			Type:  "pattern",
			Color: "#FFFFFF",
		},
	}

	// Generate key multiple times
	key1 := manager.GenerateCacheKey(styleConfig)
	key2 := manager.GenerateCacheKey(styleConfig)
	key3 := manager.GenerateCacheKey(styleConfig)

	// All keys should be identical
	assert.Equal(t, key1, key2, "Expected consistent cache key generation")
	assert.Equal(t, key2, key3, "Expected consistent cache key generation")

	// Different config should generate different key
	differentConfig := styleConfig
	differentConfig.Font.Size = 14
	differentKey := manager.GenerateCacheKey(differentConfig)

	assert.NotEqual(t, key1, differentKey, "Expected different cache keys for different configs")
}

// Test Case 8.6: Memory Usage Optimization
func TestStyleManager_MemoryOptimization(t *testing.T) {
	// Test: Check memory usage with many similar styles
	// Expected:
	// - Memory usage stays low with style reuse
	// - Cache prevents memory bloat
	// - Performance remains good

	manager := excelbuilder.NewStyleManager()

	// Create many requests for the same style
	styleConfig := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{
			Bold: true,
			Size: 12,
		},
	}

	const numRequests = 1000
	styles := make([]*excelbuilder.StyleFlyweight, numRequests)

	for i := 0; i < numRequests; i++ {
		styles[i] = manager.GetStyle(styleConfig)
	}

	// All should be the same instance
	firstStyle := styles[0]
	for i, style := range styles {
		assert.Same(t, firstStyle, style, "Expected same StyleFlyweight instance at index %d", i)
	}

	// Cache should contain only 1 style despite 1000 requests
	stats := manager.GetCacheStats()
	assert.Equal(t, 1, stats.TotalStyles, "Expected 1 unique style in cache")
	assert.Equal(t, numRequests-1, stats.CacheHits, "Expected %d cache hits", numRequests-1)
	assert.Equal(t, 1, stats.CacheMisses, "Expected 1 cache miss")
}