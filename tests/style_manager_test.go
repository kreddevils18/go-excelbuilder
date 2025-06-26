package excelbuilder_test

import (
	"sync"
	"testing"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

func TestStyleManager_GetStyle_CacheHit(t *testing.T) {
	sm := excelbuilder.NewStyleManager()
	file := excelize.NewFile()
	config := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 12},
	}

	// First call - should be a cache miss
	style1 := sm.GetStyle(config, file)
	assert.NotNil(t, style1, "First style should not be nil")

	// Second call with same config - should be a cache hit
	style2 := sm.GetStyle(config, file)
	assert.NotNil(t, style2, "Second style should not be nil")

	// Assert that the same instance is returned
	assert.Same(t, style1, style2, "Expected the same style instance to be returned from cache")

	stats := sm.GetCacheStats()
	assert.Equal(t, 1, stats.CacheHits, "Expected 1 cache hit")
	assert.Equal(t, 1, stats.CacheMisses, "Expected 1 cache miss")
	assert.Equal(t, uint64(1), stats.UniqueStyles, "Expected 1 unique style")
}

func TestStyleManager_GetStyle_CacheMiss(t *testing.T) {
	sm := excelbuilder.NewStyleManager()
	file := excelize.NewFile()
	config1 := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 12},
	}
	config2 := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Italic: true, Size: 10},
	}

	// First call
	style1 := sm.GetStyle(config1, file)
	assert.NotNil(t, style1)

	// Second call with different config
	style2 := sm.GetStyle(config2, file)
	assert.NotNil(t, style2)

	assert.NotSame(t, style1, style2, "Expected different style instances for different configs")

	stats := sm.GetCacheStats()
	assert.Equal(t, 0, stats.CacheHits, "Expected 0 cache hits")
	assert.Equal(t, 2, stats.CacheMisses, "Expected 2 cache misses")
	assert.Equal(t, uint64(2), stats.UniqueStyles, "Expected 2 unique styles")
}

func TestStyleManager_Concurrency(t *testing.T) {
	sm := excelbuilder.NewStyleManager()
	file := excelize.NewFile()
	config := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true, Size: 14, Color: "FF0000"},
	}

	var wg sync.WaitGroup
	numGoroutines := 100
	styles := make(chan *excelbuilder.StyleFlyweight, numGoroutines)

	// Launch multiple goroutines to get the same style
	for i := 0; i < numGoroutines; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			style := sm.GetStyle(config, file)
			styles <- style
		}()
	}

	wg.Wait()
	close(styles)

	// Check that all goroutines received the same instance
	var firstStyle *excelbuilder.StyleFlyweight
	for style := range styles {
		assert.NotNil(t, style)
		if firstStyle == nil {
			firstStyle = style
		}
		assert.Same(t, firstStyle, style, "All goroutines should receive the same style instance")
	}

	// Check cache stats
	stats := sm.GetCacheStats()
	assert.Equal(t, 1, stats.CacheMisses, "Should have only one cache miss for the first access")
	assert.Equal(t, numGoroutines-1, stats.CacheHits, "Should have subsequent accesses as cache hits")
	assert.Equal(t, uint64(1), stats.UniqueStyles, "Should only have one unique style in the cache")
}

func TestStyleManager_GenerateCacheKey(t *testing.T) {
	sm := excelbuilder.NewStyleManager()
	config1 := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true},
	}
	config2 := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: true},
	}
	config3 := excelbuilder.StyleConfig{
		Font: excelbuilder.FontConfig{Bold: false},
	}

	key1 := sm.GenerateCacheKey(config1)
	key2 := sm.GenerateCacheKey(config2)
	key3 := sm.GenerateCacheKey(config3)

	assert.NotEmpty(t, key1)
	assert.Equal(t, key1, key2, "Cache keys for identical configs should be the same")
	assert.NotEqual(t, key1, key3, "Cache keys for different configs should be different")
}
