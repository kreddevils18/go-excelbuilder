package excelbuilder

import (
	"crypto/sha256"
	"fmt"
	"strings"
	"sync"

	"github.com/xuri/excelize/v2"
)

// StyleManager manages style flyweights using the Flyweight pattern
// It provides thread-safe caching and reuse of style objects
type StyleManager struct {
	cache     map[string]*StyleFlyweight
	access    map[string]int64 // Track access time for LRU eviction
	mutex     sync.RWMutex
	stats     CacheStats
	maxSize   int   // Maximum cache size (0 = unlimited)
	counter   int64 // Access counter for LRU
}

// CacheStats provides statistics about the style cache
type CacheStats struct {
	TotalStyles  int
	CacheHits    int
	CacheMisses  int
	UniqueStyles uint64
}

// HitRate returns the cache hit rate as a float between 0 and 1
func (cs CacheStats) HitRate() float64 {
	totalRequests := cs.CacheHits + cs.CacheMisses
	if totalRequests == 0 {
		return 0
	}
	return float64(cs.CacheHits) / float64(totalRequests)
}

// TotalRequests returns the total number of cache requests
func (cs CacheStats) TotalRequests() int {
	return cs.CacheHits + cs.CacheMisses
}

// NewStyleManager creates a new StyleManager instance
func NewStyleManager() *StyleManager {
	return &StyleManager{
		cache:   make(map[string]*StyleFlyweight),
		access:  make(map[string]int64),
		stats:   CacheStats{},
		maxSize: 1000, // Default cache size limit
		counter: 0,
	}
}

// SetMaxCacheSize sets the maximum cache size (0 = unlimited)
func (sm *StyleManager) SetMaxCacheSize(maxSize int) {
	sm.mutex.Lock()
	defer sm.mutex.Unlock()
	sm.maxSize = maxSize
	if maxSize > 0 && len(sm.cache) > maxSize {
		sm.evictLRU(len(sm.cache) - maxSize)
	}
}

// GetMaxCacheSize returns the current maximum cache size
func (sm *StyleManager) GetMaxCacheSize() int {
	sm.mutex.RLock()
	defer sm.mutex.RUnlock()
	return sm.maxSize
}

// GetStyle returns a StyleFlyweight for the given configuration
// Uses caching to ensure memory efficiency and performance
func (sm *StyleManager) GetStyle(config StyleConfig, file *excelize.File) *StyleFlyweight {
	cacheKey := sm.GenerateCacheKey(config)

	// Try to get from cache with write lock for atomic operations
	sm.mutex.Lock()
	if flyweight, exists := sm.cache[cacheKey]; exists {
		sm.stats.CacheHits++
		sm.counter++
		sm.access[cacheKey] = sm.counter // Update access time
		sm.mutex.Unlock()
		return flyweight
	}

	// Create new flyweight
	style := convertToExcelizeStyle(config)
	styleID, err := file.NewStyle(&style)
	if err != nil {
		sm.mutex.Unlock()
		// Handle error, maybe return a default style or an error
		return nil
	}

	flyweight := NewStyleFlyweight(config, styleID)
	
	// Check if we need to evict before adding new entry
	if sm.maxSize > 0 && len(sm.cache) >= sm.maxSize {
		sm.evictLRU(1)
	}
	
	sm.cache[cacheKey] = flyweight
	sm.counter++
	sm.access[cacheKey] = sm.counter
	sm.stats.CacheMisses++
	sm.stats.TotalStyles++
	sm.mutex.Unlock()

	return flyweight
}

// evictLRU removes the least recently used entries from the cache
// Must be called with mutex already held
func (sm *StyleManager) evictLRU(count int) {
	if count <= 0 || len(sm.cache) == 0 {
		return
	}
	
	// Find the least recently used entries
	type cacheEntry struct {
		key        string
		accessTime int64
	}
	
	entries := make([]cacheEntry, 0, len(sm.access))
	for key, accessTime := range sm.access {
		entries = append(entries, cacheEntry{key: key, accessTime: accessTime})
	}
	
	// Sort by access time (ascending - oldest first)
	for i := 0; i < len(entries)-1; i++ {
		for j := i + 1; j < len(entries); j++ {
			if entries[i].accessTime > entries[j].accessTime {
				entries[i], entries[j] = entries[j], entries[i]
			}
		}
	}
	
	// Remove the oldest entries
	evicted := 0
	for i := 0; i < len(entries) && evicted < count; i++ {
		key := entries[i].key
		if _, exists := sm.cache[key]; exists {
			delete(sm.cache, key)
			delete(sm.access, key)
			evicted++
		}
	}
}

// GenerateCacheKey generates a unique cache key for a style configuration
// Optimized version that avoids JSON marshaling for better performance
func (sm *StyleManager) GenerateCacheKey(config StyleConfig) string {
	var builder strings.Builder
	builder.Grow(256) // Pre-allocate some capacity for performance
	
	// Font configuration
	if config.Font != (FontConfig{}) {
		builder.WriteString("f:")
		if config.Font.Size > 0 {
			builder.WriteString(fmt.Sprintf("sz%d,", config.Font.Size))
		}
		if config.Font.Bold {
			builder.WriteString("b,")
		}
		if config.Font.Italic {
			builder.WriteString("i,")
		}
		if config.Font.Underline {
			builder.WriteString("u,")
		}
		if config.Font.Color != "" {
			builder.WriteString("c")
			builder.WriteString(config.Font.Color)
			builder.WriteString(",")
		}
		if config.Font.Family != "" {
			builder.WriteString("fm")
			builder.WriteString(config.Font.Family)
			builder.WriteString(",")
		}
		builder.WriteString(";")
	}
	
	// Fill configuration
	if config.Fill != (FillConfig{}) {
		builder.WriteString("fl:")
		builder.WriteString(config.Fill.Type)
		builder.WriteString(",")
		builder.WriteString(config.Fill.Color)
		builder.WriteString(";")
	}
	
	// Border configuration
	if config.Border != (BorderConfig{}) {
		builder.WriteString("b:")
		if config.Border.Color != "" {
			builder.WriteString("c")
			builder.WriteString(config.Border.Color)
			builder.WriteString(",")
		}
		if config.Border.Top.Style != "" {
			builder.WriteString("t")
			builder.WriteString(config.Border.Top.Style)
			builder.WriteString(config.Border.Top.Color)
			builder.WriteString(",")
		}
		if config.Border.Bottom.Style != "" {
			builder.WriteString("bt")
			builder.WriteString(config.Border.Bottom.Style)
			builder.WriteString(config.Border.Bottom.Color)
			builder.WriteString(",")
		}
		if config.Border.Left.Style != "" {
			builder.WriteString("l")
			builder.WriteString(config.Border.Left.Style)
			builder.WriteString(config.Border.Left.Color)
			builder.WriteString(",")
		}
		if config.Border.Right.Style != "" {
			builder.WriteString("r")
			builder.WriteString(config.Border.Right.Style)
			builder.WriteString(config.Border.Right.Color)
			builder.WriteString(",")
		}
		builder.WriteString(";")
	}
	
	// Alignment configuration
	if config.Alignment != (AlignmentConfig{}) {
		builder.WriteString("a:")
		builder.WriteString(config.Alignment.Horizontal)
		builder.WriteString(",")
		builder.WriteString(config.Alignment.Vertical)
		builder.WriteString(",")
		if config.Alignment.WrapText {
			builder.WriteString("w,")
		}
		if config.Alignment.TextRotation != 0 {
			builder.WriteString(fmt.Sprintf("r%d,", config.Alignment.TextRotation))
		}
		builder.WriteString(";")
	}
	
	// Protection configuration
	if config.Protection != nil {
		builder.WriteString("p:")
		if config.Protection.Hidden {
			builder.WriteString("h,")
		}
		if config.Protection.Locked {
			builder.WriteString("l,")
		}
		builder.WriteString(";")
	}
	
	// NumberFormat configuration
	if config.NumberFormat != "" {
		builder.WriteString("nf:")
		builder.WriteString(config.NumberFormat)
		builder.WriteString(";")
	}
	
	// For short keys, return directly; for longer keys, hash to keep them manageable
	key := builder.String()
	if len(key) < 100 {
		return key
	}
	
	// Use SHA-256 for longer keys (more secure than MD5)
	hash := sha256.Sum256([]byte(key))
	return fmt.Sprintf("%x", hash)
}

// GetCacheStats returns current cache statistics
func (sm *StyleManager) GetCacheStats() CacheStats {
	sm.mutex.RLock()
	defer sm.mutex.RUnlock()
	stats := sm.stats
	stats.UniqueStyles = uint64(len(sm.cache))
	return stats
}

// ClearCache clears all cached styles (useful for testing)
func (sm *StyleManager) ClearCache() {
	sm.mutex.Lock()
	defer sm.mutex.Unlock()
	sm.cache = make(map[string]*StyleFlyweight)
	sm.access = make(map[string]int64)
	sm.stats = CacheStats{}
	sm.counter = 0
}

// GetCacheSize returns the number of cached styles
func (sm *StyleManager) GetCacheSize() int {
	sm.mutex.RLock()
	defer sm.mutex.RUnlock()
	return len(sm.cache)
}

// GetStyleFlyweight is an alias for GetStyle for backward compatibility
func (sm *StyleManager) GetStyleFlyweight(config StyleConfig, file *excelize.File) *StyleFlyweight {
	return sm.GetStyle(config, file)
}

func convertToExcelizeStyle(config StyleConfig) excelize.Style {
	style := excelize.Style{}

	// Font configuration
	if config.Font != (FontConfig{}) {
		font := &excelize.Font{}
		hasFont := false

		if config.Font.Size > 0 {
			font.Size = float64(config.Font.Size)
			hasFont = true
		}
		if config.Font.Bold {
			font.Bold = true
			hasFont = true
		}
		if config.Font.Italic {
			font.Italic = true
			hasFont = true
		}
		if config.Font.Underline {
			font.Underline = "single"
			hasFont = true
		}
		if config.Font.Color != "" {
			font.Color = config.Font.Color
			hasFont = true
		}
		if config.Font.Family != "" {
			font.Family = config.Font.Family
			hasFont = true
		}

		if hasFont {
			style.Font = font
		}
	}

	// Fill configuration
	if config.Fill != (FillConfig{}) {
		fill := &excelize.Fill{}
		hasFill := false

		if config.Fill.Type == "pattern" && config.Fill.Color != "" {
			fill.Type = "pattern"
			fill.Pattern = 1 // Solid fill
			fill.Color = []string{config.Fill.Color}
			hasFill = true
		}

		if hasFill {
			style.Fill = *fill
		}
	}

	// Border configuration
	if config.Border != (BorderConfig{}) {
		var border []excelize.Border
		if config.Border.Top.Style != "" {
			border = append(border, excelize.Border{
				Type:  "top",
				Style: getBorderStyle(config.Border.Top.Style),
				Color: getColorOrDefault(config.Border.Top.Color, config.Border.Color),
			})
		}
		if config.Border.Bottom.Style != "" {
			border = append(border, excelize.Border{
				Type:  "bottom",
				Style: getBorderStyle(config.Border.Bottom.Style),
				Color: getColorOrDefault(config.Border.Bottom.Color, config.Border.Color),
			})
		}
		if config.Border.Left.Style != "" {
			border = append(border, excelize.Border{
				Type:  "left",
				Style: getBorderStyle(config.Border.Left.Style),
				Color: getColorOrDefault(config.Border.Left.Color, config.Border.Color),
			})
		}
		if config.Border.Right.Style != "" {
			border = append(border, excelize.Border{
				Type:  "right",
				Style: getBorderStyle(config.Border.Right.Style),
				Color: getColorOrDefault(config.Border.Right.Color, config.Border.Color),
			})
		}

		if len(border) > 0 {
			style.Border = border
		}
	}

	// Alignment configuration
	if config.Alignment != (AlignmentConfig{}) {
		alignment := &excelize.Alignment{}
		hasAlignment := false
		if config.Alignment.Horizontal != "" {
			alignment.Horizontal = config.Alignment.Horizontal
			hasAlignment = true
		}
		if config.Alignment.Vertical != "" {
			alignment.Vertical = config.Alignment.Vertical
			hasAlignment = true
		}
		if config.Alignment.WrapText {
			alignment.WrapText = true
			hasAlignment = true
		}
		if config.Alignment.TextRotation != 0 {
			alignment.TextRotation = config.Alignment.TextRotation
			hasAlignment = true
		}
		if hasAlignment {
			style.Alignment = alignment
		}
	}

	// Protection configuration
	if config.Protection != nil {
		protection := &excelize.Protection{
			Hidden: config.Protection.Hidden,
			Locked: config.Protection.Locked,
		}
		style.Protection = protection
	}

	// NumberFormat configuration
	if config.NumberFormat != "" {
		style.NumFmt = 0 // Custom number format requires NumFmt to be set.
		style.CustomNumFmt = &config.NumberFormat
	}

	return style
}

func getBorderStyle(style string) int {
	switch style {
	case "thin":
		return 1
	case "medium":
		return 2
	case "thick":
		return 3
	default:
		return 0
	}
}

func getColorOrDefault(color, defaultColor string) string {
	if color != "" {
		return color
	}
	return defaultColor
}
