package excelbuilder

import (
	"crypto/md5"
	"encoding/json"
	"fmt"
	"sync"
)

// StyleManager manages style flyweights using the Flyweight pattern
// It provides thread-safe caching and reuse of style objects
type StyleManager struct {
	cache map[string]*StyleFlyweight
	mutex sync.RWMutex
	stats CacheStats
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
		cache: make(map[string]*StyleFlyweight),
		stats: CacheStats{},
	}
}

// GetStyle returns a StyleFlyweight for the given configuration
// Uses caching to ensure memory efficiency and performance
func (sm *StyleManager) GetStyle(config StyleConfig) *StyleFlyweight {
	cacheKey := sm.GenerateCacheKey(config)

	// Try to get from cache with read lock
	sm.mutex.RLock()
	if flyweight, exists := sm.cache[cacheKey]; exists {
		sm.mutex.RUnlock()
		
		// Update stats with write lock
		sm.mutex.Lock()
		sm.stats.CacheHits++
		sm.mutex.Unlock()
		
		return flyweight
	}
	sm.mutex.RUnlock()

	// Create new flyweight with write lock
	sm.mutex.Lock()
	defer sm.mutex.Unlock()

	// Double-check pattern - another goroutine might have created it
	if flyweight, exists := sm.cache[cacheKey]; exists {
		sm.stats.CacheHits++
		return flyweight
	}

	// Create new flyweight
	flyweight := NewStyleFlyweight(config)
	sm.cache[cacheKey] = flyweight
	sm.stats.CacheMisses++
	sm.stats.TotalStyles++

	return flyweight
}

// GenerateCacheKey generates a unique cache key for a style configuration
func (sm *StyleManager) GenerateCacheKey(config StyleConfig) string {
	// Serialize config to JSON for consistent key generation
	jsonData, err := json.Marshal(config)
	if err != nil {
		// Fallback to string representation if JSON fails
		return fmt.Sprintf("%+v", config)
	}

	// Generate MD5 hash for compact, consistent key
	hash := md5.Sum(jsonData)
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
	sm.stats = CacheStats{}
}

// GetCacheSize returns the number of cached styles
func (sm *StyleManager) GetCacheSize() int {
	sm.mutex.RLock()
	defer sm.mutex.RUnlock()
	return len(sm.cache)
}

// GetStyleFlyweight is an alias for GetStyle for backward compatibility
func (sm *StyleManager) GetStyleFlyweight(config StyleConfig) *StyleFlyweight {
	return sm.GetStyle(config)
}