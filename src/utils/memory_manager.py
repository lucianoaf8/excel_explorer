"""
Resource optimization utilities for Excel Explorer
Provides memory monitoring, cache management, and resource profiling.
"""

import psutil
import time
import threading
import logging
from typing import Dict, Optional, Callable, Any
from collections import OrderedDict, defaultdict
from pathlib import Path


class MemoryManager:
    """Manages system resources and memory usage for Excel analysis."""

    def __init__(self, max_memory_mb: int = 4096, warning_threshold: float = 0.8):
        self.max_memory_mb = max_memory_mb
        self.warning_threshold = warning_threshold
        self.process = psutil.Process()
        self.baseline_memory = self._get_memory_mb()
        
        # Resource tracking
        self.module_stats: Dict[str, Dict[str, Any]] = defaultdict(dict)
        self.peak_memory = self.baseline_memory
        self.start_time = time.time()
        
        # Cache management
        self._cache: OrderedDict = OrderedDict()
        self._cache_size_bytes = 0
        self._max_cache_mb = max_memory_mb // 4  # 25% of max memory for cache
        self._lock = threading.Lock()
        
        logging.getLogger(__name__).info(
            f"MemoryManager initialized: baseline={self.baseline_memory:.1f}MB, "
            f"max={max_memory_mb}MB, warning_threshold={warning_threshold:.1%}"
        )

    def _get_memory_mb(self) -> float:
        """Get current process memory usage in MB."""
        return self.process.memory_info().rss / 1_048_576

    def get_current_usage(self) -> Dict[str, float]:
        """Get current resource usage statistics."""
        current_mb = self._get_memory_mb()
        self.peak_memory = max(self.peak_memory, current_mb)
        
        return {
            'current_mb': current_mb,
            'peak_mb': self.peak_memory,
            'baseline_mb': self.baseline_memory,
            'delta_mb': current_mb - self.baseline_memory,
            'cpu_percent': self.process.cpu_percent(),
            'usage_ratio': current_mb / self.max_memory_mb,
            'elapsed_seconds': time.time() - self.start_time
        }

    def check_memory_pressure(self) -> bool:
        """Check if memory usage exceeds warning threshold."""
        usage = self.get_current_usage()
        return usage['usage_ratio'] > self.warning_threshold

    def enforce_memory_limit(self) -> None:
        """Raise exception if memory limit exceeded."""
        usage = self.get_current_usage()
        if usage['current_mb'] > self.max_memory_mb:
            raise MemoryError(
                f"Memory limit exceeded: {usage['current_mb']:.1f}MB > {self.max_memory_mb}MB"
            )

    def track_module_start(self, module_name: str, file_size_mb: float = 0) -> None:
        """Start tracking resource usage for a module."""
        with self._lock:
            self.module_stats[module_name] = {
                'start_time': time.time(),
                'start_memory_mb': self._get_memory_mb(),
                'file_size_mb': file_size_mb,
                'completed': False
            }

    def track_module_end(self, module_name: str) -> Dict[str, Any]:
        """End tracking and return module resource statistics."""
        with self._lock:
            if module_name not in self.module_stats:
                logging.warning(f"Module {module_name} not found in tracking")
                return {}
            
            stats = self.module_stats[module_name]
            end_time = time.time()
            end_memory = self._get_memory_mb()
            
            stats.update({
                'end_time': end_time,
                'end_memory_mb': end_memory,
                'duration_seconds': end_time - stats['start_time'],
                'memory_delta_mb': end_memory - stats['start_memory_mb'],
                'completed': True
            })
            
            # Calculate efficiency metrics
            if stats['file_size_mb'] > 0:
                stats['mb_per_second'] = stats['file_size_mb'] / stats['duration_seconds']
                stats['memory_multiplier'] = stats['memory_delta_mb'] / stats['file_size_mb']
            
            return stats.copy()

    def estimate_processing_time(self, file_size_mb: float, complexity_factor: float = 1.0) -> float:
        """Estimate processing time based on historical data and file characteristics."""
        # Get average processing rates from completed modules
        completed_modules = [
            stats for stats in self.module_stats.values() 
            if stats.get('completed') and stats.get('mb_per_second', 0) > 0
        ]
        
        if not completed_modules:
            # Default estimates: 0.5 seconds per MB for simple analysis
            return file_size_mb * 0.5 * complexity_factor
        
        avg_rate = sum(s['mb_per_second'] for s in completed_modules) / len(completed_modules)
        estimated_seconds = (file_size_mb / avg_rate) * complexity_factor
        
        return max(estimated_seconds, 1.0)  # Minimum 1 second

    def cache_result(self, key: str, value: Any, size_bytes: int) -> bool:
        """Cache analysis result with LRU eviction."""
        with self._lock:
            # Check if adding this item would exceed cache limit
            max_cache_bytes = self._max_cache_mb * 1_048_576
            
            if size_bytes > max_cache_bytes:
                logging.warning(f"Cache item too large: {size_bytes} bytes > {max_cache_bytes}")
                return False
            
            # Evict items if necessary
            while (self._cache_size_bytes + size_bytes > max_cache_bytes and self._cache):
                self._evict_lru_item()
            
            # Add new item
            if key in self._cache:
                # Update existing item
                old_size = self._cache[key]['size_bytes']
                self._cache_size_bytes -= old_size
            
            self._cache[key] = {
                'value': value,
                'size_bytes': size_bytes,
                'timestamp': time.time()
            }
            self._cache.move_to_end(key)  # Mark as most recently used
            self._cache_size_bytes += size_bytes
            
            return True

    def get_cached_result(self, key: str) -> Optional[Any]:
        """Retrieve cached result and update LRU order."""
        with self._lock:
            if key not in self._cache:
                return None
            
            # Move to end (most recently used)
            self._cache.move_to_end(key)
            return self._cache[key]['value']

    def _evict_lru_item(self) -> None:
        """Remove least recently used cache item."""
        if not self._cache:
            return
        
        key, item = self._cache.popitem(last=False)  # Remove oldest
        self._cache_size_bytes -= item['size_bytes']
        logging.debug(f"Evicted cache item: {key} ({item['size_bytes']} bytes)")

    def clear_cache(self) -> None:
        """Clear all cached items."""
        with self._lock:
            self._cache.clear()
            self._cache_size_bytes = 0

    def get_resource_report(self) -> Dict[str, Any]:
        """Generate comprehensive resource usage report."""
        usage = self.get_current_usage()
        
        # Module statistics summary
        module_summary = {}
        for name, stats in self.module_stats.items():
            if stats.get('completed'):
                module_summary[name] = {
                    'duration_seconds': stats['duration_seconds'],
                    'memory_delta_mb': stats['memory_delta_mb'],
                    'efficiency_mb_per_sec': stats.get('mb_per_second', 0)
                }
        
        return {
            'current_usage': usage,
            'module_statistics': module_summary,
            'cache_status': {
                'items': len(self._cache),
                'size_mb': self._cache_size_bytes / 1_048_576,
                'max_size_mb': self._max_cache_mb
            },
            'system_info': {
                'total_system_memory_gb': psutil.virtual_memory().total / 1_073_741_824,
                'available_memory_gb': psutil.virtual_memory().available / 1_073_741_824
            }
        }

    def should_reduce_processing_depth(self) -> bool:
        """Determine if processing depth should be reduced due to resource constraints."""
        usage = self.get_current_usage()
        return (
            usage['usage_ratio'] > 0.7 or  # Above 70% memory usage
            usage['cpu_percent'] > 90 or   # High CPU usage
            self.check_memory_pressure()
        )


class ResourceMonitor:
    """Context manager for automatic resource tracking."""
    
    def __init__(self, memory_manager: MemoryManager, module_name: str, file_size_mb: float = 0):
        self.memory_manager = memory_manager
        self.module_name = module_name
        self.file_size_mb = file_size_mb
        self.stats = {}

    def __enter__(self):
        self.memory_manager.track_module_start(self.module_name, self.file_size_mb)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.stats = self.memory_manager.track_module_end(self.module_name)
        
        # Log completion
        if exc_type is None:
            logging.info(
                f"Module {self.module_name} completed: "
                f"{self.stats.get('duration_seconds', 0):.1f}s, "
                f"{self.stats.get('memory_delta_mb', 0):+.1f}MB"
            )
        else:
            logging.error(f"Module {self.module_name} failed: {exc_val}")


# Global instance for easy access
_global_memory_manager: Optional[MemoryManager] = None

def get_memory_manager() -> MemoryManager:
    """Get or create global memory manager instance."""
    global _global_memory_manager
    if _global_memory_manager is None:
        _global_memory_manager = MemoryManager()
    return _global_memory_manager

def initialize_memory_manager(max_memory_mb: int = 4096, warning_threshold: float = 0.8) -> MemoryManager:
    """Initialize global memory manager with custom settings."""
    global _global_memory_manager
    _global_memory_manager = MemoryManager(max_memory_mb, warning_threshold)
    return _global_memory_manager
