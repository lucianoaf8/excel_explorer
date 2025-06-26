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
import gc


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
    
    def get_optimal_chunk_size(self, base_chunk_size: int, file_size_mb: float) -> int:
        """Calculate optimal chunk size based on current memory conditions
        
        Args:
            base_chunk_size: Default chunk size
            file_size_mb: Size of file being processed
            
        Returns:
            Optimized chunk size
        """
        usage = self.get_current_usage()
        
        # Reduce chunk size under memory pressure
        if usage['usage_ratio'] > 0.8:
            multiplier = 0.5
        elif usage['usage_ratio'] > 0.6:
            multiplier = 0.7
        else:
            multiplier = 1.0
        
        # Adjust for file size
        if file_size_mb > 100:
            multiplier *= 0.8
        elif file_size_mb > 200:
            multiplier *= 0.6
        
        optimal_size = int(base_chunk_size * multiplier)
        return max(100, optimal_size)  # Minimum chunk size
    
    def trigger_garbage_collection(self) -> bool:
        """Trigger garbage collection if memory pressure detected
        
        Returns:
            bool: True if GC was triggered
        """
        import gc
        
        if self.check_memory_pressure():
            logging.info("Triggering garbage collection due to memory pressure")
            collected = gc.collect()
            logging.debug(f"Garbage collection freed {collected} objects")
            return True
        return False
    
    def adaptive_cache_cleanup(self) -> None:
        """Perform adaptive cache cleanup based on memory pressure"""
        usage = self.get_current_usage()
        
        if usage['usage_ratio'] > 0.9:
            # Aggressive cleanup
            self.clear_cache()
            self.trigger_garbage_collection()
        elif usage['usage_ratio'] > 0.8:
            # Moderate cleanup - remove 50% of cache
            with self._lock:
                items_to_remove = len(self._cache) // 2
                for _ in range(items_to_remove):
                    if self._cache:
                        self._evict_lru_item()
        elif usage['usage_ratio'] > 0.7:
            # Light cleanup - remove oldest 25% of cache
            with self._lock:
                items_to_remove = len(self._cache) // 4
                for _ in range(items_to_remove):
                    if self._cache:
                        self._evict_lru_item()


class ResourceMonitor:
    """Context manager for automatic resource tracking with adaptive management."""
    
    def __init__(self, memory_manager: MemoryManager, module_name: str, file_size_mb: float = 0):
        self.memory_manager = memory_manager
        self.module_name = module_name
        self.file_size_mb = file_size_mb
        self.stats = {}
        self.initial_usage = None
        self.cleanup_triggered = False

    def __enter__(self):
        self.memory_manager.track_module_start(self.module_name, self.file_size_mb)
        self.initial_usage = self.memory_manager.get_current_usage()
        
        # Proactive cleanup if starting with high memory usage
        if self.initial_usage['usage_ratio'] > 0.7:
            self.memory_manager.adaptive_cache_cleanup()
            
        return self
    
    def check_memory_during_processing(self) -> bool:
        """Check memory status during processing and trigger cleanup if needed
        
        Returns:
            bool: True if processing should continue, False if should abort
        """
        current_usage = self.memory_manager.get_current_usage()
        
        if current_usage['usage_ratio'] > 0.95:
            logging.warning(f"Critical memory usage in {self.module_name}: {current_usage['usage_ratio']:.1%}")
            return False
        elif current_usage['usage_ratio'] > 0.85 and not self.cleanup_triggered:
            logging.info(f"High memory usage in {self.module_name}, triggering cleanup")
            self.memory_manager.adaptive_cache_cleanup()
            self.cleanup_triggered = True
            
        return True
    
    def get_recommended_chunk_size(self, default_size: int) -> int:
        """Get recommended chunk size based on current memory conditions"""
        return self.memory_manager.get_optimal_chunk_size(default_size, self.file_size_mb)

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.stats = self.memory_manager.track_module_end(self.module_name)
        final_usage = self.memory_manager.get_current_usage()
        
        # Log completion with memory efficiency metrics
        if exc_type is None:
            efficiency_ratio = self.stats.get('memory_delta_mb', 0) / max(1, self.file_size_mb)
            logging.info(
                f"Module {self.module_name} completed: "
                f"{self.stats.get('duration_seconds', 0):.1f}s, "
                f"{self.stats.get('memory_delta_mb', 0):+.1f}MB "
                f"(efficiency: {efficiency_ratio:.1f}x file size)"
            )
        else:
            logging.error(f"Module {self.module_name} failed: {exc_val}")
            
        # Final cleanup if module used significant memory
        if self.stats.get('memory_delta_mb', 0) > 100:  # >100MB delta
            self.memory_manager.adaptive_cache_cleanup()


# Global instance for easy access
_global_memory_manager: Optional[MemoryManager] = None


def configure_chunked_processing_defaults(file_size_mb: float) -> Dict[str, Any]:
    """Configure default chunked processing parameters based on file size
    
    Args:
        file_size_mb: Size of file to be processed
        
    Returns:
        Dict with recommended processing parameters
    """
    memory_manager = get_memory_manager()
    current_usage = memory_manager.get_current_usage()
    
    # Base parameters
    if file_size_mb < 10:
        base_chunk_size = 10000
        strategy = "row_based"
        max_memory_mb = 256
    elif file_size_mb < 50:
        base_chunk_size = 5000
        strategy = "row_based"
        max_memory_mb = 512
    elif file_size_mb < 100:
        base_chunk_size = 2000
        strategy = "adaptive"
        max_memory_mb = 1024
    else:
        base_chunk_size = 1000
        strategy = "adaptive"
        max_memory_mb = 2048
    
    # Adjust for current memory pressure
    optimal_chunk_size = memory_manager.get_optimal_chunk_size(base_chunk_size, file_size_mb)
    
    return {
        'chunk_size_rows': optimal_chunk_size,
        'strategy': strategy,
        'max_memory_mb': min(max_memory_mb, memory_manager.max_memory_mb // 2),
        'enable_progress_tracking': file_size_mb > 20,
        'intermediate_save': file_size_mb > 200,
        'memory_pressure_level': current_usage['usage_ratio']
    }

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
    
    # Log initialization with system context
    import psutil
    system_memory_gb = psutil.virtual_memory().total / 1_073_741_824
    logging.info(
        f"Memory manager initialized: {max_memory_mb}MB limit "
        f"({max_memory_mb/1024:.1f}GB) on {system_memory_gb:.1f}GB system"
    )
    
    return _global_memory_manager
