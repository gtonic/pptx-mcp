"""
Performance Optimization Module

Provides performance monitoring and optimization utilities for the PowerPoint MCP Server
to handle large presentations efficiently.
"""
import time
import functools
import logging
from typing import Any, Callable, Dict, Optional, List
import gc
import os

logger = logging.getLogger(__name__)

# Try to import psutil, fall back to basic monitoring if not available
try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False
    logger.warning("psutil not available, using basic performance monitoring")


class PerformanceMonitor:
    """Monitors and optimizes performance of PowerPoint operations."""
    
    def __init__(self):
        self.operation_stats: Dict[str, Dict[str, Any]] = {}
        self.memory_threshold_mb = 500  # MB
        self.slide_count_threshold = 50  # slides
    
    def track_operation(self, operation_name: str):
        """Decorator to track operation performance."""
        def decorator(func: Callable) -> Callable:
            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                start_time = time.time()
                start_memory = self._get_memory_usage()
                
                try:
                    result = func(*args, **kwargs)
                    success = True
                    error = None
                except Exception as e:
                    result = None
                    success = False
                    error = str(e)
                    raise
                finally:
                    end_time = time.time()
                    end_memory = self._get_memory_usage()
                    
                    # Record statistics
                    duration = end_time - start_time
                    memory_delta = end_memory - start_memory
                    
                    self._record_stats(operation_name, duration, memory_delta, success, error)
                    
                    # Log performance warnings
                    if duration > 10.0:  # 10 seconds
                        logger.warning(f"Slow operation {operation_name}: {duration:.2f}s")
                    
                    if memory_delta > 100:  # 100 MB
                        logger.warning(f"High memory operation {operation_name}: +{memory_delta:.1f}MB")
                
                return result
            return wrapper
        return decorator
    
    def _get_memory_usage(self) -> float:
        """Get current memory usage in MB."""
        if HAS_PSUTIL:
            try:
                process = psutil.Process(os.getpid())
                return process.memory_info().rss / 1024 / 1024  # Convert to MB
            except:
                pass
        
        # Fallback: basic memory estimation
        try:
            import resource
            return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024  # KB to MB on Linux
        except:
            return 0.0  # Unable to get memory info
    
    def _record_stats(self, operation: str, duration: float, memory_delta: float, 
                     success: bool, error: Optional[str]):
        """Record operation statistics."""
        if operation not in self.operation_stats:
            self.operation_stats[operation] = {
                'count': 0,
                'total_duration': 0.0,
                'total_memory_delta': 0.0,
                'success_count': 0,
                'avg_duration': 0.0,
                'avg_memory_delta': 0.0,
                'last_error': None
            }
        
        stats = self.operation_stats[operation]
        stats['count'] += 1
        stats['total_duration'] += duration
        stats['total_memory_delta'] += memory_delta
        
        if success:
            stats['success_count'] += 1
        else:
            stats['last_error'] = error
        
        # Update averages
        stats['avg_duration'] = stats['total_duration'] / stats['count']
        stats['avg_memory_delta'] = stats['total_memory_delta'] / stats['count']
    
    def get_performance_report(self) -> Dict[str, Any]:
        """Get comprehensive performance report."""
        current_memory = self._get_memory_usage()
        
        return {
            'current_memory_mb': current_memory,
            'memory_threshold_mb': self.memory_threshold_mb,
            'memory_warning': current_memory > self.memory_threshold_mb,
            'operation_stats': self.operation_stats,
            'recommendations': self._get_recommendations()
        }
    
    def _get_recommendations(self) -> List[str]:
        """Get performance recommendations based on current stats."""
        recommendations = []
        current_memory = self._get_memory_usage()
        
        if current_memory > self.memory_threshold_mb:
            recommendations.append(f"High memory usage ({current_memory:.1f}MB). Consider optimizing large presentations.")
        
        for operation, stats in self.operation_stats.items():
            if stats['avg_duration'] > 5.0:
                recommendations.append(f"Operation '{operation}' is slow (avg {stats['avg_duration']:.2f}s). Consider breaking into smaller operations.")
            
            success_rate = stats['success_count'] / stats['count'] if stats['count'] > 0 else 0
            if success_rate < 0.95 and stats['count'] > 5:
                recommendations.append(f"Operation '{operation}' has low success rate ({success_rate:.1%}). Check error handling.")
        
        return recommendations
    
    def optimize_large_presentation(self, slide_count: int) -> Dict[str, Any]:
        """Provide optimization suggestions for large presentations."""
        recommendations = []
        
        if slide_count > self.slide_count_threshold:
            recommendations.extend([
                "Consider processing slides in batches",
                "Use lazy loading for slide content",
                "Implement slide caching for repeated operations",
                "Consider splitting into multiple smaller presentations"
            ])
        
        if slide_count > 100:
            recommendations.extend([
                "Enable memory cleanup after each batch",
                "Use streaming for large data operations",
                "Consider background processing for complex operations"
            ])
        
        return {
            'slide_count': slide_count,
            'is_large_presentation': slide_count > self.slide_count_threshold,
            'recommendations': recommendations,
            'suggested_batch_size': min(10, max(1, 50 // (slide_count // 50 + 1)))
        }
    
    def cleanup_memory(self):
        """Force garbage collection to free memory."""
        gc.collect()
        logger.info("Performed memory cleanup")


class BatchProcessor:
    """Handles batch processing of large presentation operations."""
    
    def __init__(self, batch_size: int = 10):
        self.batch_size = batch_size
    
    def process_slides_in_batches(self, slides_data: List[Any], 
                                 processor_func: Callable, 
                                 *args, **kwargs) -> List[Any]:
        """Process slides in batches to optimize memory usage."""
        results = []
        total_slides = len(slides_data)
        
        logger.info(f"Processing {total_slides} slides in batches of {self.batch_size}")
        
        for i in range(0, total_slides, self.batch_size):
            batch = slides_data[i:i + self.batch_size]
            batch_start = i + 1
            batch_end = min(i + self.batch_size, total_slides)
            
            logger.info(f"Processing batch {batch_start}-{batch_end}")
            
            try:
                batch_results = []
                for slide_data in batch:
                    result = processor_func(slide_data, *args, **kwargs)
                    batch_results.append(result)
                
                results.extend(batch_results)
                
                # Memory cleanup after each batch
                if i + self.batch_size < total_slides:
                    gc.collect()
                
            except Exception as e:
                logger.error(f"Error processing batch {batch_start}-{batch_end}: {e}")
                raise
        
        return results


class CacheManager:
    """Manages caching for frequently accessed presentation elements."""
    
    def __init__(self, max_cache_size: int = 100):
        self.cache: Dict[str, Any] = {}
        self.access_times: Dict[str, float] = {}
        self.max_cache_size = max_cache_size
    
    def get(self, key: str) -> Optional[Any]:
        """Get item from cache."""
        if key in self.cache:
            self.access_times[key] = time.time()
            return self.cache[key]
        return None
    
    def set(self, key: str, value: Any) -> None:
        """Set item in cache with LRU eviction."""
        current_time = time.time()
        
        if len(self.cache) >= self.max_cache_size:
            # Remove least recently used item
            lru_key = min(self.access_times.keys(), key=self.access_times.get)
            del self.cache[lru_key]
            del self.access_times[lru_key]
        
        self.cache[key] = value
        self.access_times[key] = current_time
    
    def clear(self) -> None:
        """Clear all cached items."""
        self.cache.clear()
        self.access_times.clear()
    
    def get_stats(self) -> Dict[str, Any]:
        """Get cache statistics."""
        return {
            'cache_size': len(self.cache),
            'max_cache_size': self.max_cache_size,
            'hit_rate': getattr(self, '_hit_rate', 0.0),
            'memory_usage_estimate': sum(len(str(v)) for v in self.cache.values())
        }


# Global instances
performance_monitor = PerformanceMonitor()
batch_processor = BatchProcessor()
cache_manager = CacheManager()