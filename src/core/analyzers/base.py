"""
Base analyzer class for all analysis modules
"""

import time
import logging
from abc import ABC, abstractmethod
from typing import Dict, Any
import openpyxl


class BaseAnalyzer(ABC):
    """Base class for all analyzer modules"""
    
    def __init__(self, config: Dict[str, Any]):
        """
        Initialize base analyzer
        
        Args:
            config: Configuration dictionary from load_config()
        """
        self.config = config
        self.analysis_config = config.get('analysis', {})
        self.logger = self._setup_logger()
        self.start_time = None
        
    def _setup_logger(self) -> logging.Logger:
        """Setup logger for this analyzer"""
        logger = logging.getLogger(f'excel_analyzer.{self.__class__.__name__}')
        logger.setLevel(getattr(logging, self.config.get('logging', {}).get('level', 'INFO')))
        return logger
    
    @abstractmethod
    def analyze(self, workbook: openpyxl.Workbook) -> Dict[str, Any]:
        """
        Perform analysis on the workbook
        
        Args:
            workbook: Loaded openpyxl workbook
            
        Returns:
            Dictionary containing analysis results
        """
        pass
    
    def get_sample_limit(self) -> int:
        """Get sample row limit from configuration"""
        return self.analysis_config.get('sample_rows', 100)
    
    def get_memory_limit(self) -> int:
        """Get memory limit from configuration"""
        return self.analysis_config.get('memory_limit_mb', 500)
    
    def start_timing(self):
        """Start timing for performance tracking"""
        self.start_time = time.time()
    
    def get_duration(self) -> float:
        """Get elapsed time since start_timing()"""
        if self.start_time:
            return time.time() - self.start_time
        return 0.0
    
    def log_progress(self, message: str, level: str = 'info'):
        """Log progress message"""
        getattr(self.logger, level)(message)