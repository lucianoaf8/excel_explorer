"""
Abstract base class for all analyzer modules
Provides common interface and utilities for analysis modules.
"""

from abc import ABC, abstractmethod

class BaseAnalyzer(ABC):
    def __init__(self, config=None):
        self.config = config or {}
    
    @abstractmethod
    def analyze(self, workbook_data):
        """Perform analysis on workbook data"""
        pass
    
    @abstractmethod
    def get_results(self):
        """Return analysis results"""
        pass
