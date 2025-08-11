"""
Core functionality for Excel Explorer

Contains the main analysis engine and configuration management.
"""

from .analyzer import SimpleExcelAnalyzer
from .config import load_config, get_config_value

__all__ = ["SimpleExcelAnalyzer", "load_config", "get_config_value"]
