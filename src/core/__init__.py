"""
Core functionality for Excel Explorer

Contains the main analysis engine and configuration management.
"""

from .analyzer import SimpleExcelAnalyzer
from .config_manager import ConfigManager, get_config

__all__ = ["SimpleExcelAnalyzer", "ConfigManager", "get_config"]
