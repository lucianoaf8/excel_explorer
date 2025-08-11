#!/usr/bin/env python3
"""
Simplified Configuration Module for Excel Explorer
Replaces the complex ConfigManager with a simple function-based approach
"""

import os
import yaml
from pathlib import Path
from typing import Dict, Any


DEFAULT_CONFIG = {
    'analysis': {
        'sample_rows': 100,
        'max_formula_check': 1000,
        'memory_limit_mb': 500,
        'enable_cross_sheet_analysis': True,
        'timeout_seconds': 300,
        'detail_level': 'standard'
    },
    'reporting': {
        'output_format': 'html',
        'include_charts': True,
        'include_raw_data': False,
        'template_path': None
    },
    'logging': {
        'level': 'INFO',
        'console_output': True,
        'file_output': False
    }
}


def deep_merge(base: Dict, override: Dict) -> Dict:
    """Deep merge two dictionaries"""
    result = base.copy()
    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def load_config(config_path: str = None) -> Dict[str, Any]:
    """
    Load configuration with simple environment overrides
    
    Args:
        config_path: Path to YAML config file (optional)
        
    Returns:
        Complete configuration dictionary
    """
    # Start with defaults
    config = DEFAULT_CONFIG.copy()
    
    # Load from file if exists
    if config_path and Path(config_path).exists():
        try:
            with open(config_path) as f:
                file_config = yaml.safe_load(f) or {}
                config = deep_merge(config, file_config)
        except Exception as e:
            print(f"Warning: Failed to load config from {config_path}: {e}")
    
    # Apply simple env overrides (only for key settings)
    env_overrides = {
        'EXCEL_EXPLORER_SAMPLE_ROWS': ('analysis', 'sample_rows', int),
        'EXCEL_EXPLORER_MEMORY_LIMIT': ('analysis', 'memory_limit_mb', int),
        'EXCEL_EXPLORER_MAX_FORMULA_CHECK': ('analysis', 'max_formula_check', int),
        'EXCEL_EXPLORER_OUTPUT_FORMAT': ('reporting', 'output_format', str),
        'EXCEL_EXPLORER_LOG_LEVEL': ('logging', 'level', str)
    }
    
    for env_var, (section, key, type_fn) in env_overrides.items():
        value = os.getenv(env_var)
        if value:
            try:
                config[section][key] = type_fn(value)
            except (ValueError, TypeError):
                print(f"Warning: Invalid value for {env_var}: {value}")
    
    return config


def get_config_value(config: Dict[str, Any], section: str, key: str, default=None):
    """
    Helper function to safely get nested config values
    
    Args:
        config: Configuration dictionary
        section: Section name (e.g., 'analysis')
        key: Key name (e.g., 'sample_rows')
        default: Default value if not found
        
    Returns:
        Configuration value or default
    """
    return config.get(section, {}).get(key, default)