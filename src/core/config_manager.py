#!/usr/bin/env python3
"""
Centralized Configuration Manager for Excel Explorer
Handles configuration loading, validation, and environment variable overrides
"""

import os
import yaml
from typing import Dict, Any, Optional
from pathlib import Path


class ConfigManager:
    """
    Singleton configuration manager with environment variable support
    
    Provides centralized configuration loading with the following priority:
    1. Environment variables (highest priority)
    2. Configuration file
    3. Default values (lowest priority)
    """
    
    _instance = None
    _config = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def load_config(self, config_path: Optional[str] = None) -> Dict[str, Any]:
        """
        Load configuration with fallbacks and environment overrides
        
        Args:
            config_path: Path to YAML config file (optional)
            
        Returns:
            Complete configuration dictionary
        """
        if self._config is not None:
            return self._config
        
        # Determine config file path with priority:
        # 1. Explicit argument, 2. Environment variable, 3. Default
        if config_path:
            config_file_path = config_path
        elif os.getenv('EXCEL_EXPLORER_CONFIG'):
            config_file_path = os.getenv('EXCEL_EXPLORER_CONFIG')
        else:
            # Default to config/config.yaml relative to project root
            current_dir = Path(__file__).parent.parent.parent
            config_file_path = str(current_dir / "config" / "config.yaml")
        
        # Load base configuration
        self._config = self._load_base_config(config_file_path)
        
        # Apply environment variable overrides
        self._apply_env_overrides()
        
        # Validate configuration
        self._validate_config()
        
        return self._config
    
    def get(self, key_path: str, default: Any = None) -> Any:
        """
        Get configuration value using dot notation
        
        Args:
            key_path: Dot-separated path (e.g., 'analysis.sample_rows')
            default: Default value if key not found
            
        Returns:
            Configuration value or default
        """
        if self._config is None:
            self.load_config()
        
        keys = key_path.split('.')
        value = self._config
        
        try:
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            return default
    
    def reload_config(self, config_path: Optional[str] = None):
        """Force reload configuration"""
        self._config = None
        return self.load_config(config_path)
    
    def _load_base_config(self, config_path: str) -> Dict[str, Any]:
        """Load configuration from YAML file with fallback to defaults"""
        try:
            config_file = Path(config_path)
            if config_file.exists():
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = yaml.safe_load(f) or {}
                print(f"Configuration loaded from: {config_file}")
                return self._merge_with_defaults(config)
            else:
                print(f"Warning: Config file not found: {config_path}, using defaults")
                return self._get_default_config()
                
        except yaml.YAMLError as e:
            print(f"Warning: YAML parsing error in {config_path}: {e}")
            print("Using default configuration")
            return self._get_default_config()
        except Exception as e:
            print(f"Warning: Error loading config from {config_path}: {e}")
            print("Using default configuration")
            return self._get_default_config()
    
    def _get_default_config(self) -> Dict[str, Any]:
        """Return default configuration structure"""
        return {
            'analysis': {
                'max_cells_check': 1000,
                'max_formula_check': 1000,
                'sample_rows': 100,
                'max_sample_rows': 1000,
                'memory_limit_mb': 512,
                'timeout_per_sheet_seconds': 30,
                'enable_cross_sheet_analysis': True,
                'enable_data_quality_checks': True,
                'detail_level': 'comprehensive'
            },
            'output': {
                'json_enabled': True,
                'html_enabled': True,
                'include_raw_data': False,
                'auto_export': True,
                'timestamp_reports': True
            },
            'performance': {
                'parallel_processing': False,
                'chunk_size': 1000,
                'timeout_seconds': 300,
                'memory_warning_threshold_mb': 1024
            },
            'logging': {
                'level': 'INFO',
                'include_timestamps': True,
                'log_to_file': False,
                'log_file_path': 'excel_explorer.log'
            },
            'security': {
                'enable_pattern_detection': True,
                'scan_for_pii': True,
                'security_threshold': 8.0
            }
        }
    
    def _merge_with_defaults(self, user_config: Dict[str, Any]) -> Dict[str, Any]:
        """Merge user configuration with defaults"""
        default_config = self._get_default_config()
        return self._deep_merge(default_config, user_config)
    
    def _deep_merge(self, base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
        """Deep merge two dictionaries"""
        result = base.copy()
        
        for key, value in override.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = self._deep_merge(result[key], value)
            else:
                result[key] = value
        
        return result
    
    def _apply_env_overrides(self):
        """Apply environment variable overrides"""
        env_mappings = {
            # Analysis settings
            'EXCEL_EXPLORER_SAMPLE_ROWS': ['analysis', 'sample_rows'],
            'EXCEL_EXPLORER_MAX_FORMULA_CHECK': ['analysis', 'max_formula_check'],
            'EXCEL_EXPLORER_MEMORY_LIMIT_MB': ['analysis', 'memory_limit_mb'],
            'EXCEL_EXPLORER_TIMEOUT_SECONDS': ['analysis', 'timeout_per_sheet_seconds'],
            'EXCEL_EXPLORER_DETAIL_LEVEL': ['analysis', 'detail_level'],
            
            # Performance settings
            'EXCEL_EXPLORER_CHUNK_SIZE': ['performance', 'chunk_size'],
            'EXCEL_EXPLORER_PARALLEL_PROCESSING': ['performance', 'parallel_processing'],
            
            # Output settings
            'EXCEL_EXPLORER_AUTO_EXPORT': ['output', 'auto_export'],
            'EXCEL_EXPLORER_INCLUDE_RAW_DATA': ['output', 'include_raw_data'],
            
            # Logging settings
            'EXCEL_EXPLORER_LOG_LEVEL': ['logging', 'level'],
            'EXCEL_EXPLORER_LOG_TO_FILE': ['logging', 'log_to_file'],
        }
        
        for env_var, config_path in env_mappings.items():
            env_value = os.getenv(env_var)
            if env_value is not None:
                # Convert string to appropriate type
                converted_value = self._convert_env_value(env_value, config_path)
                self._set_nested_value(self._config, config_path, converted_value)
    
    def _convert_env_value(self, value: str, config_path: list) -> Any:
        """Convert environment variable string to appropriate type"""
        # Boolean conversion
        if value.lower() in ('true', 'false'):
            return value.lower() == 'true'
        
        # Integer conversion
        try:
            return int(value)
        except ValueError:
            pass
        
        # Float conversion
        try:
            return float(value)
        except ValueError:
            pass
        
        # Return as string
        return value
    
    def _set_nested_value(self, config: Dict[str, Any], path: list, value: Any):
        """Set a nested dictionary value using a path list"""
        current = config
        for key in path[:-1]:
            if key not in current:
                current[key] = {}
            current = current[key]
        current[path[-1]] = value
    
    def _validate_config(self):
        """Validate configuration values and apply constraints"""
        # Ensure numeric values are within reasonable bounds
        constraints = {
            ('analysis', 'sample_rows'): (1, 10000),
            ('analysis', 'max_formula_check'): (1, 100000),
            ('analysis', 'memory_limit_mb'): (64, 8192),
            ('analysis', 'timeout_per_sheet_seconds'): (1, 3600),
            ('performance', 'chunk_size'): (1, 100000),
            ('performance', 'timeout_seconds'): (1, 7200),
        }
        
        for path, (min_val, max_val) in constraints.items():
            value = self.get('.'.join(path))
            if value is not None and isinstance(value, (int, float)):
                if value < min_val:
                    print(f"Warning: Config value {'.'.join(path)}={value} too low, using minimum: {min_val}")
                    self._set_nested_value(self._config, list(path), min_val)
                elif value > max_val:
                    print(f"Warning: Config value {'.'.join(path)}={value} too high, using maximum: {max_val}")
                    self._set_nested_value(self._config, list(path), max_val)
        
        # Validate string choices
        valid_choices = {
            ('analysis', 'detail_level'): ['basic', 'standard', 'comprehensive'],
            ('logging', 'level'): ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
        }
        
        for path, choices in valid_choices.items():
            value = self.get('.'.join(path))
            if value is not None and value not in choices:
                print(f"Warning: Invalid config value {'.'.join(path)}={value}, using default")
                # Reset to default
                default_value = self._get_default_config()
                for key in path:
                    default_value = default_value[key]
                self._set_nested_value(self._config, list(path), default_value)
    
    def export_current_config(self, output_path: str):
        """Export current configuration to YAML file"""
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                yaml.dump(self._config, f, default_flow_style=False, indent=2)
            print(f"Configuration exported to: {output_path}")
        except Exception as e:
            print(f"Error: Failed to export configuration: {e}")
    
    def get_analysis_config(self) -> Dict[str, Any]:
        """Get analysis-specific configuration"""
        return self.get('analysis', {})
    
    def get_performance_config(self) -> Dict[str, Any]:
        """Get performance-specific configuration"""
        return self.get('performance', {})
    
    def get_output_config(self) -> Dict[str, Any]:
        """Get output-specific configuration"""
        return self.get('output', {})


# Global instance for easy access
config_manager = ConfigManager()


def get_config(config_path: Optional[str] = None) -> Dict[str, Any]:
    """Convenience function to get configuration"""
    return config_manager.load_config(config_path)


if __name__ == "__main__":
    # Test configuration loading
    config = ConfigManager()
    print("Loading configuration...")
    cfg = config.load_config()
    
    print("\nConfiguration loaded:")
    print(f"Sample rows: {config.get('analysis.sample_rows')}")
    print(f"Memory limit: {config.get('analysis.memory_limit_mb')} MB")
    print(f"Detail level: {config.get('analysis.detail_level')}")
    
    # Test environment variable override
    os.environ['EXCEL_EXPLORER_SAMPLE_ROWS'] = '200'
    config.reload_config()
    print(f"\nAfter env override - Sample rows: {config.get('analysis.sample_rows')}")
