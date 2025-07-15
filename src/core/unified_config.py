from dataclasses import dataclass
from typing import Dict, Any, Optional
from pathlib import Path
import yaml

@dataclass
class UnifiedConfig:
    # Analysis settings
    max_memory_mb: int = 4096
    warning_threshold: float = 0.8
    chunk_size_rows: int = 10000
    max_formulas_analyze: int = 50000
    enable_caching: bool = True
  
    # Module toggles
    include_charts: bool = True
    include_pivots: bool = True
    include_connections: bool = True
    deep_analysis: bool = False
    parallel_processing: bool = False
  
    # Module-specific settings
    module_configs: Dict[str, Dict[str, Any]] = None
  
    def __post_init__(self):
        if self.module_configs is None:
            self.module_configs = {}
  
    @classmethod
    def from_yaml(cls, path: Optional[str] = None) -> 'UnifiedConfig':
        config = cls()
        if path and Path(path).exists():
            try:
                with open(path, 'r') as f:
                    data = yaml.safe_load(f) or {}
                  
                # Map YAML structure to config
                if 'analysis' in data:
                    analysis = data['analysis']
                    config.max_memory_mb = analysis.get('max_memory_mb', config.max_memory_mb)
                    config.chunk_size_rows = analysis.get('chunk_size_rows', config.chunk_size_rows)
                    config.enable_caching = analysis.get('enable_caching', config.enable_caching)
              
                # Store module-specific configs
                for key, value in data.items():
                    if key not in ['analysis']:
                        config.module_configs[key] = value
                      
            except Exception as e:
                import logging
                logging.warning(f"Failed to load config from {path}: {e}")
      
        return config
  
    def get_module_config(self, module_name: str) -> Dict[str, Any]:
        return self.module_configs.get(module_name, {})