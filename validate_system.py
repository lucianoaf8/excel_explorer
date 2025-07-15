#!/usr/bin/env python3
"""System validation script"""

def validate_imports():
    """Test all imports work"""
    try:
        from src.core.orchestrator import ExcelExplorer
        from src.core.unified_config import UnifiedConfig
        from src.main import main
        print("✅ All imports successful")
        return True
    except Exception as e:
        print(f"❌ Import failed: {e}")
        return False

def validate_config():
    """Test configuration system"""
    try:
        from src.core.unified_config import UnifiedConfig
        config = UnifiedConfig()
        assert config.max_memory_mb > 0
        assert config.chunk_size_rows > 0
        print("✅ Configuration validation passed")
        return True
    except Exception as e:
        print(f"❌ Configuration validation failed: {e}")
        return False

def validate_modules():
    """Test module instantiation"""
    try:
        from src.modules.health_checker import HealthChecker
        from src.modules.structure_mapper import StructureMapper
        from src.modules.data_profiler import DataProfiler
      
        modules = [
            HealthChecker(),
            StructureMapper(),
            DataProfiler()
        ]
        print("✅ Module instantiation successful")
        return True
    except Exception as e:
        print(f"❌ Module validation failed: {e}")
        return False

if __name__ == "__main__":
    validations = [
        validate_imports,
        validate_config,
        validate_modules
    ]
  
    success = all(validation() for validation in validations)
  
    if success:
        print("\n🎉 System validation PASSED - Ready for testing with Excel files")
    else:
        print("\n💥 System validation FAILED - Fix errors before proceeding")
  
    exit(0 if success else 1)