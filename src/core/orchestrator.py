"""
Excel Explorer Main Orchestrator
Coordinates all analysis modules and manages the overall workflow.
"""

import argparse
import json
from pathlib import Path
from tqdm import tqdm
from ..utils import config_loader, error_handler
from ..modules.health_checker import HealthChecker
from ..modules.structure_mapper import StructureMapper


class ExcelExplorer:
    def __init__(self, config_path: str | None = None, log_file: str | None = None):
        self.config = config_loader.load_config(config_path)
        error_handler.configure_logging(log_file=log_file)

        # Instantiate modules
        self.health_checker = HealthChecker(self.config.get("health_checker", {}))
        self.structure_mapper = StructureMapper(self.config.get("structure_mapper", {}))

    def analyze_file(self, file_path: str | Path):
        """Run minimal workflow: Health Check -> print JSON report."""
        results = {}
        steps = [
            ("Health Check", self.health_checker.analyze, self.health_checker.get_results),
            ("Structure Mapping", self.structure_mapper.analyze, self.structure_mapper.get_results),
        ]

        for desc, run_fn, get_fn in tqdm(steps, desc="Analyzing", unit="module"):
            run_fn(file_path)
            results[desc.lower().replace(" ", "_")] = get_fn()

        health = results["health_check"]
        structure = results["structure_mapping"]

        print(json.dumps({"health_report": health, "structure": structure}, indent=2))


def _parse_args():
    parser = argparse.ArgumentParser(description="Excel Explorer Orchestrator")
    parser.add_argument("excel_file", help="Path to Excel file to analyze")
    parser.add_argument("--config", help="Path to YAML configuration file", default=None)
    parser.add_argument("--logfile", help="Path to write detailed log", default=None)
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    explorer = ExcelExplorer(config_path=args.config, log_file=args.logfile)
    explorer.analyze_file(args.excel_file)
