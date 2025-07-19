#!/usr/bin/env python3
"""
Setup script for Excel Explorer
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the README file
readme_path = Path(__file__).parent / "docs" / "README.md"
long_description = readme_path.read_text(encoding='utf-8') if readme_path.exists() else ""

# Read requirements
requirements_path = Path(__file__).parent / "requirements.txt"
requirements = []
if requirements_path.exists():
    requirements = requirements_path.read_text(encoding='utf-8').strip().split('\n')
    requirements = [req.strip() for req in requirements if req.strip() and not req.startswith('#')]

setup(
    name="excel-explorer",
    version="2.0.0",
    author="Excel Explorer Team",
    description="Advanced Excel File Analysis Tool with GUI and CLI interfaces",
    long_description=long_description,
    long_description_content_type="text/markdown",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    python_requires=">=3.9",
    install_requires=requirements,
    entry_points={
        "console_scripts": [
            "excel-explorer=excel_explorer.main:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "Intended Audience :: Developers",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Operating System :: OS Independent",
    ],
    keywords="excel analysis spreadsheet data-profiling security-analysis",
    project_urls={
        "Source": "https://github.com/excel-explorer/excel-explorer",
        "Documentation": "https://github.com/excel-explorer/excel-explorer/docs",
        "Bug Reports": "https://github.com/excel-explorer/excel-explorer/issues",
    },
    include_package_data=True,
    package_data={
        "excel_explorer": ["config/*.yaml"],
    },
)
