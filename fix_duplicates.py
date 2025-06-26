#!/usr/bin/env python3
"""
Script to fix duplicate functions and clean up data_profiler.py
"""

import re

def fix_duplicates():
    """Fix duplicate functions and clean up the file"""
    
    file_path = r"c:\Projects\excel_explorer\src\modules\data_profiler.py"
    
    # Read the current content
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Split into lines
    lines = content.split('\n')
    
    # Clean up the file by removing duplicates and fixing the create_data_profiler function
    fixed_lines = []
    skip_until_next_function = False
    found_create_function = False
    
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # Handle create_data_profiler function - keep only the first clean version
        if line.strip().startswith('def create_data_profiler(') and not found_create_function:
            found_create_function = True
            # Add the clean version of the function
            fixed_lines.extend([
                '',
                '# Legacy compatibility',
                'def create_data_profiler(config: dict = None) -> DataProfiler:',
                '    """Factory function for backward compatibility"""',
                '    profiler = DataProfiler()',
                '    if config:',
                '        profiler.configure(config)',
                '    return profiler'
            ])
            
            # Skip all lines until we find the next top-level definition or end of file
            i += 1
            while i < len(lines):
                current_line = lines[i]
                # Check if we've reached a new top-level definition
                if (current_line.strip().startswith('def ') and not current_line.startswith('    ') and 
                    not current_line.startswith('        ')) or \
                   (current_line.strip().startswith('class ') and not current_line.startswith('    ')):
                    break
                i += 1
            continue
        
        # Skip any additional create_data_profiler functions
        elif line.strip().startswith('def create_data_profiler(') and found_create_function:
            # Skip this duplicate function
            i += 1
            while i < len(lines):
                current_line = lines[i]
                if (current_line.strip().startswith('def ') and not current_line.startswith('    ') and 
                    not current_line.startswith('        ')) or \
                   (current_line.strip().startswith('class ') and not current_line.startswith('    ')):
                    break
                i += 1
            continue
        
        # Add normal lines
        fixed_lines.append(line)
        i += 1
    
    # Write the fixed content
    fixed_content = '\n'.join(fixed_lines)
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(fixed_content)
    
    print(f"Fixed duplicates in {file_path}")
    print(f"Original lines: {len(lines)}")
    print(f"Fixed lines: {len(fixed_lines)}")

if __name__ == "__main__":
    fix_duplicates()
