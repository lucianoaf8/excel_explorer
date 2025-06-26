#!/usr/bin/env python3
"""
Script to fix unreachable code in data_profiler.py
Removes duplicate method definitions that are incorrectly placed after return statements
"""

import re
import os

def fix_unreachable_code():
    """Fix the unreachable code issue in data_profiler.py"""
    
    file_path = r"c:\Projects\excel_explorer\src\modules\data_profiler.py"
    backup_path = file_path + ".backup"
    
    # Create backup
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    with open(backup_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"Created backup: {backup_path}")
    
    # Split content into lines
    lines = content.split('\n')
    
    # Find the create_data_profiler function and remove unreachable code after it
    fixed_lines = []
    inside_create_function = False
    function_ended = False
    
    for i, line in enumerate(lines):
        # Detect start of create_data_profiler function
        if line.strip().startswith('def create_data_profiler('):
            inside_create_function = True
            function_ended = False
            fixed_lines.append(line)
            continue
        
        # If we're inside the function and hit a return statement
        if inside_create_function and not function_ended:
            fixed_lines.append(line)
            if line.strip().startswith('return '):
                function_ended = True
                continue
        
        # Skip unreachable code after return in create_data_profiler
        if inside_create_function and function_ended:
            # Check if we've reached the next top-level definition or end of file
            if (line.strip().startswith('def ') and not line.startswith('    ') and 
                not line.startswith('        ') and not line.strip().startswith('def _')):
                # This is a new top-level function, stop skipping
                inside_create_function = False
                function_ended = False
                fixed_lines.append(line)
            elif line.strip().startswith('class ') and not line.startswith('    '):
                # This is a new class, stop skipping
                inside_create_function = False
                function_ended = False
                fixed_lines.append(line)
            elif i == len(lines) - 1:
                # End of file
                break
            # Otherwise, skip this line (it's unreachable code)
            continue
        
        # Normal line processing
        if not (inside_create_function and function_ended):
            fixed_lines.append(line)
    
    # Write the fixed content
    fixed_content = '\n'.join(fixed_lines)
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(fixed_content)
    
    print(f"Fixed unreachable code in {file_path}")
    print(f"Original file size: {len(content)} characters")
    print(f"Fixed file size: {len(fixed_content)} characters")
    print(f"Removed {len(content) - len(fixed_content)} characters of unreachable code")

if __name__ == "__main__":
    fix_unreachable_code()
