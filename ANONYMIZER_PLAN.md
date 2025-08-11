# Excel Anonymizer Implementation Plan

## Overview
Create a simple utility to anonymize Excel data by replacing sensitive values with fake ones while maintaining a mapping for reversal.

## Files to Create

### 1. Main Anonymizer Script
**File:** `src/utils/anonymizer.py`

**Purpose:** Single file containing all anonymizer functionality

**Key Features:**
- Read Excel files
- Detect columns with names/companies (simple pattern matching)
- Replace values with fake data using Faker library
- Keep mapping dictionary
- Save anonymized file and mapping

### 2. CLI Integration
**File:** `src/cli/anonymizer_command.py`

**Purpose:** Command-line interface for the anonymizer

### 3. Test File
**File:** `tests/test_anonymizer.py`

**Purpose:** Basic tests for anonymizer functionality

## Implementation Steps

### Step 1: Create Core Anonymizer
Create `src/utils/anonymizer.py` with:

1. **ExcelAnonymizer class** with methods:
   - `__init__(file_path)` - Load Excel file
   - `find_name_columns()` - Find columns containing names/companies
   - `anonymize_data(columns_to_anonymize)` - Replace values with fake ones
   - `save_files(output_path, mapping_path)` - Save anonymized Excel + mapping JSON

2. **Helper functions:**
   - Generate fake names using Faker
   - Create consistent mappings (same original → same fake)
   - Pattern matching for column detection

### Step 2: Add CLI Support
Update `src/cli/cli_runner.py` to add anonymizer options:
- `--anonymize` - Enable anonymization
- `--anonymize-columns` - Specify which columns
- `--mapping-file` - Output path for mapping JSON

### Step 3: Create Simple Test
Create `tests/test_anonymizer.py` with basic tests:
- Load sample Excel file
- Run anonymization
- Verify mappings are consistent
- Test reversal process

## Usage Example

```bash
# Anonymize with auto-detection
python main.py --mode cli --file data.xlsx --anonymize --mapping-file mappings.json

# Anonymize specific columns
python main.py --mode cli --file data.xlsx --anonymize --anonymize-columns "Name" "Company" --mapping-file mappings.json
```

## Dependencies
Add to `requirements.txt`:
```
faker
```

## Mapping File Format
Simple JSON structure:
```json
{
  "column_name": {
    "Original Name": "Fake Name",
    "Another Original": "Another Fake"
  }
}
```

## Column Detection Logic
Simple pattern matching for column headers containing:
- "name" (person names)
- "contractor", "client", "customer" (person names)
- "company", "organization", "vendor" (company names)

## Key Implementation Details

1. **Consistency:** Use dictionary to ensure same original value always gets same fake value
2. **Data Types:** Preserve Excel cell formatting and data types
3. **Empty Cells:** Skip empty/null cells
4. **Case Handling:** Handle different text cases appropriately
5. **File Output:** Create new anonymized file, don't modify original

## Testing Approach
1. Create test Excel with sample data
2. Run anonymizer and verify output
3. Check mapping file completeness
4. Test edge cases (empty cells, duplicate values)
5. Verify reversal works correctly

This simplified approach focuses on core functionality without over-engineering the solution.