# Excel Merger

A comprehensive Python tool for merging Excel files with advanced features including multiple file support, custom merge conditions, formatting options, and configuration-based operations.

## Features

- **Basic Merging**: Merge two Excel files based on matching columns
- **Multiple File Merging**: Merge more than two Excel files at once
- **Column Mapping**: Handle files with non-matching column names
- **Custom Merge Conditions**: Define custom logic for merging beyond simple equality
- **Output Formatting**: Apply formatting to the merged Excel output
- **Configuration Files**: Use JSON/YAML files to specify merge operations
- **Command Line Interface**: Full argparse-based CLI for automation
- **Progress Tracking**: Monitor progress when processing large files
- **Logging System**: Comprehensive logging for debugging and monitoring
- **Plugin System**: Register custom merge strategies
- **Async Support**: Asynchronous processing for large files

## Installation

```bash
pip install pandas openpyxl pyyaml
```

## Usage

### Command Line Interface

#### Basic Two-File Merge
```bash
python excel_merger.py file1.xlsx file2.xlsx -o output.xlsx
```

#### Specify Merge Type
```bash
python excel_merger.py file1.xlsx file2.xlsx -o output.xlsx -m outer
```

#### Specify Column to Merge On
```bash
python excel_merger.py file1.xlsx file2.xlsx -o output.xlsx -c "Employee ID"
```

#### Multiple File Merge
```bash
python excel_merger.py -o output.xlsx --multiple file1.xlsx file2.xlsx file3.xlsx
```

#### With Formatting Options
```bash
python excel_merger.py file1.xlsx file2.xlsx -o output.xlsx --formatting '{"theme": "dark", "column_widths": [15, 20, 25]}'
```

#### Using Configuration File
```bash
python excel_merger.py --config config.json
```

### Using as a Library

#### Basic Usage
```python
from excel_merger import ExcelMerger

merger = ExcelMerger()

# Load two files
merger.load_excel_files('file1.xlsx', 'file2.xlsx')

# Perform merge
output_path = merger.merge_files('output.xlsx', merge_type='inner', matching_column='ID')
```

#### Multiple File Merge
```python
merger = ExcelMerger()

# Load multiple files
merger.load_multiple_excel_files(['file1.xlsx', 'file2.xlsx', 'file3.xlsx'])

# Perform merge
output_path = merger.merge_multiple_files('output.xlsx', merge_type='inner', matching_column='ID')
```

#### Custom Merge Strategy
```python
merger = ExcelMerger()

# Define custom strategy
def custom_strategy(df1, df2):
    # Your custom merge logic here
    return pd.concat([df1, df2])

# Register strategy
merger.register_merge_strategy("concat_strategy", custom_strategy)

# Execute strategy
result = merger.execute_custom_merge_strategy("concat_strategy", 'file1.xlsx', 'file2.xlsx', 'output.xlsx')
```

#### Column Mapping for Non-Matching Names
```python
from excel_merger import ExcelMerger

merger = ExcelMerger()

# When column names are similar but not identical
result = merger.map_columns_for_merge(
    'file1.xlsx', 
    'file2.xlsx', 
    column_mapping={'Name': 'Full_Name'},  # Map Name to Full_Name
    output_path='output.xlsx'
)
```

## Configuration Files

The tool supports configuration through JSON or YAML files:

### JSON Configuration Example
```json
{
  "file_paths": ["file1.xlsx", "file2.xlsx", "file3.xlsx"],
  "output_path": "merged_output.xlsx",
  "merge_type": "inner",
  "matching_column": "ID",
  "formatting_options": {
    "theme": "dark",
    "column_widths": [15, 20, 25]
  }
}
```

### YAML Configuration Example
```yaml
file1_path: file1.xlsx
file2_path: file2.xlsx
output_path: merged_output.xlsx
merge_type: outer
matching_column: ID
formatting_options:
  theme: light
  column_widths: [15, 20, 25]
```

## Available Merge Types

- `inner`: Only rows with matching values in both files (default)
- `outer`: All rows from both files
- `left`: All rows from first file
- `right`: All rows from second file

## Error Handling

The tool provides comprehensive error handling:

- File validation (existence, permissions, format)
- Data validation (empty files, column mismatches)
- Type validation for all parameters
- Detailed error messages with context

## Logging

The application supports configurable logging:

- Multiple log levels (DEBUG, INFO, WARNING, ERROR, CRITICAL)
- Both console and file output
- Rotating file handler to prevent large log files

## Development

### Running Tests

Run unit tests:
```bash
python test_merger.py comprehensive
```

Run integration tests:
```bash
python test_merger.py integration
```

Run performance tests:
```bash
python test_merger.py performance
```

### Code Coverage

Run the coverage analysis:
```bash
python run_coverage.py
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.