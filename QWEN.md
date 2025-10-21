# Excel Handler Project

## Project Overview

This is a Python-based Excel file handling and merging tool that allows users to merge two Excel files based on matching column names. The project uses pandas for data manipulation and provides both an interactive command-line interface and a class-based API for Excel file operations.

## Key Features

- Merges two Excel files based on matching column names
- Supports different merge types (inner, outer, left, right)
- Interactive CLI for user input
- Test functionality with sample data
- Automated sample file creation for testing

## Core Components

### Main Classes

- **ExcelMerger**: The main class that handles loading, merging, and saving of Excel files
  - `load_excel_files()`: Loads two Excel files into pandas dataframes
  - `find_matching_columns()`: Identifies matching column names between the two files
  - `merge_files()`: Performs the merge operation with various merge types

### Supporting Scripts

- **create_sample1.py**: Creates a sample Excel file with employee data (ID, Name, Age, Department)
- **create_sample2.py**: Creates a sample Excel file with employee data (ID, Salary, Location, Experience)
- **test_merger.py**: Runs a test to verify the merger functionality with sample files
- **excel_merger.py**: Contains the main ExcelMerger class and interactive CLI
- **merged_output.xlsx**: Sample output file from a merge operation
- **sample_file1.xlsx/sample_file2.xlsx**: Sample Excel files for testing

## Dependencies

The project uses the following Python libraries:
- pandas: For data manipulation and Excel file handling
- openpyxl (likely as a dependency of pandas for Excel support)

## Building and Running

### Prerequisites
- Python 3.x
- Install pandas: `pip install pandas openpyxl`

### Running the Application

1. **Interactive Mode**:
   ```bash
   python excel_merger.py
   ```
   This will prompt you for two Excel file paths, output path, merge type, and column to merge on.

2. **Testing with Sample Data**:
   ```bash
   # Create sample files
   python create_sample1.py
   python create_sample2.py
   
   # Run test
   python test_merger.py
   ```

3. **Using the Merger in Code**:
   ```python
   from excel_merger import ExcelMerger
   
   merger = ExcelMerger()
   merger.load_excel_files('file1.xlsx', 'file2.xlsx')
   merger.merge_files('output.xlsx', merge_type='inner', matching_column='ID')
   ```

## Development Conventions

- Type hints are used throughout the codebase
- Comprehensive docstrings following Google style
- Error handling with appropriate exception propagation
- Input validation to ensure file existence
- Clear console output for user feedback

## Project Structure

- `excel_merger.py`: Main application logic and class
- `test_merger.py`: Test script for functionality verification
- `create_sample1.py` and `create_sample2.py`: Scripts to generate test data
- `README.md`: Minimal project documentation
- Sample output files: `merged_output.xlsx`, `sample_file1.xlsx`, `sample_file2.xlsx`

## Use Cases

This project is useful for:
- Merging datasets from different Excel files based on common identifiers
- Combining reports or data from separate sources
- Data integration tasks requiring different types of joins
- Automated Excel processing workflows