# Simple test script to run the Excel merger with sample files
import sys
import os

# Add current directory to sys.path to import excel_merger
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from excel_merger import ExcelMerger

def test_merger():
    print("Testing Excel Merger with sample files...")
    
    # Create merger instance
    merger = ExcelMerger()
    
    # Load sample files
    merger.load_excel_files('sample_file1.xlsx', 'sample_file2.xlsx')
    
    # Get matching columns
    matching_cols = merger.find_matching_columns()
    print(f"Matching columns: {matching_cols}")
    
    # Perform merge
    output_path = 'merged_output.xlsx'
    merger.merge_files(output_path, merge_type='inner', matching_column='ID')
    
    print("Test completed successfully!")

if __name__ == "__main__":
    test_merger()