import pandas as pd
import os
from typing import Optional

class ExcelMerger:
    """
    A class to merge two Excel files based on matching column names.
    """
    
    def __init__(self):
        self.df1 = None
        self.df2 = None
        
    def load_excel_files(self, file1_path: str, file2_path: str):
        """
        Load two Excel files into pandas dataframes.
        
        Args:
            file1_path (str): Path to the first Excel file
            file2_path (str): Path to the second Excel file
        """
        try:
            self.df1 = pd.read_excel(file1_path)
            self.df2 = pd.read_excel(file2_path)
            print(f"Successfully loaded {file1_path}")
            print(f"Successfully loaded {file2_path}")
            print(f"File 1 shape: {self.df1.shape}")
            print(f"File 2 shape: {self.df2.shape}")
        except Exception as e:
            print(f"Error loading Excel files: {e}")
            raise
    
    def find_matching_columns(self) -> list:
        """
        Find matching column names between the two dataframes.
        
        Returns:
            list: List of matching column names
        """
        if self.df1 is None or self.df2 is None:
            raise ValueError("Excel files not loaded. Please load files first.")
            
        matching_cols = list(set(self.df1.columns) & set(self.df2.columns))
        return matching_cols
    
    def merge_files(self, output_path: str, merge_type: str = 'inner', 
                   matching_column: Optional[str] = None) -> str:
        """
        Merge the two Excel files based on matching columns.
        
        Args:
            output_path (str): Path to save the merged Excel file
            merge_type (str): Type of merge - 'inner', 'outer', 'left', or 'right'
            matching_column (str, optional): Specific column to merge on. 
                                           If None, uses first matching column.
        
        Returns:
            str: Path to the merged Excel file
        """
        if self.df1 is None or self.df2 is None:
            raise ValueError("Excel files not loaded. Please load files first.")
            
        matching_columns = self.find_matching_columns()
        
        if not matching_columns:
            raise ValueError("No matching columns found between the two Excel files.")
        
        # If no specific column provided, use the first matching column
        if matching_column is None:
            matching_column = matching_columns[0]
            print(f"No specific column provided. Using first matching column: '{matching_column}'")
        
        if matching_column not in matching_columns:
            raise ValueError(f"Column '{matching_column}' not found in both Excel files.")
        
        print(f"Merging on column: '{matching_column}'")
        print(f"Merge type: {merge_type}")
        
        # Perform the merge
        merged_df = pd.merge(self.df1, self.df2, on=matching_column, how=merge_type)
        
        print(f"Merged dataframe shape: {merged_df.shape}")
        
        # Save to Excel file
        merged_df.to_excel(output_path, index=False)
        print(f"Merged Excel file saved to: {output_path}")
        
        return output_path


def main():
    """
    Main function to run the Excel merger application.
    """
    print("Excel Files Merger Application")
    print("="*40)
    
    # Get file paths from user
    file1_path = input("Enter the path of the first Excel file: ").strip()
    file2_path = input("Enter the path of the second Excel file: ").strip()
    
    # Validate file existence
    if not os.path.exists(file1_path):
        print(f"Error: File '{file1_path}' does not exist.")
        return
    
    if not os.path.exists(file2_path):
        print(f"Error: File '{file2_path}' does not exist.")
        return
    
    # Get output path
    output_path = input("Enter the path to save the merged Excel file: ").strip()
    
    # Create ExcelMerger instance
    merger = ExcelMerger()
    
    try:
        # Load the Excel files
        merger.load_excel_files(file1_path, file2_path)
        
        # Show matching columns
        matching_columns = merger.find_matching_columns()
        if matching_columns:
            print(f"\nMatching columns found: {matching_columns}")
        else:
            print("No matching columns found between the two Excel files.")
            return
        
        # Get merge type
        print("\nSelect merge type:")
        print("1. Inner (default) - Only rows with matching values in both files")
        print("2. Outer - All rows from both files")
        print("3. Left - All rows from first file")
        print("4. Right - All rows from second file")
        
        merge_choice = input("Enter your choice (1-4, default is 1): ").strip()
        merge_type_map = {'1': 'inner', '2': 'outer', '3': 'left', '4': 'right'}
        merge_type = merge_type_map.get(merge_choice, 'inner')
        
        # Get specific column to merge on (optional)
        specific_column = input(f"\nEnter specific column name to merge on (or press Enter to use first matching column '{matching_columns[0]}'): ").strip()
        if specific_column == "":
            specific_column = None
            
        # Perform the merge
        merger.merge_files(output_path, merge_type, specific_column)
        
        print("\nSuccessfully merged the Excel files!")
        
    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    main()