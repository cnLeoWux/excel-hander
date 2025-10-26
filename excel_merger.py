import pandas as pd
import os
from typing import Optional, List, Dict, Any, Union
import logging
from logging.handlers import RotatingFileHandler
import sys
import json
import yaml

class ExcelMerger:
    """
    A class to merge Excel files based on matching column names.
    
    This class provides functionality to load Excel files, find matching columns,
    and merge them using different types of joins (inner, outer, left, right).
    
    Attributes:
        df1 (pd.DataFrame): First loaded DataFrame
        df2 (pd.DataFrame): Second loaded DataFrame
        logger (logging.Logger): Logger instance for the merger
    """
    
    def __init__(self, log_level: str = "INFO", log_file: Optional[str] = "excel_merger.log"):
        """
        Initialize the ExcelMerger instance.
        
        Args:
            log_level (str): Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL). 
                           Defaults to "INFO".
            log_file (Optional[str]): Path to log file. If None, no file logging occurs.
                                    Defaults to "excel_merger.log".
        """
        self.df1: Optional[pd.DataFrame] = None
        self.df2: Optional[pd.DataFrame] = None
        self.logger: logging.Logger = self._setup_logger(log_level, log_file)
        self.custom_merge_strategies = {}  # Dictionary to store custom merge strategies

    def register_merge_strategy(self, name: str, strategy_func: callable) -> None:
        """
        Register a custom merge strategy.
        
        Args:
            name (str): Name of the strategy
            strategy_func (callable): Function that takes two DataFrames and returns a merged DataFrame
        """
        self.custom_merge_strategies[name] = strategy_func
        self.logger.info(f"Registered custom merge strategy: {name}")

    def execute_custom_merge_strategy(self, strategy_name: str, file1_path: str, file2_path: str, 
                                     output_path: str, **kwargs) -> str:
        """
        Execute a registered custom merge strategy.
        
        Args:
            strategy_name (str): Name of the registered strategy to execute
            file1_path (str): Path to the first Excel file
            file2_path (str): Path to the second Excel file
            output_path (str): Path to save the merged Excel file
            **kwargs: Additional arguments to pass to the strategy function
            
        Returns:
            str: Path to the merged Excel file
        """
        if strategy_name not in self.custom_merge_strategies:
            raise ValueError(f"Unknown merge strategy: {strategy_name}. "
                           f"Available strategies: {list(self.custom_merge_strategies.keys())}")
        
        self.logger.info(f"Executing custom merge strategy: {strategy_name}")
        
        # Load the files
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)
        
        if df1.empty or df2.empty:
            raise ValueError("One or both Excel files are empty.")
        
        # Execute the custom strategy
        merged_df = self.custom_merge_strategies[strategy_name](df1, df2, **kwargs)
        
        # Validate result
        if not isinstance(merged_df, pd.DataFrame):
            raise ValueError(f"Custom merge strategy {strategy_name} must return a pandas DataFrame")
        
        self.logger.info(f"Merged dataframe shape: {merged_df.shape}")
        
        # Save to Excel file
        merged_df.to_excel(output_path, index=False)
        self.logger.info(f"Merged Excel file saved to: {output_path}")
        
        return output_path
        
    def _setup_logger(self, log_level: str, log_file: Optional[str]) -> logging.Logger:
        """
        Set up a logger for the ExcelMerger class.
        
        Args:
            log_level (str): The logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
            log_file (Optional[str]): File to write logs to. If None, no file logging occurs.
            
        Returns:
            logging.Logger: Configured logger instance
        """
        # Create logger
        logger = logging.getLogger('ExcelMerger')
        logger.setLevel(getattr(logging, log_level.upper()))
        
        # Prevent adding multiple handlers if logger already has handlers
        if logger.handlers:
            logger.handlers.clear()
        
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Create console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(getattr(logging, log_level.upper()))
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
        
        # Create file handler if log_file is provided
        if log_file:
            file_handler = RotatingFileHandler(
                log_file, 
                maxBytes=10*1024*1024,  # 10MB
                backupCount=5
            )
            file_handler.setLevel(getattr(logging, log_level.upper()))
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        
        return logger
        
    def validate_file_path(self, file_path: str) -> None:
        """
        Validate that the file path exists and is accessible.
        
        Args:
            file_path (str): Path to the Excel file
            
        Raises:
            FileNotFoundError: If the file does not exist
            ValueError: If the path is not a file or not an Excel file
            PermissionError: If the file is not accessible
        """
        if not file_path:
            raise ValueError("File path cannot be empty")
        
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File does not exist: {file_path}")
        
        if not os.path.isfile(file_path):
            raise ValueError(f"Path is not a file: {file_path}")
        
        # Check if it's an Excel file by extension
        if not file_path.lower().endswith(('.xlsx', '.xls', '.xlsm')):
            raise ValueError(f"File is not an Excel file (.xlsx, .xls, .xlsm): {file_path}")
        
        self.logger.debug(f"Validated file path: {file_path}")
    
    def load_excel_files(self, file1_path: str, file2_path: str) -> None:
        """
        Load two Excel files into pandas dataframes.
        
        Args:
            file1_path (str): Path to the first Excel file
            file2_path (str): Path to the second Excel file
            
        Raises:
            FileNotFoundError: If a file does not exist
            PermissionError: If a file is not accessible
            ValueError: If a file is empty or not in Excel format
            Exception: For other errors during loading
        """
        self.logger.info(f"Loading Excel files: {file1_path} and {file2_path}")
        
        # Validate file paths
        self.validate_file_path(file1_path)
        self.validate_file_path(file2_path)
        
        try:
            # Check if files are readable
            if not os.access(file1_path, os.R_OK):
                raise PermissionError(f"File is not readable: {file1_path}")
            if not os.access(file2_path, os.R_OK):
                raise PermissionError(f"File is not readable: {file2_path}")
            
            self.df1 = pd.read_excel(file1_path)
            self.df2 = pd.read_excel(file2_path)
            
            # Validate that dataframes are not empty
            if self.df1.empty:
                raise ValueError(f"First Excel file is empty: {file1_path}")
            if self.df2.empty:
                raise ValueError(f"Second Excel file is empty: {file2_path}")
            
            self.logger.info(f"Successfully loaded {file1_path}")
            self.logger.info(f"Successfully loaded {file2_path}")
            self.logger.info(f"File 1 shape: {self.df1.shape}")
            self.logger.info(f"File 2 shape: {self.df2.shape}")
        except FileNotFoundError as e:
            self.logger.error(f"File not found error: {e}")
            raise
        except PermissionError as e:
            self.logger.error(f"Permission error: {e}")
            raise
        except pd.errors.EmptyDataError:
            self.logger.error("Error: One of the Excel files is empty or contains no data")
            raise
        except Exception as e:
            self.logger.error(f"Error loading Excel files: {e}")
            raise
    
    def find_matching_columns(self) -> List[str]:
        """
        Find matching column names between the two dataframes.
        
        Returns:
            List[str]: List of matching column names
            
        Raises:
            ValueError: If Excel files have not been loaded
        """
        if self.df1 is None or self.df2 is None:
            raise ValueError("Excel files not loaded. Please load files first.")
            
        matching_cols = list(set(self.df1.columns) & set(self.df2.columns))
        self.logger.debug(f"Found matching columns: {matching_cols}")
        return matching_cols
    
    def validate_merge_parameters(self, output_path: str, merge_type: str, 
                                matching_column: Optional[str] = None) -> None:
        """
        Validate merge parameters before performing the merge.
        
        Args:
            output_path (str): Path to save the merged Excel file
            merge_type (str): Type of merge - 'inner', 'outer', 'left', or 'right'
            matching_column (Optional[str]): Specific column to merge on
            
        Raises:
            ValueError: If any parameter is invalid
            PermissionError: If the output directory is not writable
        """
        if not output_path:
            raise ValueError("Output path cannot be empty")
        
        # Validate merge type
        valid_merge_types = ['inner', 'outer', 'left', 'right']
        if merge_type not in valid_merge_types:
            raise ValueError(f"Invalid merge type: {merge_type}. Valid types are: {valid_merge_types}")
        
        # Check if output directory exists and is writable
        output_dir = os.path.dirname(os.path.abspath(output_path))
        if output_dir and not os.path.exists(output_dir):
            raise ValueError(f"Output directory does not exist: {output_dir}")
        
        if output_dir and not os.access(output_dir, os.W_OK):
            raise PermissionError(f"Output directory is not writable: {output_dir}")
            
        self.logger.debug(f"Validated merge parameters: output_path={output_path}, merge_type={merge_type}, matching_column={matching_column}")

    def load_multiple_excel_files(self, file_paths: List[str]) -> None:
        """
        Load multiple Excel files into a list of dataframes.
        
        Args:
            file_paths (List[str]): List of paths to Excel files to load
            
        Raises:
            FileNotFoundError: If any file does not exist
            PermissionError: If any file is not accessible
            ValueError: If any file is empty or not in Excel format
        """
        self.logger.info(f"Loading {len(file_paths)} Excel files")
        
        # Validate all file paths
        for path in file_paths:
            self.validate_file_path(path)
        
        try:
            self.dataframes = []
            for i, path in enumerate(file_paths):
                # Check if file is readable
                if not os.access(path, os.R_OK):
                    raise PermissionError(f"File is not readable: {path}")
                
                df = pd.read_excel(path)
                
                # Validate that dataframe is not empty
                if df.empty:
                    raise ValueError(f"Excel file is empty: {path}")
                
                self.dataframes.append(df)
                self.logger.info(f"Successfully loaded {path}, shape: {df.shape}")
                
            self.logger.info(f"Successfully loaded {len(self.dataframes)} Excel files")
            
        except FileNotFoundError as e:
            self.logger.error(f"File not found error: {e}")
            raise
        except PermissionError as e:
            self.logger.error(f"Permission error: {e}")
            raise
        except pd.errors.EmptyDataError:
            self.logger.error("Error: One of the Excel files is empty or contains no data")
            raise
        except Exception as e:
            self.logger.error(f"Error loading Excel files: {e}")
            raise
    
    def find_common_columns(self) -> List[str]:
        """
        Find common column names across all loaded dataframes.
        
        Returns:
            List[str]: List of common column names
            
        Raises:
            ValueError: If no dataframes have been loaded
        """
        if not hasattr(self, 'dataframes') or not self.dataframes:
            raise ValueError("No Excel files loaded. Please load files first.")
            
        # Start with columns from the first dataframe
        common_cols = set(self.dataframes[0].columns)
        
        # Find intersection with all other dataframes
        for df in self.dataframes[1:]:
            common_cols &= set(df.columns)
        
        common_cols_list = list(common_cols)
        self.logger.debug(f"Found common columns: {common_cols_list}")
        return common_cols_list
    
    def merge_multiple_files(self, output_path: str, merge_type: str = 'inner',
                           matching_column: Optional[str] = None,
                           sequential: bool = True) -> str:
        """
        Merge multiple Excel files based on matching columns.
        
        Args:
            output_path (str): Path to save the merged Excel file
            merge_type (str): Type of merge - 'inner', 'outer', 'left', or 'right'
            matching_column (Optional[str]): Specific column to merge on. 
                                           If None, uses first common column.
            sequential (bool): If True, merge files sequentially. If False, use reduce.
            
        Returns:
            str: Path to the merged Excel file
            
        Raises:
            ValueError: If parameters are invalid or no common columns found
            PermissionError: If output directory is not writable
        """
        self.logger.info(f"Starting multiple file merge operation with output path: {output_path}")
        
        # Validate parameters
        self.validate_merge_parameters(output_path, merge_type, matching_column)
        
        if not hasattr(self, 'dataframes') or len(self.dataframes) < 2:
            raise ValueError("At least two Excel files must be loaded for merging.")
        
        common_columns = self.find_common_columns()
        
        if not common_columns:
            raise ValueError("No common columns found between the Excel files.")
        
        # If no specific column provided, use the first common column
        if matching_column is None:
            matching_column = common_columns[0]
            self.logger.info(f"No specific column provided. Using first common column: '{matching_column}'")
        else:
            # Validate the specified column exists in all dataframes
            for i, df in enumerate(self.dataframes):
                if matching_column not in df.columns:
                    raise ValueError(f"Column '{matching_column}' not found in Excel file {i+1}.")
            
            if matching_column not in common_columns:
                raise ValueError(f"Column '{matching_column}' not found in all Excel files.")
        
        self.logger.info(f"Merging on column: '{matching_column}'")
        self.logger.info(f"Merge type: {merge_type}")
        self.logger.info(f"Merge strategy: {'sequential' if sequential else 'reduce'}")
        
        try:
            if sequential:
                # Perform sequential merge
                merged_df = self.dataframes[0]
                for df in self.dataframes[1:]:
                    merged_df = pd.merge(merged_df, df, on=matching_column, how=merge_type)
            else:
                # Use reduce for more efficient merging
                from functools import reduce
                merged_df = reduce(
                    lambda left, right: pd.merge(left, right, on=matching_column, how=merge_type),
                    self.dataframes
                )
            
            self.logger.info(f"Merged dataframe shape: {merged_df.shape}")
            
            # Save to Excel file
            merged_df.to_excel(output_path, index=False)
            self.logger.info(f"Merged Excel file saved to: {output_path}")
            
        except Exception as e:
            self.logger.error(f"Error during multiple file merge operation: {e}")
            raise
        
        return output_path

    def find_potential_matches(self, df1: pd.DataFrame, df2: pd.DataFrame, similarity_threshold: float = 0.8) -> Dict[str, str]:
        """
        Find potential column matches based on name similarity.
        
        Args:
            df1 (pd.DataFrame): First dataframe
            df2 (pd.DataFrame): Second dataframe
            similarity_threshold (float): Threshold for considering columns as similar (0-1)
            
        Returns:
            Dict[str, str]: Mapping of columns from df1 to df2 that are similar
        """
        import difflib
        
        potential_matches = {}
        
        for col1 in df1.columns:
            # Find the best match for each column in df1 within df2
            matches = difflib.get_close_matches(col1, df2.columns.tolist(), n=1, cutoff=similarity_threshold)
            if matches:
                potential_matches[col1] = matches[0]
        
        return potential_matches

    def map_columns_for_merge(self, file1_path: str, file2_path: str, 
                             column_mapping: Optional[Dict[str, str]] = None,
                             similarity_threshold: float = 0.8) -> pd.DataFrame:
        """
        Merge two Excel files using column mapping for non-matching column names.
        
        Args:
            file1_path (str): Path to the first Excel file
            file2_path (str): Path to the second Excel file
            column_mapping (Optional[Dict[str, str]]): Explicit mapping of columns from file1 to file2
            similarity_threshold (float): Threshold for auto-detecting similar column names
            
        Returns:
            pd.DataFrame: Merged dataframe
        """
        self.logger.info(f"Merging files with column mapping: {file1_path} and {file2_path}")
        
        # Load the files
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)
        
        if df1.empty or df2.empty:
            raise ValueError("One or both Excel files are empty.")
        
        # If no explicit mapping provided, try to auto-detect similar columns
        if column_mapping is None:
            column_mapping = self.find_potential_matches(df1, df2, similarity_threshold)
            self.logger.info(f"Auto-detected column mapping: {column_mapping}")
        
        # Check if we have any mappings
        if not column_mapping:
            raise ValueError("No column mappings provided or detected for merge operation.")
        
        # Validate that the mapped columns exist in both dataframes
        for col1, col2 in column_mapping.items():
            if col1 not in df1.columns:
                raise ValueError(f"Column '{col1}' not found in first Excel file.")
            if col2 not in df2.columns:
                raise ValueError(f"Column '{col2}' not found in second Excel file.")
        
        # Perform merge using the first mapped column
        first_mapping = next(iter(column_mapping.items()))
        merged_df = pd.merge(df1, df2, left_on=first_mapping[0], right_on=first_mapping[1], how='inner')
        
        self.logger.info(f"Merged dataframe shape: {merged_df.shape}")
        
        return merged_df

    def merge_with_column_mapping(self, file_paths: List[str], output_path: str,
                                merge_type: str = 'inner',
                                column_mappings: Optional[List[Dict[str, str]]] = None,
                                similarity_threshold: float = 0.8) -> str:
        """
        Merge multiple Excel files using column mappings for non-matching column names.
        
        Args:
            file_paths (List[str]): List of paths to Excel files to merge
            output_path (str): Path to save the merged Excel file
            merge_type (str): Type of merge - 'inner', 'outer', 'left', or 'right'
            column_mappings (Optional[List[Dict[str, str]]]): List of column mappings for each pair of files
            similarity_threshold (float): Threshold for auto-detecting similar column names
            
        Returns:
            str: Path to the merged Excel file
        """
        self.logger.info(f"Starting column mapping merge operation with output path: {output_path}")
        
        # Validate parameters
        self.validate_merge_parameters(output_path, merge_type)
        
        if len(file_paths) < 2:
            raise ValueError("At least two Excel files must be provided for merging.")
        
        # Load all files
        dataframes = []
        for path in file_paths:
            self.validate_file_path(path)
            if not os.access(path, os.R_OK):
                raise PermissionError(f"File is not readable: {path}")
            
            df = pd.read_excel(path)
            if df.empty:
                raise ValueError(f"Excel file is empty: {path}")
            
            dataframes.append(df)
            self.logger.info(f"Loaded {path}, shape: {df.shape}")
        
        # If no explicit mappings provided, try to auto-detect
        if column_mappings is None:
            column_mappings = []
            for i in range(len(dataframes) - 1):
                mapping = self.find_potential_matches(dataframes[i], dataframes[i + 1], similarity_threshold)
                column_mappings.append(mapping)
                self.logger.info(f"Auto-detected mapping for files {i} and {i+1}: {mapping}")
        
        # Validate mappings
        if len(column_mappings) != len(dataframes) - 1:
            raise ValueError(f"Number of column mappings ({len(column_mappings)}) doesn't match "
                           f"number of adjacent file pairs ({len(dataframes) - 1}).")
        
        # Perform sequential merge using the mappings
        merged_df = dataframes[0]
        
        for i, mapping in enumerate(column_mappings):
            if not mapping:
                raise ValueError(f"No column mapping provided or detected for files {i} and {i+1}.")
            
            # Use the first mapping as the join key
            first_key = next(iter(mapping.keys()))
            second_key = mapping[first_key]
            
            # Perform the merge
            merged_df = pd.merge(merged_df, dataframes[i+1], 
                               left_on=first_key, right_on=second_key, how=merge_type)
        
        self.logger.info(f"Final merged dataframe shape: {merged_df.shape}")
        
        # Save to Excel file
        merged_df.to_excel(output_path, index=False)
        self.logger.info(f"Merged Excel file saved to: {output_path}")
        
        return output_path

    def merge_with_custom_condition(self, file1_path: str, file2_path: str, output_path: str,
                                   merge_type: str = 'inner',
                                   custom_condition: Optional[callable] = None) -> str:
        """
        Merge two Excel files based on a custom condition instead of simple column equality.
        
        Args:
            file1_path (str): Path to the first Excel file
            file2_path (str): Path to the second Excel file
            output_path (str): Path to save the merged Excel file
            merge_type (str): Type of merge - 'inner', 'outer', 'left', or 'right'
            custom_condition (callable, optional): A function that takes two DataFrames 
                                                  and returns a merged DataFrame based on custom logic
            
        Returns:
            str: Path to the merged Excel file
        """
        self.logger.info(f"Starting custom condition merge operation with output path: {output_path}")
        
        # Validate parameters
        self.validate_merge_parameters(output_path, merge_type)
        self.validate_file_path(file1_path)
        self.validate_file_path(file2_path)
        
        # Load the files
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)
        
        if df1.empty or df2.empty:
            raise ValueError("One or both Excel files are empty.")
        
        if not os.access(file1_path, os.R_OK):
            raise PermissionError(f"First file is not readable: {file1_path}")
        if not os.access(file2_path, os.R_OK):
            raise PermissionError(f"Second file is not readable: {file2_path}")
        
        # If no custom condition provided, default to regular merge
        if custom_condition is None:
            # Find common columns and merge on the first one
            common_cols = list(set(df1.columns) & set(df2.columns))
            if not common_cols:
                raise ValueError("No common columns found between the two Excel files.")
            
            self.logger.info(f"No custom condition provided. Using default merge on column: {common_cols[0]}")
            merged_df = pd.merge(df1, df2, on=common_cols[0], how=merge_type)
        else:
            # Apply custom merge condition
            try:
                merged_df = custom_condition(df1, df2)
                if not isinstance(merged_df, pd.DataFrame):
                    raise ValueError("Custom condition function must return a pandas DataFrame")
            except Exception as e:
                self.logger.error(f"Error applying custom merge condition: {e}")
                raise
        
        self.logger.info(f"Merged dataframe shape: {merged_df.shape}")
        
        # Save to Excel file
        merged_df.to_excel(output_path, index=False)
        self.logger.info(f"Merged Excel file saved to: {output_path}")
        
        return output_path

    def range_merge_condition(self, df1: pd.DataFrame, df2: pd.DataFrame, 
                             left_col: str, right_col: str, 
                             operator: str = '>=') -> pd.DataFrame:
        """
        A custom merge condition that merges based on a range comparison.
        
        Args:
            df1 (pd.DataFrame): First dataframe
            df2 (pd.DataFrame): Second dataframe
            left_col (str): Column name in first dataframe for comparison
            right_col (str): Column name in second dataframe for comparison
            operator (str): Comparison operator ('>=', '<=', '>', '<', '==', '!=')
            
        Returns:
            pd.DataFrame: Merged dataframe based on the range condition
        """
        import operator as op_module
        
        # Map string operators to actual operator functions
        ops = {
            '>=': op_module.ge,
            '<=': op_module.le,
            '>': op_module.gt,
            '<': op_module.lt,
            '==': op_module.eq,
            '!=': op_module.ne
        }
        
        if operator not in ops:
            raise ValueError(f"Unsupported operator: {operator}. Supported: {list(ops.keys())}")
        
        op_func = ops[operator]
        
        # Create a cross join (Cartesian product) of both dataframes first
        df1['key'] = 1
        df2['key'] = 1
        cross_join = pd.merge(df1, df2, on='key', suffixes=('_left', '_right'))
        df1.drop('key', axis=1, inplace=True)
        df2.drop('key', axis=1, inplace=True)
        cross_join.drop('key', axis=1, inplace=True)
        
        # Apply the condition
        condition = op_func(cross_join[left_col], cross_join[right_col])
        filtered_df = cross_join[condition]
        
        # Remove the suffixes from column names if they exist
        new_columns = {}
        for col in filtered_df.columns:
            if col.endswith('_left'):
                original_col = col[:-5]  # Remove '_left'
                if original_col in df2.columns and original_col not in df1.columns:
                    new_columns[col] = col  # Keep as is if no conflict
                else:
                    new_columns[col] = original_col
            elif col.endswith('_right'):
                original_col = col[:-6]  # Remove '_right'
                if original_col in df1.columns and original_col not in df2.columns:
                    new_columns[col] = col  # Keep as is if no conflict
                else:
                    new_columns[col] = f"{original_col}_right"
        
        filtered_df.rename(columns=new_columns, inplace=True)
        
        return filtered_df
    
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
        self.logger.info(f"Starting merge operation with output path: {output_path}")
        
        # Validate parameters
        self.validate_merge_parameters(output_path, merge_type, matching_column)
        
        if self.df1 is None or self.df2 is None:
            raise ValueError("Excel files not loaded. Please load files first.")
            
        matching_columns = self.find_matching_columns()
        
        if not matching_columns:
            raise ValueError("No matching columns found between the two Excel files.")
        
        # If no specific column provided, use the first matching column
        if matching_column is None:
            matching_column = matching_columns[0]
            self.logger.info(f"No specific column provided. Using first matching column: '{matching_column}'")
        else:
            # Validate the specified column exists in both dataframes
            if matching_column not in self.df1.columns:
                raise ValueError(f"Column '{matching_column}' not found in first Excel file.")
            if matching_column not in self.df2.columns:
                raise ValueError(f"Column '{matching_column}' not found in second Excel file.")
            if matching_column not in matching_columns:
                raise ValueError(f"Column '{matching_column}' not found in both Excel files.")
        
        self.logger.info(f"Merging on column: '{matching_column}'")
        self.logger.info(f"Merge type: {merge_type}")
        
        try:
            # Perform the merge
            merged_df = pd.merge(self.df1, self.df2, on=matching_column, how=merge_type)
            
            self.logger.info(f"Merged dataframe shape: {merged_df.shape}")
            
            # Save to Excel file
            merged_df.to_excel(output_path, index=False)
            self.logger.info(f"Merged Excel file saved to: {output_path}")
            
        except Exception as e:
            self.logger.error(f"Error during merge operation: {e}")
            raise
        
        return output_path

    def apply_excel_formatting(self, df: pd.DataFrame, formatting_options: Optional[Dict] = None) -> pd.DataFrame:
        """
        Apply formatting options to a DataFrame before saving to Excel.
        
        Args:
            df (pd.DataFrame): DataFrame to format
            formatting_options (Dict, optional): Dictionary containing formatting options
                                                such as column widths, header styles, etc.
                                                
        Returns:
            pd.DataFrame: Formatted DataFrame (for now just returns the original dataframe)
                         In the future, this can be extended to actually format the Excel output
        """
        # For now, just return the dataframe as is
        # In the future, this method can be extended to handle formatting
        if formatting_options:
            self.logger.info(f"Applying formatting options: {formatting_options}")
        return df

    def merge_with_formatting(self, file1_path: str, file2_path: str, output_path: str,
                              merge_type: str = 'inner', matching_column: Optional[str] = None,
                              formatting_options: Optional[Dict] = None) -> str:
        """
        Merge two Excel files with output formatting options.
        
        Args:
            file1_path (str): Path to the first Excel file
            file2_path (str): Path to the second Excel file
            output_path (str): Path to save the merged Excel file
            merge_type (str): Type of merge - 'inner', 'outer', 'left', or 'right'
            matching_column (Optional[str]): Specific column to merge on.
                                           If None, uses first matching column.
            formatting_options (Optional[Dict]): Formatting options for the output file
                                                
        Returns:
            str: Path to the merged Excel file
        """
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils.dataframe import dataframe_to_rows
        
        self.logger.info(f"Starting merge with formatting operation with output path: {output_path}")
        
        # Validate parameters
        self.validate_merge_parameters(output_path, merge_type, matching_column)
        self.validate_file_path(file1_path)
        self.validate_file_path(file2_path)
        
        # Load the files
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)
        
        if df1.empty or df2.empty:
            raise ValueError("One or both Excel files are empty.")
        
        if not os.access(file1_path, os.R_OK):
            raise PermissionError(f"First file is not readable: {file1_path}")
        if not os.access(file2_path, os.R_OK):
            raise PermissionError(f"Second file is not readable: {file2_path}")
        
        # Find matching columns
        matching_columns = list(set(df1.columns) & set(df2.columns))
        
        if not matching_columns:
            raise ValueError("No matching columns found between the two Excel files.")
        
        # If no specific column provided, use the first matching column
        if matching_column is None:
            matching_column = matching_columns[0]
            self.logger.info(f"No specific column provided. Using first matching column: '{matching_column}'")
        else:
            # Validate the specified column exists in both dataframes
            if matching_column not in df1.columns:
                raise ValueError(f"Column '{matching_column}' not found in first Excel file.")
            if matching_column not in df2.columns:
                raise ValueError(f"Column '{matching_column}' not found in second Excel file.")
            if matching_column not in matching_columns:
                raise ValueError(f"Column '{matching_column}' not found in both Excel files.")
        
        self.logger.info(f"Merging on column: '{matching_column}'")
        self.logger.info(f"Merge type: {merge_type}")
        
        try:
            # Perform the merge
            merged_df = pd.merge(df1, df2, on=matching_column, how=merge_type)
            
            # Apply formatting (the dataframe is not modified, but formatting will be applied when saving)
            formatted_df = self.apply_excel_formatting(merged_df, formatting_options)
            
            self.logger.info(f"Merged dataframe shape: {merged_df.shape}")
            
            # Save to Excel file with formatting
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                formatted_df.to_excel(writer, index=False, sheet_name='Merged_Data')
                
                # Get the workbook and worksheet to apply formatting
                workbook = writer.book
                worksheet = writer.sheets['Merged_Data']
                
                # Apply formatting based on options
                if formatting_options:
                    header_font = Font(bold=True, color='FFFFFF')
                    header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
                    header_alignment = Alignment(horizontal='center', vertical='center')
                    
                    # Format header row
                    for cell in worksheet[1:1][0]:
                        cell.font = header_font
                        cell.fill = header_fill
                        cell.alignment = header_alignment
                    
                    # Apply column width if specified
                    if 'column_widths' in formatting_options:
                        for col_idx, width in enumerate(formatting_options['column_widths'], 1):
                            col_letter = chr(64 + col_idx)  # Convert to Excel column letter (A=1, B=2, etc.)
                            worksheet.column_dimensions[col_letter].width = width
                    else:
                        # Auto-adjust column widths
                        for column in worksheet.columns:
                            max_length = 0
                            column_letter = column[0].column_letter
                            for cell in column:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass
                            adjusted_width = min(max_length + 2, 50)  # Limit max width to 50
                            worksheet.column_dimensions[column_letter].width = adjusted_width

            self.logger.info(f"Merged Excel file with formatting saved to: {output_path}")
            
        except Exception as e:
            self.logger.error(f"Error during merge with formatting operation: {e}")
            raise
        
        return output_path

    def track_progress(self, current: int, total: int, description: str = "Processing") -> None:
        """
        Track and display progress for long-running operations.
        
        Args:
            current (int): Current progress value
            total (int): Total value to reach
            description (str): Description of the operation
        """
        import sys
        
        # Calculate percentage
        percentage = (current / total) * 100 if total > 0 else 0
        
        # Create progress bar
        bar_length = 40
        filled_length = int(bar_length * current // total) if total > 0 else 0
        bar = '=' * filled_length + '-' * (bar_length - filled_length)
        
        # Display progress
        sys.stdout.write(f'\r{description}: |{bar}| {percentage:.2f}% ({current}/{total})')
        sys.stdout.flush()
        
        # Finish the progress bar when complete
        if current >= total:
            sys.stdout.write('\n')
            sys.stdout.flush()

    def merge_with_progress_tracking(self, file_paths: List[str], output_path: str,
                                   merge_type: str = 'inner',
                                   matching_column: Optional[str] = None,
                                   chunk_size: int = 1000) -> str:
        """
        Merge multiple Excel files with progress tracking for large files.
        
        Args:
            file_paths (List[str]): List of paths to Excel files to merge
            output_path (str): Path to save the merged Excel file
            merge_type (str): Type of merge - 'inner', 'outer', 'left', or 'right'
            matching_column (Optional[str]): Specific column to merge on.
                                           If None, uses first common column.
            chunk_size (int): Size of chunks to process for progress tracking
            
        Returns:
            str: Path to the merged Excel file
        """
        from tqdm import tqdm
        import os
        
        self.logger.info(f"Starting merge with progress tracking operation with output path: {output_path}")
        
        # Validate parameters
        self.validate_merge_parameters(output_path, merge_type, matching_column)
        
        if len(file_paths) < 2:
            raise ValueError("At least two Excel files must be provided for merging.")
        
        # Load all files with progress tracking
        dataframes = []
        
        print(f"Loading {len(file_paths)} Excel files...")
        
        for i, path in enumerate(file_paths):
            self.validate_file_path(path)
            if not os.access(path, os.R_OK):
                raise PermissionError(f"File is not readable: {path}")
            
            print(f"Loading file {i+1}/{len(file_paths)}: {os.path.basename(path)}")
            
            # Load file
            df = pd.read_excel(path)
            
            if df.empty:
                raise ValueError(f"Excel file is empty: {path}")
            
            dataframes.append(df)
            self.logger.info(f"Loaded {path}, shape: {df.shape}")
            
            # Show progress
            self.track_progress(i + 1, len(file_paths), "Loading files")
        
        # Find common columns
        common_columns = list(set(dataframes[0].columns))
        for df in dataframes[1:]:
            common_columns = list(set(common_columns) & set(df.columns))
        
        if not common_columns:
            raise ValueError("No common columns found between the Excel files.")
        
        # If no specific column provided, use the first common column
        if matching_column is None:
            matching_column = common_columns[0]
            self.logger.info(f"No specific column provided. Using first common column: '{matching_column}'")
        else:
            # Validate the specified column exists in all dataframes
            for i, df in enumerate(dataframes):
                if matching_column not in df.columns:
                    raise ValueError(f"Column '{matching_column}' not found in Excel file {i+1}.")
            
            if matching_column not in common_columns:
                raise ValueError(f"Column '{matching_column}' not found in all Excel files.")
        
        print(f"Merging on column: '{matching_column}'")
        print(f"Merge type: {merge_type}")
        
        # Perform merge with progress tracking for large files
        print("Merging dataframes...")
        
        # For large datasets, we'll use a simple sequential merge with progress tracking
        # since pandas merge operations don't have built-in progress tracking
        merged_df = dataframes[0]
        
        for i, df in enumerate(dataframes[1:], 1):
            print(f"Merging file {i+1}/{len(dataframes)}...")
            
            # Show row counts for progress tracking
            initial_rows = len(merged_df)
            new_rows = len(df)
            
            # Perform the merge
            merged_df = pd.merge(merged_df, df, on=matching_column, how=merge_type)
            
            final_rows = len(merged_df)
            
            print(f"  - Initial rows: {initial_rows}")
            print(f"  - Added rows from file {i+1}: {new_rows}")
            print(f"  - Final rows after merge: {final_rows}")
            
            # Show progress
            self.track_progress(i, len(dataframes) - 1, "Merging files")
        
        self.logger.info(f"Final merged dataframe shape: {merged_df.shape}")
        
        # Save to Excel file with progress tracking if it's a large file
        print(f"Saving merged file to: {output_path}")
        
        # For demonstration, we'll just use the normal save - in a real implementation
        # we might want to chunk the save operation for very large files
        merged_df.to_excel(output_path, index=False)
        
        self.logger.info(f"Merged Excel file saved to: {output_path}")
        
        return output_path

    def load_config(self, config_path: str) -> Dict[str, Any]:
        """
        Load configuration from a JSON or YAML file.
        
        Args:
            config_path (str): Path to the configuration file (.json or .yaml/.yml)
            
        Returns:
            Dict[str, Any]: Configuration dictionary
        """
        import json
        import yaml
        
        self.logger.info(f"Loading configuration from: {config_path}")
        
        # Validate file path
        self.validate_file_path(config_path)
        
        if not os.access(config_path, os.R_OK):
            raise PermissionError(f"Configuration file is not readable: {config_path}")
        
        # Determine file type and load accordingly
        if config_path.lower().endswith('.json'):
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
        elif config_path.lower().endswith(('.yaml', '.yml')):
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
        else:
            raise ValueError(f"Unsupported configuration file format: {config_path}. "
                           f"Supported formats: .json, .yaml, .yml")
        
        self.logger.info(f"Configuration loaded successfully: {config}")
        return config

    def merge_with_config(self, config_path: str) -> str:
        """
        Perform merge operation based on configuration file.
        
        Args:
            config_path (str): Path to the configuration file
            
        Returns:
            str: Path to the merged Excel file
        """
        # Load configuration
        config = self.load_config(config_path)
        
        # Extract parameters from configuration
        file1_path = config.get('file1_path')
        file2_path = config.get('file2_path')
        file_paths = config.get('file_paths', [])  # For multiple file merging
        output_path = config.get('output_path')
        merge_type = config.get('merge_type', 'inner')
        matching_column = config.get('matching_column', None)
        formatting_options = config.get('formatting_options', None)
        
        # Validate required parameters
        if not output_path:
            raise ValueError("output_path is required in configuration")
        
        # Determine if we're doing two-file merge or multiple-file merge
        if file1_path and file2_path:
            # Two file merge
            if formatting_options:
                return self.merge_with_formatting(
                    file1_path, file2_path, output_path,
                    merge_type=merge_type,
                    matching_column=matching_column,
                    formatting_options=formatting_options
                )
            else:
                # Use the original merge_files if we already loaded files or create a new method
                df1 = pd.read_excel(file1_path)
                df2 = pd.read_excel(file2_path)
                
                # Find matching columns
                matching_columns = list(set(df1.columns) & set(df2.columns))
                
                if not matching_columns:
                    raise ValueError("No matching columns found between the two Excel files.")
                
                # If no specific column provided, use the first matching column
                if matching_column is None:
                    matching_column = matching_columns[0]
                
                # Perform the merge
                merged_df = pd.merge(df1, df2, on=matching_column, how=merge_type)
                
                # Save to Excel file
                merged_df.to_excel(output_path, index=False)
                self.logger.info(f"Merged Excel file saved to: {output_path}")
                
                return output_path
        elif file_paths and len(file_paths) >= 2:
            # Multiple file merge using the existing method
            # Load all specified files
            self.load_multiple_excel_files(file_paths)
            return self.merge_multiple_files(output_path, merge_type, matching_column)
        else:
            raise ValueError("Configuration must specify either (file1_path and file2_path) or (file_paths with at least 2 files)")

    async def async_load_excel_file(self, file_path: str) -> pd.DataFrame:
        """
        Asynchronously load an Excel file.
        
        Args:
            file_path (str): Path to the Excel file
            
        Returns:
            pd.DataFrame: Loaded DataFrame
        """
        import asyncio
        
        loop = asyncio.get_event_loop()
        
        # Use run_in_executor to run the file loading in a separate thread
        # since pandas operations are blocking
        df = await loop.run_in_executor(None, pd.read_excel, file_path)
        
        self.logger.info(f"Asynchronously loaded Excel file: {file_path}, shape: {df.shape}")
        
        return df

    async def async_merge_files(self, file1_path: str, file2_path: str, output_path: str,
                               merge_type: str = 'inner', matching_column: Optional[str] = None) -> str:
        """
        Asynchronously merge two Excel files.
        
        Args:
            file1_path (str): Path to the first Excel file
            file2_path (str): Path to the second Excel file
            output_path (str): Path to save the merged Excel file
            merge_type (str): Type of merge - 'inner', 'outer', 'left', or 'right'
            matching_column (Optional[str]): Specific column to merge on.
                                           If None, uses first matching column.
                                           
        Returns:
            str: Path to the merged Excel file
        """
        import asyncio
        
        self.logger.info(f"Starting async merge operation with output path: {output_path}")
        
        # Validate parameters
        self.validate_merge_parameters(output_path, merge_type, matching_column)
        self.validate_file_path(file1_path)
        self.validate_file_path(file2_path)
        
        # Load files asynchronously
        df1_task = self.async_load_excel_file(file1_path)
        df2_task = self.async_load_excel_file(file2_path)
        
        df1, df2 = await asyncio.gather(df1_task, df2_task)
        
        if df1.empty or df2.empty:
            raise ValueError("One or both Excel files are empty.")
        
        # Find matching columns
        matching_columns = list(set(df1.columns) & set(df2.columns))
        
        if not matching_columns:
            raise ValueError("No matching columns found between the two Excel files.")
        
        # If no specific column provided, use the first matching column
        if matching_column is None:
            matching_column = matching_columns[0]
            self.logger.info(f"No specific column provided. Using first matching column: '{matching_column}'")
        else:
            # Validate the specified column exists in both dataframes
            if matching_column not in df1.columns:
                raise ValueError(f"Column '{matching_column}' not found in first Excel file.")
            if matching_column not in df2.columns:
                raise ValueError(f"Column '{matching_column}' not found in second Excel file.")
            if matching_column not in matching_columns:
                raise ValueError(f"Column '{matching_column}' not found in both Excel files.")
        
        self.logger.info(f"Merging on column: '{matching_column}'")
        self.logger.info(f"Merge type: {merge_type}")
        
        # Perform the merge (this is a CPU-bound operation, so we do it in the same thread)
        merged_df = pd.merge(df1, df2, on=matching_column, how=merge_type)
        
        self.logger.info(f"Merged dataframe shape: {merged_df.shape}")
        
        # Save to Excel file asynchronously (use executor for I/O)
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(None, merged_df.to_excel, output_path, False)  # index=False
        
        self.logger.info(f"Merged Excel file saved to: {output_path}")
        
        return output_path


def main() -> None:
    """
    Main function to run the Excel merger application.
    """
    import argparse
    
    parser = argparse.ArgumentParser(description='Merge Excel files based on matching columns.')
    
    # Define arguments
    parser.add_argument('file1', nargs='?', help='Path to the first Excel file')
    parser.add_argument('file2', nargs='?', help='Path to the second Excel file')
    parser.add_argument('-o', '--output', help='Path to save the merged Excel file')
    parser.add_argument('-m', '--merge-type', choices=['inner', 'outer', 'left', 'right'], 
                        default='inner', help='Type of merge (default: inner)')
    parser.add_argument('-c', '--column', help='Specific column to merge on (default: first matching column)')
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose logging')
    parser.add_argument('--config', help='Path to configuration file (JSON or YAML)')
    parser.add_argument('--multiple', nargs='+', help='Multiple files to merge (alternative to file1 and file2)')
    parser.add_argument('--formatting', help='JSON string with formatting options')
    parser.add_argument('--interactive', action='store_true', help='Run in interactive mode to specify files via prompts')
    
    args = parser.parse_args()
    
    # Handle configuration file
    if args.config:
        try:
            merger = ExcelMerger(log_level="DEBUG" if args.verbose else "INFO", log_file="excel_merger.log")
            result_path = merger.merge_with_config(args.config)
            print(f"Successfully merged Excel files using configuration. Output: {result_path}")
            return
        except Exception as e:
            print(f"Error using configuration: {e}")
            return
    
    # Handle interactive mode
    if args.interactive:
        print("Excel Files Merger - Interactive Mode")
        print("="*40)
        
        # Get file paths from user
        file1_path = input("Enter the path of the first Excel file: ").strip()
        file2_path = input("Enter the path of the second Excel file: ").strip()
        
        # Get output path
        output_path = input("Enter the path to save the merged Excel file: ").strip()
        
        # Create ExcelMerger instance with logging
        log_level = "DEBUG" if args.verbose else "INFO"
        merger = ExcelMerger(log_level=log_level, log_file="excel_merger.log")
        
        # Load the Excel files (validation happens inside load_excel_files)
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
        result_path = merger.merge_files(output_path, merge_type, specific_column)
        
        print(f"\nSuccessfully merged the Excel files! Output: {result_path}")
        return
    
    # For non-interactive mode, check if we have required arguments
    if not args.output:
        print("Error: Output path is required. Use -o/--output to specify output file.")
        parser.print_help()
        return
        
    if not args.multiple and (not args.file1 or not args.file2):
        parser.print_help()
        print("\nError: Either specify both input files (file1 and file2) or use --multiple option.")
        print("Or use --interactive flag to run in interactive mode.")
        return
    
    try:
        # Create ExcelMerger instance with logging
        log_level = "DEBUG" if args.verbose else "INFO"
        merger = ExcelMerger(log_level=log_level, log_file="excel_merger.log")
        
        if args.multiple:
            # Handle multiple file merging
            print(f"Merging {len(args.multiple)} Excel files...")
            
            # Load multiple files
            merger.load_multiple_excel_files(args.multiple)
            
            # Find common columns
            common_columns = merger.find_common_columns()
            if common_columns:
                print(f"Common columns found: {common_columns}")
            else:
                print("No common columns found between the Excel files.")
                return
            
            # Perform the merge
            result_path = merger.merge_multiple_files(
                args.output, 
                merge_type=args.merge_type, 
                matching_column=args.column
            )
        else:
            # Handle two file merging with optional formatting
            print(f"Merging Excel files: {args.file1} and {args.file2}")
            
            # Load the Excel files
            merger.load_excel_files(args.file1, args.file2)
            
            # Show matching columns
            matching_columns = merger.find_matching_columns()
            if matching_columns:
                print(f"Matching columns found: {matching_columns}")
            else:
                print("No matching columns found between the two Excel files.")
                return
            
            # Parse formatting options if provided
            formatting_options = None
            if args.formatting:
                import json
                try:
                    formatting_options = json.loads(args.formatting)
                except json.JSONDecodeError as e:
                    print(f"Error parsing formatting options: {e}")
                    return
            
            # Perform the merge with or without formatting
            if formatting_options:
                result_path = merger.merge_with_formatting(
                    args.file1, args.file2, args.output,
                    merge_type=args.merge_type,
                    matching_column=args.column,
                    formatting_options=formatting_options
                )
            else:
                result_path = merger.merge_files(
                    args.output, 
                    merge_type=args.merge_type, 
                    matching_column=args.column
                )
        
        print(f"Successfully merged the Excel files! Output: {result_path}")
        
    except (FileNotFoundError, PermissionError, ValueError) as e:
        print(f"Validation error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


if __name__ == "__main__":
    main()