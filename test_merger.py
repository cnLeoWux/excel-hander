# Comprehensive test suite for the Excel merger
import sys
import os
import unittest
import tempfile
import pandas as pd
from unittest.mock import patch, MagicMock

# Add current directory to sys.path to import excel_merger
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from excel_merger import ExcelMerger


class TestExcelMerger(unittest.TestCase):
    """Test suite for ExcelMerger class."""
    
    def setUp(self):
        """Set up test fixtures before each test method."""
        self.merger = ExcelMerger(log_level="ERROR")  # Set to ERROR to reduce noise during testing
        
        # Create temporary files for testing
        self.temp_dir = tempfile.mkdtemp()
        self.sample_file1 = os.path.join(self.temp_dir, 'sample1.xlsx')
        self.sample_file2 = os.path.join(self.temp_dir, 'sample2.xlsx')
        self.output_file = os.path.join(self.temp_dir, 'output.xlsx')
        
        # Create sample data
        data1 = {
            'ID': [1, 2, 3, 4, 5],
            'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
            'Age': [25, 30, 35, 28, 32]
        }
        data2 = {
            'ID': [2, 3, 4, 5, 6],
            'Salary': [50000, 60000, 55000, 62000, 58000],
            'Department': ['IT', 'Finance', 'IT', 'HR', 'Finance']
        }
        
        df1 = pd.DataFrame(data1)
        df2 = pd.DataFrame(data2)
        
        df1.to_excel(self.sample_file1, index=False)
        df2.to_excel(self.sample_file2, index=False)

    def tearDown(self):
        """Clean up test fixtures after each test method."""
        # Clean up temporary files
        for file_path in [self.sample_file1, self.sample_file2, self.output_file]:
            if os.path.exists(file_path):
                os.remove(file_path)
        os.rmdir(self.temp_dir)

    def test_validate_file_path_valid(self):
        """Test validate_file_path with valid file path."""
        # This should not raise an exception
        self.merger.validate_file_path(self.sample_file1)
        
    def test_validate_file_path_empty(self):
        """Test validate_file_path with empty path."""
        with self.assertRaises(ValueError) as context:
            self.merger.validate_file_path("")
        self.assertIn("File path cannot be empty", str(context.exception))
        
    def test_validate_file_path_nonexistent(self):
        """Test validate_file_path with nonexistent file."""
        with self.assertRaises(FileNotFoundError) as context:
            self.merger.validate_file_path("nonexistent_file.xlsx")
        self.assertIn("File does not exist", str(context.exception))
        
    def test_validate_file_path_invalid_extension(self):
        """Test validate_file_path with invalid file extension."""
        with self.assertRaises(ValueError) as context:
            self.merger.validate_file_path("invalid.txt")
        self.assertIn("File is not an Excel file", str(context.exception))
        
    def test_load_excel_files_success(self):
        """Test loading Excel files successfully."""
        self.merger.load_excel_files(self.sample_file1, self.sample_file2)
        
        self.assertIsNotNone(self.merger.df1)
        self.assertIsNotNone(self.merger.df2)
        self.assertEqual(self.merger.df1.shape[0], 5)  # 5 rows
        self.assertEqual(self.merger.df2.shape[0], 5)  # 5 rows
        
    def test_load_excel_files_nonexistent(self):
        """Test loading nonexistent Excel files."""
        with self.assertRaises(FileNotFoundError):
            self.merger.load_excel_files("nonexistent.xlsx", self.sample_file2)
            
    def test_find_matching_columns_success(self):
        """Test finding matching columns between loaded files."""
        self.merger.load_excel_files(self.sample_file1, self.sample_file2)
        matching_cols = self.merger.find_matching_columns()
        
        self.assertIn('ID', matching_cols)
        self.assertEqual(len(matching_cols), 1)  # Only 'ID' should match
        
    def test_find_matching_columns_no_files_loaded(self):
        """Test finding matching columns when no files are loaded."""
        with self.assertRaises(ValueError) as context:
            self.merger.find_matching_columns()
        self.assertIn("Excel files not loaded", str(context.exception))
        
    def test_validate_merge_parameters_valid(self):
        """Test validate_merge_parameters with valid parameters."""
        # This should not raise an exception
        self.merger.validate_merge_parameters(self.output_file, 'inner', 'ID')
        
    def test_validate_merge_parameters_invalid_merge_type(self):
        """Test validate_merge_parameters with invalid merge type."""
        with self.assertRaises(ValueError) as context:
            self.merger.validate_merge_parameters(self.output_file, 'invalid', 'ID')
        self.assertIn("Invalid merge type", str(context.exception))
        
    def test_validate_merge_parameters_empty_output_path(self):
        """Test validate_merge_parameters with empty output path."""
        with self.assertRaises(ValueError) as context:
            self.merger.validate_merge_parameters("", 'inner', 'ID')
        self.assertIn("Output path cannot be empty", str(context.exception))
        
    def test_merge_files_success(self):
        """Test merging files successfully."""
        self.merger.load_excel_files(self.sample_file1, self.sample_file2)
        
        result_path = self.merger.merge_files(self.output_file, merge_type='inner', matching_column='ID')
        
        self.assertEqual(result_path, self.output_file)
        self.assertTrue(os.path.exists(result_path))
        
        # Check that merged file has expected content
        merged_df = pd.read_excel(result_path)
        # Inner join on ID should result in 4 rows (IDs 2, 3, 4, 5 are in both files)
        self.assertEqual(merged_df.shape[0], 4)
        
    def test_merge_files_no_matching_columns(self):
        """Test merging files with no matching columns."""
        # Create a file with different columns
        different_cols_file = os.path.join(self.temp_dir, 'different_cols.xlsx')
        data = {
            'Name': ['Alice', 'Bob'],
            'City': ['New York', 'London']
        }
        pd.DataFrame(data).to_excel(different_cols_file, index=False)
        
        try:
            self.merger.load_excel_files(self.sample_file1, different_cols_file)
            with self.assertRaises(ValueError) as context:
                self.merger.merge_files(self.output_file, merge_type='inner')
            self.assertIn("No matching columns found", str(context.exception))
        finally:
            if os.path.exists(different_cols_file):
                os.remove(different_cols_file)
                
    def test_merge_files_invalid_column(self):
        """Test merging files with invalid column name."""
        self.merger.load_excel_files(self.sample_file1, self.sample_file2)
        
        with self.assertRaises(ValueError) as context:
            self.merger.merge_files(self.output_file, merge_type='inner', matching_column='NonExistent')
        self.assertIn("not found in first Excel file", str(context.exception))
        
    def test_merge_files_different_merge_types(self):
        """Test different merge types."""
        self.merger.load_excel_files(self.sample_file1, self.sample_file2)
        
        # Test inner merge
        inner_path = os.path.join(self.temp_dir, 'inner_output.xlsx')
        self.merger.merge_files(inner_path, merge_type='inner', matching_column='ID')
        inner_df = pd.read_excel(inner_path)
        
        # Inner merge should have 4 rows (common IDs: 2,3,4,5)
        self.assertEqual(inner_df.shape[0], 4)
        
        # Test outer merge
        outer_path = os.path.join(self.temp_dir, 'outer_output.xlsx')
        self.merger.merge_files(outer_path, merge_type='outer', matching_column='ID')
        outer_df = pd.read_excel(outer_path)
        
        # Outer merge should have 6 rows (all IDs: 1,2,3,4,5,6)
        self.assertEqual(outer_df.shape[0], 6)


    def test_register_and_execute_custom_strategy(self):
        """Test registering and executing custom merge strategies."""
        # Create a simple custom merge strategy
        def custom_strategy(df1, df2, **kwargs):
            # Simple strategy: concatenate the dataframes
            return pd.concat([df1, df2], ignore_index=True, sort=False)
        
        self.merger.register_merge_strategy("concat", custom_strategy)
        
        # Load test files
        self.merger.load_excel_files(self.sample_file1, self.sample_file2)
        
        # Execute custom strategy
        result_path = os.path.join(self.temp_dir, 'custom_output.xlsx')
        output_path = self.merger.execute_custom_merge_strategy(
            "concat", self.sample_file1, self.sample_file2, result_path
        )
        
        self.assertTrue(os.path.exists(output_path))
        # Concatenation would result in 10 rows (5 from each file)
        result_df = pd.read_excel(output_path)
        self.assertEqual(result_df.shape[0], 10)

    def test_load_config_json(self):
        """Test loading configuration from JSON file."""
        # Create a temporary config file
        config_path = os.path.join(self.temp_dir, 'test_config.json')
        config_data = {
            "file1_path": self.sample_file1,
            "file2_path": self.sample_file2,
            "output_path": os.path.join(self.temp_dir, 'config_output.xlsx'),
            "merge_type": "inner",
            "matching_column": "ID"
        }
        
        import json
        with open(config_path, 'w') as f:
            json.dump(config_data, f)
        
        # Test loading config
        config = self.merger.load_config(config_path)
        self.assertEqual(config['file1_path'], self.sample_file1)
        self.assertEqual(config['merge_type'], 'inner')

    def test_merge_with_config(self):
        """Test performing merge using configuration file."""
        # Create a temporary config file
        config_path = os.path.join(self.temp_dir, 'test_merge_config.json')
        config_data = {
            "file1_path": self.sample_file1,
            "file2_path": self.sample_file2,
            "output_path": os.path.join(self.temp_dir, 'configured_merge_output.xlsx'),
            "merge_type": "inner",
            "matching_column": "ID"
        }
        
        import json
        with open(config_path, 'w') as f:
            json.dump(config_data, f)
        
        # Test merge with config
        result_path = self.merger.merge_with_config(config_path)
        self.assertTrue(os.path.exists(result_path))
        result_df = pd.read_excel(result_path)
        # Inner merge on ID should result in 4 rows (IDs 2,3,4,5 are in both files)
        self.assertEqual(result_df.shape[0], 4)

    def test_apply_excel_formatting(self):
        """Test applying Excel formatting (placeholder test)."""
        df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
        formatted_df = self.merger.apply_excel_formatting(df, {"theme": "dark"})
        # Currently this returns the same dataframe, so just check it's still a dataframe
        self.assertIsInstance(formatted_df, pd.DataFrame)
        pd.testing.assert_frame_equal(df, formatted_df)

    def test_merge_with_formatting(self):
        """Test merging with formatting options."""
        output_path = os.path.join(self.temp_dir, 'formatted_output.xlsx')
        
        result_path = self.merger.merge_with_formatting(
            self.sample_file1, self.sample_file2, output_path,
            merge_type='inner', matching_column='ID',
            formatting_options={'theme': 'dark'}
        )
        
        self.assertTrue(os.path.exists(result_path))
        result_df = pd.read_excel(result_path)
        # Inner merge should result in 4 rows
        self.assertEqual(result_df.shape[0], 4)

    @unittest.skip("Async tests require event loop management in unittest environment")
    def test_async_merge_files(self):
        """Test asynchronous file merging."""
        # This test is skipped because async methods require special handling in unittest
        pass

    def test_find_potential_matches(self):
        """Test finding potential column matches based on similarity."""
        # Create test dataframes with similar column names
        df1 = pd.DataFrame({
            'Name': ['Alice', 'Bob'],
            'Age': [25, 30],
            'Email_Address': ['alice@example.com', 'bob@example.com']
        })
        
        df2 = pd.DataFrame({
            'Full_Name': ['Charlie', 'David'],  # Similar to 'Name'
            'Age': [35, 40],
            'Email': ['charlie@example.com', 'david@example.com']  # Similar to 'Email_Address'
        })
        
        # Find potential matches with high similarity threshold
        matches = self.merger.find_potential_matches(df1, df2, similarity_threshold=0.3)
        
        # Should find 'Age' as a match
        self.assertIn('Age', matches)
        self.assertEqual(matches['Age'], 'Age')

    def test_merge_with_custom_condition(self):
        """Test merging with custom condition."""
        output_path = os.path.join(self.temp_dir, 'custom_condition_output.xlsx')
        
        # Custom condition function to merge based on Age >= 30
        def age_condition(df1, df2):
            # Just merge all rows for testing purposes
            return pd.merge(df1, df2, left_on='Age', right_on='ID', how='inner')
        
        result_path = self.merger.merge_with_custom_condition(
            self.sample_file1, self.sample_file2, output_path,
            merge_type='inner', custom_condition=age_condition
        )
        
        self.assertTrue(os.path.exists(result_path))

    def test_range_merge_condition(self):
        """Test range-based merge condition."""
        df1 = pd.DataFrame({
            'ID': [1, 2, 3],
            'Value1': [10, 20, 30]
        })
        
        df2 = pd.DataFrame({
            'ID': [1, 2, 3],
            'Value2': [15, 25, 35]
        })
        
        # Test various operators
        result = self.merger.range_merge_condition(df1, df2, 'Value1', 'Value2', operator='>')
        # This should return a cross join filtered by Value1 > Value2 for matching rows
        
        # For this simple test, just ensure it returns a DataFrame
        self.assertIsInstance(result, pd.DataFrame)


def run_comprehensive_tests():
    """Run the comprehensive test suite."""
    print("Running comprehensive Excel Merger tests...")
    
    # Create a test suite
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromTestCase(TestExcelMerger)
    
    # Run the tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # Print summary
    print(f"\nTests run: {result.testsRun}")
    print(f"Failures: {len(result.failures)}")
    print(f"Errors: {len(result.errors)}")
    print(f"Success: {result.wasSuccessful()}")
    
    return result.wasSuccessful()


class TestExcelMergerIntegration(unittest.TestCase):
    """Integration tests for ExcelMerger class."""
    
    def setUp(self):
        """Set up test fixtures before each test method."""
        self.merger = ExcelMerger(log_level="ERROR")  # Set to ERROR to reduce noise during testing
        
        # Create temporary files for testing
        self.temp_dir = tempfile.mkdtemp()
        self.sample_file1 = os.path.join(self.temp_dir, 'sample1.xlsx')
        self.sample_file2 = os.path.join(self.temp_dir, 'sample2.xlsx')
        self.sample_file3 = os.path.join(self.temp_dir, 'sample3.xlsx')
        self.output_file = os.path.join(self.temp_dir, 'output.xlsx')
        
        # Create sample data
        data1 = {
            'ID': [1, 2, 3, 4, 5],
            'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
            'Age': [25, 30, 35, 28, 32]
        }
        data2 = {
            'ID': [2, 3, 4, 5, 6],
            'Salary': [50000, 60000, 55000, 62000, 58000],
            'Department': ['IT', 'Finance', 'IT', 'HR', 'Finance']
        }
        data3 = {
            'ID': [1, 3, 5, 7],
            'Location': ['New York', 'London', 'Paris', 'Tokyo'],
            'Experience': [2, 8, 5, 10]
        }
        
        df1 = pd.DataFrame(data1)
        df2 = pd.DataFrame(data2)
        df3 = pd.DataFrame(data3)
        
        df1.to_excel(self.sample_file1, index=False)
        df2.to_excel(self.sample_file2, index=False)
        df3.to_excel(self.sample_file3, index=False)

    def tearDown(self):
        """Clean up test fixtures after each test method."""
        # Clean up temporary files
        for file_path in [self.sample_file1, self.sample_file2, self.sample_file3, self.output_file]:
            if os.path.exists(file_path):
                os.remove(file_path)
        os.rmdir(self.temp_dir)

    def test_end_to_end_merge(self):
        """Test the complete workflow from loading to saving."""
        # Load files
        self.merger.load_excel_files(self.sample_file1, self.sample_file2)
        
        # Find matching columns
        matching_cols = self.merger.find_matching_columns()
        self.assertIn('ID', matching_cols)
        
        # Perform merge
        result_path = self.merger.merge_files(self.output_file, merge_type='inner', matching_column='ID')
        
        # Verify output
        self.assertTrue(os.path.exists(result_path))
        result_df = pd.read_excel(result_path)
        self.assertEqual(result_df.shape[0], 4)  # Should be 4 rows after inner join on ID

    def test_multiple_file_merge_workflow(self):
        """Test the multiple file merge workflow."""
        # Load multiple files
        self.merger.load_multiple_excel_files([self.sample_file1, self.sample_file2, self.sample_file3])
        
        # Find common columns
        common_cols = self.merger.find_common_columns()
        self.assertIn('ID', common_cols)
        
        # Perform multiple file merge
        output_path = os.path.join(self.temp_dir, 'multi_merge_output.xlsx')
        result_path = self.merger.merge_multiple_files(output_path, merge_type='inner', matching_column='ID')
        
        # Verify output
        self.assertTrue(os.path.exists(result_path))
        result_df = pd.read_excel(result_path)
        # Inner merge should result in rows where ID exists in all 3 files (only ID=3,5)
        self.assertEqual(result_df.shape[0], 2)

    def test_column_mapping_merge_workflow(self):
        """Test the column mapping merge workflow."""
        # Create files with similar but not identical column names
        alt_file = os.path.join(self.temp_dir, 'alt_sample.xlsx')
        alt_data = {
            'Employee_ID': [2, 3, 4, 5, 6],  # Similar to ID in original files
            'Wage': [50000, 60000, 55000, 62000, 58000],
            'Department': ['IT', 'Finance', 'IT', 'HR', 'Finance']
        }
        pd.DataFrame(alt_data).to_excel(alt_file, index=False)
        
        try:
            # Test auto-detection of similar columns
            result_df = self.merger.map_columns_for_merge(self.sample_file1, alt_file)
            self.assertIsInstance(result_df, pd.DataFrame)
            # This test is more complex, so we're just ensuring it doesn't error
        finally:
            if os.path.exists(alt_file):
                os.remove(alt_file)

    def test_config_based_workflow(self):
        """Test the complete workflow using configuration file."""
        # Create config file
        config_path = os.path.join(self.temp_dir, 'workflow_config.json')
        config_data = {
            "file1_path": self.sample_file1,
            "file2_path": self.sample_file2,
            "output_path": os.path.join(self.temp_dir, 'config_workflow_output.xlsx'),
            "merge_type": "outer",
            "matching_column": "ID"
        }
        
        import json
        with open(config_path, 'w') as f:
            json.dump(config_data, f)
        
        # Execute merge using config
        result_path = self.merger.merge_with_config(config_path)
        
        # Verify output
        self.assertTrue(os.path.exists(result_path))
        result_df = pd.read_excel(result_path)
        # Outer merge should have all IDs from both files (IDs 1,2,3,4,5,6)
        self.assertEqual(result_df.shape[0], 6)

    def test_custom_strategy_workflow(self):
        """Test workflow with custom merge strategies."""
        # Define a custom strategy
        def custom_avg_strategy(df1, df2):
            # Merge and add an average of a numeric column
            merged = pd.merge(df1, df2, on='ID', how='inner')
            return merged

        # Register the strategy
        self.merger.register_merge_strategy("average_merge", custom_avg_strategy)
        
        # Execute using custom strategy
        output_path = os.path.join(self.temp_dir, 'custom_strategy_output.xlsx')
        result_path = self.merger.execute_custom_merge_strategy(
            "average_merge", self.sample_file1, self.sample_file2, output_path
        )
        
        # Verify output
        self.assertTrue(os.path.exists(result_path))
        result_df = pd.read_excel(result_path)
        self.assertEqual(result_df.shape[0], 4)  # Inner merge result


def test_merger():
    """Original test function preserved for compatibility."""
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
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == 'comprehensive':
        success = run_comprehensive_tests()
        sys.exit(0 if success else 1)
    elif len(sys.argv) > 1 and sys.argv[1] == 'integration':
        print("Running Excel Merger integration tests...")
        loader = unittest.TestLoader()
        suite = loader.loadTestsFromTestCase(TestExcelMergerIntegration)
        runner = unittest.TextTestRunner(verbosity=2)
        result = runner.run(suite)
        
        print(f"\nIntegration Tests run: {result.testsRun}")
        print(f"Failures: {len(result.failures)}")
        print(f"Errors: {len(result.errors)}")
        print(f"Success: {result.wasSuccessful()}")
        
        sys.exit(0 if result.wasSuccessful() else 1)
    elif len(sys.argv) > 1 and sys.argv[1] == 'performance':
        print("Running Excel Merger performance tests...")
        
        import time
        import pandas as pd
        import tempfile
        import os
        
        # Create large test files
        temp_dir = tempfile.mkdtemp()
        large_file1 = os.path.join(temp_dir, 'large_file1.xlsx')
        large_file2 = os.path.join(temp_dir, 'large_file2.xlsx')
        output_file = os.path.join(temp_dir, 'large_merged_output.xlsx')
        
        try:
            print("Creating large test datasets...")
            
            # Create large datasets with 10,000 rows each
            size = 10000
            df1 = pd.DataFrame({
                'ID': range(1, size + 1),
                'Value1': [f'value_{i}' for i in range(1, size + 1)],
                'Number1': range(size, 0, -1)
            })
            
            df2 = pd.DataFrame({
                'ID': range(5000, size + 5001),  # Overlap with df1: IDs 5000-10000
                'Value2': [f'other_value_{i}' for i in range(5000, size + 5001)],
                'Number2': range(size * 2, 0, -2)
            })
            
            # Save large files
            start_time = time.time()
            df1.to_excel(large_file1, index=False)
            df2.to_excel(large_file2, index=False)
            save_time = time.time() - start_time
            print(f"Large files created and saved in {save_time:.2f} seconds")
            
            # Test performance of merge
            merger = ExcelMerger(log_level="ERROR")
            
            print("Starting performance test for merge operation...")
            start_time = time.time()
            merger.load_excel_files(large_file1, large_file2)
            load_time = time.time() - start_time
            print(f"Files loaded in {load_time:.2f} seconds")
            
            start_time = time.time()
            result_path = merger.merge_files(output_file, merge_type='inner', matching_column='ID')
            merge_time = time.time() - start_time
            print(f"Merge operation completed in {merge_time:.2f} seconds")
            
            # Verify output
            result_df = pd.read_excel(result_path)
            print(f"Resulting dataframe has {result_df.shape[0]} rows and {result_df.shape[1]} columns")
            
            print(f"\nPerformance Summary:")
            print(f"- File loading time: {load_time:.2f} seconds")
            print(f"- Merge operation time: {merge_time:.2f} seconds")
            print(f"- Total processed rows: ~{size} per file")
            print(f"- Result rows: {result_df.shape[0]}")
            
        finally:
            # Cleanup
            for file_path in [large_file1, large_file2, output_file]:
                if os.path.exists(file_path):
                    os.remove(file_path)
            os.rmdir(temp_dir)
        
        print("\nPerformance test completed!")
    else:
        test_merger()