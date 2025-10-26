#!/usr/bin/env python
"""
Script to generate API documentation for the Excel merger project.
This script extracts docstrings and generates documentation in various formats.
"""

import inspect
import os
from excel_merger import ExcelMerger


def generate_api_docs():
    """
    Generate API documentation for the ExcelMerger class.
    """
    print("Generating API Documentation for ExcelMerger...")
    
    # Get the ExcelMerger class
    cls = ExcelMerger
    
    # Create documentation content
    docs_content = []
    docs_content.append("# ExcelMerger API Documentation\n")
    docs_content.append(f"Class: `{cls.__name__}`\n")
    
    if cls.__doc__:
        docs_content.append(f"## Description\n{cls.__doc__}\n")
    
    # Document constructor
    docs_content.append("## Constructor\n")
    docs_content.append("```python")
    docs_content.append(f"def {cls.__init__.__name__}(self, log_level: str = 'INFO', log_file: str = 'excel_merger.log'):")
    docs_content.append("```\n")
    
    if cls.__init__.__doc__:
        docs_content.append(f"{cls.__init__.__doc__}\n")
    
    # Document all public methods
    docs_content.append("## Methods\n")
    
    # Get all methods of the class
    methods = inspect.getmembers(cls, predicate=inspect.isfunction)
    
    for name, method in methods:
        # Skip private methods and special methods (unless they provide important functionality)
        if not name.startswith('_') or name in ['__init__']:
            docs_content.append(f"### {name}\n")
            
            # Get the signature
            try:
                sig = inspect.signature(method)
                docs_content.append(f"```python")
                docs_content.append(f"def {name}{sig}:")
                docs_content.append("```\n")
            except (ValueError, TypeError):
                # If we can't get signature, just add function name
                docs_content.append(f"```python")
                docs_content.append(f"def {name}(self, ...):")
                docs_content.append("```\n")
            
            # Add docstring if available
            if method.__doc__:
                docs_content.append(f"{method.__doc__}\n")
            else:
                docs_content.append("No documentation available.\n")
    
    # Join all documentation content
    full_docs = "\n".join(docs_content)
    
    # Write to file
    docs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "API_DOCS.md")
    with open(docs_path, 'w', encoding='utf-8') as f:
        f.write(full_docs)
    
    print(f"API documentation generated: {docs_path}")
    
    return full_docs


def generate_simple_docs():
    """
    Generate simple parameter and return documentation for the key methods.
    """
    print("Generating Simple Parameter Documentation...")
    
    # Define the key methods and their expected parameters/returns
    method_docs = {
        "load_excel_files": {
            "params": [
                ("file1_path", "str", "Path to the first Excel file"),
                ("file2_path", "str", "Path to the second Excel file")
            ],
            "returns": ("None", "Loads files into internal dataframes"),
            "description": "Load two Excel files into pandas dataframes."
        },
        "merge_files": {
            "params": [
                ("output_path", "str", "Path to save the merged Excel file"),
                ("merge_type", "str", "Type of merge - 'inner', 'outer', 'left', or 'right' (default 'inner')"),
                ("matching_column", "str or None", "Specific column to merge on. If None, uses first matching column")
            ],
            "returns": ("str", "Path to the merged Excel file"),
            "description": "Merge the two Excel files based on matching columns."
        },
        "load_multiple_excel_files": {
            "params": [
                ("file_paths", "List[str]", "List of paths to Excel files to load")
            ],
            "returns": ("None", "Loads files into internal dataframes list"),
            "description": "Load multiple Excel files into a list of dataframes."
        },
        "merge_multiple_files": {
            "params": [
                ("output_path", "str", "Path to save the merged Excel file"),
                ("merge_type", "str", "Type of merge - 'inner', 'outer', 'left', or 'right' (default 'inner')"),
                ("matching_column", "str or None", "Specific column to merge on. If None, uses first common column"),
                ("sequential", "bool", "If True, merge files sequentially. If False, use reduce (default True)")
            ],
            "returns": ("str", "Path to the merged Excel file"),
            "description": "Merge multiple Excel files based on matching columns."
        },
        "register_merge_strategy": {
            "params": [
                ("name", "str", "Name of the strategy"),
                ("strategy_func", "callable", "Function that takes two DataFrames and returns a merged DataFrame")
            ],
            "returns": ("None", ""),
            "description": "Register a custom merge strategy."
        },
        "execute_custom_merge_strategy": {
            "params": [
                ("strategy_name", "str", "Name of the registered strategy to execute"),
                ("file1_path", "str", "Path to the first Excel file"),
                ("file2_path", "str", "Path to the second Excel file"),
                ("output_path", "str", "Path to save the merged Excel file"),
                ("**kwargs", "dict", "Additional arguments to pass to the strategy function")
            ],
            "returns": ("str", "Path to the merged Excel file"),
            "description": "Execute a registered custom merge strategy."
        }
    }
    
    docs_content = []
    docs_content.append("# ExcelMerger API Reference\n")
    docs_content.append("Detailed documentation for key methods in the ExcelMerger class.\n")
    
    for method_name, info in method_docs.items():
        docs_content.append(f"## {method_name}\n")
        docs_content.append(f"{info['description']}\n")
        
        # Parameters
        docs_content.append("### Parameters\n")
        for param_name, param_type, desc in info['params']:
            docs_content.append(f"- `{param_name}` ({param_type}): {desc}\n")
        
        # Returns
        return_type, return_desc = info['returns']
        docs_content.append(f"### Returns\n")
        docs_content.append(f"- ({return_type}): {return_desc}\n\n")
    
    # Write to file
    docs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "API_REFERENCE.md")
    with open(docs_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(docs_content))
    
    print(f"API reference documentation generated: {docs_path}")
    
    return "\n".join(docs_content)


if __name__ == "__main__":
    print("Excel Merger - API Documentation Generator")
    print("=" * 45)
    
    # Generate full API documentation
    full_docs = generate_api_docs()
    
    print()
    
    # Generate simplified reference
    simple_docs = generate_simple_docs()
    
    print("\nAPI documentation generation completed!")
    print("Two files created:")
    print("- API_DOCS.md: Full documentation with all methods")
    print("- API_REFERENCE.md: Key methods reference with parameters")