#!/usr/bin/env python
"""
Script to run code coverage analysis for the Excel merger project.
"""

import subprocess
import sys
import os

def run_coverage_analysis():
    """
    Run code coverage analysis on the Excel merger project.
    """
    print("Running code coverage analysis...")
    
    try:
        # Install coverage package if not already installed
        import coverage
    except ImportError:
        print("Installing coverage package...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "coverage"])
        import coverage
    
    # Create a coverage object
    cov = coverage.Coverage()
    
    # Start coverage measurement
    cov.start()
    
    # Import and run tests
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    
    # Run unit tests
    from test_merger import run_comprehensive_tests
    run_comprehensive_tests()
    
    # Run integration tests
    import unittest
    from test_merger import TestExcelMergerIntegration
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromTestCase(TestExcelMergerIntegration)
    runner = unittest.TextTestRunner(verbosity=0)
    runner.run(suite)
    
    # Stop coverage measurement
    cov.stop()
    cov.save()
    
    # Report coverage statistics
    print("\nCoverage Report:")
    print("=" * 50)
    cov.report(show_missing=True)
    
    # Generate HTML report
    html_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "htmlcov")
    cov.html_report(directory=html_dir)
    print(f"\nHTML coverage report generated in: {html_dir}")
    
    # Generate XML report for CI/CD systems
    xml_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "coverage.xml")
    cov.xml_report(outfile=xml_file)
    print(f"XML coverage report generated: {xml_file}")
    
    # Return coverage percentage
    _, _, total_statements, covered_statements = cov.analysis2(os.path.join(os.path.dirname(__file__), "excel_merger.py"))
    if total_statements > 0:
        coverage_percentage = (covered_statements / total_statements) * 100
        print(f"\nCoverage percentage for excel_merger.py: {coverage_percentage:.2f}%")
        return coverage_percentage
    else:
        print("\nCould not calculate coverage for excel_merger.py")
        return 0

def run_coverage_with_subprocess():
    """
    Alternative method using subprocess to run coverage from command line.
    """
    print("Running code coverage analysis via command line...")
    
    try:
        # Run coverage on the test file
        subprocess.check_call([
            sys.executable, "-m", "coverage", "run", "--source=excel_merger.py", 
            "-m", "unittest", "discover", "-s", ".", "-p", "test_*.py"
        ])
        
        # Generate report
        subprocess.check_call([sys.executable, "-m", "coverage", "report", "-m"])
        
        # Generate HTML report
        subprocess.check_call([sys.executable, "-m", "coverage", "html"])
        
        # Generate XML report
        subprocess.check_call([sys.executable, "-m", "coverage", "xml"])
        
        print("\nCoverage analysis completed! Reports generated.")
        
    except subprocess.CalledProcessError as e:
        print(f"Error running coverage: {e}")
        return None
    except FileNotFoundError:
        print("Coverage module not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "coverage"])
        # Retry the process
        run_coverage_with_subprocess()

if __name__ == "__main__":
    print("Excel Merger - Code Coverage Analysis")
    print("=" * 40)
    
    # Try the subprocess method first as it's more standard
    run_coverage_with_subprocess()