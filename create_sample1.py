import pandas as pd

# Create sample data for first Excel file
data1 = {
    'ID': [1, 2, 3, 4, 5],
    'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
    'Age': [25, 30, 35, 28, 32],
    'Department': ['HR', 'IT', 'Finance', 'IT', 'HR']
}

df1 = pd.DataFrame(data1)
df1.to_excel('sample_file1.xlsx', index=False)
print("Created sample_file1.xlsx")