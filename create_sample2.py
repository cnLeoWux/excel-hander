import pandas as pd

# Create sample data for second Excel file
data2 = {
    'ID': [2, 3, 4, 5, 6],
    'Salary': [50000, 60000, 55000, 62000, 58000],
    'Location': ['New York', 'London', 'Paris', 'Tokyo', 'Sydney'],
    'Experience': [3, 8, 5, 7, 6]
}

df2 = pd.DataFrame(data2)
df2.to_excel('sample_file2.xlsx', index=False)
print("Created sample_file2.xlsx")