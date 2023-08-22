import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

# Set display options to show all rows and columns
pd.set_option('display.max_rows', None)

# Find the file path
file_path = 'insert your file path here'

# Load the workbook
workbook = openpyxl.load_workbook(file_path)

# Read the first sheet from the Excel file
df = pd.read_excel(file_path, sheet_name='Sign-In', skiprows=0)

# Specify the desired row height
row_height = 30

# Find the columns we want 
selected_columns_df = df.iloc[:, [0, 1, 10, 11, 18, 19]]

# Set the column names
column_names = ['First Name', 'Last Name', 'Time In', 'Time Out', 'Authorized Pickup', 'Authorized Dropoff']

# Specify the columns to check for NaN values
columns_to_check = [0, 1]

# Filter any entries with NaN values in the specified columns
filtered_df = selected_columns_df.dropna(subset=selected_columns_df.columns[columns_to_check])

# Sort the rows by the second column (column index 1) in alphabetical order
filtered_df = filtered_df.sort_values(by=filtered_df.columns[1])

# Specify the column index number for deleting non students from the filtered data frame
column_index = 1  # Replace with the actual column index

# Filter out rows containing "IT INTERN" or "director" in the specified column
filtered_df = filtered_df[~filtered_df.iloc[:, column_index].astype(str).str.contains('IT INTERN|Director|TI|Last Name')]

# Specify the authorized pickup and dropoff columns to be combined into a single column
column1_index = 4 
column2_index = 5 

# Combine the information from the two columns into a new column
filtered_df['Combined_Column'] = filtered_df.iloc[:, column1_index].astype(str) + ' ' + filtered_df.iloc[:, column2_index].astype(str)

# Drop the old columns from the DataFrame
filtered_df = filtered_df.drop(columns=[filtered_df.columns[column1_index], filtered_df.columns[column2_index]])

# Specify the index of the column to remove the <br>'s 
column_index = 4  

# Filter out "\n" and "<br />" from the cells of the pickup column
filtered_df.iloc[:, column_index] = filtered_df.iloc[:, column_index].str.replace('\n', '').str.replace('<br />', '')

sheet = workbook.active

#print(filtered_df)

# Save the modified DataFrame to an Excel file without column numbers
output_file_path = 'Insert the file path you want the file to be made too'
filtered_df.to_excel(output_file_path, index=False, index_label=None)


