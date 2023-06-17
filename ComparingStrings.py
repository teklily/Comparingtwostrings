""" 
Requirements: 
Pandas
openpyxl
The Fuzz - Levenshtein distance library, so you don't have to write out the math in your code.
"""

import pandas as pd
from thefuzz import fuzz

input_string1 = 'Address' #Change 'Address' to the the string you are searching for in the column name for your original file'
input_string2 = 'Address' #Change 'Address' to the the string you are searching for in the column name for the file you are comparing to'
input_unique = 'key' #Change 'key' to the string you are looking for in the original file that might be a unique identifier, like a key or name'
file1 = 'input.xlsx' #location of your original file
file2 = 'official.xlsx' #location of the file you are comparing to.
sim = 0.6 #This is the similarity or more accurately, the Levenshtein distance. Enter in 0.x format.

def compare_strings(string1, string2):
    return fuzz.ratio(string1, string2) / 100

def find_matches(file1, file2, threshold=sim):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    matches = []

    string1_column1 = find_string_column(df1)
    string1_column2 = find_string_column2(df2)
    key_column = find_key_column(df1)

    for index1, row1 in df1.iterrows():
        key = str(row1[key_column])
        string1 = str(row1[string1_column1])
        for index2, row2 in df2.iterrows():
            string2 = str(row2[string1_column2])
            similarity = compare_strings(string1, string2)
            if similarity > threshold:
                match = {
                    input_unique: key,
                    ("Original " + input_string1): string1,
                    ("Matched "+ input_string2): string2,
                    'Similarity': similarity
                }
                matches.append(match)
    output_df = pd.DataFrame(matches)
    writer = pd.ExcelWriter('output.xlsx') 
    output_df.to_excel(writer, sheet_name='Matches', index=False, na_rep='NaN')

# Auto-adjust columns' width
    for column in output_df:
        column_width = max(output_df[column].astype(str).map(len).max(), len(column))
        col_idx = output_df.columns.get_loc(column)
        writer.sheets['Matches'].set_column(col_idx, col_idx, column_width)

    writer.save()

def find_string_column(df):
    for column in df.columns:
        if input_string1.lower() in str(column.lower()):
            return column
    raise ValueError("No column containing " + str({input_string1}) + " found in your file.")
    
def find_string_column2(df):
    for column in df.columns:
        if input_string2.lower() in str(column.lower()):
            return column
    raise ValueError("No column containing " + str({input_string2}) + " found in your file.")


def find_key_column(df):
    for column in df.columns:
        if input_unique.lower() in str(column.lower()):
            return column
    raise ValueError("No column containing " + str({input_unique}) +  " found in your file.")


find_matches(file1, file2, threshold=sim)
print("Data has been successfully written to output.xlsx")
