# Comparingtwostrings

Requirements: 
Pandas
openpyxl
The Fuzz - Levenshtein distance library, so you don't have to write out the math in your code.

Comparing two excel files and two strings using the Levenshtein Distance.

It will contain a unique identifier column, the original search string, the second search string and the "Similarity", but more accurately, the Levenshtein Distance.

The script will also auto adjust the width of the columns.
I tried to make the search strings variables that can be entered at the top of the search string, to make it easy for any user to use.
I debated adding an input instead, but I decided against that.

The script will write to an output.xlsx file. 


A thing to note, the output.xlsx file must not be open or you will receive an error.
