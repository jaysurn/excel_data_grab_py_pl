# excel_data_grab_py_pl

Grabs line of data from all files in given directory and outputs to an xlsx file for use after.
Same program in Python and Perl ( Client is going with Python version )
*Copy of Original with parts removed for client privacy
*Does not display name reformatter option, but essentially it would make an array of tuples from the input excel (array 1) and then pull a sorted excel of names (array 2) into an array of hashes
  *From there iterate a through a nested loop in starting with array 2 and then array 1 and check if names (regardless of order) any hash in array 2
  *If it does the hash has a proper format key to rewrite the value of the input excel
