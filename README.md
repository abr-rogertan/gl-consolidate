# QBasic Filestitch
A simple QBasic program to combine CSV files into one CSV file. To accomplish that, here are some of the components required.

## Statements
- `INPUT`: To accept input from user so that we can formulate the filename and directory.
- `OPEN`: To open/create a file based on user input.
- `KILL`: To delete a file if it turns out to be empty (optional).
- `CLOSE`: TO close the opened file as a matter of good cleanup practices.
- `PRINT`: To write into an opened write file.

## Functions
- `EOF()`: This tells us if the file pointer has reached the end of the file. Useful for traversing a file line by line.
- `LOF()`: This tells us if a file is empty or missing.
- `CHR$()`: This produces a character from the ASCII set. Usefl if we want to write a double-quote for CSV content..
