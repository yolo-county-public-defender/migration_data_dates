### TO-DO

1. there are overflow notes occassionally falling in columns D, E, and F. When this occurs, the user ID that is supposed to be in column E is pushed to column G or H, for these cases we need to:
   1. get rid of the overflow notes entirely, meaning when there is a string of acsii characters that are not equivalent to a number, clear out the cell in column of D, E, F, G, or H, and move the user Id that is a five digit number to the corresponding row's Column E.
   2. When there is "NULL" found in column E after doing all of the aforementioned, clear out said cell leaving it blank.
   3. Finally, proceed, with our initial movement of timestamp from the end of column B into column D, converting the timestamp to a 24 hour formatted timestamp.
   4. Make sure to copy ALL sheets, even the others that we aren't editing (we're editing the Note sheet) so that the final excel modified spreadsheet has all of the same content from the starting excel spreadsheet.
