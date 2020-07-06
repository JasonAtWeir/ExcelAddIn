# ExcelAddIn
My custom excel add ins

This is a bunch of junky functions I made to do stuff.  Most of it is string manipulation, find replace, and Levenshtein comparisons.  I don't rite gud code so don't expect miracles.

To use it, clone the "JasonExcelAddIn.xlam" file to your computer (you can put it anywhere, just make sure it's somewhere where it won't get moved or deleted or whatever).

Then, open Excel, go to the Developer ribbon (if you don't have that ribbon, then you gotta add it; may I suggest googling it?) and select Add-ins.  Click the browse button, then find our friendly little add-in "JasonExcelAddIn.xlam."

Once that's setup, you should have access to all the functions and crap that I've written.  Have at it hoss!

Current included functions:

	Levenshtein(s1 As String, s2 As String)
Calculates the Levenshtein distance between two strings (case sensitive)

	ArrayDupRemove(a1 As String, d1 As String)
Removes duplicate entries from an array formatted as a string with a char delimiter

	ArraySubstitute(sourceStr As String, findA As Range, replaceA As Range)
This is a quick and dirty function to replace elements in a string with a range of values

	DelArrayFromArray(s1 As String, d1 As String, s2 As String, d2 As String)
This function deletes an array of values (s2) from another array (s1). These arrays are formatted as strings (s1, s2) with a delimiter (d1, d2)
