# My Excel Macros
These are personal excel macros that I have found helpful in day to day data manipulation

## Combine
This function returns a string of concatenated values with an optional seperator. This is similar to the TEXTJOIN functionality in the 2019 version of excel.

`Combine(WorkRng As Range, Optional Sign As String = ",", Optional IgnoreEmpty As Boolean = True) As String`

*WorkRng* - The range of cells you want concatenated 
*Sign* - The delimeter you want to use [default = ',']
*IgnoreEmpty* - Whether or not you want to ignore blank cells [default = True]