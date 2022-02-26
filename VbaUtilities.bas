Attribute VB_Name = "VbaUtilities"
' ========================================
' Last Revision Date = 02/08/2022
' Module: VBA Utilities
' Developed by: Mark Kranz
' ========================================

' ========================================
' Function IsInArray
'
'Purpose:       Return a Boolean value on whether a specific
'               string exists in an array.
'
' Arguments:
'       s_stringToBeFound (string)      - Path to directory.
'       s_arr (variant) - Path to directory.
'
' Return Value (Boolean)
'       =True, then string exists
'       =False, then string does not exist
'
' ========================================

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

' ========================================
' Function NumberOfRows
'
'Purpose:       Return the number of rows in a specific sheet.
'
' Arguments:
'       SheetName (String)    - Sheet Name.
'       KeyColumn  (Integer)  - Column to reference.
'
' Return Value (Long)
'       Number of rows to last non-empty cell in specified column
'
' ========================================

Function NumberOfRows(sheetName As String, KeyColumn As Long) As Long
    Dim sheet As Worksheet
    
    Set sheet = ThisWorkbook.Sheets(sheetName)
    NumberOfRows = sheet.Cells(sheet.Rows.Count, KeyColumn).End(xlUp).row
End Function




