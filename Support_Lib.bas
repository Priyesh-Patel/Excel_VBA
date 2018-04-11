Attribute VB_Name = "Support_Lib"
Option Explicit

''-----------------------------------------------------------------------------------------------------------------------------------------
'''Name:
'''Author: Priyesh Patel
'''Puropose:
'''Dependancies:
'''Examples of Use:
'''-----------------------------------------------------------------------------------------------------------------------------------------

''Table of Contents:

''##Non Dependant
''Column Number to Letter v1.0
''Pattern Match V1.0
''Turn On Auto Filter v1.0
''Get Column Or Row Address V1.0
''Find Column v1.0
''Check Named Range Exists v1.0
''Convert Column To Number v1.0
''
''##Dependant
''Get Last Cell Address v1.0
''



''##########################################################################################################################################
''Non Dependant Helper Funtions
''##########################################################################################################################################

Function colLtr(iCol As Long) As String

''-----------------------------------------------------------------------------------------------------------------------------------------
'''Name: Column Number to Letter v1.0
'''Author: Priyesh Patel
'''Puropose: convert a Column number address to the letter.
'''Dependancies: none
'''Examples of Use: x = ColLtr(1) <- returns A
'''-----------------------------------------------------------------------------------------------------------------------------------------

If iCol > 0 And iCol <= Columns.Count Then colLtr = Replace(Cells(1, iCol).Address(0, 0), 1, "")

End Function



Public Function PatternMatch(strValueToID As String, strPattern) As Boolean
''-----------------------------------------------------------------------------------------------------------------------------------------
'''Name:Pattern Match V1.0
'''
'''Author: Priyesh Patel
'''
'''Puropose: confirms a string match using wild cards, using the LIKE operator.
'''         Website: http://analystcave.com/vba-like-operator/
'''Dependancies:
'''
'''Examples of Use:
'''
'''     *  - matches any number of characters
'''     ?  - matches any 1 character
'''     [] - matches any 1 character specified between the brackets
'''     -  - matches any range of characters e.g. [a-z] matches any non-capital 1 letter of the alphabet
'''     #  - matches any digit character
'''
'''
'''     If PatternMatch("dog", "Go*") Then  <- This will not match
'''         Debug.Print "Match"
'''     Else
'''         Debug.Print "No Match"
'''     End If
'''-----------------------------------------------------------------------------------------------------------------------------------------

        PatternMatch = False
        
        If strValueToID Like strPattern Then
        
            PatternMatch = True
            
        End If

End Function


Public Function TurnAutoFilterOn(strCellRange As String, boolTurnOff As Boolean) As Boolean

'''
'''-----------------------------------------------------------------------------------------------------------------------------------------
'''Name: Turn On Auto Filter v1.0
'''
'''Author: Priyesh Patel
'''
'''Puropose: Function will turn on auto filter using a passed variable as Reference Point, this function will also detect if autofiltering has been
'''turned on and turn it off in order to prevent error. sometimes autofiltering may be applied incorrectly either on the wrong line or if new collums
'''have been added ot will refresh and apply to all columns.
'''
'''Dependancies:None
'''
'''Examples of Use:
''' TurnAutoFilterOn("A1",True)
'''
'''-----------------------------------------------------------------------------------------------------------------------------------------

'Option to turn off AutoFilter
If boolTurnOff = True Then
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If
End If

  'check for filter, turn on if none exists
  If Not ActiveSheet.AutoFilterMode Then
    
    'Activate Filter
    ActiveSheet.Range(strCellRange).AutoFilter
    
    'Autofit Columns
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
  End If
  
  TurnAutoFilterOn = True
  
End Function

Public Function GetColumnOrRow(cell_add As String, Select_str As String) As String

'''
'''-----------------------------------------------------------------------------------------------------------------------------------------
'''Name:Get Column Or Row Address V1.0
'''Author: Priyesh Patel
'''Puropose: Pass a cell address as string or from procedure and isolate either the row or column address
'''
'''Dependancies: None
'''
'''Examples of Use:
'''GetColumnOrRow(Selection.Address, "col") <- Where Select_str can be either "col" or "row"
'''               ^                         <- Passed String must be in the Format $?$# Where ? is a Character # is a Number
'''-----------------------------------------------------------------------------------------------------------------------------------------

Select Case Select_str
    Case "col"
         GetColumnOrRow = Split(cell_add, "$")(1)
    Case "row"
         GetColumnOrRow = Split(cell_add, "$")(2)
End Select

End Function


Public Function FindColumn(strSearch As String) As String

''-----------------------------------------------------------------------------------------------------------------------------------------
'''Name: Find Column v1.0
'''Author: Priyesh Patel
'''
'''Puropose: Takes a string and finds a cell with that string, Search will be
'''Conducted on a the row of the ActiveCell.
'''
'''Dependancies: None
'''
'''Examples of Use:
'''
'''         ActiveSheet.Range("A2").Select
'''         var = FindColumn("F")
'''-----------------------------------------------------------------------------------------------------------------------------------------



'   Function will look for string in the row of the active cell
'   Returns collum address

    Dim aCell As Range

    Set aCell = ActiveSheet.Rows(ActiveCell.Row).Find(What:=strSearch, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    If Not aCell Is Nothing Then
    
        FindColumn = aCell.Column
    
    Else
        
        FindColumn = "NULL"
    
    End If

End Function


Public Function checkRangeIfExists(strRngName As String, strRngRange As String) As Integer
 
''-----------------------------------------------------------------------------------------------------------------------------------------
'''Name: Check Named Range Exists v1.0
'''Author: Priyesh Patel
'''
'''Puropose: Will take two strings one for name and one for full sheet name and range. returns 0,1,2 depending on match
''' (2)Both name and range matches (Will Do nothing)
''' (1)Name Matches only. (Will delete the range and recreate)
''' (0)No Match (Will Create the range)
'''
'''
'''Dependancies: None
'''
'''Examples of Use:
'''
'''       e.g.1  checkRangeIfExists("Hier","=Hier!$A:$D")
'''       e.g.2  checkRangeIfExists("testing", "=ePDR_Raw_Report!$A$6:$F$38")
'''-----------------------------------------------------------------------------------------------------------------------------------------

 
Dim nm As Name

'Loop through each named range in workbook
  
  For Each nm In ActiveWorkbook.Names

    If nm.Name = strRngName Then

        If nm.RefersTo = strRngRange Then

            'Debug.Print "Name & Range Match"
            checkRangeIfExists = 2
            Exit Function
            
        End If
    
     'Debug.Print "Range Match only"
     checkRangeIfExists = 1
     Exit Function
    
    End If

  Next nm

'Debug.Print "No Match"
 checkRangeIfExists = 0
End Function



'##########################################################################################################################################
'Dependant Helper Funtions
'##########################################################################################################################################

Public Function GetLastCellAddress() As String

'''
'''-----------------------------------------------------------------------------------------------------------------------------------------
'''Name:Get Last Cell Address v1.0
'''
'''Author: Priyesh Patel
'''
'''Puropose:Find the last cell address in a dataset, Dataset must
'''begin from row 1 and not have any orphen cells. i.e a cell with empty cells
'''at the row or column beginning.
'''
'''Dependancies:
'''    ~Function GetColumnOrRow()
'''
'''Examples of Use:
'''
'''   var = GetLastCellAddress() -> will return address in format ($A$1),
'''                                 use GetColumnOrRow() to split.
'''-----------------------------------------------------------------------------------------------------------------------------------------


Dim Range_address As String
Dim LastRow As String
Dim Last_Cell As String

Dim x_col As String
Dim end_col As String
Dim end_row As String

Dim c As Object

Dim LastRow_Dbl As Double
Dim Last_Cell_Dbl As Double

ActiveSheet.Range("XFD1").Select
Selection.End(xlToLeft).Select

end_col = GetColumnOrRow(Selection.Address, "col")

Last_Cell = 0

'' Last Cell is determined by checking all columns for the last cell then the highest number is used.
 For Each c In ActiveSheet.Range("A1:" & end_col & "1").Cells

        x_col = GetColumnOrRow(c.Address, "col")
        LastRow = ActiveSheet.Cells(Rows.Count, x_col).End(xlUp).Row

        LastRow_Dbl = CDbl(LastRow)
        Last_Cell_Dbl = CDbl(Last_Cell)

              If LastRow_Dbl > Last_Cell_Dbl Then

                Last_Cell = LastRow

              End If

 Next

GetLastCellAddress = "$" & end_col & "$" & Last_Cell

End Function


Public Function ConvertColToNumber(WrkSheet As String, ColumnToAffect As String)

'''
'''-----------------------------------------------------------------------------------------------------------------------------------------
'''Name:Convert Column To Number v1.0
'''
'''Author: Priyesh Patel
'''
'''Puropose:Used whenever a column of data needs to be converted into Number (The little green Triangle)
'''This will automatically find the last row
'''
'''
'''Dependancies:
'''    ~Function GetColumnOrRow()
'''    ~Function GetLastCellAddress()
'''Examples of Use:
'''
'''   Call ConvertColToNumber("Sheet7", "G2:G") -> use the format as seen "G2:G" the G2 is the First cell in the range
'''                                                the "G" will be coupled with the last used cell in the column.
'''-----------------------------------------------------------------------------------------------------------------------------------------

 Dim sht As Worksheet
 Dim DataRange As Range
 Dim LastRow As Long
 
 Set sht = ThisWorkbook.Worksheets(WrkSheet)
        
    LastRow = GetColumnOrRow(GetLastCellAddress(), "row")
    
    Set DataRange = Range(ColumnToAffect & LastRow)
    
    DataRange.Select
    
    With Selection
        Selection.NumberFormat = "0"
        .Value = .Value
    End With
   
End Function


''-----------------------------------------------------------------------------------------------------------------------------------------
'''Name:Last Row In Column v1.0
'''Author: http://www.rondebruin.nl/win/s9/win005.htm
'''Puropose: Find the Last row in a single column
'''Dependancies: None
'''Examples of Use: LastRowInColumn("A")
'''-----------------------------------------------------------------------------------------------------------------------------------------


Public Function LastRowInColumn(ColumnToUse As String) As String
'Find the last used row in a Column: column A in this example (http://www.rondebruin.nl/win/s9/win005.htm)
    Dim LastRow As Long
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, ColumnToUse).End(xlUp).Row
    End With
    LastRowInColumn = LastRow
End Function


