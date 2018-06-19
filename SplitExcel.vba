Option Explicit
Sub SplitIntoSeperateFiles()

Dim OutBook As Workbook
Dim DataSheet As Worksheet, OutSheet As Worksheet
Dim FilterRange As Range
Dim UniqueNames As New Collection
Dim LastRow As Long, LastCol As Long, _
    NameCol As Long, Index As Long
Dim OutName As String

'set references and variables up-front for ease-of-use
Set DataSheet = ThisWorkbook.Worksheets("Sheet1")
NameCol = 1
LastRow = DataSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LastCol = DataSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
Set FilterRange = Range(DataSheet.Cells(1, NameCol), DataSheet.Cells(LastRow, LastCol))

'loop through the name column and store unique names in a collection
For Index = 2 To LastRow
    On Error Resume Next
        UniqueNames.Add Item:=CStr(DataSheet.Cells(Index, NameCol).Value), Key:=CStr(DataSheet.Cells(Index, NameCol).Value)
    On Error GoTo 0
Next Index

'iterate through the unique names collection, writing
'to new workbooks and saving as the group name .xls
Application.DisplayAlerts = False
For Index = 1 To UniqueNames.Count
    Set OutBook = Workbooks.Add
    Set OutSheet = OutBook.Sheets(1)
    With FilterRange
        .AutoFilter Field:=NameCol, Criteria1:=UniqueNames(Index)
        .SpecialCells(xlCellTypeVisible).Copy OutSheet.Range("A1")
    End With
    OutName = ThisWorkbook.FullName
    OutName = Left(OutName, InStrRev(OutName, "\"))
    OutName = OutName & UniqueNames(Index)
    OutBook.SaveAs Filename:=OutName, FileFormat:=xlExcel8
    OutBook.Close SaveChanges:=False
    Call ClearAllFilters(DataSheet)
Next Index
Application.DisplayAlerts = True

End Sub

'safely clear all the filters on data sheet
Sub ClearAllFilters(TargetSheet As Worksheet)
    With TargetSheet
        TargetSheet.AutoFilterMode = False
        If .FilterMode Then
            .ShowAllData
        End If
    End With
End Sub

