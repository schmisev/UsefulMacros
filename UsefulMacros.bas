Attribute VB_Name = "UsefulMacros"
Sub TransformData(source As Range, format As Range, target As Range)
    ' Setting up filter positions
    fData = 1
    fFunc = 2
    fSort = 3
    fName = 4
    
    ' Setting up main ranges
    Dim s, f, t As Range
    Set s = source
    Set f = format
    Set t = target
    
    ' Helper constants
    Dim fCol, fRow, sCol, sRow, nCol, nRow As Long
    Let fCol = f.Columns.Count
    Let fRow = f.Rows.Count
    Let sCol = s.Columns.Count
    Let sRow = s.Rows.Count
    Let nCol = fCol
    Let nRow = sRow
    
    ' Setting up helper ranges
    Dim filter As Range
    Set filter = Range(f.Cells(fName, 1), f.Cells(fRow, fCol))
    
    ' Setting up worksheets
    Dim rootWS, tmpWS As Worksheet
    Set rootWS = ThisWorkbook.ActiveSheet
    Set tmpWS = ThisWorkbook.Worksheets.Add
    
    ' Setting up temporary target
    Dim tmpTarget As Range
    Set tmpTarget = tmpWS.Range(s.Cells(1, 1).Address, Range(s.Cells(1, 1).Address).Offset(nRow - 1, nCol - 1))
    
    ' Copy data & evaluate formulas
    For i = 1 To fCol
        If IsEmpty(f.Cells(fData, i)) Then
            ' Calculate column
            With tmpTarget.Columns(i)
                .Formula = f.Cells(fFunc, i).Formula
                '.Value = .Value
            End With
        Else
            ' Find column
            For j = 1 To sCol
                If f.Cells(fData, i).Text = s.Cells(1, j).Text Then
                    ' Copy data
                    tmpTarget.Columns(i).Value = s.Columns(j).Value
                    Exit For
                End If
            Next
        End If
        ' Rename column
        tmpTarget.Cells(1, i) = f.Cells(fName, i).Text
    Next
    
    ' Sorting data
    For i = 1 To nCol
        If f.Cells(fSort, i).Text = "<" Then
            tmpTarget.Sort key1:=tmpTarget.Cells(2, i), order1:=xlAscending, header:=xlYes
        ElseIf f.Cells(fSort, i).Text = ">" Then
            tmpTarget.Sort key1:=tmpTarget.Cells(2, i), order1:=xlDescending, header:=xlYes
        End If
    Next
    
    ' Filtering data
    tmpTarget.AdvancedFilter Action:=xlFilterInPlace, criteriarange:=filter
    
    ' Copying data to actual target
    tmpTarget.SpecialCells(xlCellTypeVisible).Copy Destination:=t
    
    ' Clean up
    Application.DisplayAlerts = False
    tmpWS.Delete
    Application.DisplayAlerts = True
End Sub
Sub UITransformData()
    Call TransformData( _
    Application.InputBox("Select a table", "Get table:", Type:=8), _
    Application.InputBox("Select a format-filter range", "Get range:", Type:=8), _
    Application.InputBox("Select a target cell", "Get cell:", Type:=8))
End Sub
