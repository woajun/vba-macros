Attribute VB_Name = "Module1"




Sub CopyNewWorkbook()

Worksheets(1).Copy
With ActiveWorkbook
    .SaveAs Filename:=Format(Now, "yyyy-mm-dd hh_nn 미수금내역"), FileFormat:=xlOpenXMLWorkbook
End With


Call Manipulate

End Sub


Sub Manipulate()

Rows(1).EntireRow.Delete

Call DeleteColumns

Call CopyAndSort

For i = 2 To 1000
 If Cells(i, 2) > -1 Then
 Range(Cells(i, 1), Cells(i, 2)).Clear
 Else
 End If
Next

For i = 2 To 1000
 If Cells(i, 4) < 1 Then
 Range(Cells(i, 3), Cells(i, 6)).Clear
 Else
 End If
Next

Columns("A:A").ColumnWidth = 20
Columns("C:C").ColumnWidth = 20

Range("B:B").NumberFormatLocal = "#,##0_ ;[빨강]-#,##0 "

Call 정리필요

End Sub


Sub DeleteColumns()

Set myUnion = Union(Columns("A:C"), Columns("E:P"), Columns("R"), Columns("U:Y"))
myUnion.Delete

End Sub

Sub CopyAndSort()

Range("A:B").EntireColumn.Insert

Range("C:D").Copy Destination:=Range("A:B")

Dim LastRow As Long

LastRow = Range("A1").CurrentRegion.Rows.Count

Range("A2:B" & LastRow).Sort Key1:=Range("B2")

Range("C2:F" & LastRow).Sort Key1:=Range("D2"), order1:=xlDescending

End Sub
Sub 위에뭐추가()

Range("E1").Select
Selection.EntireRow.Insert
Selection.EntireRow.Insert
Range("A2").Value = "미지급금"
Range("C2").Value = "미수금"
Range("D2").Value = "=SUM(R[2]C:R[398]C)"
Range("B2").Value = "=SUM(R[2]C:R[398]C)"
Range("B2").Select
Columns("B:B").EntireColumn.AutoFit
Columns("D:D").EntireColumn.AutoFit
Range("B:B,D:D").HorizontalAlignment = xlRight
Range("A1").Select
ActiveCell.FormulaR1C1 = "=TODAY()"
Rows("1:2").Select
Selection.RowHeight = 20
Selection.Font.Size = 14

End Sub


Sub 미수계산()
misu100 = 0
misu30 = 0
misu = 0

    For i = 4 To 100
If DateDiff("d", Cells(i, 5).Value, Date) > 100 Then

        If DateDiff("d", Cells(i, 5).Value, Date) = Date Then
            Exit For
        End If
        
    Cells(i, 5).Select
    Selection.Interior.Color = 10066431
    misu100 = misu100 + Cells(i, 4).Value

ElseIf DateDiff("d", Cells(i, 5).Value, Date) > 30 Then
    Cells(i, 5).Select
    Selection.Interior.Color = 65535
    misu30 = misu30 + Cells(i, 4).Value
Else
    misu = misu + Cells(i, 4).Value
End If
Next

Range("F3").Activate

Selection.EntireRow.Insert
Selection.EntireRow.Insert
Selection.EntireRow.Insert


Range("C3").Select
ActiveCell.FormulaR1C1 = "100일 초과 미수"
Selection.Interior.Color = 10066431
Range("C4").Select
ActiveCell.FormulaR1C1 = "30일 초과 미수"
Selection.Interior.Color = 65535
Range("C5").Select
ActiveCell.FormulaR1C1 = "30일 이하 미수"
Range("D3").Value = misu100
Range("D4").Value = misu30
Range("D5").Value = misu

End Sub

Sub 정리필요()

Call 위에뭐추가

Range("D4:D100").Select

Selection.FormatConditions.AddDatabar
With Selection.FormatConditions(1)
    .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=10000000
End With
With Selection.FormatConditions(1).BarColor
    .Color = 10066431
End With

Range("A1:F450").Font.Name = "나눔바른고딕"

Range("A1:E1").Select
With Selection
    .Merge
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
End With
    
Call 미수계산
    
End Sub

