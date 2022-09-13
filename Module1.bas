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
