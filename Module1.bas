Attribute VB_Name = "Module1"

Sub CopyNewWorkbook()

Worksheets(1).Copy
With ActiveWorkbook
    .SaveAs Filename:=Format(Now, "yyyy-mm-dd hh_nn �̼��ݳ���"), FileFormat:=xlOpenXMLWorkbook
End With


Manipulate

End Sub


Sub Manipulate()

Rows(1).EntireRow.Delete

End Sub
