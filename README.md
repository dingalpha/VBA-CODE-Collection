# VBA-CODE-Collection

#source: http://superuser.com/questions/561923/how-can-one-split-an-excel-xlsx-file-that-contains-multiple-sheets-into-sep


## The following sub can be used to save all the sheets in a workbook as individual xlsx files.

Sub CreateNewWBS()
Dim wbThis As Workbook
Dim wbNew As Workbook
Dim ws As Worksheet
Dim strFilename As String

Set wbThis = ThisWorkbook
For Each ws In wbThis.Worksheets
    strFilename = wbThis.Path & "/" & ws.Name
    ws.Copy
    Set wbNew = ActiveWorkbook
    wbNew.SaveAs strFilename
    wbNew.Close
Next ws
End Sub


