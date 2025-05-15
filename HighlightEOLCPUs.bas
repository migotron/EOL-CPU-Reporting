Attribute VB_Name = "Module1"
Sub HighlightEOLCPUs()
    Dim reportWS As Worksheet
    Dim cpuRange As Range
    Dim eolWB As Workbook
    Dim eolList As Variant
    Dim cell As Range
    Dim i As Long
    Dim matchFound As Boolean

    ' Set the report worksheet
    Set reportWS = ThisWorkbook.Sheets("Table")
    Set cpuRange = reportWS.Range("K2:K" & reportWS.Cells(reportWS.Rows.Count, "K").End(xlUp).Row)

    ' Open the external EOL CPU list workbook (update the path below)
    Set eolWB = Workbooks.Open("C:\Path\To\EOL_CPU_List.xlsx") ' <-- Update this path
    With eolWB.Sheets("Sheet1")
        eolList = .Range("A1:A" & .Cells(.Rows.Count, "A").End(xlUp).Row).Value
    End With

    ' Loop through CPU column and compare
    For Each cell In cpuRange
        matchFound = False
        For i = LBound(eolList, 1) To UBound(eolList, 1)
            If Trim(cell.Value) = Trim(eolList(i, 1)) Then
                matchFound = True
                Exit For
            End If
        Next i
        If matchFound Then
            cell.EntireRow.Interior.Color = RGB(255, 230, 230) ' Light red highlight
        End If
    Next cell

    ' Close the EOL workbook
    eolWB.Close SaveChanges:=False

    MsgBox "EOL CPU check complete.", vbInformation
End Sub

