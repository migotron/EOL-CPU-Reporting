Attribute VB_Name = "Module1"
Sub HighlightEOLCPUs()
    ' === Color Variables ===
    Dim colorEOL As Long: colorEOL = RGB(255, 0, 0)   ' Standard red for EOL CPUs
    Dim colorServer As Long: colorServer = RGB(0, 112, 192) ' Standard blue for Servers

    ' === Other Variables ===
    Dim reportWS As Worksheet
    Dim cpuRange As Range
    Dim eolWB As Workbook
    Dim eolList As Variant
    Dim cell As Range
    Dim i As Long
    Dim matchFound As Boolean
    Dim filePath As String
    Dim tbl As ListObject
    Dim lastRow As Long, lastCol As Long
    Dim tblRange As Range
    Dim tblRowRange As Range
    Dim downloadsPath As String
    Dim agentCell As Range
    Dim agentValue As String
    Dim rng As Range, r As Range

    ' Set the report worksheet
    Set reportWS = ThisWorkbook.Sheets("Table")

    ' Format the data as a table if not already
    If reportWS.ListObjects.Count = 0 Then
        lastRow = reportWS.Cells(reportWS.Rows.Count, 1).End(xlUp).Row
        lastCol = reportWS.Cells(1, reportWS.Columns.Count).End(xlToLeft).Column
        Set tblRange = reportWS.Range(reportWS.Cells(1, 1), reportWS.Cells(lastRow, lastCol))
        Set tbl = reportWS.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        tbl.Name = "ReportTable"
    Else
        Set tbl = reportWS.ListObjects(1)
    End If

    ' Apply "Normal" style to the table range
    tbl.Range.Style = "Normal"

    ' AutoFit all columns and rows
    tbl.Range.Columns.AutoFit
    tbl.Range.Rows.AutoFit

    ' Convert specific columns to numeric values (I, N, O)
    Application.ErrorCheckingOptions.NumberAsText = False ' Suppress "number stored as text" warning

    ' Column I - Agent Memory Total
    Set rng = reportWS.Range("I2:I" & reportWS.Cells(reportWS.Rows.Count, "I").End(xlUp).Row)
    For Each r In rng
        If IsNumeric(r.Value) Then r.Value = CDbl(r.Value)
    Next r

    ' Column N - C Drive Free Percent (convert from text like "85%" to 0.85)
    Set rng = reportWS.Range("N2:N" & reportWS.Cells(reportWS.Rows.Count, "N").End(xlUp).Row)
    For Each r In rng
        If InStr(r.Value, "%") > 0 Then
            r.Value = CDbl(Replace(r.Value, "%", "")) / 100
        ElseIf IsNumeric(r.Value) Then
            r.Value = CDbl(r.Value)
        End If
        r.NumberFormat = "0%" ' Format as percentage
    Next r


    ' Column O - Total Internal Drive
    Set rng = reportWS.Range("O2:O" & reportWS.Cells(reportWS.Rows.Count, "O").End(xlUp).Row)
    For Each r In rng
        If IsNumeric(r.Value) Then r.Value = CDbl(r.Value)
    Next r

    Application.ErrorCheckingOptions.NumberAsText = True ' Re-enable error checking

    ' Build default path to Downloads folder
    downloadsPath = Environ("USERPROFILE") & "\Downloads\EOL_CPU_List.xlsx"

    ' Check if file exists in Downloads
    If Dir(downloadsPath) <> "" Then
        filePath = downloadsPath
    Else
        filePath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Select EOL CPU List File")
        If filePath = "False" Then Exit Sub ' User cancelled
    End If

    ' Set the CPU column range (column K)
    Set cpuRange = reportWS.Range("K2:K" & reportWS.Cells(reportWS.Rows.Count, "K").End(xlUp).Row)

    ' Open the selected EOL CPU list workbook
    Set eolWB = Workbooks.Open(filePath)
    With eolWB.Sheets(1)
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

        Set tblRowRange = Intersect(cell.EntireRow, tbl.Range)

        If matchFound Then
            If Not tblRowRange Is Nothing Then
                tblRowRange.Interior.Color = colorEOL
            End If
        Else
            ' Check Agent Type in column D
            Set agentCell = reportWS.Cells(cell.Row, 4)
            agentValue = Trim(LCase(agentCell.Value))
            If agentValue = "server" Then
                If Not tblRowRange Is Nothing Then
                    If tblRowRange.Interior.Color <> colorEOL Then
                        tblRowRange.Interior.Color = colorServer
                    End If
                End If
            End If
        End If
    Next cell

    ' Close the EOL workbook
    eolWB.Close SaveChanges:=False

    MsgBox "EOL CPU check complete.", vbInformation
End Sub
