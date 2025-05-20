Attribute VB_Name = "Module1"
Sub HighlightEOLCPUs()
    ' === Color Variables ===
    Dim colorEOL As Long: colorEOL = RGB(255, 0, 0)
    Dim colorDarkRed As Long: colorDarkRed = RGB(192, 0, 0)
    Dim colorServer As Long: colorServer = RGB(0, 112, 192)
    Dim colorRAMUpgrade As Long: colorRAMUpgrade = RGB(112, 48, 160)
    Dim colorVMware As Long: colorVMware = RGB(153, 101, 21)
    Dim colorGreen As Long: colorGreen = RGB(0, 176, 80)
    Dim colorYellow As Long: colorYellow = RGB(255, 255, 0)
    Dim colorAmber As Long: colorAmber = RGB(255, 191, 0)
    Dim colorLightBlue As Long: colorLightBlue = RGB(0, 176, 240)

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
    Dim agentValue As String
    Dim mainboardValue As String
    Dim manufacturerValue As String
    Dim rng As Range, r As Range

    Set reportWS = ThisWorkbook.Sheets("Table")

    ' Format as table if not already
    If reportWS.ListObjects.Count = 0 Then
        reportWS.Cells.Style = "Normal"
        lastRow = reportWS.Cells(reportWS.Rows.Count, 1).End(xlUp).Row
        lastCol = reportWS.Cells(1, reportWS.Columns.Count).End(xlToLeft).Column
        Set tblRange = reportWS.Range(reportWS.Cells(1, 1), reportWS.Cells(lastRow, lastCol))
        Set tbl = reportWS.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        tbl.Name = "ReportTable"
        tbl.TableStyle = "TableStyleMedium15"
        tbl.Range.Columns.AutoFit
        tbl.Range.Rows.AutoFit
    Else
        Set tbl = reportWS.ListObjects(1)
    End If

    ' Normalize numeric columns
    Application.ErrorCheckingOptions.NumberAsText = False

    ' Column I - Agent Memory Total
    Set rng = reportWS.Range("I2:I" & reportWS.Cells(reportWS.Rows.Count, "I").End(xlUp).Row)
    For Each r In rng
        If Not IsEmpty(r.Value) And IsNumeric(r.Value) Then r.Value = CDbl(r.Value)
    Next r
    
    ' Column N - C Drive Free Percent
    Set rng = reportWS.Range("N2:N" & reportWS.Cells(reportWS.Rows.Count, "N").End(xlUp).Row)
    For Each r In rng
        Dim rowNum As Long: rowNum = r.Row
        Dim totalSpace As Variant: totalSpace = reportWS.Cells(rowNum, "L").Value
        Dim freeSpace As Variant: freeSpace = reportWS.Cells(rowNum, "M").Value
        Dim percentFree As Variant: percentFree = r.Value

        If Not IsEmpty(totalSpace) And Not IsEmpty(freeSpace) And Not IsEmpty(percentFree) Then
            If IsNumeric(percentFree) Then
                r.Value = CDbl(percentFree)
            ElseIf InStr(percentFree, "%") > 0 Then
                r.Value = CDbl(Replace(percentFree, "%", "")) / 100
            End If
            r.NumberFormat = "0%"
        End If
    Next r

    ' Column O - Total Internal Drive
    Set rng = reportWS.Range("O2:O" & reportWS.Cells(reportWS.Rows.Count, "O").End(xlUp).Row)
    For Each r In rng
        If Not IsEmpty(r.Value) And IsNumeric(r.Value) Then r.Value = CDbl(r.Value)
    Next r

    Application.ErrorCheckingOptions.NumberAsText = True

    ' Load EOL CPU list
    downloadsPath = Environ("USERPROFILE") & "\\Downloads\\EOL_CPU_List.xlsx"
    If Dir(downloadsPath) <> "" Then
        filePath = downloadsPath
    Else
        filePath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Select EOL CPU List File")
        If filePath = "False" Then Exit Sub
    End If

    Dim lastCpuRow As Long
    lastCpuRow = reportWS.Cells(reportWS.Rows.Count, "K").End(xlUp).Row
    If lastCpuRow < 2 Then
        MsgBox "No CPU data found in column K.", vbExclamation
        Exit Sub
    End If
    Set cpuRange = reportWS.Range("K2:K" & lastCpuRow)

    Set eolWB = Workbooks.Open(filePath)
    With eolWB.Sheets(1)
        eolList = .Range("A1:A" & .Cells(.Rows.Count, "A").End(xlUp).Row).Value
    End With

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
            If Trim(reportWS.Cells(cell.Row, 8).Value) = "Microsoft Windows 11 Pro x64" Then
                reportWS.Cells(cell.Row, 2).Interior.Color = colorDarkRed
                reportWS.Cells(cell.Row, 4).Interior.Color = colorDarkRed
                reportWS.Cells(cell.Row, 8).Interior.Color = colorDarkRed
                reportWS.Cells(cell.Row, 11).Interior.Color = colorDarkRed
            End If
        Else
            agentValue = Trim(LCase(reportWS.Cells(cell.Row, 4).Value))
            mainboardValue = Trim(reportWS.Cells(cell.Row, 7).Value)
            manufacturerValue = Trim(reportWS.Cells(cell.Row, 6).Value)
            Dim isVM As Boolean
            isVM = (mainboardValue = "VMware Virtual Platform" Or mainboardValue = "Virtual Machine" Or manufacturerValue = "VMware, Inc.")

            If agentValue = "server" Then
                If Not tblRowRange Is Nothing Then
                    If tblRowRange.Interior.Color <> colorEOL Then
                        tblRowRange.Interior.Color = colorServer
                    End If
                End If
                If isVM Then
                    reportWS.Cells(cell.Row, 6).Interior.Color = colorVMware
                    reportWS.Cells(cell.Row, 7).Interior.Color = colorVMware
                End If
            ElseIf isVM Then
                If Not tblRowRange Is Nothing Then
                    If tblRowRange.Interior.Color <> colorEOL And tblRowRange.Interior.Color <> colorServer Then
                        tblRowRange.Interior.Color = colorVMware
                    End If
                End If

                        Else
                If tblRowRange.Interior.Color <> colorEOL And tblRowRange.Interior.Color <> colorServer And tblRowRange.Interior.Color <> colorVMware Then
                    Dim osValue As String
                    osValue = Trim(reportWS.Cells(cell.Row, 8).Value)

                    If osValue = "Microsoft Windows 11 Pro x64" Then
                        tblRowRange.Interior.Color = colorGreen
                    ElseIf osValue = "Microsoft Windows 10 Pro x64" Then
                        tblRowRange.Interior.Color = colorYellow
                    ElseIf osValue = "Microsoft Windows 10 Home x64" Or _
                           osValue = "Microsoft Windows 10 x64" Or _
                           osValue = "Microsoft Windows 11 Home x64" Or _
                           osValue = "Microsoft Windows 11 x64" Then
                        tblRowRange.Interior.Color = colorAmber
                    End If
                End If
            End If
        End If

        ' RAM Upgrade Check
        If Not tblRowRange Is Nothing Then
            If tblRowRange.Interior.Color <> colorEOL And tblRowRange.Interior.Color <> colorServer Then
                If IsNumeric(reportWS.Cells(cell.Row, "I").Value) Then
                    If reportWS.Cells(cell.Row, "I").Value < 16000 Then
                        reportWS.Cells(cell.Row, "I").Interior.Color = colorRAMUpgrade
                    End If
                End If
            End If
        End If
        
        ' === SSD Upgrade Check ===
        If Not tblRowRange Is Nothing Then
            If tblRowRange.Interior.Color <> colorEOL And _
               tblRowRange.Interior.Color <> colorServer And _
               tblRowRange.Interior.Color <> colorVMware Then

                Dim freePercent As Variant
                freePercent = reportWS.Cells(cell.Row, "N").Value

                If IsNumeric(freePercent) Then
                    If freePercent <= 0.25 And freePercent <= 1 Then
                        reportWS.Cells(cell.Row, "L").Interior.Color = colorLightBlue
                        reportWS.Cells(cell.Row, "M").Interior.Color = colorLightBlue
                        reportWS.Cells(cell.Row, "N").Interior.Color = colorLightBlue
                    End If
                End If
            End If
        End If


    Next cell

    ' Close the EOL workbook
    eolWB.Close SaveChanges:=False

    MsgBox "EOL CPU check complete.", vbInformation
End Sub
