Sub split_fcst_tabs()
    ' Split econ forcasting file into seven csv files.
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    ' work on the full csv file received from econ risk team
    If ActiveWorkbook.ReadOnly Then
        Dim i As Integer
        Dim lastRowIndex1 As Integer
        Dim lastRowIndex2 As Integer
        Dim lst As Variant
        lst = Array("CCI", "SPX", "Income", "GDP", "Interest rate", "Unemployment Rate", "LPHPI")
        Dim ws As Worksheet
        Dim myPath As String
        Dim FileName As String
        Dim FileSnro() As String
        Dim Snro As String
        Dim myFileName As String
        myPath = ActiveWorkbook.Path
        FileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        FileSnro() = Split(ThisWorkbook.Name)
        ' keep three blanks between cecl 2020 Q2
        Snro = FileSnro(0) & FileSnro(1) & FileSnro(2)
        Application.ScreenUpdating = False
        Set wkb1 = ThisWorkbook
        Set sht1 = wkb1.Sheets("monthly_forecast")
        FileName = Left(wkb1.Name, Len(wkb1.Name) - 5)
        Set wkb2 = Workbooks.Open("Z:\CFRisk\VWang\4Ming\backup\Dictionary.xlsx")
        Set sht2 = wkb2.Sheets("Dictionary")
        sht2.Copy After:=sht1
        wkb2.Close savechanges:=False
        Set sht3 = wkb1.Sheets("Dictionary")

        sht1.Select
        sht1.Cells(1, 1).EntireRow.Delete
        lastRowIndex = sht1.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastRowIndex
            cell = sht1.Cells(i, 10)
            sht1.Cells(i, 2).Value = Application.VLookup(cell, sht3.Range("B1:D622"), 3, False)
        Next i

        For i = lastRowIndex To 2 Step -1
           If sht1.Cells(i, 2).Text = "#N/A" Then sht1.Rows(i).Delete
        Next i

        lastRowIndex2 = sht1.Cells(Rows.Count, 2).End(xlUp).Row
        sht1.Range("A:BW").Sort key1:=Range("B2"), _
        order1:=xlAscending, Header:=xlYes

        For Each Item In lst
            Set ws = Worksheets.Add(Type:=xlWorksheet)
            ws.Name = Item
            If Item = "LPHPI" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A2:BW454").Copy ws.Range("A2")
                
            End If
            If Item = "Unemployment Rate" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A455:BW506").Copy ws.Range("A2")

            End If
            If Item = "Interest rate" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A507:BW516").Copy ws.Range("A2")
            End If
            If Item = "GDP" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A517:BW567").Copy ws.Range("A2")
            End If
            If Item = "Income" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A568:BW618").Copy ws.Range("A2")
            End If
            If Item = "SPX" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A619:BW620").Copy ws.Range("A2")
            End If
            If Item = "CCI" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A621:BW622").Copy ws.Range("A2")
            End If
            myFileName = myPath & "\" & Snro & "_" & Item
            ActiveWorkbook.SaveAs FileName:=myFileName, FileFormat:=xlCSV
        Next

    Else
        MsgBox "File was not opened as read-only"
    End If
    myFileName = myPath & "\" & FileName & "_filtered"
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    ActiveWorkbook.SaveAs FileName:=myFileName, FileFormat:=51
    ActiveWorkbook.Close savechanges:=False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub


' CCI 1, 621-622
' SPX 1, 619-620
' Income 1, 568-618
' GDP 1, 517-567
' Interest rate 1, 507-516
' Unemployment Rate 1, 455-506
' LPHPI 1-454
        ' myPath = ActiveWorkbook.Path
        ' myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        ' myFullName = myPath & "\upload\" & myFileName

        ' Sheets(Sheets.Count).Cells.Select
        ' Sheets(Sheets.Count).Cells.EntireColumn.AutoFit
        ' sc1 = Sheets.Count
        ' For n = sc1 To 1 Step -1
        '     If Sheets(n).Name <> finshetname Then
        '         ActiveSheet.Name = finshetname
        '         Sheets(n).Delete
        '         Exit For
        '     End If
        ' Next n
        
        ' ActiveWorkbook.SaveAs Filename:=myFullName & "full", FileFormat:=xlCSV
        ' ActiveWorkbook.Close savechanges:=False
        ' Application.DisplayAlerts = True

Sub split_actual_tabs()
    ' Split econ forcasting file into seven csv files.
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
        Dim i As Integer
        Dim lastRowIndex1 As Integer
        Dim lastRowIndex2 As Integer
        Dim lst As Variant
        lst = Array("CCI", "SPX", "Income", "GDP", "Interest rate", "Unemployment Rate", "LPHPI")
        Dim ws As Worksheet
        Dim myPath As String
        Dim FileName As String
        Dim myFileName As String
        myPath = ActiveWorkbook.Path
        FileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        Application.ScreenUpdating = False
        Set wkb1 = ThisWorkbook
        Set sht1 = wkb1.Sheets("monthly_forecast")
        FileName = Left(wkb1.Name, Len(wkb1.Name) - 5)
        Set wkb2 = Workbooks.Open("U:\CFRisk\VWang\4Ming\backup\Dictionary.xlsx")
        Set sht2 = wkb2.Sheets("Dictionary")
        sht2.Copy After:=sht1
        wkb2.Close savechanges:=False
        Set sht3 = wkb1.Sheets("Dictionary")

        sht1.Select
        sht1.Cells(1, 1).EntireRow.Delete
        lastRowIndex = sht1.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastRowIndex
            cell = sht1.Cells(i, 10)
            sht1.Cells(i, 2).Value = Application.VLookup(cell, sht3.Range("B1:D622"), 3, False)
        Next i

        For i = lastRowIndex To 2 Step -1
           If sht1.Cells(i, 2).Text = "#N/A" Then sht1.Rows(i).Delete
        Next i

        lastRowIndex2 = sht1.Cells(Rows.Count, 2).End(xlUp).Row
        sht1.Range("A:BW").Sort key1:=Range("B2"), _
        order1:=xlAscending, Header:=xlYes

        For Each Item In lst
            Set ws = Worksheets.Add(Type:=xlWorksheet)
            ws.Name = Item
            If Item = "LPHPI" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A2:BW454").Copy ws.Range("A2")
                
            End If
            If Item = "Unemployment Rate" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A455:BW506").Copy ws.Range("A2")

            End If
            If Item = "Interest rate" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A507:BW516").Copy ws.Range("A2")
            End If
            If Item = "GDP" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A517:BW567").Copy ws.Range("A2")
            End If
            If Item = "Income" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A568:BW618").Copy ws.Range("A2")
            End If
            If Item = "SPX" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A619:BW620").Copy ws.Range("A2")
            End If
            If Item = "CCI" Then
                sht1.Rows(1).EntireRow.Copy ws.Range("A1")
                sht1.Range("A621:BW622").Copy ws.Range("A2")
            End If
            myFileName = myPath & "\" & Item
            ActiveWorkbook.SaveAs Filename:=myFileName, FileFormat:=xlCSV
        Next

    Else
        MsgBox "File was not opened as read-only"
    End If
    myFileName = myPath & "\" & FileName & "_full"
    Application.DisplayAlerts = False
    Application.EnableEvents = False    
    ActiveWorkbook.SaveAs FileName:=myFileName, FileFormat:=51
    ActiveWorkbook.Close savechanges:=False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

' CCI 1, 621-622
' SPX 1, 619-620
' Income 1, 568-618
' GDP 1, 517-567
' Interest rate 1, 507-516
' Unemployment Rate 1, 455-506
' LPHPI 1-454

        ' myPath = ActiveWorkbook.Path
        ' myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        ' myFullName = myPath & "\upload\" & myFileName

        ' Sheets(Sheets.Count).Cells.Select
        ' Sheets(Sheets.Count).Cells.EntireColumn.AutoFit
        ' sc1 = Sheets.Count
        ' For n = sc1 To 1 Step -1
        '     If Sheets(n).Name <> finshetname Then
        '         ActiveSheet.Name = finshetname
        '         Sheets(n).Delete
        '         Exit For
        '     End If
        ' Next n
        
        ' ActiveWorkbook.SaveAs Filename:=myFullName & "full", FileFormat:=xlCSV
        ' ActiveWorkbook.Close savechanges:=False
        ' Application.DisplayAlerts = True

