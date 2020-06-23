' Still working on ...
Sub LLR_BHCbase_3tabs()
    If ActiveWorkbook.ReadOnly Then
    ' Test to see if the Read-only attribute was assigned to the file.
    ''''''''''KEEP the last row in A column the data line!!''''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String
        Dim lst1 As Variant
        lst1 = Array("Balance", "Prov", "NonProv")

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 1
        tempswth = 6
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"

        Application.ScreenUpdating = False
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        ElseIf tempswth = 6 Then
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If

        col2copy = 5
        col2inst = 1
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For N = sc1 To 1 Step -1
            If Sheets(N).Name = finshetname Then
                Sheets(N).Delete
                Exit For
            End If
        Next N
        
        ' LLR 14A - 

        For Each Item In lst1

            datashetname = "LLR 14A - " & Item  ' make sure datafile is sheet 1  '' NEED to confirm Sheet Name!
            Sheets.Add After:=Sheets(Sheets.Count)
            finshetname = Item
            sc1 = Sheets.Count
            For N = sc1 To 1 Step -1
                If Sheets(N).Name = finshetname Then
                    Sheets(N).Delete
                    Exit For
                End If
            Next N
            Sheets(Sheets.Count).Name = finshetname
            lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
            lastColIndex = 14
        
            row2copy = lastRowIndex - col2inst
            col2trans = lastColIndex - col2copy
            recleng = col2trans * row2copy + 1
            For j = 1 To col2copy
                Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(1, j).Value
            Next j
            Sheets(finshetname).Range("G1").Value = "Adjustment Reasons"
            Sheets(finshetname).Range("G2:G" & recleng).Value = "Update by CCAR"
            Sheets(finshetname).Range("H1").Value = "Quarter"

            If Item = "Balance" Then      
                Sheets(finshetname).Range("I1").Value = "LLR Balance Override"
            Elseif Item = "Prov" Then
                Sheets(finshetname).Range("I1").Value = "LLR Build Release Override"
            Elseif Item = "NonProv" Then
                Sheets(finshetname).Range("I1").Value = "Value"
            End if

                For i = 1 To row2copy
                    For j = 1 To col2trans
                        For k = 1 To col2copy
                            Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(i + 1, k).Value
                        Next k
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + i, col2copy + j).Value
                    Next j
                Next i

            Sheets(finshetname).Columns("A:A").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Sheets(finshetname).Range("A1").Value = "Region"
            Sheets(finshetname).Range("B1").Value = "Country"
            Sheets(finshetname).Range("C1").Value = "Business"
            Sheets(finshetname).Range("D1").Value = "CCAR"
            Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
            Sheets(finshetname).Range("F1").Value = "Submitter"
            Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
            Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
            Sheets(finshetname).Range("B2:B" & recleng).Value = "US"
            
            If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
                Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
            ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
                If InStr(1, myPath, "\DM") <> 0 Then
                    If InStr(1, myFileName, "CCB ") <> 0 Then
                        Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                    ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                        Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                    End If
                Else
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
                End If
            End If

            Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
            Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
            Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
            Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
            
            Sheets(finshetname).Columns("O:O").Select
            Selection.Cut
            Sheets(finshetname).Columns("I:I").Select
            Selection.Insert Shift:=xlToRight
            
            Sheets(finshetname).Columns("L:L").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Sheets(finshetname).Range("L1").Value = "Margin Lending Indicator"
            Sheets(finshetname).Range("L2:L" & recleng).Value = "Not Applicable"
            
            Sheets(finshetname).Columns("M:M").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Sheets(finshetname).Range("M1").Value = "Currency"
            Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
            
            Sheets(finshetname).Columns("P:P").Select
            Selection.Delete
            
            Sheets(finshetname).Columns("Q:Q").Select
            Selection.Cut
            Sheets(finshetname).Columns("P:P").Select
            Selection.Insert Shift:=xlToRight
            
            Selection.NumberFormat = "0.0"

            For i = recleng To 2 Step -1
                If Sheets(finshetname).Range("O" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("O" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("L" & i).Value = "Others"  
            Next i     

            Sheets(finshetname).Cells.Select
            Sheets(finshetname).Cells.EntireColumn.AutoFit
            
            ' For Each xWs In Application.ActiveWorkbook.Worksheets
            '     If xWs.Name <> finshetname Then
            '         xWs.Delete
            '     End If
            ' Next
            ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0_"&Item, FileFormat:=51
        Next
    ActiveWorkbook.Close savechanges:=False
    Application.DisplayAlerts = True
    Application.EnableEvents = True        
    Else
        MsgBox "File is not read-only"
    End If
End Sub

Sub LLR_BHCbase_Balance()
    If ActiveWorkbook.ReadOnly Then
    ' Test to see if the Read-only attribute was assigned to the file.
    ''''''''''KEEP the last row in A column the data line!!''''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 1
        tempswth = 6
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"

        Application.ScreenUpdating = False
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        ElseIf snroswth = 6 Then
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity" 
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If

        col2copy = 5
        col2inst = 1
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For N = sc1 To 1 Step -1
            If Sheets(N).Name = finshetname Then
                Sheets(N).Delete
                Exit For
            End If
        Next N
        datashetname = "LLR 14A - Balance"  ' make sure datafile is sheet 1  '' NEED to confirm Sheet Name!
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Balance"
        sc1 = Sheets.Count
        For N = sc1 To 1 Step -1
            If Sheets(N).Name = finshetname Then
                Sheets(N).Delete
                Exit For
            End If
        Next N
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = 14
    
        row2copy = lastRowIndex - col2inst
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(1, j).Value
        Next j
        Sheets(finshetname).Range("G1").Value = "Adjustment Reasons"
        Sheets(finshetname).Range("G2:G" & recleng).Value = "Update by CCAR"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "LLR Balance Override"

            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(i + 1, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + i, col2copy + j).Value
                Next j
            Next i

        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"
        
        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("L1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("L2:L" & recleng).Value = "Not Applicable"
        
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
        
        Sheets(finshetname).Columns("P:P").Select
        Selection.Delete
        
        Sheets(finshetname).Columns("Q:Q").Select
        Selection.Cut
        Sheets(finshetname).Columns("P:P").Select
        Selection.Insert Shift:=xlToRight
        
        Selection.NumberFormat = "0.0"

        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("O" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("O" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("L" & i).Value = "Others"  
        Next i     

        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit
        
        For Each xWs In Application.ActiveWorkbook.Worksheets
            If xWs.Name <> finshetname Then
                xWs.Delete
            End If
        Next
                
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0_Balance", fileformat:=51
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
        
    Else
        MsgBox "File is not read-only"
    End If
End Sub

Sub LLR_BHCbase_Prov()
    If ActiveWorkbook.ReadOnly Then
    ' Test to see if the Read-only attribute was assigned to the file.
    ''''''''''KEEP the last row in A column the data line!!''''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String        
        
        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 1
        tempswth = 6
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"

        Application.ScreenUpdating = False
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        ElseIf snroswth = 6 Then
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If

        col2copy = 5
        col2inst = 1
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For N = sc1 To 1 Step -1
            If Sheets(N).Name = finshetname Then
                Sheets(N).Delete
                Exit For
            End If
        Next N
        datashetname = "LLR 14A - Prov"  ' make sure datafile is sheet 1  '' NEED to confirm Sheet Name!
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Prov"
        sc1 = Sheets.Count
        For N = sc1 To 1 Step -1
            If Sheets(N).Name = finshetname Then
                Sheets(N).Delete
                Exit For
            End If
        Next N
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = 14
    
        row2copy = lastRowIndex - col2inst
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(1, j).Value
        Next j
        Sheets(finshetname).Range("G1").Value = "Adjustment Reasons"
        Sheets(finshetname).Range("G2:G" & recleng).Value = "Update by CCAR"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "LLR Build Release Override"

            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(i + 1, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + i, col2copy + j).Value
                Next j
            Next i

        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"
        
        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("L1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("L2:L" & recleng).Value = "Not Applicable"
        
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
        
        Sheets(finshetname).Columns("P:P").Select
        Selection.Delete
        
        Sheets(finshetname).Columns("Q:Q").Select
        Selection.Cut
        Sheets(finshetname).Columns("P:P").Select
        Selection.Insert Shift:=xlToRight
        
        Selection.NumberFormat = "0.0"
        ' For i = recleng To 2 Step -1
        '     If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"
        ' Next i
        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("O" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("O" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("L" & i).Value = "Others"  
        Next i     

        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit
        
        For Each xWs In Application.ActiveWorkbook.Worksheets
            If xWs.Name <> finshetname Then
                xWs.Delete
            End If
        Next
                
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0_Prov", fileformat:=51
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
        
    Else
        MsgBox "File is not read-only"
    End If
End Sub

Sub LLR_BHCbase_NonProv()
    If ActiveWorkbook.ReadOnly Then
    ' Test to see if the Read-only attribute was assigned to the file.
    ''''''''''KEEP the last row in A column the data line!!''''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String        
        
        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 1
        tempswth = 6
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"

        Application.ScreenUpdating = False
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        ElseIf snroswth = 6 Then
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"        
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If

        col2copy = 5
        col2inst = 1
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For N = sc1 To 1 Step -1
            If Sheets(N).Name = finshetname Then
                Sheets(N).Delete
                Exit For
            End If
        Next N
        datashetname = "LLR 14A - NonProv"  ' make sure datafile is sheet 1  '' NEED to confirm Sheet Name!
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "NonProv"
        sc1 = Sheets.Count
        For N = sc1 To 1 Step -1
            If Sheets(N).Name = finshetname Then
                Sheets(N).Delete
                Exit For
            End If
        Next N
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = 14
    
        row2copy = lastRowIndex - col2inst
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(1, j).Value
        Next j
        Sheets(finshetname).Range("G1").Value = "Adjustment Reasons"
        Sheets(finshetname).Range("G2:G" & recleng).Value = "Update by CCAR"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "Value"

            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(i + 1, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + i, col2copy + j).Value
                Next j
            Next i

        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"
        
        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If


        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("L1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("L2:L" & recleng).Value = "Not Applicable"
        
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
        
        Sheets(finshetname).Columns("P:P").Select
        Selection.Delete
        
        Sheets(finshetname).Columns("Q:Q").Select
        Selection.Cut
        Sheets(finshetname).Columns("P:P").Select
        Selection.Insert Shift:=xlToRight
        
        Selection.NumberFormat = "0.0"
        ' For i = recleng To 2 Step -1
        '     If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"
        ' Next i         

        Sheets(finshetname).Columns("P:P").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("P1").Value = "Field Name"
        Sheets(finshetname).Range("P2:P" & recleng).Value = "NON_PROV_DUE_TO_SALES"

        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("O" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("O" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("L" & i).Value = "Others"  
        Next i     

        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit
        
        For Each xWs In Application.ActiveWorkbook.Worksheets
            If xWs.Name <> finshetname Then
                xWs.Delete
            End If
        Next
                
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0_NonProv", fileformat:=51
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
        
    Else
        MsgBox "File is not read-only"
    End If
End Sub

Sub RBLP4all()
    ' Test to see if the Read-only attribute was assigned to the file.
    If ActiveWorkbook.ReadOnly Then
    ''''''''''''''''''''''''KEEP the last row in A column the data line!!''''''''''''''''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate As String
        Dim CCARmonth As String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        ' snroswth = 1
        tempswth = 1
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"

        Application.ScreenUpdating = False

        If InStr(1, myFileName, "BHC") <> 0 Then
            If InStr(1, myFileName, "Actual") <> 0 Or InStr(1, myFileName, "Actuals") <> 0 Then
                snrotyp = "Actuals"
            ElseIf InStr(1, myFileName, "Base") <> 0 Or InStr(1, myFileName, "base") <> 0 Then
                snrotyp = "BHC Baseline"
            ElseIf InStr(1, myFileName, "BHC DS") <> 0 Or InStr(1, myFileName, "BHC_DS") <> 0 Then
                snrotyp = "BHC Stress"
            End If

        ElseIf InStr(1, myFileName, "FRB base") Or InStr(1, myFileName, "FRB Base") Or InStr(1, myFileName, "FRB_Base") <> 0 Then
            snrotyp = "Supervisory Baseline"
        ElseIf InStr(1, myFileName, "FRB Sev") Or InStr(1, myFileName, "FRB SSA") Or InStr(1, myFileName, "FRB_SSA") <> 0 Then
            snrotyp = "Supervisory Severely Adverse"
        Else
            snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If
        
    ''''''''''' Scenarios ''''''''''''''''
    '    snrotyp = "Actuals"
    '    snrotyp = "BHC Baseline"
    '    snrotyp = "BHC Stress"
    '    snrotyp = "Supervisory Baseline"
    
    '    snrotyp = "Supervisory Severely Adverse"
    
    '''''''''''' Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        Sheets(finshetname).Range("G1").Value = "Forecast Type"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "Value"
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value
                Next j
            Next i
        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"
''''''''''''''''''''''''''''''''''''''''''''''''
' Need to change other templates as well later 20200320
''''''''''''''''''''''''''''''''''''''''''''''''
        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            If InStr(1, myFileName, "_CMI_") <> 0 Then
                Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
            ElseIf InStr(1, myFileName, "_USRB_") <> 0 Then
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        Sheets(finshetname).Columns("N:N").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("P:P").Select
        Selection.Cut
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("L:L").Select
        Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
        Sheets(finshetname).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("Q2:Q" & recleng).Value = "Not Applicable"
        Sheets(finshetname).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("R1").Value = "Template Type"
        Sheets(finshetname).Range("R2:R" & recleng).Value = temptyp
        
    '    finshetname = Sheets(finshetname).Range("R2").Value & "_" & Sheets(finshetname).Range("E2").Value
    '    sc1 = Sheets.Count
    '    For n = sc1 To 1 Step -1
    '        If Sheets(n).Name = finshetname Then
    '            Sheets(n).Delete
    '            Exit For
    '        End If
    '    Next n
    '    Sheets(Sheets.Count).Name = finshetname
    
    '    For i = recleng To 2 Step -1
    '        If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
    '    Next i

        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"
            If (Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = ".") And Sheets(finshetname).Range("K" & i).Value <> "GCL_MODEL_OUTPUT" And Sheets(finshetname).Range("K" & i).Value <> "RECOVERY_AMOUNT" Then Sheets(finshetname).Rows(i).Delete
        Next i
    
        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit
        
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
        
    Else
        MsgBox "File is not read-only"
    End If
End Sub

Sub RWA4all()
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        ' snroswth = 1
        tempswth = 2
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"

        Application.ScreenUpdating = False    

        If InStr(1, myFileName, "BHC") <> 0 Then    
            If InStr(1, myFileName, "Actual") <> 0 Or InStr(1, myFileName, "Actuals") <> 0 Then
                snrotyp = "Actuals"
            ElseIf InStr(1, myFileName, "Base") <> 0 Or InStr(1, myFileName, "base") <> 0 Then       
                snrotyp = "BHC Baseline"
            ElseIf InStr(1, myFileName, "BHC DS") <> 0 Or InStr(1, myFileName, "BHC_DS") <> 0 Then
                snrotyp = "BHC Stress"
            End If

        ElseIf InStr(1, myFileName, "FRB base") Or InStr(1, myFileName, "FRB Base") Or InStr(1, myFileName, "FRB_Base") <> 0 Then
            snrotyp = "Supervisory Baseline"
        ElseIf InStr(1, myFileName, "FRB Sev") Or InStr(1, myFileName, "FRB SSA") Or InStr(1, myFileName, "FRB_SSA") <> 0 Then
            snrotyp = "Supervisory Severely Adverse"   
        Else
            snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If

        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
    
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        
        Sheets(Sheets.Count).Range("G1").Value = "Forecast Type"
        Sheets(Sheets.Count).Range("H1").Value = "Quarter"
        Sheets(Sheets.Count).Range("I1").Value = "Value"
        
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value   'copy fixed fields
                    Next k
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value 'copy fcst indicator
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value 'copy quarter indicator
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value 'copy value
                Next j
            Next i
        
        Sheets(Sheets.Count).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        Sheets(Sheets.Count).Range("A1").Value = "Region"
        Sheets(Sheets.Count).Range("B1").Value = "Country"
        ' Sheets(Sheets.Count).Range("C1").Value = "Business"
        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            If InStr(1, myFileName, "_CMI_") <> 0 Then
                Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
            ElseIf InStr(1, myFileName, "_USRB_") <> 0 Then
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If
        Sheets(Sheets.Count).Range("D1").Value = "CCAR"
        Sheets(Sheets.Count).Range("E1").Value = "Stress Test Scenario"
        Sheets(Sheets.Count).Range("F1").Value = "Submitter"
        Sheets(Sheets.Count).Range("G1").Value = "Forecast Start Date"
       
        Sheets(Sheets.Count).Range("A2:A" & recleng).Value = "NAM"
        Sheets(Sheets.Count).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(Sheets.Count).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(Sheets.Count).Range("E2:E" & recleng).Value = snrotyp
        Sheets(Sheets.Count).Range("F2:F" & recleng).Value = " XL94292"
        Sheets(Sheets.Count).Range("G2:G" & recleng).Value = FSdate
        
        Sheets(Sheets.Count).Columns("N:N").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(Sheets.Count).Columns("O:O").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("P:P").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("L:L").Select
        'Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        
        Sheets(Sheets.Count).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("M1").Value = "Currency"
        Sheets(Sheets.Count).Range("M2:M" & recleng).Value = "USD"
        
        Sheets(Sheets.Count).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(Sheets.Count).Range("Q2:Q" & recleng).Value = "Not Applicable"
        
        Sheets(Sheets.Count).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("R1").Value = "Template Type"
        Sheets(Sheets.Count).Range("R2:R" & recleng).Value = temptyp
        
        Sheets(Sheets.Count).Columns("S:S").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("K:K").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("L:L").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("O:O").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("Q:Q").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("T:T").Select
        Selection.Insert Shift:=xlToRight
        
        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
            If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
        Next i
    
        Sheets(Sheets.Count).Cells.Select
        Sheets(Sheets.Count).Cells.EntireColumn.AutoFit
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
    Else
        MsgBox "File was not opened as read-only"
    End If

End Sub

Sub Overlay4all()
        ' Check to see if the active workbook was
        ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
        ''''''''''''''''''''''''KEEP the last row in A column the data line!!''''''''''''''''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        ' snroswth = 1
        tempswth = 3
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"
        Application.ScreenUpdating = False    

        If InStr(1, myFileName, "BHC") <> 0 Then    
            If InStr(1, myFileName, "Actual") <> 0 Or InStr(1, myFileName, "Actuals") <> 0 Then
                snrotyp = "Actuals"
            ElseIf InStr(1, myFileName, "Base") <> 0 Or InStr(1, myFileName, "base") <> 0 Then       
                snrotyp = "BHC Baseline"
            ElseIf InStr(1, myFileName, "BHC DS") <> 0 Or InStr(1, myFileName, "BHC_DS") <> 0 Then
                snrotyp = "BHC Stress"
            End If

        ElseIf InStr(1, myFileName, "FRB base") Or InStr(1, myFileName, "FRB Base") Or InStr(1, myFileName, "FRB_Base") <> 0 Then
            snrotyp = "Supervisory Baseline"
        ElseIf InStr(1, myFileName, "FRB Sev") Or InStr(1, myFileName, "FRB SSA") Or InStr(1, myFileName, "FRB_SSA") <> 0 Then
            snrotyp = "Supervisory Severely Adverse"   
        Else
            snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If
        ''''''''''' Five Scenarios ''''''''''''''''
        '    snrotyp = "BHC Baseline"
        '    snrotyp = "BHC Stress"
        '    snrotyp = "Supervisory Baseline"
        
        '    snrotyp = "Supervisory Severely Adverse"
        
        '''''''''''' Four Templates '''''''''''''''
        '    temptyp = "Stress Loss"          '' RBLP (14A)
        '    temptyp = "RWA"
        '    temptyp = "Overlay"
        '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
       Application.ScreenUpdating = False     
        finshetname = "Tempname"
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
    
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        
        Sheets(finshetname).Range("G1").Value = "Forecast Type"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "Value"
        
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value
                Next j
            Next i
        
        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        ' Sheets(finshetname).Range("C1").Value = "Business"
        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            If InStr(1, myFileName, "_CMI_") <> 0 Then
                Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
            ElseIf InStr(1, myFileName, "_USRB_") <> 0 Then
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        Sheets(finshetname).Columns("N:N").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("P:P").Select
        Selection.Cut
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("L:L").Select
        '    Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
        Sheets(finshetname).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("Q2:Q" & recleng).Value = "Not Applicable"
        Sheets(finshetname).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("R1").Value = "Template Type"
        Sheets(finshetname).Range("R2:R" & recleng).Value = temptyp

        recleng = Sheets(finshetname).Cells(Rows.Count, 1).End(xlUp).Row    
        For i = recleng To 2 Step -1
            If IsEmpty(Sheets(finshetname).Range("L" & i).Value) = True Then Sheets(finshetname).Range("L" & i).Value = 0
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
            If (Sheets(finshetname).Range("H" & i).Value = "FIRST_HELOAN" Or Sheets(finshetname).Range("H" & i).Value = "HELOC") And Sheets(finshetname).Range("P" & i).Value = "Non CBNA" Then Sheets(finshetname).Rows(i).Delete
        Next i
    
        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit

        Dim r As Integer
        For r = Sheets(finshetname).UsedRange.Rows.Count To 1 Step -1
            If Cells(r, "K") = "MODEL_OUTPUT" Then
                Sheets(finshetname).Rows(r).EntireRow.Delete
            End If
        Next

        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
    Else
        MsgBox "File was not opened as read-only"
    End If
End Sub



Sub RBLP_Actuals()
    'RBLP Actuals Templates
    'RBLP Actuals 01/24/2019
    'Check snroswth/tempswth
    'Check col2trans /row2copy/recleng
    'KEEP the last row in column A the data row!!
    'KEEP the last column in row 3 the data column!!
    'Make sure data sheet is sheet 1!
    ' Test to see if the Read-only attribute was assigned to the file.
    If ActiveWorkbook.ReadOnly Then
    ''''''''''''''''''''''''KEEP the last row in A column the data line!!''''''''''''''''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate As String
        Dim CCARmonth As String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 0
        tempswth = 1
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"

        If InStr(1, myFileName, "BHC base") <> 0 Or InStr(1, myFileName, "BHC Base") <> 0 Or InStr(1, myFileName, "BHC_Base") Then        
            snrotyp = "BHC Baseline"
        ElseIf InStr(1, myFileName, "BHC DS") <> 0 Or InStr(1, myFileName, "BHC_DS") <> 0 Then
            snrotyp = "BHC Stress"
        ElseIf InStr(1, myFileName, "FRB base") Or InStr(1, myFileName, "FRB Base") Or InStr(1, myFileName, "FRB_Base") <> 0 Then
            snrotyp = "Supervisory Baseline"
        ElseIf InStr(1, myFileName, "FRB Sev") Or InStr(1, myFileName, "FRB SSA") Or InStr(1, myFileName, "FRB_SSA") <> 0 Then
            snrotyp = "Supervisory Severely Adverse"   
        ElseIf InStr(1, myFileName, "Actuals") <> 0 Then
            snrotyp = "Actuals"
        Else
            snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If
        
    ''''''''''' Scenarios ''''''''''''''''
    '    snrotyp = "Actuals"
    '    snrotyp = "BHC Baseline"
    '    snrotyp = "BHC Stress"
    '    snrotyp = "Supervisory Baseline"
    
    '    snrotyp = "Supervisory Severely Adverse"
    
    '''''''''''' Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        Sheets(finshetname).Range("G1").Value = "Forecast Type"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "Value"
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value
                Next j
            Next i
        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        Sheets(finshetname).Columns("N:N").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("P:P").Select
        Selection.Cut
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("L:L").Select
        Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
        Sheets(finshetname).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("Q2:Q" & recleng).Value = "Not Applicable"
        Sheets(finshetname).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("R1").Value = "Template Type"
        Sheets(finshetname).Range("R2:R" & recleng).Value = temptyp
        
    '    finshetname = Sheets(finshetname).Range("R2").Value & "_" & Sheets(finshetname).Range("E2").Value
    '    sc1 = Sheets.Count
    '    For n = sc1 To 1 Step -1
    '        If Sheets(n).Name = finshetname Then
    '            Sheets(n).Delete
    '            Exit For
    '        End If
    '    Next n
    '    Sheets(Sheets.Count).Name = finshetname
    
    '    For i = recleng To 2 Step -1
    '        If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
    '    Next i

        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"
            If (Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = ".") And Sheets(finshetname).Range("K" & i).Value <> "GCL_MODEL_OUTPUT" And Sheets(finshetname).Range("K" & i).Value <> "RECOVERY_AMOUNT" Then Sheets(finshetname).Rows(i).Delete
        Next i

        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit
        
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
        
    Else
        MsgBox "File is not read-only"
    End If

End Sub

Sub RBLP_BBS()
    ' Test to see if the Read-only attribute was assigned to the file.
    If ActiveWorkbook.ReadOnly Then
    ''''''''''''''''''''''''KEEP the last row in A column the data line!!''''''''''''''''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 1
        tempswth = 1
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"

        Application.ScreenUpdating = False    
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        Else
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If
        
    ''''''''''' Scenarios ''''''''''''''''
    '    snrotyp = "Actuals"
    '    snrotyp = "BHC Baseline"
    '    snrotyp = "BHC Stress"
    '    snrotyp = "Supervisory Baseline"
    
    '    snrotyp = "Supervisory Severely Adverse"
    
    '''''''''''' Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False 
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        Sheets(finshetname).Range("G1").Value = "Forecast Type"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "Value"
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value
                Next j
            Next i
        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        Sheets(finshetname).Columns("N:N").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("P:P").Select
        Selection.Cut
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("L:L").Select
        Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
        Sheets(finshetname).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("Q2:Q" & recleng).Value = "Not Applicable"
        Sheets(finshetname).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("R1").Value = "Template Type"
        Sheets(finshetname).Range("R2:R" & recleng).Value = temptyp
        
    '    finshetname = Sheets(finshetname).Range("R2").Value & "_" & Sheets(finshetname).Range("E2").Value
    '    sc1 = Sheets.Count
    '    For n = sc1 To 1 Step -1
    '        If Sheets(n).Name = finshetname Then
    '            Sheets(n).Delete
    '            Exit For
    '        End If
    '    Next n
    '    Sheets(Sheets.Count).Name = finshetname
    
    '    For i = recleng To 2 Step -1
    '        If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
    '    Next i

        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
            If (Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = ".") And Sheets(finshetname).Range("K" & i).Value <> "GCL_MODEL_OUTPUT" And Sheets(finshetname).Range("K" & i).Value <> "RECOVERY_AMOUNT" Then Sheets(finshetname).Rows(i).Delete
        Next i
    
        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit
        
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
        
    Else
        MsgBox "File is not read-only"
    End If

End Sub

Sub RBLP_BDS()
    ' Test to see if the Read-only attribute was assigned to the file.
    If ActiveWorkbook.ReadOnly Then
    ''''''''''''''''''''''''KEEP the last row in A column the data line!!''''''''''''''''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 2
        tempswth = 1	
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"
        Application.ScreenUpdating = False    
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        Else
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If
    ''''''''''' Five Scenarios ''''''''''''''''
    '    snrotyp = "BHC Baseline"
    '    snrotyp = "BHC Stress"
    '    snrotyp = "Supervisory Baseline"
    
    '    snrotyp = "Supervisory Severely Adverse"
    '''''''''''' Four Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        Sheets(finshetname).Range("G1").Value = "Forecast Type"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "Value"
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value
                Next j
            Next i

        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        Sheets(finshetname).Columns("N:N").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("P:P").Select
        Selection.Cut
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("L:L").Select
        Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
        Sheets(finshetname).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("Q2:Q" & recleng).Value = "Not Applicable"
        Sheets(finshetname).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("R1").Value = "Template Type"
        Sheets(finshetname).Range("R2:R" & recleng).Value = temptyp

    '    finshetname = Sheets(finshetname).Range("R2").Value & "_" & Sheets(finshetname).Range("E2").Value
    '    sc1 = Sheets.Count
    '    For n = sc1 To 1 Step -1
    '        If Sheets(n).Name = finshetname Then
    '            Sheets(n).Delete
    '            Exit For
    '        End If
    '    Next n
    '    Sheets(Sheets.Count).Name = finshetname
    
    '    For i = recleng To 2 Step -1
    '        If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
    '    Next i

        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
            If (Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = ".") And Sheets(finshetname).Range("K" & i).Value <> "GCL_MODEL_OUTPUT" And Sheets(finshetname).Range("K" & i).Value <> "RECOVERY_AMOUNT" Then Sheets(finshetname).Rows(i).Delete
        Next i

        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit
        
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
        
    Else
        MsgBox "File is not read-only"
    End If

End Sub
 
Sub RBLP_SBS()
    ' Test to see if the Read-only attribute was assigned to the file.
    If ActiveWorkbook.ReadOnly Then
    ''''''''''KEEP the last row in A column the data line!!'''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 3
        tempswth = 1
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"
        Application.ScreenUpdating = False             
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        Else
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If        
    ''''''''''' Five Scenarios ''''''''''''''''
    '    snrotyp = "BHC Baseline"
    '    snrotyp = "BHC Stress"
    '    snrotyp = "Supervisory Baseline"
    
    '    snrotyp = "Supervisory Severely Adverse"
    '''''''''''' Four Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        Sheets(finshetname).Range("G1").Value = "Forecast Type"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "Value"
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value
                Next j
            Next i

        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        Sheets(finshetname).Columns("N:N").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("P:P").Select
        Selection.Cut
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("L:L").Select
        Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
        Sheets(finshetname).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("Q2:Q" & recleng).Value = "Not Applicable"
        Sheets(finshetname).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("R1").Value = "Template Type"
        Sheets(finshetname).Range("R2:R" & recleng).Value = temptyp
        
    '    finshetname = Sheets(finshetname).Range("R2").Value & "_" & Sheets(finshetname).Range("E2").Value
    '    sc1 = Sheets.Count
    '    For n = sc1 To 1 Step -1
    '        If Sheets(n).Name = finshetname Then
    '            Sheets(n).Delete
    '            Exit For
    '        End If
    '    Next n
    '    Sheets(Sheets.Count).Name = finshetname
    
    '    For i = recleng To 2 Step -1
    '        If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
    '    Next i

        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"              
            If (Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = ".") And Sheets(finshetname).Range("K" & i).Value <> "GCL_MODEL_OUTPUT" And Sheets(finshetname).Range("K" & i).Value <> "RECOVERY_AMOUNT" Then Sheets(finshetname).Rows(i).Delete
        Next i
    
        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit
        
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
        
    Else
        MsgBox "File is not read-only"
    End If

End Sub

Sub RBLP_SSA()
    ' Test to see if the Read-only attribute was assigned to the file.
    If ActiveWorkbook.ReadOnly Then
    ''''''''''KEEP the last row in A column the data line!!'''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 4
        tempswth = 1
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"
        Application.ScreenUpdating = False             
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf tempswth = 6 Then
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 5 Then
        temptyp = "Sensitivity"        
        Else
        temptyp = "======== LLR ========="
        End If
    ''''''''''' Five Scenarios ''''''''''''''''
    '    snrotyp = "BHC Baseline"
    '    snrotyp = "BHC Stress"
    '    snrotyp = "Supervisory Baseline"
    
    '    snrotyp = "Supervisory Severely Adverse"
    '''''''''''' Four Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        Sheets(finshetname).Range("G1").Value = "Forecast Type"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "Value"
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value
                Next j
            Next i

        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        Sheets(finshetname).Columns("N:N").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("P:P").Select
        Selection.Cut
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("L:L").Select
        Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
        Sheets(finshetname).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("Q2:Q" & recleng).Value = "Not Applicable"
        Sheets(finshetname).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("R1").Value = "Template Type"
        Sheets(finshetname).Range("R2:R" & recleng).Value = temptyp
        
    '    finshetname = Sheets(finshetname).Range("R2").Value & "_" & Sheets(finshetname).Range("E2").Value
    '    sc1 = Sheets.Count
    '    For n = sc1 To 1 Step -1
    '        If Sheets(n).Name = finshetname Then
    '            Sheets(n).Delete
    '            Exit For
    '        End If
    '    Next n
    '    Sheets(Sheets.Count).Name = finshetname
    
    '    For i = recleng To 2 Step -1
    '        If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
    '    Next i

        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"              
            If (Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = ".") And Sheets(finshetname).Range("K" & i).Value <> "GCL_MODEL_OUTPUT" And Sheets(finshetname).Range("K" & i).Value <> "RECOVERY_AMOUNT" Then Sheets(finshetname).Rows(i).Delete
        Next i
    
        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit
        
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
        
    Else
        MsgBox "File is not read-only"
    End If

End Sub

Sub RWA_Actuals()
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 0
        tempswth = 2
        CCARmonth= "CCAR 2020-01"
        FSdate="Q1 2020"
        Application.ScreenUpdating = False    
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        Else
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If
    ''''''''''' Five Scenarios ''''''''''''''''
    '    snrotyp = "BHC Baseline"
    '    snrotyp = "BHC Stress"
    '    snrotyp = "Supervisory Baseline"
    
    '    snrotyp = "Supervisory Severely Adverse"
    '''''''''''' Four Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
    
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        
        Sheets(Sheets.Count).Range("G1").Value = "Forecast Type"
        Sheets(Sheets.Count).Range("H1").Value = "Quarter"
        Sheets(Sheets.Count).Range("I1").Value = "Value"
    
        For i = 1 To row2copy
            For j = 1 To col2trans
                For k = 1 To col2copy
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value   'copy fixed fields
                Next k
                Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value 'copy fcst indicator
                Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value 'copy quarter indicator
                Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value 'copy value
            Next j
        Next i
        
        Sheets(Sheets.Count).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        Sheets(Sheets.Count).Range("A1").Value = "Region"
        Sheets(Sheets.Count).Range("B1").Value = "Country"
        Sheets(Sheets.Count).Range("C1").Value = "Business"
        Sheets(Sheets.Count).Range("D1").Value = "CCAR"
        Sheets(Sheets.Count).Range("E1").Value = "Stress Test Scenario"
        Sheets(Sheets.Count).Range("F1").Value = "Submitter"
        Sheets(Sheets.Count).Range("G1").Value = "Forecast Start Date"
       
        Sheets(Sheets.Count).Range("A2:A" & recleng).Value = "NAM"
        Sheets(Sheets.Count).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(Sheets.Count).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(Sheets.Count).Range("E2:E" & recleng).Value = snrotyp
        Sheets(Sheets.Count).Range("F2:F" & recleng).Value = " XL94292"
        Sheets(Sheets.Count).Range("G2:G" & recleng).Value = FSdate
        
        Sheets(Sheets.Count).Columns("N:N").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(Sheets.Count).Columns("O:O").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("P:P").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("L:L").Select
        'Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        
        Sheets(Sheets.Count).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("M1").Value = "Currency"
        Sheets(Sheets.Count).Range("M2:M" & recleng).Value = "USD"
        
        Sheets(Sheets.Count).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(Sheets.Count).Range("Q2:Q" & recleng).Value = "Not Applicable"
        
        Sheets(Sheets.Count).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("R1").Value = "Template Type"
        Sheets(Sheets.Count).Range("R2:R" & recleng).Value = temptyp
        
        Sheets(Sheets.Count).Columns("S:S").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("K:K").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("L:L").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("O:O").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("Q:Q").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("T:T").Select
        Selection.Insert Shift:=xlToRight
                
        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others" 
            If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
        Next i
    
        Sheets(Sheets.Count).Cells.Select
        Sheets(Sheets.Count).Cells.EntireColumn.AutoFit
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
    Else
        MsgBox "File was not opened as read-only"
    End If
End Sub

Sub RWA_BBS()
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 1
        tempswth = 2
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"

        Application.ScreenUpdating = False    
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        Else
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If

        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
    
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        
        Sheets(Sheets.Count).Range("G1").Value = "Forecast Type"
        Sheets(Sheets.Count).Range("H1").Value = "Quarter"
        Sheets(Sheets.Count).Range("I1").Value = "Value"
        
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value   'copy fixed fields
                    Next k
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value 'copy fcst indicator
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value 'copy quarter indicator
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value 'copy value
                Next j
            Next i
        
        Sheets(Sheets.Count).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        Sheets(Sheets.Count).Range("A1").Value = "Region"
        Sheets(Sheets.Count).Range("B1").Value = "Country"
        Sheets(Sheets.Count).Range("C1").Value = "Business"
        Sheets(Sheets.Count).Range("D1").Value = "CCAR"
        Sheets(Sheets.Count).Range("E1").Value = "Stress Test Scenario"
        Sheets(Sheets.Count).Range("F1").Value = "Submitter"
        Sheets(Sheets.Count).Range("G1").Value = "Forecast Start Date"
       
        Sheets(Sheets.Count).Range("A2:A" & recleng).Value = "NAM"
        Sheets(Sheets.Count).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(Sheets.Count).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(Sheets.Count).Range("E2:E" & recleng).Value = snrotyp
        Sheets(Sheets.Count).Range("F2:F" & recleng).Value = " XL94292"
        Sheets(Sheets.Count).Range("G2:G" & recleng).Value = FSdate
        
        Sheets(Sheets.Count).Columns("N:N").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(Sheets.Count).Columns("O:O").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("P:P").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("L:L").Select
        'Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        
        Sheets(Sheets.Count).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("M1").Value = "Currency"
        Sheets(Sheets.Count).Range("M2:M" & recleng).Value = "USD"
        
        Sheets(Sheets.Count).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(Sheets.Count).Range("Q2:Q" & recleng).Value = "Not Applicable"
        
        Sheets(Sheets.Count).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("R1").Value = "Template Type"
        Sheets(Sheets.Count).Range("R2:R" & recleng).Value = temptyp
        
        Sheets(Sheets.Count).Columns("S:S").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("K:K").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("L:L").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("O:O").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("Q:Q").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("T:T").Select
        Selection.Insert Shift:=xlToRight
        
        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
            If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
        Next i
    
        Sheets(Sheets.Count).Cells.Select
        Sheets(Sheets.Count).Cells.EntireColumn.AutoFit
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
    Else
        MsgBox "File was not opened as read-only"
    End If

End Sub

Sub RWA_BDS()
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName
        snroswth = 2
        tempswth = 2
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"

        Application.ScreenUpdating = False    
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        Else
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If
    ''''''''''' Five Scenarios ''''''''''''''''
    '    snrotyp = "BHC Baseline"
    '    snrotyp = "BHC Stress"
    '    snrotyp = "Supervisory Baseline"
    
    '    snrotyp = "Supervisory Severely Adverse"
    
    '''''''''''' Four Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
    
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        
        Sheets(Sheets.Count).Range("G1").Value = "Forecast Type"
        Sheets(Sheets.Count).Range("H1").Value = "Quarter"
        Sheets(Sheets.Count).Range("I1").Value = "Value"
        
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value   'copy fixed fields
                    Next k
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value 'copy fcst indicator
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value 'copy quarter indicator
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value 'copy value
                Next j
            Next i
        
        Sheets(Sheets.Count).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        Sheets(Sheets.Count).Range("A1").Value = "Region"
        Sheets(Sheets.Count).Range("B1").Value = "Country"
        Sheets(Sheets.Count).Range("C1").Value = "Business"
        Sheets(Sheets.Count).Range("D1").Value = "CCAR"
        Sheets(Sheets.Count).Range("E1").Value = "Stress Test Scenario"
        Sheets(Sheets.Count).Range("F1").Value = "Submitter"
        Sheets(Sheets.Count).Range("G1").Value = "Forecast Start Date"
       
        Sheets(Sheets.Count).Range("A2:A" & recleng).Value = "NAM"
        Sheets(Sheets.Count).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(Sheets.Count).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(Sheets.Count).Range("E2:E" & recleng).Value = snrotyp
        Sheets(Sheets.Count).Range("F2:F" & recleng).Value = " XL94292"
        Sheets(Sheets.Count).Range("G2:G" & recleng).Value = FSdate
        
        Sheets(Sheets.Count).Columns("N:N").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(Sheets.Count).Columns("O:O").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("P:P").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("L:L").Select
        'Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        
        Sheets(Sheets.Count).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("M1").Value = "Currency"
        Sheets(Sheets.Count).Range("M2:M" & recleng).Value = "USD"
        
        Sheets(Sheets.Count).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(Sheets.Count).Range("Q2:Q" & recleng).Value = "Not Applicable"
        
        Sheets(Sheets.Count).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("R1").Value = "Template Type"
        Sheets(Sheets.Count).Range("R2:R" & recleng).Value = temptyp
        
        Sheets(Sheets.Count).Columns("S:S").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("K:K").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("L:L").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("O:O").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("Q:Q").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("T:T").Select
        Selection.Insert Shift:=xlToRight
        
        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
            If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
        Next i
    
        Sheets(Sheets.Count).Cells.Select
        Sheets(Sheets.Count).Cells.EntireColumn.AutoFit
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
    Else
        MsgBox "File was not opened as read-only"
    End If

End Sub

Sub RWA_SBS()
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 3
        tempswth = 2
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"

        Application.ScreenUpdating = False    
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        Else
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If
    ''''''''''' Five Scenarios ''''''''''''''''
    '    snrotyp = "BHC Baseline"
    '    snrotyp = "BHC Stress"
    '    snrotyp = "Supervisory Baseline"
    
    '    snrotyp = "Supervisory Severely Adverse"
    
    '''''''''''' Four Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
    
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        
        Sheets(Sheets.Count).Range("G1").Value = "Forecast Type"
        Sheets(Sheets.Count).Range("H1").Value = "Quarter"
        Sheets(Sheets.Count).Range("I1").Value = "Value"
        
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value   'copy fixed fields
                    Next k
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value 'copy fcst indicator
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value 'copy quarter indicator
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value 'copy value
                Next j
            Next i
        
        Sheets(Sheets.Count).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        Sheets(Sheets.Count).Range("A1").Value = "Region"
        Sheets(Sheets.Count).Range("B1").Value = "Country"
        Sheets(Sheets.Count).Range("C1").Value = "Business"
        Sheets(Sheets.Count).Range("D1").Value = "CCAR"
        Sheets(Sheets.Count).Range("E1").Value = "Stress Test Scenario"
        Sheets(Sheets.Count).Range("F1").Value = "Submitter"
        Sheets(Sheets.Count).Range("G1").Value = "Forecast Start Date"
       
        Sheets(Sheets.Count).Range("A2:A" & recleng).Value = "NAM"
        Sheets(Sheets.Count).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(Sheets.Count).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(Sheets.Count).Range("E2:E" & recleng).Value = snrotyp
        Sheets(Sheets.Count).Range("F2:F" & recleng).Value = " XL94292"
        Sheets(Sheets.Count).Range("G2:G" & recleng).Value = FSdate
        
        Sheets(Sheets.Count).Columns("N:N").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(Sheets.Count).Columns("O:O").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("P:P").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("L:L").Select
        'Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        
        Sheets(Sheets.Count).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("M1").Value = "Currency"
        Sheets(Sheets.Count).Range("M2:M" & recleng).Value = "USD"
        
        Sheets(Sheets.Count).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(Sheets.Count).Range("Q2:Q" & recleng).Value = "Not Applicable"
        
        Sheets(Sheets.Count).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("R1").Value = "Template Type"
        Sheets(Sheets.Count).Range("R2:R" & recleng).Value = temptyp
        
        Sheets(Sheets.Count).Columns("S:S").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("K:K").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("L:L").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("O:O").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("Q:Q").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("T:T").Select
        Selection.Insert Shift:=xlToRight
        
        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
            If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
        Next i
    
        Sheets(Sheets.Count).Cells.Select
        Sheets(Sheets.Count).Cells.EntireColumn.AutoFit
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
    Else
        MsgBox "File was not opened as read-only"
    End If

End Sub

Sub RWA_SSA()
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 4
        tempswth = 2
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"

        Application.ScreenUpdating = False    
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        Else
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If
    ''''''''''' Five Scenarios ''''''''''''''''
    '    snrotyp = "BHC Baseline"
    '    snrotyp = "BHC Stress"
    '    snrotyp = "Supervisory Baseline"
    
    '    snrotyp = "Supervisory Severely Adverse"
    
    '''''''''''' Four Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
    
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        
        Sheets(Sheets.Count).Range("G1").Value = "Forecast Type"
        Sheets(Sheets.Count).Range("H1").Value = "Quarter"
        Sheets(Sheets.Count).Range("I1").Value = "Value"
        
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value   'copy fixed fields
                    Next k
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value 'copy fcst indicator
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value 'copy quarter indicator
                    Sheets(Sheets.Count).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value 'copy value
                Next j
            Next i
        
        Sheets(Sheets.Count).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        Sheets(Sheets.Count).Range("A1").Value = "Region"
        Sheets(Sheets.Count).Range("B1").Value = "Country"
        Sheets(Sheets.Count).Range("C1").Value = "Business"
        Sheets(Sheets.Count).Range("D1").Value = "CCAR"
        Sheets(Sheets.Count).Range("E1").Value = "Stress Test Scenario"
        Sheets(Sheets.Count).Range("F1").Value = "Submitter"
        Sheets(Sheets.Count).Range("G1").Value = "Forecast Start Date"
       
        Sheets(Sheets.Count).Range("A2:A" & recleng).Value = "NAM"
        Sheets(Sheets.Count).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(Sheets.Count).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(Sheets.Count).Range("E2:E" & recleng).Value = snrotyp
        Sheets(Sheets.Count).Range("F2:F" & recleng).Value = " XL94292"
        Sheets(Sheets.Count).Range("G2:G" & recleng).Value = FSdate
        
        Sheets(Sheets.Count).Columns("N:N").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(Sheets.Count).Columns("O:O").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("P:P").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("L:L").Select
        'Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        
        Sheets(Sheets.Count).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("M1").Value = "Currency"
        Sheets(Sheets.Count).Range("M2:M" & recleng).Value = "USD"
        
        Sheets(Sheets.Count).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(Sheets.Count).Range("Q2:Q" & recleng).Value = "Not Applicable"
        
        Sheets(Sheets.Count).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(Sheets.Count).Range("R1").Value = "Template Type"
        Sheets(Sheets.Count).Range("R2:R" & recleng).Value = temptyp
        
        Sheets(Sheets.Count).Columns("S:S").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("K:K").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("L:L").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("O:O").Select
        Selection.Insert Shift:=xlToRight
        
        Sheets(Sheets.Count).Columns("Q:Q").Select
        Selection.Cut
        Sheets(Sheets.Count).Columns("T:T").Select
        Selection.Insert Shift:=xlToRight
        
        For i = recleng To 2 Step -1
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
            If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
        Next i
    
        Sheets(Sheets.Count).Cells.Select
        Sheets(Sheets.Count).Cells.EntireColumn.AutoFit
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
    Else
        MsgBox "File was not opened as read-only"
    End If

End Sub

Sub Overlay_BBS()
        ' Check to see if the active workbook was
        ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
        ''''''''''''''''''''''''KEEP the last row in A column the data line!!''''''''''''''''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 1
        tempswth = 3
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"
        Application.ScreenUpdating = False    
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        Else
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If
        ''''''''''' Five Scenarios ''''''''''''''''
        '    snrotyp = "BHC Baseline"
        '    snrotyp = "BHC Stress"
        '    snrotyp = "Supervisory Baseline"
        
        '    snrotyp = "Supervisory Severely Adverse"
        
        '''''''''''' Four Templates '''''''''''''''
        '    temptyp = "Stress Loss"          '' RBLP (14A)
        '    temptyp = "RWA"
        '    temptyp = "Overlay"
        '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
       Application.ScreenUpdating = False     
        finshetname = "Tempname"
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
    
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        
        Sheets(finshetname).Range("G1").Value = "Forecast Type"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "Value"
        
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value
                Next j
            Next i
        
        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        Sheets(finshetname).Columns("N:N").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("P:P").Select
        Selection.Cut
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("L:L").Select
        '    Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
        Sheets(finshetname).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("Q2:Q" & recleng).Value = "Not Applicable"
        Sheets(finshetname).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("R1").Value = "Template Type"
        Sheets(finshetname).Range("R2:R" & recleng).Value = temptyp

        recleng = Sheets(finshetname).Cells(Rows.Count, 1).End(xlUp).Row    
        For i = recleng To 2 Step -1
            If IsEmpty(Sheets(finshetname).Range("L" & i).Value) = True Then Sheets(finshetname).Range("L" & i).Value = 0
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
            If (Sheets(finshetname).Range("H" & i).Value = "FIRST_HELOAN" Or Sheets(finshetname).Range("H" & i).Value = "HELOC") And Sheets(finshetname).Range("P" & i).Value = "Non CBNA" Then Sheets(finshetname).Rows(i).Delete
        Next i
    
        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit

        Dim r As Integer
        For r = Sheets(finshetname).UsedRange.Rows.Count To 1 Step -1
            If Cells(r, "K") = "MODEL_OUTPUT" Then
                Sheets(finshetname).Rows(r).EntireRow.Delete
            End If
        Next

        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
    Else
        MsgBox "File was not opened as read-only"
    End If
End Sub

Sub Overlay_BDS()
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
    ''''''''''''''''''''''''KEEP the last row in A column the data line!!''''''''''''''''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName
        snroswth = 2
        tempswth = 3
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"
        Application.ScreenUpdating = False    
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        Else
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If
    ''''''''''' Five Scenarios ''''''''''''''''
    '    snrotyp = "BHC Baseline"
    '    snrotyp = "BHC Stress"
    '    snrotyp = "Supervisory Baseline"
    
    '    snrotyp = "Supervisory Severely Adverse"
    '''''''''''' Four Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False    
        finshetname = "Tempname"
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
    
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        
        Sheets(finshetname).Range("G1").Value = "Forecast Type"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "Value"
        
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value
                Next j
            Next i
        
        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        Sheets(finshetname).Columns("N:N").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("P:P").Select
        Selection.Cut
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("L:L").Select
    '    Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
            Sheets(finshetname).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("Q2:Q" & recleng).Value = "Not Applicable"
            Sheets(finshetname).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("R1").Value = "Template Type"
        Sheets(finshetname).Range("R2:R" & recleng).Value = temptyp
    
        ' For i = recleng To 2 Step -1
        '     If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
        '     If (Sheets(finshetname).Range("H" & i).Value = "FIRST_HELOAN" Or Sheets(finshetname).Range("H" & i).Value = "HELOC") And Sheets(finshetname).Range("P" & i).Value = "Non CBNA" Then Sheets(finshetname).Rows(i).Delete
        ' Next i
        
        recleng = Sheets(finshetname).Cells(Rows.Count, 1).End(xlUp).Row    
        For i = recleng To 2 Step -1
            If IsEmpty(Sheets(finshetname).Range("L" & i).Value) = True Then Sheets(finshetname).Range("L" & i).Value = 0
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
            If (Sheets(finshetname).Range("H" & i).Value = "FIRST_HELOAN" Or Sheets(finshetname).Range("H" & i).Value = "HELOC") And Sheets(finshetname).Range("P" & i).Value = "Non CBNA" Then Sheets(finshetname).Rows(i).Delete
        Next i

        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit

        Dim r As Integer
        For r = Sheets(finshetname).UsedRange.Rows.Count To 1 Step -1
            If Cells(r, "K") = "MODEL_OUTPUT" Then
                Sheets(finshetname).Rows(r).EntireRow.Delete
            End If
        Next

        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
    Else
        MsgBox "File was not opened as read-only"
    End If
End Sub

Sub Overlay_SBS()
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
    ''''''''''''KEEP the last row in A column the data line!!''''''''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 3
        tempswth = 3
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"
        Application.ScreenUpdating = False
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        Else
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If      
    ''''''''''' Five Scenarios ''''''''''''''''
    '    snrotyp = "BHC Baseline"
    '    snrotyp = "BHC Stress"
    '    snrotyp = "Supervisory Baseline"
    
    '    snrotyp = "Supervisory Severely Adverse"
    
    '''''''''''' Four Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False    
        finshetname = "Tempname"
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
    
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        
        Sheets(finshetname).Range("G1").Value = "Forecast Type"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "Value"
        
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value
                Next j
            Next i
        
        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        Sheets(finshetname).Columns("N:N").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("P:P").Select
        Selection.Cut
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("L:L").Select
    '    Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
            Sheets(finshetname).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("Q2:Q" & recleng).Value = "Not Applicable"
            Sheets(finshetname).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("R1").Value = "Template Type"
        Sheets(finshetname).Range("R2:R" & recleng).Value = temptyp
    
    '        For i = recleng To 2 Step -1
    '            If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
    '        Next i
        ' For i = recleng To 2 Step -1
        '     If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
        '     If (Sheets(finshetname).Range("H" & i).Value = "FIRST_HELOAN" Or Sheets(finshetname).Range("H" & i).Value = "HELOC") And Sheets(finshetname).Range("P" & i).Value = "Non CBNA" Then Sheets(finshetname).Rows(i).Delete
        ' Next i

        recleng = Sheets(finshetname).Cells(Rows.Count, 1).End(xlUp).Row    
        For i = recleng To 2 Step -1
            If IsEmpty(Sheets(finshetname).Range("L" & i).Value) = True Then Sheets(finshetname).Range("L" & i).Value = 0
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
            If (Sheets(finshetname).Range("H" & i).Value = "FIRST_HELOAN" Or Sheets(finshetname).Range("H" & i).Value = "HELOC") And Sheets(finshetname).Range("P" & i).Value = "Non CBNA" Then Sheets(finshetname).Rows(i).Delete
        Next i

        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit

        Dim r As Integer
        For r = Sheets(finshetname).UsedRange.Rows.Count To 1 Step -1
            If Cells(r, "K") = "MODEL_OUTPUT" Then
                Sheets(finshetname).Rows(r).EntireRow.Delete
            End If
        Next

        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
    Else
        MsgBox "File was not opened as read-only"
    End If
End Sub

Sub Overlay_SSA()
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
    ''''''''''''KEEP the last row in A column the data line!!''''''''''''''''
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\upload\" & myFileName

        snroswth = 4
        tempswth = 3
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"
        Application.ScreenUpdating = False
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        Else
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If      
    ''''''''''' Five Scenarios ''''''''''''''''
    '    snrotyp = "BHC Baseline"
    '    snrotyp = "BHC Stress"
    '    snrotyp = "Supervisory Baseline"    
    '    snrotyp = "Supervisory Severely Adverse"
    
    '''''''''''' Four Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
       
        col2copy = 6
        col2inst = 2
        Application.ScreenUpdating = False    
        finshetname = "Tempname"
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
    
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        
        Sheets(finshetname).Range("G1").Value = "Forecast Type"
        Sheets(finshetname).Range("H1").Value = "Quarter"
        Sheets(finshetname).Range("I1").Value = "Value"
        
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 7).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 8).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value
                Next j
            Next i
        
        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = "XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        Sheets(finshetname).Columns("N:N").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("O:O").Select
        Selection.Cut
        Sheets(finshetname).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("P:P").Select
        Selection.Cut
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("L:L").Select
    '    Selection.SpecialCells(xlBlanks).Value = 0#
        Selection.NumberFormat = "0.00"
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
            Sheets(finshetname).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("Q2:Q" & recleng).Value = "Not Applicable"
            Sheets(finshetname).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("R1").Value = "Template Type"
        Sheets(finshetname).Range("R2:R" & recleng).Value = temptyp
    
    '        For i = recleng To 2 Step -1
    '            If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
    '        Next i
        ' For i = recleng To 2 Step -1
        '     If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
        '     If (Sheets(finshetname).Range("H" & i).Value = "FIRST_HELOAN" Or Sheets(finshetname).Range("H" & i).Value = "HELOC") And Sheets(finshetname).Range("P" & i).Value = "Non CBNA" Then Sheets(finshetname).Rows(i).Delete
        ' Next i

        recleng = Sheets(finshetname).Cells(Rows.Count, 1).End(xlUp).Row    
        For i = recleng To 2 Step -1
            If IsEmpty(Sheets(finshetname).Range("L" & i).Value) = True Then Sheets(finshetname).Range("L" & i).Value = 0
            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
            If (Sheets(finshetname).Range("H" & i).Value = "FIRST_HELOAN" Or Sheets(finshetname).Range("H" & i).Value = "HELOC") And Sheets(finshetname).Range("P" & i).Value = "Non CBNA" Then Sheets(finshetname).Rows(i).Delete
        Next i

        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit

        Dim r As Integer
        For r = Sheets(finshetname).UsedRange.Rows.Count To 1 Step -1
            If Cells(r, "K") = "MODEL_OUTPUT" Then
                Sheets(finshetname).Rows(r).EntireRow.Delete
            End If
        Next

        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
    Else
        MsgBox "File was not opened as read-only"
    End If
End Sub

Sub Sensitivity()
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim col2copy As Integer
        Dim col2inst As Integer
        Dim col2trans As Integer
        Dim row2copy As Integer
        Dim recleng As Integer
        Dim lastRowIndex As Integer
        Dim lastColIndex As Integer
        Dim datashetname As String
        Dim finshetname As String
        Dim snrotyp As String
        Dim temptyp As String
        Dim sc1 As Integer
        Dim ws As Worksheet
        Dim snroswth As Integer
        Dim tempswth As Integer
        Dim SaveToDir As String
        Dim FileName As String
        Dim FSdate AS String
        Dim CCARmonth AS String

        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFullName = myPath & "\" & myFileName

        snroswth = 1
        tempswth = 4
        CCARmonth = "CCAR 2020-01"
        FSdate = "Q1 2020"
        Application.ScreenUpdating = False    
        If snroswth = 1 Then
        snrotyp = "BHC Baseline"
        ElseIf snroswth = 2 Then
        snrotyp = "BHC Stress"
        ElseIf snroswth = 3 Then
        snrotyp = "Supervisory Baseline"
        ElseIf snroswth = 4 Then
        snrotyp = "Supervisory Severely Adverse"   
        ElseIf snroswth = 0 Then
        snrotyp = "Actuals"
        Else
        snrotyp = "============================="
        End If
        
        If tempswth = 1 Then
        temptyp = "Stress Loss"
        ElseIf tempswth = 2 Then
        temptyp = "RWA"
        ElseIf tempswth = 3 Then
        temptyp = "Overlay"
        ElseIf tempswth = 4 Then
        temptyp = "Sensitivity"
        ElseIf tempswth = 6 Then
        temptyp = "======== LLR ========="
        End If
    '''''''''''' Four Templates '''''''''''''''
    '    temptyp = "Stress Loss"          '' RBLP (14A)
    '    temptyp = "RWA"
    '    temptyp = "Overlay"
    '    temptyp = "Sensitivity"
       
        col2copy = 5
        col2inst = 2
        
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        Sheets.Add After:=Sheets(Sheets.Count)
        finshetname = "Tempname"
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name = finshetname Then
                Sheets(n).Delete
                Exit For
            End If
        Next n
        Sheets(Sheets.Count).Name = finshetname
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 1).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(3, Columns.Count).End(xlToLeft).Column
    
        row2copy = lastRowIndex - col2inst - 1
        col2trans = lastColIndex - col2copy
        recleng = col2trans * row2copy + 1
    
        For j = 1 To col2copy
            Sheets(finshetname).Cells(1, j).Value = Sheets(datashetname).Cells(3, j).Value
        Next j
        
        Sheets(finshetname).Range("I1").Value = "Forecast Type"
        Sheets(finshetname).Range("J1").Value = "Quarter"
        Sheets(finshetname).Range("L1").Value = "Value"
        
            For i = 1 To row2copy
                For j = 1 To col2trans
                    For k = 1 To col2copy
                        Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, k).Value = Sheets(datashetname).Cells(3 + i, k).Value   'copy fixed fields
                    Next k
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 9).Value = Sheets(datashetname).Cells(col2inst, col2copy + j).Value 'copy fcst indicator
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 10).Value = Sheets(datashetname).Cells(col2inst + 1, col2copy + j).Value 'copy quarter indicator
                    Sheets(finshetname).Cells(1 + (i - 1) * col2trans + j, 12).Value = Sheets(datashetname).Cells(col2inst + 1 + i, col2copy + j).Value 'copy value
                Next j
            Next i
        
        Sheets(finshetname).Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("A1").Value = "Region"
        Sheets(finshetname).Range("B1").Value = "Country"
        Sheets(finshetname).Range("C1").Value = "Business"
        Sheets(finshetname).Range("D1").Value = "CCAR"
        Sheets(finshetname).Range("E1").Value = "Stress Test Scenario"
        Sheets(finshetname).Range("F1").Value = "Submitter"
        Sheets(finshetname).Range("G1").Value = "Forecast Start Date"
        Sheets(finshetname).Range("A2:A" & recleng).Value = "NAM"
        Sheets(finshetname).Range("B2:B" & recleng).Value = "US"

        If InStr(1, myPath, "NCL1\AdHoc") <> 0 Then
            Sheets(finshetname).Range("C2:C" & recleng).Value = "CMI"
        ElseIf InStr(1, myPath, "Forecast\CCAR") <> 0 Then
            If InStr(1, myPath, "\DM") <> 0 Then
                If InStr(1, myFileName, "CCB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CCB"
                ElseIf InStr(1, myFileName, "SBB ") <> 0 Then
                    Sheets(finshetname).Range("C2:C" & recleng).Value = "SB-CITIBUSINESS"
                End If
            Else
                Sheets(finshetname).Range("C2:C" & recleng).Value = "USCCB"
            End If
        End If

        Sheets(finshetname).Range("D2:D" & recleng).Value = CCARmonth
        Sheets(finshetname).Range("E2:E" & recleng).Value = snrotyp
        Sheets(finshetname).Range("F2:F" & recleng).Value = " XL94292"
        Sheets(finshetname).Range("G2:G" & recleng).Value = FSdate
        Sheets(finshetname).Columns("P:P").Select
        Selection.Cut
        Sheets(finshetname).Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("Q:Q").Select
        Selection.Cut
        Sheets(finshetname).Columns("J:J").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("S:S").Select
        Selection.Cut
        Sheets(finshetname).Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        Sheets(finshetname).Columns("M:M").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("M1").Value = "Currency"
        Sheets(finshetname).Range("M2:M" & recleng).Value = "USD"
        Sheets(finshetname).Columns("Q:Q").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("Q1").Value = "Margin Lending Indicator"
        Sheets(finshetname).Range("Q2:Q" & recleng).Value = "Not Applicable"
        Sheets(finshetname).Columns("R:R").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets(finshetname).Range("R1").Value = "Template Type"
        Sheets(finshetname).Range("R2:R" & recleng).Value = temptyp
        
        Sheets(finshetname).Columns("L:L").Select
        Selection.NumberFormat = "0.00"
        For i = recleng To 2 Step -1

            If Sheets(finshetname).Range("S" & i).Value = "ACT_US_OD" Or Sheets(finshetname).Range("S" & i).Value = "CMP_132791_OD" Then Sheets(finshetname).Range("Q" & i).Value = "Others"  
            If Sheets(finshetname).Range("L" & i).Value = 0# Or Sheets(finshetname).Range("L" & i).Value = "." Then Sheets(finshetname).Rows(i).Delete
        Next i
    
        Sheets(finshetname).Cells.Select
        Sheets(finshetname).Cells.EntireColumn.AutoFit
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> finshetname Then
                ActiveSheet.Name = finshetname
                Sheets(n).Delete
                Exit For
            End If
        Next n
        
        ActiveWorkbook.SaveAs FileName:=myFullName & "_upload_n0", FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = True
    Else
        MsgBox "File was not opened as read-only"
    End If
End Sub

Sub ModIDsplit()
    'Codes for splitting by Model ID 01/24/2019
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim n As Integer
        Dim sc1 As Integer
        Dim ws1 As Worksheet
        Dim datashetname As String
        Dim tabname As String
        Dim lastRowIndex As Integer 'last row of data
        Dim lastColIndex As Integer 'last column of data
        Dim modidIndex As Integer   'number of Model IDs
        Dim myPath As String
        Dim myFileName As String
        Dim myFullName As String
        
        myPath = ActiveWorkbook.Path
        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 4)
        myFullName = myPath & "\" & myFileName         
        datashetname = Sheet1.Name  ' make sure datafile is sheet 1
        lastRowIndex = Sheets(datashetname).Cells(Rows.Count, 19).End(xlUp).Row
        lastColIndex = Sheets(datashetname).Cells(1, Columns.Count).End(xlToLeft).Column
        Range("Z1", Range("Z1").End(xlDown)).Delete
        Sheets(datashetname).Range("S1:S" & lastRowIndex).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("Z1"), Unique:=True
        modidIndex = Sheets(datashetname).Cells(Rows.Count, "Z").End(xlUp).Row
        Application.DisplayAlerts = False
        sc1 = Sheets.Count
        For n = sc1 To 1 Step -1
            If Sheets(n).Name <> datashetname Then
                Sheets(n).Delete
            End If
        Next n
        For j = 2 To modidIndex
            Sheets.Add After:=Sheets(Sheets.Count)
            Sheets(Sheets.Count).Name = Sheets(datashetname).Range("Z" & j).Value
        Next j
        
        For Each ws1 In ThisWorkbook.Worksheets
            If ws1.Name <> datashetname Then
                ws1.Range("A1:U1").Value = Sheets(datashetname).Range("A1:U1").Value
            End If
        Next ws1
        
        For k = 2 To lastRowIndex
            For j = 2 To modidIndex
                If Sheets(datashetname).Range("S" & k).Value = Sheets(datashetname).Range("Z" & j).Value Then
                    Sheets(Sheets(datashetname).Range("Z" & j).Value).Range("A" & k & ":U" & k).Value = Sheets(datashetname).Range("A" & k & ":U" & k).Value
                End If
            Next j
        Next k
        
        For j = 2 To modidIndex
            For k = lastRowIndex To 2 Step -1
                If Sheets(Sheets(datashetname).Range("Z" & j).Value).Range("S" & k) = "" Then
                    Sheets(Sheets(datashetname).Range("Z" & j).Value).Rows(k).Delete
                End If
            Next k
        Next j
     
        For Each ws1 In ThisWorkbook.Worksheets
            If ws1.Name <> datashetname Then
                ws1.SaveAs Filename:=myFullName & "_" & ws1.Name, FileFormat:=xlCSV
            End If
        Next ws1
    Else
        MsgBox "File was not opened as read-only"
    End If
End Sub

Sub Example2()
    'Codes for checking if file is read-only 01/24/2019
    ' Check to see if the active workbook was
    ' opened as read-only within Microsoft Excel.
    If ActiveWorkbook.ReadOnly Then
        MsgBox "File was opened as read-only"
    Else
        MsgBox "File was not opened as read-only"
    End If
End Sub

Sub Example1()
    'Codes for checking if file is read-only 01/24/2019
    ' Test to see if the Read-only attribute was assigned to the file.
    If GetAttr("c:\test.xls") And vbReadOnly Then
        MsgBox "File is Read-only"
    Else
        MsgBox "File is not read-only"
    End If

End Sub