Sub generate_pivot()
        '''
        ' @bw11496 (vince.wang@citi.com)
        '                   - 06/05/2019
        ' To automatically generate the pivottable files in method b reporting
        ' Create temp folder in working path, such as - 
        ' U:\CFRisk\NCL1\AdHoc\CCAR\2019 CCAR\01. BHC Basecase\Support Data\MDL - B
        ' Copy model outputs to temp folder
        ' Open PivotTable file in working path and run VBA code 
        ' to copy outputs files in temp folder to PivotTable files.
        ''' 
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
        Dim pos As Integer
        Dim t_name_str As String
        Dim name_str As String
        Dim tFSO As Object
        Dim tFolder As Object
        Dim tFile As Object
        Dim tFiles As Object
        Dim tResults As Variant
        Dim my_wkbook As Workbook
        Dim my_wksheet As Worksheet
        Dim t_wkbook As Workbook
        Dim t_wksheet As Worksheet

        ' Test to see if the Read-only attribute is off.
    If ActiveWorkbook.ReadOnly = False Then        
                
        myPath = ActiveWorkbook.Path
        '        myFileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
        myFileName = ThisWorkbook.Name
        myFullName = myPath & "\" & myFileName
        myTemppath = myPath & "\temp\"
        Set tFSO = CreateObject("Scripting.FileSystemObject")
        Set tFolder = tFSO.GetFolder(myTemppath)
        Set tFiles = tFolder.Files
        
        '        v0_raw_CMI   v0_raw_CPB    v0_raw_USCCB v0_raw_GCMD

        If InStr(1, myFileName, "v0_raw_CMI") <> 0 Then
        ReDim tResults(1 To tFiles.Count)
        i = 1
        For Each tFile In tFiles
            If InStr(1, tFile.Name, "CMI_v0_raw") <> 0 Then
            tResults(i) = tFile.Name
                If InStr(1, tFile.Name, "pivot_DelqNCL_CMI") <> 0 Then
                    Sheets("Data Delq").UsedRange.ClearContents
                    Set t_wkbook = Workbooks.Open(FileName:=tFile)
                    Set t_wksheet = t_wkbook.Worksheets(1)
                    t_wksheet.UsedRange.Copy
                    Workbooks(myFileName).Worksheets("Data Delq").Activate
                    Worksheets("Data Delq").Range("A1").Select
                    Selection.PasteSpecial
                    Application.CutCopyMode = False
                    t_wkbook.Close SaveChanges:=False

                ElseIf InStr(1, tFile.Name, "pivot_NCL_CMI") <> 0 Then
                    Sheets("Data NCL").UsedRange.ClearContents
                    Set t_wkbook = Workbooks.Open(FileName:=tFile)
                    Set t_wksheet = t_wkbook.Worksheets(1)
                    t_wksheet.UsedRange.Copy
                    Workbooks(myFileName).Worksheets("Data NCL").Activate
                    Worksheets("Data NCL").Range("A1").Select
                    Selection.PasteSpecial
                    Application.CutCopyMode = False
                    t_wkbook.Close SaveChanges:=False

                End If
            i = i + 1
            End If
        Next tFile
        End If

        If InStr(1, myFileName, "v0_raw_CPB") <> 0 Then
        ReDim tResults(1 To tFiles.Count)
        i = 1
        For Each tFile In tFiles
            If InStr(1, tFile.Name, "CPB_v0_raw") <> 0 Then
            tResults(i) = tFile.Name
                If InStr(1, tFile.Name, "pivot_DelqNCL_CPB") <> 0 Then
                    Sheets("Data Delq").UsedRange.ClearContents
                    Set t_wkbook = Workbooks.Open(FileName:=tFile)
                    Set t_wksheet = t_wkbook.Worksheets(1)
                    t_wksheet.UsedRange.Copy
                    Workbooks(myFileName).Worksheets("Data Delq").Activate
                    Worksheets("Data Delq").Range("A1").Select
                    Selection.PasteSpecial
                    Application.CutCopyMode = False
                    t_wkbook.Close SaveChanges:=False
                ElseIf InStr(1, tFile.Name, "pivot_NCL_CPB") <> 0 Then
                    Sheets("Data NCL").UsedRange.ClearContents
                    Set t_wkbook = Workbooks.Open(FileName:=tFile)
                    Set t_wksheet = t_wkbook.Worksheets(1)
                    t_wksheet.UsedRange.Copy
                    Workbooks(myFileName).Worksheets("Data NCL").Activate
                    Worksheets("Data NCL").Range("A1").Select
                    Selection.PasteSpecial
                    Application.CutCopyMode = False
                    t_wkbook.Close SaveChanges:=False
                End If
            i = i + 1
            End If
        Next tFile
        End If

        If InStr(1, myFileName, "v0_raw_USRB") <> 0 Then
        ReDim tResults(1 To tFiles.Count)
        i = 1
        For Each tFile In tFiles
            If InStr(1, tFile.Name, "USCCB_v0_raw") <> 0 Then
            tResults(i) = tFile.Name
                If InStr(1, tFile.Name, "pivot_DelqNCL_USCCB") <> 0 Then
                    Sheets("Data Delq").UsedRange.ClearContents
                    Set t_wkbook = Workbooks.Open(FileName:=tFile)
                    Set t_wksheet = t_wkbook.Worksheets(1)
                    t_wksheet.UsedRange.Copy
                    Workbooks(myFileName).Worksheets("Data Delq").Activate
                    Worksheets("Data Delq").Range("A1").Select
                    Selection.PasteSpecial
                    Application.CutCopyMode = False
                    t_wkbook.Close SaveChanges:=False
                ElseIf InStr(1, tFile.Name, "pivot_NCL_USCCB") <> 0 Then
                    Sheets("Data NCL").UsedRange.ClearContents
                    Set t_wkbook = Workbooks.Open(FileName:=tFile)
                    Set t_wksheet = t_wkbook.Worksheets(1)
                    t_wksheet.UsedRange.Copy
                    Workbooks(myFileName).Worksheets("Data NCL").Activate
                    Worksheets("Data NCL").Range("A1").Select
                    Selection.PasteSpecial
                    Application.CutCopyMode = False
                    t_wkbook.Close SaveChanges:=False
                End If
            i = i + 1
            End If
        Next tFile
        End If

        If InStr(1, myFileName, "v1_ma_CMI") <> 0 Then
        ReDim tResults(1 To tFiles.Count)
        i = 1
        For Each tFile In tFiles
            If InStr(1, tFile.Name, "CMI_v1_ma") <> 0 Then
            tResults(i) = tFile.Name
                If InStr(1, tFile.Name, "pivot_DelqNCL_CMI") <> 0 Then
                    Sheets("Data Delq").UsedRange.ClearContents
                    Set t_wkbook = Workbooks.Open(FileName:=tFile)
                    Set t_wksheet = t_wkbook.Worksheets(1)
                    t_wksheet.UsedRange.Copy
                    Workbooks(myFileName).Worksheets("Data Delq").Activate
                    Worksheets("Data Delq").Range("A1").Select
                    Selection.PasteSpecial
                    Application.CutCopyMode = False
                    t_wkbook.Close SaveChanges:=False
                ElseIf InStr(1, tFile.Name, "pivot_NCL_CMI") <> 0 Then
                    Sheets("Data NCL").UsedRange.ClearContents
                    Set t_wkbook = Workbooks.Open(FileName:=tFile)
                    Set t_wksheet = t_wkbook.Worksheets(1)
                    t_wksheet.UsedRange.Copy
                    Workbooks(myFileName).Worksheets("Data NCL").Activate
                    Worksheets("Data NCL").Range("A1").Select
                    Selection.PasteSpecial
                    Application.CutCopyMode = False
                    t_wkbook.Close SaveChanges:=False
                End If
            i = i + 1
            End If
        Next tFile
        End If

        If InStr(1, myFileName, "v0_raw_GCMD") <> 0 Then
        ReDim tResults(1 To tFiles.Count)
        i = 1
        For Each tFile In tFiles
            If InStr(1, tFile.Name, "GCMD_v0_raw") <> 0 Then
            tResults(i) = tFile.Name
                If InStr(1, tFile.Name, "pivot_DelqNCL_GCMD") <> 0 Then
                    Sheets("Data Delq").UsedRange.ClearContents
                    Set t_wkbook = Workbooks.Open(FileName:=tFile)
                    Set t_wksheet = t_wkbook.Worksheets(1)
                    t_wksheet.UsedRange.Copy
                    Workbooks(myFileName).Worksheets("Data Delq").Activate
                    Worksheets("Data Delq").Range("A1").Select
                    Selection.PasteSpecial
                    Application.CutCopyMode = False
                    t_wkbook.Close SaveChanges:=False
                ElseIf InStr(1, tFile.Name, "pivot_NCL_GCMD") <> 0 Then
                    Sheets("Data NCL").UsedRange.ClearContents
                    Set t_wkbook = Workbooks.Open(FileName:=tFile)
                    Set t_wksheet = t_wkbook.Worksheets(1)
                    t_wksheet.UsedRange.Copy
                    Workbooks(myFileName).Worksheets("Data NCL").Activate
                    Worksheets("Data NCL").Range("A1").Select
                    Selection.PasteSpecial
                    Application.CutCopyMode = False
                    t_wkbook.Close SaveChanges:=False
                End If
            i = i + 1
            End If
        Next tFile
        End If


        Workbooks(myFileName).RefreshAll
        Application.DisplayAlerts = False
        Application.EnableEvents = False        
        Workbooks(myFileName).Close SaveChanges:=True        


    Else
        MsgBox "File is Read-only"
    End If

End Sub