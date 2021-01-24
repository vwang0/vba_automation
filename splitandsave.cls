
Sub SplitIntoSheets()

    Dim lsrClm As Long
    
    Dim lstRow As Long
    
    'counting last used row
    
    lstRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim uniques As Range
    
    Dim clm As String, clmNo As Long
    
    
    clmNo = Range("D" & "1").Column
    
    Set uniques = Range("D" & "2:" & "D" & lstRow)
    
    Set uniques = RemoveDuplicates(uniques)
    
    Call CreateSheets(uniques, clmNo)

End Sub


Function RemoveDuplicates(uniques As Range) As Range

ThisWorkbook.Activate

    Sheets.Add

    On Error Resume Next

    ActiveSheet.Name = "uniques"

    Sheets("uniques").Activate

    On Error GoTo 0

    uniques.Copy

    Cells(2, 1).Activate

    ActiveCell.PasteSpecial xlPasteValues

    Range("A1").Value = "uniques"


    Dim lstRow As Long

    lstRow = Cells(Rows.Count, 1).End(xlUp).Row

    Range("A2:A" & lstRow).Select

    ActiveSheet.Range(Selection.Address).RemoveDuplicates Columns:=1, Header:=xlNo

    lstRow = Cells(Rows.Count, 1).End(xlUp).Row

    Set RemoveDuplicates = Range("A2:A" & lstRow)

End Function


Sub CreateSheets(uniques As Range, clmNo As Long)

    Dim lstClm As Long

    Dim lstRow As Long
    
    Dim regex As Object


    
    For Each Unique In uniques
    
    
        If Unique Like "[A-Za-z]*" Then
            Sheet1.Activate
    
            lstRow = Cells(Rows.Count, 1).End(xlUp).Row
    
            lstClm = Cells(1, Columns.Count).End(xlToLeft).Column
    
            Dim dataSet As Range
    
            Set dataSet = Range(Cells(1, 1), Cells(lstRow, lstClm))
    
            dataSet.AutoFilter field:=clmNo, Criteria1:=Unique.Value
    
            lstRow = Cells(Rows.Count, 1).End(xlUp).Row
    
            lstClm = Cells(1, Columns.Count).End(xlToLeft).Column
    
            Debug.Print lstRow; lstClm
    
            Set dataSet = Range(Cells(1, 1), Cells(lstRow, lstClm))
    
            dataSet.Copy
    
            Sheets.Add
    
            ActiveSheet.Name = LEFT(Unique,31)
    
            ActiveCell.PasteSpecial xlPasteAll
        
        End If  
       

    Next Unique

End Sub
