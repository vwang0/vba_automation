Sub UpdateCalcu()

Dim i As Integer
Dim j As Integer
Dim strB As String
Dim strC As String
Dim strD As String
Dim strE As String


For i = 1 To 20
j = i + 5
strB = "B" & j
strC = "C" & j
strE = "E" & j
strF = "F" & j
If ThisWorkbook.Sheets("Results").Range(strC) <> 0 Then
ThisWorkbook.Sheets("TDR Template Test Step Up").Range("P1") = ThisWorkbook.Sheets("Results").Range(strB)
ThisWorkbook.Sheets("Results").Range(strE) = ThisWorkbook.Sheets("TDR Template Test Step Up").Range("P7")
ThisWorkbook.Sheets("Results").Range(strF) = ThisWorkbook.Sheets("TDR Template Test Step Up").Range("P9")
End If

Next i


End Sub

