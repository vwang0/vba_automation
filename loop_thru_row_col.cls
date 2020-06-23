Sub Sensitivity()

    Dim i As Integer
    Dim j As Integer


    For i = 1 To 5
        For j = 1 To 5
            Range("D19").Value = Range("K63").Offset(i, 0).Value
            Range("D21").Value = Range("K63").Offset(0, j).Value
            Range("K63").Offset(i, j).Value = Range("E44").Value
       
    Next j
    Next i
      
        
End Sub
