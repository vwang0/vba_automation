Sub Main()
    Dim rng_from as Range
    Dim rng_to as Range

    Set rng_from = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
    MsgBox "The cells selected were " & rng_from.Address

    Worksheets("rwa").rng_from.Copy

    Set rng_from = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
    MsgBox "The cells selected were " & rng_to.Address

    For Each cell In rng_to
    Worksheets("rblp").Range(cell).PasteSpecial Paste:=xlPasteFormats
    Next cell

End Sub



Sub Main()

Dim rng_from As Range
Dim rng_to As Range

'Set rng_from = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
'MsgBox "The cells selected were " & rng_from.Address
'
'Worksheets("HAHA").rng_from.Copy

Selection.Copy

Set rng_to = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
'MsgBox "The cells selected were " & rng_to.Address

For Each cell In rng_to
Range(cell).PasteSpecial Paste:=xlPasteFormats
Next cell

End Sub