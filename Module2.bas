Attribute VB_Name = "Module2"

Sub AddIndentSymbol()
Attribute AddIndentSymbol.VB_ProcData.VB_Invoke_Func = "A\n14"
    Dim Rng As Range

    Set Rng = Selection
    For Each Cell In Rng.Cells
        Cell.Value = ">" & Cell.Value
    Next Cell

End Sub

Sub RemoveIndentSymbol()
Attribute RemoveIndentSymbol.VB_ProcData.VB_Invoke_Func = "R\n14"
    Dim Rng As Range

    Set Rng = Selection
    For Each Cell In Rng.Cells
        Cell.Value = Replace(Cell.Value, ">", "")
    Next Cell

End Sub

