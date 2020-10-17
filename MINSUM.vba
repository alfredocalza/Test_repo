Private Function MINSUM(ByVal Rng As Range) As Double
Application.Volatile
Dim Cell As Range, FirstCell As Range, Result As Double, Ws As Worksheet
Set FirstCell = Range(Cells(Rng.Row, Rng.Column).Address) Result = FirstCell.Value Set Ws = Rng.Worksheet
With Ws
For Each Cell In Rng
Result = Application.Min(Result, Application.Sum(.Range(FirstCell.Address, .Cells(Cell.Row, Cell.Column).Address)))
Next Cell
End With
MINSUM = Result
End Function
