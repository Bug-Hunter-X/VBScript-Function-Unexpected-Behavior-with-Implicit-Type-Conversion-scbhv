Function f(x)
  If IsNumeric(x) Then 
    x = CDbl(x)  'Explicitly convert to Double
    If x > 10 Then
      f = x / 2
    Else
      f = x * 2
    End If
  Else
    f = "Invalid input: x must be a number"
  End If
End Function

MsgBox f(12) ' correct output: 6
MsgBox f(5) 'correct output: 10
MsgBox f("abc") 'correct output: Invalid Input