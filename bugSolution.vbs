Function MyFunction(param1)
  If param1 = "" Then
    Err.Raise 13, , "Type mismatch: Parameter is an empty string"
  ElseIf IsEmpty(param1) Then
    Err.Raise 13, , "Type mismatch: Parameter is empty"
  End If
  ' ... rest of the function
End Function