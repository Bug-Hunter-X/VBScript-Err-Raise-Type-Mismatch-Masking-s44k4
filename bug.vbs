Function MyFunction(param1)
  If IsEmpty(param1) Then
    Err.Raise 13, , "Param1 cannot be empty"
  End If
  ' ... rest of the function
End Function