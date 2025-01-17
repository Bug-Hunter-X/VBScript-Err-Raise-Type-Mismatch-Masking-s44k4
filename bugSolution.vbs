Function MyFunction(param1)
  On Error GoTo ErrorHandler
  If IsEmpty(param1) Then
    Err.Raise vbError, , "Param1 cannot be empty"
  End If
  ' ... rest of the function
  Exit Function

ErrorHandler:
  ' Log the actual error details
  WScript.Echo "Error Number: " & Err.Number
  WScript.Echo "Error Description: " & Err.Description
  ' Handle the error appropriately or re-raise
  Err.Raise Err.Number, Err.Source, Err.Description
End Function