Function f(a)
  If IsNull(a) Then
    'Handle Null values appropriately (e.g., assign a default value, raise an error)
    'Example: a = ""
  ElseIf IsEmpty(a) Then
    Exit Function
  ElseIf VarType(a) <> vbString Then
    Exit Function
  ElseIf Len(a) > 10 Then
    Exit Function
  End If
  Set f = CreateObject("Scripting.FileSystemObject")
end function