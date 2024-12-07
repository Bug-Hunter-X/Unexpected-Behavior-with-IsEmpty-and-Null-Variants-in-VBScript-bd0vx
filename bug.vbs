Function f(a)
  If IsEmpty(a) Then Exit Function
  If VarType(a) <> vbString Then Exit Function
  If Len(a) > 10 Then Exit Function
  Set f = CreateObject("Scripting.FileSystemObject")
end function