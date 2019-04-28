Attribute VB_Name = "CopyModule"
Function GetValue(path, file, sheet, ref)

Dim arg As String

'Checking if file exists
If Right(path, 1) <> "\" Then path = path & "\"
If Dir(path & file) = "" Then
    GetValue = "File Not Found"
    Exit Function
End If

'Create the argument
arg = "'" & path & "[" & file & "]" & sheet & "'!" & _
Range(ref).Address(, , xlR1C1)

GetValue = ExecuteExcel4Macro(arg)
    
End Function
