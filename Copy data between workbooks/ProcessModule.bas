Attribute VB_Name = "ProcessModule"
Private Function CompileWorkbookName(prefix, date1, date2)

CompileWorkbookName = prefix & Format(date1, "yymmdd") & " to " & Format(date2, "yymmdd")

End Function


Function Difference(date1, date2)

Dim sheetName As String
sheetName = CompileWorkbookName("Sheet_Name_Here ", date1, date2)

'Copy the values
v1 = CopyModule.GetValue("C:\Users\Username\Documents\", sheetName & ".xlsx", "Sheet1", "E15")
v2 = CopyModule.GetValue("C:\Users\Username\Documents\", sheetName & ".xlsx", "Sheet1", "E16")
Range("C20").Value = v1 - v2

End Function


Function CopyPb10(date1, date2)

Dim sheetName As String
sheetName = CompileWorkbookName("Sheet_Name_Here ", date1, date2)

'Copy the values
ActiveSheet.Range("C16") = CopyModule.GetValue("C:\Users\Username\Documents\", sheetName & ".xlsx", "Sheet1", "D15")
ActiveSheet.Range("C17") = CopyModule.GetValue("C:\Users\Username\Documents\", sheetName & ".xlsx", "Sheet1", "D16")
ActiveSheet.Range("C18") = CopyModule.GetValue("C:\Users\Username\Documents\", sheetName & ".xlsx", "Sheet1", "D17")
ActiveSheet.Range("C19") = CopyModule.GetValue("C:\Users\Username\Documents\", sheetName & ".xlsx", "Sheet1", "D18")

End Function
