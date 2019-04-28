VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DateForm 
   Caption         =   "UserForm1"
   ClientHeight    =   1810
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   2820
   OleObjectBlob   =   "DateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

'Activate the sheet
Sheet1.Activate
Set target = ActiveWorkbook

'Insert the period into C2 cell
Dim date1 As Date
Dim date2 As Date

date1 = DTPicker1.Value
date2 = DTPicker2.Value
Range("C2").Value = Format(date1, "dd.mm.yyyy") & " - " & Format(date2, "dd.mm.yyyy")

'Calculate from the values from another workbook
Call ProcessModule.Difference(date1, date2)

'Copy the data from source workbooks
Call ProcessModule.CopyPb10(date1, date2)

'Close user form
Unload DateForm

End Sub


