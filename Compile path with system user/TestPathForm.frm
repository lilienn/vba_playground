VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestPathForm 
   Caption         =   "UserForm1"
   ClientHeight    =   1390
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4450
   OleObjectBlob   =   "TestPathForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TestPathForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

MsgBox getSystemUser()

End Sub

Private Sub CommandButton2_Click()

Dim date1 As Date
date1 = DTPicker1.Value

MsgBox "C:\Users\" & Environ("Username") & "\OneDrive\" & Format(date1, "yy\\mmmm") & "\1.xlsx"

End Sub

Private Function getSystemUser()

getSystemUser = Environ("Username")

End Function
