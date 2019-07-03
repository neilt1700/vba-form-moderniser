VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USampleUserFormOldStyle 
   Caption         =   "Sample Form"
   ClientHeight    =   2376
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "USampleUserFormOldStyle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USampleUserFormOldStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Public Sub CmdCancel_Click()
  Me.Hide
End Sub

Private Sub CmdOK_Click()
  MsgBox "You clicked on the ""OK"" button."
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

  If CloseMode = vbFormControlMenu Then
    CmdCancel_Click
    Cancel = True
  End If
  
End Sub

