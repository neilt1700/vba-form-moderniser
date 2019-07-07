VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USampleUserForm 
   Caption         =   "Sample Form"
   ClientHeight    =   2376
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "USampleUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USampleUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) Commtap CIC 2019
' Available under the MIT license: see the LICENSE file at the root of this
' project.
' Contact: tap@commtap.org

Option Explicit


Private p_oLabelControlsManager As CLabelControlsManager

' Click methods must be declared public
Public Sub CommandButton1_Click()
  MsgBox "You clicked the command button."
End Sub

' For labels you want to be used as modern style controls:
' - add a prefix to all of them e.g. "LabelButton" (so that other labels that
'   you don't want converted into buttons aren't).
' - give each a click routine, and declare these as public - they are going to
'   be called from outside of the form.
'
Public Sub LabelButtonCancel_Click()
  Me.Hide
End Sub

Public Sub LabelButtonOK_Click()

  MsgBox "You clicked on the ""OK"" button."

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

  If CloseMode = vbFormControlMenu Then
    LabelButtonCancel_Click
    Cancel = True
  End If
  
End Sub

' Each form also needs one of these - this takes the hover off a button when
' the mouse is moved off it.
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  p_oLabelControlsManager.LabelControls.UpdateControlButtonState
End Sub

Public Sub InitiateProperties()

  ' This styles the form generally:
  ModerniseControls Me.Controls
  
  ' This converts command buttons into modern controls.
  ' The default button is the one that will run if enter is pressed.
    
  ' These must be re/initialised here.
  FormModerniserModule.ActiveButton = vbNullString
  FormModerniserModule.DefaultButton = "CommandButton1"
  
  Dim arrLabelControlsOrder() As String
  arrLabelControlsOrder = Split("CommandButton1 CommandButton2")
  
  Set p_oLabelControlsManager = Factory.CreateCLabelControlsManager(Me.Controls, _
                                                                    arrLabelControlsOrder)

End Sub
