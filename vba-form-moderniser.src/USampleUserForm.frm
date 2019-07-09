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
   Me.Hide
End Sub

Public Sub CommandButton2_Click()
  MsgBox "You clicked the ""OK"" button."
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

  If CloseMode = vbFormControlMenu Then
    CommandButton1_Click
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
    
  ' These must be re/initialised here.
  FormModerniserModule.ActiveButton = vbNullString
  FormModerniserModule.DefaultButton = "CommandButton1"
  
  ' The order of the buttons when tabbed through.
  Dim arrLabelControlsOrder() As String
  arrLabelControlsOrder = Split("CommandButton2 CommandButton1")
     
  ' This converts command buttons into modern controls.
  ' The default button is the one that will run if enter is pressed.
  Set p_oLabelControlsManager = Factory.CreateCLabelControlsManager(Me.Controls, _
                                                                    arrLabelControlsOrder)

End Sub
