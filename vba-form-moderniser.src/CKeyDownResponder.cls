VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CKeyDownResponder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright (c) Commtap CIC 2019
' Available under the MIT license: see the LICENSE file at the root of this
' project.
' Contact: tap@commtap.org

Option Explicit

' Responds to Frame move events - to reset label button states as necessary.

' TODO: completion of adding these events.

'Public WithEvents OptionButtonEvents As MSForms.OptionButton
'Public WithEvents ScrollBarEvents As MSForms.ScrollBar
'Public WithEvents ScrollBarEvents As MSForms.SpinButton

'Public WithEvents ScrollBarEvents As MSForms.ListBox
Public WithEvents TextBoxEvents As MSForms.TextBox
Attribute TextBoxEvents.VB_VarHelpID = -1
Public WithEvents CheckBoxEvents As MSForms.CheckBox
Attribute CheckBoxEvents.VB_VarHelpID = -1

Private p_ctlControl As control
Private p_oLabelControls As CLabelControls
Private p_ctlsControls As Controls

Private Sub TextBoxEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  MyEventsKeyDown KeyCode, Shift
End Sub

Private Sub TextBoxEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  MyEventsMouseDown Button, Shift, X, Y
End Sub

Private Sub CheckBoxEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  MyEventsKeyDown KeyCode, Shift
End Sub

Private Sub CheckBoxEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  MyEventsMouseDown Button, Shift, X, Y
End Sub


Private Sub MyEventsKeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  If p_oLabelControls.LabelControls.Count > 0 Then
    If p_ctlControl.Name = FormModerniserModule.LastTabbedControl Then
      If KeyCode = vbKeyReturn Then
        If FormModerniserModule.DefaultButton <> vbNullString Then
          CallByName gb_colCurrentUserForms.Item(1), FormModerniserModule.DefaultButton & "_Click", VbMethod
        End If
      ElseIf KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
        If FormModerniserModule.TabOverflow < p_oLabelControls.LabelControls.Count Then
          FormModerniserModule.TabOverflow = FormModerniserModule.TabOverflow + 1
          FormModerniserModule.DefaultButton = p_oLabelControls.LabelControls.Item(FormModerniserModule.TabOverflow).LabelName
        End If
        p_oLabelControls.UpdateControlButtonState FormModerniserModule.DefaultButton
      ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
        If FormModerniserModule.TabOverflow > 0 Then
          FormModerniserModule.TabOverflow = FormModerniserModule.TabOverflow - 1
          If FormModerniserModule.TabOverflow > 0 Then
            FormModerniserModule.DefaultButton = p_oLabelControls.LabelControls.Item(FormModerniserModule.TabOverflow).LabelName
          End If
        End If
        p_oLabelControls.UpdateControlButtonState FormModerniserModule.DefaultButton
      End If
    Else
      If KeyCode = vbKeyReturn Then
        If FormModerniserModule.DefaultButton <> vbNullString Then
          CallByName gb_colCurrentUserForms.Item(1), FormModerniserModule.DefaultButton & "_Click", VbMethod
        End If
      End If
      If FormModerniserModule.TabOverflow > 0 Then
        p_ctlsControls(FormModerniserModule.LastTabbedControl).SetFocus
      End If
    End If
  End If
  
End Sub

Private Sub MyEventsMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  FormModerniserModule.TabOverflow = 0
End Sub


Public Sub InitiateProperties(ByRef ctlControl As control, _
                              ByRef oLabelControls As CLabelControls, _
                              ByRef ctlsControls As Controls)

  Set p_ctlControl = ctlControl
  Set p_oLabelControls = oLabelControls
  Set p_ctlsControls = ctlsControls
  
  If TypeName(ctlControl) = "TextBox" Then
    Set Me.TextBoxEvents = ctlControl
    FormModerniserModule.LastTabbedControl = ctlControl.Name
  End If

End Sub



