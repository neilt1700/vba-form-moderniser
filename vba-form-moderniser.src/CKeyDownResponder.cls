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

Public WithEvents OptionButtonEvents As MSForms.OptionButton
Attribute OptionButtonEvents.VB_VarHelpID = -1
Public WithEvents ScrollBarEvents As MSForms.ScrollBar
Attribute ScrollBarEvents.VB_VarHelpID = -1
Public WithEvents SpinButtonEvents As MSForms.SpinButton
Attribute SpinButtonEvents.VB_VarHelpID = -1
Public WithEvents ListBoxEvents As MSForms.ListBox
Attribute ListBoxEvents.VB_VarHelpID = -1
Public WithEvents TabStripEvents As MSForms.TabStrip
Attribute TabStripEvents.VB_VarHelpID = -1
Public WithEvents CheckBoxEvents As MSForms.CheckBox
Attribute CheckBoxEvents.VB_VarHelpID = -1
Public WithEvents TextBoxEvents As MSForms.TextBox
Attribute TextBoxEvents.VB_VarHelpID = -1

Private p_ctlControl As control
Private p_oLabelControls As CLabelControls
Private p_ctlsControls As Controls

Private Sub OptionButtonEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  MyEventsKeyDown KeyCode, Shift
End Sub

Private Sub OptionButtonEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  MyEventsMouseDown Button, Shift, X, Y
End Sub

Private Sub ScrollBarEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  MyEventsKeyDown KeyCode, Shift
End Sub

'No mousedown event for scrollbar events.
'Private Sub ScrollBarEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'  MyEventsMouseDown Button, Shift, X, Y
'End Sub

Private Sub SpinButtonEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  MyEventsKeyDown KeyCode, Shift
End Sub

'No mousedown event for spinbutton events.
'Private Sub SpinButtonEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'  MyEventsMouseDown Button, Shift, X, Y
'End Sub

Private Sub ListBoxEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  MyEventsKeyDown KeyCode, Shift
End Sub

Private Sub ListBoxEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  MyEventsMouseDown Button, Shift, X, Y
End Sub

Private Sub TabStripEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  MyEventsKeyDown KeyCode, Shift
End Sub

Private Sub TabStripEvents_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  MyEventsMouseDown Button, Shift, X, Y
End Sub

Private Sub CheckBoxEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  MyEventsKeyDown KeyCode, Shift
End Sub

Private Sub CheckBoxEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  MyEventsMouseDown Button, Shift, X, Y
End Sub

Private Sub TextBoxEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  MyEventsKeyDown KeyCode, Shift
End Sub

Private Sub TextBoxEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
  
  Select Case TypeName(ctlControl)
    Case "TextBox"
      Set Me.TextBoxEvents = ctlControl
      FormModerniserModule.LastTabbedControl = ctlControl.Name
    Case "OptionButton"
      Set Me.OptionButtonEvents = ctlControl
      FormModerniserModule.LastTabbedControl = ctlControl.Name
    Case "ScrollBar"
      Set Me.ScrollBarEvents = ctlControl
      FormModerniserModule.LastTabbedControl = ctlControl.Name
    Case "SpinButton"
      Set Me.SpinButtonEvents = ctlControl
      FormModerniserModule.LastTabbedControl = ctlControl.Name
    Case "ListBox"
      Set Me.ListBoxEvents = ctlControl
      FormModerniserModule.LastTabbedControl = ctlControl.Name
    Case "TabStrip"
      Set Me.TabStripEvents = ctlControl
      FormModerniserModule.LastTabbedControl = ctlControl.Name
    Case "CheckBox"
      Set Me.CheckBoxEvents = ctlControl
      FormModerniserModule.LastTabbedControl = ctlControl.Name
  End Select
  
End Sub


