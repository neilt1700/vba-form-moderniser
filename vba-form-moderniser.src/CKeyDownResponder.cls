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

Private Const msMODULE As String = "CKeyDownResponder"

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

Private p_ctlControl As Control
Private p_oLabelControls As CLabelControls
Private p_ctlsControls As Controls

Private Sub OptionButtonEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

  Const sSOURCE As String = "OptionButtonEvents_KeyDown"
  On Error GoTo ErrorHandler

  MyEventsKeyDown KeyCode, Shift
  
ErrorExit:
  Exit Sub
    
ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If
  
End Sub

Private Sub OptionButtonEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

  Const sSOURCE As String = "OptionButtonEvents_MouseDown"
  On Error GoTo ErrorHandler

  MyEventsMouseDown Button, Shift, X, Y
  
ErrorExit:
  Exit Sub
    
ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If
  
End Sub

Private Sub ScrollBarEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

  Const sSOURCE As String = "ScrollBarEvents_KeyDown"
  On Error GoTo ErrorHandler

  MyEventsKeyDown KeyCode, Shift
  
ErrorExit:
  Exit Sub
    
ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If
  
End Sub

'No mousedown event for scrollbar events.
'Private Sub ScrollBarEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'  MyEventsMouseDown Button, Shift, X, Y
'End Sub

Private Sub SpinButtonEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

  Const sSOURCE As String = "SpinButtonEvents_KeyDown"
  On Error GoTo ErrorHandler

  MyEventsKeyDown KeyCode, Shift
  
ErrorExit:
  Exit Sub
    
ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If
  
End Sub

'No mousedown event for spinbutton events.
'Private Sub SpinButtonEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'  MyEventsMouseDown Button, Shift, X, Y
'End Sub

Private Sub ListBoxEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

  Const sSOURCE As String = "ListBoxEvents_KeyDown"
  On Error GoTo ErrorHandler

  MyEventsKeyDown KeyCode, Shift
  
ErrorExit:
  Exit Sub
    
ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If
  
End Sub

Private Sub ListBoxEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

  Const sSOURCE As String = "ListBoxEvents_MouseDown"
  On Error GoTo ErrorHandler

  MyEventsMouseDown Button, Shift, X, Y
  
ErrorExit:
  Exit Sub
    
ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If
  
End Sub

Private Sub TabStripEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

  Const sSOURCE As String = "TabStripEvents_KeyDown"
  On Error GoTo ErrorHandler

  MyEventsKeyDown KeyCode, Shift
  
ErrorExit:
  Exit Sub
    
ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If
  
End Sub

Private Sub TabStripEvents_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

  Const sSOURCE As String = "TabStripEvents_MouseDown"
  On Error GoTo ErrorHandler

  MyEventsMouseDown Button, Shift, X, Y
  
ErrorExit:
  Exit Sub
    
ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If
  
End Sub

Private Sub CheckBoxEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

  Const sSOURCE As String = "CheckBoxEvents_KeyDown"
  On Error GoTo ErrorHandler

  MyEventsKeyDown KeyCode, Shift
  
ErrorExit:
  Exit Sub
    
ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If
  
End Sub

Private Sub CheckBoxEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

  Const sSOURCE As String = "CheckBoxEvents_MouseDown"
  On Error GoTo ErrorHandler

  MyEventsMouseDown Button, Shift, X, Y
  
ErrorExit:
  Exit Sub
    
ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If
  
End Sub

Private Sub TextBoxEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

  Const sSOURCE As String = "TextBoxEvents_KeyDown"
  On Error GoTo ErrorHandler

  MyEventsKeyDown KeyCode, Shift
  
ErrorExit:
  Exit Sub
    
ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If
  
End Sub

Private Sub TextBoxEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

  Const sSOURCE As String = "TextBoxEvents_MouseDown"
  On Error GoTo ErrorHandler

  MyEventsMouseDown Button, Shift, X, Y
  
ErrorExit:
  Exit Sub
    
ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If
  
End Sub

Private Sub MyEventsKeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

  Const sSOURCE As String = "MyEventsKeyDown"
  On Error GoTo ErrorHandler

  If p_oLabelControls.LabelControls.Count > 0 Then
    If p_ctlControl.Name = FormModerniserModule.LastTabbedControl Then
      If KeyCode = vbKeyReturn Then
        If FormModerniserModule.DefaultButton <> vbNullString Then
          CallByName gb_colCurrentUserForms.Item(1), SourceButtonName(FormModerniserModule.DefaultButton) & "_Click", VbMethod
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
          CallByName gb_colCurrentUserForms.Item(1), SourceButtonName(FormModerniserModule.DefaultButton) & "_Click", VbMethod
        End If
      End If
      If FormModerniserModule.TabOverflow > 0 Then
        p_ctlsControls(FormModerniserModule.LastTabbedControl).SetFocus
      End If
    End If
  End If
  
  Exit Sub
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Sub

Private Sub MyEventsMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

  Const sSOURCE As String = "MyEventsMouseDown"
  On Error GoTo ErrorHandler

  FormModerniserModule.TabOverflow = 0
  
  Exit Sub
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Sub


Public Sub InitiateProperties(ByRef ctlControl As Control, _
                              ByRef oLabelControls As CLabelControls, _
                              ByRef ctlsControls As Controls)

  Const sSOURCE As String = "InitiateProperties"
  On Error GoTo ErrorHandler
  
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
    
  Exit Sub

ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub
