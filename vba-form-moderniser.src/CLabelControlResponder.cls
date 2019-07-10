VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLabelControlResponder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright (c) Commtap CIC 2019
' Available under the MIT license: see the LICENSE file at the root of this
' project.
' Contact: tap@commtap.org

Option Explicit

Private Const msMODULE As String = "CLabelControlResponder"

' Responds to events on behalf of a CLabelControl.
' For adding additional control events onto the control surface.

Public WithEvents LabelEvents As MSForms.Label
Attribute LabelEvents.VB_VarHelpID = -1

Private p_stButtonName As String
Private p_oLabelControls As CLabelControls
Private p_oLabelControl As CLabelControl

Private Sub LabelEvents_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  Const sSOURCE As String = "LabelEvents_MouseMove"
  On Error GoTo ErrorHandler
  
  If FormModerniserModule.ActiveButton = p_stButtonName Then
    p_oLabelControls.UpdateControlButtonState p_stButtonName, lctlActive, True, X, Y
  Else
    p_oLabelControls.UpdateControlButtonState p_stButtonName, lctlHover, True, X, Y
  End If
  
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

Private Sub LabelEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  Const sSOURCE As String = "LabelEvents_MouseDown"
  On Error GoTo ErrorHandler
  
  FormModerniserModule.DefaultButton = p_stButtonName
  FormModerniserModule.ActiveButton = p_stButtonName
  p_oLabelControls.UpdateControlButtonState p_stButtonName, lctlActive
  
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

Private Sub LabelEvents_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  Const sSOURCE As String = "LabelEvents_MouseUp"
  On Error GoTo ErrorHandler
  
  p_oLabelControls.UpdateControlButtonState p_stButtonName, lctlHover, True, X, Y
  FormModerniserModule.ActiveButton = vbNullString
  
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


Public Sub InitiateProperties(ByVal oLabelControl As CLabelControl, _
                              ByRef oLabelControls As CLabelControls)

  Const sSOURCE As String = "InitiateProperties"
  On Error GoTo ErrorHandler
  
  Set p_oLabelControls = oLabelControls
  Set p_oLabelControl = oLabelControl
  
  p_stButtonName = oLabelControl.LabelName
    
  Set Me.LabelEvents = oLabelControl.ControlSurface
  
  Exit Sub

ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub

