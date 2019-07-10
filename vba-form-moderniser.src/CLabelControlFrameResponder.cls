VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLabelControlFrameResponder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright (c) Commtap CIC 2019
' Available under the MIT license: see the LICENSE file at the root of this
' project.
' Contact: tap@commtap.org

Option Explicit

Private Const msMODULE As String = "CLabelControlFrameResponder"

' Responds to Frame move events - to reset label button states as necessary.

Public WithEvents FrameEvents As MSForms.Frame
Attribute FrameEvents.VB_VarHelpID = -1

Private p_oLabelControls As CLabelControls

Private Sub FrameEvents_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  Const sSOURCE As String = "FrameEvents_MouseMove"
  On Error GoTo ErrorHandler
  
  p_oLabelControls.UpdateControlButtonState
  
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


Public Sub InitiateProperties(ByRef ctlFrameControl As control, _
                              ByRef oLabelControls As CLabelControls)

  Const sSOURCE As String = "InitiateProperties"
  On Error GoTo ErrorHandler
  
  Set p_oLabelControls = oLabelControls
  Set Me.FrameEvents = ctlFrameControl
  
  Exit Sub

ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub
