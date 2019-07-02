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

' Responds to events on behalf of a CLabelControl.
' For adding additional control events onto the control surface.
' Only the click event for the control surface (label) should be defined on the
' form build UI initially.

Public WithEvents LabelEvents As MSForms.Label
Attribute LabelEvents.VB_VarHelpID = -1

Private p_stButtonName As String
Private p_oLabelControls As CLabelControls
Private p_oLabelControl As CLabelControl

Private Sub LabelEvents_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  If FormModerniserModule.ActiveButton = p_stButtonName Then
    p_oLabelControls.UpdateControlButtonState p_stButtonName, lctlActive, True, X, Y
  Else
    p_oLabelControls.UpdateControlButtonState p_stButtonName, lctlHover, True, X, Y
  End If
    
End Sub

Private Sub LabelEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   
  FormModerniserModule.DefaultButton = p_stButtonName
  FormModerniserModule.ActiveButton = p_stButtonName
  p_oLabelControls.UpdateControlButtonState p_stButtonName, lctlActive
  
End Sub

Private Sub LabelEvents_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

  p_oLabelControls.UpdateControlButtonState p_stButtonName, lctlHover, True, X, Y
  FormModerniserModule.ActiveButton = vbNullString

End Sub


Public Sub InitiateProperties(ByVal oLabelControl As CLabelControl, _
                              ByRef oLabelControls As CLabelControls)
  
  Set p_oLabelControls = oLabelControls
  Set p_oLabelControl = oLabelControl
  
  p_stButtonName = oLabelControl.LabelName
    
  Set Me.LabelEvents = oLabelControl.ControlSurface

End Sub


