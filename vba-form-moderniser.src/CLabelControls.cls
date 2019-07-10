VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLabelControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright (c) Commtap CIC 2019
' Available under the MIT license: see the LICENSE file at the root of this
' project.
' Contact: tap@commtap.org

Option Explicit

Private Const msMODULE As String = "CLabelControls"

Private p_colLabelControls As Collection
Private p_arrLabelControlsOrder() As String
Private p_ctlsControls As Controls

Public Property Get LabelControls() As Collection
  Set LabelControls = p_colLabelControls
End Property

Public Sub CreateLayeredControls()

  Const sSOURCE As String = "CreateLayeredControls"
  On Error GoTo ErrorHandler

  Dim ctlControl As control
  
  Dim colLabelControls As Collection
  Set colLabelControls = New Collection

  For Each ctlControl In p_ctlsControls
    If TypeName(ctlControl) = "CommandButton" Then
      colLabelControls.Add VFMFactory.CreateCLabelControl(p_ctlsControls, ctlControl), ctlControl.Name
    End If
  Next ctlControl
  
  Set p_colLabelControls = New Collection
  Dim stLabelControl As Variant
  
  For Each stLabelControl In p_arrLabelControlsOrder
    If KeyExistsInCollection(colLabelControls, stLabelControl) Then
      p_colLabelControls.Add colLabelControls(stLabelControl), stLabelControl
      colLabelControls.Remove stLabelControl
    End If
  Next stLabelControl
  
  ' Add on any remaining items that weren't in the label control order list:
  Dim oLabelControl As CLabelControl
  For Each oLabelControl In colLabelControls
    p_colLabelControls.Add oLabelControl, oLabelControl.LabelName
  Next oLabelControl
  
  Exit Sub

ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub

Public Sub UpdateControlButtonState(Optional ByVal stControlName As String, _
                                    Optional ByVal lsState As lctlState = lctlInactive, _
                                    Optional ByVal boolLocationSet As Boolean = False, _
                                    Optional ByVal X As Double, _
                                    Optional ByVal Y As Double)

  Const sSOURCE As String = "UpdateControlButtonState"
  On Error GoTo ErrorHandler

  Dim stDefaultControlName As String
  stDefaultControlName = FormModerniserModule.DefaultButton

  Dim oLabelControl As CLabelControl
  
  For Each oLabelControl In p_colLabelControls
    If oLabelControl.LabelName = stControlName Then
      If oLabelControl.LabelName = stDefaultControlName Then
        oLabelControl.SetState True, lsState, boolLocationSet, X, Y
      Else
        oLabelControl.SetState False, lsState, boolLocationSet, X, Y
      End If
    Else
      If oLabelControl.LabelName = stDefaultControlName Then
        oLabelControl.SetState True, lctlInactive, boolLocationSet, X, Y
      Else
        oLabelControl.SetState False, lctlInactive, boolLocationSet, X, Y
      End If
    End If
  Next oLabelControl

  Exit Sub

ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub


Public Sub InitiateProperties(ByRef ctlsControls As MSForms.Controls, _
                              ByRef arrLabelControlsOrder() As String)

  Const sSOURCE As String = "InitiateProperties"
  On Error GoTo ErrorHandler
  
  p_arrLabelControlsOrder = arrLabelControlsOrder
  
  Set p_ctlsControls = ctlsControls
  CreateLayeredControls

  Exit Sub

ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub
