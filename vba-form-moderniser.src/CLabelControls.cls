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

Private p_colLabelControls As Collection
Private p_arrLabelControlsOrder() As String
Private p_ctlsControls As Controls

Public Property Get LabelControls() As Collection
  Set LabelControls = p_colLabelControls
End Property

Public Sub CreateLayeredControls(Optional ByVal stNamePrefix As String)

  Dim ctlControl As control
  Dim colLabelControls As Collection
  Set colLabelControls = New Collection

  For Each ctlControl In p_ctlsControls
    If TypeName(ctlControl) = "Label" Then
      If Mid(ctlControl.Name, 1, Len(stNamePrefix)) = stNamePrefix Then
        colLabelControls.Add Factory.CreateCLabelControl(p_ctlsControls, ctlControl), ctlControl.Name
      End If
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

End Sub

Public Sub UpdateControlButtonState(Optional ByVal stControlName As String, _
                                    Optional ByVal lsState As lctlState = lctlInactive, _
                                    Optional ByVal boolLocationSet As Boolean = False, _
                                    Optional ByVal X As Double, _
                                    Optional ByVal Y As Double)

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

End Sub

Public Sub InitiateProperties(ByRef ctlsControls As MSForms.Controls, _
                              ByVal stIdentifier As String, _
                              ByRef arrLabelControlsOrder() As String)


  p_arrLabelControlsOrder = arrLabelControlsOrder
  
  Set p_ctlsControls = ctlsControls
  CreateLayeredControls stIdentifier

End Sub
