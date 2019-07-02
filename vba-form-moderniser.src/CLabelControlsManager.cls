VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLabelControlsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright (c) Commtap CIC 2019
' Available under the MIT license: see the LICENSE file at the root of this
' project.
' Contact: tap@commtap.org

Option Explicit

Private p_oLabelControls As CLabelControls
Private p_ctlsControls As Controls
Private p_colControlResponders As Collection
Private p_colFrameControls As Collection

Private p_colKeyDownControls As Collection

Public Property Get LabelControls() As CLabelControls
  Set LabelControls = p_oLabelControls
End Property

Public Sub CreateControlResponders()

  Set p_colControlResponders = New Collection

  Dim colLabelControls As Collection
  Set colLabelControls = p_oLabelControls.LabelControls
  
  FormModerniserModule.DefaultButton = vbNullString
  FormModerniserModule.ActiveButton = vbNullString
  
  Dim oLabelControl As CLabelControl
  For Each oLabelControl In colLabelControls
  
    p_colControlResponders.Add Factory.CreateCLabelControlResponder(oLabelControl, _
                                         p_oLabelControls)
  
  Next oLabelControl

End Sub

Public Sub CreateFrameControlResponders(Optional ByVal stNamePrefix As String)

  Dim ctlControl As control
  Set p_colFrameControls = New Collection

  For Each ctlControl In p_ctlsControls
    If TypeName(ctlControl) = "Frame" Then
      If Mid(ctlControl.Name, 1, Len(stNamePrefix)) = stNamePrefix Then
        p_colFrameControls.Add Factory.CreateCLabelControlFrameResponder(ctlControl, p_oLabelControls), _
                                  ctlControl.Name
      End If
    End If
  Next ctlControl
  
End Sub

Public Sub CreateKeyDownResponders()

  Dim ctlControl As control
  
  FormModerniserModule.TabOverflow = 0
  FormModerniserModule.LastTabbedControl = vbNullString
  
  Set p_colKeyDownControls = New Collection
  
  For Each ctlControl In p_ctlsControls
    If TypeName(ctlControl) = "TextBox" Then
      p_colKeyDownControls.Add Factory.CreateCKeyDownResponder(ctlControl, p_oLabelControls, p_ctlsControls), _
                                        ctlControl.Name
    End If
  Next ctlControl
  
End Sub


Public Sub InitiateProperties(ByRef ctlsControls As MSForms.Controls, _
                              Optional ByVal stIdentifier As String = vbNullString)
  
  Set p_ctlsControls = ctlsControls
  
  Set p_oLabelControls = Factory.CreateCLabelControls(ctlsControls, stIdentifier)
  p_oLabelControls.UpdateControlButtonState
   
  Me.CreateControlResponders
  Me.CreateFrameControlResponders
  Me.CreateKeyDownResponders

End Sub

