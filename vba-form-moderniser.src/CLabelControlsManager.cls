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

Private Const msMODULE As String = "CLabelControlsManager"

Private p_oLabelControls As CLabelControls
Private p_ctlsControls As Controls
Private p_colControlResponders As Collection
Private p_colFrameControls As Collection

Private p_colKeyDownControls As Collection

Public Property Get LabelControls() As CLabelControls
  Set LabelControls = p_oLabelControls
End Property

Public Sub CreateControlResponders()

  Const sSOURCE As String = "CreateControlResponders"
  On Error GoTo ErrorHandler

  Set p_colControlResponders = New Collection

  Dim colLabelControls As Collection
  Set colLabelControls = p_oLabelControls.LabelControls
    
  Dim oLabelControl As CLabelControl
  For Each oLabelControl In colLabelControls
  
    p_colControlResponders.Add VFMFactory.CreateCLabelControlResponder(oLabelControl, _
                                         p_oLabelControls)
  
  Next oLabelControl

  Exit Sub

ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub

Public Sub CreateFrameControlResponders(Optional ByVal stNamePrefix As String)

  Const sSOURCE As String = "CreateFrameControls"
  On Error GoTo ErrorHandler

  Dim ctlControl As Control
  Set p_colFrameControls = New Collection

  For Each ctlControl In p_ctlsControls
    If TypeName(ctlControl) = "Frame" Then
      If Mid(ctlControl.Name, 1, Len(stNamePrefix)) = stNamePrefix Then
        p_colFrameControls.Add VFMFactory.CreateCLabelControlFrameResponder(ctlControl, p_oLabelControls), _
                                  ctlControl.Name
      End If
    End If
  Next ctlControl
  
  Exit Sub

ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub

Public Sub CreateKeyDownResponders()
 
  Const sSOURCE As String = "CreateKeyDownResponders"
  On Error GoTo ErrorHandler
 
  Dim ctlControl As Control
  
  FormModerniserModule.TabOverflow = 0
  FormModerniserModule.LastTabbedControl = vbNullString
  
  Set p_colKeyDownControls = New Collection
  
  For Each ctlControl In p_ctlsControls
    If TypeName(ctlControl) = "CheckBox" Or _
       TypeName(ctlControl) = "OptionButton" Or _
       TypeName(ctlControl) = "ScrollBar" Or _
       TypeName(ctlControl) = "SpinButton" Or _
       TypeName(ctlControl) = "ListBox" Or _
       TypeName(ctlControl) = "TabStrip" Or _
       TypeName(ctlControl) = "TextBox" Then
      p_colKeyDownControls.Add VFMFactory.CreateCKeyDownResponder(ctlControl, p_oLabelControls, p_ctlsControls), _
                                        ctlControl.Name
    End If
  Next ctlControl
  
  Exit Sub

ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Sub

' Checks that the default buttons are correct for this set of controls and
' resets them if not.
Public Sub CheckDefaults()
  
  If Not ControlExists(p_ctlsControls, FormModerniserModule.ActiveButton) Then
    FormModerniserModule.ActiveButton = vbNullString
  End If
  
  If Not ControlExists(p_ctlsControls, FormModerniserModule.DefaultButton) Then
    FormModerniserModule.DefaultButton = vbNullString
  End If
  
End Sub

Public Sub InitiateProperties(ByRef ctlsControls As MSForms.Controls, _
                              ByRef arrLabelControlsOrder() As String)

  Const sSOURCE As String = "InitiateProperties"
  On Error GoTo ErrorHandler
  
  Set p_ctlsControls = ctlsControls
  
  Set p_oLabelControls = VFMFactory.CreateCLabelControls(ctlsControls, arrLabelControlsOrder)
  CheckDefaults
  
  p_oLabelControls.UpdateControlButtonState
   
  Me.CreateControlResponders
  Me.CreateFrameControlResponders
  Me.CreateKeyDownResponders
    
  Exit Sub

ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub

