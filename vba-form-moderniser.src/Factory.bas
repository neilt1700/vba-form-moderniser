Attribute VB_Name = "Factory"
' Copyright (c) Commtap CIC 2019
' Available under the MIT license: see the LICENSE file at the root of this
' project.
' Contact: tap@commtap.org

Option Explicit

Private Const msMODULE As String = "Factory"

' Provides a way to initiate a new object with arguments. See:
' http://stackoverflow.com/questions/15224113/pass-arguments-to-constructor-in-vba

' Label Controls
Public Function CreateCLabelControl(ByRef ctlsUserFormControls As MSForms.Controls, _
                                    ByRef ctlLabelControl As MSForms.control, _
                                    Optional ByVal boolDefault As Boolean) As CLabelControl
  Set CreateCLabelControl = New CLabelControl
  CreateCLabelControl.InitiateProperties ctlsUserFormControls:=ctlsUserFormControls, _
                                         ctlLabelControl:=ctlLabelControl, _
                                         boolDefault:=boolDefault
End Function

Public Function CreateCLabelControlResponder(ByVal oLabelControl As CLabelControl, _
                                             ByRef oLabelControls As CLabelControls) As CLabelControlResponder
  Set CreateCLabelControlResponder = New CLabelControlResponder
  CreateCLabelControlResponder.InitiateProperties oLabelControl:=oLabelControl, _
                                                  oLabelControls:=oLabelControls
End Function

Public Function CreateCLabelControlFrameResponder(ByRef ctlFrameControl As control, _
                                                  ByRef oLabelControls As CLabelControls) As CLabelControlFrameResponder
  Set CreateCLabelControlFrameResponder = New CLabelControlFrameResponder
  CreateCLabelControlFrameResponder.InitiateProperties ctlFrameControl:=ctlFrameControl, _
                                                       oLabelControls:=oLabelControls
End Function

Public Function CreateCKeyDownResponder(ByRef ctlControl As control, _
                                        ByRef oLabelControls As CLabelControls, _
                                        ByRef ctlsControls As Controls) As CKeyDownResponder
  Set CreateCKeyDownResponder = New CKeyDownResponder
  CreateCKeyDownResponder.InitiateProperties ctlControl:=ctlControl, _
                                             oLabelControls:=oLabelControls, _
                                             ctlsControls:=ctlsControls
End Function

Public Function CreateCLabelControls(ByRef ctlsControls As MSForms.Controls, _
                                     ByVal stIdentifier As String) As CLabelControls
  Set CreateCLabelControls = New CLabelControls
  CreateCLabelControls.InitiateProperties ctlsControls:=ctlsControls, _
                                          stIdentifier:=stIdentifier
End Function

Public Function CreateCLabelControlsManager(ByRef ctlsControls As MSForms.Controls, _
                                            Optional ByVal stIdentifier As String) As CLabelControlsManager
  Set CreateCLabelControlsManager = New CLabelControlsManager
  CreateCLabelControlsManager.InitiateProperties ctlsControls:=ctlsControls, _
                                                 stIdentifier:=stIdentifier
End Function


' User Form Objects Creation
' ==========================

Public Function CreateUSampleUserForm() As USampleUserForm
  Set CreateUSampleUserForm = New USampleUserForm
  CreateUSampleUserForm.InitiateProperties
End Function

