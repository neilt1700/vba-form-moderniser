Attribute VB_Name = "VFMUtility"
Option Explicit

' This method from:
' https://stackoverflow.com/a/218727/1382318
Public Function KeyExistsInCollection(ByVal col As Collection, ByVal key As String) As Boolean
  Dim Var As Variant
  Dim errNumber As Long

  KeyExistsInCollection = False
  Set Var = Nothing

  Err.Clear
  On Error Resume Next
    Var = col.Item(key)
    errNumber = CLng(Err.Number)
  On Error GoTo 0

  ' 5 is not in, 0 and 438 represent in collection
  If errNumber = 5 Then ' it is 5 if not in collection
    KeyExistsInCollection = False
  Else
    KeyExistsInCollection = True
  End If

End Function

Public Function ControlExists(ByVal ctlsControls As Controls, ByVal stControlName As String) As Boolean
  On Error Resume Next
    ControlExists = Not ctlsControls(stControlName) Is Nothing
  On Error GoTo 0
End Function

' Placeholder function for error handling
' Note, general error handling is not implemented in this project, however the
' code to call it is mostly present. If you want to implement this error
' handling please refer to:
' "Professional Excel Development - Second Edition" - Rob Bovey,
' Dennis Wallentin, Stephen Bullen and John Green. 2009. Published by Addison
' Wesley.
Public Function bCentralErrorHandler(ByVal sModule As String, _
                                     ByVal sProc As String, _
                                     Optional ByVal sFile As String, _
                                     Optional ByVal bEntryPoint As Boolean, _
                                     Optional ByVal bReThrow As Boolean = True) As Boolean
  bCentralErrorHandler = False
End Function



