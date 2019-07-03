Attribute VB_Name = "Utility"
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
