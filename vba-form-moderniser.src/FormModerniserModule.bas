Attribute VB_Name = "FormModerniserModule"
' Copyright (c) Commtap CIC 2019
' Available under the MIT license: see the LICENSE file at the root of this
' project.
' Contact: tap@commtap.org

' Note buttons in office generally have a standard height and width.
' TODO: have a switch to switch off form modernisation.

Option Explicit

Private Const msMODULE As String = "FormModerniserModule"

Public Const g_stVERSION As String = "2.11-beta"

' Used for styling the label buttons.
Private m_stDefaultButton As String
Private m_stActiveButton As String

' Suffix for label control surface created from a command button:
Public Const g_stLABEL_CONTROL_SUFFIX As String = "VFMLabelControl_"

' Used to capture returns and tabbing from controls to the label buttons.
' Tab as in keyboard tab.
Private m_lngTabOverflow As Long
Private m_stLastTabbedControl As String

' Store reference userform
' Need one for each type of form in the project - early binding
' otherwise callbyname won't work.
Public gb_colCurrentUserForms As Collection

' General styling
Public Const g_lngFORE_COLOUR As Long = &H464646
Public Const g_stFONT_NAME As String = "Calibri"
Public Const g_lngFONT_SIZE = 10
Public Const g_lngFORM_BACK_COLOUR = &HE6E6E6
Public Const g_lngBACK_COLOUR = &HFFFFFF
Public Const g_lngBORDER_COLOUR As Long = &HA9A9A9
Public Const g_lngSPECIAL_EFFECT As Long = fmSpecialEffectFlat
Public Const g_lngTEXTBOX_BORDERSTYLE As Long = fmBorderStyleSingle

' Labels used as buttons specific styling
Public Const g_dblBTN_BORDER_WIDTH As Double = 1
Public Const g_dblBTN_DEFAULT_BORDER_WIDTH As Double = 2
Public Const g_dblBTN_DEFAULT_ACTIVE_BORDER_WIDTH As Double = 3

' These colours apply to the options pane in PowerPoint: these colours vary
' between Office products.
'Public Const g_lngBTN_ACTIVE_DEFAULT_BORDER_COLOUR As Long = &H565D71
'Public Const g_lngBTN_HOVER_DEFAULT_BORDER_COLOUR As Long = &H7E95C4
'Public Const g_lngBTN_HOVER_BORDER_COLOUR As Long = &H7E95C4
'Public Const g_lngBTN_INACTIVE_DEFAULT_BORDER_COLOUR As Long = &H3959DC
'Public Const g_lngBTN_INACTIVE_BORDER_COLOUR As Long = &HABABAB
'
'Public Const g_lngBTN_ACTIVE_DEFAULT_BACKGROUND_COLOUR As Long = &H9DBAF5
'Public Const g_lngBTN_HOVER_DEFAULT_BACKGROUND_COLOUR As Long = &HDCE4FC
'Public Const g_lngBTN_HOVER_BACKGROUND_COLOUR As Long = &HDCE4FC
'Public Const g_lngBTN_INACTIVE_DEFAULT_BACKGROUND_COLOUR As Long = &HFDFDFD
'Public Const g_lngBTN_INACTIVE_BACKGROUND_COLOUR As Long = &HFDFDFD

' Shades of blue
Public Const g_lngBTN_ACTIVE_DEFAULT_BORDER_COLOUR As Long = &H9E8671
Public Const g_lngBTN_HOVER_DEFAULT_BORDER_COLOUR As Long = &HD77800 ' Done
Public Const g_lngBTN_HOVER_BORDER_COLOUR As Long = &HD77800 ' Done
Public Const g_lngBTN_INACTIVE_DEFAULT_BORDER_COLOUR As Long = &HD77800 ' Done 2x width
Public Const g_lngBTN_INACTIVE_BORDER_COLOUR As Long = &HADADAD ' Done

Public Const g_lngBTN_ACTIVE_DEFAULT_BACKGROUND_COLOUR As Long = &HF7E4CC ' done 4x width
Public Const g_lngBTN_HOVER_DEFAULT_BACKGROUND_COLOUR As Long = &HFBF1E5 ' Done
Public Const g_lngBTN_HOVER_BACKGROUND_COLOUR As Long = &HFBF1E5 ' Done
Public Const g_lngBTN_INACTIVE_DEFAULT_BACKGROUND_COLOUR As Long = &HE1E1E1 ' done
Public Const g_lngBTN_INACTIVE_BACKGROUND_COLOUR As Long = &HE1E1E1 ' done

Public Enum lctlState
  lctlInactive
  lctlHover
  lctlActive
End Enum

Public Property Let DefaultButton(ByVal stValue As String)
  If Mid(stValue, 1, Len(g_stLABEL_CONTROL_SUFFIX)) = g_stLABEL_CONTROL_SUFFIX Then
    m_stDefaultButton = stValue
  Else
    m_stDefaultButton = g_stLABEL_CONTROL_SUFFIX & stValue
  End If
End Property

Public Property Get DefaultButton() As String
  DefaultButton = m_stDefaultButton
End Property

Public Property Let ActiveButton(ByVal stValue As String)
  If Mid(stValue, 1, Len(g_stLABEL_CONTROL_SUFFIX)) = g_stLABEL_CONTROL_SUFFIX Then
    m_stActiveButton = stValue
  Else
    m_stActiveButton = g_stLABEL_CONTROL_SUFFIX & stValue
  End If
End Property

Public Property Get ActiveButton() As String
  ActiveButton = m_stActiveButton
End Property

Public Property Let TabOverflow(ByVal stValue As Long)
  m_lngTabOverflow = stValue
End Property

Public Property Get TabOverflow() As Long
  TabOverflow = m_lngTabOverflow
End Property

Public Property Let LastTabbedControl(ByVal stValue As String)
  m_stLastTabbedControl = stValue
End Property

Public Property Get LastTabbedControl() As String
  LastTabbedControl = m_stLastTabbedControl
End Property


Public Sub ModerniseForm(ByRef uUserForm As UserForm)

  Const sSOURCE As String = "ModerniseForm"
  On Error GoTo ErrorHandler

  uUserForm.ForeColor = g_lngFORE_COLOUR
  uUserForm.Font.Name = g_stFONT_NAME
  uUserForm.Font.Size = g_lngFONT_SIZE
  uUserForm.BackColor = g_lngFORM_BACK_COLOUR
  uUserForm.BorderColor = g_lngBORDER_COLOUR
  uUserForm.SpecialEffect = g_lngSPECIAL_EFFECT

  Exit Sub
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub

Public Sub ModerniseControls(ByRef ctlsControls As Controls)

  Const sSOURCE As String = "ModerniseControls"
  On Error GoTo ErrorHandler

  Dim ctlControl As Control
  
  For Each ctlControl In ctlsControls
    With ctlControl
      ' General:
      .BackColor = g_lngBACK_COLOUR
    
      Select Case TypeName(ctlControl)
        Case "Label"
          .Font.Name = g_stFONT_NAME
          .Font.Size = g_lngFONT_SIZE
          .ForeColor = g_lngFORE_COLOUR
    
        Case "TextBox"
          .Font.Name = g_stFONT_NAME
          .Font.Size = g_lngFONT_SIZE
          .BorderStyle = g_lngTEXTBOX_BORDERSTYLE
          .BorderColor = g_lngBORDER_COLOUR
          .SpecialEffect = g_lngSPECIAL_EFFECT
          .ForeColor = g_lngFORE_COLOUR
      
        Case "Frame"
          .Font.Name = g_stFONT_NAME
          .Font.Size = g_lngFONT_SIZE
          .BorderStyle = g_lngTEXTBOX_BORDERSTYLE
          .BorderColor = g_lngBORDER_COLOUR
          .SpecialEffect = g_lngSPECIAL_EFFECT
          .ForeColor = g_lngFORE_COLOUR
   
        Case "CheckBox"
          .Font.Name = g_stFONT_NAME
          .Font.Size = g_lngFONT_SIZE
          .SpecialEffect = g_lngSPECIAL_EFFECT
          .ForeColor = g_lngFORE_COLOUR
   
        Case "OptionButton"
          .Font.Name = g_stFONT_NAME
          .Font.Size = g_lngFONT_SIZE
          .SpecialEffect = g_lngSPECIAL_EFFECT
          .ForeColor = g_lngFORE_COLOUR
      
        Case "ScrollBar"
          .ForeColor = g_lngFORE_COLOUR
      
        Case "SpinButton"
          .ForeColor = g_lngFORE_COLOUR
      
        Case "ListBox"
          .Font.Name = g_stFONT_NAME
          .Font.Size = g_lngFONT_SIZE
          .SpecialEffect = g_lngSPECIAL_EFFECT
          .BorderStyle = g_lngTEXTBOX_BORDERSTYLE
          .BorderColor = g_lngBORDER_COLOUR
       
      End Select
    End With
    
  Next ctlControl
  
  Exit Sub
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub

' Removes the label command suffix.
Public Function SourceButtonName(ByVal stButtonName As String) As String
  If Mid(stButtonName, 1, Len(g_stLABEL_CONTROL_SUFFIX)) = g_stLABEL_CONTROL_SUFFIX Then
    SourceButtonName = Mid(stButtonName, Len(g_stLABEL_CONTROL_SUFFIX) + 1)
  Else
    SourceButtonName = stButtonName
  End If
End Function

'Control utilities

Public Function ControlHasParent(ctlControl As Object) As Boolean
  
  On Error GoTo ErrorHandler
  
  Dim stParentType As String
  stParentType = TypeName(ctlControl.Parent)
  ControlHasParent = True
  
  Exit Function
  
ErrorHandler:
  ControlHasParent = False
End Function

Public Sub AbsolutePosition(ByRef ctlControl As Control, ByRef dblLeft As Double, ByRef dblTop As Double)

  Const sSOURCE As String = "AbsolutePosition"
  On Error GoTo ErrorHandler

  dblLeft = ctlControl.Left
  dblTop = ctlControl.Top
  
  AbsolutePositionRecurse ctlControl, dblLeft, dblTop
  
  Exit Sub
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Sub

Public Sub AbsolutePositionRecurse(ByRef ctlControl As Object, ByRef dblLeft As Double, ByRef dblTop As Double)

  Const sSOURCE As String = "AbsolutePositionRecurse"
  On Error GoTo ErrorHandler

  If ControlHasParent(ctlControl) Then
    dblLeft = dblLeft + ctlControl.Parent.Left
    dblTop = dblTop + ctlControl.Parent.Top
    AbsolutePositionRecurse ctlControl.Parent, dblLeft, dblTop
  End If
  
  Exit Sub
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Sub
