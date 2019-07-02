Attribute VB_Name = "FormModerniserModule"
' Copyright (c) Commtap CIC 2019
' Available under the MIT license: see the LICENSE file at the root of this
' project.
' Contact: tap@commtap.org

' Note font size is not settable programmatically it seems.

Option Explicit

Private Const msMODULE As String = "FormModerniserModule"

' Used for styling the label buttons.
Private m_stDefaultButton As String
Private m_stActiveButton As String

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
Public Const g_lngFONT_SIZE = 10 ' This has to be set manually on forms:
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

' Note these colours are the ones used in forms in the workspace for Word,
' PowerPoint and Excel (shades of blue).
' The active border style is slightly simplified.
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
  m_stDefaultButton = stValue
End Property

Public Property Get DefaultButton() As String
  DefaultButton = m_stDefaultButton
End Property

Public Property Let ActiveButton(ByVal stValue As String)
  m_stActiveButton = stValue
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

' TODO
' Apply Styles Something Like This:
Public Sub ApplyStyles(ByRef ctlControl As control, ByVal stControlType As String)

  Select Case stControlType
    Case "Label"
     ' ctlControl.BorderColor = whatever
  
  End Select


End Sub

Public Sub ModerniseForm(ByRef uUserForm As UserForm)

  uUserForm.ForeColor = g_lngFORE_COLOUR
  uUserForm.Font.Name = g_stFONT_NAME
  'uUserForm.Font.Size = g_lngFONT_SIZE
  uUserForm.BackColor = g_lngFORM_BACK_COLOUR
  uUserForm.BorderColor = g_lngBORDER_COLOUR
  uUserForm.SpecialEffect = g_lngSPECIAL_EFFECT

End Sub

Public Sub ModerniseControls(ByRef ctlsControls As Controls)

  ' TODO: at the moment this attempts to apply styles to everything regardless
  '       change this to target elements with styles they can accept - and then
  '       get rid of the on error resume next.

  Dim ctlControl As control
  On Error Resume Next
  For Each ctlControl In ctlsControls
    ctlControl.ForeColor = g_lngFORE_COLOUR
    ctlControl.Font.Name = g_stFONT_NAME
    ctlControl.Font.Size = g_lngFONT_SIZE
    ctlControl.BackColor = g_lngBACK_COLOUR
    ctlControl.BorderColor = g_lngBORDER_COLOUR
    ctlControl.SpecialEffect = g_lngSPECIAL_EFFECT
    
    If TypeName(ctlControl) = "TextBox" Then
      ctlControl.BorderStyle = g_lngTEXTBOX_BORDERSTYLE
    End If
    
    If TypeName(ctlControl) = "Frame" Then
      ctlControl.BorderStyle = g_lngTEXTBOX_BORDERSTYLE
    End If
    
  Next ctlControl

End Sub
