VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLabelControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright (c) Commtap CIC 2019
' Available under the MIT license: see the LICENSE file at the root of this
' project.
' Contact: tap@commtap.org

Private Const msMODULE As String = "CLabelControl"

Option Explicit

Public WithEvents LabelControlEvents As MSForms.Label
Attribute LabelControlEvents.VB_VarHelpID = -1

Private p_ctlSurface As MSForms.Control
Private p_ctlsUserFormControls As MSForms.Controls

Private p_ctlText As MSForms.Control
Private p_ctlBackground As MSForms.Control
Private p_ctlBorder As MSForms.Control

Private p_stName As String

Private p_boolDefault As Boolean
Private p_boolActive As Boolean

Private p_stOnClickMethod As String

Private Const p_stTEXTLAYER As String = "Caption"
Private Const p_stBACKGROUNDLAYER As String = "Background"
Private Const p_stBORDERLAYER As String = "Border"

Public Property Get LabelName() As String
  LabelName = p_stName
End Property

Public Property Get ControlSurface() As MSForms.Control
  Set ControlSurface = p_ctlSurface
End Property

Private Sub LabelControlEvents_Click()
  CallByName gb_colCurrentUserForms.Item(1), p_stOnClickMethod, VbMethod
End Sub

' The control consists of (from top to bottom):
' - Control Face Layer (where events are captured) - the original label but
'   transparent and without any text.
' - Text layer - containing the orginal text - centred.
' - Fill layer - colour modifies according to the control state - slightly
'   reduced in height and width compared to the top and bottom layer.
' - Border layer.
'
Public Sub CreateLayers()
  
  Const sSOURCE As String = "CreateLayers"
  On Error GoTo ErrorHandler
  
  Dim stCaption As String
  Dim dblTop As Double
  Dim dblLeft As Double
  Dim dblWidth As Double
  Dim dblHeight As Double
  
  Dim dblTextHeight As Double
  
  With p_ctlSurface
    p_stName = .Name
    stCaption = .Caption
    dblTop = .Top
    dblLeft = .Left
    dblWidth = .Width
    dblHeight = .Height
  End With
      
  ' Border layer
  Set p_ctlBorder = p_ctlsUserFormControls.Add("Forms.Label.1", p_stName & "_" & p_stBORDERLAYER, True)
  With p_ctlBorder
    .Caption = vbNullString
    .BackStyle = fmBackStyleOpaque
    .BorderStyle = fmBorderStyleNone
    
    If p_boolDefault Then
      .BackColor = g_lngBTN_INACTIVE_DEFAULT_BORDER_COLOUR
    Else
      .BackColor = g_lngBTN_INACTIVE_BORDER_COLOUR
    End If
      
    .Top = dblTop
    .Left = dblLeft
    .Width = dblWidth
    .Height = dblHeight
    .ZOrder 0
  End With
  
  ' Background
  Set p_ctlBackground = p_ctlsUserFormControls.Add("Forms.Label.1", p_stName & "_" & p_stBACKGROUNDLAYER, True)
  With p_ctlBackground
    .Caption = vbNullString
    .BackStyle = fmBackStyleOpaque
    .BorderStyle = fmBorderStyleNone
    
    If p_boolDefault Then
      .BackColor = g_lngBTN_INACTIVE_DEFAULT_BACKGROUND_COLOUR
    Else
      .BackColor = g_lngBTN_INACTIVE_BACKGROUND_COLOUR
    End If
    .ZOrder 0
      
  End With
  
  Me.SetBorderWidth
  
  ' Caption - needs to stay just below control face.
  Set p_ctlText = p_ctlsUserFormControls.Add("Forms.Label.1", p_stName & "_" & p_stTEXTLAYER, True)
  With p_ctlText
    ' 1.35 appears to be the best factor to use (takes into account the padding
    ' added around text.
    .Font.Size = g_lngFONT_SIZE
    dblTextHeight = CDbl(.Font.Size) * 1.35
    .Caption = stCaption
    .BackStyle = fmBackStyleTransparent
    .BorderStyle = fmBorderStyleNone
    .ForeColor = g_lngFORE_COLOUR
    .Font.Name = "Calibri"
    .Top = (dblTop + (dblHeight / 2)) - (dblTextHeight / 2)
    .Left = dblLeft
    .Width = dblWidth
    .Height = dblTextHeight
    .TextAlign = fmTextAlignCenter
    .ZOrder 0
  End With
  
  ' The original item is the "control surface" and should stay
  ' on top.
  With p_ctlSurface
    .Caption = vbNullString
    .BackStyle = fmBackStyleTransparent
    .BorderStyle = fmBorderStyleNone
    .ZOrder 0
  End With
      
  Exit Sub
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
      
End Sub


Public Sub SetState(ByVal boolDefault As Boolean, _
                    ByVal lsState As lctlState, _
                    Optional ByVal boolLocationSet As Boolean = False, _
                    Optional ByVal X As Double, _
                    Optional ByVal Y As Double)

  Const sSOURCE As String = "SetState"
  On Error GoTo ErrorHandler
  
  p_boolDefault = boolDefault
  p_boolActive = False
   
  ' Check if mouse is over the button:
  If boolLocationSet Then
    If X < 0 Or X > p_ctlSurface.Width Or _
       Y < 0 Or Y > p_ctlSurface.Height Then
       lsState = lctlInactive
    End If
  End If

  Select Case lsState
    Case lctlInactive
      If boolDefault Then
        p_ctlBackground.BackColor = g_lngBTN_INACTIVE_DEFAULT_BACKGROUND_COLOUR
        p_ctlBorder.BackColor = g_lngBTN_INACTIVE_DEFAULT_BORDER_COLOUR
      Else
        p_ctlBackground.BackColor = g_lngBTN_INACTIVE_BACKGROUND_COLOUR
        p_ctlBorder.BackColor = g_lngBTN_INACTIVE_BORDER_COLOUR
      End If
      
    Case lctlHover
      If boolDefault Then
        p_ctlBackground.BackColor = g_lngBTN_HOVER_DEFAULT_BACKGROUND_COLOUR
        p_ctlBorder.BackColor = g_lngBTN_HOVER_DEFAULT_BORDER_COLOUR
      Else
        p_ctlBackground.BackColor = g_lngBTN_HOVER_BACKGROUND_COLOUR
        p_ctlBorder.BackColor = g_lngBTN_HOVER_BORDER_COLOUR
      End If
      
    Case lctlActive
      ' Always default
      p_ctlBackground.BackColor = g_lngBTN_ACTIVE_DEFAULT_BACKGROUND_COLOUR
      p_ctlBorder.BackColor = g_lngBTN_ACTIVE_DEFAULT_BORDER_COLOUR
      p_boolActive = True

  End Select
    
  Me.SetBorderWidth
                    
  Exit Sub
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
                    
End Sub

Public Sub SetBorderWidth()
  
  Const sSOURCE As String = "SetBorderWidth"
  On Error GoTo ErrorHandler
  
  Dim dblBorderWidth As Double
  
  Dim dblTop As Double
  Dim dblLeft As Double
  Dim dblWidth As Double
  Dim dblHeight As Double
  
  With p_ctlSurface
    dblTop = .Top
    dblLeft = .Left
    dblWidth = .Width
    dblHeight = .Height
  End With
  
  Set p_ctlBackground = p_ctlsUserFormControls(p_stName & "_" & p_stBACKGROUNDLAYER)
  With p_ctlBackground
    If p_boolDefault Then
      If p_boolActive Then
        dblBorderWidth = g_dblBTN_DEFAULT_ACTIVE_BORDER_WIDTH
      Else
        dblBorderWidth = g_dblBTN_DEFAULT_BORDER_WIDTH
      End If
    Else
      dblBorderWidth = g_dblBTN_BORDER_WIDTH
    End If
      
    .Top = dblTop + dblBorderWidth
    .Left = dblLeft + dblBorderWidth
    .Width = dblWidth - (dblBorderWidth * 2)
    .Height = dblHeight - (dblBorderWidth * 2)
  End With

  Exit Sub
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub

Private Function CreateControlSurface(ByRef ctlCommandButton As Control) As Control

  Const sSOURCE As String = "CreateControlSurface"
  On Error GoTo ErrorHandler

  Dim stCaption As String
  Dim stName As String
  
  Dim dblTop As Double
  Dim dblLeft As Double
  Dim dblWidth As Double
  Dim dblHeight As Double
  
  Dim ctlSurface As Control
  
  AbsolutePosition ctlCommandButton, dblLeft, dblTop

  With ctlCommandButton
    stName = g_stLABEL_CONTROL_SUFFIX & .Name
    stCaption = .Caption
    'dblTop = .Top
    'dblLeft = .Left
    dblWidth = .Width
    dblHeight = .Height
  End With

  Set ctlSurface = p_ctlsUserFormControls.Add("Forms.Label.1", stName, True)
  With ctlSurface
    .Caption = stCaption
    .Top = dblTop
    .Left = dblLeft
    .Width = dblWidth
    .Height = dblHeight
  End With
  
  Set CreateControlSurface = ctlSurface
  
  Exit Function
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Function



Public Sub InitiateProperties(ByRef ctlsUserFormControls As MSForms.Controls, _
                              ByRef ctlCommandButton As MSForms.Control, _
                              Optional ByVal boolDefault As Boolean = False)

  Const sSOURCE As String = "InitiateProperties"
  On Error GoTo ErrorHandler

  Set p_ctlsUserFormControls = ctlsUserFormControls
  
  Set p_ctlSurface = CreateControlSurface(ctlCommandButton)
  Set Me.LabelControlEvents = p_ctlSurface
  
  p_stOnClickMethod = ctlCommandButton.Name & "_Click"
  
  p_ctlsUserFormControls(ctlCommandButton.Name).Visible = False
  
  p_boolDefault = boolDefault
  
  Me.CreateLayers

  Exit Sub
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub
