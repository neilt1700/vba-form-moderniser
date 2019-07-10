Attribute VB_Name = "FormModerniserDevPPT"
Option Explicit

' This module should be loaded in PowerPoint only.

' Get the VBProject
Public Function VFM_ImportModule(ByVal stModulePath As String) As Object
  Set VFM_ImportModule = Application.ActivePresentation.VBProject.VBComponents.Import(stModulePath)
End Function

Public Function VFM_RemoveModule(ByVal stModuleName As String)
  With Application.ActivePresentation.VBProject
    On Error Resume Next
      .VBComponents.Remove .VBComponents(stModuleName)
    On Error GoTo 0
  End With
End Function

Public Function VFM_ExportModules(ByVal stModuleNames As String, ByVal stFolderPath As String)

  Const vbext_ct_StdModule = 1
  Const vbext_ct_ClassModule = 2
  Const vbext_ct_MSForm = 3
  
  Dim cmpComponent
  Dim stFileName As String

  stModuleNames = " " & stModuleNames & " "

  With Application.ActivePresentation.VBProject
    For Each cmpComponent In .VBComponents
      If InStr(stModuleNames, " " & cmpComponent.Name & " ") Then
        stFileName = vbNullString
        Select Case .VBComponents(cmpComponent.Name).Type
          Case vbext_ct_ClassModule
            stFileName = cmpComponent.Name & ".cls"
          Case vbext_ct_MSForm
            stFileName = cmpComponent.Name & ".frm"
          Case vbext_ct_StdModule
            stFileName = cmpComponent.Name & ".bas"
        End Select
        If stFileName <> vbNullString Then
          cmpComponent.Export VFMFileAddTrailingSlash(stFolderPath) & stFileName
        End If
      End If
    Next cmpComponent
    
  End With
End Function
