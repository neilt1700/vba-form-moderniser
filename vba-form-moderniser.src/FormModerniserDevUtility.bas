Attribute VB_Name = "FormModerniserDevUtility"
Option Explicit

Private Const m_stFormModerniserModule As String = "FormModerniserModule.bas"
Private Const m_stDevUtilityVersion As String = "1.2"

Private m_stCurrentVersion As String
Private m_boolCurrentVersionLoaded As Boolean

' Use to import modules into your project:
Private Sub VFMImport()
  VFMImportModules
End Sub

' Run this if asked to:
Private Sub VFMStoreVersionNumber()
  VFMStoreCurrentVersionNumber
End Sub

Private Sub VFMExport()
  VFMExportToFiles
End Sub

' =============================================================================

Public Function VFMFileNames(Optional ByVal boolImport As Boolean = True) As String
  ' Names must be separated by spaces.
  ' The file extension is included for importing - but not exporting.
  If boolImport Then
    VFMFileNames = "CKeyDownResponder.cls " & _
                   "CLabelControl.cls " & _
                   "CLabelControlFrameResponder.cls " & _
                   "CLabelControlResponder.cls " & _
                   "CLabelControls.cls " & _
                   "CLabelControlsManager.cls " & _
                   "FormModerniserModule.bas " & _
                   "VFMFactory.bas"
  Else
    VFMFileNames = "CKeyDownResponder " & _
                   "CLabelControl " & _
                   "CLabelControlFrameResponder " & _
                   "CLabelControlResponder " & _
                   "CLabelControls " & _
                   "CLabelControlsManager " & _
                   "FormModerniserDevPPT " & _
                   "FormModerniserDevWord " & _
                   "FormModerniserDevExcel " & _
                   "FormModerniserDevUtility " & _
                   "FormModerniserModule " & _
                   "VFMFactory"
  End If
End Function

Private Function VFMStoreCurrentVersionNumber() As String
   m_stCurrentVersion = FormModerniserModule.g_stVERSION
   m_boolCurrentVersionLoaded = True
   MsgBox "Current version number (" & m_stCurrentVersion & ") stored.", vbInformation
End Function

Private Function VFMVersionFromFile(ByVal stFolderPath As String, Optional ByVal stVersionType As String = "macro") As String

  Dim stFileName As String
  Dim stVersionIdentifier As String
  
  Select Case stVersionType
    Case "macro"
      stFileName = m_stFormModerniserModule
      stVersionIdentifier = "Public Const g_stVERSION As String = "
    Case "devutility"
      stFileName = "FormModerniserDevUtility.bas"
      stVersionIdentifier = "Private Const m_stDevUtilityVersion As String = "
    Case Else
      Exit Function
  End Select

  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  Const ForReading As Long = 1
  
  Dim stFilePath As String
  stFilePath = VFMFileAddTrailingSlash(stFolderPath) & stFileName
  
  If Not fso.FileExists(stFilePath) Then
    Exit Function
  End If
  
  Dim boolVersionFound As Boolean
  Dim stLine As String
  Dim stVersion As String
    
  Set f = fso.OpenTextFile(FileName:=stFilePath, iomode:=ForReading, Format:=0)
  Do While f.AtEndOfStream <> True And boolVersionFound <> True
    stLine = Trim$(f.Readline)
    If Mid(stLine, 1, Len(stVersionIdentifier)) = stVersionIdentifier Then
      stVersion = Mid(stLine, Len(stVersionIdentifier) + 1)
      ' Removes leading and trailing quotes:
      stVersion = Mid(stVersion, 2, Len(stVersion) - 2)
      boolVersionFound = True
    End If
  Loop
  f.Close

  VFMVersionFromFile = stVersion

End Function

Private Sub VFMImportModules()

  ' As the current version exists in a module we want to replace, we
  ' cannot get anything from that module when attempting to replace it -
  ' so a separate procedure must be run first to store the version before
  ' running this one.
  If Not m_boolCurrentVersionLoaded Then
    MsgBox "Please run ""VFMStoreVersionNumber"" first.", vbExclamation
    Exit Sub
  End If

  Dim stFileNames() As String
  stFileNames = Split(VFMFileNames)
  
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  Dim stFolderPath As String
  stFolderPath = VFMGetFolder(vbNullString)
  
  ' (1) Browse to source folder
  If stFolderPath = vbNullString Then
    Exit Sub
  End If
  
  Dim boolAllFilesExist As Boolean
  boolAllFilesExist = True
  
  Dim stFileName As Variant
  For Each stFileName In stFileNames
    If Not fso.FileExists(stFolderPath & "\" & stFileName) Then
      boolAllFilesExist = False
      Exit For
    End If
  Next stFileName

  If Not boolAllFilesExist Then
    MsgBox "Not all the expected files exist. Aborting. Nothing has been imported.", vbExclamation + vbOKOnly, "Import"
    Exit Sub
  End If
  
  Dim stModuleNameList As String
  stModuleNameList = Join(stFileNames, vbCrLf)
  
  Dim stCurrentVersion As String
  stCurrentVersion = m_stCurrentVersion
  
  Dim stVersionFromFile As String
  stVersionFromFile = VFMVersionFromFile(stFolderPath)
  
  Dim stDevUtilityVersion As String
  stDevUtilityVersion = m_stDevUtilityVersion
  Dim stDevUtilityVersionFromFile As String
  stDevUtilityVersionFromFile = VFMVersionFromFile(stFolderPath, "devutility")
  
  Dim stDevUtilityMsg As String
  If stDevUtilityVersionFromFile <> vbNullString Then
    If stDevUtilityVersion <> stDevUtilityVersionFromFile Then
    
      Dim stDiffVersion As String
    
      If CDbl(stDevUtilityVersion) > CDbl(stDevUtilityVersionFromFile) Then
        stDiffVersion = "*newer*"
      Else
        stDiffVersion = "older"
      End If
      
      stDevUtilityMsg = "The version of the DevUtility module you are " & _
                        "using (" & stDevUtilityVersion & ") is " & stDiffVersion & " than the version " & _
                        "of the DevUtility module in the folder you are " & _
                        "importing from (" & stDevUtilityVersionFromFile & "). " & _
                        "You should import that manually first before " & _
                        "proceeding."
                        
      MsgBox stDevUtilityMsg, vbOKOnly + vbExclamation, "Form Moderniser Import"
      m_boolCurrentVersionLoaded = False
      Exit Sub
    End If
  End If
  
  Dim stOtherInfoMsg As String
  stOtherInfoMsg = "Note: this does NOT load the VFMUtility module - do this " & _
                   "manually as and when necessary. You will " & _
                   "also need to have the appropriate FormModerniserDevPPT/Excel/Word " & _
                   "module."
  
  Dim stTargetDocument As String
  stTargetDocument = VFMCurrentDocument
  
  Dim stBackupPath As String
  stBackupPath = VFMBackupFilePath
  
  ' (3)
  Dim stMsg As String
  stMsg = stOtherInfoMsg & vbCrLf & vbCrLf
  stMsg = stMsg & "Current version: " & stCurrentVersion & vbCrLf
  stMsg = stMsg & "Version to be imported: " & stVersionFromFile & vbCrLf & vbCrLf
  stMsg = stMsg & "Source Folder: " & stFolderPath & vbCrLf
  stMsg = stMsg & "Target Document: " & stTargetDocument & vbCrLf
  stMsg = stMsg & "Backup path for current document: " & stBackupPath & vbCrLf & vbCrLf
  stMsg = stMsg & "The following modules will be imported - replacing existing modules with the same name: " & vbCrLf
  stMsg = stMsg & stModuleNameList & vbCrLf & vbCrLf
  stMsg = stMsg & "Make sure you have saved this document before proceding. " & _
                  "This utility does create a backup - but only of the last " & _
                  "saved version." & vbCrLf & vbCrLf
  stMsg = stMsg & "Do you want to continue?"
  
  If MsgBox(stMsg, vbYesNoCancel + vbInformation, "Form Moderniser Module Import") <> vbYes Then
    Exit Sub
  End If
  
  Dim stCurrentDocumentPath As String
  stCurrentDocumentPath = VFMFileAddTrailingSlash(VFMCurrentFolder) & VFMCurrentDocument
  
  VFMBackup stCurrentDocumentPath, stBackupPath
  
  Dim stModulePath As String
  
  For Each stFileName In stFileNames
    stModulePath = stFolderPath & "\" & stFileName
    VFM_RemoveModule VFMFileName(stModulePath, False)
    VFM_ImportModule stModulePath
  Next stFileName
  
  m_boolCurrentVersionLoaded = False

End Sub

Private Sub VFMExportToFiles()

  Dim stFolderPath As String
  stFolderPath = VFMGetFolder(vbNullString)
  
  Dim stModuleNames As String
  stModuleNames = Replace(VFMFileNames(False), " ", vbCrLf)
  
  Dim stMsg As String
  stMsg = "Are you sure you would like to export the following modules (where they exist):" & _
        vbCrLf & stModuleNames & vbCrLf & "to " & stFolderPath & "?" & vbCrLf & _
        "Any existing files will be overwritten."
  
  If MsgBox(stMsg, vbYesNoCancel + vbInformation, "Module Export") <> vbYes Then
    Exit Sub
  End If
  
  If stFolderPath <> vbNullString Then
    VFM_ExportModules VFMFileNames(False), stFolderPath
  End If

End Sub

' File Functions
' ==============

' Returns unique file path for given file path - by altering the file name.
Private Function VFMUniqueFilePath(stFilePath As String) As String

  Dim stUniqueFilePath As String
  Dim i As Long
  
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")

  stUniqueFilePath = stFilePath
  
  If fso.FileExists(stFilePath) Then
    
    i = 0
    Dim stFileFolderPath As String
    Dim stFileName As String
    Dim stFileExt As String
    Dim boolFileFound As Boolean
    
    stFileFolderPath = VFMFileFolderPath(stFilePath)
    stFileName = VFMFileName(stFilePath, False)
    stFileExt = VFMFileExt(stFilePath)
    
    If stFileExt <> vbNullString Then
      stFileExt = "." & stFileExt
    End If
    
    ' Looking for " (dd)" at end of file name
    Do While boolFileFound = False
      i = i + 1
      stUniqueFilePath = stFileFolderPath & stFileName & " (" & i & ")" & stFileExt
      If fso.FileExists(stUniqueFilePath) <> True Then
        boolFileFound = True
      End If
    Loop
  End If
  
  VFMUniqueFilePath = stUniqueFilePath

End Function

Private Function VFMTrailingSlash(varIn As Variant) As String
  If Len(varIn) > 0& Then
    If Right$(varIn, 1&) = "\" Then
      VFMTrailingSlash = varIn
    Else
      VFMTrailingSlash = varIn & "\"
    End If
  End If
End Function

Private Function VFMFileExt(stPath) As String
  If InStr(stPath, ".") > 0 Then
    VFMFileExt = Right$(stPath, Len(stPath) - InStrRev(stPath, "."))
  Else
    VFMFileExt = vbNullString
  End If
End Function

Private Function VFMFileStripTrailingSlash(stPath) As String
  VFMFileStripTrailingSlash = stPath
  If Len(stPath) > 0 Then
    If Right$(stPath, 1) = "\" Then
      VFMFileStripTrailingSlash = Mid$(stPath, 1, Len(stPath) - 1)
    End If
  End If
End Function

Public Function VFMFileAddTrailingSlash(stPath) As String
  VFMFileAddTrailingSlash = stPath
  If Len(stPath) > 0 Then
    If Not Right$(stPath, 1) = "\" Then
      VFMFileAddTrailingSlash = stPath & "\"
    End If
  Else
    VFMFileAddTrailingSlash = "\"
  End If
End Function

' Returns the file name without the extension:
' For folders, the trailing slash - if any - is stripped off first
Private Function VFMFileName(ByVal stPath As String, Optional ByVal lower_case As Boolean = True, _
                         Optional ByVal boolFolder = False) As String
  
  Dim stFileName As String
  
  stFileName = VFMFileStripTrailingSlash(stPath)
  stFileName = Mid$(stFileName, InStrRev(stFileName, "\") + 1)
  
  If InStrRev(stFileName, ".") <> 0 And InStr(stFileName, ".") <> 1 And boolFolder <> True Then
    stFileName = Mid$(stFileName, InStrRev(stFileName, "\") + 1, InStrRev(stFileName, ".") - InStrRev(stFileName, "\") - 1)
  End If
  
  If lower_case = True Then
    stFileName = LCase$(stFileName)
  End If
  
  VFMFileName = stFileName
  
End Function

Private Function VFMFileNameWithExt(stPath) As String
  VFMFileNameWithExt = Mid$(stPath, InStrRev(stPath, "\") + 1)
End Function

' Gets the folder path from a full path to a file
Private Function VFMFileFolderPath(stPath) As String
  Dim stFileNameWithExt
  stFileNameWithExt = VFMFileNameWithExt(stPath)
  VFMFileFolderPath = Mid$(stPath, 1, Len(stPath) - Len(stFileNameWithExt))
End Function

Private Function VFMGetFolder(ByVal strPath As String) As String

  Dim fldr As FileDialog
  Dim sItem As String
  sItem = vbNullString
  Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
  With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
  End With
NextCode:
  VFMGetFolder = sItem
  Set fldr = Nothing
    
End Function

' Information about the current document
' ======================================

Private Function VFMCurrentDocument() As String
  Dim stCurrentDocument As String
  
  Select Case Application.Name
    Case "Microsoft PowerPoint"
      stCurrentDocument = CallByName(CallByName(Application, "ActivePresentation", VbGet), "Name", VbGet)
    Case "Microsoft Excel"
      stCurrentDocument = CallByName(CallByName(Application, "ActiveWorkbook", VbGet), "Name", VbGet)
    Case "Microsoft Word"
      stCurrentDocument = CallByName(CallByName(Application, "ActiveDocument", VbGet), "Name", VbGet)
  End Select

  VFMCurrentDocument = stCurrentDocument
End Function

Private Function VFMCurrentFolder() As String
  Dim stCurrentFolder As String
  
  Select Case Application.Name
    Case "Microsoft PowerPoint"
      stCurrentFolder = CallByName(CallByName(Application, "ActivePresentation", VbGet), "Path", VbGet)
    Case "Microsoft Excel"
      stCurrentFolder = CallByName(CallByName(Application, "ActiveWorkbook", VbGet), "Path", VbGet)
    Case "Microsoft Word"
      stCurrentFolder = CallByName(CallByName(Application, "ActiveDocument", VbGet), "Path", VbGet)
  End Select

  VFMCurrentFolder = stCurrentFolder
End Function

Private Function VFMBackupFilePath() As String
  VFMBackupFilePath = VFMUniqueFilePath(VFMFileAddTrailingSlash(VFMCurrentFolder) & VFMCurrentDocument & ".bak")
End Function

Private Function VFMBackup(ByVal stSource As String, ByVal stTarget As String)

  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  fso.CopyFile stSource, stTarget, False

End Function
