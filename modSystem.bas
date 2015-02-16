Attribute VB_Name = "modSystem"
Option Explicit
'==============================================================================
' module for system options
' requirements for references
'    {420B2830-E718-11CF-893D-00A0C9054228} | 1 | 0 | Scripting
'    {F935DC20-1CF0-11D0-ADB9-00C04FD58A0B} | 1 | 0 | IWshRuntimeLibrary
'==============================================================================
' declare for system options
Public Const PROJ_NAME = "projName"
Public initReady As Boolean
Public usr_name As String
'==============================================================================
' declare for logging options
Public log_dir, log_file_full As String
'==============================================================================

'------------------------------------------------------------------------------
' procedure for Workbook_Open()
'------------------------------------------------------------------------------
Sub autoOpen()
  Dim fso As New FileSystemObject
  modeScripting
  If fso.GetExtensionName(ActiveWorkbook.name) <> "xlsm" Then
    MsgBox "This workbook have incorrect format, need to be xlsm", vbOKOnly, PROJ_NAME & " error"
    modeHuman
    ActiveWorkbook.Close False
    End
  End If
  removeBrockenRefs ' https://github.com/mcgr0g/vbaCore/blob/master/modReference.bas
  log_file_full = loggerInit(getTempFolder(PROJ_NAME))
  init
  installRefsFromList ' https://github.com/mcgr0g/vbaCore/blob/master/modReference.bas
  'blabla  
  modeHuman
End Sub

'------------------------------------------------------------------------------
' initialization of project variables
'------------------------------------------------------------------------------
Sub init()
  If initReady Then Exit Sub
  initReady = True
  log_dir = setSlash2End(getTempFolder(PROJ_NAME))
  If log_file_full = "" Then log_file_full = loggerInit(log_dir, True)
  logger "started in folder " & Application.ActiveWorkbook.FullName
  'C:\Users\%user%\AppData\Local\Temp\projName\
  localSettings(0) = log_file_full
  usr_name = getUserLogin
  logger "current user is " & usr_name
  localSettings(1) = ActiveWorkbook.name
End Sub

'------------------------------------------------------------------------------
'use it at the begin of huge macro
'------------------------------------------------------------------------------
Sub modeScripting()
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  Application.CutCopyMode = False
  Application.Calculation = xlCalculationManual
End Sub

'------------------------------------------------------------------------------
'use it at the end of huge macro
'------------------------------------------------------------------------------
Sub modeHuman()
  Application.CutCopyMode = False
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Application.Calculation = xlCalculationAutomatic
End Sub

'------------------------------------------------------------------------------
'to avoid errors, wrap windows path in this
'------------------------------------------------------------------------------
Function setSlash2End(ByVal path As String) As String
  If Right(path, 1) <> "\" Then path = path & "\"
  setSlash2End = path
End Function

'------------------------------------------------------------------------------
' deactivate ALL buttons
'------------------------------------------------------------------------------
Sub hideButtons()
Dim cmdButton As OLEObject
Dim ws As Worksheet
  Application.EnableEvents = False
  For Each ws In ActiveWorkbook
    For Each cmdButton In ws
      cmdButton.Enabled = False
    Next
  Next
  Application.EnableEvents = True
End Sub

'------------------------------------------------------------------------------
'resolve username
'http://social.msdn.microsoft.com/Forums/en-US/a978fc57-88a2-49d2-89a9-ba645e1f5651/how-do-you-find-out-the-windows-user-name-from-vba?forum=isvvba
'reference to Windows Scripting Host Object Model
'------------------------------------------------------------------------------
Function getUserLogin() As String
On Error GoTo errHandler:
Dim objNetwork As IWshRuntimeLibrary.WshNetwork
  Set objNetwork = New IWshRuntimeLibrary.WshNetwork
  If objNetwork Is Nothing Then
    Shell "Regsvr32 -s scrun.dll", vbHide
    Application.Wait Now + (1# / 3600 / 24)
    Shell "Regsvr32 -s wshom.ocx", vbHide
    Application.Wait Now + (1# / 3600 / 24)
    Set objNetwork = CreateObject("WScript.Network")
  End If
  If objNetwork Is Nothing Then GoTo errHandler:
  logger "current user: " & objNetwork.username
  getUserLogin = objNetwork.username
  Set objNetwork = Nothing
  Exit Function
errHandler:
  handleError "[getUserLogin]", True
End Function

'----------------------------------------------
' resolving temp folder for logs
'------------------------------------------------------------------------------
Function getTempFolder(Optional ByVal subfolder As String = "nothing") As String
Dim temp_path As String
Dim fso As Object
  logger "resolve temp folder (with subdir = " & subfolder & ")"
  Set fso = CreateObject("Scripting.FileSystemObject")
  temp_path = setSlash2End(fso.GetSpecialFolder(2))
  If subfolder <> "nothing" Then
    temp_path = setSlash2End(temp_path) & subfolder
    If Len(Dir(temp_path, vbDirectory)) = 0 Then
      MkDir temp_path ' NO slash at the end
      logger "creating " & temp_path
    Else: logger "subdir is already existed"
    End If
  End If
  getTempFolder = temp_path
End Function

'------------------------------------------------------------------------------
'universal error handler
'------------------------------------------------------------------------------
Sub handleError(Optional errLocation As String = "", Optional showUI As Boolean = False)
  Dim errorMessage, msg_answ As String
  errorMessage = "Error in " & errLocation & ", {" & Err.Source & "} : vba error " & Err.Number & vbCrLf & Err.Description
  If Len(log_file_full) Then
    logger "----------------------------------------"
    logger errorMessage
    logger "----------------------------------------"
  Else
    Debug.Print "----------------------------------------"
    Debug.Print ">> " & errorMessage
    Debug.Print "----------------------------------------"
  End If
  If showUI Then
    msg_answ = MsgBox(errorMessage & vbCrLf _
      & "If the error occurs next time contact with developers" & vbCrLf _
      & "do it now?", vbCritical, PROJ_NAME & " error")
    Select Case msg_answ
      Case vbYes
        ThisWorkbook.FollowHyperlink Address:="https://github.com/mcgr0g"
      Case vbNo
      ' something else
    End Select
  End If
End Sub

'##############################################################################
' options for logging
'##############################################################################
'------------------------------------------------------------------------------
'initialization of logging
'------------------------------------------------------------------------------
Function loggerInit(ByVal path2log As String, Optional Reinit As Boolean = False) As String
Dim log_file As Integer
Dim msg As String
log_file = FreeFile
  path2log = setSlash2End(path2log) & "log.txt"
  msg = "initialization of logging into " & path2log
  If Reinit Then
    msg = "re" + msg
    Open path2log For Append As #log_file ' write to old file
  Else
    Open path2log For Output As #log_file ' clear & write
  End If
  Debug.Print msg
  Print #log_file, msg
  Close #log_file
  loggerInit = path2log
End Function

'------------------------------------------------------------------------------
'logging into:
'0 - only into Immediate Window
'1 - like option 1 and into file
'2 - like option 2 and mailing to somebody
'------------------------------------------------------------------------------
Sub logger(ByVal data_string As String, Optional level As Byte = 1)
Dim log_file As Integer
log_file = FreeFile
  If level > 0 Then
    On Error GoTo errNotInited:
    Open log_file_full For Append As #log_file
    Print #log_file, "# " & data_string
    Close #log_file
    If level = 2 Then mailAdmin "debug", data_string
    On Error GoTo 0
  End If
  Debug.Print "# " & data_string
  Exit Sub
errNotInited:
  If Err.Number = 75 Then
    Debug.Print "> logger is not inited, executing:"
    Debug.Print ">> " & data_string 
  Else
    handleError "[logger]"
  End If
End Sub

Sub getLogDir()
  Dim msg_answ As Integer
  msg_answ = MsgBox(log_dir & vbCrLf & "open?", vbQuestion + vbYesNo, PROJ_NAME & " log folder")
  If msg_answ = vbYes Then ThisWorkbook.FollowHyperlink Address:=log_dir
End Sub

Sub getLogFile()
  Dim msg_answ As Integer
  msg_answ = MsgBox(log_file_full & vbCrLf & "open?", vbQuestion + vbYesNo, PROJ_NAME & " log file")
  If msg_answ = vbYellow Then ThisWorkbook.FollowHyperlink Address:=log_file_full
End Sub
