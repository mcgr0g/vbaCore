Attribute VB_Name = "modAddin"
Option Explicit
'==============================================================================
' module for addin run, transmit and receive parameter
' requirements for modSystem.bas
'==============================================================================
' declare for addin options
Public localSettings(0 To 10) As String
Public inputSettings(0 To 10) As String
Public addin_param, addin_name As String
'==============================================================================

'------------------------------------------------------------------------------
'check array is empty
'------------------------------------------------------------------------------
Public Function isArrayEmpty(Arr As Variant) As Boolean
Dim LB, UB As Long
  Err.Clear
  On Error Resume Next
  If IsArray(Arr) = False Then
    IsArrayEmpty = True ' we weren't passed an array, return True
  End If
  UB = UBound(Arr, 1)
  If (Err.Number <> 0) Then
    IsArrayEmpty = True
  Else
    ' On rare occassion, under circumstances I cannot reliably replictate,
    ' Err.Number will be 0 for an unallocated, empty array.
    ' On these occassions, LBound is 0 and UBoung is -1.
    ' To accomodate the weird behavior, test to see if LB > UB.
    ' If so, the array is not allocated.
    Err.Clear
    LB = LBound(Arr)
    If LB > UB Then
      IsArrayEmpty = True
    Else
      IsArrayEmpty = False
    End If
  End If
End Function
'##############################################################################
'options for addin solver.xlam,
'##############################################################################

'------------------------------------------------------------------------------
' no comment :)
'------------------------------------------------------------------------------
Sub installAddIn(ByVal addinDir As String, ByVal addinName As String)
Dim AI As Excel.AddIn
  logger "installing addin " & addinName
  On Error GoTo errHandler:
  Set AI = Application.AddIns.Add(Filename:=addinDir & addinName & ".xlam")
  AI.Installed = True
  logger "success"
  Exit Sub
errHandler:
  handleError "[installAddIn]"
End Sub

'------------------------------------------------------------------------------
' no comment :)
'------------------------------------------------------------------------------
Function isInstalled(ByVal addinName As String) As Boolean
  On Error GoTo errHandler:
  If AddIns(addinName).Installed Then
    logger addinName & " is installed"
    isInstalled = True
  End If
  Exit Function
errHandler:
  handleError "[isInstalled]"
End Function


'------------------------------------------------------------------------------
'you can run its macro whith array localSettings as parameter
'------------------------------------------------------------------------------
Sub remoteRunInAddin(ByVal addin_name As String)
  On Error GoTo errHandler:
  logger "remore run macro " & addin_name
  If IsArrayEmpty(localSettings) Then
    logger "local settings are not initialized, run without it"
    Application.Run "'" & addin_name & "'.xlam!" & "initialize"
  Else
    logger "local settings are initialized, run with it"
    Application.Run "'" & addin_name & "'.xlam!" & "initialize", localSettings
  End If
  Exit Sub
errHandler:
  handleError "[remoteRunInAddin]"
End Sub

'------------------------------------------------------------------------------
'this can help to get execution stream  from it
'------------------------------------------------------------------------------
Sub receptionParams(ByRef inputSettings() As String)
  If Not initReady Then init
  logger "listening-in for addin parameter"
On Error GoTo errorHandler:
  If Not IsArrayEmpty(inputSettings) Then
    logger "there is something.."
    addin_param = inputSettings(0)
    logger "received addin_param = " & addin_param
  End If
  Exit Sub
errorHandler:
  handleError "[receptionParams]"
End Sub

Sub closeAddin(wbName As String)
Dim lanBook As Workbook
  logger "trying to close " & wbName
  On Error GoTo errHandler:
  Set lanBook = Workbooks(wbName)
  lanBook.Close
  Exit Sub
errHandler:
  If Err.Number = 9 Then
    logger "already closed"
  Else
    handleError "[closeAddin]"
  End If
End Sub
