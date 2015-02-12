Attribute VB_Name = "modReference"
Option Explicit
'==============================================================================
' module for Reference connect, disconnect and clearing
' requirements for 
'   modSystem.bas
'   need mark checkbox in
'   -->Tools|macro|security|trusted publishers tab|trust access to visual basic project
'==============================================================================
' declare for addin options
Public localSettings(0 To 10) As String
Public inputSettings(0 To 10) As String
Public addin_param, addin_name As String
'==============================================================================

'##############################################################################
' options for ownRef reference
'##############################################################################
'------------------------------------------------------------------------------
'check out reference and print all list
'------------------------------------------------------------------------------
Function addedRef(Optional refName As String = "nothing", Optional showAll As Boolean = False) As Boolean
Dim theRef As Object
Dim res As Boolean
  logger "check out reference list"
  For Each theRef In ActiveWorkbook.VBProject.References
    If showAll Then logger theRef.GUID & " , " & theRef.Major & " , " & theRef.Minor & " ' " & theRef.name
    If (refName <> "nothing") And (refName = theRef.name) Then
      res = True
      logger "найдено: " & refName
    End If
  Next theRef
  addedRef = res
  Exit Function
  
errHandler:
  handleError "[addedRef]"
End Function

'-------------------------------------------------
'no comments :)
'-------------------------------------------------
Sub removeRef(refName)
Dim ref As Reference
Dim vbPr As VBProject
On Error GoTo errHandler:
  logger "removing reference " & refName
  Set vbPr = ActiveWorkbook.VBProject
  Set ref = vbPr.References(refName)
  vbPr.References.Remove ref
  Exit Sub
errHandler:
  If Err.Number = 9 Then
    logger "reference already removed"
  Else
    handleError "[removeRef]"
  End If
End Sub

'------------------------------------------------------------------------------
'no comments :)
'------------------------------------------------------------------------------
Sub removeBrockenRefs()
Dim i As Integer
Dim theRef As Object
  logger "removing broken reference"
  On Error GoTo errHandler:
  For i = ThisWorkbook.VBProject.References.Count To 1 Step -1
    Set theRef = ThisWorkbook.VBProject.References.Item(i)
    'logger i & ": checking:  " & theRef.name
    If theRef.IsBroken Then
      ThisWorkbook.VBProject.References.Remove theRef
      logger "removed broken reference: " & theRef.name
    End If
  Next i
  Exit Sub
errHandler:
  handleError "[removeBrockenRefs]"
End Sub

'------------------------------------------------------------------------------
' list of references for this project
' you can modify it as needed
'------------------------------------------------------------------------------
Private Function getRef4install() As String()
Dim refList() As String
Dim element As Variant
On Error GoTo errHandler:
  ReDim refList(2)
  refList(0) = "{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}" 'IWshRuntimeLibrary
  refList(1) = "{420B2830-E718-11CF-893D-00A0C9054228}" 'Scripting
  'refList(2) = "{00062FFF-0000-0000-C000-000000000046}" 'Outlook
  getRef4install = refList
  Exit Function
errHandler:
  handleError "[getRef4install]"
End Function

'------------------------------------
' connect reference by GUID
' used predefined list
'------------------------------------
Sub installRefsFromList()
Dim guidNo As Variant
Dim refList() As String
  logger "connect reference by guid list"
  On Error GoTo errHandler:
  refList = getRef4install
  With ThisWorkbook.VBProject.References
    For Each guidNo In refList
      .AddFromGuid guidNo, 1, 0
      logger "success for " & guidNo
    Next guidNo
  End With
  Exit Sub
errHandler:
  Select Case Err.Number
    Case Is = 1004
      MsgBox "need mark checkbox in" & vbCrLf _
        & "Tools|macro|security|trusted publishers tab|trust access to visual basic project"
      End
    Case Is = 32813
      handleError "connect ref " & guidNo & ", it's already connected"
      Resume Next
    Case Is = -2147319779
      logger "refList too big, there are null entries", 0
    Case Else
      handleError "[installRefsFromList], its guid is " & guidNo
      Resume Next
  End Select
End Sub

'-------------------------------------------------------
' connect reference from file
' based on http://support.microsoft.com/kb/308340/en-us
'-------------------------------------------------------
Function addRefFromFile(targetFilePath As String) As Boolean
Dim vbRef As Object
Dim projRefs As References
Dim proj As VBProject
Dim toImport, isConnected As Boolean
Dim targetFileName As String
  isConnected = False
  removeBrockenRefs
  logger "connect reference from file " & targetFilePath
  toImport = True
  Set proj = ThisWorkbook.VBProject
  Set projRefs = proj.References
  targetFileName = Dir(targetFilePath)
  For Each vbRef In projRefs
    If vbRef.fullpath = targetFilePath Then
      toImport = False
      isConnected = True
      logger "it's already connected " & targetFileName
    End If
  Next
  
  On Error GoTo errorHandler:
  If toImport Then
    If proj.protection = vbext_pp_locked Then
      logger "workbook is protected, you need to enter pass" ' based on http://stackoverflow.com/a/2505006
    End If
    projRefs.AddFromFile targetFilePath
    
    isConnected = True
    logger "success for " & targetFileName
  End If
  
exitHere:
  addRefFromFile = isConnected
  Exit Function
  
errorHandler:
  handleError "[addRefFromFile]"
  GoTo exitHere:
End Function
