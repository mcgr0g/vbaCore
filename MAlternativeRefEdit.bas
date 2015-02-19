Attribute VB_Name = "MAlternativeRefEdit"
Option Explicit
' based on http://peltiertech.com/refedit-control-alternative/
''============================================================================
Sub ShowDialog(tCaption As String)
  Dim frmAlternativeRefEdit As FAlternativeRefEdit
  Dim rRange As range
  Dim sRange As String
  Dim bCanceled As Boolean
  
  If ActiveSheet Is Nothing Then Exit Sub
  
  If TypeName(Selection) = "Range" Then
    Set rRange = Selection.Areas(1)
    sRange = Selection.Parent.Name
    If InStr(rRange.Address(, , , True), "'") > 0 Then
      sRange = "'" & sRange & "'"
    End If
    sRange = sRange & "!"
    sRange = sRange & rRange.Address
  End If
  
  Set frmAlternativeRefEdit = New FAlternativeRefEdit
  With frmAlternativeRefEdit
    .Address = sRange
    .caption = PROJ_NAME & tCaption
    .Show
    ' wait for user input
    bCanceled = .Cancel
    If Not bCanceled Then
      sRange = .Address
    End If
  End With
  
  Unload frmAlternativeRefEdit
  Set frmAlternativeRefEdit = Nothing

  If bCanceled Then GoTo ExitSub
  
  If IsRange(sRange) Then
    Application.GoTo range(sRange)
  End If
  
  Call handlerFormParams(tCaption, rRange)
  
ExitSub:
End Sub
''============================================================================
Public Function IsRange(ByVal sRangeAddress As String) As Boolean
  Dim TestRange As range
  
  IsRange = True
  On Error Resume Next
  Set TestRange = range(sRangeAddress)
  If Err.Number <> 0 Then
    IsRange = False
  End If
  Err.Clear
  On Error GoTo 0
  Set TestRange = Nothing
End Function
''============================================================================
