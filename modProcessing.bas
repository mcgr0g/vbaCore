Attribute VB_Name = "modProcessing"
Option Explicit

Sub handlerFormParams(todo As String, ByRef wRange As range)
  init
'  logger "starting handlerFormParams"
'  On Error GoTo errHandler:
  Select Case todo
    Case OPTION1
      logger "there is valid option 'todo'=" & todo
      Call setPersentFormat(wRange)
    Case Else
      logger "there is invalid option 'todo'"
  End Select
  Exit Sub
errHandler:
  handleError "[handlerFormParams] with" & vbCrLf _
    & "todo= " & todo
End Sub

Private Sub setPersentFormat(ByRef wRange As range)
Dim iRow, iCol As Object
  logger "starting setPersentFormat"
  On Error GoTo errHandler:
  modeScripting
  For Each iRow In wRange.Rows
    For Each iCol In wRange.Columns
      With ActiveSheet.Cells(iRow.Row, iCol.Column)
        .Style = "Percent"
        .NumberFormat = "0.00%" 'Percentage
        .Value = .Value / 100
      End With
    Next
  Next
exitHere:
  modeHuman
  Exit Sub
errHandler:
  handleError "[setPersentFormat]"
  GoTo exitHere:
End Sub
