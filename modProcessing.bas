Attribute VB_Name = "modProcessing"
Option Explicit

Sub handlerFormParams(todo As String, ByRef wRange As Range)
  init
'  logger "starting handlerFormParams"
'  On Error GoTo errHandler:
  Select Case todo
    Case OPTION1
      logger "there is valid option 'todo'=" & todo
      Call setPersentFormat(wRange)
    Case OPTION2
      logger "there is valid option 'todo'=" & todo
      Call setContinuity(wRange)
    Case Else
      logger "there is invalid option 'todo'"
  End Select
  Exit Sub
errHandler:
  handleError "[handlerFormParams] with" & vbCrLf _
    & "todo= " & todo
End Sub

Sub setPersentFormat(ByRef wRange As Range)
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

Sub setContinuity(ByRef wRange As Range)
Dim iRow, iCol As Object
Dim iCell
Dim startPoint, secondPoint, startRng, destRng, samlpeRng As Range
Dim workWithRow As Boolean ' true-работаем со строками, иначе со столбцами.
Dim blankStartRow, blankStartCol, blankEndRow, blankEndCol, i, counterttl As Integer
Dim sampl, dest, samplVal, destVal As Variant
  logger "starting setContinuity"
  'On Error GoTo errHandler:
  modeScripting
  blankStartRow = wRange.Row
  blankStartCol = wRange.Column
  blankEndRow = wRange.Row + wRange.Rows.Count - 1
  blankEndCol = wRange.Column + wRange.Columns.Count - 1
  Set samlpeRng = Range(Cells(blankStartRow, blankStartCol), Cells(blankEndRow, blankEndCol))
  Set startPoint = ActiveSheet.Range(wRange.Cells(1, 1).Address())
  If wRange.Columns.Count < 2 Then
    If wRange.Rows.Count < 2 Then
      MsgBox "Выделена только одна ячейка", vbOKOnly, PROJ_NAME & " error"
      GoTo errHandler:
    Else
      logger "работаем со столбцом", 0
      workWithRow = False
      'Set secondPoint = Cells(wRange.Row + 1, wRange.Column)
      'MsgBox "Вторая ячейка массиа имеет значение " & secondPoint.Value, vbOKOnly, PROJ_NAME & " info"
    End If
  Else
    If wRange.Rows.Count < 2 Then
      logger "работаем со строкой", 0
      workWithRow = True
    Else
      MsgBox "Выделен двумерный массив вместо отдномерного", vbOKOnly, PROJ_NAME & " error"
      GoTo errHandler:
    End If
  End If
  
  If workWithRow Then
    ' работаем со строкой
    Set secondPoint = Cells(blankStartRow, blankStartCol + 1)
    ' копируем строку
    Rows(blankStartRow).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    blankStartRow = blankStartRow + 1
    blankEndRow = blankEndRow + 1
    'задаем 2 ячейки, находящуююся выше болванки на 1 строку
    Set startRng = Range(Cells(blankStartRow - 1, blankStartCol), Cells(blankStartRow - 1, blankStartCol + 1))
    'задаем образец, находящийся выше болванки на 1 строку
    Set destRng = Range(Cells(blankStartRow - 1, blankStartCol), Cells(blankStartRow - 1, blankEndCol))
    startRng.Cells(1, 1).Value = startPoint.Value
    startRng.Cells(1, 2).Value = Cells(blankStartRow, blankStartCol + 1).Value
    startRng.Select
    Selection.AutoFill Destination:=destRng, Type:=xlFillDefault
    ' задаем размерность цикла
    counterttl = samlpeRng.Columns.Count
    i = 1
    ' сверяем значения болванки с образцом
    Do Until i > counterttl Or i > 100
      sampl = samlpeRng(1, i).Value
      dest = destRng(1, i).Value
      If IsNumeric(sampl) And IsNumeric(dest) Then
        samplVal = Val(sampl)
        destVal = Val(dest)
      Else
        If IsDate(sampl) And IsDate(dest) Then
          samplVal = CDate(sampl)
          destVal = CDate(dest)
        Else
          MsgBox "неизвестный формат"
          End Sub
        End If
      End If
      If samplVal <> destVal Then
        Debug.Print "нашел косяк " & samlpeRng.Cells(1, i).Address
        Columns(blankStartCol + i - 1).Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        blankEndCol = blankEndCol + 1
        counterttl = counterttl + 1
        'расширили массив
        Set samlpeRng = Nothing
        Set destRng = Nothing
        Set samlpeRng = Range(Cells(blankStartRow, blankStartCol), Cells(blankEndRow, blankEndCol))
        Set destRng = Range(Cells(blankStartRow - 1, blankStartCol), Cells(blankStartRow - 1, blankEndCol))
        startRng.Select
        Selection.AutoFill Destination:=destRng, Type:=xlFillDefault
        samlpeRng(1, i).Value = destRng(1, i).Value
      End If
    'Next i
    i = i + 1
    Loop
    Rows(blankStartRow - 1).Select
    Selection.Delete Shift:=xlUp
  Else
    'работаем со столбцом
    Set secondPoint = Cells(blankStartRow + 1, blankStartCol)
    ' копируем строку
    Columns(blankStartCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    blankStartCol = blankStartCol + 1
    blankEndCol = blankEndCol + 1
    'задаем 2 ячейки, находящуююся левее болванки на 1 строку
    Set startRng = Range(Cells(blankStartRow, blankStartCol - 1), Cells(blankStartRow + 1, blankStartCol - 1))
    'задаем образец, находящийся левее болванки на 1 строку
    Set destRng = Range(Cells(blankStartRow, blankStartCol - 1), Cells(blankEndRow, blankStartCol - 1))
    startRng.Cells(1, 1).Value = startPoint.Value
    startRng.Cells(2, 1).Value = Cells(blankStartRow + 1, blankStartCol).Value
    startRng.Select
    Selection.AutoFill Destination:=destRng, Type:=xlFillDefault
    ' сверяем значения болванки с образцом
    counterttl = samlpeRng.Rows.Count
    i = 1
    Do Until i > counterttl Or i > 100
    'For i  = 1 To counterttl + 2 ' не спрашивайте, почему так
      sampl = samlpeRng(i, 1).Value
      dest = destRng(i, 1).Value
      If IsNumeric(sampl) And IsNumeric(dest) Then
        samplVal = Val(sampl)
        destVal = Val(dest)
      Else
        If IsDate(sampl) And IsDate(dest) Then
          samplVal = CDate(sampl)
          destVal = CDate(dest)
        Else
          MsgBox "неизвестный формат"
          End Sub
        End If
      End If
      If samplVal <> destVal Then
        Debug.Print "нашел косяк " & samlpeRng.Cells(i, 1).Address
        Rows(blankStartRow + i - 1).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        blankEndRow = blankEndRow + 1
        counterttl = counterttl + 1
        'расширили массив
        Set samlpeRng = Range(Cells(blankStartRow, blankStartCol), Cells(blankEndRow, blankEndCol))
        Set destRng = Range(Cells(blankStartRow, blankStartCol - 1), Cells(blankEndRow, blankEndCol - 1))
        startRng.Select
        Selection.AutoFill Destination:=destRng, Type:=xlFillDefault
        samlpeRng.Cells(i, 1).Value = destRng.Cells(i, 1).Value
      End If
    i = i + 1
    Loop
    Columns(blankStartCol - 1).Select
    Selection.Delete Shift:=xlToLeft
  End If
  
exitHere:
  modeHuman
  Exit Sub
errHandler:
  handleError "[setContinuity]"
  GoTo exitHere:
End Sub
