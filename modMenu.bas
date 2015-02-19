Attribute VB_Name = "modMenu"
Option Explicit
Public Const MENU_BAR = PROJ_NAME & " menu"

Sub createMenu()
Dim cmBar As CommandBar
Dim butn1, butn2 As CommandBarButton
  ' panel creation
  Set cmBar = Application.CommandBars.Add(Name:=MENU_BAR, Position:=msoBarTop, MenuBar:=False, Temporary:=True)
  ' button creation
  Set butn1 = addMenuItem(cmBar, "transformPersentFormat", OPTION1, 1162)
  Set butn2 = addMenuItem(cmBar, "RunMyMacro1", "Òåñò", 1398)
  Set butn2 = addMenuItem(cmBar, "RunMyMacro1", "Òåñò234", 1397)
  cmBar.Visible = True
End Sub

Sub deleteMenu()
Dim i As Integer
    On Error Resume Next 'in case the menu item has already been deleted
    For i = 0 To Application.CommandBars(MENU_BAR).Controls.Count
      Application.CommandBars(MENU_BAR).Controls(i).Delete
    Next i
    Application.CommandBars(MENU_BAR).Delete
    On Error GoTo 0
End Sub

Private Function addMenuItem(menu As CommandBar, _
  ByVal onAction As String, ByVal caption As String, Optional fcid As Integer = 2950) As CommandBarButton
  Dim menuItem As CommandBarButton
  Set menuItem = menu.Controls.Add(Type:=msoControlButton, ID:=2950)
  With menuItem
    menuItem.Style = msoButtonIconAndWrapCaption
    If fcid <> 2950 Then menuItem.FaceId = fcid
    menuItem.onAction = onAction
    menuItem.caption = caption
  End With
  Set addMenuItem = menuItem
End Function


Private Sub RunMyMacro1()
MsgBox "This is test macro", vbInformation, PROJ_NAME
End Sub
