VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FAlternativeRefEdit 
   Caption         =   "Alternative RefEdit"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "FAlternativeRefEdit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FAlternativeRefEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
''============================================================================
Private mbCancel As Boolean
''============================================================================
Private Sub btnCancel_Click()
  mbCancel = True
  Me.Hide
End Sub

Private Sub btnOK_Click()
  mbCancel = False
  Me.Hide
End Sub
''============================================================================
Public Property Let Address(sAddress As String)
  CheckAddress sAddress
End Property

Public Property Get Address() As String
  Dim sAddress As String
  
  sAddress = Me.txtRefChtData.Text
  If IsRange(sAddress) Then
    Address = sAddress
  Else
    sAddress = Application.ConvertFormula(sAddress, xlR1C1, xlA1)
    If IsRange(sAddress) Then
      Address = sAddress
    End If
  End If

End Property

Public Property Get Cancel() As Boolean
  Cancel = mbCancel
End Property
''============================================================================
Private Sub txtRefChtData_DropButtonClick()
  Dim var As Variant
  Dim rng As range
  Dim sFullAddress As String
  Dim sAddress As String
  
  Me.Hide
  
  On Error Resume Next
  var = Application.InputBox("Select the range containing your data", _
      "Select Chart Data", Me.txtRefChtData.Text, Me.Left + 2, _
       Me.Top - 86, , , 0)
  On Error GoTo 0
  
  If TypeName(var) = "String" Then
    CheckAddress CStr(var)
  End If
  
  Me.Show
End Sub
''============================================================================
Private Sub CheckAddress1(sAddress As String)
  Dim rng As range
  Dim sFullAddress As String
  
  If Left$(sAddress, 1) = "=" Then sAddress = Mid$(sAddress, 2, 256)
  If Left$(sAddress, 1) = Chr(34) Then sAddress = Mid$(sAddress, 2, 255)
  If Right$(sAddress, 1) = Chr(34) Then sAddress = Left$(sAddress, Len(sAddress) - 1)
  
  If IsRange(sAddress) Then
    Set rng = range(sAddress)
  Else
    sAddress = Application.ConvertFormula(sAddress, xlR1C1, xlA1)
    If IsRange(sAddress) Then
      Set rng = range(sAddress)
    End If
  End If
  
  If Not rng Is Nothing Then
    sFullAddress = rng.Address(, , Application.ReferenceStyle, True)
    If Left$(sFullAddress, 1) = "'" Then
      sAddress = "'"
    Else
      sAddress = ""
    End If
    sAddress = sAddress & Mid$(sFullAddress, InStr(sFullAddress, "]") + 1)
    
    rng.Parent.Activate
    
    Me.txtRefChtData.Text = sAddress
  End If

End Sub

Private Sub CheckAddress(sAddress As String)
  ' changed following advice of Julien steelandt@yahoo.fr
  Dim rng As range
  Dim sFullAddress As String

  If Left$(sAddress, 1) = "=" Then sAddress = Mid$(sAddress, 2, 256)
  If Left$(sAddress, 1) = Chr(34) Then sAddress = Mid$(sAddress, 2, 255)
  If Right$(sAddress, 1) = Chr(34) Then sAddress = Left$(sAddress, Len(sAddress) - 1)

  On Error Resume Next
  sAddress = Application.ConvertFormula(sAddress, xlR1C1, xlA1)

  If IsRange(sAddress) Then
    Set rng = range(sAddress)
  End If

  If Not rng Is Nothing Then
    sFullAddress = rng.Address(, , Application.ReferenceStyle, True)
    If Left$(sFullAddress, 1) = "'" Then
      sAddress = "'"
    Else
      sAddress = ""
    End If
    sAddress = sAddress & Mid$(sFullAddress, InStr(sFullAddress, "]") + 1)

    rng.Parent.Activate

    Me.txtRefChtData.Text = sAddress
  End If

End Sub
''============================================================================
Private Sub UserForm_Initialize()
  Me.txtRefChtData.DropButtonStyle = fmDropButtonStyleReduce
  Me.txtRefChtData.ShowDropButtonWhen = fmShowDropButtonWhenAlways
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = vbFormControlMenu Then
    Cancel = True
    btnCancel_Click
  End If
End Sub
''============================================================================
