VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Odesa Fast File Search System - Alper ESKIKILIÇ -www.odesayazilim.com"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   Icon            =   "Odesafastsearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Odesafastsearch.frx":628A
   ScaleHeight     =   7830
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin Project1.Button Button1 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   7440
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   661
      ButtonStyle     =   8
      CaptionEffect   =   3
      BackColor       =   -2147483647
      BackColorPressed=   -2147483647
      BackColorHover  =   -2147483647
      BorderColor     =   -2147483647
      BorderColorPressed=   -2147483647
      BorderColorHover=   -2147483647
      Caption         =   "About Program"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.Button Command1 
      Height          =   495
      Left            =   6720
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ButtonStyle     =   6
      ButtonStyleColors=   3
      ButtonTheme     =   6
      CaptionEffect   =   3
      BackColor       =   2504331
      BackColorPressed=   4349166
      BackColorHover  =   4678655
      Caption         =   "Search"
      Picture         =   "Odesafastsearch.frx":26813
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.Button Command2 
      Height          =   495
      Left            =   5400
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ButtonStyle     =   6
      ButtonStyleColors=   3
      ButtonTheme     =   6
      CaptionEffect   =   3
      BackColor       =   2504331
      BackColorPressed=   4349166
      BackColorHover  =   4678655
      Caption         =   "Stop"
      Picture         =   "Odesafastsearch.frx":26B44
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Include directory names"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5340
      TabIndex        =   15
      Top             =   2385
      Width           =   1995
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5340
      TabIndex        =   2
      Top             =   1350
      Width           =   2400
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   7065
      Width           =   7575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Case sensitive"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5340
      TabIndex        =   3
      Top             =   2115
      Width           =   1320
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00400000&
      Height          =   3735
      ItemData        =   "Odesafastsearch.frx":2C766
      Left            =   210
      List            =   "Odesafastsearch.frx":2C768
      TabIndex        =   4
      Top             =   3285
      Width           =   7530
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7005
      TabIndex        =   1
      Text            =   "*"
      Top             =   720
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "Odesafastsearch.frx":2C76A
      Left            =   8160
      List            =   "Odesafastsearch.frx":2C76C
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3600
      Width           =   420
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   2760
      Hidden          =   -1  'True
      Left            =   2775
      System          =   -1  'True
      TabIndex        =   7
      Top             =   90
      Width           =   2490
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00400000&
      Height          =   2790
      Left            =   210
      TabIndex        =   6
      Top             =   90
      Width           =   2490
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5340
      TabIndex        =   0
      Top             =   360
      Width           =   2400
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Search for in-file text:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5340
      TabIndex        =   14
      Top             =   1080
      Width           =   1485
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Other options:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5340
      TabIndex        =   13
      Top             =   1845
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   810
      TabIndex        =   12
      Top             =   3015
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Matches:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   3015
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Extension:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5340
      TabIndex        =   10
      Top             =   765
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Search for filenames:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5340
      TabIndex        =   9
      Top             =   90
      Width           =   1485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Dim Alper As Boolean

Private Sub Button1_Click()
Form2.Show
End Sub

Private Sub Command1_Click()
Alper = False
search

End Sub

'Dim Alper ESKIKILIC As Visual Basic Programmer
'Dim Go.to url = "www.odesayazilim.com"
'Very Very Fast File Search System
'Made In Turkey
'www.odesayazilim.com


Private Sub Command2_Click()
Alper = True

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path

End Sub

Private Sub Form_Load()
On Error Resume Next
If PathFileExists(App.Path & "\odesafastsearch.ini") = 1 Then
  Open App.Path & "\odesafastsearch.ini" For Input As #1
    Line Input #1, X
    Dir1.Path = X
    Line Input #1, X
    Check1.Value = Val(X)
    Line Input #1, X
    Check2.Value = Val(X)
  Close #1
Else
  Dir1.Path = Mid(App.Path, 1, InStr(1, App.Path, "\"))
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Open App.Path & "\odesafastsearch.ini" For Output As #1
  Print #1, Dir1.Path
  Print #1, CStr(Check1.Value)
  Print #1, CStr(Check2.Value)
Close #1

End

End Sub

Private Sub List2_Click()
Text3.Text = List2.List(List2.ListIndex)

End Sub

Private Sub List2_DblClick()
MsgBox List2.List(List2.ListIndex)

End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text2_Change()
File1.Pattern = "*." & Text2.Text

End Sub

Private Sub search()
Dim i As Integer, tI As Integer, X As Integer, oDir As String

List1.Clear
List2.Clear
Label4.Caption = "0"
List1.AddItem Dir1.Path
oDir = Dir1.Path

tI = 1

While tI <> 0
  For i = 0 To tI - 1
    Dir1.Path = List1.List(i)
    DoEvents
    If File1.ListCount <> 0 Then
      If Check1.Value = 1 Then
        checkListCS
      Else
        checkList
      End If
    End If
    For X = 0 To Dir1.ListCount - 1
      List1.AddItem Dir1.List(X)
    Next X
  Next i
  For i = 0 To tI - 1
    List1.RemoveItem 0
  Next i
  tI = List1.ListCount
  If Alper = True Then GoTo quit
Wend

quit:
Close #2
Dir1.Path = oDir
Dir1.Refresh

End Sub

Private Sub checkListCS()
Dim temp As String, i As Integer, add As Byte

If InStr(1, Dir1.Path, Text1.Text) <> 0 And Text4.Text = Empty And Check2.Value = 1 Then List2.AddItem Dir1.Path

For i = 0 To File1.ListCount - 1
  add = 0
  If InStr(1, File1.List(i), Text1.Text) <> 0 Then add = 1

  If Text4.Text <> Empty And add = 1 Then
    temp = String(FileLen(IIf(Right(Dir1.Path, 1) = "\", Dir1.Path & File1.List(i), Dir1.Path & "\" & File1.List(i))), Chr(0))
    
    Open IIf(Right(Dir1.Path, 1) = "\", Dir1.Path & File1.List(i), Dir1.Path & "\" & File1.List(i)) For Binary As #1
      Get #1, , temp
    Close #1
    
    If InStr(1, temp, Text4.Text) <> 0 Then
      add = 1
    Else
      add = 0
    End If
  End If
  
  If add = 1 Then List2.AddItem IIf(Right(Dir1.Path, 1) = "\", Dir1.Path & File1.List(i), Dir1.Path & "\" & File1.List(i))
Next i

Label4.Caption = List2.ListCount

End Sub

Private Sub checkList()
Dim temp As String, i As Integer, add As Integer

If InStr(1, LCase(Dir1.Path), LCase(Text1.Text)) <> 0 And Text4.Text = Empty And Check2.Value = 1 Then List2.AddItem Dir1.Path

For i = 0 To File1.ListCount - 1
  add = 0
  If InStr(1, LCase(File1.List(i)), LCase(Text1.Text)) <> 0 Then add = 1

  If Text4.Text <> "" And add = 1 Then
    temp = String(FileLen(IIf(Right(Dir1.Path, 1) = "\", Dir1.Path & File1.List(i), Dir1.Path & "\" & File1.List(i))), Chr(0))
    
    Open IIf(Right(Dir1.Path, 1) = "\", Dir1.Path & File1.List(i), Dir1.Path & "\" & File1.List(i)) For Binary As #1
      Get #1, , temp
    Close #1
    
    If InStr(1, LCase(temp), LCase(Text4.Text)) <> 0 Then
      add = 1
    Else
      add = 0
    End If
  End If
  
  If add = 1 Then List2.AddItem IIf(Right(Dir1.Path, 1) = "\", Dir1.Path & File1.List(i), Dir1.Path & "\" & File1.List(i))
Next i

Label4.Caption = List2.ListCount

End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)

End Sub
