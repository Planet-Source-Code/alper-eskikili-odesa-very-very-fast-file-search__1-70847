VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About Code"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9525
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   2970
      Left            =   240
      Picture         =   "Form2.frx":0000
      Top             =   0
      Width           =   9000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000F&
      X1              =   0
      X2              =   9480
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Image Image3 
      Height          =   1080
      Left            =   6120
      Picture         =   "Form2.frx":55B9
      Top             =   3120
      Width           =   2250
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   360
      Picture         =   "Form2.frx":8F38
      Top             =   3360
      Width           =   420
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
      Caption         =   "Programmed By: Alper ESKÝKILIÇ    E-Mail: odesayazilim@gmail.com "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Odesa Yazýlým Fast File Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000F&
      X1              =   4680
      X2              =   4680
      Y1              =   3000
      Y2              =   4320
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
