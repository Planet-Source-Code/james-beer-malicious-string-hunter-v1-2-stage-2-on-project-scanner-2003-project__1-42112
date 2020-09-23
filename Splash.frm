VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   LinkTopic       =   "Form2"
   ScaleHeight     =   5280
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Don't Show this Again"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   1935
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005EC6F2&
      Height          =   2220
      ItemData        =   "Splash.frx":0000
      Left            =   360
      List            =   "Splash.frx":000D
      TabIndex        =   2
      Top             =   2400
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1605
      Left            =   1080
      Picture         =   "Splash.frx":0085
      ScaleHeight     =   1605
      ScaleWidth      =   4650
      TabIndex        =   0
      Top             =   240
      Width           =   4650
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "[Project Scanner 2003] Project Credits:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   120
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
On Error Resume Next
MSH_Settings.ShowSplash = Check1.Value
Kill App.Path & "\Settings.MPC" 'To make sure it doesn't just
'reallocate data into the file instead create a new one.
'Another way would be to buffer the data but that's pointless
'for this kind of file
Open App.Path & "\Settings.MPC" For Binary Access Write Lock Write As #3
Put #3, , MSH_Settings
Close #3
End Sub

Private Sub Command1_Click()
Form1.Enabled = True
Form1.Show
Unload Splash
End Sub
