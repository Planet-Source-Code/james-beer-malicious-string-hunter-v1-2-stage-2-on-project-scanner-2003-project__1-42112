VERSION 5.00
Begin VB.Form Proposial 
   Caption         =   "Proposial for Visual Basic users on PSC!"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   Icon            =   "Proposial.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4425
   ScaleWidth      =   7050
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Back to Program"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3960
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1605
      Left            =   240
      Picture         =   "Proposial.frx":0442
      ScaleHeight     =   107
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   310
      TabIndex        =   0
      Top             =   120
      Width           =   4650
   End
   Begin VB.Label Label2 
      Caption         =   $"Proposial.frx":1942
      Height          =   1575
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   $"Proposial.frx":1A05
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   6495
   End
End
Attribute VB_Name = "Proposial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Proposial
End Sub
