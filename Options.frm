VERSION 5.00
Begin VB.Form Options 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "General Settings"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   LinkTopic       =   "Form2"
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   513
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Scan Options:"
      Height          =   1935
      Left            =   3960
      TabIndex        =   11
      Top             =   480
      Width           =   3495
      Begin VB.CheckBox Option_ExcludeClass 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exclude Suspicious Class Strings"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   3255
      End
      Begin VB.CheckBox Option_ExcludeClass 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exclude Potential Class Strings"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "* Note: I realised it not really a good idea to make it delete certains strings as some programs depend on them."
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Report Settings:"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3735
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Report File:"
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   3495
         Begin VB.OptionButton ReportFileOps 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Create Increments (e.g File0001.log)"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   3255
         End
         Begin VB.OptionButton ReportFileOps 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Overwrite Existing File"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox Report_Filename 
            Height          =   285
            Left            =   1800
            TabIndex        =   8
            Text            =   "ScanReportLog"
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Default Log Filename:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create Report Log Automatically"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   3495
      End
      Begin VB.OptionButton ReportType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Basic - Summary of Scan Only"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   3495
      End
      Begin VB.OptionButton ReportType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Advanced - Files/Malicious Strings List Only"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3495
      End
      Begin VB.OptionButton ReportType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Full"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Log Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Label Options_Tool_Tip 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   3840
      Width           =   7695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Finally finished it, Enjoy :)   Note: A feature was remove because it wasn't necessary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   17
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4320
      Top             =   2520
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   24
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   1
      X1              =   512
      X2              =   512
      Y1              =   0
      Y2              =   24
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   512
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   3
      X1              =   0
      X2              =   512
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Shape Options_Button_OL 
      BorderWidth     =   3
      Height          =   255
      Index           =   1
      Left            =   6240
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Shape Options_Button_OL 
      BorderWidth     =   3
      Height          =   255
      Index           =   0
      Left            =   4800
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Options_Button 
      Alignment       =   2  'Center
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Options_Button 
      Alignment       =   2  'Center
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3255
      Left            =   120
      Top             =   480
      Width           =   7455
   End
   Begin VB.Label ScanDoneLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "General Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check2_Click()
If Check2.Value = 0 Then
Option_BufferLimit.Enabled = False
Option_BufferLimit = 0
Else
Option_BufferLimit.Enabled = True
Option_BufferLimit = 128
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 0 Then
Frame2.Enabled = False
Report_Filename.BackColor = &HC0C0C0
ReportFileOps(0).ForeColor = &HC0C0C0
ReportFileOps(1).ForeColor = &HC0C0C0
Else
Frame2.Enabled = True
Report_Filename.BackColor = vbWhite
ReportFileOps(0).ForeColor = vbBlack
ReportFileOps(1).ForeColor = vbBlack
End If
End Sub

Private Sub Form_Load()
Options.Image1.Picture = Form1.ImageList1.ListImages.Item(1).Picture
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
For i = 0 To 1
If Options_Button(i).BackColor <> &HC0C0C0 Then Options_Button(i).BackColor = &HC0C0C0
Next i
End Sub



Private Sub Options_Button_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
Options_Button(Index).BackColor = vbRed
Select Case Index
Case Is = 0

'Sets the log creation type
'0 = Basic only
'1 = Advanced only
'2 = Both

MSH_Settings.AutoCreateLog = Not (Check1.Value - 1)
If ReportType(0).Value = True Then
    MSH_Settings.LogCreateType = 2
ElseIf ReportType(1).Value = True Then
    MSH_Settings.LogCreateType = 1
ElseIf ReportType(2).Value = True Then
    MSH_Settings.LogCreateType = 0
End If

'To store the default filename in memory for easy access
MSH_Settings.DefaultLogFilename = Report_Filename

'If it's going to overwrite an existing file if found,
'otherwise increments are added.
If ReportFileOps(0).Value = True Then
    MSH_Settings.OverWriteLog = True
ElseIf ReportFileOps(1).Value = True Then
    MSH_Settings.OverWriteLog = False
End If

'Exclude classes, this is if you don't want to waste your time
'scrolling through a list of these classes if found.
MSH_Settings.ExcludePOTENTIALClass = Not (Option_ExcludeClass(0).Value) - 1
MSH_Settings.ExcludeSUSPICIOUSClass = Not (Option_ExcludeClass(1).Value) - 1

Kill App.Path & "\Settings.MPC" 'To make sure it doesn't just
'reallocate data into the file instead create a new one.
'Another way would be to buffer the data but that's pointless
'for this kind of file
Open App.Path & "\Settings.MPC" For Binary Access Write Lock Write As #3
Put #3, , MSH_Settings
Close #3

Case Is = 1

End Select
Form1.Enabled = True
Unload Options
End Sub

Private Sub Options_Button_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Options_Button(Index).BackColor <> vbYellow Then
For i = 0 To 1
Options_Button(i).BackColor = &HC0C0C0
Next i
Options_Button(Index).BackColor = vbYellow
End If
End Sub
