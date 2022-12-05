VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BeApp Programs - Malicious String Hunter v1.2 [Project Scanner 2003 Project] by James Beer"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11145
   Icon            =   "MSH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   550
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   743
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add_File"
            Object.ToolTipText     =   "Add a new File"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Remove_File"
            Object.ToolTipText     =   "Remove Selected File"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Scan_Files"
            Object.ToolTipText     =   "Scan Files"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Show_Report"
            Object.ToolTipText     =   "Show Existing Report"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Erase_Info"
            Object.ToolTipText     =   "Erase Information"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Options"
            Object.ToolTipText     =   "General Settings"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar MSHStatus 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   7890
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   13124
            MinWidth        =   13124
            Picture         =   "MSH.frx":0442
            Text            =   "Not Ready"
            TextSave        =   "Not Ready"
         EndProperty
      EndProperty
   End
   Begin VB.Frame ShowWork 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   3840
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Shape Working_Button_OL 
         BorderWidth     =   3
         Height          =   255
         Left            =   2280
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Malicious String Hunter is Working"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   3735
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   9
         Left            =   1920
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   8
         Left            =   1800
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   7
         Left            =   1680
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   6
         Left            =   1560
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   5
         Left            =   1440
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   1320
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   1200
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   1080
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   960
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   840
         Top             =   1080
         Width           =   135
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   240
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanning in Progress... Please Wait"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Working_Button 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Left            =   2280
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   0
         Left            =   120
         Top             =   360
         Width           =   3495
      End
   End
   Begin MSComctlLib.ListView MSL 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   4080
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Malicious/Accessory String"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Line"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Class"
         Object.Width           =   2664
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1605
      Left            =   6405
      Picture         =   "MSH.frx":0A18
      ScaleHeight     =   107
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   310
      TabIndex        =   23
      Top             =   480
      Width           =   4650
   End
   Begin VB.TextBox CodeLine_Label 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   22
      Top             =   7080
      Width           =   9855
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   9960
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer LightEffect1 
      Interval        =   25
      Left            =   9960
      Top             =   1920
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9960
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":1F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":2B6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":3177
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":374D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":3D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":49D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":5622
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":6274
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":6EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":7B18
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":876A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":8BBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":900E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSH.frx":9460
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ProjectFiles 
      Height          =   1455
      Left            =   5760
      TabIndex        =   2
      Top             =   2520
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Project"
         Object.Width           =   2752
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File"
         Object.Width           =   2778
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Strings Found"
         Object.Width           =   2099
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Scan Information/Progress"
      Height          =   3255
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   5415
      Begin VB.TextBox Path_Label 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1200
         Width           =   5175
      End
      Begin MSComctlLib.ListView FTSList 
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1508
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Path"
            Object.Width           =   5010
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Filename"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton Scan_Files 
         Caption         =   "Scan File(s)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   20
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton Remove_File 
         Caption         =   "Remove File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton Add_File 
         Caption         =   "Add File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Files to Scan:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   5775
      End
      Begin VB.Label FileType_Label 
         BackStyle       =   0  'Transparent
         Caption         =   "<No File>"
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "File Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Size_Label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0 KB"
         Height          =   255
         Left            =   4080
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Located At:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.Label File_Label 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Process File:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Malicious String List:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Line:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Files (If Project File Loaded): "
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   7335
      Left            =   120
      Top             =   480
      Width           =   10935
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_File_AddFile 
         Caption         =   "Add &File/Project"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu_File_RemoveFile 
         Caption         =   "Remove Selected File/Project"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu_File_Scan 
         Caption         =   "&Scan File(s)/Project(s)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_File_Propose 
         Caption         =   "Read Proposial!!!"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnu_Info 
      Caption         =   "&Information"
      Begin VB.Menu mnu_Info_Clear 
         Caption         =   "Clear Information"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnu_Info_Report 
         Caption         =   "Display Existing Report"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
      Begin VB.Menu mnu_Options_Settings 
         Caption         =   "General Settings"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnu_About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FTSSelectedItem

Private Function ScanFilesListed()
Dim i As Long
Dim t As String

MSHStatus.Panels(1).Picture = Form1.ImageList1.ListImages.Item(5).Picture
MSHStatus.Panels(1) = "Scanning in Progress..."

If InUse = False Then
    For i = 1 To 6
        Toolbar1.Buttons.Item(i).Enabled = False
    Next i
    mnu_File_AddFile.Enabled = False
    mnu_File_RemoveFile.Enabled = False
    mnu_File_Scan.Enabled = False
    mnu_Info_Clear.Enabled = False
    mnu_Info_Report.Enabled = False
    mnu_Options_Settings.Enabled = False

    InUse = True
    EndScan = False

    ReDim GetErrors(0)

    For i = 0 To MS_Total
        ReportI.MSOccur(i) = 0
    Next i

    ReportI.TOTALLines = 0
    ReportI.TotalMaliciousCount = 0
    ReportI.POTENTIAL = 0
    ReportI.SUSPICIOUS = 0
    ReportI.CAUTION = 0
    ReportI.WARNING = 0
    ReportI.DANGER = 0
    ReportI.DESTRUCTIVE = 0
    ReportI.TOTALBYTES = 0

    ProgressLight = 0
    ShowWork.Visible = True
    LightEffect1.Enabled = True

    DoEvents
    
    CodeLine_Label = Empty
    ClearList 'clears the lists

    For i = 1 To UBound(FilesToScan)
        Path_Label = FilesToScan(i).Pathname
        Scan FilesToScan(i).Pathname, FilesToScan(i).Filename
    Next i

    ShowWork.Visible = False
    LightEffect1.Enabled = False

    If EndScan = False Then
        CreateReport
        InUse = False
            For i = 1 To 6
            Toolbar1.Buttons.Item(i).Enabled = True
            Next i
        mnu_File_AddFile.Enabled = True
        mnu_File_RemoveFile.Enabled = True
        mnu_File_Scan.Enabled = True
        mnu_Info_Clear.Enabled = True
        mnu_Info_Report.Enabled = True
        mnu_Options_Settings.Enabled = True
    Else
        InUse = False
        Path_Label = Empty
        File_Label = Empty
        ClearList
        MSHStatus.Panels.Item(1) = "Scan Cancelled - Awaiting new scan"
        MSHStatus.Panels.Item(1).Picture = Form1.ImageList1.ListImages.Item(3).Picture
    
    End If
    
Else

    MSHStatus.Panels.Item(1) = "Scan is in Use or Failed to Shutdown - Cannot Reinitiate"
    MSHStatus.Panels.Item(1).Picture = Form1.ImageList1.ListImages.Item(4).Picture
    
End If
End Function

Private Sub Add_File_Click()
AddFile
End Sub


Private Sub Form_Load()
Open App.Path & "\Settings.MPC" For Binary Access Read Lock Read As #3
Get #3, , MSH_Settings
Close #3

ReDim FilesToScan(0)
MSHStatus.Panels.Item(1) = "Not Ready"
MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(3).Picture
Image1(0).Picture = ImageList1.ListImages.Item(5).Picture
Load_MSData 'Obtains the String Reference file to this program

If MSH_Settings.ShowSplash = 0 Then
Form1.Hide
Load Splash
Splash.Show
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
TerminateApp
End Sub



Private Sub LightEffect1_Timer()
Dim i As Long

'This progress bar effect is for program that have undetermined
'processing times, it indicates only to the user that it's
'not frozen

ProgressLight = ProgressLight + 1

If ProgressLight = 10 Then ProgressLight = 0
For i = 0 To 9
    Form1.Shape2(i).FillColor = vbBlack
Next i

If ProgressLight > 0 Then
    Shape2(ProgressLight - 1).FillColor = &HC000&
Else
    Shape2(9).FillColor = &HC000&
End If

If ProgressLight > 1 Then
    Shape2(ProgressLight - 2).FillColor = &H8000&
Else
    Shape2(9 - (1 - ProgressLight)).FillColor = &H8000&
End If

Shape2(ProgressLight).FillColor = vbGreen
End Sub

Private Sub mnu_About_Click()
ShowSplash
End Sub

Private Sub mnu_File_AddFile_Click()
AddFile
End Sub

Private Sub mnu_File_Exit_Click()
Dim x As Integer

If InUse = True Then
MsgBox "A Scan is still running, Are you sure?", vbExclamation + vbYesNo, "Scanning In Progress"
If x = vbYes Then TerminateApp
Else
TerminateApp
End If

End Sub

Private Sub mnu_File_Propose_Click()
Load Proposial
Proposial.Show
End Sub

Private Sub mnu_File_RemoveFile_Click()
RemoveFile
End Sub

Private Sub mnu_File_Scan_Click()
ScanFilesListed
End Sub

Private Sub mnu_Info_Clear_Click()
If InUse = False Then ClearList
End Sub

Private Sub mnu_Info_Report_Click()
If InUse = False Then
ReDim GetErrors(0)
CreateReport
End If
End Sub

Private Sub mnu_Options_Settings_Click()
If InUse = False Then
ShowOptions
End If
End Sub

Private Sub MSL_ItemClick(ByVal Item As MSComctlLib.ListItem)
CodeLine_Label = Item.SubItems(2)
CodeLine_Label.ForeColor = Item.ListSubItems(2).ForeColor
End Sub

Private Sub ProjectFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim i As Long
For i = 1 To ProjectFiles.ListItems.Count
If i <> Item.Index Then
ProjectFiles.ListItems.Item(i).ForeColor = vbBlack
ProjectFiles.ListItems.Item(i).Bold = False
ProjectFiles.ListItems.Item(i).ListSubItems.Item(1).ForeColor = vbBlack
ProjectFiles.ListItems.Item(i).ListSubItems.Item(1).Bold = False
ProjectFiles.ListItems.Item(i).ListSubItems.Item(2).ForeColor = vbBlack
ProjectFiles.ListItems.Item(i).ListSubItems.Item(2).Bold = False
Else
ProjectFiles.ListItems.Item(i).ForeColor = vbBlue
ProjectFiles.ListItems.Item(i).Bold = True
ProjectFiles.ListItems.Item(i).ListSubItems.Item(1).ForeColor = vbBlue
ProjectFiles.ListItems.Item(i).ListSubItems.Item(1).Bold = True
ProjectFiles.ListItems.Item(i).ListSubItems.Item(2).ForeColor = vbBlue
ProjectFiles.ListItems.Item(i).ListSubItems.Item(2).Bold = True
End If
Next i
For i = 1 To MSL.ListItems.Count
    If MSL.ListItems.Item(i) = Item.SubItems(1) Then
        MSL.ListItems.Item(i).ForeColor = vbBlue
    Else
        MSL.ListItems.Item(i).ForeColor = vbBlack
    End If
Next i
ProjectFiles.Refresh
MSL.Refresh
End Sub

Private Sub Remove_File_Click()
RemoveFile
End Sub

Private Sub Scan_Files_Click()
ScanFilesListed
End Sub

Private Sub ShowWork_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Working_Button.BackColor = &HC0C0C0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case Is = "Add_File"
AddFile
Case Is = "Remove_File"
RemoveFile
Case Is = "Scan_Files"
ScanFilesListed
Case Is = "Show_Report"
ReDim GetErrors(0)
CreateReport
Case Is = "Erase_Info"
ClearList
Case Is = "Options"
ShowOptions
Case Is = "Syntax_Validation"
ShowSyntaxValidate
End Select
End Sub

Private Sub Working_Button_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
EndScan = True
End Sub

Function Scan(ByVal PathN As String, ByVal FileN As String)
Dim i As Long
Dim x As Long

Filename = FileN
Pathname = PathN

'I seperated the path and file data into two variables because
'I wanted the list to only include file names of where that
'string orginated from

If UCase(Right(Filename, 4)) = ".VBP" Then
ProjectFile = Filename
GetVBPFiles Pathname & Filename
i = 0
    Do Until i >= UBound(SubFiles)
    'since the loop is based on a dynamic array which
    'expands as it's loops, DO...LOOP is the only other
    'fastest way other than FOR...NEXT
    i = i + 1
        Filename = SubFiles(i) 'The files of which this
        'project file has
        
        Size_Label = GetFileSizeProper(FileLen(Pathname & Filename))
        ReportI.TOTALBYTES = ReportI.TOTALBYTES + FileLen(Pathname & Filename)
        'NOTE: The buffer only handles the amount of bytes
        'for each file as they're processed, then it's
        'erased and refilled with the next file's data
        'the buffer doesn't expand to accompany all files.
        'Only the one it's scanning.
        
        LoadFile Pathname, Filename 'loads and scans the file instantly
        ProjectFiles.ListItems.Add , , ProjectFile
        ProjectFiles.ListItems.Item(ProjectFiles.ListItems.Count).SubItems(1) = SubFiles(i)
        ProjectFiles.ListItems.Item(ProjectFiles.ListItems.Count).SubItems(2) = MaliciousCount
        ProjectFiles.ListItems.Item(ProjectFiles.ListItems.Count).SmallIcon = IIf(MaliciousCount > 0, 3, 1)
    If i > 500 Then Exit Do ' A failsafe if it loops
    'out of control, Doevents slows it down too much
    'so this is a alternative way to prevent lock-ups.
    'NO Visual Basic project should ever excess this amount
    'Well I've never seen one that does anyway.
    Loop
    Filename = ProjectFile 'Since the scan is primarily based
    'on the project file it should give it's name back in case
    'of a re-scan.
Else
ProjectFile = Empty
Size_Label = GetFileSizeProper(FileLen(Pathname & Filename))
ReportI.TOTALBYTES = ReportI.TOTALBYTES + FileLen(Pathname & Filename)
LoadFile Pathname, Filename 'loads and scans the file instantly
End If
End Function

Function GetFileSizeProper(ByVal Size As Single) As String
Select Case Size
Case Is < 1024
GetFileSizeProper = Size & " Bytes"
Case 1024 To ((1024 ^ 2) - 1)
GetFileSizeProper = CCur(Size / 1024) & " KB"
Case Is >= (1024 ^ 2)
GetFileSizeProper = CCur(Size / (1024 ^ 2)) & " MB"
End Select
End Function

Function AddFile()
Dim x As Integer
Dim i As Long
Dim Existing As Boolean
CD1.Filename = Empty 'So if you cancel it, it will have
'no filename string pre-stored therefore won't execute
'the add file process
CD1.DialogTitle = "Select File"
CD1.Filter = "Visual Basic 6.0 Projects|*.vbp|" & _
             "Visual Basic 6.0 Forms|*.frm|" & _
             "Visual Basic 6.0 Modules|*.bas|" & _
             "Visual Basic 6.0 Class Modules|*.cls|" & _
             "Visual Basic 6.0 Resource File|*.res|" & _
             "Visual Basic 6.0 Form Binary Files|*.frx|" & _
             "Text Document|*.txt|" & _
             "VB Script|*.vbs"
'File filters, in case you're a newbie interested these define
'the file types that should be allowed to be added.
CD1.ShowOpen 'Easy isn't it, one line give you about
             '30 lines of action plus even more
If CD1.FilterIndex <> 1 And CD1.FilterIndex <> 7 And _
   CD1.FilterIndex <> 8 And CD1.Filename <> Empty Then
'For if you chose a VB file which really doesn't
'need to be added as a project file could be referenced to it
x = MsgBox("You can scan it's project file instead to scan every linked file in it" & Chr(13) & _
           "this is more efficient and save adding every file, Do you still want to do this?", vbExclamation + vbYesNo, "Optimize Scanning Tip")
If x = vbNo Then Exit Function
End If

If CD1.Filename <> Empty And Existing = False Then
For i = 1 To UBound(FilesToScan)
'to find if an existing file is already added
'this prevent wasting memory on the same file
If Left(CD1.Filename, Len(CD1.Filename) - Len(CD1.FileTitle)) = FilesToScan(i).Pathname And _
   CD1.FileTitle = FilesToScan(i).Filename Then
Existing = True
End If
If Existing = True Then Exit For
Next i

If Existing = True Then MsgBox "Duplicate file and path found, you don't need to scan it twice", vbExclamation, "File/Path conflict": Exit Function

'Add the file to the array which is processed
ReDim Preserve FilesToScan(UBound(FilesToScan) + 1)
FilesToScan(UBound(FilesToScan)).Pathname = Left(CD1.Filename, Len(CD1.Filename) - Len(CD1.FileTitle))
FilesToScan(UBound(FilesToScan)).Filename = CD1.FileTitle
FTSList.ListItems.Add , , FilesToScan(UBound(FilesToScan)).Pathname
FTSList.ListItems.Item(FTSList.ListItems.Count).SubItems(1) = FilesToScan(UBound(FilesToScan)).Filename

MSHStatus.Panels.Item(1) = "Ready - " & UBound(FilesToScan) & " File(s) Selected"
MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(2).Picture
End If
End Function

Function RemoveFile()
Dim i As Long
If FTSList.ListItems.Count > 0 Then
FTSList.ListItems.Remove (FTSList.SelectedItem.Index)

ReDim FilesToScan(FTSList.ListItems.Count)
For i = 1 To FTSList.ListItems.Count
FilesToScan(UBound(FilesToScan)).Pathname = FTSList.ListItems.Item(i)
FilesToScan(UBound(FilesToScan)).Filename = FTSList.ListItems.Item(i).SubItems(1)
Next i
End If

If FTSList.ListItems.Count = 0 Then
ReDim FilesToScan(0)
MSHStatus.Panels.Item(1) = "Not Ready"
MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(3).Picture
End If
End Function

Function CreateReport()
Dim i As Long
'This generate a brief report which is easier to
'understand, the option to save a log of this
'scan is optional as long as there's malicious code
'found.
'This acts like a error handler and debug too
Load ReportWindow
ReportWindow.Show
Form1.Enabled = False

With ReportWindow
For i = 0 To MS_Total
If ReportI.MSOccur(i) > 0 Then
.Report_MSL.ListItems.Add , , MS(i).MStr
.Report_MSL.ListItems.Item(.Report_MSL.ListItems.Count).SubItems(1) = MS(i).Class
.Report_MSL.ListItems.Item(.Report_MSL.ListItems.Count).SubItems(2) = ReportI.MSOccur(i)
End If
Next i
.Result(0) = ReportI.TotalMaliciousCount
.Result(1) = ReportI.POTENTIAL
.Result(2) = ReportI.SUSPICIOUS
.Result(3) = ReportI.CAUTION
.Result(4) = ReportI.WARNING
.Result(5) = ReportI.DANGER
.Result(6) = ReportI.DESTRUCTIVE
.Result(7) = ReportI.TOTALLines
.Result(8) = GetFileSizeProper(ReportI.TOTALBYTES)

If MSL.ListItems.Count > 0 Then
'this show the report that the scan worked successfully but
'found some malicious strings :(
.ScanDoneLabel = "Scan Completed Successfully - Malicious/Accessory Strings Found"
.Image2.Picture = Form1.ImageList1.ListImages.Item(3).Picture
.Report_Header = "This code contains malicious/Accessory code and should not be executed without Investigating"
.Report_Information.Visible = True
.Report_Button(0).Visible = True
.Report_Button_OL(0).Visible = True
MSHStatus.Panels.Item(1) = "Scan Completed Successfully - Malicous/Accessory Strings Found"
MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(3).Picture
Else
'same thing except no malicious strings were found :)
.ScanDoneLabel = "Scan Completed Successfully - Data is Clean"
.Image2.Picture = Form1.ImageList1.ListImages.Item(1).Picture
.Report_Header = "This code is considered clean by MSH program resources"
.Report_Information.Visible = True
.Report_Button(0).Visible = False
.Report_Button_OL(0).Visible = False
MSHStatus.Panels.Item(1) = "Scan Completed Successfully - All Clean"
MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(2).Picture
End If

If UBound(FilesToScan) = 0 Then
'If the filestoscan array has no files listed
.ScanDoneLabel = "Scan Cannot Initiate - No File Specified"
.Image2.Picture = Form1.ImageList1.ListImages.Item(2).Picture
.Report_Header = "No projects or files were selected to scan, please add files by using the Add File button"
.Report_Information.Visible = False
.Report_Button(0).Visible = False
.Report_Button_OL(0).Visible = False
Beep
MSHStatus.Panels.Item(1) = "Not Ready - Error"
MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(3).Picture
End If

If UBound(GetErrors) > 0 Then
'if it completed but encountered Visual Basic Run-time errors
.ScanDoneLabel = "Scan Completed but Errors Occurred"
.Image2.Picture = Form1.ImageList1.ListImages.Item(4).Picture
.Report_Header = "The Follow Errors Occurred while Scanning:"
For i = 1 To UBound(GetErrors)
.Report_Header = .Report_Header & Chr(13) & Chr(10) & _
                             GetErrors(i).Def & ":" & GetErrors(i).File
.Report_Information.Visible = True
.Report_Button(0).Visible = True
.Report_Button_OL(0).Visible = True
Next i

MSHStatus.Panels.Item(1) = "Scan Complete but with Errors"
MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(4).Picture
End If

On Error Resume Next
If EndScan = True Or UBound(FilesToScan) = 0 Then Exit Function
If MSH_Settings.AutoCreateLog = True Then
    .Report_Button(0).Visible = False
    .Report_Button_OL(0).Visible = False
    
    CD1.Filename = App.Path & "\" & MSH_Settings.DefaultLogFilename & ".log"
    FileLen App.Path & "\" & MSH_Settings.DefaultLogFilename & ".log"
    
    If Err.Number <> 53 And MSH_Settings.OverWriteLog = False Then
    Err.Clear 'Using this clear the last error that was
    'created so the next function can regenerate it.
        i = 0
            Do Until Err.Number = 53 Or i > 500
            'Obtains which increment number has been used
            '* Instead of open statement use filelen because
            'it generates the same error and save you time
            i = i + 1
            FileLen App.Path & "\" & MSH_Settings.DefaultLogFilename & i & ".log"
                DoEvents
            Loop
        CD1.Filename = App.Path & "\" & MSH_Settings.DefaultLogFilename & i & ".log"
    End If
ReportWindow.CreateLog
End If

End With
End Function

Function ShowOptions()
Form1.Enabled = False
Load Options

With Options
.Show
'This adds the data if stored from previous changes
.Check1.Value = (Not MSH_Settings.AutoCreateLog) + 1
 
If .Check1.Value = 1 Then
.Frame2.Enabled = True
.Report_Filename.BackColor = vbWhite
.ReportFileOps(0).ForeColor = vbBlack
.ReportFileOps(1).ForeColor = vbBlack
Else
.Frame2.Enabled = False
.Report_Filename.BackColor = &HC0C0C0
.ReportFileOps(0).ForeColor = &HC0C0C0
.ReportFileOps(1).ForeColor = &HC0C0C0
End If

If MSH_Settings.LogCreateType = 2 Then
.ReportType(0).Value = True
ElseIf MSH_Settings.LogCreateType = 1 Then
.ReportType(1).Value = True
ElseIf MSH_Settings.LogCreateType = 0 Then
.ReportType(2).Value = True
End If

'To store the default filename in memory for easy access
.Report_Filename = MSH_Settings.DefaultLogFilename

'If it's going to overwrite an existing file if found,
'otherwise increments are added.

If MSH_Settings.OverWriteLog = True Then
    .ReportFileOps(0).Value = True
ElseIf MSH_Settings.OverWriteLog = False Then
    .ReportFileOps(1).Value = True
End If

'Exclude classes, this is if you don't want to waste your time
'scrolling through a list of these classes if found.
.Option_ExcludeClass(0).Value = (Not MSH_Settings.ExcludePOTENTIALClass) + 1
.Option_ExcludeClass(1).Value = (Not MSH_Settings.ExcludeSUSPICIOUSClass) + 1
End With
End Function

Function ShowSplash()
Form1.Enabled = False
Load Splash
Splash.Show
Splash.Check1 = MSH_Settings.ShowSplash
End Function
