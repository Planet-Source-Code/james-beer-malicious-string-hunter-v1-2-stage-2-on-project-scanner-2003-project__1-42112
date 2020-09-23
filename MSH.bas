Attribute VB_Name = "Module1"
Option Explicit

Global Const MS_Total As Long = 74

Global Pathname As String
Global ProjectFile As String
Global Filename As String
Global MaliciousCount As Long
Global InUse As Boolean
Global EndScan As Boolean
Global ProgressLight As Long

Type ErrorData
Def As String
File As String
End Type

Type Options_Settings
LogCreateType As Byte 'to tell report generator what type of
                      'log it should make
AutoCreateLog As Boolean 'instead of clicking 'save scan report'
                         'the program will automatically create one by
                         'using the default filename
OverWriteLog As Boolean 'If it is allowed to overwrite an existing log
DefaultLogFilename As String 'Obvious
ExcludePOTENTIALClass As Boolean 'Excludes all POTENTIAL Strings
ExcludeSUSPICIOUSClass As Boolean 'Excludes all Suspicious Strings
'Everything won't get those classes information if their set true
ShowSplash As Byte
End Type

Public Enum StringFoundAction
Delete = 1
Comment = 2
Prompt = 3
Syntax_IfValid_Delete = 11
Syntax_IfValid_Comment = 12
Syntax_IfValid_Prompt = 13
End Enum

Global MSH_Settings As Options_Settings

Global GetErrors() As ErrorData 'In case my scan fails to work on something
'so the report can view it

Type Report_Info
TotalMaliciousCount As Long
TOTALLines As Long
TOTALBYTES As Single 'for debugging purposes of checking speed
POTENTIAL As Long
SUSPICIOUS As Long
CAUTION As Long
WARNING As Long
DANGER As Long
DESTRUCTIVE As Long
MSOccur(MS_Total) As Long
End Type

Public ReportI As Report_Info

Public Type Malicious_Info
    MStr As String
    Class As String
End Type

Type FileToScan_Info
Pathname As String
Filename As String
End Type

Public FilesToScan() As FileToScan_Info
Public SubFiles() As String 'Files which are found inside a project
'not programmed into the file to scan at start
Public MS(MS_Total) As Malicious_Info
Public CodeLine() As String

Function LoadFile(ByVal Path As String, ByVal File As String)
Dim MaxLine As Long 'To Manage the CodeLine Ubound easier
Dim MaxFile As Long 'To Manage the Files Ubound easier
Dim c As Long
Dim TCL As String 'Ucase Temp for CodeLine
Dim FRXLoaded As Boolean 'If this file is a form and
'it's FRX binary file is loaded

On Error GoTo Err_Check
MaliciousCount = 0
ReDim CodeLine(0)
Close #1
Open Path & File For Input As #1
Form1.File_Label = File
Form1.FileType_Label = GetFileType(File)
Do Until EOF(1) Or EndScan = True

ReDim Preserve CodeLine(UBound(CodeLine) + 1)
Line Input #1, CodeLine(UBound(CodeLine))
MaxLine = UBound(CodeLine)
TCL = UCase(CodeLine(MaxLine))

If Right(File, 4) = ".FRM" And FRXLoaded = False Then
c = InStr(1, TCL, Left(File, Len(File) - 4) & ".FRX") 'Searchs for
'Form Binary File Extensions
    If c > 0 Then
        ReDim Preserve SubFiles(UBound(SubFiles) + 1)
        SubFiles(UBound(SubFiles)) = Left(File, Len(File) - 4) & ".FRX"
        'Adds the file to the list
        FRXLoaded = True
    End If
End If

ScanCodeLine (UBound(CodeLine)) ' this is the line which does the scan
Loop
Close #1

Exit Function
Err_Check:
ReDim Preserve GetErrors(UBound(GetErrors) + 1)
GetErrors(UBound(GetErrors)).Def = Err.Description
GetErrors(UBound(GetErrors)).File = "LoadFile: " & File
Exit Function
End Function

Function GetVBPFiles(ByVal ProjectFile As String)
Dim i As Long
On Error GoTo Err_Check
'this allows you to just give the project file and my program will handle the rest
'not bad eh?
'However if you do select a form file to scan it will bypass this and that form file
'only is checked and this applies to all other VB files, hell you can even use it
'on documents and vbscript files if you wanted to.
'Next Version will extend to cover VB Group Project Files as well

'the data can be both lower or upper case as it converted to upper case anyway

Dim MaxLine As Long 'To Manage the CodeLine Ubound easier
Dim MaxFile As Long 'To Manage the Files Ubound easier
Dim TCL As String 'Temp for CodeLine Upper-cased
Dim TSF As String 'Temp for Sub File

Dim x As Long 'type of VB file to add
Dim c As Long 'colon ; location if field is Class or Module, because they
'store their object name first then their file name
'or it's used as a instr locater too.

ReDim SubFiles(0)
ReDim CodeLine(0)

Open ProjectFile For Input As #1
    Do Until EOF(1)
    
    ReDim Preserve CodeLine(UBound(CodeLine) + 1)
    
    MaxLine = UBound(CodeLine)
    
    Line Input #1, CodeLine(MaxLine)
    TCL = UCase(CodeLine(MaxLine))
    'This is a optimization trick, don't ever use ucase more than once
    'instead store it in a non-array string so the code that needs this
    'info will be able to load it faster
    '(Using arrays in InStr slows it down a bit)

    x = 0
    If InStr(1, TCL, "FORM=") = 1 Then x = 1
    If InStr(1, TCL, "MODULE=") = 1 Then x = 2
    If InStr(1, TCL, "CLASS=") = 1 Then x = 3
    If InStr(1, TCL, "RESFILE32=") = 1 Then x = 4

        If x > 0 Then

        Do While c > 0
            c = InStr(1, TCL, Chr(34))
            If c = 0 Then Exit Do
            Mid(TCL, c, 1) = " "
        Loop

    ReDim Preserve SubFiles(UBound(SubFiles) + 1)
    MaxFile = UBound(SubFiles)
    If x = 1 Then TSF = Mid(TCL, 6, Len(TCL) - 5)
    If x = 4 Then TSF = Mid(TCL, 11, Len(TCL) - 10)
        If x = 2 Or x = 3 Then
            c = InStr(1, TCL, "; ")
                If c > 0 Then
                TSF = Mid(TCL, (c + 2), Len(TCL) - (c + 1))
                End If
        End If
    End If
    SubFiles(MaxFile) = Trim(TSF)
    Loop
    Close #1
   
Exit Function
Err_Check:
ReDim Preserve GetErrors(UBound(GetErrors) + 1)
GetErrors(UBound(GetErrors)).Def = Err.Description
GetErrors(UBound(GetErrors)).File = ProjectFile
Exit Function
End Function


Function AddToList(ByVal MSNum As Long, ByVal CLNum As Long)
Dim x As Long
Dim IconNum As Long

If MSH_Settings.ExcludePOTENTIALClass = True And MS(MSNum).Class = "POTENTIAL" Or _
   MSH_Settings.ExcludeSUSPICIOUSClass = True And MS(MSNum).Class = "SUSPICIOUS" Then Exit Function
'To make it more efficient and not look for possible malicious strings
'you can exclude two of the six classes so only caution or higher classed
'are reported.

'IF...Elseif...End if are said to be 3% faster than select case
'I think because it streams for the right answer rather than
'look for it straight away.
'Use Else statement for the equilivent of Case Else
If MS(MSNum).Class = "POTENTIAL" Then
x = vbBlack: IconNum = 7
ElseIf MS(MSNum).Class = "SUSPICIOUS" Then
x = vbBlack: IconNum = 7
ElseIf MS(MSNum).Class = "CAUTION" Then
x = 64: IconNum = 2
ElseIf MS(MSNum).Class = "WARNING" Then
x = 128: IconNum = 3
ElseIf MS(MSNum).Class = "DANGER" Then
x = 192: IconNum = 4
ElseIf MS(MSNum).Class = "DESTRUCTIVE" Then
x = 255: IconNum = 4
End If

ReportI.MSOccur(MSNum) = ReportI.MSOccur(MSNum) + 1

With Form1
    .MSL.ListItems.Add , , Filename
    .MSL.ListItems.Item(.MSL.ListItems.Count).SmallIcon = IconNum
    .MSL.ListItems.Item(.MSL.ListItems.Count).SubItems(1) = MS(MSNum).MStr
    .MSL.ListItems.Item(.MSL.ListItems.Count).SubItems(2) = CLNum & " " & CodeLine(CLNum)
    .MSL.ListItems.Item(.MSL.ListItems.Count).SubItems(3) = MS(MSNum).Class
    .MSL.ListItems.Item(.MSL.ListItems.Count).ListSubItems. _
    Item(1).ForeColor = x
    .MSL.ListItems.Item(.MSL.ListItems.Count).ListSubItems. _
    Item(2).ForeColor = x
    .MSL.ListItems.Item(.MSL.ListItems.Count).ListSubItems. _
    Item(3).ForeColor = x
End With

        MaliciousCount = MaliciousCount + 1
        ReportI.TotalMaliciousCount = ReportI.TotalMaliciousCount + 1
        
        'I presume a select case with a if...elseif...end if
        'would optimize speed of declaring what level it is
        If MS(MSNum).Class = "POTENTIAL" Then
            ReportI.POTENTIAL = ReportI.POTENTIAL + 1
        ElseIf MS(MSNum).Class = "SUSPICIOUS" Then
            ReportI.SUSPICIOUS = ReportI.SUSPICIOUS + 1
        ElseIf MS(MSNum).Class = "CAUTION" Then
            ReportI.CAUTION = ReportI.CAUTION + 1
        ElseIf MS(MSNum).Class = "WARNING" Then
            ReportI.WARNING = ReportI.WARNING + 1
        ElseIf MS(MSNum).Class = "DANGER" Then
            ReportI.DANGER = ReportI.DANGER + 1
        ElseIf MS(MSNum).Class = "DESTRUCTIVE" Then
        ReportI.DESTRUCTIVE = ReportI.DESTRUCTIVE + 1
        End If

End Function

Function ClearList()
Form1.MSL.ListItems.Clear
Form1.ProjectFiles.ListItems.Clear
End Function

Private Sub ScanCodeLine(ByVal Line_Number As Long)
On Error GoTo Err_Check
Dim i As Long
Dim a As Long
Dim b As Long
Dim UMS As String 'Upper case of Malicious String
Dim TCL As String 'Upper case of CodeLine

TCL = UCase$(CodeLine(Line_Number))
ReportI.TOTALLines = ReportI.TOTALLines + 1
For i = 0 To UBound(MS)
        UMS = UCase(MS(i).MStr) 'I couldn't make this than better than
        If TestContext(TCL, UMS) Then
            AddToList i, Line_Number
        End If
Next i

DoEvents

Exit Sub
Err_Check:
ReDim Preserve GetErrors(UBound(GetErrors) + 1)
GetErrors(UBound(GetErrors)).Def = Err.Description
GetErrors(UBound(GetErrors)).File = ProjectFile & " on ScanCodeLine"
Exit Sub
End Sub

Function Load_MSData()

'NOTE: I did this intentionally, so that when compiled it can't be edited by
'users as easily, there's 61 malicious/accessory strings listed below
'Danger: if modified or hacked some of these can be triggered by using
'Shell with the array variable data!!!
'.msh' is included in this list due to it's libraries containing
'dangerous code, so if other projects than this one use it, it could
'be a disguised attack on your PC if you have those libraries. :(
'This version (From the last version) I promised an Encryption method
'to make it so no other program can use these strings for malicious
'purposes without first decrypting them.

'Designed for Libary with Now 75 Strings! (62 Originally, 13 More)
Open App.Path & "\Reference1.MSH" For Binary Access Read Lock Read As #2
Get #2, , MS
Close #2

CryptReference

End Function

Function CryptReference()
Dim i As Long
Dim j As Long
'It's a two way method the same function decrypts and encrypts depending
'if the data is encrypted or normal.
'Like the trick with app.CompanyName?
'Makes it easy to track down other programs using it.
'This doubles as a security feature, when compiled this would be impossible
'to access the library 100% without my company name in the exe.
'Feel free to use this encryption idea I don't care,
'it took me only 5 minutes to make (10 minutes to Debug however).
'Just add a few other things and it would make a good encryption method.

Dim DS As String
Dim MSC As String
Dim CSC As String

For i = 0 To UBound(MS)
DS = Empty
    For j = 1 To Len(MS(i).MStr)
    MSC = Mid(MS(i).MStr, j, 1)
    CSC = Mid(App.CompanyName, (i Mod Len(App.CompanyName) + 1), 1)
    DS = DS & Chr(Asc(MSC) Xor Asc(CSC))
    Next j
MS(i).MStr = DS
Next i
End Function


Function MakeLibrary()
'this function I only use, the code for is on my machine I paste it here
'and use debug to execute it so a new resource file is created and then
'remove it leaving a easy place to paste it again otherwise my program
'integrity is unreliable with all those malicious/Accessory strings.
End Function


Function GetFileType(File As String) As String

Dim UcaseFileExt As String
UcaseFileExt = UCase(Right(File, 4))

If UcaseFileExt = ".VBP" Then
GetFileType = "Visual Basic 6.0 Project"
ElseIf UcaseFileExt = ".FRM" Then
GetFileType = "Visual Basic 6.0 Form"
ElseIf UcaseFileExt = ".BAS" Then
GetFileType = "Visual Basic 6.0 Module"
ElseIf UcaseFileExt = ".CLS" Then
GetFileType = "Visual Basic 6.0 Class Module"
ElseIf UcaseFileExt = ".RES" Then
GetFileType = "Visual Basic 6.0 Resource File"
ElseIf UcaseFileExt = ".FRX" Then
GetFileType = "Visual Basic 6.0 Form Binary File"
ElseIf UcaseFileExt = ".TXT" Then
GetFileType = "Text Document"     'In case it's hidden in a text file
ElseIf UcaseFileExt = ".VBS" Then 'and is renamed into a script.
GetFileType = "VB Script" 'Real virus are made using this language variant
End If                    'of Visual Basic e.g 'Love Bug' so this also
                          'covers basic-level viruses too!
End Function

Function TerminateApp()
'Thanks to Coding Genius article, I Did use a proper way to close this app
'(a more efficient way). (If you don't, I think unallocated memory wastes
'your RAM therefore slowing your machine down, but only a huge App could
'do that like a 3-D game, but it's highly reccommended for all apps)

'Deallocates and erases the dynamic array memory
Erase MS
Erase CodeLine
Erase FilesToScan
Erase SubFiles
Erase GetErrors
'Unloads the forms nuff said
Unload Form1
Unload Proposial
Unload ReportWindow
Unload Splash
'Finishs the app
'The mistake is that the 'END' Statement does NOT do the above for you
'like i though, those above are manually required to unload properly.
End
End Function

Private Function TestContext(ByVal TestLine As String, ByVal TestWord As String) As Boolean
'Copyright 2002 Roger Gilchrist
'This code was provided by this author on PSC, so you must obtain his
'permissions and copyrights of this code if you want to use it.
'Roger I modified it slightly to suit my program and for
'speed optimization, e.g ByVal makes
'it recieve strings by values not by reference making it faster and
'change select case statement to IF...THEN...ELSEIF...ENDIF Statement also
'for fractional speed.

'Roger's Notes (I thought I'll leave them in to help other users and
'understand what he's fixed/done):
'I liked 'Malicious String Hunter' a lot but found that it made some unnecessary
'warnings when
'I ran it against a very large project I have.
'EXAMPLES
'MZ was found in 'Function LenTrimZero',
'output in  'Function OutPutToList',
'.sys in 'SysForm.sysButton(4).Enabled'
'.dll in Private Declare Sub CoTaskMemFree Lib "ole32.dll" (byVal.... etc
'and several over-dramatically named routines with names like KillRow(it clears
'and removes a row in a MSGrid control)
'SO
'I created this routine and sent it to you
'I also rewrote Function ScanCodeLine as Sub ScanCodeLine (It doesn't return
'anything) and
'simplified it to use this Function.

  Dim TestChars As String 'String of characters which legitimately could delimit
'the TestWord
  Dim Pos As Long         'Position of TestWord in TestLine
  Dim TPos As Long        'File extention special case test
  Dim CommentTest As Long
  
  TestChars = " .*;':!?<>+-_=" & Chr(34)

    'The first two tests were originally in ScanCodeLine
    'I just moved them here for better encapsulation of tests
    '-----------------------------------------------------------
    'Test 1 TestWord is in TestLine
    Pos = InStr(1, TestLine, TestWord)
        If Pos = 0 Then
        Exit Function ' its not there so don't list it
    End If
    '-----------------------------------------------------------
    'Test 2 is TestWord in a comment so not dangerous
    CommentTest = InStr(1, TestLine, "'")
    If CommentTest < Pos And CommentTest > 0 Then
        Exit Function ' its in a comment so don't list it
    End If
''Un-comment the next two lines and you have the original tests
'    TestContext = True
'    Exit Function
'-----------------------------------------------------------
''Below this line are my additions to the program
    'Test 3 Just in case the whole line is a dangerous word
    '(This would fail Test 4 as it has no Before/After characters)
    If TestLine = TestWord Then
        TestContext = True
        Exit Function ' the whole line is a suspicious word so include it
    End If
'-----------------------------------------------------------
    'sub-test: File extention is a special case
    TPos = Pos
    If Left(TestWord, 1) = "." Then 'If TestWord is a file extention only check
    'that it
        TPos = 1                    'not embedded at the end of the word
    End If
 
'-----------------------------------------------------------
'Test 4 check surrounding characters (where they exist)
    If TPos = 1 Then 'Only test After TestWord
        TestContext = InStr(TestChars, Mid(TestLine, Pos + Len(TestWord), 1)) > 0
    ElseIf TPos = (Len(TestLine) - Len(TestWord)) Then 'Only test Before TestWord
        TestContext = InStr(TestChars, Mid(TestLine, Pos - 1, 1)) > 0
    Else 'Test Before and After TestWord
        TestContext = InStr(TestChars, Mid(TestLine, Pos - 1, 1)) > 0
        If TestContext Then 'second half of test only if first half is True
        TestContext = TestContext And InStr(TestChars, Mid(TestLine, Pos + Len(TestWord), 1)) > 0
        End If
    End If
End Function
