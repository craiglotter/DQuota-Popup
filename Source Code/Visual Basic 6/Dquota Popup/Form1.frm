VERSION 5.00
Begin VB.Form MainDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Network Space Advisor"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7080
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.Label errorlabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Left            =   480
      TabIndex        =   12
      Top             =   1200
      Width           =   6015
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   1770
      Width           =   3375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   1770
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   1770
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Currently Used:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Quota:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1770
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   600
      X2              =   3900
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   3900
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Network space available for "
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":509A
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   3240
      Left            =   -600
      Picture         =   "Form1.frx":5169
      Top             =   0
      Width           =   7680
   End
End
Attribute VB_Name = "MainDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
   End Type

   Private Type PROCESS_INFORMATION
      hProcess As Long
      hThread As Long
      dwProcessID As Long
      dwThreadID As Long
   End Type

   Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
      hHandle As Long, ByVal dwMilliseconds As Long) As Long

   Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
      lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long

   Private Declare Function CloseHandle Lib "kernel32" _
      (ByVal hObject As Long) As Long

   Private Declare Function GetExitCodeProcess Lib "kernel32" _
      (ByVal hProcess As Long, lpExitCode As Long) As Long

   Private Const NORMAL_PRIORITY_CLASS = &H20&
   Private Const INFINITE = -1&

   Public Function ExecCmd(cmdline$)
      Dim proc As PROCESS_INFORMATION
      Dim start As STARTUPINFO

      ' Initialize the STARTUPINFO structure:
      start.cb = Len(start)

      ' Start the shelled application:
      ret& = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, _
         NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)

      ' Wait for the shelled application to finish:
         ret& = WaitForSingleObject(proc.hProcess, INFINITE)
         Call GetExitCodeProcess(proc.hProcess, ret&)
         Call CloseHandle(proc.hThread)
         Call CloseHandle(proc.hProcess)
         ExecCmd = ret&
   End Function

   





Private Sub Command1_Click()
Me.Hide
Unload Me
End Sub

Sub Read_Files()
    Dim fso As New FileSystemObject, fil1 As File, ts As TextStream
    If fso.FileExists("c:\dquotaresult.txt") Then
            Set fil1 = fso.GetFile("c:\dquotaresult.txt")
        ' Read the contents of the file.
        Set ts = fil1.OpenAsTextStream(ForReading)
        Dim s As String
        If Not ts.AtEndOfStream Then s = ts.ReadLine
            If Not ts.AtEndOfStream Then s = Trim(ts.ReadLine)
                If StringStartsWith(s, "DQUOTA:", vbTextCompare) = False Then
                    If StringStartsWith(s, "Directory quota", vbTextCompare) = True Then
                        Dim username As String
                        username = s
                        Dim loopcounter As Integer
                        loopcounter = 0
                        While (StringStartsWith(s, "disk", vbTextCompare) = False) And (loopcounter <= 10)
                            If Not ts.AtEndOfStream Then s = Trim(ts.ReadLine)
                            username = username & s
                            loopcounter = loopcounter + 1
                        Wend
                        username = Replace(username, "Directory quota for ", "", , , vbTextCompare)
                        Dim checkvalue As Integer
                        checkvalue = InStr(username, "(")
                        If checkvalue > 0 Then
                            username = Trim(Left(username, checkvalue - 1))
                        End If
                        Label2.Caption = "Network space available for " & username
                        Dim diskquota, diskinuse, remaining As Double
                        diskquota = Trim(Replace(s, "Disk Quota:", ""))
                        If StringStartsWith(diskquota, "No", vbTextCompare) = False Then
                            diskquota = CDbl(diskquota)
                            Label6.Caption = diskquota & " MB"
                        Else
                            Label6.Caption = diskquota
                        End If
                        If Not ts.AtEndOfStream Then s = Trim(ts.ReadLine)
                            diskinuse = CDbl(Trim(Replace(s, "Disk in use:", "")))
                            Dim diskinusepercentage, remainingpercentage As Double
                            If StringStartsWith(Label6.Caption, "No", vbTextCompare) = False Then
                                diskinusepercentage = Round(((diskinuse / diskquota) * 100), 1)
                                Label9.Caption = "(" & diskinusepercentage & "%)"
                            End If
                            Label7.Caption = diskinuse & " MB"
                            If Not ts.AtEndOfStream Then s = Trim(ts.ReadLine)
                                If Not ts.AtEndOfStream Then s = Trim(ts.ReadLine)
                                    If StringEndsWith(s, "*", vbTextCompare) = False Then
    remaining = CDbl(Trim(Replace(s, "Disk Available:", "")))
    remainingpercentage = Round((100 - diskinusepercentage), 1)
Label8.Caption = remaining & " MB"
If diskinusepercentage <= 80 Then
Label8.ForeColor = &H8000&
Label10.ForeColor = &H8000&
End If
If diskinusepercentage > 80 And diskinusepercentage <= 90 Then
Label8.ForeColor = &H80FF&
Label10.ForeColor = &H80FF&
End If
If diskinusepercentage > 90 Then
Label8.ForeColor = &HFF&
Label10.ForeColor = &HFF&
End If

Label10.Caption = "(" & remainingpercentage & "%)"
Else
s = Trim(Replace(s, "Disk Available:", "", , , vbTextCompare))
s = Trim(Replace(s, ",", "", , , vbTextCompare))
s = Trim(Replace(s, "*", "", , , vbTextCompare))
Label8.Caption = s & " MB"
Label11.Caption = "available on server."
End If
Else
Line1.Visible = False
Line2.Visible = False
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
errorlabel.Caption = "This program was unable to complete your request. You either do not have a Home Directory Attribute set for Novell or you have not been logged into the Novell network correctly. Please hit the 'Close' button to exit this screen."
End If
Else
Line1.Visible = False
Line2.Visible = False
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
errorlabel.Caption = "This program was unable to complete your request. You either do not have a Home Directory Attribute set for Novell or you have not been logged into the Novell network correctly. Please hit the 'Close' button to exit this screen."
End If
    ts.Close
    fil1.Delete (True)
  
    Else
Line1.Visible = False
Line2.Visible = False
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
errorlabel.Caption = "This program was unable to complete your request. You either do not have a Home Directory Attribute set for Novell or you have not been logged into the Novell network correctly. Please hit the 'Close' button to exit this screen."
    End If
End Sub

Public Function StringStartsWith(ByVal strValue As String, _
  CheckFor As String, Optional CompareType As VbCompareMethod _
   = vbBinaryCompare) As Boolean
   
'Determines if a string starts with the same characters as
'CheckFor string

'True if starts with CheckFor, false otherwise
'Case sensitive by default.  If you want non-case sensitive, set
'last parameter to vbTextCompare
    
    'Examples:
    'MsgBox StringStartsWith("Test", "TE") 'false
    'MsgBox StringStartsWith("Test", "TE", vbTextCompare) 'True
    
  Dim sCompare As String
  Dim lLen As Long
   
  lLen = Len(CheckFor)
  If lLen > Len(strValue) Then Exit Function
  sCompare = Left(strValue, lLen)
  StringStartsWith = StrComp(sCompare, CheckFor, CompareType) = 0

End Function

Public Function StringEndsWith(ByVal strValue As String, _
   CheckFor As String, Optional CompareType As VbCompareMethod _
   = vbBinaryCompare) As Boolean
 'Determines if a string ends with the same characters as
 'CheckFor string
 
 'True if end with CheckFor, false otherwise

 'Case sensitive by default.  If you want non-case sensitive, set
 'last parameter to vbTextCompare
 
  'Examples
  'MsgBox StringEndsWith("Test", "ST") 'False
  'MsgBox StringEndsWith("Test", "ST", vbTextCompare) 'True

  Dim sCompare As String
  Dim lLen As Long
   
  lLen = Len(CheckFor)
  If lLen > Len(strValue) Then Exit Function
  sCompare = Right(strValue, lLen)
  StringEndsWith = StrComp(sCompare, CheckFor, CompareType) = 0

End Function


Private Sub Form_Load()
'Dim PID As Double
'PID = Shell("rundquota.bat", vbMinimizedNoFocus)
Dim retval As Long
retval = -1
retval = ExecCmd(App.Path & "\rundquota.bat")

If retval = 0 Then
    Read_Files
Else
    retval = ExecCmd(App.Path & "\rundquota.bat")

    If retval = 0 Then
        Read_Files
    End If
End If
End Sub


