VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3645
   ClientLeft      =   3075
   ClientTop       =   2940
   ClientWidth     =   6435
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar p1 
      Height          =   225
      Left            =   2070
      TabIndex        =   5
      Top             =   4770
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1665
      Top             =   3840
   End
   Begin VB.Frame Frame1 
      Height          =   3750
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   6450
      Begin VB.Label lblabout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   45
         TabIndex        =   4
         Top             =   3315
         Width           =   840
      End
      Begin VB.Label lblversion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5775
         TabIndex        =   3
         Top             =   120
         Width           =   630
      End
      Begin VB.Image imgLogo 
         Height          =   480
         Left            =   420
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblProductName 
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER INFORMATION MAINTENANCE SYSTEM"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   24
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   1755
         TabIndex        =   2
         Top             =   465
         Width           =   4050
      End
      Begin VB.Label lblLicense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   45
         TabIndex        =   1
         Top             =   120
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
(ByRef lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0

Dim i As Integer

Private Sub Form_Initialize()
Call CheckVersion
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Timer1.Enabled = False
    If frmlogin.txtsourceusername.Text = "" And frmlogin.txtsourcepassword.Text = "" Then
        Unload Me
        MsgBox "This program is running for the first time.Please set an system" & vbCrLf & "authentication in order to prevent unauthorized access.", vbApplicationModal + vbInformation, "Welcome To System LogIn"
        frmchange.Show
    ElseIf frmlogin.txtsourceusername.Text <> "" And frmlogin.txtsourcepassword.Text <> "" Then
        Unload Me
        frmlogin.Show
    End If
End If
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
p1.Max = 100
p1.Min = 0
lblLicense.Caption = "Licensed To : " & GetMachineUserName
lblversion.Caption = "Version Info : " & App.Major & "." & App.Minor & "." & App.Revision
lblabout.Caption = "Developed By : Shouvik Choudhury [Contact : 98831-74688]" & Chr(10) & "Supported And Advised By : Mr. Subrata Santra [Contact : 98303-89744]"
Timer1_Timer
End Sub

Private Sub Timer1_Timer()
i = i + 5
p1.value = i
If i = 100 Then
    If frmlogin.txtsourceusername.Text = "" And frmlogin.txtsourcepassword.Text = "" Then
        Unload Me
        MsgBox "This program is running for the first time.Please set an system" & vbCrLf & "authentication in order to prevent unauthorized access.", vbApplicationModal + vbInformation, "Welcome To System LogIn"
        frmchange.Show
    ElseIf frmlogin.txtsourceusername.Text <> "" And frmlogin.txtsourcepassword.Text <> "" Then
        Unload Me
        frmlogin.Show
    End If
End If
End Sub

Public Sub CheckVersion()
Dim info As OSVERSIONINFO
Dim txt As String
Dim vno As Double

info.dwOSVersionInfoSize = Len(info)
GetVersionEx info
If info.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    txt = "Windows95 "
Else
    txt = "WindowsNT "
End If
vno = Format$(info.dwMajorVersion) & "." & Format$(info.dwMinorVersion)
If txt = "WindowsNT " Then
    Exit Sub
ElseIf txt = "Windows95 " Then
    MsgBox "This program is not compatible with this version of windows.", vbApplicationModal + vbCritical, "Fatal Error"
    End
End If
End Sub
