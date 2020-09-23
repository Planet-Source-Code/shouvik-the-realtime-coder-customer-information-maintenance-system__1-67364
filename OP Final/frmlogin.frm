VERSION 5.00
Begin VB.Form frmlogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System LogIn"
   ClientHeight    =   1065
   ClientLeft      =   3375
   ClientTop       =   3240
   ClientWidth     =   4830
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   4830
   Begin VB.TextBox txtsourcepassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   7635
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   975
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.TextBox txtsourceusername 
      Height          =   300
      Left            =   7605
      TabIndex        =   7
      Top             =   585
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2880
      TabIndex        =   3
      Top             =   735
      Width           =   945
   End
   Begin VB.CommandButton cmdchange 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3870
      TabIndex        =   4
      Top             =   735
      Width           =   945
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1890
      TabIndex        =   2
      Top             =   735
      Width           =   945
   End
   Begin VB.TextBox txtusername 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   1410
      TabIndex        =   0
      Top             =   45
      Width           =   3420
   End
   Begin VB.TextBox txtpassword 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1410
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   390
      Width           =   3420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   450
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   105
      Width           =   1005
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim value

Private Sub cmdcancel_Click()
confirm = MsgBox("Sure to exit ?", vbApplicationModal + vbYesNo + vbQuestion, "Confirm Exit")
If confirm = vbNo Then
    txtusername.SetFocus
    Exit Sub
ElseIf confirm = vbYes Then
    End
End If
End Sub

Private Sub cmdchange_Click()
Unload Me
frmverify.Show
End Sub

Private Sub cmdOK_Click()
If Trim(txtusername.Text) = Trim(txtsourceusername.Text) And Trim(txtpassword.Text) = Trim(txtsourcepassword.Text) Then
    Unload Me
    frmmain.Show
ElseIf Trim(txtusername.Text) <> Trim(txtsourceusername.Text) And Trim(txtpassword.Text) <> Trim(txtsourcepassword.Text) Then
    MsgBox "The system username and password are incorrect.", vbApplicationModal + vbExclamation, "System Login Error"
    txtusername.Text = ""
    txtpassword.Text = ""
    txtusername.SetFocus
ElseIf Trim(txtusername.Text) = Trim(txtsourceusername.Text) And Trim(txtpassword.Text) <> Trim(txtsourcepassword.Text) Then
    MsgBox "The system password you typed is incorrect.", vbApplicationModal + vbExclamation, "System Login Error"
    txtpassword.Text = ""
    txtpassword.SetFocus
ElseIf Trim(txtusername.Text) <> Trim(txtsourceusername.Text) And Trim(txtpassword.Text) = Trim(txtsourcepassword.Text) Then
    MsgBox "The system username you typed is incorrect.", vbApplicationModal + vbExclamation, "System Login Error"
    txtusername.Text = ""
    txtusername.SetFocus
End If
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

value = GetSetting("System", "Login", "Username", "{not-saved}")
If value <> "{not-saved}" Then
    txtsourceusername.Text = value
End If

value = GetSetting("System", "Login", "Password", "{not-saved}")
If value <> "{not-saved}" Then
    txtsourcepassword.Text = value
End If
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub txtusername_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdOK_Click
End Sub
