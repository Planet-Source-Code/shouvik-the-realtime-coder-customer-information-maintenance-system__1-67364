VERSION 5.00
Begin VB.Form frmchange 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System LogIn Settings"
   ClientHeight    =   1365
   ClientLeft      =   2295
   ClientTop       =   4080
   ClientWidth     =   5625
   ControlBox      =   0   'False
   Icon            =   "frmchange.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5625
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
      Left            =   4650
      TabIndex        =   4
      Top             =   1035
      Width           =   945
   End
   Begin VB.TextBox txtconfirmpassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2205
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   705
      Width           =   3420
   End
   Begin VB.TextBox txtnewpassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2205
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   375
      Width           =   3420
   End
   Begin VB.TextBox txtnewusername 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2205
      TabIndex        =   0
      Top             =   45
      Width           =   3420
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
      Left            =   3675
      TabIndex        =   3
      Top             =   1035
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm New Password :"
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
      Left            =   60
      TabIndex        =   7
      Top             =   765
      Width           =   1995
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New User Name :"
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
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password :"
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
      Left            =   60
      TabIndex        =   5
      Top             =   435
      Width           =   1290
   End
End
Attribute VB_Name = "frmchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
If Trim(frmlogin.txtsourceusername.Text) = "" And Trim(frmlogin.txtsourcepassword.Text) = "" Then
    MsgBox "You must set the system authentication in" & vbCrLf & "order to proceed for the first time.", vbApplicationModal + vbInformation, "Access Denied"
    txtnewusername.SetFocus
Else
    Unload Me
    frmlogin.Show
End If
End Sub

Private Sub cmdok_Click()
If Trim(txtnewusername.Text) <> "" And Trim(txtnewpassword.Text) <> "" And Trim(txtconfirmpassword.Text) = Trim(txtnewpassword.Text) Then
    SaveSetting "System", "Login", "Username", Trim(txtnewusername.Text)
    SaveSetting "System", "Login", "Password", Trim(txtnewpassword.Text)
    SaveSetting "System", "Login", "ConfirmPassword", Trim(txtconfirmpassword.Text)
    frmlogin.txtsourceusername.Text = Trim(txtnewusername.Text)
    frmlogin.txtsourcepassword.Text = Trim(txtnewpassword.Text)
    Unload Me
    frmlogin.Show
Else
    MsgBox "Unable to save system login settings.", vbApplicationModal + vbExclamation, "System Login Settings"
    txtnewusername.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
End Sub

Private Sub txtconfirmpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdok_Click
End Sub

Private Sub txtnewpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdok_Click
End Sub

Private Sub txtnewusername_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdok_Click
End Sub
