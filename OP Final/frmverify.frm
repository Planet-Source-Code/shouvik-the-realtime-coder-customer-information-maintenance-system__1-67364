VERSION 5.00
Begin VB.Form frmverify 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Verification"
   ClientHeight    =   735
   ClientLeft      =   2520
   ClientTop       =   2370
   ClientWidth     =   6300
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
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
      Left            =   4350
      TabIndex        =   3
      Top             =   390
      Width           =   945
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
      Left            =   5340
      TabIndex        =   2
      Top             =   390
      Width           =   945
   End
   Begin VB.TextBox txtverify 
      Alignment       =   2  'Center
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
      Left            =   2895
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   30
      Width           =   3420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Current System Password :"
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
      TabIndex        =   1
      Top             =   75
      Width           =   2775
   End
End
Attribute VB_Name = "frmverify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
Unload Me
frmlogin.Show
End Sub

Private Sub cmdok_Click()
If Trim(txtverify.Text) = "" Then
    Beep
    Exit Sub
ElseIf Trim(txtverify.Text) <> "" Then
    If Trim(txtverify.Text) = Trim(frmlogin.txtsourcepassword.Text) Then
        Unload Me
        frmchange.Show
    Else
        MsgBox "The system password you typed is incorrect.", vbApplicationModal + vbExclamation, "Password Verification Error"
        txtverify.SelStart = 0
        txtverify.SelLength = Len(Trim(txtverify.Text))
        txtverify.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
End Sub

Private Sub txtverify_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdok_Click
End Sub
