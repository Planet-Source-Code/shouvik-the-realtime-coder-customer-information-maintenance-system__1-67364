VERSION 5.00
Begin VB.Form frmcheque 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheque Details"
   ClientHeight    =   1725
   ClientLeft      =   3870
   ClientTop       =   3315
   ClientWidth     =   4215
   ControlBox      =   0   'False
   Icon            =   "frmcheque.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4215
   Begin VB.Frame Frame1 
      Caption         =   "Payment Mode - Cheque"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   7
      TabIndex        =   5
      Top             =   15
      Width           =   4200
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
         Left            =   2160
         TabIndex        =   3
         Top             =   1305
         Width           =   945
      End
      Begin VB.TextBox txtbranch 
         Height          =   300
         Left            =   1665
         TabIndex        =   2
         Top             =   975
         Width           =   2460
      End
      Begin VB.TextBox txtbank 
         Height          =   300
         Left            =   1665
         TabIndex        =   1
         Top             =   645
         Width           =   2460
      End
      Begin VB.TextBox txtchequeno 
         Height          =   300
         Left            =   1665
         TabIndex        =   0
         Top             =   315
         Width           =   2460
      End
      Begin VB.CommandButton cmdback 
         Caption         =   "Back"
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
         Left            =   3150
         TabIndex        =   4
         Top             =   1305
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch :"
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
         Left            =   105
         TabIndex        =   8
         Top             =   1065
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drawn On Bank :"
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
         Left            =   105
         TabIndex        =   7
         Top             =   735
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque Number :"
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
         Left            =   105
         TabIndex        =   6
         Top             =   390
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmcheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdback_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
If Trim(txtchequeno.Text) <> "" And Trim(txtbank.Text) <> "" And Trim(txtbranch.Text) <> "" Then
    recchno = Trim(txtchequeno.Text)
    recbank = Trim(txtbank.Text)
    recbranch = Trim(txtbranch.Text)
    Unload Me
    frmrenewadd.txtprice.SetFocus
Else
    MsgBox "Incomplete cheque information.", vbApplicationModal + vbExclamation, "Cheque Details"
    txtchequeno.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next

txtchequeno.Text = recchno
txtbank.Text = recbank
txtbranch.Text = recbranch
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
End Sub

Private Sub txtbank_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtbranch.SetFocus
End Sub

Private Sub txtbank_LostFocus()
txtbank.Text = StrConv(Trim(txtbank.Text), vbProperCase)
End Sub

Private Sub txtbranch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdok.SetFocus
End Sub

Private Sub txtbranch_LostFocus()
txtbranch.Text = StrConv(Trim(txtbranch.Text), vbProperCase)
End Sub

Private Sub txtchequeno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtbank.SetFocus
End Sub

Private Sub txtchequeno_LostFocus()
txtchequeno.Text = StrConv(Trim(txtchequeno.Text), vbProperCase)
End Sub
