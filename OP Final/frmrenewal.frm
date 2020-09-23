VERSION 5.00
Begin VB.Form frmrenewal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Renew Customer Membership"
   ClientHeight    =   1575
   ClientLeft      =   3030
   ClientTop       =   3150
   ClientWidth     =   6375
   Icon            =   "frmrenewal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6375
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   0
      TabIndex        =   6
      Top             =   1530
      Width           =   6375
   End
   Begin VB.CommandButton cmdrenew 
      Caption         =   "Renew"
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
      Left            =   3765
      TabIndex        =   1
      Top             =   1170
      Width           =   945
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
      Left            =   4755
      TabIndex        =   2
      Top             =   1170
      Width           =   945
   End
   Begin VB.ComboBox cbobno 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2827
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   750
      Width           =   2925
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   660
      Width           =   6375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Customer Bill No :"
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
      Left            =   622
      TabIndex        =   5
      Top             =   810
      Width           =   2025
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmrenewal.frx":164A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   15
      TabIndex        =   3
      Top             =   0
      Width           =   6420
   End
End
Attribute VB_Name = "frmrenewal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database, rs As Recordset
Dim todate As Variant

Private Sub cmdback_Click()
Unload Me
frmmain.Show
End Sub

Private Sub cmdrenew_Click()
If cbobno.ListIndex = 0 Then
    MsgBox "Select a billno from the list above to renew information.", vbApplicationModal + vbExclamation, "No billno selected"
    Exit Sub
Else
    selbno = cbobno.List(cbobno.ListIndex)
    Me.Hide
    frmrenewadd.Show
End If
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

Set db = OpenDatabase(App.Path & "\company.mdb")
Set rs = db.OpenRecordset("info", dbOpenTable)
If rs.RecordCount > 0 Then rs.MoveFirst
cbobno.Clear
cbobno.AddItem "-=Select BillNo=-"
Set rs = db.OpenRecordset("select *from info where endcamp<=" & "'" & Format(Date, "dd/mm/yyyy") & "'")
If rs.RecordCount > 0 Then
    rs.MoveFirst
    Do Until rs.EOF
        cbobno.AddItem rs("billno")
        rs.MoveNext
    Loop
End If

If cbobno.ListCount = 1 Then
    MsgBox "No customer membership has been expiring today.", vbApplicationModal + vbInformation, "Renew"
    Unload Me
    frmmain.Show
ElseIf cbobno.ListCount > 1 Then
    cbobno.ListIndex = 0
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdback_Click
End Sub
