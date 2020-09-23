VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmsearch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search & Query"
   ClientHeight    =   3705
   ClientLeft      =   2805
   ClientTop       =   3075
   ClientWidth     =   8085
   Icon            =   "frmsearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8085
   Begin VB.TextBox txtresult 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3345
      Width           =   5250
   End
   Begin VB.CheckBox chkdomain 
      Caption         =   "By Domain"
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
      Left            =   2100
      TabIndex        =   4
      Top             =   -15
      Width           =   1335
   End
   Begin VB.CheckBox chkendcamp 
      Caption         =   "By End-Campaigning"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3660
      TabIndex        =   3
      Top             =   -30
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid GridInfo 
      Bindings        =   "frmsearch.frx":0442
      Height          =   2895
      Left            =   0
      OleObjectBlob   =   "frmsearch.frx":045B
      TabIndex        =   2
      Top             =   405
      Width           =   8085
   End
   Begin VB.Data datasearch 
      Caption         =   "datasearch"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3795
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4500
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Back"
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
      Left            =   6600
      TabIndex        =   1
      ToolTipText     =   "Closes search and backs to main screen"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
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
      Left            =   6600
      TabIndex        =   0
      ToolTipText     =   "Searches by category"
      Top             =   30
      Width           =   1455
   End
   Begin VB.Label lblsearchresult 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Result :"
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
      Left            =   15
      TabIndex        =   6
      Top             =   3345
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Searching Options :"
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
      Left            =   15
      TabIndex        =   5
      Top             =   30
      Width           =   1620
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkdomain_Click()
If chkdomain.value = 1 Then
    chkendcamp.value = 0
End If
End Sub

Private Sub chkendcamp_Click()
If chkendcamp.value = 1 Then
    chkdomain.value = 0
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
frmmain.Show
End Sub

Private Sub cmdsearch_Click()
If chkdomain.value = 0 And chkendcamp.value = 0 Then
    MsgBox "Select a searching option to begin search.", vbApplicationModal + vbInformation, "Search & Query"
    Exit Sub
ElseIf chkdomain.value = 1 Then
    frmdomains.Show vbModal
ElseIf chkendcamp.value = 1 Then
    frmendcampdates.Show vbModal
End If
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
datasearch.DatabaseName = App.Path & "\company.mdb"
txtresult.Text = "No Action.Total 0 Record Found"
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdclose_Click
End Sub
