VERSION 5.00
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Info Maintenance"
   ClientHeight    =   4290
   ClientLeft      =   3375
   ClientTop       =   2415
   ClientWidth     =   5340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5340
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
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
      Left            =   4035
      TabIndex        =   8
      Top             =   3870
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
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
      Left            =   2790
      TabIndex        =   7
      Top             =   3870
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "Available Services"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   5340
      Begin VB.OptionButton optAbout 
         Caption         =   "About Customer Information Maintenance System "
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
         Left            =   420
         TabIndex        =   13
         Top             =   2085
         Width           =   4710
      End
      Begin VB.OptionButton optnew 
         Caption         =   "Add New Customer Information"
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
         Left            =   405
         TabIndex        =   1
         Top             =   300
         Width           =   3330
      End
      Begin VB.OptionButton optmod 
         Caption         =   "Modify Existing Customer Information"
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
         Left            =   405
         TabIndex        =   2
         Top             =   585
         Width           =   3675
      End
      Begin VB.OptionButton optreport 
         Caption         =   "View Report of All Customers"
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
         Left            =   405
         TabIndex        =   4
         Top             =   1185
         Width           =   2805
      End
      Begin VB.OptionButton optsearch 
         Caption         =   "Search Customer Information"
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
         Left            =   405
         TabIndex        =   3
         Top             =   885
         Width           =   3495
      End
      Begin VB.OptionButton optrenew 
         Caption         =   "Renew Customer Membership"
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
         Left            =   405
         TabIndex        =   6
         Top             =   1785
         Width           =   2925
      End
      Begin VB.OptionButton optpayment 
         Caption         =   "Change Customer Payment Details"
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
         Left            =   405
         TabIndex        =   5
         Top             =   1500
         Width           =   3300
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   0
      TabIndex        =   11
      Top             =   780
      Width           =   5265
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a service from the available options below and click OK to proceed. To exit the program click on Cancel."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   45
      TabIndex        =   10
      Top             =   900
      Width           =   5265
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER INFORMATION MAINTENANCE SYSTEM"
      Height          =   210
      Left            =   375
      TabIndex        =   9
      Top             =   480
      Width           =   4515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "WELCOME TO"
      Height          =   570
      Left            =   2160
      TabIndex        =   0
      Top             =   15
      Width           =   945
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database, rs As Recordset

Private Sub cmdcancel_Click()
Dim confirm As Integer

confirm = MsgBox("Sure to exit ?", vbApplicationModal + vbYesNo + vbQuestion, "Confirm Exit")
If confirm = vbNo Then
    Exit Sub
ElseIf confirm = vbYes Then
    End
End If
End Sub

Private Sub cmdOK_Click()
On Error GoTo xx

If optnew.value = True Then
    Me.Hide
    frmentry.Show
ElseIf optmod.value = True Then
    If rs.RecordCount > 0 Then
        Me.Hide
        frmedit.Show
    ElseIf rs.RecordCount = 0 Then
        MsgBox "No customer information available for modification.", vbApplicationModal + vbExclamation, "Modification"
        Exit Sub
    End If
ElseIf optreport.value = True Then
    If rs.RecordCount > 0 Then
        Me.Hide
        frmreport.Show
    ElseIf rs.RecordCount = 0 Then
        MsgBox "No customer information available to see as report.", vbApplicationModal + vbExclamation, "Report"
        Exit Sub
    End If
ElseIf optsearch.value = True Then
    If rs.RecordCount > 0 Then
        Me.Hide
        frmsearch.Show
    ElseIf rs.RecordCount = 0 Then
        MsgBox "No customer information available to search.", vbApplicationModal + vbExclamation, "Search"
        Exit Sub
    End If
ElseIf optpayment.value = True Then
    If rs.RecordCount > 0 Then
        Me.Hide
        frmpaymentdetails.Show
    ElseIf rs.RecordCount = 0 Then
        MsgBox "No customer information available to change payment details.", vbApplicationModal + vbExclamation, "Payment Details"
        Exit Sub
    End If
ElseIf optrenew.value = True Then
    If rs.RecordCount > 0 Then
        Me.Hide
        frmrenewal.Show
    ElseIf rs.RecordCount = 0 Then
        MsgBox "No customer information available for renewal.", vbApplicationModal + vbExclamation, "Renewal"
        Exit Sub
    End If
    ElseIf optAbout.value = True Then
        Me.Hide
        frmAbout.Show
Else
    MsgBox "You have to select an option in order to proceed.", vbApplicationModal + vbCritical, "CustInfo"
    Exit Sub
End If

xx:
If Err.Number = 364 Then
    Unload frmrenewal
    Exit Sub
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdcancel_Click
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

Set db = OpenDatabase(App.Path & "\company.mdb")
Set rs = db.OpenRecordset("info", dbOpenTable)
If rs.RecordCount > 0 Then rs.MoveFirst
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim confirm As Integer

confirm = MsgBox("Sure to exit ?", vbApplicationModal + vbYesNo + vbQuestion, "Confirm Exit")
If confirm = vbNo Then
    Cancel = vbNo
    Exit Sub
ElseIf confirm = vbYes Then
    End
End If
End Sub

Private Sub optAbout_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub optmod_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub optnew_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub optpayment_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub optrenew_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub optreport_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub optsearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdOK_Click
End Sub
