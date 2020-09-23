VERSION 5.00
Begin VB.Form frminvpmtinfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirm Paid Amount"
   ClientHeight    =   2670
   ClientLeft      =   3120
   ClientTop       =   2760
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frminvpmtinfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5670
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
      Left            =   3720
      TabIndex        =   14
      Top             =   2340
      Width           =   945
   End
   Begin VB.CommandButton cmdback 
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
      Left            =   4710
      TabIndex        =   13
      Top             =   2340
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Caption         =   "Purchasing Details"
      Height          =   1710
      Left            =   0
      TabIndex        =   5
      Top             =   570
      Width           =   5685
      Begin VB.TextBox txtpaid 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1950
         TabIndex        =   0
         Top             =   1350
         Width           =   1770
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Last Paid :"
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   1095
         Width           =   1575
      End
      Begin VB.Label lblLastpaid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1980
         TabIndex        =   16
         Top             =   1095
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
         Height          =   195
         Left            =   1665
         TabIndex        =   15
         Top             =   1410
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paid :"
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   1410
         Width           =   450
      End
      Begin VB.Label lblprice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1980
         TabIndex        =   11
         Top             =   570
         Width           =   45
      End
      Begin VB.Label lbldue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1980
         TabIndex        =   10
         Top             =   840
         Width           =   45
      End
      Begin VB.Label lbldomain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1980
         TabIndex        =   9
         Top             =   300
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Domain Value :"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   570
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due :"
         Height          =   195
         Left            =   105
         TabIndex        =   7
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Domain Purchased :"
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   300
         Width           =   1665
      End
   End
   Begin VB.Label lblcname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1965
      TabIndex        =   4
      Top             =   270
      Width           =   45
   End
   Begin VB.Label lblbno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1965
      TabIndex        =   3
      Top             =   15
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name :"
      Height          =   195
      Left            =   165
      TabIndex        =   2
      Top             =   270
      Width           =   1410
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No :"
      Height          =   195
      Left            =   165
      TabIndex        =   1
      Top             =   15
      Width           =   585
   End
End
Attribute VB_Name = "frminvpmtinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database, rs As Recordset, tprice, lpaid, due As Variant

Private Sub cmdback_Click()
Unload Me
Unload frmreport
selbno = ""
currentform = ""
frmmain.Show
tprice = 0
End Sub

Private Sub cmdOK_Click()
If Trim(txtpaid.Text) = "" Then
    MsgBox "Paid amount should not be left blank.", vbApplicationModal + vbInformation, "Paid Amount"
    txtpaid.SetFocus
ElseIf Val(Trim(txtpaid.Text)) <= 0 Then
    MsgBox "Paid amount should be greater than 0.", vbApplicationModal + vbInformation, "Wrong Input"
    txtpaid.SelStart = 0
    txtpaid.SelLength = Len(Trim(txtpaid.Text))
    txtpaid.SetFocus
ElseIf Val(Trim(txtpaid.Text)) > due Then
    MsgBox "Paid should not more than total domain value.", vbApplicationModal + vbInformation, "Overflow"
    txtpaid.SelStart = 0
    txtpaid.SelLength = Len(Trim(txtpaid.Text))
    txtpaid.SetFocus
ElseIf Trim(txtpaid.Text) <> "" And Val(Trim(txtpaid.Text)) > 1 And Val(Trim(txtpaid.Text)) <= tprice And due >= 0 Then
    currentform = "invoice"
    frminvoice.Show
    frmoptions.Show
    frminvoice.lbldue.Caption = Val(tprice) - (Val(Trim(txtpaid.Text)) + Val(lpaid))
    frminvoice.lblLastpaid.Caption = Val(lpaid)
    Me.Hide
End If
End Sub

Private Sub Form_Activate()
txtpaid.SelStart = 0
txtpaid.SelLength = Len(Trim(txtpaid.Text))
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

Set db = OpenDatabase(App.Path & "\company.mdb")
Set rs = db.OpenRecordset("info", dbOpenTable)
If rs.RecordCount > 0 Then rs.MoveFirst

Call GetBillInfo
End Sub

Public Sub GetBillInfo()
On Error Resume Next

Set rs = db.OpenRecordset("select *from info where billno=" & "'" & selbno & "'")
If rs.RecordCount > 0 Then
    rs.MoveFirst
    lblbno.Caption = rs("billno")
    lblcname.Caption = rs("companyname")
    lbldomain.Caption = rs("domain")
    tprice = rs("price")
    lblprice.Caption = "Rs. " & tprice
    lbldue.Caption = "Rs. " & rs("due")
    lpaid = rs("lastpaid")
    If Val(lpaid) = 0 Then
        lblLastpaid.Caption = "Rs. " & 0
        txtpaid.Text = rs("paid")
    ElseIf Val(lpaid) <> 0 Then
        lblLastpaid.Caption = "Rs. " & lpaid
        txtpaid.Text = ""
    End If
End If
Set rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdback_Click
End Sub

Private Sub txtpaid_Change()
due = Val(tprice) - (Val(Trim(txtpaid.Text)) + Val(lpaid))
lbldue.Caption = "Rs. " & Val(tprice) - (Val(Trim(txtpaid.Text)) + Val(lpaid))
End Sub

Private Sub txtpaid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub txtpaid_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
