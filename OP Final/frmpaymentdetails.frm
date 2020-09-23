VERSION 5.00
Begin VB.Form frmpaymentdetails 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Details"
   ClientHeight    =   3735
   ClientLeft      =   2730
   ClientTop       =   2400
   ClientWidth     =   6450
   Icon            =   "frmpaymentdetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6450
   Begin VB.Frame Frame1 
      Caption         =   "Customer Payment Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   0
      TabIndex        =   12
      Top             =   420
      Width           =   6465
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
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
         Left            =   3450
         TabIndex        =   8
         Top             =   2655
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
         Left            =   5400
         TabIndex        =   10
         Top             =   2655
         Width           =   945
      End
      Begin VB.CommandButton cmdreset 
         Caption         =   "Reset"
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
         Left            =   4425
         TabIndex        =   9
         Top             =   2655
         Width           =   945
      End
      Begin VB.TextBox txtchequeno 
         Height          =   300
         Left            =   2130
         TabIndex        =   2
         Top             =   675
         Width           =   4245
      End
      Begin VB.TextBox txtbank 
         Height          =   300
         Left            =   2130
         TabIndex        =   3
         Top             =   1005
         Width           =   4245
      End
      Begin VB.TextBox txtbranch 
         Height          =   300
         Left            =   2130
         TabIndex        =   4
         Top             =   1335
         Width           =   4245
      End
      Begin VB.ComboBox cbopayment 
         Height          =   315
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   4245
      End
      Begin VB.TextBox txtprice 
         Height          =   300
         Left            =   2130
         TabIndex        =   5
         Top             =   1680
         Width           =   4245
      End
      Begin VB.TextBox txtpaid 
         Height          =   300
         Left            =   2130
         TabIndex        =   6
         Top             =   2010
         Width           =   4245
      End
      Begin VB.TextBox txtdue 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2340
         Width           =   4245
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000A&
         BorderWidth     =   2
         Height          =   15
         Left            =   15
         Top             =   3045
         Width           =   6435
      End
      Begin VB.Label lblstatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label17"
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
         TabIndex        =   24
         Top             =   3075
         Width           =   660
      End
      Begin VB.Label lblcount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label18"
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
         Left            =   5760
         TabIndex        =   23
         Top             =   3075
         Width           =   660
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
         Left            =   120
         TabIndex        =   22
         Top             =   750
         Width           =   1425
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
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1365
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
         Left            =   120
         TabIndex        =   20
         Top             =   1410
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Domain Value :"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1740
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paid :"
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
         Left            =   120
         TabIndex        =   18
         Top             =   2085
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due :"
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
         Left            =   120
         TabIndex        =   17
         Top             =   2415
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Mode :"
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
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Width           =   1365
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
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
         Left            =   1860
         TabIndex        =   15
         Top             =   1740
         Width           =   255
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
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
         Left            =   1860
         TabIndex        =   14
         Top             =   2085
         Width           =   255
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
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
         Left            =   1860
         TabIndex        =   13
         Top             =   2415
         Width           =   255
      End
   End
   Begin VB.ComboBox cbobillno 
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
      Left            =   2775
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   15
      Width           =   3030
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Customer BillNo :"
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
      Left            =   645
      TabIndex        =   11
      Top             =   75
      Width           =   1980
   End
End
Attribute VB_Name = "frmpaymentdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database, rs As Recordset, rsd As Recordset, recd As Recordset, lpaid As Variant

Private Sub cbobillno_Click()
If cbobillno.ListIndex = 0 Then
    Exit Sub
Else
    Call GetPayment
    Call DoChange
    cmdsave.Enabled = True
    cmdreset.Enabled = True
    cmdback.Enabled = False
    cbobillno.Enabled = False
    Label13.Enabled = False
    lblstatus.Caption = "Service Status : Changing Payment Details"
End If
End Sub

Private Sub cbopayment_Click()
If cbopayment.ListIndex = 0 Then
    Exit Sub
ElseIf cbopayment.ListIndex = 1 Then
    txtchequeno.Text = ""
    txtbank.Text = ""
    txtbranch.Text = ""
    txtchequeno.Enabled = False
    txtbank.Enabled = False
    txtbranch.Enabled = False
    txtchequeno.BackColor = vbButtonFace
    txtbank.BackColor = vbButtonFace
    txtbranch.BackColor = vbButtonFace
    txtprice.SetFocus
ElseIf cbopayment.ListIndex = 2 Then
    txtchequeno.Enabled = True
    txtbank.Enabled = True
    txtbranch.Enabled = True
    txtchequeno.BackColor = vbWhite
    txtbank.BackColor = vbWhite
    txtbranch.BackColor = vbWhite
    txtchequeno.SetFocus
End If
End Sub

Private Sub cmdback_Click()
Unload Me
frmmain.Show
End Sub

Private Sub cmdreset_Click()
Call ClearFields
Call DisableAll
cmdsave.Enabled = False
cmdreset.Enabled = False
cmdback.Enabled = True
lblstatus.Caption = "Service Status : Payment Details Reseted"
cbobillno.Enabled = True
Label13.Enabled = True
cbobillno.ListIndex = 0
cbobillno.SetFocus
End Sub

Private Sub cmdsave_Click()
If cbopayment.ListIndex = 0 Then
    MsgBox "Select a proper payment mode.", vbApplicationModal + vbInformation, "No payment mode selected"
    cbopayment.SetFocus
ElseIf cbopayment.ListIndex = 1 Then
    If Trim(txtprice.Text) = "" Or Trim(txtpaid.Text) = "" Then
        MsgBox "Domain value and paid amount must be mentioned.", vbApplicationModal + vbInformation, "Incomplete data"
        txtprice.SetFocus
    ElseIf Val(Trim(txtpaid.Text)) > Val(Trim(txtprice.Text)) Then
        MsgBox "Paid amount should not be more than domain value.", vbApplicationModal + vbCritical, "Wrong paid amount"
        txtpaid.SelStart = 0
        txtpaid.SelLength = Len(Trim(txtpaid.Text))
        txtpaid.SetFocus
    ElseIf Val(Trim(txtpaid.Text)) <= Val(Trim(txtprice.Text)) Then
        Call DoSave
        Exit Sub
    End If
ElseIf cbopayment.ListIndex = 2 Then
    If Trim(txtchequeno.Text) = "" Or Trim(txtbank.Text) = "" Or Trim(txtbranch.Text) = "" Or _
         Trim(txtprice.Text) = "" Or Trim(txtpaid.Text) = "" Then
        MsgBox "Incomplete payment information.", vbApplicationModal + vbInformation, "Incomplete data"
        txtchequeno.SetFocus
    ElseIf Val(Trim(txtpaid.Text)) > Val(Trim(txtprice.Text)) Then
        MsgBox "Paid amount should not be more than domain value.", vbApplicationModal + vbCritical, "Wrong paid amount"
        txtpaid.SelStart = 0
        txtpaid.SelLength = Len(Trim(txtpaid.Text))
        txtpaid.SetFocus
    ElseIf Trim(txtchequeno.Text) <> "" And Trim(txtbank.Text) <> "" And Trim(txtbranch.Text) <> "" And _
         Trim(txtprice.Text) <> "" And Trim(txtpaid.Text) <> "" And Val(Trim(txtpaid.Text)) <= Val(Trim(txtprice.Text)) Then
         Call DoSave
        Exit Sub
    End If
End If
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

Set db = OpenDatabase(App.Path & "\company.mdb")
Set rs = db.OpenRecordset("info", dbOpenTable)
If rs.RecordCount > 0 Then rs.MoveFirst

Set rsd = db.OpenRecordset("info")
cbobillno.Clear
cbobillno.AddItem "-=Select BillNo=-"
Do Until rsd.EOF
    cbobillno.AddItem rsd("billno")
    rsd.MoveNext
Loop
cbobillno.ListIndex = 0
Call DisableAll
cmdsave.Enabled = False
cmdreset.Enabled = False
lblcount.Caption = "Total Record(s) Found : " & rs.RecordCount
lblstatus.Caption = "Payment Details : No Action"
End Sub

Public Sub ActiveAll()
Dim ctl As Object

For Each ctl In Me.Controls
    If TypeOf ctl Is TextBox Then
        ctl.Enabled = True
        ctl.BackColor = vbWhite
        DoEvents
    End If
Next ctl
cbopayment.Enabled = True
cbopayment.BackColor = vbWhite
End Sub

Public Sub DisableAll()
Dim ctl As Object

For Each ctl In Me.Controls
    If TypeOf ctl Is TextBox Then
        ctl.Enabled = False
        ctl.BackColor = vbButtonFace
        DoEvents
    End If
Next ctl
cbopayment.Enabled = False
cbopayment.BackColor = vbButtonFace
End Sub

Public Sub ClearFields()
On Error Resume Next
Dim ctlControl As Object

For Each ctlControl In Me.Controls
    If TypeOf ctlControl Is TextBox Then
        ctlControl.Text = ""
        DoEvents
    End If
Next ctlControl
cbopayment.Clear
End Sub

Public Sub GetPayment()
On Error Resume Next
Dim pmode As String

Call ActiveAll

cbopayment.Clear
With cbopayment
    .AddItem "-=Select Payment Mode=-"
    .AddItem "CASH"
    .AddItem "CHEQUE"
End With

Set recd = db.OpenRecordset("select *from info where billno=" & "'" & cbobillno.List(cbobillno.ListIndex) & "'")
If recd.RecordCount > 0 Then
    recd.MoveFirst
    txtchequeno.Text = recd("ch_no")
    txtbank.Text = recd("bank")
    txtbranch.Text = recd("branch")
    txtprice.Text = recd("price")
    txtpaid.Text = recd("paid")
    txtdue.Text = recd("due")
    pmode = recd("payment")
    lpaid = recd("lastpaid")
    If pmode = "" Then
        cbopayment.ListIndex = 0
    ElseIf pmode = "CASH" Then
        cbopayment.ListIndex = 1
    ElseIf pmode = "CHEQUE" Then
        cbopayment.ListIndex = 2
    End If
End If
Set recd = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cmdback.Enabled = True Then
    cmdback_Click
ElseIf cmdback.Enabled = False Then
    MsgBox "Exit not available now.", vbApplicationModal + vbExclamation, "Permission Denied"
    Cancel = vbNo
End If
End Sub

Private Sub txtpaid_Change()
txtdue.Text = Val(Trim(txtprice.Text)) - Val(Trim(txtpaid.Text))
End Sub

Private Sub txtpaid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdsave.SetFocus
End Sub

Private Sub txtpaid_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtprice_Change()
txtdue.Text = Val(Trim(txtprice.Text)) - Val(Trim(txtpaid.Text))
End Sub

Private Sub txtprice_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtpaid.SetFocus
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtbank_LostFocus()
txtbank.Text = StrConv(Trim(txtbank.Text), vbProperCase)
End Sub

Private Sub txtbranch_LostFocus()
txtbranch.Text = StrConv(Trim(txtbranch.Text), vbProperCase)
End Sub

Private Sub txtchequeno_LostFocus()
txtchequeno.Text = StrConv(Trim(txtchequeno.Text), vbProperCase)
End Sub

Public Sub DoChange()
If cbopayment.Text = "" Then
    MsgBox "ss"
Else
    Set rs = db.OpenRecordset("select * from info where billno='" & cbobillno.List(cbobillno.ListIndex) & "'")
End If
If rs.EditMode = dbEditNone Then
    rs.Edit
End If
End Sub

Public Sub DoSave()
confirm = MsgBox("Sure to save this payment details ?", vbApplicationModal + vbYesNo + vbQuestion, "Confirm Save Payment")
If confirm = vbNo Then
    cbopayment.SetFocus
    lblstatus.Caption = "Service Status : Changing Payment Details"
    Exit Sub
ElseIf confirm = vbYes Then
    If cmdsave.Caption = "Save" Then
        rs("payment") = cbopayment.Text
        rs("ch_no") = Trim(txtchequeno.Text)
        rs("bank") = Trim(txtbank.Text)
        rs("branch") = Trim(txtbranch.Text)
        rs("price") = Trim(txtprice.Text)
        rs("paid") = Trim(txtpaid.Text)
        rs("due") = Trim(txtdue.Text)
        rs("lastpaid") = Val(Trim(lpaid))
        If Val(Trim(txtdue.Text)) = 0 Then
            rs("remarks") = "Full Paid"
        Else
            rs("remarks") = "Due"
        End If
        rs.Update
        Call ClearFields
        Call DisableAll
        cmdsave.Enabled = False
        cmdreset.Enabled = False
        cmdback.Enabled = True
        cbobillno.Enabled = True
        Label13.Enabled = True
        cbobillno.ListIndex = 0
        cbobillno.SetFocus
        lblstatus.Caption = "Service Status : Payment Details Changed"
    End If
End If
End Sub

Private Sub txtchequeno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtbank.SetFocus
End Sub

Private Sub txtbank_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtbranch.SetFocus
End Sub

Private Sub txtbranch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtprice.SetFocus
End Sub
