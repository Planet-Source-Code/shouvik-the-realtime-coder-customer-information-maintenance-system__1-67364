VERSION 5.00
Begin VB.Form frmedit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Existing Customer Information"
   ClientHeight    =   5940
   ClientLeft      =   2295
   ClientTop       =   1710
   ClientWidth     =   6555
   Icon            =   "frmedit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   6555
   Begin VB.Frame Frame1 
      Caption         =   "Customer Info Sheet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5550
      Left            =   0
      TabIndex        =   13
      Top             =   375
      Width           =   6570
      Begin VB.ComboBox cbostday 
         Height          =   315
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   3540
         Width           =   675
      End
      Begin VB.ComboBox cbostmonth 
         Height          =   315
         Left            =   4125
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   3555
         Width           =   675
      End
      Begin VB.ComboBox cbostyear 
         Height          =   315
         Left            =   5385
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   3555
         Width           =   1095
      End
      Begin VB.ComboBox cboendday 
         Height          =   315
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3885
         Width           =   675
      End
      Begin VB.ComboBox cboendmonth 
         Height          =   315
         Left            =   4125
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3885
         Width           =   675
      End
      Begin VB.ComboBox cboendyear 
         Height          =   315
         Left            =   5385
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3885
         Width           =   1095
      End
      Begin VB.TextBox txtbillno 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   4245
      End
      Begin VB.TextBox txtdate 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2235
         TabIndex        =   5
         Top             =   300
         Width           =   4245
      End
      Begin VB.TextBox txtdomain 
         Height          =   300
         Left            =   2235
         TabIndex        =   10
         Top             =   4230
         Width           =   4245
      End
      Begin VB.TextBox txtcompanyname 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2235
         TabIndex        =   7
         Top             =   960
         Width           =   4245
      End
      Begin VB.TextBox txtemail 
         Height          =   300
         Left            =   2235
         TabIndex        =   11
         Top             =   4560
         Width           =   4245
      End
      Begin VB.TextBox txtdescription 
         Height          =   1110
         Left            =   2235
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   2430
         Width           =   4245
      End
      Begin VB.TextBox txtaddress 
         Height          =   1110
         Left            =   2235
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   1290
         Width           =   4245
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
         Left            =   4560
         TabIndex        =   3
         Top             =   4905
         Width           =   945
      End
      Begin VB.CommandButton cmdmodify 
         Caption         =   "Modify"
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
         Left            =   2610
         TabIndex        =   1
         Top             =   4905
         Width           =   945
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "Update"
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
         Left            =   3585
         TabIndex        =   2
         Top             =   4905
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
         Left            =   5535
         TabIndex        =   4
         Top             =   4905
         Width           =   945
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         Left            =   3480
         TabIndex        =   33
         Top             =   3615
         Width           =   540
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   2250
         TabIndex        =   32
         Top             =   3615
         Width           =   405
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   4905
         TabIndex        =   31
         Top             =   3615
         Width           =   390
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   2250
         TabIndex        =   30
         Top             =   3945
         Width           =   405
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   4920
         TabIndex        =   29
         Top             =   3945
         Width           =   390
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         Left            =   3480
         TabIndex        =   28
         Top             =   3945
         Width           =   540
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
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
         Left            =   135
         TabIndex        =   24
         Top             =   390
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Campaigning :"
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
         Left            =   135
         TabIndex        =   23
         Top             =   3660
         Width           =   1665
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Campaigning :"
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
         Left            =   135
         TabIndex        =   22
         Top             =   3990
         Width           =   1530
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
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
         Left            =   135
         TabIndex        =   21
         Top             =   2445
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill No :"
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
         Left            =   135
         TabIndex        =   20
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Domain :"
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
         Left            =   135
         TabIndex        =   19
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail :"
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
         Left            =   135
         TabIndex        =   18
         Top             =   4650
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address :"
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
         Left            =   135
         TabIndex        =   17
         Top             =   1335
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name :"
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
         Left            =   135
         TabIndex        =   16
         Top             =   1035
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000A&
         BorderWidth     =   2
         Height          =   15
         Left            =   15
         Top             =   5280
         Width           =   6525
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
         TabIndex        =   15
         Top             =   5310
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
         Left            =   5850
         TabIndex        =   14
         Top             =   5310
         Width           =   660
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
      Left            =   2835
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   30
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
      Left            =   705
      TabIndex        =   12
      Top             =   90
      Width           =   1980
   End
End
Attribute VB_Name = "frmedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database, rs As Recordset, rsd As Recordset
Dim stdate, enddate, lpaid As Variant, invoice As String
Dim stday, stmonth, styear, endday, endmonth, endyear As Integer

Private Sub cbobillno_Click()
Call GetCustomerData
End Sub

Private Sub cmdback_Click()
Unload Me
frmmain.Show
End Sub

Private Sub cmdcancel_Click()
Call ClearFields
Call DisableAll
cmdmodify.Enabled = True
cmdupdate.Enabled = False
cmdcancel.Enabled = False
cmdback.Enabled = True
cbobillno.Enabled = True
cbobillno.ListIndex = 0
cbobillno.SetFocus
Label13.Enabled = True
lblstatus.Caption = "Service Status : Data Modification Cancelled"
End Sub

Private Sub cmdmodify_Click()
If cbobillno.Text = "-=Select BillNo=-" Then
    MsgBox "Please select a customer billbo for modification.", vbApplicationModal + vbExclamation, "No Bill Selected"
    cbobillno.SetFocus
Else
    Call DoEditing
End If
End Sub

Private Sub cmdupdate_Click()
On Error GoTo UpdateError

Call CheckDates

UpdateError:
If Err.Number = 3421 Then
    MsgBox "CustInfo cannot update this record.", vbApplicationModal + vbCritical, "Error in updating data"
    cmdcancel_Click
    cmdback_Click
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

Set db = OpenDatabase(App.Path & "\company.mdb")
Set rs = db.OpenRecordset("info", dbOpenTable)
If rs.RecordCount > 0 Then rs.MoveFirst

cbobillno.Clear
Do Until rs.EOF
    cbobillno.AddItem rs("billno")
    rs.MoveNext
Loop
cbobillno.AddItem "-=Select BillNo=-", 0
cbobillno.ListIndex = 0

lblstatus.Caption = "Modification Status : No Action"
lblcount.Caption = "Total " & rs.RecordCount & " Record(s) Found"
cmdupdate.Enabled = False
cmdcancel.Enabled = False

Call DisableAll
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cmdback.Enabled = True Then
    cmdback_Click
ElseIf cmdback.Enabled = False Then
    MsgBox "Exit not available now.", vbApplicationModal + vbExclamation, "Permission Denied"
    Cancel = vbNo
End If
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
cbostday.Enabled = True
cbostday.BackColor = vbWhite
cboendday.Enabled = True
cboendday.BackColor = vbWhite
cbostmonth.Enabled = True
cbostmonth.BackColor = vbWhite
cboendmonth.Enabled = True
cboendmonth.BackColor = vbWhite
cbostyear.Enabled = True
cbostyear.BackColor = vbWhite
cboendyear.Enabled = True
cboendyear.BackColor = vbWhite
Label18.Enabled = True
Label19.Enabled = True
Label20.Enabled = True
Label21.Enabled = True
Label22.Enabled = True
Label23.Enabled = True
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
cbostday.Enabled = False
cbostday.BackColor = vbButtonFace
cboendday.Enabled = False
cboendday.BackColor = vbButtonFace
cbostmonth.Enabled = False
cbostmonth.BackColor = vbButtonFace
cboendmonth.Enabled = False
cboendmonth.BackColor = vbButtonFace
cbostyear.Enabled = False
cbostyear.BackColor = vbButtonFace
cboendyear.Enabled = False
cboendyear.BackColor = vbButtonFace
Label18.Enabled = False
Label19.Enabled = False
Label20.Enabled = False
Label21.Enabled = False
Label22.Enabled = False
Label23.Enabled = False
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
cbobillno.Enabled = True
cbobillno.ListIndex = 0
Label13.Enabled = True
cbostday.Clear
cbostmonth.Clear
cbostyear.Clear
cboendday.Clear
cboendmonth.Clear
cboendyear.Clear
End Sub

Public Sub GetCustomerData()
On Error Resume Next
Dim i, j, k, l As Integer

cbostday.Clear
cbostmonth.Clear
cbostyear.Clear
cboendday.Clear
cboendmonth.Clear
cboendyear.Clear

For i = 1 To 31 Step 1
    If i < 10 Then
        cbostday.AddItem "0" & i
    ElseIf i >= 10 Then
        cbostday.AddItem i
    End If
Next i

For j = 1 To 12 Step 1
    If j < 10 Then
        cbostmonth.AddItem "0" & j
    ElseIf j >= 10 Then
        cbostmonth.AddItem j
    End If
Next j

For k = 1 To 31 Step 1
    If k < 10 Then
        cboendday.AddItem "0" & k
    ElseIf k >= 10 Then
        cboendday.AddItem k
    End If
Next k

For l = 1 To 12 Step 1
    If l < 10 Then
        cboendmonth.AddItem "0" & l
    ElseIf l >= 10 Then
        cboendmonth.AddItem l
    End If
Next l

Set rsd = db.OpenRecordset("select *from info where billno='" & cbobillno.List(cbobillno.ListIndex) & "'", dbOpenDynaset)
If rsd.RecordCount > 0 Then
    rsd.MoveFirst
    txtdate.Text = rsd("date")
    txtbillno.Text = rsd("billno")
    txtcompanyname.Text = rsd("companyname")
    txtaddress.Text = rsd("address")
    txtdescription.Text = rsd("description")
    txtdomain.Text = rsd("domain")
    txtemail.Text = rsd("email")
    stdate = rsd("startcamp")
    enddate = rsd("endcamp")
    invoice = rsd("inv_status")
    lpaid = rsd("lastpaid")
    Call SetDates
    Set rsd = Nothing
End If
End Sub

Public Sub DoEditing()
If cmdmodify.Caption = "Modify" Then
    If txtbillno.Text = "" Then
        MsgBox "ss"
    Else
        Set rs = db.OpenRecordset("select * from info where billno='" & cbobillno.List(cbobillno.ListIndex) & "'")
    End If
    confirm = MsgBox("Sure to modify bill " & txtbillno.Text & " ?", vbYesNo + vbQuestion, "Confirm Modify")
    If confirm = vbNo Then
        Call ClearFields
        cbobillno.ListIndex = 0
        cbobillno.SetFocus
        lblstatus.Caption = "Service Status : Data Modification Ignored"
        Exit Sub
    ElseIf confirm = vbYes Then
        Call ActiveAll
        cmdmodify.Enabled = False
        Label13.Enabled = False
        cbobillno.Enabled = False
        cmdback.Enabled = False
        cmdcancel.Enabled = True
        cmdupdate.Enabled = True
        lblstatus.Caption = "Service Status : Modifying Customer Data"
    End If
    If rs.EditMode = dbEditNone Then
        rs.Edit
        txtdate.SelStart = 0
        txtdate.SelLength = Len(Trim(txtdate.Text))
        txtdate.SetFocus
    End If
End If
End Sub

Public Sub DoUpdate()
confirm = MsgBox("Sure to update bill " & txtbillno.Text & " ?", vbApplicationModal + vbYesNo + vbQuestion, "Confirm Update")
If confirm = vbNo Then
    txtdate.SetFocus
    lblstatus.Caption = "Service Status : Modifying Customer Data"
    Exit Sub
ElseIf confirm = vbYes Then
    If cmdupdate.Caption = "Update" Then
        rs("date") = Trim(txtdate.Text)
        rs("billno") = Trim(txtbillno.Text)
        rs("companyname") = Trim(txtcompanyname.Text)
        rs("address") = Trim(txtaddress.Text)
        rs("description") = Trim(txtdescription.Text)
        rs("startcamp") = Trim(stdate)
        rs("endcamp") = Trim(enddate)
        rs("domain") = Trim(txtdomain.Text)
        rs("email") = Trim(txtemail.Text)
        rs("inv_status") = Trim(invoice)
        rs("lastpaid") = Val(Trim(lpaid))
        rs.Update
        Call ClearFields
        Call DisableAll
        cmdmodify.Enabled = True
        cmdupdate.Enabled = False
        cmdcancel.Enabled = False
        cmdback.Enabled = True
        Label13.Enabled = True
        cbobillno.Enabled = True
        cbobillno.ListIndex = 0
        cbobillno.SetFocus
        lblstatus.Caption = "Service Status : Customer Data Updated"
    End If
End If
End Sub

Private Sub txtaddress_LostFocus()
txtaddress.Text = StrConv(Trim(txtaddress.Text), vbProperCase)
End Sub

Private Sub txtcompanyname_LostFocus()
txtcompanyname.Text = StrConv(Trim(txtcompanyname.Text), vbProperCase)
End Sub

Private Sub txtdate_LostFocus()
Call CheckPaymentDate
End Sub

Private Sub txtdescription_LostFocus()
txtdescription.Text = StrConv(Trim(txtdescription.Text), vbProperCase)
End Sub

Public Sub SetDates()
stday = Val(Left(Trim(stdate), 2))
If stday < 10 Then
    stday = Right(Trim(stday), 1)
ElseIf stday >= 10 Then
    stday = stday
End If

stmonth = Val(Mid(Trim(stdate), 4, 2))
If stmonth < 10 Then
    stmonth = Right(Trim(stmonth), 1)
ElseIf stmonth >= 10 Then
    stmonth = stmonth
End If

styear = Val(Right(Trim(stdate), 4))

endday = Val(Left(Trim(enddate), 2))
If endday < 10 Then
    endday = Right(Trim(endday), 1)
ElseIf endday >= 10 Then
    endday = endday
End If

endmonth = Val(Mid(Trim(enddate), 4, 2))
If endmonth < 10 Then
    endmonth = Right(Trim(endmonth), 1)
ElseIf endmonth >= 10 Then
    endmonth = endmonth
End If

endyear = Val(Right(Trim(enddate), 4))

cbostday.ListIndex = stday - 1
cbostmonth.ListIndex = stmonth - 1
cboendday.ListIndex = endday - 1
cboendmonth.ListIndex = endmonth - 1

cbostyear.AddItem styear
cboendyear.AddItem endyear
cbostyear.ListIndex = 0
cboendyear.ListIndex = 0
End Sub

Private Sub txtdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtbillno.SetFocus
End Sub

Private Sub txtbillno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtcompanyname.SetFocus
End Sub

Private Sub txtcompanyname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtaddress.SetFocus
End Sub

Private Sub txtaddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtdescription.SetFocus
End Sub

Private Sub txtdescription_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cbostday.SetFocus
End Sub

Private Sub cbostday_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cbostmonth.SetFocus
End Sub

Private Sub cbostmonth_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cbostyear.SetFocus
End Sub

Private Sub cbostyear_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboendday.SetFocus
End Sub

Private Sub cboendday_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboendmonth.SetFocus
End Sub

Private Sub cboendmonth_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboendyear.SetFocus
End Sub

Private Sub cboendyear_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtdomain.SetFocus
End Sub

Private Sub txtdomain_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtemail.SetFocus
End Sub

Private Sub txtemail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdupdate.SetFocus
End Sub

Private Sub cbostday_Click()
stdate = cbostday.Text & "/" & cbostmonth.Text & "/" & cbostyear.Text
enddate = cboendday.Text & "/" & cboendmonth.Text & "/" & cboendyear.Text
End Sub

Private Sub cbostmonth_Click()
stdate = cbostday.Text & "/" & cbostmonth.Text & "/" & cbostyear.Text
enddate = cboendday.Text & "/" & cboendmonth.Text & "/" & cboendyear.Text
End Sub

Private Sub cbostyear_Click()
stdate = cbostday.Text & "/" & cbostmonth.Text & "/" & cbostyear.Text
enddate = cboendday.Text & "/" & cboendmonth.Text & "/" & cboendyear.Text
End Sub

Private Sub cboendday_Click()
stdate = cbostday.Text & "/" & cbostmonth.Text & "/" & cbostyear.Text
enddate = cboendday.Text & "/" & cboendmonth.Text & "/" & cboendyear.Text
End Sub

Private Sub cboendmonth_Click()
stdate = cbostday.Text & "/" & cbostmonth.Text & "/" & cbostyear.Text
enddate = cboendday.Text & "/" & cboendmonth.Text & "/" & cboendyear.Text
End Sub

Private Sub cboendyear_Click()
stdate = cbostday.Text & "/" & cbostmonth.Text & "/" & cbostyear.Text
enddate = cboendday.Text & "/" & cboendmonth.Text & "/" & cboendyear.Text
End Sub

Public Sub CheckPaymentDate()
If Trim(txtdate.Text) <> "" Then
    If Not IsDate(txtdate.Text) Then
        MsgBox "The date you entered is invalid.", vbApplicationModal + vbCritical, "Invalid Date"
        txtdate.SelStart = 0
        txtdate.SelLength = Len(Trim(txtdate.Text))
        txtdate.SetFocus
    ElseIf IsDate(txtdate.Text) Then
        If Trim(txtdate.Text) = Format(Trim(txtdate.Text), "dd/mm/yyyy") And Len(Trim(txtdate.Text)) = 10 And _
              Trim(txtdate.Text) >= Format(Date, "dd/mm/yyyy") Then
            Exit Sub
        Else
            MsgBox "The date format is invalid.Maintain the following rules while entering the date :-" & Chr(10) & Chr(10) _
            & "1. The date should be in dd/mm/yyyy format." & Chr(10) & "2. Its length should be 10." & Chr(10) _
            & "3.The date should be greater than or equal to current date." & Chr(10) & Chr(10) _
            & Chr(10) & "For example,15/08/2006 - represents 15th of August,2006 .", vbApplicationModal + vbExclamation, "Invalid date format"
            txtdate.SelStart = 0
            txtdate.SelLength = Len(Trim(txtdate.Text))
            txtdate.SetFocus
        End If
    End If
ElseIf Trim(txtdate.Text) = "" Then
    Exit Sub
End If
End Sub

Public Sub CheckDates()
If CDate(stdate) < Format(Date, "dd/mm/yyyy") And CDate(enddate) < Format(Date, "dd/mm/yyyy") Then
    MsgBox "The camp dates must be greater than current date.", vbApplicationModal + vbInformation, "Wrong Date Range"
    cbostday.SetFocus
ElseIf CDate(stdate) < Format(Date, "dd/mm/yyyy") And CDate(enddate) >= Format(Date, "dd/mm/yyyy") Then
    MsgBox "Start Campaigning date must be greater than current date.", vbApplicationModal + vbInformation, "Wrong Date Range"
    cbostday.SetFocus
ElseIf CDate(stdate) >= Format(Date, "dd/mm/yyyy") And CDate(enddate) < Format(Date, "dd/mm/yyyy") Then
    MsgBox "End Campaigning date must be greater than current date.", vbApplicationModal + vbInformation, "Wrong Date Range"
    cboendday.SetFocus
ElseIf CDate(stdate) = CDate(enddate) Then
    MsgBox "Start and End Campaigning dates must be different.", vbApplicationModal + vbInformation, "Wrong Date Range"
    cbostday.SetFocus
ElseIf CDate(stdate) > CDate(enddate) Then
    MsgBox "Start Campaigning date must be lower than End Campaigning date.", vbApplicationModal + vbInformation, "Wrong Date Range"
    cbostday.SetFocus
ElseIf CDate(stdate) >= Format(Date, "dd/mm/yyyy") And CDate(enddate) >= Format(Date, "dd/mm/yyyy") And CDate(stdate) <> CDate(enddate) _
            And CDate(stdate) < CDate(enddate) Then
    Call DoUpdate
    Exit Sub
End If
End Sub
