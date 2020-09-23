VERSION 5.00
Begin VB.Form frmentry 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Customer Information"
   ClientHeight    =   5610
   ClientLeft      =   2775
   ClientTop       =   1875
   ClientWidth     =   6540
   Icon            =   "frmentry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6540
   Begin VB.ComboBox cbostyear 
      Height          =   315
      Left            =   5385
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   3540
      Width           =   1095
   End
   Begin VB.ComboBox cbostmonth 
      Height          =   315
      Left            =   4110
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   3540
      Width           =   675
   End
   Begin VB.ComboBox cbostday 
      Height          =   315
      Left            =   2700
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   3540
      Width           =   675
   End
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
      Height          =   5595
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6555
      Begin VB.ComboBox cboendyear 
         Height          =   315
         Left            =   5385
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   3885
         Width           =   1095
      End
      Begin VB.ComboBox cboendmonth 
         Height          =   315
         Left            =   4110
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   3870
         Width           =   675
      End
      Begin VB.ComboBox cboendday 
         Height          =   315
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   3870
         Width           =   675
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
         TabIndex        =   3
         Top             =   4905
         Width           =   945
      End
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
         Left            =   3585
         TabIndex        =   1
         Top             =   4905
         Width           =   945
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add"
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
         TabIndex        =   0
         Top             =   4905
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
         Left            =   4560
         TabIndex        =   2
         Top             =   4905
         Width           =   945
      End
      Begin VB.TextBox txtaddress 
         Height          =   1110
         Left            =   2235
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   1290
         Width           =   4245
      End
      Begin VB.TextBox txtdescription 
         Height          =   1110
         Left            =   2235
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   2415
         Width           =   4245
      End
      Begin VB.TextBox txtemail 
         Height          =   300
         Left            =   2235
         TabIndex        =   10
         Top             =   4545
         Width           =   4245
      End
      Begin VB.TextBox txtcompanyname 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2235
         TabIndex        =   6
         Top             =   960
         Width           =   4245
      End
      Begin VB.TextBox txtdomain 
         Height          =   300
         Left            =   2235
         TabIndex        =   9
         Top             =   4215
         Width           =   4245
      End
      Begin VB.TextBox txtdate 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2235
         TabIndex        =   4
         Top             =   300
         Width           =   4245
      End
      Begin VB.TextBox txtbillno 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   630
         Width           =   4245
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
         Left            =   3495
         TabIndex        =   31
         Top             =   3930
         Width           =   540
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
         Left            =   4905
         TabIndex        =   30
         Top             =   3930
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
         Left            =   2235
         TabIndex        =   29
         Top             =   3930
         Width           =   405
      End
      Begin VB.Label Label19 
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
         Left            =   3495
         TabIndex        =   28
         Top             =   3615
         Width           =   540
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
         TabIndex        =   27
         Top             =   3615
         Width           =   390
      End
      Begin VB.Label Label17 
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
         Left            =   2235
         TabIndex        =   26
         Top             =   3615
         Width           =   405
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
         TabIndex        =   22
         Top             =   5340
         Width           =   660
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
         TabIndex        =   21
         Top             =   5340
         Width           =   660
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000A&
         BorderWidth     =   2
         Height          =   15
         Left            =   15
         Top             =   5295
         Width           =   6525
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
         TabIndex        =   20
         Top             =   1035
         Width           =   1410
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
         TabIndex        =   19
         Top             =   1335
         Width           =   780
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
         Top             =   4605
         Width           =   600
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
         TabIndex        =   17
         Top             =   4260
         Width           =   735
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
         TabIndex        =   16
         Top             =   720
         Width           =   585
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
         TabIndex        =   15
         Top             =   2445
         Width           =   1050
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
         TabIndex        =   14
         Top             =   3915
         Width           =   1530
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
         TabIndex        =   13
         Top             =   3615
         Width           =   1665
      End
      Begin VB.Label Label13 
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
         TabIndex        =   12
         Top             =   390
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database, rs As Recordset
Dim stdate, enddate As Variant
Dim campday, campmonth As Integer

Public Sub GenBillNo()
Dim lastbno, bnochar, curbno As Variant
Dim bnonum, l As Integer
Dim rsd As Recordset

Set rsd = db.OpenRecordset("info", dbOpenTable)
If rsd.RecordCount = 0 Then
    curbno = "MB0112"
    txtbillno.Text = curbno
    Exit Sub
ElseIf rsd.RecordCount > 0 Then
    rsd.MoveLast
    lastbno = rsd("billno")
    bnochar = Left(lastbno, 3)
    l = Len(lastbno)
    bnonum = Mid(lastbno, 4, l) + 1
    curbno = bnochar & bnonum
    txtbillno.Text = curbno
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
Label17.Enabled = True
Label18.Enabled = True
Label19.Enabled = True
Label20.Enabled = True
Label21.Enabled = True
Label22.Enabled = True
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
Label17.Enabled = False
Label18.Enabled = False
Label19.Enabled = False
Label20.Enabled = False
Label21.Enabled = False
Label22.Enabled = False
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
cbostday.Clear
cbostmonth.Clear
cbostyear.Clear
cboendday.Clear
cboendmonth.Clear
cboendyear.Clear
End Sub

Private Sub cmdadd_Click()
Dim i, j, k, l As Integer

rs.AddNew
Call ActiveAll
Call GenBillNo
txtdate.Text = Format(Date, "dd/mm/yyyy")
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
cbostday.ListIndex = campday - 1

For j = 1 To 12 Step 1
    If j < 10 Then
        cbostmonth.AddItem "0" & j
    ElseIf j >= 10 Then
        cbostmonth.AddItem j
    End If
Next j
cbostmonth.ListIndex = campmonth - 1

For k = 1 To 31 Step 1
    If k < 10 Then
        cboendday.AddItem "0" & k
    ElseIf k >= 10 Then
        cboendday.AddItem k
    End If
Next k
cboendday.ListIndex = campday - 1

For l = 1 To 12 Step 1
    If l < 10 Then
        cboendmonth.AddItem "0" & l
    ElseIf l >= 10 Then
        cboendmonth.AddItem l
    End If
Next l
cboendmonth.ListIndex = campmonth - 1

cbostyear.AddItem Year(Date)
cboendyear.AddItem Year(Date)
cbostyear.ListIndex = 0
cboendyear.ListIndex = 0

txtdate.SetFocus
txtdate.SelStart = 0
txtdate.SelLength = Len(Trim(txtdate.Text))
lblstatus.Caption = "Service Status : Adding New Data"
cmdadd.Enabled = False
cmdback.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
End Sub

Private Sub cmdback_Click()
Unload Me
frmmain.Show
End Sub

Private Sub cmdcancel_Click()
Call ClearFields
Call DisableAll
cmdadd.Enabled = True
cmdback.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
lblstatus.Caption = "Service Status : Addition Cancelled"
End Sub

Private Sub cmdsave_Click()
On Error GoTo SaveError

Call CheckDates

SaveError:
If Err.Number = 3421 Then
    MsgBox "CustInfo cannot save this record.", vbApplicationModal + vbCritical, "Error In Saving Data"
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

lblstatus.Caption = "Addition Status : No Action"
lblcount.Caption = "So Far Added : " & rs.RecordCount
cmdsave.Enabled = False
cmdcancel.Enabled = False
Call DisableAll
campday = Val(Day(Date))
campmonth = Val(Month(Date))
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cmdback.Enabled = True Then
    cmdback_Click
ElseIf cmdback.Enabled = False Then
    MsgBox "Exit not available now.", vbApplicationModal + vbExclamation, "Permission Denied"
    Cancel = vbNo
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
    Call DoSave
    Exit Sub
End If
End Sub

Public Sub DoSave()
confirm = MsgBox("Sure to save this record ?", vbApplicationModal + vbYesNo + vbQuestion, "Confirm Save Record")
If confirm = vbNo Then
    txtdate.SetFocus
    Exit Sub
ElseIf confirm = vbYes Then
    If rs.EditMode = dbEditAdd Then
        rs("date") = Trim(txtdate.Text)
        rs("billno") = Trim(txtbillno.Text)
        rs("companyname") = Trim(txtcompanyname.Text)
        rs("address") = Trim(txtaddress.Text)
        rs("description") = Trim(txtdescription.Text)
        rs("startcamp") = Trim(stdate)
        rs("endcamp") = Trim(enddate)
        rs("domain") = Trim(txtdomain.Text)
        rs("email") = Trim(txtemail.Text)
        rs("inv_status") = "Invoice Due"
        rs("lastpaid") = 0
        rs.Update
    End If
    lblstatus.Caption = "Service Status : Record Saved"
    lblcount.Caption = "So Far Added : " & rs.RecordCount
    Call ClearFields
    Call DisableAll
    cmdsave.Enabled = False
    cmdcancel.Enabled = False
    cmdadd.Enabled = True
    cmdback.Enabled = True
End If
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
If KeyCode = vbKeyReturn Then cmdsave.SetFocus
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
