VERSION 5.00
Begin VB.Form frmrenewadd 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Renew Customer Membership"
   ClientHeight    =   6885
   ClientLeft      =   2055
   ClientTop       =   1185
   ClientWidth     =   6585
   Icon            =   "frmrenewadd.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   6585
   Begin VB.Frame Frame1 
      Caption         =   "Customer Renewal Info Sheet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6885
      Left            =   0
      TabIndex        =   18
      Top             =   -15
      Width           =   6600
      Begin VB.TextBox txtdate 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2235
         TabIndex        =   43
         Top             =   300
         Width           =   4245
      End
      Begin VB.TextBox txtdue 
         Height          =   300
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   5880
         Width           =   4245
      End
      Begin VB.TextBox txtpaid 
         Height          =   300
         Left            =   2235
         TabIndex        =   14
         Top             =   5550
         Width           =   4245
      End
      Begin VB.TextBox txtprice 
         Height          =   300
         Left            =   2235
         TabIndex        =   13
         Top             =   5220
         Width           =   4245
      End
      Begin VB.ComboBox cbopayment 
         Height          =   315
         Left            =   2235
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4875
         Width           =   4245
      End
      Begin VB.ComboBox cbostday 
         Height          =   315
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3540
         Width           =   675
      End
      Begin VB.ComboBox cbostmonth 
         Height          =   315
         Left            =   4110
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3540
         Width           =   675
      End
      Begin VB.ComboBox cbostyear 
         Height          =   315
         Left            =   5385
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3540
         Width           =   1095
      End
      Begin VB.TextBox txtbillno 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   630
         Width           =   4245
      End
      Begin VB.TextBox txtdomain 
         Height          =   300
         Left            =   2235
         TabIndex        =   10
         Top             =   4215
         Width           =   4245
      End
      Begin VB.TextBox txtcompanyname 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2235
         TabIndex        =   1
         Top             =   960
         Width           =   4245
      End
      Begin VB.TextBox txtemail 
         Height          =   300
         Left            =   2235
         TabIndex        =   11
         Top             =   4545
         Width           =   4245
      End
      Begin VB.TextBox txtdescription 
         Height          =   1110
         Left            =   2235
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   2415
         Width           =   4260
      End
      Begin VB.TextBox txtaddress 
         Height          =   1110
         Left            =   2235
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   1290
         Width           =   4260
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
         Left            =   5505
         TabIndex        =   17
         Top             =   6225
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
         Left            =   4530
         TabIndex        =   16
         Top             =   6225
         Width           =   945
      End
      Begin VB.ComboBox cboendday 
         Height          =   315
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3870
         Width           =   675
      End
      Begin VB.ComboBox cboendmonth 
         Height          =   315
         Left            =   4110
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3870
         Width           =   675
      End
      Begin VB.ComboBox cboendyear 
         Height          =   315
         Left            =   5385
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3885
         Width           =   1095
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
         Left            =   1965
         TabIndex        =   42
         Top             =   5955
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
         Left            =   1965
         TabIndex        =   41
         Top             =   5625
         Width           =   255
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
         Left            =   1965
         TabIndex        =   40
         Top             =   5280
         Width           =   255
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
         Left            =   135
         TabIndex        =   39
         Top             =   4965
         Width           =   1365
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
         Left            =   135
         TabIndex        =   38
         Top             =   5955
         Width           =   420
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
         Left            =   135
         TabIndex        =   37
         Top             =   5625
         Width           =   450
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
         Left            =   135
         TabIndex        =   36
         Top             =   5280
         Width           =   1245
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   3615
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
         TabIndex        =   33
         Top             =   3915
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   4305
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
         TabIndex        =   29
         Top             =   4635
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   1035
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000A&
         BorderWidth     =   2
         Height          =   15
         Left            =   15
         Top             =   6600
         Width           =   6570
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
         TabIndex        =   26
         Top             =   6630
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
         TabIndex        =   25
         Top             =   6630
         Width           =   660
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   3615
         Width           =   390
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
         TabIndex        =   22
         Top             =   3615
         Width           =   540
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
         TabIndex        =   21
         Top             =   3930
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
         Left            =   4905
         TabIndex        =   20
         Top             =   3930
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
         Left            =   3495
         TabIndex        =   19
         Top             =   3930
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmrenewadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database, rs As Recordset, recd As Recordset
Dim stdate, enddate, todate, lpaid As Variant, pbox As String
Dim campday, campmonth As Integer

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

Private Sub cbopayment_Click()
If pbox = "" Then
    Exit Sub
ElseIf pbox = "got" Then
    If cbopayment.ListIndex = 0 Then
        Exit Sub
    ElseIf cbopayment.ListIndex = 1 Then
        txtprice.SetFocus
    ElseIf cbopayment.ListIndex = 2 Then
        frmcheque.Show vbModal
    End If
End If
End Sub

Private Sub cbopayment_GotFocus()
pbox = "got"
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

Private Sub cmdcancel_Click()
Call ClearFields
Unload Me
frmrenewal.Show
frmrenewal.cbobno.ListIndex = 0
frmrenewal.cbobno.SetFocus
End Sub

Private Sub cmdsave_Click()
On Error GoTo SaveError

If cbopayment.ListIndex = 0 Then
    MsgBox "Select a proper payment mode either cash or cheque.", vbApplicationModal + vbInformation, "No payment mode selected"
    cbopayment.ListIndex = 0
ElseIf cbopayment.ListIndex = 1 Or cbopayment.ListIndex = 2 Then
    Call CheckDates
End If

SaveError:
If Err.Number = 3421 Then
    MsgBox "CustInfo cannot save this record.", vbApplicationModal + vbCritical, "Error In Saving Data"
    Call ClearFields
    Unload Me
    Unload frmrenewal
    frmmain.Show
    Exit Sub
End If
End Sub

Private Sub Form_Activate()
Call GetData
Call SetDates
Call GenBillNo
pbox = ""
rs.AddNew
txtdate.SelStart = 0
txtdate.SelLength = Len(Trim(txtdate.Text))
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

Set db = OpenDatabase(App.Path & "\company.mdb")
Set rs = db.OpenRecordset("info", dbOpenTable)
Set recd = db.OpenRecordset("info", dbOpenTable)
If rs.RecordCount > 0 Then rs.MoveFirst
If recd.RecordCount > 0 Then recd.MoveFirst

lblstatus.Caption = "Service Status : Renewing Customer Data"
lblcount.Caption = "So Far Added : " & rs.RecordCount
campday = Val(Day(Date))
campmonth = Val(Month(Date))

With cbopayment
    .AddItem "-=Select Payment Mode=-"
    .AddItem "CASH"
    .AddItem "CHEQUE"
End With
txtdate.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdcancel_Click
End Sub

Private Sub txtdate_LostFocus()
Call CheckPaymentDate
End Sub

Private Sub txtpaid_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
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

Private Sub txtprice_Change()
txtdue.Text = Val(Trim(txtprice.Text)) - Val(Trim(txtpaid.Text))
End Sub

Private Sub txtpaid_Change()
txtdue.Text = Val(Trim(txtprice.Text)) - Val(Trim(txtpaid.Text))
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
cbostday.Clear
cbostmonth.Clear
cbostyear.Clear
cboendday.Clear
cboendmonth.Clear
cboendyear.Clear
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

Public Sub GenBillNo()
Dim lastbno, bnochar, curbno As Variant
Dim bnonum, l As Integer
Dim rsd As Recordset

Set rsd = db.OpenRecordset("info", dbOpenTable)
rsd.MoveLast
lastbno = rsd("billno")
bnochar = Left(lastbno, 3)
l = Len(lastbno)
bnonum = Mid(lastbno, 4, l) + 1
curbno = bnochar & bnonum
txtbillno.Text = curbno
End Sub

Public Sub SetDates()
Dim i, j, k, l As Integer

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
End Sub

Public Sub GetData()
On Error Resume Next
Dim pmode As String

Set recd = db.OpenRecordset("select *from info where billno=" & "'" & selbno & "'")
If recd.RecordCount > 0 Then recd.MoveFirst
    txtcompanyname.Text = recd("companyname")
    txtaddress.Text = recd("address")
    txtdescription.Text = recd("description")
    txtdomain.Text = recd("domain")
    txtemail.Text = recd("email")
    txtprice.Text = recd("price")
    txtpaid.Text = recd("paid")
    txtdue.Text = recd("due")
    pmode = recd("payment")
    If pmode = "CASH" Then
        cbopayment.ListIndex = 1
    ElseIf pmode = "CHEQUE" Then
        cbopayment.ListIndex = 2
    ElseIf pmode = "" Then
        cbopayment.ListIndex = 0
    End If
    recchno = recd("ch_no")
    recbank = recd("bank")
    recbranch = recd("branch")
    txtdate.SetFocus
Set recd = Nothing
End Sub

Public Sub DoSave()
confirm = MsgBox("Sure to save this renewal ?", vbApplicationModal + vbYesNo + vbQuestion, "Confirm Save Renewal")
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
        rs("payment") = cbopayment.List(cbopayment.ListIndex)
        rs("price") = Trim(txtprice.Text)
        rs("paid") = Trim(txtpaid.Text)
        rs("due") = Trim(txtdue.Text)
        rs("ch_no") = recchno
        rs("bank") = recbank
        rs("branch") = recbranch
        rs("inv_status") = "Invoice Due"
        rs("lastpaid") = 0
        If Val(Trim(txtdue.Text)) = 0 Then
            rs("remarks") = "Full Paid"
        Else
            rs("remarks") = "Due"
        End If
        rs.Update
    End If
    lblstatus.Caption = "Service Status : Renewal Successfully Done"
    lblcount.Caption = "So Far Added : " & rs.RecordCount
    Call ClearFields
    Unload Me
    frmrenewal.Show
End If
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
If KeyCode = vbKeyReturn Then cbopayment.SetFocus
End Sub

Private Sub txtpaid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdsave.SetFocus
End Sub
