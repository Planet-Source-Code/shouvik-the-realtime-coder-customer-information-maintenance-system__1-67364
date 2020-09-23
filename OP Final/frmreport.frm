VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmreport 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Payment Report"
   ClientHeight    =   6105
   ClientLeft      =   1710
   ClientTop       =   1770
   ClientWidth     =   9900
   Icon            =   "frmreport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9900
   Begin VB.TextBox txtinvoice 
      Height          =   330
      Left            =   8010
      TabIndex        =   8
      Top             =   7230
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton cmdmoneyreceipt 
      Caption         =   "Money Receipt"
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
      Left            =   1545
      TabIndex        =   6
      ToolTipText     =   "Displays individual customer bill"
      Top             =   5775
      Width           =   1500
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "Invoice"
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
      Left            =   15
      TabIndex        =   5
      ToolTipText     =   "Displays individual customer bill"
      Top             =   5775
      Width           =   1500
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Bindings        =   "frmreport.frx":0442
      Height          =   1095
      Left            =   4605
      TabIndex        =   4
      Top             =   6840
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   1931
      _Version        =   393216
   End
   Begin VB.Data datapaidamount 
      Caption         =   "datapaidamount"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   2970
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7170
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data datatotalpaid 
      Caption         =   "datatotalpaid"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3030
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7530
      Visible         =   0   'False
      Width           =   1470
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmreport.frx":045F
      Height          =   570
      Left            =   1845
      OleObjectBlob   =   "frmreport.frx":047B
      TabIndex        =   3
      Top             =   7200
      Visible         =   0   'False
      Width           =   1005
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
      Height          =   300
      Left            =   8400
      TabIndex        =   2
      ToolTipText     =   "Closes report and backs to main screen"
      Top             =   5775
      Width           =   1500
   End
   Begin MSComctlLib.ImageList reportimage 
      Left            =   7290
      Top             =   7140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreport.frx":2222
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvreport 
      Height          =   5385
      Left            =   15
      TabIndex        =   7
      Top             =   15
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   9499
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Shape Shape1 
      Height          =   5415
      Left            =   0
      Top             =   0
      Width           =   9900
   End
   Begin VB.Label lblamount 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9285
      TabIndex        =   1
      Top             =   5490
      Width           =   585
   End
   Begin VB.Label lbltotalpaid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   5490
      Width           =   585
   End
   Begin VB.Shape Shape2 
      Height          =   285
      Left            =   0
      Top             =   5445
      Width           =   9900
   End
End
Attribute VB_Name = "frmreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim custdb As Database, custrs As Recordset, custrsd As Recordset, rsd As Recordset
Dim totalpaid As Integer, totalrecords As Integer
Dim pmode, status As String
Dim colx As ColumnHeader
Dim lstitem As ListItem

Private Sub cmdclose_Click()
selbno = ""
pmode = ""
status = ""
txtinvoice.Text = ""
Unload Me
frmmain.Show
End Sub

Private Sub cmdmoneyreceipt_Click()
If selbno = "" Then
    MsgBox "Select a billno from the grid to see the money receipt.", vbApplicationModal + vbExclamation, "No record selected"
ElseIf selbno <> "" Then
    If pmode = "CASH" And status = "Full Paid" Or status = "Due" Then
        If Trim(txtinvoice.Text) = "Invoice Due" Then
            currentform = "cashbill"
            Me.Hide
            frmcashbill.Show
        ElseIf Trim(txtinvoice.Text) = "Invoice Settled" Then
            MsgBox "MoneyReceipt to BillNo. " & selbno & " has been settled.", vbApplicationModal + vbInformation, "Money Receipt"
        End If
    ElseIf status = "" Then
        MsgBox "No payment has not been made yet against Bill-No." & selbno, vbApplicationModal + vbInformation, "Report"
        selbno = ""
        Exit Sub
    ElseIf pmode = "CHEQUE" And status = "Full Paid" Or status = "Due" Then
        If Trim(txtinvoice.Text) = "Invoice Due" Then
            currentform = "chequebill"
            Me.Hide
            frmchequebill.Show
        ElseIf Trim(txtinvoice.Text) = "Invoice Settled" Then
            MsgBox "MoneyReceipt to BillNo. " & selbno & " has been settled.", vbApplicationModal + vbInformation, "Invoice"
        End If
    ElseIf status = "" Then
        MsgBox "No payment has not been made yet against Bill-No." & selbno, vbApplicationModal + vbInformation, "Report"
        selbno = ""
        Exit Sub
    End If
End If
End Sub

Private Sub cmdshow_Click()
If selbno = "" Then
    MsgBox "Select a billno from the grid to see the invoice.", vbApplicationModal + vbExclamation, "No record selected"
ElseIf selbno <> "" Then
    If status = "Full Paid" Or status = "Due" Then
        If Trim(txtinvoice.Text) = "Invoice Due" Then
            Me.Hide
            frminvpmtinfo.Show
        ElseIf Trim(txtinvoice.Text) = "Invoice Settled" Then
            MsgBox "Invoice to BillNo. " & selbno & " has been settled.", vbApplicationModal + vbInformation, "Invoice"
        End If
    ElseIf status = "" Then
        MsgBox "No payment has not been made yet against Bill-No." & selbno, vbApplicationModal + vbInformation, "Report"
        selbno = ""
    End If
End If
End Sub

Private Sub Form_Activate()
totalpaid = DBGrid2.ApproxCount
lbltotalpaid.Caption = "Total Paid :  " & DBGrid2.ApproxCount & " of " & custrs.RecordCount
If Grid1.TextMatrix(1, 1) = "" Then
    lblamount.Caption = "No Amount Has Not Been Full Paid Yet"
ElseIf Grid1.TextMatrix(1, 1) <> "" Then
    lblamount.Caption = "Total Amount Full Paid : Rs. " & Grid1.TextMatrix(1, 1)
End If
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
Set custdb = OpenDatabase(App.Path & "\company.mdb")
Set custrs = custdb.OpenRecordset("info", dbOpenTable)
Set custrsd = custdb.OpenRecordset("info", dbOpenTable)
Set rsd = custdb.OpenRecordset("info", dbOpenTable)
If custrs.RecordCount > 0 Then custrs.MoveFirst
If custrsd.RecordCount > 0 Then custrsd.MoveFirst
If rsd.RecordCount > 0 Then rsd.MoveFirst

totalrecords = custrs.RecordCount
datatotalpaid.DatabaseName = App.Path & "\company.mdb"
datatotalpaid.RecordSource = "select *from info where remarks='" & "Full Paid" & "'"
datapaidamount.DatabaseName = App.Path & "\company.mdb"
datapaidamount.RecordSource = "select sum(paid) from info where remarks='" & "Full Paid" & "'"

Call GetData
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdclose_Click
End Sub

Public Sub GetData()
On Error Resume Next

Set colx = lvreport.ColumnHeaders.Add(, , "Bill No", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "Date", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "Company Name", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "Address", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "Description", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "Domain", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "Email", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "Start Campaigning", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "End Campaigning", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "Payment Mode", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "Domain Value", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "Amount Paid", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "Amount Due", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "Remarks", lvreport.Width / 4)
Set colx = lvreport.ColumnHeaders.Add(, , "Invoice Status", lvreport.Width / 4)
lvreport.View = lvwReport
lvreport.Icons = reportimage
lvreport.SmallIcons = reportimage
If custrsd.RecordCount > 0 Then custrsd.MoveFirst
While Not custrsd.EOF
    Set lstitem = lvreport.ListItems.Add(, custrsd!billno, (custrsd!billno))
        lstitem.Icon = 1
        lstitem.SmallIcon = 1
        lstitem.SubItems(1) = CStr(custrsd!Date)
        lstitem.SubItems(2) = CStr(custrsd!CompanyName)
        lstitem.SubItems(3) = CStr(custrsd!address)
        lstitem.SubItems(4) = CStr(custrsd!Description)
        lstitem.SubItems(5) = CStr(custrsd!domain)
        lstitem.SubItems(6) = CStr(custrsd!email)
        lstitem.SubItems(7) = CStr(custrsd!startcamp)
        lstitem.SubItems(8) = CStr(custrsd!endcamp)
        lstitem.SubItems(9) = CStr(custrsd!payment)
        lstitem.SubItems(10) = CStr(custrsd!price)
        lstitem.SubItems(11) = CStr(custrsd!paid)
        lstitem.SubItems(12) = CStr(custrsd!due)
        lstitem.SubItems(13) = CStr(custrsd!remarks)
        lstitem.SubItems(14) = CStr(custrsd!inv_status)
        custrsd.MoveNext
Wend
End Sub

Private Sub lvreport_ItemClick(ByVal Item As MSComctlLib.ListItem)
selbno = Item.Key
status = Item.SubItems(13)
pmode = Item.SubItems(9)
txtinvoice.Text = Item.SubItems(14)
End Sub
