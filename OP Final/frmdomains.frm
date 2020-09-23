VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmdomains 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Company Domain"
   ClientHeight    =   3810
   ClientLeft      =   3300
   ClientTop       =   2490
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmdomains.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5205
   Begin MSComctlLib.ImageList domainimage 
      Left            =   8880
      Top             =   2805
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
            Picture         =   "frmdomains.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvdomain 
      Height          =   3180
      Left            =   15
      TabIndex        =   5
      Top             =   240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5609
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtalldomains 
      Height          =   1290
      Left            =   8160
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   1245
      Visible         =   0   'False
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
      Left            =   4230
      TabIndex        =   0
      Top             =   3480
      Width           =   945
   End
   Begin VB.ListBox lstdomain 
      Height          =   1230
      Left            =   9270
      TabIndex        =   2
      Top             =   1215
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Shape Shape1 
      Height          =   3210
      Left            =   0
      Top             =   225
      Width           =   5205
   End
   Begin VB.Label lblcount 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Available Domain(s) :"
      Height          =   195
      Left            =   3120
      TabIndex        =   3
      Top             =   15
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Available Domain(s) :"
      Height          =   195
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   2055
   End
End
Attribute VB_Name = "frmdomains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database, rs As Recordset, rsd As Recordset
Dim curdomain, alldomains As Variant, pos As Integer
Dim colx As ColumnHeader
Dim lstitem As ListItem

Private Sub cmdback_Click()
Unload Me
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

Set db = OpenDatabase(App.Path & "\company.mdb")
Set rs = db.OpenRecordset("info", dbOpenTable)
If rs.RecordCount > 0 Then rs.MoveFirst

frmsearch.datasearch.DatabaseName = App.Path & "\company.mdb"
Call AddAfterCheck
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Public Sub AddAfterCheck()
Set rsd = db.OpenRecordset("info")
If rsd.RecordCount > 0 Then rsd.MoveFirst

Set colx = lvdomain.ColumnHeaders.Add(, , "Company Domain", lvdomain.Width)
lvdomain.View = lvwReport
lvdomain.Icons = domainimage
lvdomain.SmallIcons = domainimage
Do Until rsd.EOF
    alldomains = UCase(Trim(txtalldomains.Text))
    curdomain = UCase(Trim(rsd("domain")))
    pos = InStr(alldomains, curdomain)
    If pos = 0 Then
        txtalldomains.SelText = rsd("domain") & vbCrLf
        lstdomain.AddItem rsd("domain")
        lblcount.Caption = "Found : " & lstdomain.ListCount
        Set lstitem = lvdomain.ListItems.Add(, rsd!domain, (rsd!domain))
        lstitem.Icon = 1
        lstitem.SmallIcon = 1
        rsd.MoveNext
    ElseIf pos > 0 Then
        txtalldomains.SelText = rsd("domain") & vbCrLf
        rsd.MoveNext
    End If
Loop
End Sub

Private Sub lvdomain_Click()
Dim fields, recd As String

fields = "domain"
recd = lvdomain.SelectedItem.Text
frmsearch.datasearch.RecordSource = "select *from info where " & fields & "='" & recd & "'"
frmsearch.datasearch.Refresh
frmsearch.txtresult.Text = "Total " & frmsearch.GridInfo.ApproxCount & " Record(s) Found For Domain=" & Chr(34) & recd & Chr(34)
frmsearch.Caption = "Search & Query - Last Query Searched By Domain"
fields = ""
recd = ""
Unload Me
End Sub
