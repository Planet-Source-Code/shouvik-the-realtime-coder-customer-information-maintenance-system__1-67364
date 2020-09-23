VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmendcampdates 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select EndCamp Date"
   ClientHeight    =   3810
   ClientLeft      =   2925
   ClientTop       =   2940
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmendcampdates.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5250
   Begin VB.ListBox lstendcampdate 
      Height          =   1230
      Left            =   9285
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   1305
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
      Left            =   4290
      TabIndex        =   0
      Top             =   3480
      Width           =   945
   End
   Begin VB.TextBox txtallendcampdates 
      Height          =   1290
      Left            =   8175
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1230
      Visible         =   0   'False
      Width           =   945
   End
   Begin MSComctlLib.ImageList endcampimage 
      Left            =   8895
      Top             =   2790
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
            Picture         =   "frmendcampdates.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvendcampdate 
      Height          =   3180
      Left            =   15
      TabIndex        =   5
      Top             =   240
      Width           =   5220
      _ExtentX        =   9208
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
   Begin VB.Shape Shape1 
      Height          =   3210
      Left            =   0
      Top             =   225
      Width           =   5250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Available EndCamp Date(s) :"
      Height          =   195
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   2640
   End
   Begin VB.Label lblcount 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Available Domain(s) :"
      Height          =   195
      Left            =   3135
      TabIndex        =   3
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmendcampdates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database, rs As Recordset, rsd As Recordset
Dim curendcamp, allendcamps As Variant, pos As Integer
Dim colx As ColumnHeader
Dim lstitem As ListItem
Dim fields As String, recd, curdate As Variant

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

Set colx = lvendcampdate.ColumnHeaders.Add(, , "End Campaigning Date", lvendcampdate.Width)
lvendcampdate.View = lvwReport
lvendcampdate.Icons = endcampimage
lvendcampdate.SmallIcons = endcampimage
Do Until rsd.EOF
    allendcamps = UCase(Trim(txtallendcampdates.Text))
    curendcamp = UCase(Trim(rsd("endcamp")))
    pos = InStr(allendcamps, curendcamp)
    If pos = 0 Then
        txtallendcampdates.SelText = rsd("endcamp") & vbCrLf
        lstendcampdate.AddItem rsd("endcamp")
        lblcount.Caption = "Found : " & lstendcampdate.ListCount
        Set lstitem = lvendcampdate.ListItems.Add(, rsd!endcamp, (rsd!endcamp))
        lstitem.Icon = 1
        lstitem.SmallIcon = 1
        rsd.MoveNext
    ElseIf pos > 0 Then
        txtallendcampdates.SelText = rsd("endcamp") & vbCrLf
        rsd.MoveNext
    End If
Loop
End Sub

Private Sub lvendcampdate_Click()
curdate = Format(Date, "dd/mm/yyyy")
recd = lvendcampdate.SelectedItem.Text

If recd = curdate Then
    frmsearch.datasearch.RecordSource = "select *from info where endcamp=" & "'" & recd & "'"
    frmsearch.datasearch.Refresh
    frmsearch.txtresult.Text = "Total " & frmsearch.GridInfo.ApproxCount & " Record(s) Found For EndCampDate=" & Chr(34) & recd & Chr(34)
    frmsearch.Caption = "Search & Query - Last Query Searched By End Campaigning"
    fields = ""
    recd = ""
    curdate = ""
    Unload Me
ElseIf recd < curdate Then
    frmsearch.datasearch.RecordSource = "select *from info where endcamp>=" & "'" & recd & "'" & " and endcamp<=" & "'" & curdate & "'"
    frmsearch.datasearch.Refresh
    frmsearch.txtresult.Text = "Total " & frmsearch.GridInfo.ApproxCount & " Record(s) Found For EndCampDate=" & Chr(34) & recd & Chr(34)
    frmsearch.Caption = "Search & Query - Last Query Searched By End Campaigning"
    fields = ""
    recd = ""
    curdate = ""
    Unload Me
ElseIf recd > curdate Then
    frmsearch.datasearch.RecordSource = "select *from info where endcamp>=" & "'" & curdate & "'" & " and endcamp<=" & "'" & recd & "'"
    frmsearch.datasearch.Refresh
    frmsearch.txtresult.Text = "Total " & frmsearch.GridInfo.ApproxCount & " Record(s) Found For EndCampDate=" & Chr(34) & recd & Chr(34)
    frmsearch.Caption = "Search & Query - Last Query Searched By End Campaigning"
    fields = ""
    recd = ""
    curdate = ""
    Unload Me
End If
End Sub
