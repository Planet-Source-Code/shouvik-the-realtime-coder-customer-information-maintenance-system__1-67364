VERSION 5.00
Begin VB.Form frmEmail 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Money Receipt"
   ClientHeight    =   1005
   ClientLeft      =   3540
   ClientTop       =   3360
   ClientWidth     =   5520
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtbody 
      Height          =   3660
      Left            =   540
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   2775
      Visible         =   0   'False
      Width           =   5085
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
      Left            =   4530
      TabIndex        =   3
      Top             =   690
      Width           =   945
   End
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
      Left            =   3555
      TabIndex        =   2
      Top             =   690
      Width           =   945
   End
   Begin VB.TextBox txtsubject 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1275
      TabIndex        =   0
      Top             =   360
      Width           =   4245
   End
   Begin VB.TextBox txtmailid 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1275
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   30
      Width           =   4245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject :"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   420
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   90
      Width           =   300
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database, rs As Recordset, rsd As Recordset

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If Trim(txtsubject.Text) <> "" Then
        Call SaveBody
        ShellExecute Me.hwnd, "", "mailto: " & txtmailid.Text & "?subject=" & txtsubject.Text, "", "", vbNormalFocus
        If currentform = "cashbill" Then
            Unload Me
            Unload frmcashbill
            Unload frmreport
            selbno = ""
            currentform = ""
            frmmain.Show
        ElseIf currentform = "chequebill" Then
            Unload Me
            Unload frmchequebill
            Unload frmreport
            selbno = ""
            currentform = ""
            frmmain.Show
        End If
Else
    MsgBox "Please enter a subject for this mail.", vbApplicationModal + vbExclamation, "No subject mentioned"
    txtsubject.SetFocus
End If
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

Set db = OpenDatabase(App.Path & "\company.mdb")
Set rs = db.OpenRecordset("info", dbOpenTable)
If rs.RecordCount > 0 Then rs.MoveFirst

Set rs = db.OpenRecordset("select *from info where billno=" & "'" & selbno & "'")
If rs.RecordCount > 0 Then
    rs.MoveFirst
    txtmailid.Text = rs("email")
End If
Set rs = Nothing

Call GetMessage
End Sub

Public Sub GetMessage()
On Error Resume Next

Set rsd = db.OpenRecordset("select *from info where billno=" & "'" & selbno & "'")
If rsd.RecordCount > 0 Then
    rsd.MoveFirst
    If currentform = "cashbill" Then
        txtbody.Text = "Receipt VoucherNo. :" & rsd("billno") & "                      Payment Date : " & rsd("date") & vbCrLf _
                                   & vbCrLf & "Received with thanks from " & UCase(frmcashbill.lblcname.Caption) & vbCrLf _
                                   & "The sum of Rupees " & frmcashbill.lblmoney.Caption & vbCrLf _
                                   & "Out of total price Rupees " & frmcashbill.lbltotalvalue.Caption & vbCrLf _
                                   & "For purchase of domain : " & frmcashbill.lbldomain.Caption & vbCrLf _
                                   & "By Cash" & vbCrLf & vbCrLf & "Remarks : " & frmcashbill.lbltotalprice.Caption _
                                   & "                                Total : Rs. " & frmcashbill.lblamount.Caption
    ElseIf currentform = "chequebill" Then
        txtbody.Text = "Receipt VoucherNo. " & rsd("billno") & "                      Payment Date : " & rsd("date") & vbCrLf _
                                   & vbCrLf & "Received with thanks from " & UCase(frmcashbill.lblcname.Caption) & vbCrLf _
                                   & "The sum of Rupees " & frmcashbill.lblmoney.Caption & vbCrLf _
                                   & "Out of total price Rupees " & frmcashbill.lbltotalvalue.Caption & vbCrLf _
                                   & "For purchase of domain : " & frmcashbill.lbldomain.Caption & vbCrLf _
                                   & "By Cheque No. :" & frmchequebill.lblpayment.Caption _
                                   & "                          Drawn on : " & frmchequebill.lblbankbranch.Caption & vbCrLf & vbCrLf _
                                   & "Remarks : " & frmchequebill.lbltotalprice.Caption _
                                   & "                                Total : Rs. " & frmchequebill.lblamount.Caption
    End If
End If
Set rsd = Nothing
End Sub

Private Sub txtsubject_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub txtsubject_LostFocus()
txtsubject.Text = StrConv(Trim(txtsubject.Text), vbProperCase)
End Sub

Public Sub SaveBody()
Open App.Path & "\Message.txt" For Output As #1
    Print #1, Trim(txtbody.Text)
Close #1
MsgBox "The body of the message has been saved under the folder " & Chr(34) & App.Path & Chr(34) _
        & "with name " & Chr(34) & "Message.txt" & Chr(34) & ".Attach this file to this mail.", vbApplicationModal + vbInformation, "Mail"
End Sub
