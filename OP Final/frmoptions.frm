VERSION 5.00
Begin VB.Form frmoptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill Options"
   ClientHeight    =   705
   ClientLeft      =   3615
   ClientTop       =   4035
   ClientWidth     =   2460
   Icon            =   "frmoptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   2460
   Begin VB.CommandButton cmdback 
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
      Left            =   15
      TabIndex        =   1
      Top             =   375
      Width           =   2430
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print Invoice"
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
      TabIndex        =   0
      Top             =   45
      Width           =   2430
   End
End
Attribute VB_Name = "frmoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database, rs As Recordset

Private Sub cmdback_Click()
Unload frminvoice
Unload Me
Unload frmreport
Unload frminvpmtinfo
selbno = ""
frmmain.Show
End Sub

Private Sub cmdprint_Click()
On Error GoTo printerror

confirm = MsgBox("Sure to print this invoice ?", vbApplicationModal + vbYesNo + vbQuestion, "Confirm Print Invoice")
If confirm = vbNo Then
    cmdback_Click
ElseIf confirm = vbYes Then
    'frminvoice.PrintForm
    Call UpdatePaidAmount
    cmdback_Click
End If

printerror:
If Err.Number = 28663 Then
    MsgBox "No default printer found on your system.Print aborted.", vbApplicationModal + vbCritical, "Print Error"
    cmdback_Click
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\company.mdb")
Set rs = db.OpenRecordset("info", dbOpenTable)
If rs.RecordCount > 0 Then rs.MoveFirst
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized And currentform = "invoice" Then
        frminvoice.WindowState = vbMinimized
        frminvoice.Hide
ElseIf Me.WindowState = vbNormal And currentform = "invoice" Then
        frminvoice.WindowState = vbNormal
        frminvoice.Show
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdback_Click
End Sub

Public Sub UpdatePaidAmount()
Dim paid, due, lpaid As Variant

paid = Val(frminvoice.lblpaid.Caption)
due = Val(frminvoice.lbldue.Caption)
lpaid = Val(frminvoice.lblLastPaid.Caption)

If frminvoice.lblpaid.Caption = "" Then
    MsgBox "ss"
Else
    Set rs = db.OpenRecordset("select *from info where billno=" & "'" & frminvoice.lblbno.Caption & "'")
End If
If rs.EditMode = dbEditNone Then
    rs.Edit
End If
If cmdprint.Caption = "Print Invoice" Then
    rs("paid") = paid + lpaid
    rs("due") = due
    rs("lastpaid") = lpaid + paid
    If due = 0 Then
        rs("remarks") = "Full Paid"
        rs("inv_status") = "Invoice Settled"
    ElseIf due <> 0 Then
        rs("remarks") = "Due"
        rs("inv_status") = "Invoice Due"
    End If
    rs.Update
End If
End Sub
