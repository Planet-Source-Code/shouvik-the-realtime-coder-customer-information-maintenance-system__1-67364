VERSION 5.00
Begin VB.Form frmcashbill 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Receipt Voucher"
   ClientHeight    =   3000
   ClientLeft      =   3030
   ClientTop       =   2955
   ClientWidth     =   7020
   Icon            =   "frmcashbill.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   7020
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
      Left            =   5655
      TabIndex        =   14
      Top             =   2685
      Width           =   1350
   End
   Begin VB.CommandButton cmdsend 
      Caption         =   "Send"
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
      Left            =   15
      TabIndex        =   13
      Top             =   2685
      Width           =   1350
   End
   Begin VB.Label lbltotalvalue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      DataField       =   "coursename"
      DataSource      =   "Data4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2715
      TabIndex        =   12
      Top             =   1155
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Out of total price Rupees................................................."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   11
      Top             =   1125
      Width           =   6870
   End
   Begin VB.Label lbltotalprice 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2000"
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
      Left            =   1275
      TabIndex        =   10
      Top             =   2280
      Width           =   435
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   345
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   2205
      Width           =   2760
   End
   Begin VB.Label lblamount 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2000"
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
      Left            =   6000
      TabIndex        =   9
      Top             =   2280
      Width           =   450
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   345
      Left            =   5595
      Shape           =   4  'Rounded Rectangle
      Top             =   2205
      Width           =   1245
   End
   Begin VB.Label lblbillno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No. "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   165
      TabIndex        =   8
      Top             =   105
      Width           =   675
   End
   Begin VB.Label lbldate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6285
      TabIndex        =   7
      Top             =   105
      Width           =   555
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Received with thanks from................................................"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   75
      TabIndex        =   6
      Top             =   405
      Width           =   6870
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The sum of Rupees........................................................"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   765
      Width           =   6855
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Cash"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   75
      TabIndex        =   4
      Top             =   1830
      Width           =   750
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For purchase of Domain..................................................."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   75
      TabIndex        =   3
      Top             =   1485
      Width           =   6870
   End
   Begin VB.Label lblcname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      DataField       =   "name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2745
      TabIndex        =   2
      Top             =   435
      Width           =   45
   End
   Begin VB.Label lblmoney 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      DataField       =   "coursefee"
      DataSource      =   "Data3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2025
      TabIndex        =   1
      Top             =   795
      Width           =   45
   End
   Begin VB.Label lbldomain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      DataField       =   "coursename"
      DataSource      =   "Data4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2475
      TabIndex        =   0
      Top             =   1515
      Width           =   45
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BorderWidth     =   2
      Height          =   2625
      Left            =   15
      Shape           =   4  'Rounded Rectangle
      Top             =   15
      Width           =   7005
   End
End
Attribute VB_Name = "frmcashbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim billdb As Database, billrs As Recordset

Private Sub cmdback_Click()
Unload Me
Unload frmreport
selbno = ""
currentform = ""
frmmain.Show
End Sub

Private Sub cmdsend_Click()
frmEmail.Show vbModal
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

Set billdb = OpenDatabase(App.Path & "\company.mdb")
Set billrs = billdb.OpenRecordset("info", dbOpenTable)
If billrs.RecordCount > 0 Then billrs.MoveFirst
Call GetBillInfo
End Sub

Public Sub GetBillInfo()
Dim status As String

Set billrs = billdb.OpenRecordset("select *from info where billno='" & selbno & "'")
If billrs.RecordCount > 0 Then
    billrs.MoveFirst
    lblbillno.Caption = "Bill No. " & billrs("billno")
    lbldate.Caption = "Payment Date : " & billrs("date")
    lblcname.Caption = StrConv(billrs("companyname"), vbProperCase)
    lbldomain.Caption = billrs("domain")
    lblamount.Caption = "Rs. " & billrs("paid")
    valnum = Format(billrs("paid"), ".00")
    Call WordConvert
    lblmoney.Caption = valsent
    status = billrs("remarks")
    If status = "Full Paid" Then
        lbltotalprice.Caption = "FULL PAID"
    ElseIf status = "Due" Then
        lbltotalprice.Caption = "DUE - Rs. " & billrs("due")
    End If
    valnum = Format(billrs("price"), ".00")
    Call WordConvert
    lbltotalvalue.Caption = valsent
End If
Set billrs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdback_Click
End Sub
