VERSION 5.00
Begin VB.Form frminvoice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   7605
   ClientLeft      =   2655
   ClientTop       =   915
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame9 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   4770
      TabIndex        =   19
      Top             =   6360
      Width           =   1785
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   4755
      TabIndex        =   18
      Top             =   7020
      Width           =   30
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   2565
      TabIndex        =   17
      Top             =   7020
      Width           =   30
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5205
      Left            =   4755
      TabIndex        =   16
      Top             =   1515
      Width           =   30
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   2565
      TabIndex        =   9
      Top             =   1515
      Width           =   30
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   30
      TabIndex        =   8
      Top             =   7230
      Width           =   6525
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   30
      TabIndex        =   7
      Top             =   1785
      Width           =   6525
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   4590
      TabIndex        =   1
      Top             =   705
      Width           =   1965
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   4560
      TabIndex        =   0
      Top             =   30
      Width           =   30
   End
   Begin VB.Label lblLastPaid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hh"
      Height          =   195
      Left            =   915
      TabIndex        =   37
      Top             =   5280
      Width           =   210
   End
   Begin VB.Label lbldue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hh"
      Height          =   195
      Left            =   900
      TabIndex        =   36
      Top             =   4995
      Width           =   210
   End
   Begin VB.Label lbladdress 
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      Height          =   915
      Left            =   855
      TabIndex        =   35
      Top             =   315
      Width           =   3690
   End
   Begin VB.Label lblmoney 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      Height          =   195
      Left            =   1695
      TabIndex        =   34
      Top             =   6735
      Width           =   660
   End
   Begin VB.Label lblpaid 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label25"
      Height          =   195
      Left            =   5355
      TabIndex        =   33
      Top             =   7305
      Width           =   660
   End
   Begin VB.Label lblcno 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label24"
      Height          =   195
      Left            =   3330
      TabIndex        =   32
      Top             =   7305
      Width           =   660
   End
   Begin VB.Label lblpayment 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label23"
      Height          =   195
      Left            =   900
      TabIndex        =   31
      Top             =   7305
      Width           =   660
   End
   Begin VB.Label lbltotal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label22"
      Height          =   195
      Left            =   5340
      TabIndex        =   30
      Top             =   6435
      Width           =   660
   End
   Begin VB.Label lblprice 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label21"
      Height          =   195
      Left            =   5340
      TabIndex        =   29
      Top             =   1860
      Width           =   660
   End
   Begin VB.Label lblperiod 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      Height          =   195
      Left            =   3345
      TabIndex        =   28
      Top             =   1860
      Width           =   660
   End
   Begin VB.Label lbldescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Label19"
      Height          =   2625
      Left            =   60
      TabIndex        =   27
      Top             =   3645
      Width           =   2475
   End
   Begin VB.Label lbldomain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      Height          =   195
      Left            =   60
      TabIndex        =   26
      Top             =   1860
      Width           =   660
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label17"
      Height          =   195
      Left            =   5265
      TabIndex        =   25
      Top             =   1095
      Width           =   660
   End
   Begin VB.Label lblbno 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label16"
      Height          =   195
      Left            =   5265
      TabIndex        =   24
      Top             =   330
      Width           =   660
   End
   Begin VB.Label lblcompanyname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      Height          =   195
      Left            =   1485
      TabIndex        =   23
      Top             =   45
      Width           =   660
   End
   Begin VB.Label lblemail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      Height          =   195
      Left            =   720
      TabIndex        =   22
      Top             =   1230
      Width           =   660
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   4305
      TabIndex        =   21
      Top             =   6435
      Width           =   435
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount(in words) :"
      Height          =   195
      Left            =   60
      TabIndex        =   20
      Top             =   6735
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount(Rs.)"
      Height          =   195
      Index           =   1
      Left            =   5160
      TabIndex        =   15
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount(Rs.)"
      Height          =   195
      Left            =   5145
      TabIndex        =   14
      Top             =   7020
      Width           =   1080
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment"
      Height          =   195
      Left            =   855
      TabIndex        =   13
      Top             =   7005
      Width           =   765
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Cheque No."
      Height          =   195
      Left            =   2985
      TabIndex        =   12
      Top             =   7020
      Width           =   1425
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Period"
      Height          =   195
      Left            =   3360
      TabIndex        =   11
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domain  +  Description"
      Height          =   195
      Left            =   300
      TabIndex        =   10
      Top             =   1560
      Width           =   1920
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   585
      Left            =   15
      Top             =   7005
      Width           =   6570
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   5190
      Left            =   15
      Top             =   1530
      Width           =   6570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail :"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   1230
      Width           =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   315
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name :"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   45
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date."
      Height          =   195
      Index           =   0
      Left            =   4620
      TabIndex        =   3
      Top             =   750
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
      Height          =   195
      Left            =   4620
      TabIndex        =   2
      Top             =   45
      Width           =   540
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1470
      Left            =   15
      Top             =   15
      Width           =   6570
   End
End
Attribute VB_Name = "frminvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE As Long = -20
Private Const LWA_COLORKEY As Long = &H1
Private Const LWA_ALPHA As Long = &H2
Private Const WS_EX_LAYERED As Long = &H80000

Dim lOldStyle As Long
Dim invdb As Database, invrs As Recordset

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

lOldStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
SetWindowLong hwnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
SetLayeredWindowAttributes hwnd, vbWhite, 200, LWA_ALPHA

Set invdb = OpenDatabase(App.Path & "\company.mdb")
Set invrs = invdb.OpenRecordset("info", dbOpenTable)
If invrs.RecordCount > 0 Then invrs.MoveFirst
Call GetInvoiceInfo
End Sub

Public Sub GetInvoiceInfo()
On Error Resume Next

Set invrs = invdb.OpenRecordset("select *from info where billno=" & "'" & selbno & "'")
If invrs.RecordCount > 0 Then
    invrs.MoveFirst
    lblbno.Caption = invrs("billno")
    lbldate.Caption = invrs("date")
    lblcompanyname.Caption = invrs("companyname")
    lbladdress.Caption = invrs("address")
    lbldomain.Caption = invrs("domain")
    lblemail.Caption = invrs("email")
    lblDescription.Caption = invrs("description")
    lblperiod.Caption = invrs("startcamp") & Chr(10) & "To" & Chr(10) & invrs("endcamp")
    lblprice.Caption = invrs("price")
    lbltotal.Caption = invrs("price")
    lblpayment.Caption = invrs("payment")
    If lblpayment.Caption = "CASH" Then
        lblcno.Caption = "-----"
    ElseIf lblpayment.Caption = "CHEQUE" Then
        lblcno.Caption = invrs("ch_no")
    End If
    lblpaid.Caption = Trim(frminvpmtinfo.txtpaid.Text)
    valnum = Format(invrs("price"), ".00")
    Call WordConvert
    lblmoney.Caption = valsent
End If
Set invrs = Nothing
End Sub
