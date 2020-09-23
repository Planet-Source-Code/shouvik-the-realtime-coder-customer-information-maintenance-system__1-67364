Attribute VB_Name = "modmoney"
Option Explicit

Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Public selbno, chequeno, bankname, branchname, currentform As String
Public recchno, recbank, recbranch As String
Public confirm As Integer
Public valnum, valsent As String
Private n, intpart, realpart, numchar, intword, realword, spltval, spltword As String
Private flag As Boolean

Public Sub WordConvert()
n = ""
intpart = ""
realpart = ""
numchar = ""
intword = ""
realword = ""
spltval = ""
spltword = ""
valsent = ""
If valnum = "." Then valnum = "0.00"
If valnum = "" Then Exit Sub

intpart = Format(Int(valnum), "000000000")
realpart = Right(valnum, 2)

spltval = realpart
Call ValFind
If spltword <> "" Then realword = spltword
spltval = Mid(intpart, 1, 2)
Call ValFind
If spltword <> "" Then intword = spltword + "Crore "
spltval = Mid(intpart, 3, 2)
Call ValFind
If spltword <> "" Then intword = intword + spltword + "Lakh "
spltval = Mid(intpart, 5, 2)
Call ValFind
If spltword <> "" Then intword = intword + spltword + "Thousand "
n = Mid(intpart, 7, 1)
Call ONES
If numchar <> "" Then intword = intword + numchar + "Hundred "
spltval = Mid(intpart, 8, 2)
If intword <> "" And Val(spltval) > 0 And realword = "" Then intword = intword + "AND "
Call ValFind
If spltword <> "" Then intword = intword + spltword
If intword <> "" And realword <> "" Then valsent = intword + " AND Paise " + realword + "Only"
If intword <> "" And realword = "" Then valsent = intword + "Only"
If intword = "" And realword <> "" Then valsent = "Paise: " + realword + "Only"
End Sub

Private Sub ValFind()
n = ""
spltword = ""
If Val(spltval) = 0 Then Exit Sub
n = Left(spltval, 1)
Call TENS
spltword = numchar
If flag = False Then n = Right(spltval, 1): Call ONES: spltword = spltword + numchar
End Sub

Private Sub ONES()
numchar = ""
If n = 0 Then numchar = ""
If n = 1 Then numchar = "One "
If n = 2 Then numchar = "Two "
If n = 3 Then numchar = "Three "
If n = 4 Then numchar = "Four "
If n = 5 Then numchar = "Five "
If n = 6 Then numchar = "Six "
If n = 7 Then numchar = "Seven "
If n = 8 Then numchar = "Eight "
If n = 9 Then numchar = "Nine "
End Sub

Private Sub TENS()
numchar = ""
If n = 1 Then n = Right(spltval, 1): Call TEENS: flag = True: Exit Sub Else flag = False
If n = 0 Then numchar = ""
If n = 2 Then numchar = "Twenty "
If n = 3 Then numchar = "Thirty "
If n = 4 Then numchar = "Fourty "
If n = 5 Then numchar = "Fifty "
If n = 6 Then numchar = "Sixty "
If n = 7 Then numchar = "Seventy "
If n = 8 Then numchar = "Eighty "
If n = 9 Then numchar = "Ninety "
End Sub

Private Sub TEENS()
numchar = ""
If n = 0 Then numchar = "Ten "
If n = 1 Then numchar = "Eleven "
If n = 2 Then numchar = "Twelve "
If n = 3 Then numchar = "Thirteen "
If n = 4 Then numchar = "Fourteen "
If n = 5 Then numchar = "Fifteen "
If n = 6 Then numchar = "Sixteen "
If n = 7 Then numchar = "Seventeen "
If n = 8 Then numchar = "Eighteen "
If n = 9 Then numchar = "Nineten "
End Sub

Public Function GetMachineUserName() As String
Dim UserName As String * 255

Call GetUserName(UserName, 255)
GetMachineUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
