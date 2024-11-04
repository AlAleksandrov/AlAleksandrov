VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmd_login_Click()

Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim strSQL As String

If Trim(Me.txt_username.Value & vbNullString) = vbNullString Then
MsgBox prompt:="Username should not be left blank.", buttons:=vbInformation, title:="Username Required"
Me.txt_username.SetFocus
Exit Sub
End If

If Trim(Me.txt_password.Value & vbNullString) = vbNullString Then
MsgBox prompt:="Password should not be left blank.", buttons:=vbInformation, title:="Password Required"
Me.txt_password.SetFocus
Exit Sub
End If

' Hash the password
Dim hashedPassword As String
hashedPassword = HashPassword(Me.txt_password.Value)

'Query to check if login details are correct
strSQL = "SELECT FirstName FROM tbl_users WHERE Username = """ & Me.txt_username.Value & """ AND Password = """ & hashedPassword & """"

Set db = CurrentDb
Set rst = db.OpenRecordset(strSQL)
If rst.EOF Then
MsgBox prompt:="Incorrect username/password. Try again.", buttons:=vbCritical, title:="Login Error"
Me.txt_username.SetFocus
Else
MsgBox prompt:="Hello, " & rst.Fields(0).Value & ".", buttons:=vbOKOnly, title:="Login Successful"
DoCmd.Close acForm, "frm_login", acSaveYes
End If

Set db = Nothing
Set rst = Nothing

End Sub

Private Sub cmd_cancel_Click()
   DoCmd.Quit acQuitSaveAll
End Sub

'Function HashPassword(Password As String) As String
'    'Hashes the password using SHA-256 algorithm
'    Dim shaObj As Object
'    Set shaObj = CreateObject("System.Security.Cryptography.SHA256Managed")
'    Dim bytes() As Byte
'    bytes = StrConv(Password, vbFromUnicode)
'    bytes = shaObj.ComputeHash_2((bytes))
'    HashPassword = StrConv(bytes, vbUnicode)
'    Set shaObj = Nothing
'End Function

Function HashPassword(password As String) As String
    'Hashes the password using MD5 algorithm
    Dim md5Obj As Object
    Set md5Obj = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    Dim bytes() As Byte
    bytes = StrConv(password, vbFromUnicode)
    bytes = md5Obj.ComputeHash_2((bytes))
    Dim hash As String
    Dim i As Long
    For i = LBound(bytes) To UBound(bytes)
        hash = hash & Right("0" & Hex(bytes(i)), 2)
    Next i
    HashPassword = Left(hash, 16)
    Set md5Obj = Nothing
End Function
