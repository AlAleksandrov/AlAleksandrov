VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmd_cancel_Click()
   DoCmd.Close acForm, "frm_register", acSaveYes
End Sub

Private Sub cmd_register_Click()

Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim strSQL As String
Dim computerName As String

computerName = Environ("COMPUTERNAME")

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

' Check if the admin code is correct
If Trim(Me.txt_admincode.Value & vbNullString) <> "REG01" Then
    MsgBox prompt:="Please ask for Admin registration approval.", buttons:=vbCritical, title:="Registration Error"
    Me.txt_admincode.SetFocus
    Exit Sub
End If

' Check if the username already exists in the database
strSQL = "SELECT COUNT(*) FROM tbl_users WHERE Username = """ & Me.txt_username.Value & """"
Set db = CurrentDb
Set rst = db.OpenRecordset(strSQL)
If rst.Fields(0).Value > 0 Then
    MsgBox prompt:="Username already exists. Please choose a different username.", buttons:=vbCritical, title:="Registration Error"
    Me.txt_username.SetFocus
    Exit Sub
End If

' Hash the password and insert the user into the database
strSQL = "INSERT INTO tbl_users (Username, Password, FirstName, LastName, LeoniNo,HostMachine) VALUES(""" & Me.txt_username.Value & """, """ & HashPassword(Me.txt_password.Value) & """, """ & Me.txt_firstname.Value & """, """ & Me.txt_lastname.Value & """, """ & Me.txt_leoniNo.Value & """,""" & computerName & """)"
db.Execute strSQL, dbFailOnError

' Show success message and close the form
MsgBox prompt:="Registration successful. You can now log in with your new account.", buttons:=vbOKOnly, title:="Registration Successful"
DoCmd.Close acForm, "frm_register", acSaveYes

Set db = Nothing
Set rst = Nothing

End Sub

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

'Function HashPassword(Password As String) As String
''Hashes the password using SHA-256 algorithm
'Dim shaObj As Object
'Set shaObj = CreateObject("System.Security.Cryptography.SHA256Managed")
'Dim bytes() As Byte
'bytes = StrConv(Password, vbFromUnicode)
'bytes = shaObj.ComputeHash_2((bytes))
'
'Dim hash As String
'Dim i As Long
'For i = LBound(bytes) To UBound(bytes)
'    hash = hash & Right("0" & Hex(bytes(i)), 2)
'Next i
'HashPassword = hash
'
'Set shaObj = Nothing

'End Function

