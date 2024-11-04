VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MAIN MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Command121_Click()

Call BackupFile

End Sub

Private Sub Form_Current()

End Sub


'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-----------------Separator-------------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Private Sub OpenFile_CM_Click()

Dim OpenExplorer As String
    OpenExplorer = "\\SVBG1FILE01\Groups\AO\01-Change managment\" ' Replace with the folder path you want to open
    Call OpenFileExplorer(OpenExplorer)
End Sub

Private Sub OpenFile_DBMS_Click()

Dim OpenExplorer As String
    OpenExplorer = "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\01-Database Master Data\" ' Replace with the folder path you want to open
    Call OpenFileExplorer(OpenExplorer)
End Sub

Private Sub OpenFile_DImp_Click()

Dim OpenExplorer As String
    OpenExplorer = "\\Svbg1pj002\dataimport$\" ' Replace with the folder path you want to open
    Call OpenFileExplorer(OpenExplorer)
End Sub

Private Sub OpenFile_E3_Click()

Dim OpenExplorer As String
    OpenExplorer = "C:\E3Compare\" ' Replace with the folder path you want to open
    Call OpenFileExplorer(OpenExplorer)
End Sub

Private Sub OpenFile_KBL_Click()

Dim OpenExplorer As String
    OpenExplorer = "\\Desktop\KBL\" ' Replace with the folder path you want to open
    Call OpenFileExplorer(OpenExplorer)
End Sub

Private Sub OpenFile_MD_Click()

Dim OpenExplorer As String
    OpenExplorer = "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\" ' Replace with the folder path you want to open
    Call OpenFileExplorer(OpenExplorer)
End Sub

Private Sub OpenFile_SP_Click()

Dim OpenExplorer As String
    OpenExplorer = "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\Spare parts\" ' Replace with the folder path you want to open
    Call OpenFileExplorer(OpenExplorer)
End Sub

Private Sub OpenFile_TPH_Click()

Dim OpenExplorer As String
    OpenExplorer = "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\TPH\" ' Replace with the folder path you want to open
    Call OpenFileExplorer(OpenExplorer)
End Sub

Private Sub OpenFile_W206_Click()

Dim OpenExplorer As String
    OpenExplorer = "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\BR206 - C class\W206 & S206\" ' Replace with the folder path you want to open
    Call OpenFileExplorer(OpenExplorer)
End Sub

Private Sub OpenFile_X254_Click()

Dim OpenExplorer As String
    OpenExplorer = "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\BR206 - C class\X254\" ' Replace with the folder path you want to open
    Call OpenFileExplorer(OpenExplorer)
End Sub
Private Sub OpenFile_EVA2_Click()

Dim OpenExplorer As String
    OpenExplorer = "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\BR297 - EVA2\V297 & V295\" ' Replace with the folder path you want to open
    Call OpenFileExplorer(OpenExplorer)
End Sub
Private Sub StartFile_E3_Click()

Dim appPath As String, directoryPath As String
    appPath = "C:\E3Start\E3Start.bat"    'Replace with the folder path you want to open
    directoryPath = "C:\E3Start\"    'Replace with the folder path you want to open
    Call OpenApplicationInDirectory(appPath, directoryPath)
    
End Sub

Sub OpenFileExplorer(Optional OpenExplorer As String)

    Dim objShell As Object
    Set objShell = CreateObject("Shell.Application")
    objShell.Open (OpenExplorer)   ' Replace with the directory you want to open
    Set objShell = Nothing
    
End Sub

Sub OpenApplicationInDirectory(appPath As String, directoryPath As String)
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    objShell.CurrentDirectory = directoryPath
    objShell.Run appPath
    Set objShell = Nothing
End Sub

Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-----------------Separator-------------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Private Sub Form_Load()
    'Log user login time
    LogAction "User Login", "User " & Environ("USERNAME") & " logged in at " & Now()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Log user logout time
    LogAction "User Logout", "User " & Environ("USERNAME") & " logged out at " & Now()
End Sub

Private Sub Form_Timer()
    'Log user activity every minute
    LogAction "User Activity", "User " & Environ("USERNAME") & " is active at " & Now()
End Sub

Private Sub LogAction(ActionType As String, ActionDescription As String)
    'Log user action to a table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("tblUserActions", dbOpenDynaset)
    
    rs.AddNew
    rs!username = Environ("USERNAME")
    rs!ActionType = ActionType
    rs!ActionDescription = ActionDescription
    rs!ActionTime = Now()
    rs.Update
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

