Attribute VB_Name = "01-BackupModule"
Option Compare Database

Function BackupFile() As Boolean
 Dim Source As String
 Dim Target As String
 Dim retval As Integer
 Source = CurrentDb.Name
 Target = "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\01-Database Master Data\02-Startup Archive\Master_Data_Process_All "
 Target = Target & Format(Date, "yyyy-mm-dd") & " "
 Target = Target & Format(Time, "hh-mm") & ".accdb"
' create the backup
 retval = 0
 Dim objFSO As Object
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 retval = objFSO.CopyFile(Source, Target, True)
 Set objFSO = Nothing
End Function

