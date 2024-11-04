Attribute VB_Name = "08c-Check_PRG_Xcode"
Option Compare Database

Function Check_PRG_Xcode()

Dim x As Integer, y As Long, z As Integer, z1 As Integer, z2 As Integer, last_row As Integer, MyLinkDest As String, last_row1 As Integer, MyRange As Integer, OpenAt As Variant, MyIndex As String, MyColor As String, MyXcode1 As String, MyXcode2 As String, MyPin1 As String, MyPin2 As String, MyWire As String, MyProject As String, MyHand As Integer, MyString As String, MyTables(1) As Variant

Dim report_path As String, file_name As String, strCriteria As String
    
MyTables(1) = "PRG Test Table Import"

OpenAt = "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\01-Database Master Data\06-Import\06 - Check PRG\"

report_path = BrowseForFolder(OpenAt) & "\"


If Right(report_path, 7) = "COC LL\" Or Right(report_path, 7) = "COC RL\" Or Right(report_path, 7) = "MRA LL\" Or Right(report_path, 7) = "MRA RL\" Then
    If Right(report_path, 12) = "W206\COC LL\" Or Right(report_path, 12) = "W206\COC RL\" Or Right(report_path, 12) = "W206\MRA LL\" Or Right(report_path, 12) = "W206\MRA RL\" Then
        MyProject = Left(Right(report_path, 12), 4)
    ElseIf Right(report_path, 12) = "V297\COC LL\" Or Right(report_path, 12) = "V297\COC RL\" Or Right(report_path, 12) = "V297\MRA LL\" Or Right(report_path, 12) = "V297\MRA RL\" Then
        MyProject = Left(Right(report_path, 12), 4)
    ElseIf Right(report_path, 12) = "V295\COC LL\" Or Right(report_path, 12) = "V295\COC RL\" Or Right(report_path, 12) = "V295\MRA LL\" Or Right(report_path, 12) = "V295\MRA RL\" Then
        MyProject = Left(Right(report_path, 12), 4)
    ElseIf Right(report_path, 12) = "X254\COC LL\" Or Right(report_path, 12) = "X254\COC RL\" Or Right(report_path, 12) = "X254\MRA LL\" Or Right(report_path, 12) = "X254\MRA RL\" Then
        MyProject = Left(Right(report_path, 12), 4)
    ElseIf Right(report_path, 12) = "C236\COC LL\" Or Right(report_path, 12) = "C236\COC RL\" Or Right(report_path, 12) = "C236\MRA LL\" Or Right(report_path, 12) = "C236\MRA RL\" Then
        MyProject = Left(Right(report_path, 12), 4)
    ElseIf Mid(file_name, 2, 3) = "" Then
        a = MsgBox("File not found or in wrong naming scheme!", , "ERROR")
    End If
End If

DoCmd.OpenTable TableName:="PRG Test Table Import"

If Right(Dir(report_path & "*.xlsx", vbDirectory), 4) = "xlsx" Then
    file_name = Dir(report_path & "*.xlsx", vbDirectory)
End If

Do While file_name <> vbNullString
        
    MyLinkDest = (report_path & file_name)
'MsgBox (MyLinkDest)
    Workbooks.OpenText FileName:=MyLinkDest _
        , Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier _
        :=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:= _
        False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1) _
        , TrailingMinusNumbers:=True
    last_row = ActiveWorkbook.ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    last_column = ActiveWorkbook.ActiveSheet.Cells(1, columns.Count).End(xlToLeft).Column
    MyIndex = Left(file_name, 9)
    
    For y = 1 To last_row
        last_row1 = DCount("[Xcode1]", MyTables(1))
        If y = 1 Then
            For x = 1 To last_column
                If Cells(y, x) = Empty Then
                    GoTo Label1
                Else
                    MyString = Cells(y, x)
                    y = y + 1
                    GoTo Label3
                End If
Label1:
            Next
        Else
            x = 1
            MyWire = Cells(y, x)
            MyXcode1 = Cells(y, x + 1)
            MyPin1 = Cells(y, x + 2)
            MyXcode2 = Cells(y, x + 3)
            MyPin2 = Cells(y, x + 4)
            MyColor = Cells(y, x + 5)
        End If
            strCriteria = "INSERT INTO [" & MyTables(1) & "] ([Wire Name], [Xcode1], [Pin1], [Xcode2], [Pin2], [Color], [Index])" & _
                          "VALUES ('" & MyWire & "', '" & MyXcode1 & "', '" & MyPin1 & "', '" & MyXcode2 & "', '" & MyPin2 & "', '" & MyColor & "', '" & MyIndex & "')"
            DoCmd.SetWarnings (WarningsOff)
            DoCmd.RunSQL strCriteria
Label3:
    Next
    
    ActiveWindow.Close
    file_name = Dir
    
Loop
    
    MsgBox "The All Data was Imported. Nice, a!", vbInformation


End Function


