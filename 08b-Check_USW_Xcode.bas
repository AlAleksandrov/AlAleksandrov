Attribute VB_Name = "08b-Check_USW_Xcode"
Option Compare Database

Function Check_USW_Xcode()

Dim x As Integer, y As Long, z As Integer, z1 As Integer, z2 As Integer, last_row As Integer, MyLinkDest As String, last_row1 As Integer, MyRange As Integer, OpenAt As Variant, MyAnumber As String, MyVariant As String, MyXcodeShunk As String, MyXcode As String, MyTyp As String, MyColor As String, MyXcode1 As String, MyPin As String, MyPin1 As String, MyWire As String, MyCrossSection As String, MyATermminal1 As String, MyASeal1 As String, MyPTermminal1 As String, MyPSeal1 As String, MyATermminal2 As String, MyASeal2 As String, MyPTermminal2 As String, MyPSeal2 As String, MyLength As Integer, MyProject As String, MyHand As Integer, MyString As String, MyTables(1) As Variant

Dim report_path As String, file_name As String, strCriteria As String
    
MyTables(1) = "USW_SPLICES"

OpenAt = "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\01-Database Master Data\08-Upload"

report_path = BrowseForFolder(OpenAt) & "\"


If Right(report_path, 12) = "USW_SPLICES\" Then
    If Right(report_path, 17) = "W206\USW_SPLICES\" Then
        MyProject = Left(Right(report_path, 17), 4)
    ElseIf Right(report_path, 17) = "V297\USW_SPLICES\" Then
        MyProject = Left(Right(report_path, 17), 4)
    ElseIf Right(report_path, 17) = "V295\USW_SPLICES\" Then
        MyProject = Left(Right(report_path, 17), 4)
    ElseIf Right(report_path, 17) = "X254\USW_SPLICES\" Then
        MyProject = Left(Right(report_path, 17), 4)
    ElseIf Right(report_path, 17) = "C236\USW_SPLICES\" Then
        MyProject = Left(Right(report_path, 17), 4)
    ElseIf Mid(file_name, 2, 3) = "" Then
        a = MsgBox("File not found or in wrong naming scheme!", , "ERROR")
    End If
End If

DoCmd.OpenTable TableName:="USW_SPLICES"

If Right(Dir(report_path & "*.mdl", vbDirectory), 3) = "mdl" Then
    file_name = Dir(report_path & "*.mdl", vbDirectory)
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
    
    For y = 1 To last_row
        MyXcode = ""
        MyWire = ""
        MyCrossSection = ""
        z = 0
        z1 = 0
        z2 = 0
        last_row1 = DCount("[ID]", MyTables(1))
            For x = 1 To Len(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value)
                If x - 1 = z And (Mid(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x, 1) <> ";") Then
                    If z = 0 Then
                        MyXcodeShunk = Mid(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x, 1)
                        MyXcode = Mid(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x, 1)
                        z = z + 1
                        GoTo Label1
                    Else
                        If Mid(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x, 1) = "#" Or Mid(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x, 1) = "x" Then
                            If Mid(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x, 1) = "#" Then
                                MyXcodeShunk = MyXcodeShunk & "#"
                                MyXcode = MyXcode & "/"
                                z = z + 1
                                GoTo Label1
                            Else
                                MyXcodeShunk = MyXcodeShunk & "x"
                                MyXcode = MyXcode & "*"
                                z = z + 1
                                GoTo Label1
                            End If
                        Else
                            MyXcodeShunk = MyXcodeShunk & Mid(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x, 1)
                            MyXcode = MyXcode & Mid(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x, 1)
                            z = z + 1
                            GoTo Label1
                        End If
                    End If
                End If
                If (Mid(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x, 1) = ";") Then
                    If z1 = 0 Then
                        x = x + 1
                        MyWire = Mid(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x, 1)
                        z1 = z1 + 1
                    Else
                        If Left(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x) = (MyXcodeShunk & ";" & MyWire & ";" & MyCrossSection & ";") Then
                            GoTo Label2
                        Else
                        x = x + 1
                        MyCrossSection = Mid(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x, 1)
                        z2 = z2 + 1
                        End If
                    End If
                Else
                    If z2 = 0 Then
                        MyWire = MyWire & Mid(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x, 1)
                        z1 = z1 + 1
                        GoTo Label1
                    Else
                        MyCrossSection = MyCrossSection & Mid(ActiveWorkbook.ActiveSheet.Cells(y, 1).Value, x, 1)
                        z2 = z2 + 1
                        GoTo Label1
                    End If
                End If
Label1:
            Next
Label2:
            strCriteria = "INSERT INTO [" & MyTables(1) & "] ([Index-Nr], [USW Addresses], [WireNumber], [CrossSection], [Project Name])" & _
                          "VALUES (Left('" & file_name & "', 9), '" & MyXcode & "', '" & MyWire & "', '" & MyCrossSection & "', '" & MyProject & "')"
            DoCmd.SetWarnings (WarningsOff)
            DoCmd.RunSQL strCriteria

    
    Next
    
    ActiveWindow.Close
    file_name = Dir
    
Loop
    
    MsgBox "The All Data was Imported. Nice, a!", vbInformation


End Function

