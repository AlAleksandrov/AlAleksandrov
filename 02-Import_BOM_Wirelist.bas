Attribute VB_Name = "02-Import_BOM_Wirelist"
Option Compare Database

Function Import_BOM_Wirelist_Fixings_Foaming()

Dim x As Integer, y As Long, z As Integer, x1 As Integer, y1 As Long, z1 As Long, last_row As Integer, last_row1 As Integer, MyRange As Integer, MyAnumber As String, MyVariant As String, MyXcode As String, MyTyp As String, MyColor As String, MyXcode1 As String, MyPin As String, MyPin1 As String, MyWire As String, MyCrossSection As String, MyATermminal1 As String, MyASeal1 As String, MyPTermminal1 As String, MyPSeal1 As String, MyATermminal2 As String, MyASeal2 As String, MyPTermminal2 As String, MyPSeal2 As String, MyLength As Integer, MyProject As Integer, MyHand As Integer, MyString As String, MyTables(4) As Variant

Dim report_path As String, file_name As String, strCriteria As String, OpenAt As Variant
    
MyTables(1) = "Wirelist Import"
MyTables(2) = "BOM Import"
MyTables(3) = "Foaming Import"
MyTables(4) = "Fixings Import"

OpenAt = "\\leoni.local\dfsroot\BG1\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\01-Database Master Data\06-Import\01 - Export_from_Drawings\"

report_path = BrowseForFolder(OpenAt) & "\"

If Right(Dir(report_path & "*.csv", vbDirectory), 3) = "csv" Then
    file_name = Dir(report_path & "*.csv", vbDirectory)
Else
    file_name = Dir(report_path & "*.xlsx", vbDirectory)
End If

If Mid(file_name, 2, 3) = "206" Then
    MyProject = DLookup("[ProjectID]", "Project", "[Project Name] = '" & Mid(file_name, 1, 4) & "'")
ElseIf Mid(file_name, 2, 3) = "297" Then
    If Mid(file_name, 1, 4) = "V297" Then
        MyProject = DLookup("[ProjectID]", "Project", "[Project Name] = '" & Mid(file_name, 1, 4) & "'")
    Else
        MyProject = 9
    End If
ElseIf Mid(file_name, 2, 3) = "295" Then
    MyProject = 3
ElseIf Mid(file_name, 2, 3) = "254" Then
    MyProject = DLookup("[ProjectID]", "Project", "[Project Name] = '" & Mid(file_name, 1, 4) & "'")
ElseIf Mid(file_name, 2, 3) = "236" Then
    MyProject = DLookup("[ProjectID]", "Project", "[Project Name] = '" & Mid(file_name, 1, 4) & "'")
ElseIf Mid(file_name, 2, 3) = "" Then
    a = MsgBox("File not found or in wrong naming scheme!", , "ERROR")
    End
End If


Do While file_name <> vbNullString
 
    Call mImportWirelistBOMFiles(report_path, file_name, strCriteria, MyTables(), MyProject)
    
    If (Right(file_name, 12) = "Wirelist.csv") Then
        last_row = DCount("[ID]", "[Wirelist Import]")
        z = DLookup("[ID]", "[Wirelist Import]")
        For y = z To z + last_row
            If (Left(DLookup("[Field1]", "Wirelist Import", "[ID] = " & y), 2) = "A ") Then
                MyAnumber = DLookup("[Field1]", "Wirelist Import", "[ID] = " & y)
                MyVariant = DLookup("[Field2]", "Wirelist Import", "[ID] = " & y)
                GoTo Label1
            End If
            If (Left(DLookup("[Field1]", "Wirelist Import", "[ID] = " & y), 3) = "Ltg") Or IsNull(DLookup("[Field8]", "Wirelist Import", "[ID] = " & y)) Then GoTo Label1
            MyWire = DLookup("[Field1]", "Wirelist Import", "[ID] = " & y)
            MyTyp = DLookup("[Field2]", "Wirelist Import", "[ID] = " & y)
            MyCrossSection = DLookup("[Field3]", "Wirelist Import", "[ID] = " & y)
            MyColor = DLookup("[Field4]", "Wirelist Import", "[ID] = " & y)
            MyXcode = DLookup("[Field5]", "Wirelist Import", "[ID] = " & y)
            MyPin = DLookup("[Field6]", "Wirelist Import", "[ID] = " & y)
            MyXcode1 = DLookup("[Field8]", "Wirelist Import", "[ID] = " & y)
            MyPin1 = DLookup("[Field9]", "Wirelist Import", "[ID] = " & y)
            If IsNull(DLookup("[Field11]", "Wirelist Import", "[ID] = " & y)) Then
                MyATerminal1 = Empty
            Else
                MyATerminal1 = DLookup("[Field11]", "Wirelist Import", "[ID] = " & y)
            End If
            If IsNull(DLookup("[Field12]", "Wirelist Import", "[ID] = " & y)) Then
                MyASeal1 = Empty
            Else
                MyASeal1 = DLookup("[Field12]", "Wirelist Import", "[ID] = " & y)
            End If
            If IsNull(DLookup("[Field13]", "Wirelist Import", "[ID] = " & y)) Then
                MyPTerminal1 = Empty
            Else
                MyPTerminal1 = DLookup("[Field13]", "Wirelist Import", "[ID] = " & y)
            End If
            If IsNull(DLookup("[Field14]", "Wirelist Import", "[ID] = " & y)) Then
                MyPSeal1 = Empty
            Else
                MyPSeal1 = DLookup("[Field14]", "Wirelist Import", "[ID] = " & y)
            End If
            If IsNull(DLookup("[Field15]", "Wirelist Import", "[ID] = " & y)) Then
                MyATerminal2 = Empty
            Else
                MyATerminal2 = DLookup("[Field15]", "Wirelist Import", "[ID] = " & y)
            End If
            If IsNull(DLookup("[Field16]", "Wirelist Import", "[ID] = " & y)) Then
                MyASeal2 = Empty
            Else
                MyASeal2 = DLookup("[Field16]", "Wirelist Import", "[ID] = " & y)
            End If
            If IsNull(DLookup("[Field17]", "Wirelist Import", "[ID] = " & y)) Then
                MyPTerminal2 = Empty
            Else
                MyPTerminal2 = DLookup("[Field17]", "Wirelist Import", "[ID] = " & y)
            End If
            If IsNull(DLookup("[Field18]", "Wirelist Import", "[ID] = " & y)) Then
                MyPSeal2 = Empty
            Else
            MyPSeal2 = DLookup("[Field18]", "Wirelist Import", "[ID] = " & y)
            End If
            'If MyProject = 2 Or MyProject = 3 Or MyProject = 1 Or MyProject = 4 Or MyProject = 5 Or MyProject = 9 Then
                MyLength = DLookup("[Field19]", "Wirelist Import", "[ID] = " & y)
            'Else
                'MyLength = DLookup("[Field15]", "Wirelist Import", "[ID] = " & y)
            'End If
            last_row1 = DCount("[WireListID]", "[WireList]")
            strCriteria = "INSERT INTO WireList ([A-Nr], [WireNumber], [Typ], [CrossSection], [Color], [Xcode], [PinNumber], [Xcode1], [PinNumber1], [A-Terminal1], [A-Seal1], [P-Terminal1], [P-Seal1], [A-Terminal2], [A-Seal2], [P-Terminal2], [P-Seal2], [Length])" & _
                          "VALUES ('" & MyAnumber & "', '" & MyWire & "', '" & MyTyp & "', '" & MyCrossSection & "', '" & MyColor & "', '" & MyXcode & "', '" & MyPin & "', '" & MyXcode1 & "', '" & MyPin1 & "', '" & MyATerminal1 & "', '" & MyASeal1 & "', '" & MyPTerminal1 & "', '" & MyPSeal1 & "', '" & MyATerminal2 & "', '" & MyASeal2 & "', '" & MyPTerminal2 & "', '" & MyPSeal2 & "', '" & MyLength & "')"
            'If DLookup("[A-Nr+]", "A-Nr Variants") = MyAnumber Then
            DoCmd.SetWarnings (WarningsOff)
            DoCmd.RunSQL strCriteria
            'Else
                'DoCmd.SetWarnings (WarningsOn)
            'End If
Label1:
        Next
    'DoCmd.RunSQL "DELETE * from [Wirelist Import]"
    DoCmd.RunSQL "DROP TABLE [" & MyTables(1) & "]"
    ElseIf (Right(file_name, 7) = "BOM.csv") Then
        last_row = DCount("[ID]", "[BOM Import]")
        z = DLookup("[ID]", "[BOM Import]")
        For y = z To z + last_row
            If (Left(DLookup("[Field1]", "BOM Import", "[ID] = " & y), 2) = "A ") Then
                MyAnumber = DLookup("[Field1]", "BOM Import", "[ID] = " & y)
                MyVariant = DLookup("[Field2]", "BOM Import", "[ID] = " & y)
                GoTo Label2
            End If
            'If MyProject = 2 Or MyProject = 3 Or MyProject = 1 Or MyProject = 4 Or MyProject = 5 Or MyProject = 9 Then
                If (Left(DLookup("[Field6]", "BOM Import", "[ID] = " & y), 6) = "F152/3") Or (Left(DLookup("[Field3]", "BOM Import", "[ID] = " & y), 6) = "MUTTER") Then
                    last_row1 = DCount("[BOMID]", "[BOM]")
                    strCriteria = "INSERT INTO BOM ([A-Nr], [Address], [Description], [Place], [Fuse&Relay Number])" & _
                                "VALUES ('" & MyAnumber & "', '" & DLookup("[Field6]", "BOM Import", "[ID] = " & y) & "', '" & DLookup("[Field3]", "BOM Import", "[ID] = " & y) & "', '" & DLookup("[Field5]", "BOM Import", "[ID] = " & y) & "', '" & DLookup("[Field2]", "BOM Import", "[ID] = " & y) & "')"
                    'If DLookup("[A-Nr+]", "A-Nr Variants") = MyAnumber Then
                    DoCmd.SetWarnings (WarningsOff)
                    DoCmd.RunSQL strCriteria
                    'Else
                        'DoCmd.SetWarnings (WarningsOn)
                    'End If
                End If
            'Else
                'If (Left(DLookup("[Field5]", "BOM Import", "[ID] = " & y), 6) = "F152/3") Or (Left(DLookup("[Field2]", "BOM Import", "[ID] = " & y), 6) = "MUTTER") Then
                    'last_row1 = DCount("[BOMID]", "[BOM]")
                    'strCriteria = "INSERT INTO BOM ([A-Nr], [Address], [Description], [Place], [Fuse&Relay Number])" & _
                                "VALUES ('" & MyAnumber & "', '" & DLookup("[Field5]", "BOM Import", "[ID] = " & y) & "', '" & DLookup("[Field2]", "BOM Import", "[ID] = " & y) & "', '" & DLookup("[Field4]", "BOM Import", "[ID] = " & y) & "', '" & DLookup("[Field1]", "BOM Import", "[ID] = " & y) & "')"
                    'If DLookup("[A-Nr+]", "A-Nr Variants") = MyAnumber Then
                    'DoCmd.SetWarnings (WarningsOff)
                    'DoCmd.RunSQL strCriteria
                    'Else
                        'DoCmd.SetWarnings (WarningsOn)
                    'End If
                'End If
            'End If
Label2:
        Next
    'DoCmd.RunSQL "DELETE * from [BOM Import]"
    DoCmd.RunSQL "DROP TABLE [" & MyTables(2) & "]"
    ElseIf (Right(file_name, 7) = "FOA.csv") Then
        last_row = DCount("[ID]", "[Foaming Import]")
        z = DLookup("[ID]", "[Foaming Import]")
        'MyProject = DLookup("[ProjectID]", "Project", "[Project Name] = '" & Mid(file_name, 1, 4) & "'")
        MyHand = DLookup("[HandID]", "Hands", "[Hands] = '" & Mid(file_name, 6, 2) & "'")
        For y = z To z + last_row - 1
            MyWire = DLookup("[No]", "Foaming Import", "[ID] = " & y)
            MyCrossSection = DLookup("[CSA]", "Foaming Import", "[ID] = " & y)
            MyString = DLookup("[Modules]", "Foaming Import", "[ID] = " & y)
            For x = 1 To 1000
                If Mid(MyString, x, 2) = "A " Then
                    MyAnumber = Mid(MyString, x, 15)
                    last_row1 = DCount("[FoamingID]", "[Foaming]")
                    z1 = Nz(DLookup("[FoamingID]", "[Foaming]"), "0")
                    If Mid(MyAnumber, 3, 3) = "206" Then
                        MyProject = DLookup("[ProjectID]", "Project", "[Project Name] = '" & Mid(file_name, 1, 4) & "'")
                    ElseIf Mid(MyAnumber, 3, 3) = "297" Then
                        If Mid(file_name, 1, 4) = "V297" Then
                            MyProject = DLookup("[ProjectID]", "Project", "[Project Name] = '" & Mid(file_name, 1, 4) & "'")
                        Else
                            MyProject = 9
                        End If
                    ElseIf Mid(MyAnumber, 3, 3) = "295" Then
                        MyProject = 3
                    ElseIf Mid(MyAnumber, 3, 3) = "254" Then
                        MyProject = DLookup("[ProjectID]", "Project", "[Project Name] = '" & Mid(file_name, 1, 4) & "'")
                    ElseIf Mid(MyAnumber, 3, 3) = "236" Then
                        MyProject = DLookup("[ProjectID]", "Project", "[Project Name] = '" & Mid(file_name, 1, 4) & "'")
                    Else
                        GoTo Label3
                    End If
                    'Call CheckForDuplicates(file_name, y, y1, z1, last_row1, MyWire, MyCrossSection, MyAnumber, MyProject, MyHand, Myrange)
                    If MyRange = 1 Then
                        GoTo Label3
                    End If
                    strCriteria = "INSERT INTO Foaming ([Wirenumber], [CrossSection], [A-Nr], [ProjectID], [HandID])" & _
                                  "VALUES ('" & MyWire & "', '" & MyCrossSection & "', '" & MyAnumber & "', '" & MyProject & "', '" & MyHand & "')"
                    'If DLookup("[A-Nr+]", "A-Nr Variants") = MyAnumber Then
                    DoCmd.SetWarnings (WarningsOff)
                    DoCmd.RunSQL strCriteria
                    'Else
                        'DoCmd.SetWarnings (WarningsOn)
                    'End If
Label3:
                MyRange = 0
                x = x + 15
                End If
            Next
        Next
    'DoCmd.RunSQL "DELETE * from [Foaming Import]"
    DoCmd.RunSQL "DROP TABLE [" & MyTables(3) & "]"
    ElseIf (Right(file_name, 12) = "Fixings.xlsx") Then
        last_row = DCount("[FixingsID]", "[Fixings Import]")
        z = DLookup("[FixingsID]", "[Fixings Import]")
        For y = z To z + last_row - 1
            MyXcode = DLookup("[Id]", "Fixings Import", "[FixingsID] = " & y)
            MyString = DLookup("[Assigned modules]", "Fixings Import", "[FixingsID] = " & y)
            For x = 1 To 500
                If Mid(MyString, x, 3) = "A 2" Then
                    MyAnumber = Mid(MyString, x, 15)
                    last_row1 = DCount("[ID]", "[Fixings]")
                    strCriteria = "INSERT INTO Fixings ([Fixings], [Part description], [Fixing type], [Assigned module])" & _
                                  "VALUES ('" & MyXcode & "', '" & DLookup("[Part description]", "Fixings Import", "[FixingsID] = " & y) & "', '" & DLookup("[Fixing type]", "Fixings Import", "[FixingsID] = " & y) & "', '" & MyAnumber & "')"
                    'If DLookup("[A-Nr+]", "A-Nr Variants") = MyAnumber Then
                    DoCmd.SetWarnings (WarningsOff)
                    DoCmd.RunSQL strCriteria
                    'Else
                        'DoCmd.SetWarnings (WarningsOn)
                    'End If
                End If
            Next
        Next
    'DoCmd.RunSQL "DELETE * from [Fixings Import]"
    DoCmd.RunSQL "DROP TABLE [" & MyTables(4) & "]"
    End If
    
    file_name = Dir
    
Loop
    
    MsgBox "The All Data was Imported. Nice, a!", vbInformation


End Function

'------------------------------------------------------------
' mImportWirelistBOMFiles
'
'------------------------------------------------------------
Function mImportWirelistBOMFiles(report_path As String, file_name As String, strCriteria As String, MyTables() As Variant, MyProject As Integer)
On Error GoTo mImportWirelistFiles_Err
    
    If (Right(file_name, 12) = "Wirelist.csv") Then
       'If MyProject = 2 Or MyProject = 3 Or MyProject = 1 Or MyProject = 4 Or MyProject = 5 Or MyProject = 9 Then
            DoCmd.TransferText acImportDelim, "Wirelist Import Specification", "Wirelist Import", report_path & file_name, True, ""
            strCriteria = "ALTER TABLE [" & MyTables(1) & "]" & _
                          "ADD Column ID Counter"
            DoCmd.RunSQL strCriteria
        'Else
            'DoCmd.TransferText acImportDelim, "Wirelist Import Specification Other", "Wirelist Import", report_path & file_name, True, ""
            'strCriteria = "ALTER TABLE [" & MyTables(1) & "]" & _
                          "ADD Column ID Counter"
            'DoCmd.RunSQL strCriteria
        'End If
    ElseIf (Right(file_name, 7) = "BOM.csv") Then
       'If MyProject = 2 Or MyProject = 3 Or MyProject = 1 Or MyProject = 4 Or MyProject = 5 Or MyProject = 9 Then
            DoCmd.TransferText acImportDelim, "BOM Import Specification", "BOM Import", report_path & file_name, True, ""
            strCriteria = "ALTER TABLE [" & MyTables(2) & "]" & _
                          "ADD Column ID Counter"
            DoCmd.RunSQL strCriteria
        'Else
            'DoCmd.TransferText acImportDelim, "BOM Import Specification Other", "BOM Import", report_path & file_name, True, ""
            'strCriteria = "ALTER TABLE [" & MyTables(2) & "]" & _
                          "ADD Column ID Counter"
            'DoCmd.RunSQL strCriteria
        'End If
    ElseIf (Right(file_name, 7) = "FOA.csv") Then
    'Call Open_Excel_files(report_path, file_name)
        DoCmd.TransferText acImportDelim, "Foaming Import Specification", "Foaming Import", report_path & file_name, True, ""
            strCriteria = "ALTER TABLE [" & MyTables(3) & "]" & _
                          "ADD Column ID Counter, ProjectID INTEGER, HandID INTEGER"
            DoCmd.RunSQL strCriteria
    ElseIf (Right(file_name, 12) = "Fixings.xlsx") Then
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "Fixings Import", report_path & file_name, True, ""
            strCriteria = "ALTER TABLE [" & MyTables(4) & "]" & _
                          "ADD Column FixingsID Counter"
            DoCmd.RunSQL strCriteria
    End If
    
mImportWirelistFiles_Exit:
    Exit Function

mImportWirelistFiles_Err:
    MsgBox Error$
    Resume mImportWirelistFiles_Exit

End Function

'------------------------------------------------------------
' CheckForDuplicates
'
'------------------------------------------------------------
Function CheckForDuplicates(file_name As String, y As Long, y1 As Long, z1 As Integer, last_row1 As Integer, MyWire As String, MyCrossSection As String, MyAnumber As String, MyProject As Integer, MyHand As Integer, MyRange As Integer)
On Error GoTo CheckForDuplicates_Err
    
    If (Right(file_name, 12) = "Wirelist.csv") Then
    
    ElseIf (Right(file_name, 7) = "BOM.csv") Then
    
    ElseIf (Right(file_name, 7) = "FOA.csv") Then
        For y1 = z1 To z1 + last_row1 - 1
            If MyAnumber = DLookup("[A-Nr]", "Foaming", "[FoamingID] =  " & y1) Then
                If MyWire = DLookup("[WireNumber]", "Foaming", "[FoamingID] =  " & y1) Then
                    If MyCrossSection = DLookup("[CrossSection]", "Foaming", "[FoamingID] =  " & y1) Then
                        If MyHand = DLookup("[HandID]", "Foaming", "[FoamingID] =  " & y1) Then
                            If MyProject = DLookup("[ProjectID]", "Foaming", "[FoamingID] =  " & y1) Then
                                MyRange = 1
                            End If
                        End If
                    End If
                End If
            End If
        Next
    ElseIf (Right(file_name, 12) = "Fixings.xlsx") Then
    
    End If
    
CheckForDuplicates_Exit:
    Exit Function

CheckForDuplicates_Err:
    MsgBox Error$
    Resume CheckForDuplicates_Exit

End Function

'''------------------------------------------------------------''
'''                        FolderSelector                      ''
'''                         AI Optimized                       ''
'''------------------------------------------------------------''
'Function BrowseForFolder(Optional OpenAt As Variant, Optional WindowTitle As String = "Please choose a folder") As Variant
'    'Function purpose: To browse for a user-selected folder.
'    'If the "OpenAt" path is provided, open the browser at that directory.
'    'NOTE: If invalid, it will open at the desktop level.
'
'    'Create a file browser window at the default folder.
'    Dim ShellApp As Shell32.Shell
'    Set ShellApp = New Shell32.Shell
'
'    'Handle errors creating the Shell.Application object.
'    On Error Resume Next
'    If OpenAt = "" Then
'        Set ShellApp = New Shell32.Shell
'    Else
'        Set ShellApp = New Shell32.Shell
'        Set ShellApp = ShellApp.BrowseForFolder(0, WindowTitle, 0, OpenAt)
'    End If
'    On Error GoTo 0
'
'    'Set the folder to the selected path. (On error in case cancelled)
'    On Error Resume Next
'    BrowseForFolder = ShellApp.self.Path
'    On Error GoTo 0
'
'    'Destroy the Shell.Application object.
'    Set ShellApp = Nothing
'
'    'Check for invalid or non-entries and send to the Invalid error handler if found.
'    'Valid selections can begin L: (where L is a letter) or \\ (as in \\servername\sharename). All others are invalid.
'    Select Case Mid(BrowseForFolder, 2, 1)
'        Case Is = ":"
'            If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
'        Case Is = "\"
'            If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
'        Case Else
'            GoTo Invalid
'    End Select
'
'    Exit Function
'
'Invalid:
'    'If it was determined that the selection was invalid, set to False.
'    BrowseForFolder = False
'End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------
' FolderSelector
'
'------------------------------------------------------------
Function BrowseForFolder(Optional OpenAt As Variant) As Variant
     'Function purpose:  To Browser for a user selected folder.
     'If the "OpenAt" path is provided, open the browser at that directory
     'NOTE:  If invalid, it will open at the Desktop level
    Dim ShellApp As Object

    'If OpenAt = "" Then OpenAt = "\\SVBG1FILE01\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\01-Database Master Data\06-Import\01 - Export_from_Drawings\"


     'Create a file browser window at the default folder
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Please choose a folder", 0, OpenAt)

     'Set the folder to that selected.  (On error in case cancelled)
    On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
    On Error GoTo 0

     'Destroy the Shell Application
    Set ShellApp = Nothing

     'Check for invalid or non-entries and send to the Invalid error
     'handler if found
     'Valid selections can begin L: (where L is a letter) or
     '\\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
    Case Is = ":"
        If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select

    Exit Function

Invalid:
     'If it was determined that the selection was invalid, set to False
    BrowseForFolder = False
End Function

