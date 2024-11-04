Attribute VB_Name = "09b-AEM_Import_TXT_to_Access_DBMS"
 Option Compare Database

Function AEM_Import_TXT_to_Access_DBMS()
'
' TXT_to_Access_DBMS Macro
' Save data from .txt to Access DBMS
'
Dim x As Integer, y As Long, z As Long, MyLinkDest As String, MyLink As String, MyPDFName(1000, 1), MyType(2), MyFile As String, intPath As Integer, intPDFFile As Long

Dim i As Integer, j As Integer, k As Integer, l As Integer, m As Integer, n As Integer, a As Integer, b As Integer, c As Integer, last_row As Integer, Response As Integer, project As String, MyProjectAEM As String, MyProjectModule As String, AEMNumber As String, ChangeRequest As String, AEMDate As Date, DrawingDate As Date, Types As String, Hand1 As String, Hand2 As String, Hand As String, Description As String, strCriteria As String, strZGSNumber As String, strAEMNumber As String, Deviation As String

Dim KickOffArray(8, 100) As Variant, AEMUpdateArray(1000, 7) As Variant, report_path As String, OpenAt As Variant

OpenAt = "\\leoni.local\dfsroot\BG1\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\01-Database Master Data\06-Import\05 - AEMs\"

report_path = BrowseForFolder(OpenAt) & "\"

If Right(report_path, 4) = "TXT\" Then
    If Right(report_path, 9) = "W206\TXT\" Then
        project = Left(Right(report_path, 9), 4)
    ElseIf Right(report_path, 9) = "V297\TXT\" Then
        project = Left(Right(report_path, 9), 4)
    ElseIf Right(report_path, 9) = "V295\TXT\" Then
        project = Left(Right(report_path, 9), 4)
    ElseIf Right(report_path, 9) = "X254\TXT\" Then
        project = Left(Right(report_path, 9), 4)
    ElseIf Right(report_path, 9) = "C236\TXT\" Then
        project = Left(Right(report_path, 9), 4)
    ElseIf Mid(file_name, 2, 3) = "" Then
        a = MsgBox("File not found or in wrong naming scheme!", , "ERROR")
    End If
End If

MyLink = OpenAt

For x = 0 To 1
    If x = 0 Then
        MyType(x) = "COC"
        MyLinkDest = (MyLink & project & "\TXT\")
        intPath = Len(Dir(MyLinkDest, vbDirectory))
        If intPath > 0 Then
            intPDFFile = Len(Dir(MyLinkDest & MyType(x) & "_*.pdf.txt"))
            If intPDFFile = 0 Then GoTo Label1
            MyFile = Dir(MyLinkDest & MyType(x) & "_*.pdf.txt")
            While (MyFile <> "")
                MyPDFName(i, x) = MyFile
                i = i + 1
                MyFile = Dir
            Wend
        Else
            GoTo Label1
        End If
    ElseIf x = 1 Then
        'i = 0
        MyType(x) = "MRA"
        MyLinkDest = (MyLink & project & "\TXT\")
        intPath = Len(Dir(MyLinkDest, vbDirectory))
        If intPath > 0 Then
            intPDFFile = Len(Dir(MyLinkDest & MyType(x) & "_*.pdf.txt"))
            If intPDFFile = 0 Then GoTo Label1
            MyFile = Dir(MyLinkDest & MyType(x) & "_*.pdf.txt")
            While (MyFile <> "")
                MyPDFName(i, x) = MyFile
                i = i + 1
                MyFile = Dir
            Wend
        Else
            GoTo Label1
        End If
    End If
    For y = y To i - 1
        MyLinkDest = (MyLink & project & "\TXT\")
'MsgBox (MyLinkDest)
        ChDir MyLinkDest
        MyLinkDest = (MyLink & project & "\TXT\" & MyPDFName(y, x))
'MsgBox (MyLinkDest)
        Workbooks.OpenText FileName:=MyLinkDest _
            , Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier _
            :=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:= _
            False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1) _
            , TrailingMinusNumbers:=True
        last_row = ActiveWorkbook.ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
        For z = 1 To last_row
            If ActiveWorkbook.ActiveSheet.Cells(z, 1).Value = "Datum/Date" Then
                AEMDate = Cells(z + 1, 1)
                KickOffArray(4, y) = AEMDate
                z = z + 1
            End If
            If Left(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 26) = "Änderungsmeldung Lfd.-Nr. " Then
                b = InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, "BNE")
                If b > 0 Then
                    If KickOffArray(2, y) <> "" Then GoTo Label2
                    AEMNumber = Mid(Cells(z, 1), 27, b - 1 - 27)
                    KickOffArray(2, y) = AEMNumber
                    AEMNumber = ""
                    b = 0
                Else
                    If KickOffArray(2, y) <> "" Then GoTo Label2
                    AEMNumber = Mid(Cells(z, 1), 27, 100)
                    KickOffArray(2, y) = AEMNumber
                    AEMNumber = ""
                End If
            End If
            If InStr(Mid(KickOffArray(2, y), 4, 3), "206") = 1 Then
                MyProjectAEM = "W206"
            ElseIf InStr(Mid(KickOffArray(2, y), 4, 3), "236") = 1 Then
                MyProjectAEM = "C236"
            ElseIf InStr(KickOffArray(2, y), "297") > 0 Then
                MyProjectAEM = "V297"
            ElseIf InStr(KickOffArray(2, y), "295") > 0 Then
                MyProjectAEM = "V295"
            ElseIf InStr(Mid(KickOffArray(2, y), 4, 3), "254") = 1 Then
                MyProjectAEM = "X254"
            End If
Label2:
            If InStr(Left(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 2), "BR") = 1 And InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, MyType(x)) > 0 Or InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, Right(MyProjectAEM, 3)) > 0 Then
                KickOffArray(0, y) = MyProjectAEM
                b = InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, "AMG")
                If b > 0 Then
                    Deviation = "AMG"
                    If Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 9) = "LL/RL AMG" Or Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 9) = "AMG LL/RL" Then
                        If KickOffArray(6, y) <> "" Then GoTo Label3
                        If KickOffArray(7, y) <> "" Then GoTo Label3
                        If Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 9) = "LL/RL AMG" Then
                            Types = Left(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 13), 3)
                            Hand = Left(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 9), 5)
                        ElseIf Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 9) = "AMG LL/RL" Then
                            Types = Left(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 13), 3)
                            Hand = Right(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 9), 5)
                        End If
                        KickOffArray(6, y) = Types
                        KickOffArray(7, y) = Hand
                    ElseIf Left(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 6), 2) = "RL" Or Left(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 6), 2) = "LL" Then
                        If KickOffArray(6, y) <> "" Then GoTo Label3
                        If KickOffArray(7, y) <> "" Then GoTo Label3
                        Types = Left(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 10), 3)
                        Hand = Left(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 6), 2)
                        Hand1 = Hand
                        Hand2 = Hand
                        KickOffArray(6, y) = Types
                        KickOffArray(7, y) = Hand
                    ElseIf Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 7) = "MRA AMG" Then
                        If KickOffArray(6, y) <> "" Then GoTo Label3
                        If KickOffArray(7, y) <> "" Then GoTo Label3
                        Types = Left(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 7), 3)
                        KickOffArray(6, y) = Types
                    End If
                Else
                    Deviation = "Serial"
                    If Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 5) = "LL/RL" Then
                        If KickOffArray(6, y) <> "" Then GoTo Label3
                        If KickOffArray(7, y) <> "" Then GoTo Label3
                        Types = Left(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 9), 3)
                        Hand = Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 5)
                        KickOffArray(6, y) = Types
                        KickOffArray(7, y) = Hand
                    ElseIf Right(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 5), 2) = "RL" Or Right(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 5), 2) = "LL" Then
                        If KickOffArray(6, y) <> "" Then GoTo Label3
                        If KickOffArray(7, y) <> "" Then GoTo Label3
                        Types = Left(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 6), 3)
                        Hand = Right(Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 5), 2)
                        Hand1 = Hand
                        Hand2 = Hand
                        KickOffArray(6, y) = Types
                        KickOffArray(7, y) = Hand
                    ElseIf Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 3) = "MRA" Or Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 3) = "COC" Then
                        If KickOffArray(6, y) <> "" Then GoTo Label3
                        If KickOffArray(7, y) <> "" Then GoTo Label3
                        Types = Right(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 3)
                        KickOffArray(6, y) = Types
                    End If
                End If
            End If
Label3:
            If Left(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 20) = "Änderungsliste Nr.: " Then
                If InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, "Fest") > 0 Then
                    b = InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, "Fest")
                    If KickOffArray(3, y) <> "" Then GoTo Label4
                    ChangeRequest = Mid(Cells(z, 1), 21, b - 1 - 21)
                    KickOffArray(3, y) = ChangeRequest
                    b = 0
                Else
                    If KickOffArray(3, y) <> "" Then GoTo Label4
                    ChangeRequest = Mid(Cells(z, 1), 21, 30)
                    KickOffArray(3, y) = ChangeRequest
                End If
            End If
Label4:
            If InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, "Nacharbeit") And InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, "NEIN") Then
                If KickOffArray(8, y) <> "" Then GoTo Label6
                While Left(Cells(z + 1, 1), 7) <> "Einsatz"
                    If Cells(z + 1, 1) <> " " Then
                        Description = Description & Cells(z + 1, 1) & " "
                    End If
                    z = z + 1
                    If Left(Cells(z + 1, 1), 6) = "Gültig" Or Left(Cells(z + 1, 1), 12) = "Sehr geehrte" Then GoTo Label5
                Wend
Label5:
                KickOffArray(8, y) = Description
                Description = ""
            End If
Label6:
            If Hand1 = Empty And Hand2 = Empty Then
                If InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, "Änderungsbeschreibung") And InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, "Bemerkung") Then
                    z = z + 1
                    If InStr(Left(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 2), "BR") = 1 Or InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, MyType(x)) > 0 Or InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, Right(MyProjectAEM, 3)) > 0 Then
                        Hand1 = Right(Cells(z, 1), 2)
                    End If
                End If
            Else
                GoTo Label7
            End If
Label7:
            If Left(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 11) = "KEM aktuell" Then
                z = z + 1
                While (Cells(z, 1) <> "Die Änderung erfolgt unter folgenden Kriterien:")
                    If Left(Cells(z, 1), 2) = "A" & Mid(project, 2, 1) Then
                        AEMUpdateArray(j, 0) = KickOffArray(2, y)
                        If InStr(Mid(Cells(z, 1), 2, 3), "206") = 1 Then
                            MyProjectModule = "W206"
                        ElseIf InStr(Mid(Cells(z, 1), 2, 3), "236") = 1 Then
                            MyProjectModule = "C236"
                        ElseIf InStr(Mid(Cells(z, 1), 2, 3), "295") = 1 Then
                            MyProjectModule = "V295"
                        ElseIf InStr(Mid(Cells(z, 1), 2, 3), "297") = 1 Then
                            MyProjectModule = "V297"
                        ElseIf InStr(Mid(Cells(z, 1), 2, 3), "296") = 1 Or InStr(Mid(Cells(z, 1), 2, 3), "294") = 1 Then
                            z = z + 1
                            GoTo Label9
                        ElseIf InStr(Mid(Cells(z, 1), 2, 3), "254") = 1 Then
                            MyProjectModule = "X254"
                        End If
                        AEMUpdateArray(j, 1) = Cells(z, 1)
                        AEMUpdateArray(j, 2) = Cells(z + 1, 1)
                        If strAEMNumber = "" And strZGSNumber = "" Then
                            strAEMNumber = AEMUpdateArray(j, 1)
                            strZGSNumber = AEMUpdateArray(j, 2)
                        Else
                            strAEMNumber = strAEMNumber & "; " & AEMUpdateArray(j, 1)
                            strZGSNumber = strZGSNumber & "; " & AEMUpdateArray(j, 2)
                        End If
                        If MyProjectAEM = project Then
                            AEMUpdateArray(j, 3) = project
                        Else
                            AEMUpdateArray(j, 3) = MyProjectModule
                        End If
                        AEMUpdateArray(j, 4) = MyType(x)
                        AEMUpdateArray(j, 6) = Deviation
                        If Hand2 = Empty Then
                            AEMUpdateArray(j, 5) = Hand1
                        Else
                            AEMUpdateArray(j, 5) = Hand2
                        End If
                        If Left(Cells(z + 4, 1), 14) = "ZB EL.LTG.SATZ" Then
                            l = 2
                            z = z + 4
                            GoTo Label8
                        ElseIf InStr(Cells(z + 4, 1), "EL.LTG.SATZ") > 0 Then
                            l = 1
                            z = z + 4
                            GoTo Label8
                        End If
                        If Left(Cells(z + 5, 1), 14) = "ZB EL.LTG.SATZ" Then
                            l = 2
                            z = z + 5
                            GoTo Label8
                        ElseIf InStr(Cells(z + 5, 1), "EL.LTG.SATZ") > 0 Then
                            l = 1
                            z = z + 5
                            GoTo Label8
                        End If
                        While Cells(z, 1) <> "" Or Left(Cells(z, 1), 1) = "A"
Label8:
                            If Cells(z, 1) <> "" Or Left(Cells(z, 1), 1) = "A" Then
                                If Right(Cells(z, 1), 7) = "COCKPIT" Or Right(Cells(z, 1), 8) = "MOT.RAUM" Then
                                    Description = Description & Cells(z, 1)
                                    GoTo Label13
                                Else
                                    If l = 1 Then
                                        If Description = "" Then
                                            Description = "ZB " & Description & Cells(z, 1) & " "
                                        Else
                                            Description = Description & Cells(z, 1) & " "
                                        End If
                                    ElseIf l = 2 Then
                                        Description = Description & Cells(z, 1) & " "
                                    End If
                                End If
                            End If
                            z = z + 1
                        Wend
Label13:
                        AEMUpdateArray(j, 7) = Description
                        Description = ""
                        If DCount("[A-Nr]", "All Modules", "[A-Nr] = '" & AEMUpdateArray(j, 1) & "'") = 1 Then
                            GoTo Label9
                        Else
                            m = j
                            n = n + 1
                            GoTo Label9
                        End If
Label9:
                        l = 0
                        j = j + 1
                        z = z + 2
                    End If
                    z = z + 1
                    If Hand2 = Empty Then
                    If InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, "Änderungsbeschreibung") And InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, "Bemerkung") Then
                        z = z + 1
                        If InStr(Left(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 2), "BR") = 1 Or InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, MyType(x)) > 0 Or InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, Right(MyProjectAEM, 3)) > 0 Then
                                Hand2 = Right(Cells(z, 1), 2)
                            End If
                        End If
                    Else
                        GoTo Label10
                    End If
Label10:
                Wend
            End If
            'If Mid(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 2, 3) = MyProjectAEM Then
            '    z = z + 6
            '    If InStr(ActiveWorkbook.ActiveSheet.Cells(z + 2, 1).Value, "Neues Zeichnungsdatum") Or Left(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 25) = "– Neues Zeichnungsdatum: " Then
            '        If KickOffArray(5, y) <> "" Then GoTo Label11
            '        DrawingDate = Mid(Cells(z, 1), 26, 10)
            '        KickOffArray(5, y) = DrawingDate
            '    End If
            If InStr(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, "Zeichnungs-Kriterien") Then
                If InStr(ActiveWorkbook.ActiveSheet.Cells(z + 2, 1).Value, "Neues Zeichnungsdatum") Or Left(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 25) = "– Neues Zeichnungsdatum: " Then
                    If KickOffArray(5, y) <> "" Then GoTo Label11
                    DrawingDate = Mid(Cells(z + 2, 1), 26, 10)
                    KickOffArray(5, y) = DrawingDate
                ElseIf InStr(ActiveWorkbook.ActiveSheet.Cells(z + 3, 1).Value, "Neues Zeichnungsdatum") Or Left(ActiveWorkbook.ActiveSheet.Cells(z, 1).Value, 25) = "– Neues Zeichnungsdatum: " Then
                    If KickOffArray(5, y) <> "" Then GoTo Label11
                    DrawingDate = Mid(Cells(z + 3, 1), 26, 10)
                    KickOffArray(5, y) = DrawingDate
                End If
            End If
Label11:
        Next
        If KickOffArray(4, y) = Empty Then
            last_row = ActiveWorkbook.ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row
            For z = 1 To last_row
                If ActiveWorkbook.ActiveSheet.Cells(z, 2).Value = "Datum/Date" Then
                    AEMDate = Cells(z + 1, 2)
                    KickOffArray(4, y) = AEMDate
                End If
            Next
        End If
        ActiveWindow.Close
        If Hand = "" And Hand1 <> "" And Hand2 = "" Then
            Hand = Hand1
        ElseIf Hand = "" And Hand1 = "" And Hand2 <> "" Then
            Hand = Hand1
        ElseIf Hand = "" And Hand1 <> "" And Hand2 <> "" And Hand1 <> Hand2 Then
            Hand = Hand1 & "/" & Hand2
        ElseIf Hand = "" And Hand1 <> "" And Hand2 <> "" And Hand1 = Hand2 Then
            Hand = Hand1
        End If
        DoCmd.OpenTable "Project", acViewNormal, acReadOnly
        For k = 1 To 8
            If DLookup("[Project Name]", "Project", "[ProjectID] = " & k) = MyProjectAEM Then GoTo Label12
        Next
Label12:
        DoCmd.Close
        strCriteria = "INSERT INTO [Kick Off] ([AEM-Nr], [CR], [AEM Date], [Drawing Date], [Type], [Hand], [Description], [ProjectID], [AEM Phase])" & _
                        "VALUES ('" & KickOffArray(2, y) & "', '" & KickOffArray(3, y) & "', '" & KickOffArray(4, y) & "', '" & KickOffArray(5, y) & "', '" & KickOffArray(6, y) & "', '" & KickOffArray(7, y) & "', '" & KickOffArray(8, y) & "', '" & k & "', '" & DLookup("[PhaseID]", "Project", "[ProjectID] = " & k) & "')"
        DoCmd.SetWarnings (WarningsOff)
        DoCmd.RunSQL strCriteria
        If n > 0 Then
            Call Add_New_Modules_From_AEMs(KickOffArray(), AEMUpdateArray(), m, n, y)
            GoTo Label101
        End If
        DoCmd.OpenForm FormName:="Kick Off"
        Forms![Kick Off]![ProjectID] = k
        Forms![Kick Off]![AEM Phase] = DLookup("[PhaseID]", "Project", "[ProjectID] = " & k)
        Forms![Kick Off]![AEM-Nr] = KickOffArray(2, y)
        Forms![Kick Off]![CR] = KickOffArray(3, y)
        Forms![Kick Off]![AEM Date] = KickOffArray(4, y)
        Forms![Kick Off]![Drawing Date] = KickOffArray(5, y)
        Forms![Kick Off]![Type] = KickOffArray(6, y)
        Forms![Kick Off]![Hand] = KickOffArray(7, y)
        Forms![Kick Off]![Description] = KickOffArray(8, y)
        
        Forms![Kick Off]![AEM Update]![AEM_Nr] = KickOffArray(2, y)
        Forms![Kick Off]![AEM Update]![List9].RowSource = strAEMNumber
        Forms![Kick Off]![AEM Update]![Text108] = Forms![Kick Off]![AEM Update]![List9].ListCount
        Forms![Kick Off]![AEM Update]![List15].RowSource = strZGSNumber
        Forms![Kick Off]![AEM Update]![Text188] = Forms![Kick Off]![AEM Update]![List15].ListCount
        
        Response = MsgBox("Start process A-Nr and ZGS in AEM Update?", vbYesNo + vbCritical + vbDefaultButton1, "AEM Update", "Help", 1000)
        
        If Response = vbYes Then    ' User chose Yes.
            Call [Form_AEM Update].Command39_Click
            'MyString = "Yes"    ' Perform some action.
        Else    ' User chose No.
            'MyString = "No"    ' Perform some action.
            GoTo Label1
        End If
        
        'DoCmd.Close
Label101:
n = 0
strAEMNumber = ""
strZGSNumber = ""
Hand1 = ""
Hand2 = ""
Hand = ""
Types = ""
    Next
Label1:
Next

MsgBox "The All AEMs was Imported. Nice, a!", vbInformation

End Function

Function Add_New_Modules_From_AEMs(KickOffArray() As Variant, AEMUpdateArray() As Variant, m As Integer, n As Integer, y As Long)

Dim MyArray(1000, 7), MyTables(1) As Variant, strCriteria As String, MyAnumber As String, MyLiumf As String, MyDescription As String

Dim i As Integer, j As Integer, k As Integer, z As Integer, a As Integer, b As Integer, c As Integer

MyTables(1) = "New Modules Table"
k = 0

For i = m + 1 - n To m
    For j = 0 To 7
        If j = 0 Then
            MyArray(i, j) = KickOffArray(5, y)
        ElseIf j = 1 Then
            If DCount("[A-Nr]", "New Modules Table", "[A-Nr] = '" & AEMUpdateArray(i, j) & "'") = 1 Then
                MyAnumber = AEMUpdateArray(i, j)
                GoTo Label1
            Else
                MyArray(i, j) = AEMUpdateArray(i, j)
            End If
        ElseIf j = 2 Then
            If AEMUpdateArray(i, j) = "" Then
                MyArray(i, j) = 1
            Else
                MyArray(i, j) = AEMUpdateArray(i, j)
            End If
        ElseIf j = 3 Then
            For z = 1 To 10
                a = z
                If DLookup("[Project Name]", "Project", "[ProjectID] = " & z) = AEMUpdateArray(i, j) Then GoTo Label2
            Next
Label2:
            MyArray(i, j) = a
        ElseIf j = 4 Then
            For z = 1 To 2
                b = z
                If DLookup("[Types]", "Types", "[TypeID] = " & z) = AEMUpdateArray(i, j) Then GoTo Label3
            Next
Label3:
            MyArray(i, j) = DLookup("[PhaseID]", "Project", "[ProjectID] = " & a)
        ElseIf j = 5 Then
            For z = 1 To 3
                c = z
                If DLookup("[Hands]", "Hands", "[HandID] = " & z) = AEMUpdateArray(i, j) Then GoTo Label4
            Next
Label4:
        ElseIf j = 6 Then
            For z = 1 To 5
                If DLookup("[Deviation]", "Deviations", "[Car Deviation ID] = " & z) = AEMUpdateArray(i, j) Then GoTo Label5
            Next
Label5:
            MyArray(i, j) = z
        ElseIf j = 7 Then
            MyArray(i, j) = AEMUpdateArray(i, j)
        End If
    Next
    
    strCriteria = "INSERT INTO [" & MyTables(1) & "] ([A-Nr], [Drawing Date], [Description], [Car Deviation], [LIUMF], [ConfigurationID], [ProjectID], [ZGS], [Phase Start], [Phase Current/End])" & _
                "VALUES ('" & MyArray(i, 1) & "', '" & MyArray(i, 0) & "', '" & MyArray(i, 7) & "', '" & MyArray(i, 6) & "', '" & DLookup("[LIUMF]", "Configuration", "[ProjectID] = " & a & "And" & "[TypeID] = " & b & "And" & "[HandID] = " & c) & "', '" & DLookup("[ConfigID]", "Configuration", "[ProjectID] = " & a & "And" & "[TypeID] = " & b & "And" & "[HandID] = " & c) & "', '" & MyArray(i, 3) & "', '" & MyArray(i, 2) & "', '" & MyArray(i, 4) & "', '" & MyArray(i, 4) & "')"
    DoCmd.SetWarnings (WarningsOff)
    DoCmd.RunSQL strCriteria

'Check for Dublicates
Label1:
    If MyAnumber <> "" Then
        MyDescription = (Left(DLookup("[Description]", "New Modules Table", "[A-Nr] = '" & MyAnumber & "'"), 14) & "    " & Right(DLookup("[Description]", "New Modules Table", "[A-Nr] = '" & MyAnumber & "'"), Len(DLookup("[Description]", "New Modules Table", "[A-Nr] = '" & MyAnumber & "'")) - 15))
        MyLiumf = (Left(DLookup("[LIUMF]", "New Modules Table", "[A-Nr] = '" & MyAnumber & "'"), 4) & "%" & Right(DLookup("[LIUMF]", "New Modules Table", "[A-Nr] = '" & MyAnumber & "'"), 3))
        strCriteria = "UPDATE [" & MyTables(1) & "]" & _
                "SET [" & MyTables(1) & "].[Description]= '" & MyDescription & "'" & _
                "WHERE [" & MyTables(1) & "].[A-Nr]= '" & MyAnumber & "'"
        DoCmd.SetWarnings (WarningsOff)
        DoCmd.RunSQL strCriteria
        strCriteria = "UPDATE [" & MyTables(1) & "]" & _
                "SET [" & MyTables(1) & "].[LIUMF]= '" & MyLiumf & "'" & _
                "WHERE [" & MyTables(1) & "].[A-Nr]= '" & MyAnumber & "'"
        DoCmd.SetWarnings (WarningsOff)
        DoCmd.RunSQL strCriteria
        MyAnumber = ""
        k = k + 1
DoCmd.OpenTable MyTables(1)
    End If

Next



If n - k = 1 Then
    If k = 0 Then
        Response = MsgBox("" & n - k & " NEW module from the " & KickOffArray(2, y) & " was added to DataBase. Please insert index for him!", vbInformation)
    Else
        Response = MsgBox("" & n - k & " NEW LL/RL module from the " & KickOffArray(2, y) & " was added to DataBase. Please insert index for him!", vbInformation)
    End If
Else
    If k = 0 Then
        Response = MsgBox("" & n - k & " NEW modules from the " & KickOffArray(2, y) & " was added to DataBase. Please insert indexes for them!", vbInformation)
    Else
        Response = MsgBox("" & n - k & " NEW modules (" & k & " of them are LL/RL) from the " & KickOffArray(2, y) & " was added to DataBase. Please insert indexes for them!", vbInformation)
    End If
End If

End Function

