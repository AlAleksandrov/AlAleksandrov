Attribute VB_Name = "08a-PRG_&_USW_Export"
Option Compare Database

Function Export_PRG_USS(project As String, Combined As Boolean)

Dim x As Integer, y As Long, z As Integer, x1 As Integer, y1 As Long, z1 As Long, last_row As Integer, last_row1 As Integer, MyProject As String, MyType As String, MyHand As String, MyStatus As String, MyPRGs(1000, 3), MyUSWs(1000, 3), MyDrawingPRGs(1000, 3), MyDrawingUSWs(1000, 3), MyTables(2) As Variant, MyPhasePRGs(1000, 3), MyPhaseUSWs(1000, 3) As String

Dim i As Integer, j As Integer, ProjectNumber As Integer, MySource(1000, 11), MyDest(1000, 11), MyLink As String, MyLinkSourse As String, MyLinkDest As String, intPath As Integer, a As Integer, b As Integer, c As Integer, d As Integer, CountResult As Integer
'Dim Project As String, Combined As Boolean
'Project = "V297"
'Combined = True


If project = "W206" Then ProjectNumber = 1
If project = "V297" Or project = "V295" Then ProjectNumber = 2
If project = "V297" Or project = "V295" Then Combined = True
If project = "X254" Then ProjectNumber = 3
If project = "C236" Then ProjectNumber = 4

MyTables(1) = "[PRG List Data]"
last_row = DCount("[ID]", MyTables(1))
MyLink = "\\leoni.local\dfsroot\BG1\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\01-Database Master Data\"

For y = 1 To last_row
    MyProject = DLookup("[Project]", MyTables(1), "[ID] = " & y)
    If project = MyProject Then
        MyStatus = DLookup("[Status]", MyTables(1), "[ID] = " & y)
        If MyStatus = "C" Or MyStatus = "N" Then
            MyType = DLookup("[Type]", MyTables(1), "[ID] = " & y)
            MyHand = DLookup("[Hand]", MyTables(1), "[ID] = " & y)
            If (MyType & " " & MyHand) = "COC LL" Then
                For x = 0 To 0
                    MyPRGs(z, x) = DLookup("[Files]", MyTables(1), "[ID] = " & y)
                    MyDrawingPRGs(z, x) = DLookup("[Drawing A-Nr+]", MyTables(1), "[ID] = " & y)
                    MyPhasePRGs(z, x) = DLookup("[Future Implementation]", MyTables(1), "[ID] = " & y)
                    If ProjectNumber = 1 Then
                        If Right(MyPhasePRGs(z, x), 3) = "SOP" Then
                            MyPhasePRGs(z, x) = "SOP"
                        Else
                            MyPhasePRGs(z, x) = "AJ"
                        End If
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyPhasePRGs(z, x) & "\" & MyType & " " & MyHand & "\" & MyDrawingPRGs(z, x) & "\ete\" & MyPRGs(z, x))
                        MySource(z, x) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\" & MyPRGs(z, x))
                        MyDest(z, x) = MyLinkDest
                    Else
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingPRGs(z, x) & "\ete\" & MyPRGs(z, x))
                        MySource(z, x) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\" & MyPRGs(z, x))
                        MyDest(z, x) = MyLinkDest
                    End If
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject)
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject)
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG")
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\PRG\")
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand)
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\")
'MsgBox (MySource(z, x))
'MsgBox (MyDest(z, x))
                    FileCopy MySource(z, x), MyDest(z, x)
                    z = z + 1
                Next
            ElseIf (MyType & " " & MyHand) = "COC RL" Then
                For x = 1 To 1
                    MyPRGs(z, x) = DLookup("[Files]", MyTables(1), "[ID] = " & y)
                    MyDrawingPRGs(z, x) = DLookup("[Drawing A-Nr+]", MyTables(1), "[ID] = " & y)
                    MyPhasePRGs(z, x) = DLookup("[Future Implementation]", MyTables(1), "[ID] = " & y)
                    If ProjectNumber = 1 Then
                        If Right(MyPhasePRGs(z, x), 3) = "SOP" Then
                            MyPhasePRGs(z, x) = "SOP"
                        Else
                            MyPhasePRGs(z, x) = "AJ"
                        End If
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyPhasePRGs(z, x) & "\" & MyType & " " & MyHand & "\" & MyDrawingPRGs(z, x) & "\ete\" & MyPRGs(z, x))
                        MySource(z, x) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\" & MyPRGs(z, x))
                        MyDest(z, x) = MyLinkDest
                    Else
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingPRGs(z, x) & "\ete\" & MyPRGs(z, x))
                        MySource(z, x) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\" & MyPRGs(z, x))
                        MyDest(z, x) = MyLinkDest
                    End If
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject)
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject)
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG")
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\PRG\")
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand)
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\")
'MsgBox (MySource(z, x))
'MsgBox (MyDest(z, x))
                    FileCopy MySource(z, x), MyDest(z, x)
                    z = z + 1
                Next
            ElseIf (MyType & " " & MyHand) = "MRA LL" Then
                For x = 2 To 2
                    MyPRGs(z, x) = DLookup("[Files]", MyTables(1), "[ID] = " & y)
                    MyDrawingPRGs(z, x) = DLookup("[Drawing A-Nr+]", MyTables(1), "[ID] = " & y)
                    MyPhasePRGs(z, x) = DLookup("[Future Implementation]", MyTables(1), "[ID] = " & y)
                    If ProjectNumber = 1 Then
                        If Right(MyPhasePRGs(z, x), 3) = "SOP" Then
                            MyPhasePRGs(z, x) = "SOP"
                            If Combined = True Then
                                MyDrawingPRGs(z, x) = "A_206_540_00_00"
                            Else
                                MyDrawingPRGs(z, x) = DLookup("[Drawing A-Nr+]", MyTables(1), "[ID] = " & y)
                            End If
                        Else
                            MyPhasePRGs(z, x) = "AJ"
                            If Combined = True Then
                                MyDrawingPRGs(z, x) = "A_206_540_00_02"
                            Else
                                MyDrawingPRGs(z, x) = DLookup("[Drawing A-Nr+]", MyTables(1), "[ID] = " & y)
                            End If
                        End If
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyPhasePRGs(z, x) & "\" & MyType & " " & MyHand & "\" & MyDrawingPRGs(z, x) & "\ete\" & MyPRGs(z, x))
                        MySource(z, x) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\" & MyPRGs(z, x))
                        MyDest(z, x) = MyLinkDest
                    ElseIf ProjectNumber = 2 Then
                        If MyDrawingPRGs(z, x) = "A_295_540_31_07" Then
                            If Combined = True Then
                                MyDrawingPRGs(z, x) = "MODULFREIGABE_V295"
                            End If
                        ElseIf MyDrawingPRGs(z, x) = "A_297_540_77_14" Then
                            If Combined = True Then
                                MyDrawingPRGs(z, x) = "MODULFREIGABE_V297"
                            End If
                        End If
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingPRGs(z, x) & "\ete\" & MyPRGs(z, x))
                        MySource(z, x) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\" & MyPRGs(z, x))
                        MyDest(z, x) = MyLinkDest
                    Else
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingPRGs(z, x) & "\ete\" & MyPRGs(z, x))
                        MySource(z, x) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\" & MyPRGs(z, x))
                        MyDest(z, x) = MyLinkDest
                    End If
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject)
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject)
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG")
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\PRG\")
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand)
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\")
'MsgBox (MySource(z, x))
'MsgBox (MyDest(z, x))
                    FileCopy MySource(z, x), MyDest(z, x)
                    z = z + 1
                Next
            ElseIf (MyType & " " & MyHand) = "MRA RL" Then
                For x = 3 To 3
                    MyPRGs(z, x) = DLookup("[Files]", MyTables(1), "[ID] = " & y)
                    MyDrawingPRGs(z, x) = DLookup("[Drawing A-Nr+]", MyTables(1), "[ID] = " & y)
                    MyPhasePRGs(z, x) = DLookup("[Future Implementation]", MyTables(1), "[ID] = " & y)
                    If ProjectNumber = 1 Then
                        If Right(MyPhasePRGs(z, x), 3) = "SOP" Then
                            MyPhasePRGs(z, x) = "SOP"
                            If Combined = True Then
                                MyDrawingPRGs(z, x) = "A_206_540_00_01"
                            Else
                                MyDrawingPRGs(z, x) = DLookup("[Drawing A-Nr+]", MyTables(1), "[ID] = " & y)
                            End If
                        Else
                            MyPhasePRGs(z, x) = "AJ"
                            If Combined = True Then
                                MyDrawingPRGs(z, x) = "A_206_540_00_03"
                            Else
                                MyDrawingPRGs(z, x) = DLookup("[Drawing A-Nr+]", MyTables(1), "[ID] = " & y)
                            End If
                        End If
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyPhasePRGs(z, x) & "\" & MyType & " " & MyHand & "\" & MyDrawingPRGs(z, x) & "\ete\" & MyPRGs(z, x))
                        MySource(z, x) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\" & MyPRGs(z, x))
                        MyDest(z, x) = MyLinkDest
                    ElseIf ProjectNumber = 2 Then
                        If MyDrawingPRGs(z, x) = "A_295_540_32_07" Then
                            If Combined = True Then
                                MyDrawingPRGs(z, x) = "MODULFREIGABE_V295"
                            End If
                        ElseIf MyDrawingPRGs(z, x) = "A_297_540_78_14" Then
                            If Combined = True Then
                                MyDrawingPRGs(z, x) = "MODULFREIGABE_V297"
                            End If
                        End If
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingPRGs(z, x) & "\ete\" & MyPRGs(z, x))
                        MySource(z, x) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\" & MyPRGs(z, x))
                        MyDest(z, x) = MyLinkDest
                    Else
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingPRGs(z, x) & "\ete\" & MyPRGs(z, x))
                        MySource(z, x) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\" & MyPRGs(z, x))
                        MyDest(z, x) = MyLinkDest
                    End If
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject)
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject)
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG")
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\PRG\")
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand)
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\PRG\" & MyType & " " & MyHand & "\")
'MsgBox (MySource(z, x))
'MsgBox (MyDest(z, x))
                    
                    FileCopy MySource(z, x), MyDest(z, x)
                    z = z + 1
                Next
            End If
        End If
    End If
            If (MyType & " " & MyHand) = "COC LL" Then
            a = z
            ElseIf (MyType & " " & MyHand) = "COC RL" Then
            b = z - a
            ElseIf (MyType & " " & MyHand) = "MRA LL" Then
            c = z - (b + a)
            ElseIf (MyType & " " & MyHand) = "MRA RL" Then
            d = z - (c + b + a)
            End If
Next
            CountResult = MsgBox("PRG FILES BY FOLDER" & vbCrLf & _
            "COC LL  =  " & a & vbCrLf & _
            "COC RL  =  " & b & vbCrLf & _
            "MRA LL  =  " & c & vbCrLf & _
            "MRA RL  =  " & d & vbCrLf & _
            "TOTAL    =  " & z & vbCrLf _
            , vbOKOnly + vbInformation, "PRG COUNT")

MyTables(2) = "[USW Modules]"
last_row1 = DCount("[ID]", MyTables(2))
For y1 = 1 To last_row1
    MyProject = DLookup("[Project]", MyTables(2), "[ID] = " & y1)
    If project = MyProject Then
        MyStatus = DLookup("[Status]", MyTables(2), "[ID] = " & y1)
        If MyStatus = "C" Or MyStatus = "N" Then
            MyType = DLookup("[Type]", MyTables(2), "[ID] = " & y1)
            MyHand = DLookup("[Hand]", MyTables(2), "[ID] = " & y1)
            If (MyType & " " & MyHand) = "COC LL" Then
                For x1 = 0 To 0
                    MyUSWs(z1, x1) = DLookup("[Files]", MyTables(2), "[ID] = " & y1)
                    MyDrawingUSWs(z1, x1) = DLookup("[Drawing A-Nr+]", MyTables(2), "[ID] = " & y1)
                    MyPhaseUSWs(z1, x1) = DLookup("[Future Implementation]", MyTables(2), "[ID] = " & y1)
                    If ProjectNumber = 1 Then
                        If Right(MyPhaseUSWs(z1, x1), 3) = "SOP" Then
                            MyPhaseUSWs(z1, x1) = "SOP"
                        Else
                            MyPhaseUSWs(z1, x1) = "AJ"
                        End If
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyPhaseUSWs(z1, x1) & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\usw\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 4) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 4) = MyLinkDest
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyPhaseUSWs(z1, x1) & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 8) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 8) = MyLinkDest
                    Else
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\usw\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 4) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 4) = MyLinkDest
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 8) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 8) = MyLinkDest
                    End If
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject)
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject)
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW")
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\USW\")
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES")
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\")
'MsgBox (MySource(z1, x1+4))
'MsgBox (MyDest(z1, x1+4))
                    FileCopy MySource(z1, x1 + 4), MyDest(z1, x1 + 4)
'MsgBox (MySource(z1, x1+8))
'MsgBox (MyDest(z1, x1+8))
                    FileCopy MySource(z1, x1 + 8), MyDest(z1, x1 + 8)
                    z1 = z1 + 1
                Next
            ElseIf (MyType & " " & MyHand) = "COC RL" Then
                For x1 = 1 To 1
                    MyUSWs(z1, x1) = DLookup("[Files]", MyTables(2), "[ID] = " & y1)
                    MyDrawingUSWs(z1, x1) = DLookup("[Drawing A-Nr+]", MyTables(2), "[ID] = " & y1)
                    MyPhaseUSWs(z1, x1) = DLookup("[Future Implementation]", MyTables(2), "[ID] = " & y1)
                    If ProjectNumber = 1 Then
                        If Right(MyPhaseUSWs(z1, x1), 3) = "SOP" Then
                            MyPhaseUSWs(z1, x1) = "SOP"
                        Else
                            MyPhaseUSWs(z1, x1) = "AJ"
                        End If
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyPhaseUSWs(z1, x1) & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\usw\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 4) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 4) = MyLinkDest
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyPhaseUSWs(z1, x1) & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 8) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 8) = MyLinkDest
                    Else
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\usw\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 4) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 4) = MyLinkDest
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 8) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 8) = MyLinkDest
                    End If
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject)
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject)
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW")
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\USW\")
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES")
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\")
'MsgBox (MySource(z1, x1 + 4))
'MsgBox (MyDest(z1, x1 + 4))
                    FileCopy MySource(z1, x1 + 4), MyDest(z1, x1 + 4)
'MsgBox (MySource(z1, x1 + 8))
'MsgBox (MyDest(z1, x1 + 8))
                    FileCopy MySource(z1, x1 + 8), MyDest(z1, x1 + 8)
                    z1 = z1 + 1
                Next
            ElseIf (MyType & " " & MyHand) = "MRA LL" Then
                For x1 = 2 To 2
                    MyUSWs(z1, x1) = DLookup("[Files]", MyTables(2), "[ID] = " & y1)
                    MyDrawingUSWs(z1, x1) = DLookup("[Drawing A-Nr+]", MyTables(2), "[ID] = " & y1)
                    MyPhaseUSWs(z1, x1) = DLookup("[Future Implementation]", MyTables(2), "[ID] = " & y1)
                    If ProjectNumber = 1 Then
                        If Right(MyPhaseUSWs(z1, x1), 3) = "SOP" Then
                            MyPhaseUSWs(z1, x1) = "SOP"
                            If Combined = True Then
                                MyDrawingUSWs(z1, x1) = "A_206_540_00_00"
                            Else
                                MyDrawingUSWs(z1, x1) = DLookup("[Drawing A-Nr+]", MyTables(2), "[ID] = " & y1)
                            End If
                        Else
                            MyPhaseUSWs(z1, x1) = "AJ"
                            If Combined = True Then
                                MyDrawingUSWs(z1, x1) = "A_206_540_00_02"
                            Else
                                MyDrawingUSWs(z1, x1) = DLookup("[Drawing A-Nr+]", MyTables(2), "[ID] = " & y1)
                            End If
                        End If
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyPhaseUSWs(z1, x1) & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\usw\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 4) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 4) = MyLinkDest
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyPhaseUSWs(z1, x1) & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 8) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 8) = MyLinkDest
                    ElseIf ProjectNumber = 2 Then
                        If MyDrawingUSWs(z1, x1) = "A_295_540_31_07" Then
                            If Combined = True Then
                                MyDrawingUSWs(z1, x1) = "MODULFREIGABE_V295"
                            End If
                        ElseIf MyDrawingUSWs(z1, x1) = "A_297_540_77_14" Then
                            If Combined = True Then
                                MyDrawingUSWs(z1, x1) = "MODULFREIGABE_V297"
                            End If
                        End If
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\usw\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 4) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 4) = MyLinkDest
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 8) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 8) = MyLinkDest
                    Else
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\usw\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 4) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 4) = MyLinkDest
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 8) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 8) = MyLinkDest
                    End If
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject)
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject)
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW")
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\USW\")
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES")
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\")
'MsgBox (MySource(z1, x1 + 4))
'MsgBox (MyDest(z1, x1 + 4))
                    FileCopy MySource(z1, x1 + 4), MyDest(z1, x1 + 4)
'MsgBox (MySource(z1, x1+8))
'MsgBox (MyDest(z1, x1+8))
                    FileCopy MySource(z1, x1 + 8), MyDest(z1, x1 + 8)
                    z1 = z1 + 1
                Next
            ElseIf (MyType & " " & MyHand) = "MRA RL" Then
                For x1 = 3 To 3
                    MyUSWs(z1, x1) = DLookup("[Files]", MyTables(2), "[ID] = " & y1)
                    MyDrawingUSWs(z1, x1) = DLookup("[Drawing A-Nr+]", MyTables(2), "[ID] = " & y1)
                    MyPhaseUSWs(z1, x1) = DLookup("[Future Implementation]", MyTables(2), "[ID] = " & y1)
                    If ProjectNumber = 1 Then
                        If Right(MyPhaseUSWs(z1, x1), 3) = "SOP" Then
                            MyPhaseUSWs(z1, x1) = "SOP"
                            If Combined = True Then
                                MyDrawingUSWs(z1, x1) = "A_206_540_00_01"
                            Else
                                MyDrawingUSWs(z1, x1) = DLookup("[Drawing A-Nr+]", MyTables(2), "[ID] = " & y1)
                            End If
                        Else
                            MyPhaseUSWs(z1, x1) = "AJ"
                            If Combined = True Then
                                MyDrawingUSWs(z1, x1) = "A_206_540_00_03"
                            Else
                                MyDrawingUSWs(z1, x1) = DLookup("[Drawing A-Nr+]", MyTables(2), "[ID] = " & y1)
                            End If
                        End If
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyPhaseUSWs(z1, x1) & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\usw\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 4) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 4) = MyLinkDest
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyPhaseUSWs(z1, x1) & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 8) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 8) = MyLinkDest
                    ElseIf ProjectNumber = 2 Then
                        If MyDrawingUSWs(z1, x1) = "A_295_540_32_07" Then
                            If Combined = True Then
                                MyDrawingUSWs(z1, x1) = "MODULFREIGABE_V295"
                            End If
                        ElseIf MyDrawingUSWs(z1, x1) = "A_297_540_78_14" Then
                            If Combined = True Then
                                MyDrawingUSWs(z1, x1) = "MODULFREIGABE_V297"
                            End If
                        End If
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\usw\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 4) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 4) = MyLinkDest
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 8) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 8) = MyLinkDest
                    Else
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\usw\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 4) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 4) = MyLinkDest
                        MyLinkSourse = (MyLink & "06-Import\01 - Export_from_Drawings\02-KBL\" & MyProject & "\" & MyType & " " & MyHand & "\" & MyDrawingUSWs(z1, x1) & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MySource(z1, x1 + 8) = MyLinkSourse
                        MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\" & MyUSWs(z1, x1))
                        MyDest(z1, x1 + 8) = MyLinkDest
                    End If
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject)
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject)
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW")
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\USW\")
                    MyLinkDest = (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES")
                    intPath = Len(Dir(MyLinkDest, vbDirectory))
                    If intPath = 0 Then MkDir (MyLink & "08-Upload\" & MyProject & "\USW_SPLICES\")
'MsgBox (MySource(z1, x1 + 4))
'MsgBox (MyDest(z1, x1 + 4))
                    FileCopy MySource(z1, x1 + 4), MyDest(z1, x1 + 4)
'MsgBox (MySource(z1, x1+8))
'MsgBox (MyDest(z1, x1+8))
                    FileCopy MySource(z1, x1 + 8), MyDest(z1, x1 + 8)
                    z1 = z1 + 1
                Next
            End If
        End If
    End If
Next

            CountResult = MsgBox("NUMBER OF USW SPLICE FILES" & vbCrLf & _
            "TOTAL    =  " & z1 & vbCrLf _
            , vbOKOnly + vbInformation, "USS COUNT")

DoCmd.Close acForm, "frm_select_project_PRG_USS_Export", acSaveYes
MsgBox "Action Completed"
    
End Function
