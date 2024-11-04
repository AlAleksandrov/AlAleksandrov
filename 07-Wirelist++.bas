Attribute VB_Name = "07-Wirelist++"
Option Compare Database

Function Index_Order_Wirelist()

Dim x As Integer, y As Long, z As Integer, x1 As Integer, y1 As Long, z1 As Long, last_row As Integer, last_row1 As Integer, MyVariant As String, MyVariants As String, MyWire As String, MyProject As String, MyType As String, MyHand As String, MyIndexes(), MyTables(1) As Variant, MyColumns(), MySortingWire As String, MyRange As Integer, MyAnumber As String

Dim strCriteria As String
    
MyTables(1) = "[Wirelist++]"
MyColumns() = Array("[Index-Nr]", "[Index-Nr2]", "[Index-Nr3]", "[Index-Nr4]", "[Index-Nr5]", "[Index-Nr6]", "[Index-Nr7]", "[Index-Nr8]", "[Index-Nr9]", "[Index-Nr10]", "[Index-Nr11]", "[Index-Nr12]", "[Index-Nr13]", "[Index-Nr14]", "[Index-Nr15]", "[Index-Nr16]", "[Index-Nr17]", "[Index-Nr18]", "[Index-Nr19]", "[Index-Nr20]", "[Index-Nr21]", "[Index-Nr22]", "[Index-Nr23]", "[Index-Nr24]", "[Index-Nr25]")
MyIndexes() = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
last_row = DCount("[WirelistID]", MyTables(1))

For y = 1 To last_row
    MyWire = DLookup("[WireNumber]", MyTables(1), "[WirelistID] = " & y)
    MyVariant = DLookup("[Variants]", MyTables(1), "[WirelistID] = " & y)
    MyProject = DLookup("[Project]", MyTables(1), "[WirelistID] = " & y)
    MyType = DLookup("[Type]", MyTables(1), "[WirelistID] = " & y)
    MyHand = DLookup("[Hand]", MyTables(1), "[WirelistID] = " & y)
    MySortingWire = DLookup("[SortingWire]", MyTables(1), "[WirelistID] = " & y)
    MyRange = DCount("[SortingWire]", MyTables(1), "[SortingWire] = '" & MySortingWire & "'")
    If y >= 2 Then
        If MyWire = DLookup("[WireNumber]", MyTables(1), "[WirelistID] = " & y - 1) And MyProject = DLookup("[Project]", MyTables(1), "[WirelistID] = " & y - 1) And MyType = DLookup("[Type]", MyTables(1), "[WirelistID] = " & y - 1) And MyHand = DLookup("[Hand]", MyTables(1), "[WirelistID] = " & y - 1) Then
            z1 = z1 + 1
        Else
            GoTo Label3
        End If
    Else
        GoTo Label1
    End If
Label3:
    For x = 0 To 24
        If IsNull(DLookup(MyColumns(x), MyTables(1), "[WirelistID] = " & y)) Then
            x = 24
            GoTo Label2
        Else
            If z1 = 0 Then
                z = x
                MyIndexes(z) = DLookup(MyColumns(z), MyTables(1), "[WirelistID] = " & y)
            Else
                z = x + z1
                MyIndexes(z) = DLookup(MyColumns(x), MyTables(1), "[WirelistID] = " & y)
            End If
        End If
Label2:
    Next
    If z = 0 Then
        GoTo Label1
    Else
        If z1 = MyRange - 1 Then
            For x1 = 1 To z
                y1 = y - z1
                strCriteria = "UPDATE " & MyTables(1) & "" & _
                            "SET " & MyTables(1) & "." & MyColumns(x1) & "= '" & MyIndexes(x1) & "'" & _
                            "WHERE " & MyTables(1) & ".[WirelistID]= " & y1 & ""
                DoCmd.SetWarnings (WarningsOff)
                DoCmd.RunSQL strCriteria
                MyVariants = (DLookup("[Variants]", MyTables(1), "[WirelistID] = " & y1) & DLookup("[Variants]", MyTables(1), "[WirelistID] = " & y1 + x1))
                strCriteria = "UPDATE " & MyTables(1) & "" & _
                            "SET " & MyTables(1) & ".[Variants]= '" & MyVariants & "'" & _
                            "WHERE " & MyTables(1) & ".[WirelistID]= " & y1 & ""
                DoCmd.RunSQL strCriteria
                strCriteria = "DELETE FROM " & MyTables(1) & "" & _
                            "WHERE " & MyTables(1) & ".[WirelistID]= " & y1 + x1 & ""
                DoCmd.RunSQL strCriteria
                DoCmd.SetWarnings (WarningsOn)
            Next
            MyIndexes() = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            z1 = 0
        End If
    End If
Label1:
Next
End Function


