Attribute VB_Name = "06-Wirelist+"
Option Compare Database

Function Sorting_Variant_Wirelist()

Dim x As Integer, y As Long, last_row As Integer, MyVariant As String, MyProject As String, MySortingVariant As String, MyTables(1) As Variant

Dim strCriteria As String
    
MyTables(1) = "Wirelist+"
last_row = DCount("[WirelistID]", "[Wirelist+]")

For y = 1 To last_row
    MyProject = DLookup("[Project]", "Wirelist+", "[WirelistID] = " & y)
    If Left(MyProject, 3) = "V29" Then
        MyVariant = DLookup("[Variant]", "Wirelist+", "[WirelistID] = " & y)
        MyVariant = Left(MyVariant, Len(MyVariant) - 2)
    Else
        MyVariant = DLookup("[Variant]", "Wirelist+", "[WirelistID] = " & y)
    End If
    For x = 1 To 10
        If Mid(MyVariant, 1, 1) = "V" Then
            MyString = Right(MyVariant, Len(MyVariant) - 1)
            If IsNumeric(MyString) = True Then
                If Len(MyVariant) - 1 = 5 Then
                    MySortingVariant = ("" & MyString & "")
                    GoTo Label2
                ElseIf Len(MyVariant) - 1 = 4 Then
                    MySortingVariant = ("" & MyString & "")
                ElseIf Len(MyVariant) - 1 = 3 Then
                    MySortingVariant = (0 & "" & MyString & "")
                ElseIf Len(MyVariant) - 1 = 2 Then
                    MySortingVariant = (0 & "" & MyString & "")
                    MySortingVariant = (0 & "" & MySortingVariant & "")
                ElseIf Len(MyVariant) - 1 = 1 Then
                    MySortingVariant = (0 & "" & MyString & "")
                    MySortingVariant = (0 & "" & MySortingVariant & "")
                    MySortingVariant = (0 & "" & MySortingVariant & "")
                End If
            Else
                If Len(MyVariant) - 1 = 5 Then
                    MySortingVariant = ("" & MyString & "")
                    GoTo Label2
                ElseIf Len(MyVariant) - 1 = 4 Then
                    MyString = Mid(MyVariant, 2, Len(MyVariant) - 2)
                    MySortingVariant = (0 & "" & MyString & "")
                    MySortingVariant = ("" & MySortingVariant & "" & Right(MyVariant, 1))
                ElseIf Len(MyVariant) - 1 = 3 Then
                    MyString = Mid(MyVariant, 2, Len(MyVariant) - 2)
                    MySortingVariant = (0 & "" & MyString & "")
                    MySortingVariant = (0 & "" & MySortingVariant & "")
                    MySortingVariant = ("" & MySortingVariant & "" & Right(MyVariant, 1))
                ElseIf Len(MyVariant) - 1 = 2 Then
                    MyString = Mid(MyVariant, 2, Len(MyVariant) - 2)
                    MySortingVariant = (0 & "" & MyString & "")
                    MySortingVariant = (0 & "" & MySortingVariant & "")
                    MySortingVariant = (0 & "" & MySortingVariant & "")
                    MySortingVariant = ("" & MySortingVariant & "" & Right(MyVariant, 1))
                End If
            End If
Label2:
            If Left(MyProject, 3) = "V29" Then
                MySortingVariant = ("V" & "" & MySortingVariant & "" & Right(DLookup("[Variant]", "Wirelist+", "[WirelistID] = " & y), 2))
            Else
                MySortingVariant = ("V" & "" & MySortingVariant & "")
            End If
                x = 10
        End If
    Next
Label1:
    strCriteria = "UPDATE [Wirelist+]" & _
                "SET [Wirelist+].[SortingVariant]= '" & MySortingVariant & "'" & _
                "WHERE [Wirelist+].[WirelistID]= " & y & ""
    DoCmd.SetWarnings (WarningsOff)
    DoCmd.RunSQL strCriteria
Next
End Function

