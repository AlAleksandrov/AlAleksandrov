VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_AEM Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command24_Click()

Me.List15.SetFocus
DoCmd.RunCommand acCmdEditListItems
Text188 = Me.List15.ListCount

End Sub

Private Sub Command28_Click()

Me.List9.SetFocus
DoCmd.RunCommand acCmdEditListItems
Text108 = Me.List9.ListCount

End Sub

Sub Command39_Click()

Call Search

End Sub

Function Search()

Dim MyArray(1000, 4), MyAnumbers As Variant, MyZGS As Variant, MyAnumbersText As String, strCriteria As String

Dim i As Integer, j As Integer, a As Integer

Text108 = Me.List9.ListCount
Text188 = Me.List15.ListCount
a = Text108

For i = 0 To a
    If IsNull(List9.Column(0, i)) Then
        GoTo Label1
    Else
        MyArray(i, 0) = List9.Column(0, i)
        MyArray(i, 1) = List15.Column(0, i)
        Me.List9.Selected(i) = True
        Me.List15.Selected(i) = True
        MyAnumbers = MyArray(i, 0)
        MyAnumbersText = "" & MyAnumbers & ""
        Me.A_NumberField = MyAnumbers
        MyZGS = MyArray(i, 1)
        Me.ZGS = MyZGS
    End If
    Call SaveQuote_Click
    MyArray(i, 3) = DCount("[A_Nomer]", "AEM Update", "[A_Nomer] = '" & MyAnumbersText & "'")
    Me.Text55 = MyArray(i, 3)
    Call SaveQuote_Click
    For j = 0 To 0
        If DMax("[No of Increased Index]", "AEM Update", "[A_Nomer] = '" & MyAnumbersText & "'") <> DLookup("[No of AEM's]", "AEM Update", "[A_Nomer] = '" & MyAnumbersText & "'") Then
            strCriteria = "UPDATE [AEM Update]" & _
                          "SET [AEM Update].[No of AEM's]= '" & DMax("[No of Increased Index]", "AEM Update", "[A_Nomer] = '" & MyAnumbersText & "'") & "'" & _
                          "WHERE ([AEM Update].[A_Nomer] = '" & MyAnumbersText & "')"
            DoCmd.SetWarnings (WarningsOff)
            DoCmd.RunSQL strCriteria
        End If
    Next j
    
    DoCmd.GoToRecord , "", acNewRec
    
Label1:
Next i

Me.List9.RowSource = ""
Me.List15.RowSource = ""
Text108 = Me.List9.ListCount
Text188 = Me.List15.ListCount

End Function

Private Sub SaveQuote_Click()

If Me.Dirty = True Then
Me.Dirty = False
End If

End Sub


