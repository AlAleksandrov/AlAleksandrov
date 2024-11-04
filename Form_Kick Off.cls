VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Kick Off"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub AEM_AfterUpdate()

Me.AEM_Update!AEM_Nr = AEM
Me.Requery

End Sub

Private Sub Combo140_AfterUpdate()
Me.Requery
End Sub

Private Sub Command146_Click()

DoCmd.RunSQL "DELETE * from [New Modules Table]"
DoEvents
DoCmd.RunSQL "DELETE * from [Kick Off]"
DoEvents
DoCmd.RunSQL "DELETE * from [AEM Update]"
DoEvents
DoCmd.RunSQL "DELETE * from [Connect AEMs to A-Nr]"
DoEvents
DoCmd.RunSQL "DELETE * from [New Indexes]"
DoEvents
DoCmd.RunSQL "DELETE * from [Status Update]"
DoEvents
DoCmd.RunSQL "DELETE * from [DBLB - tbl_Update_Weight_Date_etc]"
DoEvents
DoCmd.RunSQL "DELETE * from [tbl_AddNewData]"
DoEvents
DoCmd.RunSQL "DELETE * from [tbl_Update]"
DoEvents
DoCmd.RunSQL "DELETE * from [DBLB 03 - tbl_Deleted Modules]"
DoEvents

End Sub

Private Sub Form_Current()

If Me.Combo140 = "13" Then
   Me.Combo150.Enabled = True
   Me.Combo152.Enabled = True
ElseIf Me.Combo140 <> "13" Then
   Me.Combo150.Enabled = False
   Me.Combo152.Enabled = False
End If



End Sub
