VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_select_project_archive_indevidual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'MD Archive Indevidual form variable'

Dim project As String

Private Sub Archive_IndevidualOK_Click()

project = Me.Combo_Select_Project_Archive_Indevidual

Call MD_Archive_Indevidual_Update(project)
End Sub

