VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_select_project_delete_import_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim project As String

Private Sub ChangeIndexOK_Click()

project = Me.Combo_Select_Project

Call Index_State_Update(project)
End Sub

