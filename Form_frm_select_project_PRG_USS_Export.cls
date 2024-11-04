VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_select_project_PRG_USS_Export"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim project As String, Combined As Boolean

Private Sub Combo_Select_Project_AfterUpdate()

Me.Requery

    If Me.Combo_Select_Project = "W206" Then

        Me.Combined_Drawing.Enabled = True
   
    ElseIf Me.Combo_Select_Project <> "W206" Then
   
        Me.Combined_Drawing.Enabled = False

    End If
    
Me.Requery

End Sub

Private Sub Form_Current()

    If Me.Combo_Select_Project = "W206" Then

        Me.Combined_Drawing.Enabled = True
   
    ElseIf Me.Combo_Select_Project <> "W206" Then
   
        Me.Combined_Drawing.Enabled = False

    End If

End Sub

Private Sub ExportPrgUssOK_Click()

    project = Me.Combo_Select_Project

    Combined = Me.Combined_Drawing

    Call Export_PRG_USS(project, Combined)

End Sub
