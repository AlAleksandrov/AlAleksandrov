VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Menu Form Archive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()

    Dim answer As Integer
    
    answer = MsgBox("Run Archive All Query?", vbQuestion + vbYesNo + vbDefaultButton2, "Archive All Query")

    If answer = vbYes Then
        DoCmd.OpenQuery "AEM Archive Query"
        DoCmd.OpenQuery "Index Archive Query"
        DoCmd.OpenQuery "Delete Archived Indexes"
        DoCmd.OpenQuery "Module Archive Query"
        DoCmd.OpenQuery "Module Persists In Index"
        DoCmd.OpenQuery "Delete Impersistent Modules"
    Else
        MsgBox "Action Canseled"
    End If

End Sub

Private Sub Command16_Click()

    Dim answer As Integer
    
    answer = MsgBox("Run Archive All Master Data Query?", vbQuestion + vbYesNo + vbDefaultButton2, "Archive Master Data")

    If answer = vbYes Then
        DoCmd.OpenQuery "Archive BWN Query"
        DoCmd.OpenQuery "Archive Foaming Query"
        DoCmd.OpenQuery "Archive OGC Query"
        DoCmd.OpenQuery "Archive PRG Query"
        DoCmd.OpenQuery "Archive RFA Query"
        DoCmd.OpenQuery "Archive Screwing Query"
    Else
        MsgBox "Action Canseled"
    End If

End Sub

Private Sub Command2_Click()
    
    Dim answer As Integer
    
    answer = MsgBox("Run AEM Archive Query?", vbQuestion + vbYesNo + vbDefaultButton2, "AEM Archive Query")

    If answer = vbYes Then
        DoCmd.OpenQuery "AEM Archive Query"
    Else
        MsgBox "Action Canseled"
    End If

End Sub

Private Sub Command5_Click()
    
    Dim answer As Integer
    
    answer = MsgBox("Run Index Archive Query?", vbQuestion + vbYesNo + vbDefaultButton2, "Index Archive Query")

    If answer = vbYes Then
        DoCmd.OpenQuery "Index Archive Query"
        DoCmd.OpenQuery "Delete Archived Indexes"
    Else
        MsgBox "Action Canseled"
    End If
    
End Sub

Private Sub Command6_Click()

    Dim answer As Integer
    
    answer = MsgBox("Run Module Archive Query?", vbQuestion + vbYesNo + vbDefaultButton2, "Module Archive Query")

    If answer = vbYes Then
        DoCmd.OpenQuery "Module Archive Query"
        DoCmd.OpenQuery "Module Persists In Index"
        DoCmd.OpenQuery "Delete Impersistent Modules"
    Else
        MsgBox "Action Canseled"
    End If

End Sub
