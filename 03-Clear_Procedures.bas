Attribute VB_Name = "03-Clear_Procedures"
Option Compare Database

Function Delete_Master_Data_Archive()

    Dim answer As Integer
        answer = MsgBox("Confirm Delete Master Data Archive?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Data")
    If answer = vbYes Then
        DoCmd.RunSQL "DELETE * from [Master Data Archive BRW]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Master Data Archive Foaming]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Master Data Archive OGC]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Master Data Archive PRG]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Master Data Archive RFA]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Master Data Archive Screwing]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Master Data Archive USW Splice]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Master Data Archive USW Modules]"
        DoEvents
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If
DoCmd.SetWarnings (WarningsOff)

End Function

Function Delete_Archive()

    Dim answer As Integer
        answer = MsgBox("Confirm Delete Archive?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Data")
    If answer = vbYes Then
        DoCmd.RunSQL "DELETE * from [Archive AEM]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Archive Index]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Archive Modules]"
        DoEvents
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If
DoCmd.SetWarnings (WarningsOff)

End Function

Function Delete_Q4()

    Dim answer As Integer
        answer = MsgBox("Confirm Delete Q4?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Data")
    If answer = vbYes Then
        DoCmd.RunSQL "DELETE * from [Q4]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [QDE Final]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Number of Wires]"
        DoEvents
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If
DoCmd.SetWarnings (WarningsOff)

End Function


Function Delete_Special_Process()

    Dim answer As Integer
        answer = MsgBox("Confirm Delete Special Process?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Data")

    If answer = vbYes Then
        DoCmd.RunSQL "DELETE * from [BRW Modules All Table]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Foaming-LL-RL]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [OGC-LL-RL]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [PRG List Data]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [RFA]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Screwing]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [USW Splice]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [USW Modules]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [MSG Basic Modules]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [MSG Basic Modules+PIN+Kabel]"
        DoEvents

        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If
DoCmd.SetWarnings (WarningsOff)

End Function

Function Delete_Output_Tables()

    Dim answer As Integer
        answer = MsgBox("Confirm Delete Output Tables?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Data")

    If answer = vbYes Then
        DoCmd.RunSQL "DELETE * from [All Used Modules]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Index Package Output Table]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [SAP Tansaction Output Table]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [AEM Package Output Table]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [QGate Output Table]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [QGzone Output Table]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [KDMAT Output Table]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [MD Revise Output Table]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [MFG Request Output Table]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [ROS Output Table]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Module Time Output Table]"
        DoEvents
        DoCmd.RunSQL "DELETE * from [Project Weight Output Table]"
        DoEvents
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If
DoCmd.SetWarnings (WarningsOff)


End Function



Function Delete_Confirm()

    Dim answer As Integer
        answer = MsgBox("Confirm Delete?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Data")
    
    DoCmd.SetWarnings (WarningsOff)

End Function


Function Delete_Wirelist_Plus_and_Plus_Plus()

    Dim answer As Integer
        answer = MsgBox("Confirm Delete Wirelists?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Data")

    If answer = vbYes Then
        DoCmd.RunSQL "DELETE * from [Wirelist+]"
        DoCmd.RunSQL "DELETE * from [Wirelist++]"
        DoCmd.RunSQL "DELETE * from [Wirelist+ no z]"
        DoCmd.RunSQL "DELETE * from [Wirelist++ Autokan]"
        DoCmd.RunSQL "DELETE * from [Wirelist+ -Antennas]"
        DoEvents
        
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If
DoCmd.SetWarnings (WarningsOff)


End Function

Function Delete_USW_SPLICES_CHECK()

    Dim answer As Integer
        answer = MsgBox("Confirm CLEAR USW_SPLICES CHECK TABLE?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Data")

    If answer = vbYes Then
        DoCmd.RunSQL "DELETE * from [USW_SPLICES]"
        DoEvents
        
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If
DoCmd.SetWarnings (WarningsOff)


End Function
