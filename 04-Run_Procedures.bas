Attribute VB_Name = "04-Run_Procedures"
Option Compare Database

Function Run_Q4()
    
Dim answer As Integer
    answer = MsgBox("Run Q4 Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "Q4 Process")
    
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    
    If answer = vbYes Then
        dbs.Execute "Q4 Main Query"
        dbs.Execute "Q4 Line Set-up Query"
        dbs.Execute "Append Cliplist to QDE Final"
        dbs.Execute "Number of Wires Query"
        dbs.Execute "Update Index - Number of Wires"
        dbs.Execute "Q4 OGC & Screwing Update"
        
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If

End Function

Function Run_Master_Data()
    
    Dim answer As Integer
    answer = MsgBox("Run Master Data Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "Master Data Process")
    
    Dim dbs As DAO.Database
    Set dbs = CurrentDb

    If answer = vbYes Then
        'BWN
        dbs.Execute "Step 01 - BRW Modules All Query"
        dbs.Execute "BRW Black List Mark"
        dbs.Execute "BRW Delete Marked"
        
        'OGC
        dbs.Execute "OGC ClipsList Query"
        dbs.Execute "OGC White List Append"
        dbs.Execute "OGC LED Update Query"
        dbs.Execute "OGC Blacklist Mark Query"
        dbs.Execute "OGC Black List Delete Query"
        dbs.Execute "OGC LED Special Request"
        
        'PRG
        dbs.Execute "ALTER TABLE [PRG List Data] ALTER COLUMN [ID] COUNTER(1,1)"
        dbs.Execute "PRG List Query"
        dbs.Execute "PRG with Clips"
        
        'RFA
        dbs.Execute "Fuses Query"
        dbs.Execute "Relays Query"
        
        'Foaming
        dbs.Execute "Foaming Query"
        
        'Screwing
        dbs.Execute "Filtered Screwing Data To Screwing"
        
        'USW
        dbs.Execute "ALTER TABLE [USW Modules] ALTER COLUMN [ID] COUNTER(1,1)"
        dbs.Execute "USW Modules Query"
        dbs.Execute "USW Splice Query"
        
        'Picking
        dbs.Execute "MSG Basic Modules w/o Wirelist Append Query"
        dbs.Execute "MSG Basic Modules w Wirelist Data Append Query"
        dbs.Execute "MSG Special Request Query"
        dbs.Execute "MSG Special Request Query1"
        
    
        MsgBox "Action Completed!"
        
    Else
        MsgBox "Action Canseled"
    End If
        
End Function


Function Q4_Indevidual_Change()

Dim answer As Integer
    answer = MsgBox("Run Indevidual Q4 Process Update?", vbQuestion + vbYesNo + vbDefaultButton2, "Q4 Process Update")

    If answer = vbYes Then
        DoCmd.OpenForm ("frm_select_project_Q4")
    Else
        MsgBox "Action Canseled"
    End If

End Function

Function Run_Q4_Indevidual(project As String)
    
    Dim answer As Integer
        answer = MsgBox("Run Q4 Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "Q4 Process")
    
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    
    If answer = vbYes Then
        dbs.Execute "Q4 Main Query"
        dbs.Execute "Q4 Line Set-up Query"
        dbs.Execute "Append Cliplist to QDE Final"
        dbs.Execute "Number of Wires Query"
        dbs.Execute "Update Index - Number of Wires"
        dbs.Execute "Q4 OGC & Screwing Update"
        
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If

End Function

Function Special_Indevidual_Change()

Dim answer As Integer
    answer = MsgBox("Run Indevidual Special Process Update?", vbQuestion + vbYesNo + vbDefaultButton2, "Special Process Update")

    If answer = vbYes Then
        DoCmd.OpenForm ("frm_select_project_Special")
    Else
        MsgBox "Action Canseled"
    End If

End Function


Function Run_Master_Data_Indevidual(project As String)
    
    Dim answer As Integer, dbs As DAO.Database
    Set dbs = CurrentDb
    
    answer = MsgBox("Run Master Data Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "Master Data Process")
    

    If answer = vbYes Then
        'BWN
        dbs.Execute "BRW Modules ALL"
        
        'OGC
        dbs.Execute "OGC ClipsList Query"
        dbs.Execute "OGC White list Append"
        dbs.Execute "OGC LED Update Query"
        dbs.Execute "OGC Blacklist Mark Query"
        dbs.Execute "OGC Black List Delete Query"
        
        'PRG
        dbs.Execute "PRG List Query"
        dbs.Execute "PRG with Clips"
        
        'RFA
        dbs.Execute "Fuses Query"
        dbs.Execute "Relays Query"
        
        'Foaming
        dbs.Execute "Foaming Query"
        
        'Screwing
        dbs.Execute "Filtered Screwing Data To Screwing"

    Else
        MsgBox "Action Canseled"
    End If

End Function

Function Run_Output_Tables()

    Dim answer As Integer, dbs As DAO.Database
    Set dbs = CurrentDb
    
    answer = MsgBox("Run Output Table Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "Master Data Process")
    

    If answer = vbYes Then
    
        dbs.Execute "All Used Modules New Index"
        dbs.Execute "All Used Modules Old Index"
        dbs.Execute "Index Package"
        dbs.Execute "SAP Transaction"
        dbs.Execute "AEM Package"
        dbs.Execute "QGate Query"
        dbs.Execute "QGzone"
        dbs.Execute "KDMAT"
        dbs.Execute "MD Revise Output"
        dbs.Execute "MD Revise Update Predecessor Part"
        dbs.Execute "MFG Request"
        dbs.Execute "ROS"
        dbs.Execute "Projects Weight"
        dbs.Execute "Module Time Append Query"
        
        MsgBox "Action Completed!"

    Else
        MsgBox "Action Canseled"
    End If


End Function

Function Run_Master_Data_Archive_To_Active()
    
    Dim answer As Integer
    answer = MsgBox("Run MD Archive to Active Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "Master Data Process")
    
    Dim dbs As DAO.Database
    Set dbs = CurrentDb

    If answer = vbYes Then
        
        'BWN
        dbs.Execute "Update BWN Status Archive"
        dbs.Execute "BRW Archive to Active MD"
        
        'OGC
        dbs.Execute "Update OGC Status Archive"
        dbs.Execute "OGC Archive to Active MD"

        
        'PRG
        dbs.Execute "Update PRG Status Archive"
        dbs.Execute "PRG Archive to Active MD"
        
        'RFA
        dbs.Execute "Update RFA Status Archive"
        dbs.Execute "RFA Archive to Active MD"
        
        'Foaming
        dbs.Execute "Update Foaming Status Archive"
        dbs.Execute "Foaming Archive to Active MD"
        
        'Screwing
        dbs.Execute "Update Screwing Status Archive"
        dbs.Execute "Screwing Archive to Active MD"
        
        'USW
        dbs.Execute "Update USW Modules Status Archive"
        dbs.Execute "USW Modules Archive to Active MD"
        dbs.Execute "Update USW Splice Status Archive"
        dbs.Execute "USW Splice Archive to Active MD"

        'Picking
        
        dbs.Execute "Update MSG Basic Modules Status Archive"
        dbs.Execute "MSG Basic Modules Archive to Active MD"
        dbs.Execute "Update MSG Basic Modules+PIN+Kabel Status Archive"
        dbs.Execute "MSG Basic Modules+PIN+Kabel Archive to Active MD"
        
        MsgBox "Action Completed!"
        
    Else
        MsgBox "Action Canseled"
    End If
        
End Function


Function Run_Wirelist_Plus()
    
    Dim answer As Integer
    answer = MsgBox("Run Wirelist Plus Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "PE Process")
    
    Dim dbs As DAO.Database
    Set dbs = CurrentDb

    If answer = vbYes Then
    
        dbs.Execute "ALTER TABLE [Wirelist+] ALTER COLUMN [WirelistID] COUNTER(1,1)"
        dbs.Execute "ALTER TABLE [Wirelist++] ALTER COLUMN [WirelistID] COUNTER(1,1)"
        dbs.Execute "Wirelist+ Append Query"
        dbs.Execute "Wirelist+ BWN1 Query"
        dbs.Execute "Wirelist+ BWN2 Query"
        dbs.Execute "Wirelist+ USS1 Query"
        dbs.Execute "Wirelist+ USS2 Query"
        dbs.Execute "Update WireList PartNo"
        dbs.Execute "Update WireList PartNo V295"
        dbs.Execute "Update WireList PartNo C236"
        dbs.Execute "Wirelist+ Strip Lenght1"
        dbs.Execute "Wirelist+ Strip Lenght2"
        dbs.Execute "Wirelist+ to Wirelist++"
        dbs.Execute "Copy Wirelist+"
        dbs.Execute "Copy Wirelist++"
        dbs.Execute "Copy Wirelist+ to Wirelist+ - Antennas"
        dbs.Execute "Wirelist+ Add Z Xcode"
        dbs.Execute "Wirelist+ Add Z Xcode1"
        
        Call Sorting_Variant_Wirelist
        
        MsgBox "Action Completed!"
        
    Else
        MsgBox "Action Canseled"
    End If
        
End Function

Function Run_Wirelist_Plus_Plus()
    
    Dim answer As Integer
    answer = MsgBox("Run Wirelist ++ Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "PE Process")
    
    Dim dbs As DAO.Database
    Set dbs = CurrentDb

    If answer = vbYes Then
    
        Call Index_Order_Wirelist
        
        MsgBox "Action Completed!"
        
    Else
        MsgBox "Action Canseled"
    End If
        
End Function

Function Run_Autokan()
    
    Dim answer As Integer
    answer = MsgBox("Run Autokan Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "PE Process")
    
    Dim dbs As DAO.Database
    Set dbs = CurrentDb

    If answer = vbYes Then

        'dbs.Execute "Wirelist+ Append Query"
        'dbs.Execute "Wirelist+ BWN1 Query"
        'dbs.Execute "Wirelist+ BWN2 Query"
        'dbs.Execute "Wirelist+ USS1 Query"
        'dbs.Execute "Wirelist+ USS2 Query"
        'dbs.Execute "Update WireList PartNo"
        'dbs.Execute "Update WireList PartNo V295"
        'dbs.Execute "Update WireList PartNo C236"
        'dbs.Execute "Wirelist+ Strip Lenght1"
        'dbs.Execute "Wirelist+ Strip Lenght2"
        'dbs.Execute "Copy Wirelist+"
        'dbs.Execute "Wirelist+ Add Z Xcode"
        'dbs.Execute "Wirelist+ Add Z Xcode1"
        
        MsgBox "Action Completed!"
        
    Else
        MsgBox "Action Canseled"
    End If
        
End Function

Function Run_AEM_TO_TXT()
    
Dim answer As Integer
    answer = MsgBox("Run AEM TO TXT Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "AEM Process")
    
    If answer = vbYes Then
        Call AEM_Preparation_PDF_to_TXT
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If

End Function

Function Run_AEM_TXT_TO_DB()
    
Dim answer As Integer
    answer = MsgBox("Run AEM TXT TO DB IMPORT Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "AEM Process")

        If answer = vbYes Then
        DoCmd.Close acForm, "Kick Off", acSaveYes
        Call AEM_Import_TXT_to_Access_DBMS
        
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If

End Function

Function Run_AEM_AUTO_IMPORT()
    
Dim answer As Integer
    answer = MsgBox("Run AEM AUTO IMPORT Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "AEM Process")
        
    If answer = vbYes Then
        Call AEM_Preparation_PDF_to_TXT
        Call AEM_Import_TXT_to_Access_DBMS
        
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If

End Function

Function Run_MD_Export_PRG_USS()

Dim answer As Integer
    answer = MsgBox("Run PRG & USS EXPORT Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "PRG & USS")

    If answer = vbYes Then
        DoCmd.OpenForm ("frm_select_project_PRG_USS_Export")
    Else
        MsgBox "Action Canseled"
    End If

End Function

Function Run_MD_Check_USS()

Dim answer As Integer
    answer = MsgBox("Run USS CHECK Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "USS")

    Dim dbs As DAO.Database
    Set dbs = CurrentDb

    If answer = vbYes Then
        Call Check_USW_Xcode
        dbs.Execute "USW Sprices Homologation"
        dbs.Execute "USW Sprices Homologation 2"
        
    Else
        MsgBox "Action Canseled"
    End If

End Function
