Attribute VB_Name = "05-Update_Procedures"
Option Compare Database
Function Persists_In_MasterData_Update()

Dim answer As Integer
    answer = MsgBox("Run Master Data Persistance Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "Master Data Persistance")


    If answer = vbYes Then
        dbs.Execute "Master Data Reset Query"
        dbs.Execute "Persists In Master Data Query"
        dbs.Execute "Master Data Update Query"
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If

End Function


Function QIndex_Update()

Dim answer As Integer
    answer = MsgBox("Run QIndex Update Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "QIndex Update")
    
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
 
    If answer = vbYes Then
        dbs.Execute "ISIR to QIndex 297"
        dbs.Execute "ISIR to QIndex 206"
        dbs.Execute "ISIR to QIndex 254"
        dbs.Execute "ISIR to QIndex 295"

        dbs.Execute "QIndex Update Query"
        dbs.Execute "Highest QIndex on Module"
        dbs.Execute "Update QIndex Inheritence"
    
        DoCmd.RunSQL "DELETE * from [Highest QIndex]"
        DoCmd.RunSQL "DELETE * from [QIndex Update]"
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If


End Function

Function Generators_Update()

Dim answer As Integer
    answer = MsgBox("Run QIndex Update Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "Generators Update")
    
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
 
    If answer = vbYes Then
        dbs.Execute "A-Nr+ Generator"
        dbs.Execute "LIUMF Generaor +"
        dbs.Execute "Update LIUMF"
        DoCmd.RunSQL "DELETE * from [LIUMF]"
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If

End Function

Function All_Used_Modules_Update()

Dim answer As Integer
    answer = MsgBox("Run All Used Modules Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "All Used Modules Update")
    
    Dim dbs As DAO.Database
    Set dbs = CurrentDb

    If answer = vbYes Then
        dbs.Execute "Clear All Used Modules"
        dbs.Execute "All Used Modules New Index"
        dbs.Execute "All Used Modules Old Index"
        DoCmd.RunSQL "DROP TABLE [All Used Modules]"
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If



End Function

Function Change_Index_State()

Dim answer As Integer
    answer = MsgBox("Run Change Index State Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "Change Index State")

    If answer = vbYes Then
        DoCmd.OpenForm ("frm_select_project")
    Else
        MsgBox "Action Canseled"
    End If

End Function

Function Index_State_Update(project As String)

Dim ProjectNumber As Integer, answer As Integer, Index As String

    If project = "W206" Then
        ProjectNumber = 1
        Index = "91V"
    End If
    If project = "V297" Then
        ProjectNumber = 2
        Index = "91W"
    End If
    If project = "V295" Then
        ProjectNumber = 3
        Index = "91W"
    End If
    If project = "X254" Then
        ProjectNumber = 4
        Index = "91X"
    End If
    If project = "C236" Then
        ProjectNumber = 5
        Index = "91V"
    End If
    
    DoCmd.RunSQL "UPDATE [db_owner_tbl_data_Index] SET [db_owner_tbl_data_Index].Status = 'X' WHERE ((([db_owner_tbl_data_Index].Status)='C' Or ([db_owner_tbl_data_Index].Status)='N') AND (Left([db_owner_tbl_data_Index].[Index-Nr],3)= '" & Index & "') AND (Mid([db_owner_tbl_data_Index].[A-Nr],2,3)= '" & Mid(project, 2, 3) & "'))"
    
    DoCmd.RunSQL "UPDATE [All Index] SET [All Index].Status = 'X' WHERE ((([All Index].Status)='C' Or ([All Index].Status)='N') AND (([All Index].ProjectID)= " & ProjectNumber & "))"
    On Error Resume Next
    DoCmd.Close acForm, "frm_select_project", acSaveYes

    DoCmd.Close acForm, "frm_select_project", acSaveYes
        MsgBox "Action Completed"


End Function


Function MD_Archive_Indevidual()

'MD Archive Indevidual confirm message box

Dim answer As Integer
    answer = MsgBox("Run Archive Indevidual MD Procedure?", vbQuestion + vbYesNo + vbDefaultButton2, "Archive to Active")

    If answer = vbYes Then
        DoCmd.OpenForm ("frm_select_project_archive_indevidual")
    Else
        MsgBox "Action Canseled"
    End If

End Function


Function MD_Archive_Indevidual_Update(project As String)

'MD Archive Indevidual action function

Dim answer As Integer
Dim SQL As String
Dim project_acode As String

    If project = "W206" Then project_acode = "A 206*"
    If project = "V297" Then project_acode = "A 297*"
    If project = "V295" Then project_acode = "A 295*"
    If project = "X254" Then project_acode = "A 254*"
    If project = "C236" Then project_acode = "A 236*"
    
'Delete Queries MD Archive


    DoCmd.RunSQL "DELETE * FROM [Master Data Archive BRW] WHERE ([Master Data Archive BRW].[Project Name]= '" & project & "')"
    DoCmd.RunSQL "DELETE * FROM [Master Data Archive Foaming] WHERE ([Master Data Archive Foaming].[Project Name]= '" & project & "')"
    DoCmd.RunSQL "DELETE * FROM [Master Data Archive OGC] WHERE ([Master Data Archive OGC].[Project Name]= '" & project & "')"
    DoCmd.RunSQL "DELETE * FROM [Master Data Archive PRG] WHERE ([Master Data Archive PRG].[Project]= '" & project & "')"
    DoCmd.RunSQL "DELETE * FROM [Master Data Archive RFA] WHERE ([Master Data Archive RFA].[Project Name]= '" & project & "')"
    DoCmd.RunSQL "DELETE * FROM [Master Data Archive Screwing] WHERE ([Master Data Archive Screwing].[Project]= '" & project & "')"
    DoCmd.RunSQL "DELETE * FROM [Master Data Archive USW Modules] WHERE ([Master Data Archive USW Modules].[Project]= '" & project & "')"
    DoCmd.RunSQL "DELETE * FROM [Master Data Archive USW Splice] WHERE ([Master Data Archive USW Splice].[Project Name]= '" & project & "')"
    DoCmd.RunSQL "DELETE * FROM [Master Data Archive MSG Basic Modules] WHERE ([Master Data Archive MSG Basic Modules].[Project Name]= '" & project & "')"
    DoCmd.RunSQL "DELETE * FROM [Master Data Archive MSG Basic Modules+PIN+Kabel] WHERE ([Master Data Archive MSG Basic Modules+PIN+Kabel].[Project Name]= '" & project & "')"

'Wirelist Queries Wirelist

    DoCmd.RunSQL "DELETE * FROM [WireList Previous Phase] WHERE ([WireList Previous Phase].[A-Nr] LIKE '" & project_acode & "')"
    DoCmd.RunSQL "INSERT INTO [WireList Previous Phase] SELECT * FROM [WireList] WHERE ([WireList].[A-Nr] LIKE '" & project_acode & "')"


'Run MD Queries
    
    'BRW
    SQL = "INSERT INTO [Master Data Archive BRW] ( [A-Nr], [Project Name], [Variant], [Index-Nr], [WireNumber], [BWN Address], [CrossSection], [LIUMF], [Status], [Hands], [Type], [Phase Current/End], [Future Implementation], [Serial Production] )" & _
            "SELECT [BRW Modules All Table].[A-Nr], [BRW Modules All Table].[Project Name], [BRW Modules All Table].[Variant], [BRW Modules All Table].[Index-Nr], [BRW Modules All Table].[WireNumber], [BRW Modules All Table].[BWN Address], [BRW Modules All Table].[CrossSection], [BRW Modules All Table].[LIUMF], [BRW Modules All Table].[Status], [BRW Modules All Table].[Hands], [BRW Modules All Table].[Type], [BRW Modules All Table].[Phase Current/End], [BRW Modules All Table].[Future Implementation], [BRW Modules All Table].[Serial Production]" & _
            "FROM [BRW Modules All Table]" & _
            "WHERE ((([BRW Modules All Table].[Notes])='Actual') AND (([BRW Modules All Table].[Project Name])= '" & project & "'))"
    
    DoCmd.RunSQL SQL

    'Foaming
    SQL = "INSERT INTO [Master Data Archive Foaming] ( [A-Nr], [Project Name], [Variant], [Index-Nr], [GROMMET], [WireNumber], [CrossSection], [Hands], [Status], [Phase], [Future Implementation], [Serial Production] )" & _
            "SELECT [Foaming-LL-RL].[A-Nr], [Foaming-LL-RL].[Project Name], [Foaming-LL-RL].[Variant], [Foaming-LL-RL].[Index-Nr], [Foaming-LL-RL].[GROMMET], [Foaming-LL-RL].[WireNumber], [Foaming-LL-RL].[CrossSection], [Foaming-LL-RL].[Hands], [Foaming-LL-RL].[Status], [Foaming-LL-RL].[Phase], [Foaming-LL-RL].[Future Implementation], [Foaming-LL-RL].[Serial Production]" & _
            "FROM [Foaming-LL-RL]" & _
            "WHERE ((([Foaming-LL-RL].[Notes])='Actual') AND (([Foaming-LL-RL].[Project Name])= '" & project & "'))"
    
    DoCmd.RunSQL SQL

    'OGC
    SQL = "INSERT INTO [Master Data Archive OGC] ( [Clips], [Project Name], [LED], [Module/Index-Nr], [Workplace], [Status], [Hands], [Phase], [Future Implementation], [Serial Production], [A-Nr], [Variant] )" & _
            "SELECT [OGC-LL-RL].[Clips], [OGC-LL-RL].[Project Name], [OGC-LL-RL].[LED], [OGC-LL-RL].[Module/Index-Nr], [OGC-LL-RL].[Workplace], [OGC-LL-RL].[Status], [OGC-LL-RL].[Hands], [OGC-LL-RL].[Phase], [OGC-LL-RL].[Future Implementation], [OGC-LL-RL].[Serial Production], [OGC-LL-RL].[A-Nr], [OGC-LL-RL].[Variant]" & _
            "FROM [OGC-LL-RL]" & _
            "WHERE ((([OGC-LL-RL].[Notes])='Actual') AND (([OGC-LL-RL].[Project Name])= '" & project & "'))"

    DoCmd.RunSQL SQL

    'PRG
    SQL = "INSERT INTO [Master Data Archive PRG] ( [A-Nr], [Project], [Variant], [Index-Nr], [Status], [Type], [Hand], [Drawing A-Nr], [PRG with Clips], [Phase Current/End], [Future Implementation], [Serial Production], [Files] )" & _
            "SELECT [PRG List Data].[A-Nr], [PRG List Data].[Project], [PRG List Data].[Variant], [PRG List Data].[Index-Nr], [PRG List Data].[Status], [PRG List Data].[Type], [PRG List Data].[Hand], [PRG List Data].[Drawing A-Nr], [PRG List Data].[PRG with Clips], [PRG List Data].[Phase Current/End], [PRG List Data].[Future Implementation], [PRG List Data].[Serial Production], [PRG List Data].[Files]" & _
            "FROM [PRG List Data]" & _
            "WHERE ((([PRG List Data].[Notes])='Actual') AND (([PRG List Data].[Project])= '" & project & "'))"

    DoCmd.RunSQL SQL

    'RFA
    SQL = "INSERT INTO [Master Data Archive RFA] ( [A-Nr], [Project Name], [Variant], [Module], [SIBO], [Section], [Place], [PPSNumber], [Status], [Hands], [Phase Current/End], [Future Implementation], [Serial Production] )" & _
            "SELECT RFA.[A-Nr], [RFA].[Project Name], [RFA].[Variant], [RFA].[Module], [RFA].[SIBO], [RFA].[Section], [RFA].[Place], [RFA].[PPSNumber], [RFA].[Status], [RFA].[Hands], [RFA].[Phase Current/End], [RFA].[Future Implementation], [RFA].[Serial Production]" & _
            "FROM [RFA]" & _
            "WHERE ((([RFA].[Notes])='Actual') AND (([RFA].[Project Name])= '" & project & "'))"

    DoCmd.RunSQL SQL

    'Screwing
    SQL = "INSERT INTO [Master Data Archive Screwing] ( [LIUMF], [Module], [SIBO], [Section], [Place], [PPSNumber], [A-Nr], [Phase Current/End], [Variant], [Future Implementation], [Serial Production], [Project], [Hands] )" & _
            "SELECT [Screwing].[LIUMF], [Screwing].[Module], [Screwing].[SIBO], [Screwing].[Section], [Screwing].[Place], [Screwing].[PPSNumber], [Screwing].[A-Nr], [Screwing].[Phase Current/End], [Screwing].[Variant], [Screwing].[Future Implementation], [Screwing].[Serial Production], [Screwing].[Project], [Screwing].[Hands]" & _
            "FROM [Screwing]" & _
            "WHERE ((([Screwing].[Notes])='Actual') AND (([Screwing].[Project])= '" & project & "'))"

    DoCmd.RunSQL SQL

    'USW Modules
    SQL = "INSERT INTO [Master Data Archive USW Modules] ( [A-Nr], [Project], [Variant], [Index], [Status], [Drawing A-Nr], [Files], [Type], [Hand], [Deviation], [Phase], [Future Implementation], [Serial Production] )" & _
            "SELECT [USW Modules].[A-Nr], [USW Modules].[Project], [USW Modules].[Variant], [USW Modules].[Index], [USW Modules].[Status], [USW Modules].[Drawing A-Nr], [USW Modules].[Files], [USW Modules].[Type], [USW Modules].[Hand], [USW Modules].[Deviation], [USW Modules].[Phase], [USW Modules].[Future Implementation], [USW Modules].[Serial Production]" & _
            "FROM [USW Modules]" & _
            "WHERE ((([USW Modules].[Notes])='Actual') AND (([USW Modules].[Project])= '" & project & "'))"

    DoCmd.RunSQL SQL

    'USW Splice
    SQL = "INSERT INTO [Master Data Archive USW Splice] ( [A-Nr], [Project Name], [Variant], [Index-Nr], [WireNumber], [USW Addresses], [Cross Section], [Deviation], [Status], [Hand], [Types] )" & _
            "SELECT [USW Splice].[A-Nr], [USW Splice].[Project Name], [USW Splice].[Variant], [USW Splice].[Index-Nr], [USW Splice].[WireNumber], [USW Splice].[USW Addresses], [USW Splice].[Cross Section], [USW Splice].[Deviation], [USW Splice].[Status], [USW Splice].[Hand], [USW Splice].[Types]" & _
            "FROM [USW Splice]" & _
            "WHERE ((([USW Splice].[Notes])='Actual') AND (([USW Splice].[Project Name])= '" & project & "'))"

    DoCmd.RunSQL SQL
    
'MSG BASIC Modules
    SQL = "INSERT INTO [Master Data Archive MSG Basic Modules] ( [ModuleNumber], [ItemName], [ModuleIndex], [Index-Nr], [LIUMF], [Status], [A-Nr], [Variant], [Project Name], [Engine Type], [Hands] )" & _
            "SELECT [MSG Basic Modules].[ModuleNumber], [MSG Basic Modules].[ItemName], [MSG Basic Modules].[ModuleIndex], [MSG Basic Modules].[Index-Nr], [MSG Basic Modules].[LIUMF], [MSG Basic Modules].[Status], [MSG Basic Modules].[A-Nr], [MSG Basic Modules].[Variant], [MSG Basic Modules].[Project Name], [MSG Basic Modules].[Engine Type], [MSG Basic Modules].[Hands]" & _
            "FROM [MSG Basic Modules]" & _
            "WHERE ((([MSG Basic Modules].Notes)='Actual') AND (([MSG Basic Modules].[Project Name])= '" & project & "'))"

    DoCmd.RunSQL SQL

'MSG BASIC Modules+PIN+KABEL

    SQL = "INSERT INTO [Master Data Archive MSG Basic Modules+PIN+Kabel] ( [ModuleNumber], [ItemName], [ModuleIndex], [Index-Nr], [LIUMF], [Status], [A-Nr], [Variant], [Project Name], [Engine Type], [Xcode], [PinNumber], [WireNumber], [Hands] )" & _
            "SELECT [MSG Basic Modules+PIN+Kabel].[ModuleNumber], [MSG Basic Modules+PIN+Kabel].[ItemName], [MSG Basic Modules+PIN+Kabel].[ModuleIndex], [MSG Basic Modules+PIN+Kabel].[Index-Nr], [MSG Basic Modules+PIN+Kabel].[LIUMF], [MSG Basic Modules+PIN+Kabel].[Status], [MSG Basic Modules+PIN+Kabel].[A-Nr], [MSG Basic Modules+PIN+Kabel].[Variant], [MSG Basic Modules+PIN+Kabel].[Project Name], [MSG Basic Modules+PIN+Kabel].[Engine Type], [MSG Basic Modules+PIN+Kabel].[Xcode], [MSG Basic Modules+PIN+Kabel].[PinNumber], [MSG Basic Modules+PIN+Kabel].[WireNumber], [MSG Basic Modules+PIN+Kabel].[Hands]" & _
            "FROM [MSG Basic Modules+PIN+Kabel]" & _
            "WHERE ((([MSG Basic Modules+PIN+Kabel].Notes)='Actual') AND (([MSG Basic Modules+PIN+Kabel].[Project Name])= '" & project & "'))"

    DoCmd.RunSQL SQL

    On Error Resume Next
    DoCmd.Close acForm, "frm_select_project_archive_indevidual", acSaveYes

    DoCmd.Close acForm, "frm_select_project_archive_indevidual", acSaveYes
    MsgBox "Action Completed"


End Function

Function Autokan_Previous_Phase_Update()

    Dim answer As Integer
        answer = MsgBox("Confirm Autokan Previous Phase Update?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Data")

    If answer = vbYes Then
        DoCmd.RunSQL "DELETE * from [Wirelist+ No Z Previous Phase]"
        DoCmd.RunSQL "DELETE * from [Wirelist++ Autokan Previous Phase]"

        DoCmd.RunSQL "INSERT INTO [Wirelist+ No Z Previous Phase] SELECT * FROM [Wirelist+ no Z]"
        DoCmd.RunSQL "INSERT INTO [Wirelist++ Autokan Previous Phase] SELECT * FROM [Wirelist++ Autokan]"

        DoEvents
        
        MsgBox "Action Completed!"
    Else
        MsgBox "Action Canseled"
    End If
DoCmd.SetWarnings (WarningsOff)



End Function

Function Process_AEMs()

Dim answer As Integer, answer1 As Integer, answer2 As Integer, answer3 As Integer, answer4 As Integer, answer5 As Integer, answer6 As Integer

Dim dbs As DAO.Database
Set dbs = CurrentDb

answer = MsgBox("Are you have New Modules for IMPORT in DBMS?", vbQuestion + vbYesNo + vbDefaultButton2, "AEM Process")
If answer = vbYes Then
    answer1 = MsgBox("Do you have all the needed information in a table: '00 - New Modules Table'?", vbQuestion + vbYesNo + vbDefaultButton2, "AEM Process")
    If answer1 = vbYes Then
        dbs.Execute "New Module Step 01 - Import - All Module from New Module Table"
        dbs.Execute "New Module Step 02 - Import - All Index from New Module Table"
        dbs.Execute "New Module Step 03 - IMPORT-DBLB_MODULE_from_New_Modules_Table"
        dbs.Execute "New Module Step 04 - IMPORT-DBLB_INDEX_from_New_Module_Table"
        dbs.Execute "New Module Step 05 - IMPORT-DBLB_ISIR_from_New_Module_Table"
        MsgBox "Action Completed! The new modules are imported!"
    Else
        MsgBox "Action Canceled! Please insert correct information in a table: '00 - New Modules Table' and try again!"
        'GoTo Label1
    End If
Else
    MsgBox "Go ahead, please!"
End If
answer2 = MsgBox("Are you have Deleted Modules for UPDATE in DBMS?", vbQuestion + vbYesNo + vbDefaultButton2, "AEM Process")
If answer2 = vbYes Then
    answer3 = MsgBox("Do you have all the needed information in a table: 'DBLB 03 - tbl_Deleted Modules'?", vbQuestion + vbYesNo + vbDefaultButton2, "AEM Process")
    If answer3 = vbYes Then
        dbs.Execute "Delete Module Step 01 - UPDATE-DBLB_MODULE"
        dbs.Execute "Delete Module Step 02 - UPDATE-DBLB_INDEX"
        MsgBox "Action Completed! The information for deleted modules is updated!"
    Else
        MsgBox "Action Canceled! Please insert correct information in a table: 'DBLB 03 - tbl_Deleted Modules' and try again!"
        'GoTo Label1
    End If
Else
    MsgBox "Go ahead, please!"
End If
answer4 = MsgBox("Are you have changed Drawing Dates, Weights or ZGS for some modules to be UPDATED in DBMS?", vbQuestion + vbYesNo + vbDefaultButton2, "AEM Process")
If answer4 = vbYes Then
    answer5 = MsgBox("Do you have all the needed information in a table: 'DBLB - tbl_Update_Weight_Date_etc'?", vbQuestion + vbYesNo + vbDefaultButton2, "AEM Process")
    If answer5 = vbYes Then
        dbs.Execute "IMPORT-DBLB_UPDATES_from_tbl_Update_Weight_Date_etc"
        MsgBox "Action Completed! The information for New Drawing Date, Weights or ZGS is updated!"
    Else
        MsgBox "Action Canceled! Please insert correct information in a table: 'DBLB - tbl_Update_Weight_Date_etc' and try again!"
        'GoTo Label1
    End If
Else
    MsgBox "Go ahead, please!"
End If
answer6 = MsgBox("Finally you start AEM's process! Are you sure?", vbQuestion + vbYesNo + vbDefaultButton2, "AEM Process")
If answer6 = vbYes Then
    dbs.Execute "AEM Update Query"
    'dbs.Execute "Step 1 - Retrive Highest Index"
    'dbs.Execute "Step 05 - Connect AEMs to A-Nr"
    Call Update_Connect_AEMs_to_Last_Index
    dbs.Execute "Step 3 - Append New Status to Status Update"
    dbs.Execute "DBLB Step 02 - UPDATE tbl_Update"
    dbs.Execute "Step 2 - Upload New Index to All Index"
    dbs.Execute "DBLB Step 01 - UPDATE tbl_AddNewData"
    dbs.Execute "Step 4 - Update Index Status and Phase"
    dbs.Execute "IMPORT-DBLB_UPDATES_from_tbl_Update"
    dbs.Execute "DBLB Step 00 - IMPORT-DBLB_ADD_NEW_DATAMODUL_from_tbl_AddNewData"
    dbs.Execute "Step 01 - IMPORT-DBLB_ADD_NEW_DATA_from_tbl_AddNewData"
    dbs.Execute "Step 02 - IMPORT-DBLB_ADD_NEW_DATA_ISIR_from_tbl_AddNewData"
    MsgBox "Action Completed! The AEM's are processed successfully!"
Else
    MsgBox "Action Canceled!"
End If
    
Label1:
End Function

Function Update_Connect_AEMs_to_Last_Index()

Dim x As Integer, y As Integer, z As String, last_row As Integer, last_column As Integer, MyAnumber As String, MyIndex As String, MyAEM As Variant, MyString As String, MyTables(2) As Variant, MyQuerys(1) As Variant

Dim strCriteria As String, MyAEMs(300, 1), MyField As String

MyTables(1) = "Connect AEMs to A-Nr"
MyTables(2) = "Collect AEMs"
MyQuerys(1) = "Step 05 - Connect AEMs to A-Nr"

DoCmd.CopyObject , MyTables(2), acTable, MyQuerys(1)

strCriteria = "ALTER TABLE [" & MyTables(2) & "] " & _
                "ADD Column ID COUNTER;"
DoCmd.RunSQL strCriteria

last_row = DCount("[A-Nr]", MyTables(2))
last_column = CurrentDb.TableDefs(MyTables(2)).Fields.Count
    
    For y = 1 To last_row
        x = 2
        MyField = "[" & CurrentDb.TableDefs(MyTables(2)).Fields(x - 2).Name & "]"
        MyAnumber = DLookup(MyField, MyTables(2), "[ID]= " & y)
        MyField = "[" & CurrentDb.TableDefs(MyTables(2)).Fields(x - 1).Name & "]"
        MyIndex = DLookup(MyField, MyTables(2), "[ID]= " & y)
        For x = 2 To last_column - 2
            MyField = "[" & CurrentDb.TableDefs(MyTables(2)).Fields(x).Name & "]"
            MyAEM = DLookup(MyField, MyTables(2), "[ID]= " & y)
            If IsNull(MyAEM) Then
                GoTo Label1
            Else
                If MyAEMs(y, 0) = Empty Then
                    MyAEMs(y, 0) = MyAEM
                Else
                    MyAEMs(y, 0) = MyAEMs(y, 0) & ", "
                    MyAEMs(y, 0) = MyAEMs(y, 0) & MyAEM
                End If
            End If
Label1:
        Next
        strCriteria = "INSERT INTO [" & MyTables(1) & "] ([A-Nr], [MaxOfNewIndex], [AEM's])" & _
                          "VALUES ('" & MyAnumber & "', '" & MyIndex & "', '" & MyAEMs(y, 0) & "')"
        DoCmd.SetWarnings (WarningsOff)
        DoCmd.RunSQL strCriteria
    Next

DoCmd.RunSQL "DROP TABLE [" & MyTables(2) & "]"
End Function

