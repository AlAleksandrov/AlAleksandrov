Attribute VB_Name = "09a-AEM_Preparation_PDF_to_TXT"
Option Compare Database

Function AEM_Preparation_PDF_to_TXT()
'
' PDF_to_TXT Macro
' Save .pdf to .txt
'
Dim x As Integer, y As Long, MyLinkSource As String, MyLinkDest As String, MyLink As String, MyPDFName(1000, 1), MyType(2), MyFile As String, project As String, intPath As Integer, intPDFFile As Long, OpenAt As Variant

Dim report_path As String

OpenAt = "\\leoni.local\dfsroot\BG1\Groups\AO\01-Change managment\01-C Class Project\03 Masterdata_LEPS\01-Database Master Data\06-Import\05 - AEMs\"

report_path = BrowseForFolder(OpenAt) & "\"

If Right(report_path, 4) = "COC\" Or Right(report_path, 4) = "MRA\" Then
    If Right(report_path, 9) = "W206\COC\" Or Right(report_path, 9) = "W206\MRA\" Then
        project = Left(Right(report_path, 9), 4)
    ElseIf Right(report_path, 9) = "V297\COC\" Or Right(report_path, 9) = "V297\MRA\" Then
        project = Left(Right(report_path, 9), 4)
    ElseIf Right(report_path, 9) = "V295\COC\" Or Right(report_path, 9) = "V295\MRA\" Then
        project = Left(Right(report_path, 9), 4)
    ElseIf Right(report_path, 9) = "X254\COC\" Or Right(report_path, 9) = "X254\MRA\" Then
        project = Left(Right(report_path, 9), 4)
    ElseIf Right(report_path, 9) = "C236\COC\" Or Right(report_path, 9) = "C236\MRA\" Then
        project = Left(Right(report_path, 9), 4)
    ElseIf Mid(file_name, 2, 3) = "" Then
        a = MsgBox("File not found or in wrong naming scheme!", , "ERROR")
    End If
End If

MyLink = OpenAt

For x = 0 To 1
    If x = 0 Then
        MyType(x) = "COC"
        MyLinkSource = (MyLink & project & "\" & MyType(x) & "\")
        intPath = Len(Dir(MyLinkSource, vbDirectory))
        If intPath > 0 Then
            intPDFFile = Len(Dir(MyLinkSource & "ÄM*.pdf"))
            If intPDFFile = 0 Then GoTo Label1
            MyFile = Dir(MyLinkSource & "ÄM*.pdf")
            While (MyFile <> "")
                MyFile = (MyType(x) & "_" & MyFile)
                MyPDFName(i, x) = MyFile
                i = i + 1
                MyFile = Dir
            Wend
        Else
            GoTo Label1
        End If
    ElseIf x = 1 Then
        i = 0
        MyType(x) = "MRA"
        MyLinkSource = (MyLink & project & "\" & MyType(x) & "\")
        intPath = Len(Dir(MyLinkSource, vbDirectory))
        If intPath > 0 Then
            intPDFFile = Len(Dir(MyLinkSource & "ÄM*.pdf"))
            If intPDFFile = 0 Then GoTo Label1
            MyFile = Dir(MyLinkSource & "ÄM*.pdf")
            While (MyFile <> "")
                MyFile = (MyType(x) & "_" & MyFile)
                MyPDFName(i, x) = MyFile
                i = i + 1
                MyFile = Dir
            Wend
        Else
            GoTo Label1
        End If
    End If
    For y = 0 To i - 1
        MyLinkSource = (MyLink & project & "\" & MyType(x) & "\")
        intPath = Len(Dir(MyLinkSource, vbDirectory))
        If intPath = 0 Then MkDir (MyLink & project & "\" & MyType(x) & "\")
        MyLinkDest = (MyLink & project & "\TXT\")
        intPath = Len(Dir(MyLinkDest, vbDirectory))
        If intPath = 0 Then MkDir (MyLink & project & "\TXT\")
'MsgBox (MyLinkSource)
        'ChangeFileOpenDirectory MyLinkSource
        MyLinkSource = (MyLink & project & "\" & MyType(x) & "\" & Mid(MyPDFName(y, x), 5, 100))
'MsgBox (MyLinkSource)
        Documents.Open FileName:=MyLinkSource _
            , ConfirmConversions:=False, readOnly:=False, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", Revert:=False, _
            WritePasswordDocument:="", WritePasswordTemplate:="", Format:= _
            wdOpenFormatAuto, XMLTransform:=""
'MsgBox (MyLinkDest)
        ChangeFileOpenDirectory MyLinkDest
        MyLinkDest = (MyLink & project & "\TXT\" & MyPDFName(y, x) & ".txt")
'MsgBox (MyLinkDest)
        ActiveDocument.SaveAs2 FileName:=MyLinkDest _
            , FileFormat:=wdFormatText, LockComments:=False, password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False, Encoding:=1252, InsertLineBreaks:=False _
            , AllowSubstitutions:=False, LineEnding:=wdCRLF, CompatibilityMode:=0
        ActiveDocument.Close
    Next
Label1:
Next
End Function
