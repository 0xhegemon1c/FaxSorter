Sub FaxSorter()

'coded by Jacqueline Xavier'
'Manipulating within Outlook (defining attachments and folders)`
Dim ns As Outlook.NameSpace
Set ns = Application.GetNamespace("MAPI")
Dim objItem As Outlook.MailItem
Dim oAttachment As Outlook.Attachment

'Naming the new folders with the date the code is run'
Dim sFolderName As String
sFolderName = Format(Now, "mm-dd-yyyy")

'VBA FileSystemObject to work with files + folders on our system (I:) and to convert in Adobe'
Dim FSO As Object
Dim FSOFile As Object
Dim FSOLibrary As Object
Dim FSOFolder As Object
Dim fldrname As String

Dim createFile As Boolean
createFile = False
Dim i As Integer: i = 1

'Define path to the target folder in Outlook'
Set moveToFolderMayo = ns.Folders("LabSS.Faxes").Folders("Path_Reports").Folders("Mayo")
Set moveToFolderNeo = ns.Folders("LabSS.Faxes").Folders("Path_Reports").Folders("Neogenomics")
Set moveToFolderFou = ns.Folders("LabSS.Faxes").Folders("Path_Reports").Folders("Foundation")
Set moveToFolderInter = ns.Folders("LabSS.Faxes").Folders("Path_Reports").Folders("Interpace Diagnostics")
Set moveToFolderThyro = ns.Folders("LabSS.Faxes").Folders("Path_Reports").Folders("Thyroseq")
Set moveToFolderArkana = ns.Folders("LabSS.Faxes").Folders("Path_Reports").Folders("Arkana")
Set moveToFolderOncoType = ns.Folders("LabSS.Faxes").Folders("Path_Reports").Folders("OncoType")
Set moveToFolderOncoCyte = ns.Folders("LabSS.Faxes").Folders("Path_Reports").Folders("OncoCyte")

'check if user has selected an email'
If Application.ActiveExplorer.Selection.Count = 0 Then
   MsgBox ("No email selected")
   Exit Sub
End If

For Each objItem In Application.ActiveExplorer.Selection
    If objItem.Class = olMail Then
        For Each oAttachment In objItem.Attachments
          If InStr(oAttachment.DisplayName, ".pdf") Then
          Set FSO = CreateObject("Scripting.FileSystemObject")
          Set FSOLibrary = CreateObject("Scripting.FileSystemObject")
          fldrname = BrowseForFolder("I:\Lab\Path_Reports\")
          datename = "\" & sFolderName
          Do While createFile = False
           newfldr = fldrname & datename & " " & "(" & i & ")"
           If Dir(newfldr, vbDirectory) = "" Then
               MkDir newfldr
               createFile = True
            Else
               i = i + 1
            End If
            Loop
          oAttachment.SaveAsFile newfldr & "\" & Format(DateAdd("d", -1, objItem.ReceivedTime), "mm-dd-yyyy" & "_" & "H-mm") & "_" & oAttachment.DisplayName
          Set FSOFolder = FSOLibrary.GetFolder(newfldr)
          Set FSOFile = FSOFolder.Files
          Debug.Print "report downloaded to i: drive in new folder."
            For Each FSOFile In FSOFile
              Call SavePDFAsJPG(FSOFile)
              Debug.Print "pages converted to JPG"
            Next
          End If
        Next
          objItem.UnRead = False
          Select Case fldrname
          Case "I:\Lab\Path_Reports\Mayo"
          objItem.Move moveToFolderMayo
          Debug.Print "Fax sorted fax sorted to Mayo folder"
          Case "I:\Lab\Path_Reports\Neogenomics"
          objItem.Move moveToFolderNeo
          Debug.Print "Fax sorted to Neogenomics folder"
          Case "I:\Lab\Path_Reports\Foundation"
          objItem.Move moveToFolderFou
          Debug.Print "Fax sorted to Foundation folder"
          Case "I:\Lab\Path_Reports\Interpace Diagnostics"
          objItem.Move moveToFolderInter
          Debug.Print "Fax sorted to Interpace Diagnostics folder"
          Case "I:\Lab\Path_Reports\Thyroseq"
          objItem.Move moveToFolderThyro
          Debug.Print "Fax sorted to Thyroseq folder"
          Case "I:\Lab\Path_Reports\Arkana"
          objItem.Move moveToFolderArkana
          Debug.Print "Fax sorted to Arkana folder"
          Case "I:\Lab\Path_Reports\OncoType"
          objItem.Move moveToFolderOncoType
          Debug.Print "Fax sorted to OncoType folder"
          Case "I:\Lab\Path_Reports\OncoCyte"
          objItem.Move moveToFolderOncoCyte
          Debug.Print "Fax sorted to OncoCyte folder"
          End Select
  End If
Next

End Sub

'sub to convert pdf documents into images using Adobe Acrobat'
Public Sub SavePDFAsJPG(PDFPath As Object)
    Dim objAcroApp      As Acrobat.AcroApp
    Dim objAcroAVDoc    As Acrobat.AcroAVDoc
    Dim objAcroPDDoc    As Acrobat.AcroPDDoc
    Dim objJSO          As Object
    Dim boResult        As Boolean
    Dim ExportFormat    As String
    Dim NewFilePath     As String

    Set objAcroApp = CreateObject("AcroExch.App")
    Set objAcroAVDoc = CreateObject("AcroExch.AVDoc")
    boResult = objAcroAVDoc.Open(PDFPath, "")
    Set objAcroPDDoc = objAcroAVDoc.GetPDDoc
    Set objJSO = objAcroPDDoc.GetJSObject
    ExportFormat = "com.adobe.acrobat.jpeg"
    NewFilePath = Replace(PDFPath, ".pdf", ".jpg")

    boResult = objJSO.SaveAs(NewFilePath, ExportFormat)
    boResult = objAcroAVDoc.Close(True)
    boResult = objAcroApp.Exit

    Set objAcroPDDoc = Nothing
    Set objAcroAVDoc = Nothing
    Set objAcroApp = Nothing
End Sub

'select box code that allows users to choose a save destination'
Function BrowseForFolder(Optional OpenAt As Variant) As Variant
  Dim ShellApp As Object
  Set ShellApp = CreateObject("Shell.Application"). _
 BrowseForFolder(0, "Please choose a folder", 0, OpenAt)

 On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
 On Error GoTo 0

 Set ShellApp = Nothing
    Select Case Mid(BrowseForFolder, 2, 1)
        Case Is = ":"
            If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
        Case Is = "\"
            If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
        Case Else
            GoTo Invalid
    End Select
 Exit Function

Invalid:
 BrowseForFolder = False
End Function
