Attribute VB_Name = "FilesService"
'@Folder("ReportMergerProject.src")
Option Explicit

Public Function GetNewFiles(ByVal ReportsFolder As String) As Variant
    Dim Folder As Object: Set Folder = CreateObject("Scripting.FileSystemObject").GetFolder(ReportsFolder)
    Dim Buffer As ArrayList: Set Buffer = New ArrayList

    Dim File As Object
    For Each File In Folder.Files
        If File.Attributes <> VbFileAttribute.vbArchive Then GoTo Continue
        Buffer.Add File.Path
Continue:
    Next

    GetNewFiles = Buffer.ToArray()
End Function
