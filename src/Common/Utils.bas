Attribute VB_Name = "Utils"
'@Folder("ReportMergerProject.src.Common")
Option Explicit

Public Sub DevMode()
    Static Dev As Boolean: Dev = Not Dev
    Dim Visibility As XlSheetVisibility: Visibility = IIf(Dev, xlSheetVisible, xlSheetVeryHidden)

    Logs.Visible = Visibility
End Sub
