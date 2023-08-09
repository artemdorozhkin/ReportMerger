VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReportMerger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ReportMergerProject.src")
Option Explicit

Private Type TReportMerger
    ResultBook As Workbook
    ReportBook As Workbook
End Type

Private this As TReportMerger

Public Sub Merge(ByVal ResultFolder As String, ByRef NewReports As Variant)
    IPMExcel.DisableAllerts
    Dim ErrMsg As String

    Dim ResultPath As String: ResultPath = CreateNewResultBook(ResultFolder)
    If IsFalse(ResultPath) Then
        ErrMsg = FString("Ну удалось найти файл '{0}'.", Constants.ResultFileName)
        Err.Raise ErrorService.GetErrorNumber(ErrorService.PathNotFound), Description:=ErrMsg
    End If
    Set this.ResultBook = Workbooks.Open(ResultPath)

    Dim NewReport As Variant
    For Each NewReport In NewReports
        Set this.ReportBook = Workbooks.Open(NewReport)
        Dim SheetIndex As Integer: SheetIndex = FindSheetIndex()
        If SheetIndex = 0 Then
            SaveChanges
            ErrMsg = FString("Ну удалось найти лист с данными в файле '{0}'.", this.ReportBook.Name)
            Err.Raise ErrorService.GetErrorNumber(ErrorService.MissingSheet), Description:=ErrMsg
        End If
        Dim Status As String: Status = FString("Копирую данные файла: {0}", this.ReportBook.Name)
        IPMExcel.PrintStatus Status
        LogsReportService.WriteLog Status
        CopyData SheetIndex
        SaveChanges
        this.ReportBook.Close False
    Next

    this.ResultBook.Close
    IPMExcel.EnableAllerts
End Sub

Private Sub SaveChanges()
    this.ResultBook.Save
End Sub

Private Function CreateNewResultBook(ByVal ResultFolder As String) As String
    Dim Book As Workbook: Set Book = Workbooks.Add()
    Book.SaveAs NodePath.Join(ResultFolder, FString(Constants.ResultFileName, DateTime.Date))
    CreateNewResultBook = Book.FullName
    Book.Close False
End Function

Private Function FindSheetIndex() As Integer
    Dim Sheet As Worksheet
    For Each Sheet In this.ReportBook.Worksheets
        Dim ControlColumn As Range: Set ControlColumn = Sheet.UsedRange.Rows(1).Find(Constants.ControlColumnName, _
                                                                                     LookIn:=XlFindLookIn.xlValues, _
                                                                                     LookAt:=XlLookAt.xlPart, _
                                                                                     MatchCase:=False)
        If IsFalse(ControlColumn) Then GoTo Continue

        FindSheetIndex = Sheet.Index
        Exit Function
Continue:
    Next
End Function

Private Sub CopyData(ByVal SourceSheetIndex As Integer)
    ShowAllSheetData this.ReportBook.Worksheets(SourceSheetIndex)

    Dim ReportTable As Range
    Set ReportTable = this.ReportBook.Worksheets(SourceSheetIndex).UsedRange
    Dim ResultTable As Range
    Set ResultTable = this.ResultBook.ActiveSheet.UsedRange
    ReportTable.Rows(1).Copy
    ResultTable.Rows(1).PasteSpecial xlPasteAll

    Set ReportTable = ReportTable.Offset(1).Resize(ReportTable.Rows.Count - 1)
    ReportTable.Copy

    Dim Destination As Range: Set Destination = ResultTable.Cells(ResultTable.Rows.Count + 1, 1)
    Destination.PasteSpecial xlPasteAll
End Sub

Private Sub ShowAllSheetData(ByVal Sheet As Worksheet)
    With Sheet
        On Error Resume Next
        .ShowAllData
        With .UsedRange
            .Columns.EntireColumn.Hidden = False
            .Rows.EntireColumn.Hidden = False
        End With
    End With
End Sub