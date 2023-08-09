Attribute VB_Name = "LogsReportService"
'@Folder("ReportMergerProject.src")
Option Explicit

Const DeveloperEmail As String = ""

Public Sub WriteLog(ByVal Text As String)
    Const MasLogsCount As Integer = 500

    Dim LogsCount As Integer: LogsCount = Logs.UsedRange.Rows.Count
    If LogsCount >= MasLogsCount Then
        Logs.UsedRange.Rows(FString("{0}:{1}", LogsCount, MasLogsCount)).Delete
    End If
    Logs.Rows(1).Insert
    Dim Cell As Range: Set Cell = Logs.Cells(1, 1)

    Cell.Value = FString("{0} | {1}", Time, Text)
End Sub

Public Sub SendReport()
    Dim Body As String: Body = GetBody()
    Dim Letter As Letter: Set Letter = NewLetter(DeveloperEmail, Subject:=Constants.AppNameEn, Body:=Body, NeedSend:=True)
    Dim Sender As Sender: Set Sender = NewSender()
    Sender.CreateLetter Letter
    MsgBox "Успешно.", vbInformation, Constants.AppNameRu
End Sub

Private Function GetBody() As String
    Dim Data As Variant: Data = Logs.UsedRange.Value
    If Not IsArray(Data) Then Data = Array(Data)
    Dim Buffer As ArrayList: Set Buffer = New ArrayList
    Buffer.AddRange Data

    GetBody = Strings.Join(Buffer.ToArray(), "<br />")
End Function
