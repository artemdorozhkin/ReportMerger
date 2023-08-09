Attribute VB_Name = "App"
'@Folder "ReportMergerProject.src"
Option Explicit

Public Sub Main()
    On Error GoTo Catch
    IPMExcel.DisableSettings

    Dim ErrMsg As String
    Config.App = Constants.AppNameEn
    Config.Section = Constants.SettingsSection

    IPMExcel.PrintStatus "Проверка настроек..."
    Dim ReportFolder As String: ReportFolder = Config.GetValue(Constants.ReportsFolderKey)
    Dim ResultFolder As String: ResultFolder = Config.GetValue(Constants.ResultsFolderKey)

    If IsFalse(ReportFolder) Then
        ErrMsg = "Не указан путь к папке с отчетами. Проверьте настройки макроса."
        Err.Raise ErrorService.GetErrorNumber(ErrorService.PathNotSet), Description:=ErrMsg
        GoTo ExitSub
    End If

    If IsFalse(ResultFolder) Then
        ErrMsg = "Не указан путь к папке для сохранения результатов. Проверьте настройки макроса."
        Err.Raise ErrorService.GetErrorNumber(ErrorService.PathNotSet), Description:=ErrMsg
        GoTo ExitSub
    End If

    Dim FS As FS: Set FS = New FS
    If Not FS.DirExists(ReportFolder) Then
        ErrMsg = "Не удалось найти папку с отчетами по указанному в настройках пути. Возможно она была удалена или перемещена.\\n\\nПроверьте правильность указанных данных в настройках макроса."
        Err.Raise ErrorService.GetErrorNumber(ErrorService.PathNotFound), Description:=ErrMsg
        GoTo ExitSub
    End If

    If Not FS.DirExists(ResultFolder) Then
        ErrMsg = "Не удалось найти папку для сохранения результатов по указанному в настройках пути. Возможно она была удалена или перемещена.\\n\\nПроверьте правильность указанных данных в настройках макроса."
        Err.Raise ErrorService.GetErrorNumber(ErrorService.PathNotFound), Description:=ErrMsg
        GoTo ExitSub
    End If

    Dim Files As Variant: Files = FilesService.GetNewFiles(ReportFolder)
    If UBound(Files) = -1 Then
        MsgBox "Нет новых отчетов.", vbInformation, Constants.AppNameRu
        GoTo ExitSub
    End If

    Dim ReportMerger As ReportMerger: Set ReportMerger = New ReportMerger
    ReportMerger.Merge ResultFolder, Files

    MsgBox "Готово.", vbInformation, Constants.AppNameRu
    IPMExcel.OpenPath ResultFolder

ExitSub:
    IPMExcel.PrintStatus
    IPMExcel.EnableSettings
Exit Sub

Catch:
    Dim Unexpected As String: Unexpected = "Непредвиденная ошибка. Попробуйте повторить попытку.\\nЕcли вы не знаете как ее устранить, обратитесь в отдел разработки.\\n\\n#{0}\\nОшибка: {1}"
    Dim Style As VbMsgBoxStyle: Style = vbExclamation
    If Not ErrorService.IsUserError(Err.Number) Then
        Err.Description = FString(Unexpected, Err.Number, Err.Description)
        Style = vbCritical
    End If

    Dim Msg As String: Msg = FString(Err.Description)
    LogsReportService.WriteLog Msg
    MsgBox Msg, Style, Constants.AppNameRu
    GoTo ExitSub
End Sub

