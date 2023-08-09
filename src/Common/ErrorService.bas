Attribute VB_Name = "ErrorService"
'@Folder "ReportMergerProject.src.Common"
''' <summary>
''' Модуль ErrorService предоставляет функциональность для работы с ошибками в приложении GPPOperationsProject.
''' </summary>
''' <remarks>
''' Модуль ErrorService предоставляет методы для работы с ошибками, проверки пользовательских ошибок и обработки исключений.
''' </remarks>
''' <example>
''' <code>
''' Dim errorNumber As Long = ErrorService.GetErrorNumber("ColumnNotFound")
''' Dim isUserError As Boolean = ErrorService.IsUserError(errorNumber)
'''
''' ErrorService.ThrowError("Source")
''' ErrorService.Handle(Err)
''' </code>
''' </example>
Option Explicit

Public Const PathNotSet As String = "PathNotSet"
Public Const PathNotFound As String = "PathNotFound"
Public Const MissingSheet As String = "MissingSheet"

''' <summary>
''' Возвращает объект, содержащий пользовательские ошибки.
''' </summary>
''' <returns>
''' Объект Scripting.Dictionary, представляющий пользовательские ошибки.
''' </returns>
''' <example>
''' <code>
''' Dim userErrors As Object = ErrorService.GetUserErrors()
''' </code>
''' </example>
Private Function GetUserErrors() As Object
    Dim Buffer As Object: Set Buffer = CreateObject("Scripting.Dictionary")

    Buffer(GetErrorNumber(PathNotSet)) = PathNotSet
    Buffer(GetErrorNumber(PathNotFound)) = PathNotFound
    Buffer(GetErrorNumber(MissingSheet)) = MissingSheet

    Set GetUserErrors = Buffer
End Function

''' <summary>
''' Возвращает числовое представление ошибки по заданному имени.
''' </summary>
''' <param name="Name">
''' Строка, представляющая имя ошибки.
''' </param>
''' <returns>
''' Числовое представление ошибки.
''' </returns>
''' <example>
''' <code>
''' Dim errorNumber As Long = ErrorService.GetErrorNumber("ColumnNotFound")
''' </code>
''' </example>
Public Function GetErrorNumber(ByVal Name As String) As Long
    Dim Number As Long
    Dim i As Integer
    For i = 1 To Strings.Len(Name)
        Number = Number + Strings.Asc(Strings.Mid(Name, i, 1))
    Next

    GetErrorNumber = Number
End Function

''' <summary>
''' Проверяет, является ли указанное число ошибкой пользователя.
''' </summary>
''' <param name="Number">
''' Число, представляющее ошибку.
''' </param>
''' <returns>
''' True, если число является ошибкой пользователя, в противном случае - False.
''' </returns>
''' <example>
''' <code>
''' Dim isUserError As Boolean = ErrorService.IsUserError(errorNumber)
''' </code>
''' </example>
Public Function IsUserError(ByVal Number As Long) As Boolean
    Dim UserErrors As Object: Set UserErrors = GetUserErrors()

    IsUserError = UserErrors.Exists(Number)
End Function

''' <summary>
''' Создает исключение с указанным источником.
''' </summary>
''' <remarks>
''' Если источник совпадает с название проекта, то в качестве источника устанавливается переданная строка Source.
''' </remarks>
''' <param name="Source">
''' Строка, представляющая источник исключения.
''' </param>
''' <example>
''' <code>
''' ErrorService.ThrowError "Source"
''' </code>
''' </example>
Public Sub ThrowError(ByVal Source As String)
    If Err.Source = Constants.ProjectName Then
        Err.Source = Source
    End If

    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

''' <summary>
''' Обрабатывает исключение и выводит сообщение об ошибке.
''' </summary>
''' <param name="Err">
''' Объект ErrObject, представляющий исключение.
''' </param>
''' <example>
''' <code>
''' ErrorService.Handle Err
''' </code>
''' </example>
Public Sub Handle(ByRef Err As ErrObject)
    Dim Msg As String: Msg = Err.Description
    Dim Style As VbMsgBoxStyle: Style = vbExclamation

    If Err.Source = Constants.ProjectName Or _
       Not ErrorService.IsUserError(Err.Number) Then
        Msg = "Неизвестная ошибка.\\n" & _
        "Если Вы не знаете как ее исправить, обратитесь к разработчикам\\n\\n" & _
        "#{0}\\nОшибка: {1}\\nИсточник: {2}"
        Style = vbCritical
    End If

    MsgBox FString(Msg, Err.Number, Err.Description, Err.Source), Style, Constants.AppNameRu
    Debug.Print FString("{0} | ERR | {1}", Time, Err.Source)
End Sub