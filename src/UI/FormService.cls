VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ReportMergerProject.src.UI"
''' <summary>
''' Класс FormService предоставляет функциональность для управления формами в приложении.
''' </summary>
''' <remarks>
''' Класс FormService предоставляет методы для конструирования формы, установки тэгов на элементы управления, закрытия формы и загрузки данных.
''' </remarks>
''' <example>
''' <code>
''' Dim form As New MSForms.UserForm
''' Dim controller As New FormService
''' controller.Constructor(form, Tags)
''' controller.SetTags()
''' controller.CloseForm()
''' controller.LoadData()
''' </code>
''' </example>
Option Explicit

Private Type TFormController
    Tags As ArrayList
    Form As MSForms.UserForm
End Type

Private this As TFormController

''' <summary>
''' Инициализирует объект FormService с указанной формой и заданными тэгами для элементов управления.
''' </summary>
''' <param name="Form">
''' Объект MSForms.UserForm, представляющий форму.
''' </param>
''' <param name="Tags">
''' Массив Variant, представляющий тэги для элементов управления.
''' </param>
''' <example>
''' <code>
''' Dim controller As New FormService
''' controller.Constructor(form, Tags)
''' </code>
''' </example>
Public Sub Constructor(ByRef Form As MSForms.UserForm, ParamArray Tags() As Variant)
    Set this.Form = Form

    Set this.Tags = New ArrayList
    Dim TagsValue As Variant: TagsValue = Tags
    If IsArray(TagsValue(0)) Then this.Tags.AddRange TagsValue(0) Else this.Tags.AddRange TagsValue
    SetTags
End Sub

''' <summary>
''' Устанавливает тэги на элементы управления в соответствии с заданными значениями.
''' </summary>
Public Sub SetTags()
    Dim Control As MSForms.Control
    For Each Control In this.Form.Controls
        Dim ControlName As String: ControlName = Strings.Replace(Control.Name, TypeName(Control), Empty)
        If this.Tags.Contains(ControlName, Compare:=vbTextCompare) Then Control.Tag = ControlName
    Next
End Sub

''' <summary>
''' Закрывает форму.
''' </summary>
Public Sub CloseForm()
    Unload this.Form
End Sub

''' <summary>
''' Загружает данные на форму, включая флажки, выпадающие списки, текстовые поля и список.
''' </summary>
Public Sub LoadData()
    LoadCheckBoxes
    LoadComboBoxes
    LoadTextBoxes
    LoadListBoxes
End Sub

''' <summary>
''' Загружает значения в текстовые поля на форме.
''' </summary>
Public Sub LoadTextBoxes()
    Dim Control As MSForms.Control
    For Each Control In this.Form.Controls
        If IsTextBox(Control) Then
            Control.Value = Config.GetValue(Control.Tag, Empty)
            Control.ControlTipText = Control.Value
        End If
    Next
End Sub

''' <summary>
''' Загружает значения в флажки на форме.
''' </summary>
Public Sub LoadCheckBoxes()
    Dim Control As MSForms.Control
    For Each Control In this.Form.Controls
        If IsCheckBox(Control) Then Control.Value = Config.GetValue(Control.Tag, False)
    Next
End Sub

''' <summary>
''' Загружает значения в список на форме.
''' </summary>
''' <param name="Delimiter">Разделитель для разбиения строки на элементы списка</param>
Public Sub LoadListBoxes(Optional ByVal Delimiter As String = ";")
    Dim Control As MSForms.Control
    For Each Control In this.Form.Controls
        If IsListBox(Control) Then
            Control.List = Strings.Split(Config.GetValue(Control.Tag), Delimiter)
        End If
    Next
End Sub

''' <summary>
''' Загружает значения в список на форме.
''' </summary>
''' <param name="Delimiter">Разделитель для разбиения строки на элементы списка</param>
Public Sub LoadComboBoxes(Optional ByVal Delimiter As String = ";")
    Dim Control As MSForms.Control
    For Each Control In this.Form.Controls
        If IsComboBox(Control) Then
            Control.List = Strings.Split(Config.GetValue(Control.Tag), Delimiter)
            Dim FirstValue As String
            If Control.ListCount > -1 Then
                FirstValue = Control.List(0)
            End If
            Control.Value = Config.GetValue(Control.Tag & "Sel", FirstValue)
        End If
    Next
End Sub

''' <summary>
''' Устанавливает путь к файлу в текстовом поле на форме.
''' </summary>
''' <param name="TagName">Тег, связанный с текстовым полем на форме</param>
''' <param name="Caption">Заголовок для диалогового окна выбора файла</param>
Public Sub SetFilePath(ByVal TagName As String, Optional ByVal Caption As String = "Укажите путь к файлу '{0}'")
    Dim FileName As String: FileName = Strings.Replace(GetLabelByTag(TagName).Caption, ":", Empty)
    Dim Path As Variant
    Path = GetFilePath(FString(Caption, FileName), _
                        GetLastFolder(), _
                        Array("Excel файлы", "*.xls*"))
    If IsFalse(Path) Then Exit Sub

    Dim PathBox As MSForms.TextBox: Set PathBox = GetTextBoxByTag(TagName)
    PathBox.Value = Path
    PathBox.ControlTipText = Path
    SetBorderColor vbActiveBorder, Array(PathBox)
    UpdateLastFolder Path
End Sub

''' <summary>
''' Устанавливает путь к папке в текстовом поле на форме.
''' </summary>
''' <param name="TagName">Тег, связанный с текстовым полем на форме</param>
''' <param name="Caption">Заголовок для диалогового окна выбора папки</param>
Public Sub SetFolderPath(ByVal TagName As String, Optional ByVal Caption As String = "Укажите путь к папке с файлами '{0}'")
    Dim Path As Variant
    Path = GetFolderPath(FString(Caption, Strings.Replace(GetLabelByTag(TagName).Caption, ":", Empty)), _
                            GetLastFolder())
    If IsFalse(Path) Then Exit Sub

    Dim PathBox As MSForms.TextBox: Set PathBox = GetTextBoxByTag(TagName)
    PathBox.Value = Path
    PathBox.ControlTipText = Path
    SetBorderColor vbActiveBorder, Array(PathBox)
    UpdateLastFolder Path
End Sub

''' <summary>
''' Возвращает последнюю сохраненную папку.
''' </summary>
''' <returns>Строка с путем к последней сохраненной папке</returns>
Private Function GetLastFolder() As String
    Dim Folder As String
    Folder = Config.GetValue(TypeName(this.Form) & "LastFolder", ThisWorkbook.Path)
    If Strings.Right(Folder, 1) <> NodePath.Sep Then
        Folder = Folder & NodePath.Sep
    End If
    GetLastFolder = Folder
End Function

''' <summary>
''' Обновляет последнюю сохраненную папку.
''' </summary>
''' <param name="Path">Путь к файлу или папке</param>
Private Sub UpdateLastFolder(ByVal Path As String)
    Dim FS As FS: Set FS = New FS
    Dim Folder As String
    If FS.IsDir(Path) Then
        Folder = Path
    Else
        Folder = NodePath.DirName(Path)
    End If
    Config.SetValue TypeName(this.Form) & "LastFolder", Folder
End Sub

''' <summary>
''' Сохраняет данные с формы.
''' </summary>
Public Sub SaveData()
    SaveCheckBoxes
    SaveTextBoxes
    SaveListBoxes
    SaveComboBoxes
End Sub

''' <summary>
''' Сохраняет значения из флажков на форме.
''' </summary>
Public Sub SaveCheckBoxes()
    Dim Control As MSForms.Control
    For Each Control In this.Form.Controls
        If IsCheckBox(Control) Then Config.SetValue Control.Tag, Control.Value
    Next
End Sub

''' <summary>
''' Сохраняет значения из текстовых полей на форме.
''' </summary>
Public Sub SaveTextBoxes()
    Dim Control As MSForms.Control
    For Each Control In this.Form.Controls
        If IsTextBox(Control) Then Config.SetValue Control.Tag, Control.Value
    Next
End Sub

''' <summary>
''' Сохраняет значения из списков на форме.
''' </summary>
''' <param name="Delimiter">Разделитель для объединения элементов списка в строку</param>
Public Sub SaveListBoxes(Optional ByVal Delimiter As String = ";")
    Dim Control As MSForms.Control
    For Each Control In this.Form.Controls
        If IsListBox(Control) Then Config.SetValue Control.Tag, Strings.Join(Control.List, Delimiter)
    Next
End Sub

''' <summary>
''' Сохраняет значения из списков на форме.
''' </summary>
''' <param name="Delimiter">Разделитель для объединения элементов списка в строку</param>
Public Sub SaveComboBoxes(Optional ByVal Delimiter As String = ";")
    Dim Control As MSForms.Control
    For Each Control In this.Form.Controls
        If IsComboBox(Control) Then
            Config.SetValue Control.Tag, Strings.Join(Control.List, Delimiter)
            Config.SetValue Control.Tag & "Sel", Control.Value
        End If
    Next
End Sub

''' <summary>
''' Проверяет правильность заполнения текстовых полей.
''' </summary>
''' <param name="Controls">Массив текстовых полей или параметр массива</param>
''' <returns>Значение True, если все текстовые поля заполнены правильно, в противном случае - False</returns>
Public Function ValidatePaths(ParamArray Controls() As Variant) As Boolean
    Dim Incorrect As ArrayList: Set Incorrect = New ArrayList
    Dim ControlsValue As Variant: ControlsValue = Controls
    If IsArray(Controls(0)) Then ControlsValue = Controls(0)

    Dim Control As Variant
    For Each Control In ControlsValue
        If Not IsTextBox(Control) Then GoTo Continue
        If IsValide(Control) Then GoTo Continue
        Incorrect.Add Control
Continue:
    Next

    If Incorrect.Count > 0 Then
        SetBorderColor vbRed, Incorrect.ToArray()
        Exit Function
    End If

    ValidatePaths = True
End Function

''' <summary>
''' Устанавливает цвет рамки элементов управления.
''' </summary>
''' <param name="Color">Цвет рамки в формате Long</param>
''' <param name="Controls">Массив элементов управления или параметр массива</param>
Public Sub SetBorderColor(ByVal Color As Long, ParamArray Controls() As Variant)
    Dim ControlsValue As Variant: ControlsValue = Controls
    If IsArray(Controls(0)) Then ControlsValue = Controls(0)

    Dim Control As Variant
    For Each Control In ControlsValue
        Control.BorderColor = Color
    Next
End Sub

''' <summary>
''' Возвращает текстовое поле по заданному тегу.
''' </summary>
''' <param name="TagName">Тег текстового поля</param>
''' <returns>Текстовое поле с заданным тегом</returns>
Private Function GetTextBoxByTag(ByVal TagName As String) As MSForms.TextBox
    Dim TextBox As MSForms.Control
    For Each TextBox In this.Form.Controls
        If IsTextBox(TextBox) And TextBox.Tag = TagName Then Exit For
    Next

    Set GetTextBoxByTag = TextBox
End Function

''' <summary>
''' Возвращает метку по заданному тегу.
''' </summary>
''' <param name="TagName">Тег метки</param>
''' <returns>Метка с заданным тегом</returns>
Private Function GetLabelByTag(ByVal TagName As String) As MSForms.Label
    Dim Label As MSForms.Control
    For Each Label In this.Form.Controls
        If IsLabel(Label) And Label.Tag = TagName Then Exit For
    Next

    Set GetLabelByTag = Label
End Function

''' <summary>
''' Проверяет, является ли значение пути в текстовом поле допустимым.
''' </summary>
''' <param name="Control">Текстовое поле</param>
''' <returns>True, если значение допустимо, иначе False</returns>
Private Function IsValide(ByVal Control As MSForms.TextBox) As Boolean
    If IsFalse(Control.Value) Then Exit Function

    Dim FS As FS: Set FS = New FS
    If IsFalse(FS.DirExists(Control.Value) Or FS.FileExists(Control.Value)) Then Exit Function

    If FS.IsDir(Control.Value) Then
        IsValide = FS.DirExists(Control.Value)
    Else
        IsValide = FS.FileExists(Control.Value)
    End If
End Function


