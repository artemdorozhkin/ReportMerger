VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsForm 
   Caption         =   "Настройки"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9180.001
   OleObjectBlob   =   "SettingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ReportMergerProject.src.UI"
Option Explicit

Private Type TSettingsForm
    FormService As FormService
End Type

Private this As TSettingsForm

Private Sub UserForm_Initialize()
    Config.App = Constants.AppNameEn
    Config.Section = Constants.SettingsSection

    Set this.FormService = NewFormService(Me, Constants.ReportsFolderKey, _
                                              Constants.ResultsFolderKey)
    this.FormService.LoadTextBoxes
    SaveCommandButton.Enabled = False
End Sub

Private Sub ReportsFolderCommandButton_Click()
    this.FormService.SetFolderPath ReportsFolderCommandButton.Tag, "Укажите путь к папке с отчетами"
End Sub

Private Sub ResultsFolderCommandButton_Click()
    this.FormService.SetFolderPath ResultsFolderCommandButton.Tag, "Укажите путь к папке для сохранения результата"
End Sub

Private Sub ReportsFolderTextBox_Change()
    SaveCommandButton.Enabled = True
    ReportsFolderTextBox.BorderColor = vbActiveBorder
End Sub

Private Sub ResultsFolderTextBox_Change()
    SaveCommandButton.Enabled = True
    ResultsFolderTextBox.BorderColor = vbActiveBorder
End Sub

Private Sub SendLogsCommandButton_Click()
    LogsReportService.SendReport
End Sub

Private Sub SaveCommandButton_Click()
    If this.FormService.ValidatePaths(ReportsFolderTextBox, ResultsFolderTextBox) Then
        this.FormService.SaveTextBoxes
        SaveCommandButton.Enabled = False
    End If
End Sub

Private Sub CloseCommandButton_Click()
    this.FormService.CloseForm
End Sub
