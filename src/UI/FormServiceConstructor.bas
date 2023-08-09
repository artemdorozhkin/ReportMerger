Attribute VB_Name = "FormServiceConstructor"
'@Folder "ReportMergerProject.src.UI"
Option Explicit

Public Function NewFormService(ByRef Form As MSForms.UserForm, ParamArray Tags() As Variant) As FormService
    Set NewFormService = New FormService
    NewFormService.Constructor Form, Tags
End Function
