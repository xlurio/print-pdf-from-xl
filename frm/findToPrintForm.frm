VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} findToPrintForm 
   Caption         =   "Caminho para imprimir"
   ClientHeight    =   1440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "findToPrintForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "findToPrintForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub okBtn_Click()
    If pathTxtBox.Value <> "" Then
        findToPrintForm.Hide
    ElseIf pathTxtBox.Value = "" Then
        MsgBox ("Favor informar um caminho para procurar os desenhos!")
    Else
        MsgBox ("Erro: pathTxtBox.Value = " & pathTxtBox.Value)
    End If
End Sub

Private Sub searchBtn_Click()
    pathTxtBox.Value = GetFolder(ActiveWorkbook.Path)
End Sub

Function GetFolder(strPath As String) As String

    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing

End Function
