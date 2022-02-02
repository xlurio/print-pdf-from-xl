Attribute VB_Name = "Print_Especified_Item"
' Main subprocess
Sub PrintItem()
    ' Declare variables
    Dim RetVal, drawFolder As String, drawFile As String
    Dim drawName As String, c As Range, drawPath As String
    Dim sumatraPath As String, drawZip As String, unzipperPath As String
    Dim tempFolderPath As String
    ' Get drawings path
    sumatraPath = Workbooks("PrintItem-Sumatra.xlam").Path & "\SumatraPDFPortable\SumatraPDFPortable.exe"
    unzipperPath = Workbooks("PrintItem-Sumatra.xlam").Path & "\7-Zip\7z.exe"
    tempFolderPath = Workbooks("PrintItem-Sumatra.xlam").Path & "\temp\"
    drawFolder = getDrawFolder()
    
    For Each c In Selection.Rows
    
        drawName = Trim(Replace(CStr(c.Value), "(01 RH e 01 LH)", ""))
        drawName = Trim(Replace(CStr(drawName), "_LH", ""))
        
        If Right(drawFolder, 1) <> "\" Then
            drawFolder = drawFolder & "\"
        End If
        
        drawFile = Dir(drawFolder & "*" & drawName & "*.pdf")
        Debug.Print drawFolder & drawFile
        
        If Len(drawFile) <= 0 Then
            drawZip = Dir(drawFolder & "*" & drawName & "*.zip")
            Debug.Print drawFolder & drawZip
            
            If Len(drawZip) <= 0 Then
                MsgBox ("Erro: " & drawName & " n�o encontrado na pasta especificada (" & drawFolder & ")")
            End If
            
            Do While Len(drawZip) > 0
                drawPath = drawFolder & drawZip
                RetVal = Shell(unzipperPath & " e " & """" & drawPath & """" & " -o" & """" & tempFolderPath & """", 1)
                drawPath = tempFolderPath & drawFile
                Application.Wait (Now + TimeValue("00:00:05"))
                drawFile = Dir(tempFolderPath & "*pdf")
                printFilesLoop drawFile, tempFolderPath, RetVal, sumatraPath, drawPath, True
                On Error Resume Next
                    drawZip = Dir
                On Error GoTo 0
            Loop
            MsgBox ("Desenhos impressos!")
            Exit Sub
        End If
        ' Print drawing
        printFilesLoop drawFile, drawFolder, RetVal, sumatraPath, drawPath
    Next
    
    MsgBox ("Desenhos impressos!")
End Sub

Function getDrawFolder() As String
    ' Declare function variables
    Dim pathTxtFile As String, fileId As Integer
    ' Get drawing path from txt file
    If Len(Dir(ActiveWorkbook.Path & "\toprintpath.txt")) > 0 Then
        pathTxtFile = ActiveWorkbook.Path & "\toprintpath.txt"
        MsgBox ("Utilizando caminho do arquivo: " & pathTxtFile)
        fileId = FreeFile
        Open pathTxtFile For Input As #fileId
        getDrawFolder = Trim(CStr(Input(LOF(fileId), fileId)))
        Close #fileId
        getDrawFolder = Replace(getDrawFolder, "é", ChrW(233))
        getDrawFolder = Replace(getDrawFolder, "â", ChrW(226))
        getDrawFolder = Replace(getDrawFolder, "ç", ChrW(231))
        getDrawFolder = Replace(getDrawFolder, "ã", ChrW(227))
        Exit Function
    ElseIf Len(Dir(ActiveWorkbook.Path & "\toprintpath.txt")) <= 0 Then
        ' Get drawing path from form
        findToPrintForm.Show
        getDrawFolder = Trim(CStr(findToPrintForm.pathTxtBox.Value))
        Unload findToPrintForm
        Exit Function
    End If
    MsgBox ("Erro: Dir(ActiveWorkbook.Path & '\toprintpath.txt') = " & Dir(ActiveWorkbook.Path & "\toprintpath.txt"))
End Function

Sub printFilesLoop(drawFile As String, drawFolder As String, RetVal, sumatraPath As String, drawPath As String, Optional toDelete As Boolean = False)
    Do While Len(drawFile) > 0
        drawPath = drawFolder & drawFile
        RetVal = Shell(sumatraPath & " -print-settings " & """" & "fit, paper=A4" & """" & " -print-to-default " & """" & drawPath & """", 1)
        Application.Wait (Now + TimeValue("00:00:05"))
        drawFile = Dir
        If toDelete Then
            Kill drawPath
        End If
    Loop
End Sub
