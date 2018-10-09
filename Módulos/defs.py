
'Apaga último registro'
Public Sub del_ur()
Application.ScreenUpdating = False
    ActiveSheet.Range("D5:G5").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    ActiveWorkbook.Save
Application.ScreenUpdating = True
End Sub

'Habilita inicio de um registro'
Public Sub iniciar_leitura()
Application.ScreenUpdating = False
    If ActiveSheet.Range("E1").Value = "RETIRADO" Then
        Range("D1, F1").Select
        Selection.ClearContents
    
    ElseIf Range("E1").Value = "RELOTEADO" Then
        Range("D1, F1").Select
        Selection.ClearContents
    
    Else
        Range("D1:F1").Select
        Selection.ClearContents
    End If

    Range("A1").Select
    ActiveWorkbook.Save
Application.ScreenUpdating = True
End Sub


'Apaga todos os registros da folha'
Public Sub limpar_registros()
Application.ScreenUpdating = False
    
    ActiveSheet.Range("D5").Select
    
    If Range("D5").Value <> "" Then
        Rows("5:5").Select
        
        If Range("D6").Value <> "" Then
            Range(Selection, Selection.End(xlDown)).Select
        End If
        Selection.Delete Shift:=xlUp

    End If
    Call iniciar_leitura

Application.ScreenUpdating = True
End Sub


'Salvar registros'
Public Sub salvar_registros()
    Dim PlanilhaAtiva As String
    PlanilhaAtiva = ActiveSheet.Name
    ActiveSheet.Range("D5").Select
    
    If Range("D5").Value <> "" Then
        Range(Selection, Selection.End(xlToRight)).Select
        
        If Range("D6") <> "" Then
            Range(Selection, Selection.End(xlDown)).Select
        End If
        Selection.Copy
        
        Sheets("log").Select
        Range("A2").Select
        Selection.Insert Shift:=xlDown
        Application.CutCopyMode = False
        Range("A1").Select

        Sheets(PlanilhaAtiva).Select
        Selection.ClearContents


        MsgBox "REGISTROS SALVOS EM LOG!"
    Else
        Range("A1").Select
        MsgBox "NÃO HÁ REGISTROS PARA SALVAR."
    End If
    Call iniciar_leitura
    
End Sub

'Exportar log'

    'Insere registros no log'
    Public Sub insert_log()
        Dim PlanilhaAtiva As String
        PlanilhaAtiva = ActiveSheet.Name
        ActiveSheet.Range("D5").Select
        
        If Range("D5").Value <> "" Then
            Range(Selection, Selection.End(xlToRight)).Select
            
            If Range("D6") <> "" Then
                Range(Selection, Selection.End(xlDown)).Select
            End If
            Selection.Copy
            
            Sheets("log").Select
            Range("A2").Select
            Selection.Insert Shift:=xlDown
            Application.CutCopyMode = False
            Range("A1").Select

            Sheets(PlanilhaAtiva).Select
            Selection.ClearContents
            Range("A1").Select
        End If
        Call iniciar_leitura
    End Sub

    Public Sub export_log()
    Application.ScreenUpdating = False
    EnableEvents = False

        Sheets("entrada").Select
            Call insert_log

        Sheets("relot").Select
            Call insert_log

        Sheets("saida").Select
        Call insert_log

        Sheets("transfer").Select
        Call insert_log

        Call ExportCSV
        
    EnableEvents = True
    Application.ScreenUpdating = True
    End Sub
