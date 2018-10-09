'EXPORTA CSV'
    Option Explicit
    Public Sub ExportCSV()
    Application.ScreenUpdating = False

        Sheets("log").Select
        If Range("A2") <> "" Then
                   
            Dim MyFileName As String
            Dim CurrentWB As Workbook, TempWB As Workbook
            
            Set CurrentWB = ActiveWorkbook
            ActiveWorkbook.ActiveSheet.UsedRange.Copy
        
            Set TempWB = Application.Workbooks.Add(1)
            With TempWB.Sheets(1).Range("A1")
              .PasteSpecial xlPasteValues
              .PasteSpecial xlPasteFormats
            End With
             
            MyFileName = CurrentWB.Path & "\" & "Log dia_" & Format(Now, "yyyy-mm-dd") & "_hora_" & Format(Now, "hh-mm")
        
            'Application.DisplayAlerts = False '
            TempWB.SaveAs Filename:=MyFileName, FileFormat:=xlNormal, CreateBackup:=False, Local:=True
            TempWB.Close SaveChanges:=False
            'Application.DisplayAlerts = True '
            
            
            Rows("2:2").Select
            If Range("A3") <> "" Then
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Delete Shift:=xlUp
            Else
                Selection.Delete Shift:=xlUp
                
            End If
                Range("A1").Select
                Sheets("home").Select
                MsgBox "LOG EXPORTADO COM SUCESSO"

            Else
                Range("A1").Select
                Sheets("home").Select
                MsgBox "NÃO HÁ REGISTROS PARA EXPORTAR"

        End If
        
    Application.ScreenUpdating = True
    End Sub
