Option Explicit
Private Sub btexct_Click()
Application.ScreenUpdating = False

'variáveis'
Dim W               As Worksheet
Dim WNew            As Workbook
Dim ArqParaAbrir    As Variant
Dim A               As Integer
Dim NomeArquivo     As String

'capturar arquivos para tratamento'
ArqParaAbrir = Application.GetOpenFilename("Arquivo do Excel(*.xls), *.xl*", _
            Title:="Selecione os arquivos", _
            MultiSelect:=True)

If Not IsArray(ArqParaAbrir) Then
    If ArqParaAbrir = "" Or ArqParaAbrir = False Then
        MsgBox "Operação cancelada - Não há arquivos selcionados ou válidos!"
        Exit Sub
        
    End If
End If



'iniciar importaçao'
Set W = Sheets("Plan1")
Columns("A:D").Select
Selection.Delete Shift:=xlToLeft
Range("A1").Select


'loop importação'

For A = LBound(ArqParaAbrir) To UBound(ArqParaAbrir)

    NomeArquivo = ArqParaAbrir(A)
    
    Application.Workbooks.Open (NomeArquivo)
    Set WNew = ActiveWorkbook
    ActiveSheet.Range("A2").CurrentRegion.Select
    
    Selection.Copy Destination:=W.Cells(W.Rows.Count, 1).End(xlUp).Offset(1, 0)
    
    Application.DisplayAlerts = False
    
        ActiveWorkbook.Close savechanges:=False
    Application.DisplayAlerts = True
    
    MsgBox "import ok"
    
Next A



Application.ScreenUpdating = True
End Sub
