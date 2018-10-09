'registra um endereço'
Public Sub registra_endereco()

    ActiveSheet.Range("E1") = Range("A4").Value
    
End Sub

'registra um ID'
Public Sub registra_id()

    ActiveSheet.Range("D1") = Range("A4").Value
    Call registra_movimento
    

End Sub


'registra movimento'
Public Sub registra_movimento()

    ActiveSheet.Range("F1") = Now()
    Range("D1:G1").Select
    Selection.Copy
    Range("D5").Select
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False

End Sub

