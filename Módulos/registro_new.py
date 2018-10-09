Public Sub registos_novo()

Dim entrada As String
acao = Sheets("entrada").Range("G1").Value


'valida ação'
    If acao = "TRANSFERÊNCIA" or acao = "ENTRADA" Then

        Call valida_in

    ElseIf acao = "SAÍDA" or acao = "RELOTEAMENTO" Then
        
        Call valida_out

    Else
        Sheets("home").Select
        Exit Sub 
    End If



    'Tratamento'

        'Capitaliza, retira ";" e retira último caracter'
        Range("A2").Select
        ActiveCell.FormulaR1C1 = _
            "=UPPER(SUBSTITUTE(LEFT(R[-1]C,LEN(R[-1]C)-1),"";"",""VAZIO""))"

        'Informa os 2 primeiros caracteres'
        Range("A3").Select
        ActiveCell.FormulaR1C1 = "=LEFT(R[-1]C,2)"

        'Retorna string tratada'
        Range("A4").Select
        ActiveCell.FormulaR1C1 = _
            "=IF(R[-1]C=""id"",R[-2]C,IF(R[-1]C=""en"",RIGHT(R[-2]C,LEN(R[-2]C)-3),""VALOR INVÁLIDO""))"
        Range("A5").Select









    'Validação'

        'Valida entrada / transferência'
        If Range("G1").Value = "ENTRADA" Or Range("G1").Value = "TRANSFERÊNCIA" Then
        
            MsgBox "ok"
        Else
            MsgBox "deu ruim"
        End If

End Sub