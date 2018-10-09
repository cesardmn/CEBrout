'Tratamentos e validações de inputs'
    Public Sub TrataIn()

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

        'Validações'
            If Range("A1") = "" Then
                MsgBox "REGISTRE UM ENDEREÇO!"
                
            ElseIf Range("A3") = "EN" Then
                Call registra_endereco
            
            ElseIf Range("A3") = "ID" Then
                If Range("E1") = "" Then    'Verifica se tem endereço'
                    MsgBox "REGISTRE UM ENDEREÇO!"
                Else
                    Call registra_id
                End If
         
            Else
                MsgBox "INSIRA UM ENDEREÇO VÁLIDO."
            End If
        
    End Sub

Public Sub TrataOut()

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

        'Validações'
            If Range("A1") = "" Then
                MsgBox "INFORME UM ID PARA RETIRADA!"
            
            ElseIf Range("A3") = "ID" Then
                If Range("E1") = "" Then    'Verifica se tem endereço'
                    MsgBox "REGISTRE UM ENDEREÇO!"
                Else
                    Call registra_id
                End If
         
            Else
                MsgBox "INSIRA UM ID VÁLIDO."
            End If
        
    End Sub
