Attribute VB_Name = "TratamentoValidacao"
'TrataIn'
    Public Sub TratarEntrada(entrada As String)

        'Retira caracteres especiais'
            entrada = Replace(entrada, ";", "")
            entrada = Replace(entrada, "-", "")
            entrada = Replace(entrada, " ", "")
        
        'Retira "a" final'
            If Right(entrada, 1) = "a" Then
                entrada = Left(entrada, Len(entrada) - 1)
            End If
                    
        'Capitaliza'
           entrada = UCase(entrada)
           
           Call ValidarTipo(entrada)
        
    End Sub
    
    
    'ValidaTipo'
    Public Sub ValidarTipo(entrada As String)

        
        If Left(entrada, 2) = "ID" Then
            Call ValidarID(entrada)

        ElseIf Left(entrada, 3) = "END" Then
            Call ValidarEND(entrada)
            
        ElseIf entrada = "" Then
            Call Vazio
            Exit Sub
            
        Else
            Call Invalido

        End If
    End Sub
