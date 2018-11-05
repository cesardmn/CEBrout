Attribute VB_Name = "Avisos"
Public Sub Vazio()

    MsgBox "O VALOR NÃO PODE SER VAZIO.", vbExclamation, "ATENÇÃO!"
    
End Sub

Public Sub Invalido()
    
    MsgBox "VALOR INVÁLIDO.", vbCritical, "ERRO"
    
End Sub

Public Sub MsgReloteado()

    MsgBox "Não é necessário registrar endereço para reloteamento, apenas o ID."
    
End Sub

Public Sub MsgSaida()

    MsgBox "Não é necessário registrar endereço para retirada, apenas o ID."
    
End Sub

