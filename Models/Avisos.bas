Attribute VB_Name = "Avisos"
Public Sub Vazio()

    MsgBox "O VALOR N�O PODE SER VAZIO.", vbExclamation, "ATEN��O!"
    
End Sub

Public Sub Invalido()
    
    MsgBox "VALOR INV�LIDO.", vbCritical, "ERRO"
    
End Sub

Public Sub MsgReloteado()

    MsgBox "N�o � necess�rio registrar endere�o para reloteamento, apenas o ID."
    
End Sub

Public Sub MsgSaida()

    MsgBox "N�o � necess�rio registrar endere�o para retirada, apenas o ID."
    
End Sub

