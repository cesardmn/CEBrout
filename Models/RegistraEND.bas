Attribute VB_Name = "RegistraEND"
'ValidaEND'
    Public Sub ValidarEND(endereco As String)
        
        endereco = Right(endereco, Len(endereco) - 3)
        If endereco <> "" Then
            Call RegistrarEND(endereco)
            
        Else
            Call Vazio
            Exit Sub
        
        End If

    End Sub


'RegistraEND'
    Public Sub RegistrarEND(endereco As String)
      
        If frm_registros.dsp_lbl_end = "RELOTEADO" Then
            Call MsgReloteado
        
        ElseIf frm_registros.dsp_lbl_end = "RETIRADO" Then
            Call MsgSaida
            
        Else
            frm_registros.dsp_lbl_end.Caption = endereco
        
        End If
    End Sub

