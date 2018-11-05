Attribute VB_Name = "RegistraID"
'ValidaID'
    Public Sub ValidarID(id As String)

        If Right(id, Len(id) - 2) <> "" Then
            Call RegistrarID(id)
            
        Else
            Call Vazio
            Exit Sub
        
        End If
    
    End Sub


'RegistraID'
    Public Sub RegistrarID(id As String)
    Dim endereco As String
    endereco = frm_registros.dsp_lbl_end.Caption
    
    If endereco <> "" Then
        frm_registros.dsp_lbl_id.Caption = id
        frm_registros.dsp_lbl_reg.Caption = Now()
        Call Registrar_Log
        Call Registrar_ID
        
    Else
        MsgBox "INSIRA UM ENDEREÇO."
        Exit Sub
    End If

End Sub



