VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_registros 
   Caption         =   "Registro de movimentos BRbid OUTLET"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10980
   OleObjectBlob   =   "frm_registros.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_registros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim registro As String
Dim vsql As String


Private Sub btn_valida_Enter()

txt_entrada.Visible = True
txt_entrada.SetFocus


End Sub

Private Sub lbl_consultas_Click()

frm_registros.Hide
frm_consultas.Show

End Sub

Private Sub lbl_devel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Application.Visible = True
Sheets("registros").Range("A1").CurrentRegion.ClearContents
frm_registros.Hide

End Sub



Private Sub txt_entrada_AfterUpdate()
    
registro = txt_entrada.Value

    Call TratarEntrada(registro)


txt_entrada.Value = Null
txt_entrada.SetFocus

End Sub



Private Sub lbl_entrada_Click()
    
    lbl_entrada.BackColor = &HE0E0E0
    
    lbl_registrar.Caption = "Registrar Entrada"
    dsp_lbl_reg.Caption = ""
    dsp_lbl_id.Caption = ""
    dsp_lbl_end.Caption = ""
    dsp_lbl_mov.Caption = "ENTRADA"
    
    lbl_registrar.ForeColor = &H8000&
    
    lbl_transferencia.BackColor = &HFFFFFF
    lbl_reloteamento.BackColor = &HFFFFFF
    lbl_saida.BackColor = &HFFFFFF
    
    txt_entrada.Visible = True
    txt_entrada.SetFocus
    

End Sub



Private Sub lbl_transferencia_Click()

    lbl_transferencia.BackColor = &HE0E0E0
    
    lbl_registrar.Caption = "Registrar Transferência"
    dsp_lbl_reg.Caption = ""
    dsp_lbl_id.Caption = ""
    dsp_lbl_end.Caption = ""
    dsp_lbl_mov.Caption = "TRANSFERÊNCIA"
    
    lbl_registrar.ForeColor = &H8080&
    
    lbl_entrada.BackColor = &HFFFFFF
    lbl_reloteamento.BackColor = &HFFFFFF
    lbl_saida.BackColor = &HFFFFFF
    
    txt_entrada.Visible = True
    txt_entrada.SetFocus
    
End Sub



Private Sub lbl_reloteamento_Click()


    lbl_reloteamento.BackColor = &HE0E0E0
    
    lbl_registrar.Caption = "Registrar Reloteamento"
    dsp_lbl_reg.Caption = ""
    dsp_lbl_id.Caption = ""
    dsp_lbl_mov.Caption = "RELOTEAMENTO"
    dsp_lbl_end.Caption = "RELOTEADO"
    lbl_registrar.ForeColor = &H800080
    
    lbl_entrada.BackColor = &HFFFFFF
    lbl_transferencia.BackColor = &HFFFFFF
    lbl_saida.BackColor = &HFFFFFF
    
    txt_entrada.Visible = True
    txt_entrada.SetFocus

End Sub

Private Sub lbl_saida_Click()



    lbl_saida.BackColor = &HE0E0E0
    
    lbl_registrar.Caption = "Registrar Saída"
    dsp_lbl_reg.Caption = ""
    dsp_lbl_id.Caption = ""
    dsp_lbl_end.Caption = "RETIRADO"
    dsp_lbl_mov.Caption = "SAÍDA"
    
    lbl_registrar.ForeColor = &HC0&
    
    lbl_entrada.BackColor = &HFFFFFF
    lbl_transferencia.BackColor = &HFFFFFF
    lbl_reloteamento.BackColor = &HFFFFFF
    
    txt_entrada.Visible = True
    txt_entrada.SetFocus

End Sub




Private Sub UserForm_QueryClose _
  (Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

