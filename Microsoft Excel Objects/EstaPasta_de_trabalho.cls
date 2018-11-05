VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

Application.Visible = False
frm_registros.Show

    frm_registros.lbl_entrada.BackColor = &HE0E0E0
    
    frm_registros.lbl_registrar.Caption = "Registrar Entrada"
    frm_registros.dsp_lbl_reg.Caption = ""
    frm_registros.dsp_lbl_id.Caption = ""
    frm_registros.dsp_lbl_end.Caption = ""
    frm_registros.dsp_lbl_mov.Caption = "ENTRADA"
    
    frm_registros.lbl_registrar.ForeColor = &H8000&
    
    frm_registros.lbl_transferencia.BackColor = &HFFFFFF
    frm_registros.lbl_reloteamento.BackColor = &HFFFFFF
    frm_registros.lbl_saida.BackColor = &HFFFFFF
    
    frm_registros.txt_entrada.Visible = True
    frm_registros.txt_entrada.SetFocus

End Sub
