VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_consultas 
   Caption         =   "Consultas"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3930
   OleObjectBlob   =   "frm_consultas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_consultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String

Private Sub btn_base_Click()

Application.Visible = True
frm_consultas.Hide

SQL = "SELECT * FROM tb_id ORDER BY REGISTRO DESC"
ConsultarBase (SQL)

End Sub

Private Sub btn_entradas_Click()

Application.Visible = True
frm_consultas.Hide

SQL = "SELECT ID, ENDERECO, REGISTRO, MOVIMENTO FROM tb_log WHERE MOVIMENTO IN ('ENTRADA') ORDER BY REGISTRO DESC"
ConsultarBase (SQL)

End Sub

Private Sub btn_estoque_Click()

Application.Visible = True
frm_consultas.Hide

SQL = "SELECT ID, ENDERECO, REGISTRO, MOVIMENTO FROM tb_id WHERE MOVIMENTO NOT IN ('RELOTEAMENTO', 'SAÍDA') ORDER BY REGISTRO DESC"
ConsultarBase (SQL)

End Sub

Private Sub btn_reloteamentos_Click()
Application.Visible = True
frm_consultas.Hide

SQL = "SELECT ID, ENDERECO, REGISTRO, MOVIMENTO FROM tb_log WHERE MOVIMENTO IN ('RELOTEAMENTO') ORDER BY REGISTRO DESC"
ConsultarBase (SQL)
End Sub

Private Sub btn_retirados_Click()

Application.Visible = True
frm_consultas.Hide

SQL = "SELECT ID, ENDERECO, REGISTRO, MOVIMENTO FROM tb_log WHERE MOVIMENTO IN ('SAÍDA') ORDER BY REGISTRO DESC"
ConsultarBase (SQL)

End Sub

Private Sub btn_sair_Click()

frm_consultas.Hide
frm_registros.Show


End Sub

Private Sub btn_transferidos_Click()
Application.Visible = True
frm_consultas.Hide

SQL = "SELECT ID, ENDERECO, REGISTRO, MOVIMENTO FROM tb_log WHERE MOVIMENTO IN ('TRANSFERÊNCIA') ORDER BY REGISTRO DESC"
ConsultarBase (SQL)
End Sub



Private Sub UserForm_QueryClose _
  (Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

