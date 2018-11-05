Attribute VB_Name = "SQL"
'---------------------------------------------------------------------------------------------------------------------------------------
'VARIÁVEIS GLOBAIS

Option Explicit
Dim Conexao             As ADODB.Connection
Dim RS                  As ADODB.Recordset
Dim FD                  As ADODB.Field
Dim SQL, csRS           As String
Dim W                   As Worksheet





'---------------------------------------------------------------------------------------------------------------------------------------
'CONEXÕES

Public Sub Conectar()
    Dim caminhoDB As String
    caminhoDB = ActiveWorkbook.Path & "\broutdb.accdb"
    caminhoDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & caminhoDB & "; Persist Security Info=False" ';Jet Oledb:DataBase password=ff123"
        
    Set Conexao = New ADODB.Connection
    Conexao.Open caminhoDB
End Sub

Public Sub Desconectar()
    If Not Conexao Is Nothing Then
        Conexao.Close
        Set Conexao = Nothing
    End If
End Sub

Public Sub DesconectarRS()
    If Not RS Is Nothing Then
        RS.Close
        Set RS = Nothing
    End If
End Sub






'---------------------------------------------------------------------------------------------------------------------------------------
'CONSULTAS

Public Sub ConsultarBase(SQL As String)
Dim col As Long

'Conectar
Set RS = New ADODB.Recordset
Conectar

'Consulta
RS.Open SQL, Conexao


'Prepara a planilha para receber dados
    Set W = Sheets("registros")
    W.Select
    W.Range("A2").Select
    W.Range("A2").CurrentRegion.ClearContents   'limpa dados anteriores
    col = 1                                     'marco inicial para inserção dos dados

 'Cabeçalho
    If RS.EOF = False Then
        W.Range("A1").Select
        For Each FD In RS.Fields
            With W.Cells(1, col)
                .Value = FD.Name
                .Font.Bold = True
             End With
            col = col + 1
        Next FD
        
        W.Cells(2, 1).CopyFromRecordset RS      'insere consulta na planilha
        W.UsedRange.EntireColumn.AutoFit        'ajusta tamanho das colunas
    End If
    
  'Desconectar
    DesconectarRS
    Desconectar
End Sub





'---------------------------------------------------------------------------------------------------------------------------------------
'REGISTROS


Public Sub Registrar_Log()

Set RS = New ADODB.Recordset
Conectar
RS.Open "tb_log", Conexao, adOpenKeyset, adLockOptimistic
RS.AddNew

RS!id = frm_registros.dsp_lbl_id.Caption
RS!endereco = frm_registros.dsp_lbl_end.Caption
RS!registro = frm_registros.dsp_lbl_reg.Caption
RS!Movimento = frm_registros.dsp_lbl_mov.Caption
RS.Update

DesconectarRS
Desconectar

End Sub

Public Sub Registrar_ID()
    
    Conectar
        SQL = "DELETE FROM tb_id WHERE ID = '" & frm_registros.dsp_lbl_id.Caption & "'"
        Conexao.Execute SQL
    Desconectar

    Set RS = New ADODB.Recordset
    Conectar
        RS.Open "tb_id", Conexao, adOpenKeyset, adLockOptimistic
        RS.AddNew
        
        RS!id = frm_registros.dsp_lbl_id.Caption
        RS!endereco = frm_registros.dsp_lbl_end.Caption
        RS!registro = frm_registros.dsp_lbl_reg.Caption
        RS!Movimento = frm_registros.dsp_lbl_mov.Caption
        RS.Update
        DesconectarRS
    Desconectar

End Sub

