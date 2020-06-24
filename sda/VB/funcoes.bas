Attribute VB_Name = "Modfuncoes"
Dim i As Integer
Dim Rs As ADODB.Recordset
Dim Sql As String

Public Function ExecutaSqlRs(strSql As String) As ADODB.Recordset
  
  'Cria o Objeto ADO
  Dim objRs     As ADODB.Recordset
  Dim objCmd    As ADODB.Command
  'Inicializa os Objetos
  Set objRs = New ADODB.Recordset
  Set objCmd = New ADODB.Command
  'Informo os parâmetros do Objeto
  strStringConexao = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=false;Data Source=" & App.Path & "/Bd/auditoria.mdb" & ";User id=Admin;"
  objCmd.ActiveConnection = strStringConexao
  objCmd.CommandType = adCmdText
  objCmd.CommandText = strSql
  'Executa a query somente para leitura
  objRs.CursorLocation = adUseClient
  objRs.Open objCmd, , adOpenStatic, adLockReadOnly
  'Retorna o resultado do recordset
  Set ExecutaSqlRs = objRs
  'Desconectar
  Set objCmd.ActiveConnection = Nothing
  Set objCmd = Nothing
  Set objRs.ActiveConnection = Nothing
  Exit Function
End Function

Public Sub ExecutaSql(strSql As String)
  
  'Cria o Objeto ADO
  
  Dim objCmd    As ADODB.Command
  'Inicializa os Objetos
  
  Set objCmd = New ADODB.Command
  'Informo os parâmetros do Objeto
  strStringConexao = "Provider=MSDASQL.1;Data Source=sda;"
  objCmd.ActiveConnection = strStringConexao
  objCmd.CommandType = adCmdText
  objCmd.CommandText = strSql
  objCmd.Execute
  'Desconectar
  Set objCmd.ActiveConnection = Nothing
  Set objCmd = Nothing
  Exit Sub
End Sub
Public Sub SelecionaItens(ByVal sig_orgao_processo As String, ByVal seq_processo As Integer, ByVal ano_processo As Integer, ByVal seq_sa As Integer, ByVal seq_sa_complementar As String, ByVal seq_assunto As Variant, ByVal seq_area As Variant)
Sql = " select sa.seq_assunto,seq_area, sa.seq_item_sa, " & _
        " iif (descr_item_modificado_sa <> null, descr_item_modificado_sa,descr_item_sa) " & _
        " as descr_item," & _
        " seq_ordem_impressao " & _
        "   from sa_item_auditoria sa inner join item_sa i" & _
        "       on sa.seq_item_sa = i.seq_item_sa" & _
        " and sa.seq_assunto = i.seq_assunto" & _
        " Where" & _
        " sig_orgao_processo ='" & sig_orgao_processo & _
        "' and ano_processo = " & ano_processo & _
        " and seq_processo = " & seq_processo & _
        " and seq_sa = " & seq_sa & _
        " and seq_sa_complementar = '" & seq_sa_complementar & "'"
        If seq_assunto <> "" Then
            Sql = Sql & " and sa.seq_assunto=" & seq_assunto
        End If
        If seq_area <> "" Then
            Sql = Sql & " and seq_area=" & seq_area
        End If
        Sql = Sql & " order by seq_ordem_impressao"
        
        frmSAConsultar.rdcItens.DataSourceName = "sda"
        frmSAConsultar.rdcItens.Sql = Sql
        frmSAConsultar.rdcItens.Refresh
End Sub
Sub CarregaAssunto(Rs As ADODB.Recordset, Form As Form)
i = 0

Do While Not Rs.EOF
    seq_assunto = Rs("seq_assunto")
    descr_assunto = Rs("descr_assunto")
    Form.SelAssunto.AddItem descr_assunto, i
    Form.SelAssunto.ItemData(i) = seq_assunto
    Rs.MoveNext
    i = i + 1
Loop
End Sub

Sub CarregaSA(ByRef Rs As ADODB.Recordset, Form As Form, sig_orgao_processo As String, seq_processo As Integer, ano_processo As Integer)

Do While Not Rs.EOF
    Form.SelSa.AddItem Rs("seq_sa") & IIf(Rs("seq_sa_complementar") = "0", "", "-" & Rs("seq_sa_complementar"))
    Rs.MoveNext
Loop

End Sub

Sub CarregaArea(ByRef Rs As ADODB.Recordset, Form As Form)

i = 0
Do While Not Rs.EOF
    Form.SelArea.AddItem Rs("descr_area"), i
    Form.SelArea.ItemData(i) = Rs("seq_area")
Rs.MoveNext
i = i + 1
Loop
End Sub

Sub CarregaProcesso(ByRef Rs As ADODB.Recordset, Form As Form)
    Form.txtNumProcesso.Text = "PA" & "-" & Rs("sig_orgao_processo") & "-" & Right("000" & Rs("seq_processo"), 3) & "/" & Rs("ano_processo")
End Sub



Sub CarregaChecklist(ByRef Rs As ADODB.Recordset, Form As Form)
    i = 0
    Do While Not Rs.EOF
        Form.SelChklst.AddItem Rs("descr_checklist"), i
        Form.SelChklst.ItemData(i) = Rs("seq_checklist")
    Rs.MoveNext
    i = i + 1
    Loop
End Sub

Sub CarregaPergunta(ByRef Rs As ADODB.Recordset, Form As Form)
    i = 0
    Do While Not Rs.EOF
        Form.SelItemSa.AddItem Rs("descr_item_sa"), i
        Form.SelItemSa.ItemData(i) = Rs("seq_item_sa")
    Rs.MoveNext
    i = i + 1
    Loop
End Sub

Sub CarregaTipoTemplate(ByRef Rs As ADODB.Recordset, Form As Form)
    i = 0
    Do While Not Rs.EOF
        Form.SelTipoTemplate.AddItem Rs("descr_tipo_documento"), i
        Form.SelTipoTemplate.ItemData(i) = Rs("seq_tipo_documento")
    Rs.MoveNext
    i = i + 1
    Loop
End Sub


Sub CarregaTopicoTemplate(ByRef Rs As ADODB.Recordset, Form As Form)
    i = 0
    Do While Not Rs.EOF
        Form.SelTopicoDoc.AddItem Rs("descr_conteudo"), i
        Form.SelTopicoDoc.ItemData(i) = Rs("seq_topico_doc")
    Rs.MoveNext
    i = i + 1
    Loop
End Sub

Sub CarregaTemplate(ByRef Rs As ADODB.Recordset, Form As Form)
    i = 0
    Do While Not Rs.EOF
        Form.SelTemplate.AddItem Rs("descr_template"), i
        Form.SelTemplate.ItemData(i) = Rs("seq_template")
    Rs.MoveNext
    i = i + 1
    Loop
End Sub
Sub CarregaCampoChave(ByRef Rs As ADODB.Recordset, Form As Form)
    i = 0
    Do While Not Rs.EOF
        Form.SelPalavraChave.AddItem Rs("campo_chave"), i
        Rs.MoveNext
        i = i + 1
    Loop
End Sub


