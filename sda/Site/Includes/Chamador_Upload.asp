
<!--#include file="FileUpload.asp"-->
<!--#include file="SplitFileName.asp"-->
<!--#include file="FileMakeDir.asp"-->
<!--#include file="FileExists.asp"-->
<!--#include file="StrToBin.asp"-->
<!--#include file="BinToStr.asp"-->


<%

' ====== Declaração das Constantes

' ====== Fim Declaração das Constantes

' ====== Declaração das variaveis

Dim Controle
Dim Buffer
Dim NomeCompleto
Dim Nome
Dim Extensao
Dim NomeDoArquivo
Dim Altura:Altura=0
Dim Largura:Largura=0
Dim I
Dim CodigoDoErro:CodigoDoErro=0
Dim Mensagem
Dim Tamanho
Dim Pos
' ====== Fim Declaração das variaveis

Sub BancoDadosIncluir
	
	Set Controle = GetControls()
	Buffer = Controle.Item("nomearquivo").Item("Value")
	NomeCompleto = Controle.Item("nomearquivo").Item("FileName")
	
	'Aux enquete =  Controle.Item("Aux").Item("Value")

	Call SplitFileName(NomeCompleto, "", "", Nome, Extensao)
	NomeDoArquivo = Nome & "." & Extensao

	if Extensao = "MDB" or Extensao = "mdb" then
		If CodigoDoErro=0 Then
			Call FileMakeDir(server.MapPath("bd\Importacao"))
			NomeCompleto = server.MapPath("bd\Importacao") & "\" & NomeDoArquivo
			'If Not FileExists(NomeCompleto) Then
				Call UploadFileBuffer(NomeCompleto,Buffer)
				Set Controle = Nothing
			'End If
		End If
	else
		Response.Write"<script>alert('Usar apenas arquivos do tipo mdb para importação');window.history.back(-1);</script>"  
		Response.End 
	end if  
End Sub

Sub PlanilhaIncluir(sig_orgao_processo,seq_processo,ano_processo,seq_sa,seq_sa_complementar,seq_assunto,seq_item_sa)
	dim fso
	Dim seq_imagem
	Set Controle = GetControls()
	Buffer = Controle.Item("nomearquivo").Item("Value")
	'Response.Write "buffer:" & Buffer
	
	NomeCompleto = Controle.Item("nomearquivo").Item("FileName")
	'Response.Write "NomeCompleto:" & NomeCompleto
	

	Call SplitFileName(NomeCompleto, "", "", Nome, Extensao)
	NomeDoArquivo = Nome & "." & Extensao
     
	if Extensao = "XLS" or Extensao = "xls" then
		If CodigoDoErro=0 Then
			'Response.Write server.MapPath("Planilhas")
			'Response.End
			Call FileMakeDir(server.MapPath("Planilhas"))
			NomeCompleto = server.MapPath("Planilhas") & "\" & NomeDoArquivo
			If Not FileExists(NomeCompleto) Then
				Set RS = Server.CreateObject("ADODB.Recordset")
				Call UploadFileBuffer(NomeCompleto,Buffer)
				Set fso = CreateObject("Scripting.FileSystemObject")
				if err.number <> 0 then TrataErro
				sql = "select isnull(max(seq_imagem),0) as seq_imagem from teste_imagem"
				Rs.open sql, cn,2,2
				if not rs.eof then
					seq_imagem = cint(Rs("seq_imagem"))
				end if
				Rs.close
				seq_imagem = seq_imagem + 1
				
				NomeArquivo = sig_orgao_processo & "-" & seq_processo & "-" & ano_processo & "-" & seq_sa & "-" & seq_sa_complementar & "-" & seq_assunto & "-" & seq_item_sa & "-" & cstr(seq_imagem)
				fso.CopyFile server.MapPath("Planilhas") & "\" & NomeDoArquivo ,server.MapPath("Planilhas") & "\" & NomeArquivo &  ".xls",true	
				if err.number <> 0 then TrataErro
								
				RS.Open "teste_imagem", Cn, 2, 2
				RS.AddNew
				RS("seq_imagem") = seq_imagem
				RS("img_excel").AppendChunk MultiByteToBinary(Controle.Item("nomearquivo").Item("Value"))
				RS.Update
				RS.Close
				Cn.Close
				
				
				fso.DeleteFile server.MapPath("Planilhas") & "\" & NomeDoArquivo
				Set Controle = Nothing
				set fso = nothing
			End If
		End If
	else
		Response.Write"<script>alert('Usar apenas arquivos do tipo xls para importação');window.history.back(-1);</script>"  
		Response.End 
	end if  
End Sub



Function MultiByteToBinary(MultiByte)
  ' This Function converts multibyte string to real binary data (VT_UI1 | VT_ARRAY)
  ' Using recordset
  Dim RS, LMultiByte, Binary
  Const adLongVarBinary = 205
  Set RS = CreateObject("ADODB.Recordset")
  LMultiByte = LenB(MultiByte)
  RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
  RS.Open
  RS.AddNew
    RS("mBinary").AppendChunk MultiByte & ChrB(0)
  RS.Update
  Binary = RS("mBinary").GetChunk(LMultiByte)
  MultiByteToBinary = Binary
End Function



Sub Mensagem1(pMensagem)
%>
	<table width="50%" border="0" cellspacing="1" cellpadding="4">
		<tr>
			<td class="a3-coluna">Aviso</td>
		</tr>
		<tr>
			<td bgcolor="#EEEEEE" align=center><b><%=pMensagem%></b></td>
		</tr>
	</table>
	<br>

<%
End Sub

%>


