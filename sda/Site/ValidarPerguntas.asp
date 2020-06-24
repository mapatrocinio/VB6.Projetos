<%@ Language=VBScript %>
<!-- #include file="Includes/ValidaSessao.asp" -->
<!-- #include file="Includes/Funcoes.asp" -->

<!--#include file="Includes/ConexaobancoDados.asp"-->
<!--#include file="Includes/TrataErro.asp"-->
<% 
Response.ContentType = "text/xml"
response.write "<?xml version=""1.0"" encoding=""Windows-1250""?>"

Dim xmlDoc 
Dim strListaItemSa
Dim strArea
Dim strAssunto
Dim strSql
Dim Sig_orgao_processo
Dim Ano_processo
Dim Seq_processo
Dim seq_sa
Dim seq_sa_complementar

strListaItemSa 			= request("Lista") 
strArea 						= request("cod_area")
strAssunto 					= request("cod_assunto")

Sig_orgao_processo 	= request("Sig_orgao_processo")
Ano_processo				= request("Ano_processo")
Seq_processo				= request("Seq_processo")
seq_sa							= request("seq_sa")
seq_sa_complementar	= request("seq_sa_complementar")


'Persist the Recorset in XML format to the ASP Response object. 
'The constant value for adPersistXML is 1.

Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = CreateObject("ADODB.Recordset")
'Persist the Recorset in XML format to the DOMDocument object.
'The constant value for adPersistXML is 1.
strSql = "Select DISTINCT descr_item_sa, area_auditoria.descr_area " & _
	"from item_sa inner join sa_item_auditoria on item_sa.seq_assunto = sa_item_auditoria.seq_assunto and " & _
		" item_sa.seq_item_sa = sa_item_auditoria.seq_item_sa " & _
		"inner join area_auditoria on area_auditoria.seq_area = sa_item_auditoria.seq_area " & _
		" where item_sa.seq_assunto = " & strAssunto & _
		" and item_sa.seq_item_sa in (" & strListaItemSa & ")" & _
		" and sa_item_auditoria.ano_processo = " & Ano_processo & _
		" and sa_item_auditoria.sig_orgao_processo = '" & Sig_orgao_processo & "'" & _
		" and sa_item_auditoria.seq_processo = " & Seq_processo & _
		" and sa_item_auditoria.seq_sa = " & seq_sa & _
		" and sa_item_auditoria.seq_sa_complementar = '" & seq_sa_complementar & "'" & _
		" and sa_item_auditoria.seq_area <> " & strArea

objRs.Open strSql,cn,0,1		

objRs.Save xmlDoc,1

objRs.Close
Set objRs = Nothing

Response.Write xmldoc.xml
%>