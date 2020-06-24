<%Response.Expires = 0%>
<!--#include file="Includes/ConexaobancoDados.asp"-->

<% 
dim cod_uo
dim ano_processo
dim Data


cod_uo = Request("cod_uo")
Data= Request("Data")
if IsDate(Data) then
	ano_processo = year(Data)
else
	ano_processo = ""
end if
Response.ContentType = "text/xml"
response.write "<?xml version=""1.0"" encoding=""Windows-1250""?>"


Dim rs 
Dim xmlDoc 


Set rs=Server.CreateObject("ADODB.Recordset") 

'Replace the ADO Connection string attributes
'in the following line of code to point to your
'instance of SQL Server, and to specify the 
'required security credentials for User ID and Password.

rs.CursorLocation = 3
if ano_processo <> "" and cod_uo <> "" then
	rs.Open "Select 'PA' + '-' + sig_orgao_processo + '-' + right('000' + convert(varchar(3),seq_processo),3) + '/' + convert(varchar(4),ano_processo) as ProcessoAudin" & _                
			" from processo_auditoria where ano_processo = " & ano_processo & _
			" and cod_uo='" & cod_uo & "' and sig_tipo_auditoria = 'O'" ,cn
else
	rs.Open "Select seq_processo as ProcessoAudin from processo_auditoria where ano_processo =000 ",cn
end if

'Persist the Recorset in XML format to the ASP Response object. 
'The constant value for adPersistXML is 1.

Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")

'Persist the Recorset in XML format to the DOMDocument object.
'The constant value for adPersistXML is 1.

rs.Save xmlDoc,1

rs.Close
cn.Close

Set rs = Nothing
Set cn = Nothing 

'Write out the xml property of the DOMDocument
'object to the client Browser
Response.Write xmldoc.xml
%>