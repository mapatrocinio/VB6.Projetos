<%
'**********************************************************************
'ROTINA PARA UPLOAD DE ARQUIVOS
'PARAMETROS:
'FileName -> NOME DO ARQUIVO
'Buffer -> BUFFER COM O ARQUIVO A SER GRAVADO
'RETORNO:
'NENHUM
'DEPENDENCIAS:
'StrToBin
'BinToStr
'**********************************************************************
'PARA FAZER O UPLOAD DE ARQUIVOS E' NECESSARIO QUE:
'O FORMULARIO SEJA DEFINIDO COMO enctype="multipart/form-data"
'LER OS CONTROLES ATRAVES DA COLECAO CRIADA PELA FUNCAO GetControls
'EXEMPLO:
'NA PAGINA DO FORMULARIO TERIAMOS
'<FORM action="Upload.asp" method=POST id=form1 name=form1 enctype="multipart/form-data">
'<INPUT type="file" id="file1" name="file1">
'<INPUT type="submit" value="Submit" id=submit1 name=submit1></FORM>
'
'NO FORMULARIO DE ACTION (Upload.asp) SERIA
'Set Controle = GetControls()
'Buffer = Controle.Item("file1").Item("Value")
'NomeCompleto = Controle.Item("file1").Item("FileName")
'Call SplitFileName(NomeCompleto, "", "", Nome, Extensao)
'NomeCompleto = Nome & "." & Extensao
'Call UploadFileBuffer(NomeCompleto,Buffer)
'Set Controle = Nothing

Function GetControls()
'RETORNA UMA COLECAO DOS CAMPOS ENVIADOS VIA POST
'A COLECAO RETORNARA O NOME DO CAMPO COMO Name E O CONTEUDO COMO Value (NO CASO DE CAMPO TIPO FILE E' O CONTEUDO DO ARQUIVO)
'NO CASO DE CAMPOS TIPO FILE TAMBEM RETORNA FileName (NOME COMPLETO DO ARQUIVO) e ContentType (TIPO DE CONTEUDO)
'PARAMETROS:
'NENHUM
'RETORNO:
'COLECAO DE CONTROLES

	Dim Name					'NOME DO CONTROLE
	Dim Value					'VALOR DO CONTROLE
	Dim FileName				'NOME DO ARQUIVO
	Dim ContentType				'TIPO DE ARQUIVO
		
	Dim Dicionario				'OBJETO DICIONARIO DE DADOS
	Dim Inicio					'INICIO DA SUBSTRING
	Dim Fim						'FIM DA SUBSTRING
	Dim Limite					'STRING DE LIMITE DE CADA CONTROLE
	Dim PosicaoLimite			'POSICAO LIMITE
	Dim Controle				'OBJETO DO CONTROLE
	Dim Buffer					'BUFFER COM OS CAMPOS DO FORMULARIO
	
    Buffer = Request.BinaryRead(Request.TotalBytes)
	Set Dicionario = Server.CreateObject("Scripting.Dictionary")
    Inicio = 1
    Fim = InStrB(Inicio, Buffer, ChrB(13))
    Limite = MidB(Buffer, Inicio, Fim-Inicio)
    PosicaoLimite=1
    Do Until (PosicaoLimite=InstrB(Buffer, Limite & StrToBin("--")))
	'MONTA UMA ESTRUTURA PARA CADA CONTROLE PASSADO PELO METODO POST
		Set Controle = CreateObject("Scripting.Dictionary")
		Inicio = InstrB(PosicaoLimite, Buffer, StrToBin("Content-Disposition"))
		'INICIO DA RECUPERACAO DO NOME DO CONTROLE O CONTROLE VEM NO FORMATO name="XXXXXX"
		Inicio = InstrB(Inicio, Buffer, StrToBin("name=")) + 6
		Fim = InstrB(Inicio, Buffer, chrB(34))
		Name = BinToStr(MidB(Buffer, Inicio, Fim-Inicio))
		'FIM DA RECUPERACAO DO NOME DO CONTROLE
		'VERIFICA SE O CONTROLE E' DO TIPO FILE
		Inicio = Fim + 3
		If BinToStr(MidB(Buffer,Inicio,9))="filename=" Then
		'E' UM CONTROLE DO TIPO FILE
			Inicio = Inicio + 10
			Fim = InStrB(Inicio, Buffer, ChrB(34))
			'RECUPERA O NOME DO ARQUIVO
			FileName = BinToStr(MidB(Buffer, Inicio, Fim-Inicio))
			Controle.Add "FileName", FileName
			'RECUPERA O TIPO DE CONTEUDO
			Inicio = InstrB(Fim, Buffer, StrToBin("Content-Type:")) + 14
			Fim = InstrB(Inicio, Buffer, ChrB(13))
			ContentType = Trim(BinToStr(MidB(Buffer, Inicio, Fim-Inicio)))
			Controle.Add "ContentType", ContentType
			'RECUPERA O CONTEUDO DO ARQUIVO
			Inicio = Fim + 4
			
			Fim = InStrB(Inicio, Buffer, Limite)-2
			Value = MidB(Buffer, Inicio, Fim-Inicio)
		Else
		'E' UM CONTROLE QUALQUER
			Inicio = InstrB(Fim, Buffer, ChrB(13)) + 4
			Fim = InstrB(Inicio, Buffer, Limite)-2
			'RECUPERA OS DADOS DO CONTROLE COMUM
			Value = BinToStr(MidB(Buffer, Inicio, Fim-Inicio))
		End If
		Controle.Add "Value" , Value
		Dicionario.Add Name, Controle
		PosicaoLimite=InStrB(PosicaoLimite+LenB(Limite), Buffer, Limite)
			' leo
			'Response.Write name & " - "
			'Response.Write value & "<BR>"
			' fim leo
		Set Controle = Nothing
    Loop
	Set GetControls = Dicionario
	Set Dicionario = Nothing
End Function

Sub UploadFileBuffer(FileName,Buffer)
'SALVA O BUFFER EM UM ARQUIVO NO SERVIDOR
'PARAMETROS:
'FileName -> NOME DO ARQUIVO
'Buffer -> BUFFER COM O ARQUIVO
'RETORNO:
'NENHUM
	Dim NomeCompleto
	Dim ArquivoObjeto
	Dim Arquivo
	Dim I
	Dim Tam
	Dim P
	Dim Bloco
	Dim Blocos
	Dim TamanhoDoBloco

    ''If Left(Filename,2)="//" Or Left(Filename,2)="\\" Or Mid(Filename,2,1)=":" Then
    ''CAMINHO ABSOLUTO
    '	NomeCompleto=FileName
    'Else
    'CAMINHO RELATIVO AO SITE
    '	NomeCompleto = Server.MapPath("/" & FileName)
    'End If

	NomeCompleto=FileName

    for mCount = 1 to 1000
   
		Set ArquivoObjeto = Server.CreateObject("Scripting.FileSystemObject")
		Set Arquivo = ArquivoObjeto.CreateTextFile(NomeCompleto, True)
        
        if err.number = 0 then
           exit for
        end if
   
    next
     
        if err.number <> 0 then 
           Response.Write "<script>alert('O servidor virou carroça, tente executar a operação novamente');window.history.back(-1);</script>"      
           Response.End 
        end if
    
    Tam=LenB(Buffer)
    TamanhoDoBloco=128
    Blocos=Tam \ TamanhoDoBloco
    P=1
    For I = 1 to Blocos
    'SALVA CADA BYTE DO ARQUIVO
		Bloco=""
		For P=P to I*TamanhoDoBloco
			Bloco=Bloco & Chr(AscB(MidB(Buffer, P, 1)))
		Next
		Arquivo.Write Bloco
    Next
	Bloco=""
	For P = P to Tam
		Bloco=Bloco & Chr(AscB(MidB(Buffer, P, 1)))
	Next
	Arquivo.Write Bloco
    Arquivo.Close
    Set Arquivo = Nothing
    Set ArquivoObjeto = Nothing
End Sub

%>