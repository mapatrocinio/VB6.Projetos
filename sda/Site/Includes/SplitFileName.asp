<%
'**********************************************************************
'SEPARA UM NOME DE ARQUIVO EM SUAS PARTE OU
'SE O PARAMETRO FileName FOR UMA STRING NULA,
'JUNTA AS PARTES DO NOME
'PARAMETROS:
'FileName -> NOME DO ARQUIVO COMPLETO
'Drive <- DISCO
'Path <- CAMINHO
'Name <- NOME DO ARQUIVO
'Extension <- EXTENSAO
'RETORNO:
'NOME COMPLETO DO ARQUIVO
'DEPENDENCIAS:
'NENHUMA

'**********************************************************************
Function SplitFileName(FileName, Drive, Path, Name, Extension)
    Dim Pos
    Dim NomeCompleto
    
	If FileName="" And (Drive<>"" Or Path<>"" Or Name<>"" Or Extension<>"") Then
	'ESTA FORMANDO O NOME COMPLETO
		NomeCompleto=""
		If Drive<>"" Then
		'EXISTE DRIVE
			NomeCompleto=NomeCompleto & Drive
			If Len(Drive)=1 Then
			'NAO EXISTE : 
				NomeCompleto=NomeCOmpleto & ":"
			End If
		End If
		If Path<>"" Then
		'EXISTE CAMINHO
			NomeCompleto=NomeCompleto & Path
			If Right(Path,1)<>"\" Then
			'NAO EXISTE \
				NomeCompleto=NomeCompleto & "\"
			End If
		End If
		If Name<>"" Then
		'EXISTE NOME
			NomeCompleto=NomeCompleto & Name
		End If
		If Extension<>"" Then
		'EXISTE EXTENSAO
			If Left(Extension,1)<>"." And Right(NomeCompleto,1)<>"." Then
			'NAO EXISTE .
				NomeCompleto=NomeCOmpleto & "."
			End IF
			NomeCompleto=NomeCompleto & Extension
		End If
	ElseIf FileName<>"" Then
	'ESTA' SEPARANDO O NOME
		NomeCompleto=FileName
		If Len(NomeCompleto)>=2 Then
		'PODE TER UM DISCO
			If Mid(NomeCompleto,2,1)=":" Then
			'E' UM DISCO
				Drive=Left(NomeCompleto,2)
				NomeCompleto=Mid(NomeCompleto,3)
			End If
		End If
		If NomeCompleto<>"" Then
		'PODE TER UM CAMINHO
			Pos = InStrRev(NomeCompleto, "\")
			If Pos<>0 Then
			'EXISTE CAMINHO
				Path=Left(NomeCompleto,Pos)
				NomeCompleto=Mid(NomeCompleto,Pos+1)
			End If
		End If
		If NomeCompleto<>"" Then
		'PODE TER UM NOME
			Pos = InStrRev(NomeCompleto, ".")
			If Pos=0 Then
			'NAO EXISTE EXTENSAO
				Pos=Len(NomeCompleto)+1
			End If
			Name=Left(NomeCompleto,Pos-1)
			NomeCompleto=Mid(NomeCompleto,Pos+1)
		End If
		If NomeCompleto<>"" Then
		'EXISTE EXTENSAO
			Extension=NomeCompleto
		End If
		NomeCompleto=FileName
	End If
	SplitFileName=NomeCompleto
End Function
%>