<%
'**********************************************************************
'CONVERTE UMA STRING PARA UM BUFFER DE BYTES
'PARAMETROS:
'Str -> STRING
'RETORNO:
'EQUIVALENTE BINARIO DA STRING
'DEPENDENCIAS:
'NENHUMA

'**********************************************************************
Function StrToBin(Str)
    Dim Bin
    Dim I
    
    For I = 1 to Len(Str)
    'CONVERTE CADA CARACTERE PARA BYTE
		Bin = Bin & ChrB(Asc(Mid(Str, I, 1)))
    Next
    StrToBin=Bin
End Function
%>	