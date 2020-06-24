Attribute VB_Name = "Function"
Option Explicit

Global Const gsDiretor = "DIR"
Global Const gsGerente = "GER"
Global Const gsRecepcao = "REC"
Global Const gsPortaria = "POR"
Global Const gsAdmin = "ADM"
Global Const gsEstoque = "EST"

Enum TpTipoDados
  tpDados_Texto '0 a 255
  tpDados_Memo 'Sem Limite
  tpDados_Inteiro '-32767 a 32767
  tpDados_Longo 'Sem limite
  tpDados_DataHora 'MM/DD/YYYY hh:mm:ss
  tpDados_Moeda '121212.98
End Enum

Enum tpAceitaNulo
  tpNulo_Aceita
  tpNulo_NaoAceita
End Enum

Public Function Formata_Dados(pValor As Variant, pTipoDados As TpTipoDados, pAceitaNulo As tpAceitaNulo, Optional pTamanhoCampo As Integer) As Variant
  On Error GoTo trata
  '
  Dim vRetorno As Variant
  Dim sData As String
  '
  Select Case pTipoDados
  Case TpTipoDados.tpDados_Texto
    If Len(Trim(pValor & "")) = 0 Then
      If pAceitaNulo = tpNulo_Aceita Then
        vRetorno = "Null"
      Else
        vRetorno = "' '"
      End If
    Else
      vRetorno = "'" & Tira_Plic(Trim(pValor & "")) & "'"
    End If
  Case TpTipoDados.tpDados_Longo
    If Not IsNumeric(pValor) Then
      If pAceitaNulo = tpNulo_Aceita Then
        vRetorno = "Null"
      Else
        vRetorno = "0"
      End If
    Else
      vRetorno = CLng(pValor)
    End If
  Case TpTipoDados.tpDados_DataHora
    'Converter para Data
    sData = ""
    If Len(pValor & "") = 10 Then
      If Mid(pValor & "", 1, 2) <> "__" Then
        'Data no Formato DD/MM/YYYY
        sData = "#" & Mid(pValor, 4, 2) & "/" & Mid(pValor, 1, 2) & "/" & Mid(pValor, 7, 4) & "#"
      Else
        sData = "null"
      End If
    ElseIf Len(pValor & "") = 16 Then
      'Data no Formato DD/MM/YYYY hh:mm
      If Mid(pValor & "", 1, 2) <> "__" Then
        'Data no Formato DD/MM/YYYY
        sData = "#" & Mid(pValor, 4, 2) & "/" & Mid(pValor, 1, 2) & "/" & Mid(pValor, 7, 4) & " " & Mid(pValor, 12, 2) & ":" & Mid(pValor, 15, 2) & "#"
      Else
        sData = "null"
      End If
    Else
      sData = ""
    End If
    If Len(sData) = 0 Then
      If pAceitaNulo = tpNulo_Aceita Then
        vRetorno = "Null"
      Else
        vRetorno = "#01/01/1900#"
      End If
    Else
      vRetorno = sData
    End If
  Case TpTipoDados.tpDados_Moeda
    If Not IsNumeric(pValor) Then
      If pAceitaNulo = tpNulo_Aceita Then
        vRetorno = "Null"
      Else
        vRetorno = "0"
      End If
    Else
      vRetorno = Replace(pValor, ".", "")
      vRetorno = Replace(vRetorno, ",", ".")
    End If
  Case Else
  End Select
  '
  Formata_Dados = vRetorno
  '
  Exit Function
trata:
  
End Function

Public Function Tira_Plic(pValor As String) As Variant
  Tira_Plic = Replace(pValor, "'", "''")
End Function

