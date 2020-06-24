VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function ReadBinFile(ByVal bfilename As String) As Variant
  Dim fl As Long
  Dim FileNum As Long
  Dim binbyte() As Byte
  Dim binfilestr As String
  Dim objCon As ADODB.Connection
  
  On Error GoTo trata
  
  FileNum = FreeFile
  Open bfilename For Binary Access Read As #FileNum
  
  fl = FileLen(bfilename)
  ReDim binbyte(fl)
  
  Get #FileNum, , binbyte
  
  Close #FileNum
  
  ReadBinFile = binbyte
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source, _
            Err.Description
End Function

Function RetornaStrCon(ByVal intBd As Integer, ByVal strPath As Variant) As String
Dim strConexao As String
' Quando intBD for:
'  - 1 = >  conexão com o Auditoria no SQL SERVER
'  - 2 = >  conexão com o Auditoria no Access
'  - 3 = >  conexão com o CtrAcesso no SqlServer

'''Priodução
'''If intBd = 1 Then
'''    strConexao = "Provider=SQLOLEDB.1;Password=audin_usr;Persist Security Info=True;User ID=usr_audin;Initial Catalog=auditoria;Data Source=rweb02s"
'''ElseIf intBd = 2 Then
'''    strConexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & "\Auditoria.mdb" & ";Persist Security Info=False"
'''ElseIf intBd = 3 Then
'''    strConexao = "Provider=SQLOLEDB.1;Password=audin_usr;Persist Security Info=True;User ID=usr_audin;Initial Catalog=CtrAcesso;Data Source=XINM01S"
'''End If

If intBd = 1 Then
  'AUDITORIA SITE INTRANET
  'Desenvolvimento (máquina analista)
  strConexao = "Provider=SQLOLEDB.1;Password=midesenv;Persist Security Info=True;User ID=mimontreal;Initial Catalog=auditoria_teste;Data Source=rdes01s"
  'Desenvolvimento (máquina RDES01S)
  'strConexao = "Provider=SQLOLEDB.1;Password=midesenv;Persist Security Info=True;User ID=mimontreal;Initial Catalog=auditoria;Data Source=rdes01s"
  'Priodução
  'strConexao = "Provider=SQLOLEDB.1;Password=audin_usr;Persist Security Info=True;User ID=usr_audin;Initial Catalog=auditoria;Data Source=rweb02s"
ElseIf intBd = 2 Then
  'MÓDULO NOTEBOOK
  'strConexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & "\Auditoria.mdb" & ";Persist Security Info=False"
  strConexao = "Provider=SQLOLEDB.1;Password=midesenv;Persist Security Info=True;User ID=mimontreal;Initial Catalog=auditoria;Data Source=sinfo-42rc"
   'strConexao = "sdanote,null,null"
'ElseIf intBd = 3 Then
    'strConexao = "Provider=SQLOLEDB.1;Password=midesenv;Persist Security Info=True;User ID=mimontreal;Initial Catalog=CtrAcesso;Data Source=xdes01s"
End If

RetornaStrCon = strConexao


End Function
