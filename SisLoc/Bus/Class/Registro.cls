VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub InicializaRegister(ByRef strTituloSistema As String, _
                              ByRef strReportPath As String, _
                              ByRef strAppPath As String, _
                              ByRef strNomeUsu As String, _
                              ByRef strNomeEmpresa As String, _
                              ByRef strBMPPath As String, _
                              ByRef strIconsPath As String, _
                              ByRef strBMP As String, _
                              ByRef strPathBackup As String, _
                              ByRef strNomeServidorBD As String, _
                              ByRef strConnectRpt As String)
  On Error GoTo trata
  'Inicializar variáveis do register
  CapturaParametrosRegister strTituloSistema, _
                            strReportPath, _
                            strAppPath, _
                            strNomeUsu, _
                            strNomeEmpresa, _
                            strBMPPath, _
                            strIconsPath, _
                            strBMP, _
                            strPathBackup, _
                            strNomeServidorBD

  TratarDimensaoBotoes strTituloSistema
  
  strConnectRpt = "ODBC;DSN=SisLoc;UID=SA;PWD=SHOGUM2806;DATABASE=SisLoc"
  
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub CapturaParametrosRegister(ByRef strTituloSistema As String, _
                                     ByRef strReportPath As String, _
                                     ByRef strAppPath As String, _
                                     ByRef strNomeUsu As String, _
                                     ByRef strNomeEmpresa As String, _
                                     ByRef strBMPPath As String, _
                                     ByRef strIconsPath As String, _
                                     ByRef strBMP As String, _
                                     ByRef strPathBackup As String, _
                                     ByRef strNomeServidorBD As String)
  On Error GoTo trata
  'Capturar parâmetros do register
  Dim objRegistro As datSisLoc.clsRegistro
  Set objRegistro = New datSisLoc.clsRegistro
  objRegistro.CapturaParametrosRegister strTituloSistema, _
                                        strReportPath, _
                                        strAppPath, _
                                        strNomeUsu, _
                                        strNomeEmpresa, _
                                        strBMPPath, _
                                        strIconsPath, _
                                        strBMP, _
                                        strPathBackup, _
                                        strNomeServidorBD
  Set objRegistro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub SalvaParametrosRegisterUsuario(strTituloSistema As String, _
                                          strNomeUsu As String, _
                                          strNivel As String)
  On Error GoTo trata
  Dim objRegistro As datSisLoc.clsRegistro
  Set objRegistro = New datSisLoc.clsRegistro
  objRegistro.SalvaParametrosRegisterUsuario strTituloSistema, _
                                             strNomeUsu, _
                                             strNivel
  Set objRegistro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub SalvaParametrosRegisterBMP(strTituloSistema As String, _
                                      strBMP As String)
  On Error GoTo trata
  Dim objRegistro As datSisLoc.clsRegistro
  Set objRegistro = New datSisLoc.clsRegistro
  objRegistro.SalvaParametrosRegisterBMP strTituloSistema, _
                                         strBMP
  Set objRegistro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub SalvaParametrosRegisterNivel(strTituloSistema As String)
  On Error GoTo trata
  Dim objRegistro As datSisLoc.clsRegistro
  Set objRegistro = New datSisLoc.clsRegistro
  objRegistro.SalvaParametrosRegisterNivel strTituloSistema
  Set objRegistro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub TratarDimensaoBotoes(strTitulosSistema As String)
  On Error GoTo trata
  Dim objRegistro As datSisLoc.clsRegistro
  Set objRegistro = New datSisLoc.clsRegistro
  objRegistro.TratarDimensaoBotoes strTitulosSistema
  Set objRegistro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub SalvarChaveRegistro(strTitulosSistema As String, _
                               strKey As String, _
                               strValor As String)
  On Error GoTo trata
  Dim objRegistro As datSisLoc.clsRegistro
  Set objRegistro = New datSisLoc.clsRegistro
  objRegistro.SalvarChaveRegistro strTitulosSistema, _
                                  strKey, _
                                  strValor
  Set objRegistro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function RetornarChaveRegistro(strTitulosSistema As String, _
                                      strKey As String)
  On Error GoTo trata
  Dim objRegistro As datSisLoc.clsRegistro
  Set objRegistro = New datSisLoc.clsRegistro
  RetornarChaveRegistro = objRegistro.RetornarChaveRegistro(strTitulosSistema, _
                                                            strKey)
  Set objRegistro = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

