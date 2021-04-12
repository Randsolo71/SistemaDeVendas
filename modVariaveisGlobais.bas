Attribute VB_Name = "modVariaveisGlobais"
Option Explicit

Public ConexaoBanco As New ADODB.Connection
Public pstrUsuarioLogado As String
Public Const COR_SELECAO = &HF0FFFF

Enum enuTipoEntradaDados
   adEntradaTexto = 0
   adEntradaNumero = 1
   adEntradaData = 2
   adEntradaBoolean = 3
   adEntradaDataVazia = 4
   adRetornaNumeroMenosUm = 5
   adEntradaDataVaziaTexto = 6
End Enum


