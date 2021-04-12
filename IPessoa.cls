VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Property Let Codigo(ByVal lngCodigo_ As Long)
End Property
Public Property Get Codigo() As Long
End Property
Public Property Let Nome(ByVal strNome_ As String)
End Property
Public Property Get Nome() As String
End Property
Public Property Let Telefone(ByVal strTelefone_ As String)
End Property
Public Property Get Telefone() As String
End Property
Public Property Let Celular(ByVal strCelular_ As String)
End Property
Public Property Get Celular() As String
End Property
Public Function Gravar() As Boolean
End Function
Public Function Excluir() As Boolean
End Function


