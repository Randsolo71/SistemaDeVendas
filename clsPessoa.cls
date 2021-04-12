VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Implements IPessoa

Private mlngCodigo As Long
Private mstrNome As String
Private mstrTelefone As String
Private mstrCelular As String
Private mstrTipo As String
Public Property Let Codigo(ByVal lngCodigo_ As Long)
    mlngCodigo = lngCodigo_
End Property
Public Property Get Codigo() As Long
    Codigo = mlngCodigo
End Property
Public Property Let Nome(ByVal strNome_ As String)
    mstrNome = strNome_
End Property
Public Property Get Nome() As String
    Nome = mstrNome
End Property
Public Property Let Telefone(ByVal strTelefone_ As String)
    mstrTelefone = strTelefone_
End Property
Public Property Get Telefone() As String
    Telefone = mstrTelefone
End Property
Public Property Let Celular(ByVal strCelular_ As String)
    mstrCelular = strCelular_
End Property
Public Property Get Celular() As String
    Celular = mstrCelular
End Property
Public Property Let Tipo(ByVal strTipo_ As String)
    mstrTipo = strTipo_
End Property
Public Property Get Tipo() As String
    Tipo = mstrTipo
End Property
Public Function VerificarSePessoaJaExiste() As Long
   Dim strSQL As String
   Dim rs As New ADODB.Recordset
   On Error GoTo CATCH

   VerificarSePessoaJaExiste = -1
   strSQL = "SELECT Nome" & vbCr & _
            "FROM Pessoa" & vbCr & _
            "WHERE Tipo = '" & mstrTipo & "'" & vbCr & _
            "  AND Nome = '" & mstrNome & "'"

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic
   
   If rs.RecordCount > 0 Then
      mlngCodigo = Nulo(rs!Codigo)
      VerificarSePessoaJaExiste = mlngCodigo
   End If
   
   rs.Close
   Set rs = Nothing

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método VerificarSePessoaJaExiste.Módulo de classe.clsPessoa" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Public Function Gravar() As Long
   Dim strSQL As String
   
   Gravar = -1
   
   If mlngCodigo = -1 Then
      mlngCodigo = GerarNovoCodigo("Pessoa", "Codigo")
      strSQL = "INSERT INTO Pessoa " & vbCr & _
               "(Codigo, Nome, Telefone, Celular, Tipo)" & vbCr & _
               "VALUES (" & vbCr & _
               "" & mlngCodigo & "" & vbCr & _
               ",'" & mstrNome & "'" & vbCr & _
               ",'" & mstrTelefone & "'" & vbCr & _
               ",'" & mstrCelular & "'" & vbCr & _
               ",'" & mstrTipo & "'" & vbCr & _
               ")"
   Else
      strSQL = "UPDATE Pessoa SET " & vbCr & _
               " Nome = '" & mstrNome & "'" & vbCr & _
               ",Telefone = '" & mstrTelefone & "'" & vbCr & _
               ",Celular = '" & mstrCelular & "'" & vbCr & _
               ",Tipo = '" & mstrTipo & "'" & vbCr & _
               "WHERE Codigo = " & mlngCodigo
   End If
   Call ConexaoBanco.Execute(strSQL)
   Gravar = mlngCodigo

End Function
Public Function Excluir() As Boolean
   Dim strSQL As String
   On Error GoTo CATCH

   Excluir = False
   
   strSQL = "DELETE FROM Pessoa " & vbCr & _
            "WHERE Codigo = " & mlngCodigo
   Call ConexaoBanco.Execute(strSQL)
   
   Excluir = True

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método Excluir.Módulo de classe.clsPessoa" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function

