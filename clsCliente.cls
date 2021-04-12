VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private objPessoa As New clsPessoa

Private mlngCodigo As Long
Private mdblLimiteCredito As Double
Private mlngCodigoPessoa As Long
Public Sub class_Initialize()
   Set objPessoa = New clsPessoa
End Sub
Public Sub class_terminate()
   Set objPessoa = Nothing
End Sub
Public Property Let CodigoPessoa(ByVal lngCodigoPessoa_ As Long)
   mlngCodigoPessoa = lngCodigoPessoa_
End Property
Public Property Get CodigoPessoa() As Long
   CodigoPessoa = mlngCodigoPessoa
End Property
Public Property Let Nome(ByVal strNome_ As String)
   objPessoa.Nome = strNome_
End Property
Public Property Get Nome() As String
   Nome = objPessoa.Nome
End Property
Public Property Let Telefone(ByVal strTelefone_ As String)
   objPessoa.Telefone = strTelefone_
End Property
Public Property Get Telefone() As String
   Telefone = objPessoa.Telefone
End Property
Public Property Let Celular(ByVal strCelular_ As String)
   objPessoa.Celular = strCelular_
End Property
Public Property Let Tipo(ByVal strTipo_ As String)
    objPessoa.Tipo = strTipo_
End Property
Public Property Get Tipo() As String
    Tipo = objPessoa.Tipo
End Property
Public Property Get Celular() As String
   Celular = objPessoa.Celular
End Property
Public Property Let Codigo(ByVal lngCodigo_ As Long)
   mlngCodigo = lngCodigo_
End Property
Public Property Get Codigo() As Long
   Codigo = mlngCodigo
End Property
Public Property Let LimiteCredito(ByVal dblLimiteCredito_ As Double)
   mdblLimiteCredito = dblLimiteCredito_
End Property
Public Property Get LimiteCredito() As Double
   LimiteCredito = mdblLimiteCredito
End Property

Public Function rsCliente(Optional strWhere_ As String = "") As ADODB.Recordset
   Dim strSQL As String
   Dim rs As New ADODB.Recordset

   On Error GoTo CATCH

   strSQL = "SELECT C.Codigo,C.CodigoPessoa,C.LimiteCredito," & vbCr & _
          "       P.Nome, P.Telefone, P.Celular" & vbCr & _
          "FROM Cliente as C INNER JOIN Pessoa as P ON P.Codigo = C.CodigoPessoa" & vbCr & _
          "WHERE P.Tipo = 'C'" & vbCr
   If strWhere_ <> "" Then
      strSQL = strSQL & "  AND " & strWhere_
   End If

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic
   Set rsCliente = rs

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método rsCliente.Módulo de classe.clsCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Private Sub BuscarCodigoPorCodigoPessoa()
   Dim strSQL As String
   Dim rs As New ADODB.Recordset

   On Error GoTo CATCH

   mlngCodigo = -1
   strSQL = "SELECT Codigo" & vbCr & _
            "FROM Cliente" & vbCr & _
            "WHERE CodigoPessoa = " & mlngCodigoPessoa & ""

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic
   
   If rs.RecordCount > 0 Then
      mlngCodigo = Nulo(rs!Codigo)
   End If
   
   rs.Close
   Set rs = Nothing

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método BuscarCodigoPorCodigoPessoa.Módulo de classe.clsCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Public Function VerificarSeJaExiste() As Long
   On Error GoTo CATCH

   mlngCodigoPessoa = objPessoa.VerificarSePessoaJaExiste
   If mlngCodigoPessoa > 0 Then
      Call BuscarCodigoPorCodigoPessoa
   Else
      mlngCodigo = -1
   End If
   VerificarSeJaExiste = mlngCodigo

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método VerificarSeJaExiste.Módulo de classe.clsCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Public Function Gravar() As Boolean
   Dim strSQL As String
   On Error GoTo CATCH

   Gravar = False
   ConexaoBanco.BeginTrans
   
   If mlngCodigo = -1 Then
      mlngCodigo = GerarNovoCodigo("Cliente", "Codigo")
      
      objPessoa.Codigo = -1
      mlngCodigoPessoa = objPessoa.Gravar
      
      strSQL = "INSERT INTO Cliente " & vbCr & _
               "(Codigo, CodigoPessoa, LimiteCredito)" & vbCr & _
               "VALUES (" & vbCr & _
               "" & mlngCodigo & "" & vbCr & _
               "," & mlngCodigoPessoa & "" & vbCr & _
               "," & FormatarNumeroSQL(Str(mdblLimiteCredito)) & "" & vbCr & _
               ")"
   Else
   
      mlngCodigoPessoa = objPessoa.Gravar
   
      strSQL = "UPDATE Cliente SET " & vbCr & _
               "  LimiteCredito = " & FormatarNumeroSQL(Str(mdblLimiteCredito)) & "" & vbCr & _
               "WHERE Codigo = " & mlngCodigo
   End If
   Call ConexaoBanco.Execute(strSQL)
   ConexaoBanco.CommitTrans
   
   Gravar = True
   Exit Function
CATCH:
   ConexaoBanco.RollbackTrans
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método Gravar.Módulo de classe.clsCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Public Function Excluir() As Boolean
   Dim strSQL As String
   Dim objPedido As New clsPedido
   On Error GoTo CATCH

   Excluir = False
   
   objPedido.CodigoCliente = mlngCodigo
   If Not objPedido.VerificarPedidoPorCliente Then
      ConexaoBanco.BeginTrans
      
      objPessoa.Codigo = mlngCodigoPessoa
      strSQL = "DELETE FROM Cliente " & vbCr & _
               "WHERE Codigo = " & mlngCodigo
      Call ConexaoBanco.Execute(strSQL)
      Excluir = objPessoa.Excluir
      
      ConexaoBanco.CommitTrans
   Else
      Call MsgBox("Não é possivel excluir pois existem Pedidos para o cliente !", vbExclamation, "Atenção")
   End If

   Exit Function
CATCH:
   ConexaoBanco.RollbackTrans
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método Excluir.Módulo de classe.clsCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Public Sub BuscarCliente()
   Dim strSQL As String
   Dim rs As New ADODB.Recordset

   On Error GoTo CATCH

   mlngCodigoPessoa = -1
   Nome = ""
   Telefone = ""
   Celular = ""
   Tipo = ""
   mdblLimiteCredito = 0
   
   strSQL = "SELECT C.Codigo,C.CodigoPessoa,C.LimiteCredito," & vbCr & _
          "       P.Nome, P.Telefone, P.Celular, P.Tipo " & vbCr & _
          "FROM Cliente as C INNER JOIN Pessoa as P ON P.Codigo = C.CodigoPessoa" & vbCr & _
          "WHERE C.Codigo = " & mlngCodigo

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic
   If rs.RecordCount > 0 Then
      mlngCodigoPessoa = Nulo(rs!CodigoPessoa, adEntradaNumero)
      Nome = Nulo(rs!Nome, adEntradaTexto)
      mdblLimiteCredito = Nulo(rs!LimiteCredito, adEntradaNumero)
      Telefone = Nulo(rs!Telefone, adEntradaTexto)
      Celular = Nulo(rs!Celular, adEntradaTexto)
      Tipo = Nulo(rs!Tipo, adEntradaTexto)
   End If
   
   rs.Close
   Set rs = Nothing

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método BuscarCliente.Módulo de classe.clsCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Public Sub AtualizarLimiteCredito()
Dim strSQL  As String

   strSQL = "UPDATE Cliente SET " & vbCr & _
            "  LimiteCredito = " & FormatarNumeroSQL(Str(mdblLimiteCredito)) & "" & vbCr & _
            "WHERE Codigo = " & mlngCodigo
   Call ConexaoBanco.Execute(strSQL)

End Sub
