VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim objPedidoItem As New clsPedidoItens

Private mlngCodigo As Long
Private mlngCodigoCliente As Long
Private mdblValorTotal As Double
Public Property Let CodigoCliente(ByVal lngCodigoCliente_ As Long)
   mlngCodigoCliente = lngCodigoCliente_
End Property
Public Property Get CodigoCliente() As Long
   CodigoCliente = mlngCodigoCliente
End Property
Public Property Let Codigo(ByVal lngCodigo_ As Long)
   mlngCodigo = lngCodigo_
End Property
Public Property Get Codigo() As Long
   Codigo = mlngCodigo
End Property
Public Property Let ValorTotal(ByVal dblValorTotal_ As Double)
   mdblValorTotal = dblValorTotal_
End Property
Public Property Get ValorTotal() As Double
   ValorTotal = mdblValorTotal
End Property

Public Function rsPedido(Optional strWhere_ As String = "") As ADODB.Recordset
   Dim strSQL As String
   Dim rs As New ADODB.Recordset

   On Error GoTo CATCH

   strSQL = "SELECT Codigo,CodigoCliente,ValorTotal," & vbCr & _
            "FROM Pedido " & vbCr
   If strWhere_ <> "" Then
      strSQL = strSQL & "WHERE " & strWhere_
   End If

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic
   Set rsPedido = rs

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método rsPedido.Módulo de classe.clsCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Public Function Gravar() As Boolean
   Dim strSQL As String
   On Error GoTo CATCH

   Gravar = False
   
   If mlngCodigo = -1 Then
      mlngCodigo = GerarNovoCodigo("Pedido", "Codigo")
      
      strSQL = "INSERT INTO Pedido " & vbCr & _
               "(Codigo, CodigoCliente, ValorTotal)" & vbCr & _
               "VALUES (" & vbCr & _
               "" & mlngCodigo & "" & vbCr & _
               "," & mlngCodigoCliente & "" & vbCr & _
               "," & FormatarNumeroSQL(Str(mdblValorTotal)) & vbCr & _
               ")"
   Else
      strSQL = "UPDATE Cliente SET " & vbCr & _
               "  ValorTotal = " & FormatarNumeroSQL(Str(mdblValorTotal)) & "" & vbCr & _
               "WHERE Codigo = " & mlngCodigo
   End If
   Call ConexaoBanco.Execute(strSQL)
   
   Gravar = True
   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método Gravar.Módulo de classe.clsCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Public Function Excluir() As Boolean
   Dim strSQL As String
   On Error GoTo CATCH

   Excluir = False
   
   ConexaoBanco.BeginTrans
   
   strSQL = "DELETE FROM Cliente " & vbCr & _
            "WHERE Codigo = " & mlngCodigo
   Call ConexaoBanco.Execute(strSQL)
   
   Excluir = objPedidoItem.Excluir
   
   ConexaoBanco.CommitTrans

   Exit Function
CATCH:
   ConexaoBanco.RollbackTrans
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método Excluir.Módulo de classe.clsCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function

Public Function VerificarPedidoPorCliente() As Boolean
   Dim strSQL As String
   Dim rs As New ADODB.Recordset

   On Error GoTo CATCH

   strSQL = "SELECT Count(Codigo) AS QuantidadePedidos" & vbCr & _
            "FROM Pedido " & vbCr & _
            "WHERE CodigoCliente = " & CodigoCliente

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic
   
   VerificarPedidoPorCliente = Nulo(rs!QuantidadePedidos, adEntradaNumero) > 0

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método rsPedido.Módulo de classe.clsCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function


