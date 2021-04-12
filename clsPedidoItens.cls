VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPedidoItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private objProduto As New clsProduto

Private mlngCodigoPedido As Long
Private mlngCodigoProduto As Long
Private mdblQuantidade As Double
Private mdblPreco As Double
Private mdblValorTotal As Double
Public Sub class_Initialize()
   Set objProduto = New clsProduto
End Sub
Public Sub class_terminate()
   Set objProduto = Nothing
End Sub
Public Property Let CodigoPedido(ByVal lngCodigoPedido_ As Long)
   mlngCodigoPedido = lngCodigoPedido_
End Property
Public Property Get CodigoPedido() As Long
   CodigoPedido = mlngCodigoPedido
End Property
Public Property Let CodigoProduto(ByVal lngCodigoProduto_ As Long)
   mlngCodigoProduto = lngCodigoProduto_
End Property
Public Property Get CodigoProduto() As Long
   CodigoProduto = mlngCodigoProduto
End Property
Public Property Let Quantidade(ByVal dblQuantidade_ As Double)
   mdblQuantidade = dblQuantidade_
End Property
Public Property Get Quantidade() As Double
   Quantidade = mdblQuantidade
End Property
Public Property Let Preco(ByVal dblPreco_ As Double)
   mdblPreco = dblPreco_
End Property
Public Property Get Preco() As Double
   Preco = mdblPreco
End Property
Public Property Let ValorTotal(ByVal dblValorTotal_ As Double)
   mdblValorTotal = dblValorTotal_
End Property
Public Property Get ValorTotal() As Double
   ValorTotal = mdblValorTotal
End Property
Public Function rsPedidoItem(Optional strWhere_ As String = "") As ADODB.Recordset
   Dim strSQL As String
   Dim rs As New ADODB.Recordset

   On Error GoTo CATCH

   strSQL = "SELECT PI.CodigoPedido,PI.CodigoProduto,PI.Quantidade,PI.Preco,PI.ValorTotal, " & vbCr & _
          "       P.Nome, P.Preco AS PrecoOriginal" & vbCr & _
          "FROM PedidoItem AS PI INNER JOIN Produto AS P ON P.Codigo = PI.CodigoProduto" & vbCr
   If strWhere_ <> "" Then
      strSQL = strSQL & "WHERE " & strWhere_
   End If

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic
   Set rsPedidoItem = rs

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método rsPedidoItem.Módulo de classe.clsPedidoItem" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Private Sub BuscarItensDoPedido()
   Dim strSQL As String
   Dim rs As New ADODB.Recordset

   On Error GoTo CATCH
   'TODO
'   mlngCodigo = -1
'   strSQL = "SELECT Codigo" & vbCr & _
'            "FROM PedidoItem" & vbCr & _
'            "WHERE CodigoProduto = " & mlngCodigoProduto & ""
'
'   rs.CursorLocation = adUseClient
'   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic
'
'   If rs.RecordCount > 0 Then
'      mlngCodigo = Nulo(rs!Codigo)
'   End If
'
'   rs.Close
'   Set rs = Nothing

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método BuscarCodigoPorCodigoProduto.Módulo de classe.clsPedidoItem" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Public Function Gravar() As Boolean
   Dim strSQL As String
   On Error GoTo CATCH

   Gravar = False
   
   strSQL = "INSERT INTO PedidoItem " & vbCr & _
            "(CodigoPedido, CodigoProduto, Quantidade, Preco, ValorTotal)" & vbCr & _
            "VALUES (" & vbCr & _
            "" & mlngCodigoPedido & "" & vbCr & _
            "," & mlngCodigoProduto & "" & vbCr & _
            "," & FormatarNumeroSQL(Str(mdblQuantidade)) & "" & vbCr & _
            "," & FormatarNumeroSQL(Str(mdblPreco)) & "" & vbCr & _
            "," & FormatarNumeroSQL(Str(mdblValorTotal)) & "" & vbCr & _
            ")"
   Call ConexaoBanco.Execute(strSQL)
   
   Gravar = True
   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método Gravar.Módulo de classe.clsPedidoItem" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Public Function Excluir() As Boolean
   Dim strSQL As String
   On Error GoTo CATCH

   Excluir = False
   
   objProduto.Codigo = mlngCodigoProduto
   strSQL = "DELETE FROM PedidoItem " & vbCr & _
            "WHERE CodigoPedido = " & mlngCodigoPedido
   Call ConexaoBanco.Execute(strSQL)

   Excluir = True
   
   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método Excluir.Módulo de classe.clsPedidoItem" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
