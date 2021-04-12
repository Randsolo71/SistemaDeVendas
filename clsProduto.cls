VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngCodigo As Long
Private mstrNome As String
Private mdblPreco As Double
Private mstrCodigoExterno As String
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
Public Property Let Preco(ByVal dblPreco_ As Double)
   mdblPreco = dblPreco_
End Property
Public Property Get Preco() As Double
   Preco = mdblPreco
End Property
Public Property Let CodigoExterno(ByVal strCodigoExterno_ As String)
   mstrCodigoExterno = strCodigoExterno_
End Property
Public Property Get CodigoExterno() As String
   CodigoExterno = mstrCodigoExterno
End Property

Public Function rsProduto(Optional strWhere_ As String = "") As ADODB.Recordset
   Dim strSQL As String
   Dim rs As New ADODB.Recordset

   On Error GoTo CATCH

   strSQL = "SELECT Codigo,Nome,Preco, CodigoExterno" & vbCr & _
          "FROM Produto " & vbCr
   If strWhere_ <> "" Then
      strSQL = strSQL & "WHERE " & strWhere_
   End If

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic
   Set rsProduto = rs

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método rsProduto.Módulo de classe.clsProduto" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Private Sub BuscarCodigo()
   Dim strSQL As String
   Dim rs As New ADODB.Recordset

   On Error GoTo CATCH

   mlngCodigo = -1
   strSQL = "SELECT Codigo" & vbCr & _
            "FROM Produto" & vbCr & _
            "WHERE Codigo = " & mlngCodigo & ""

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic
   
   If rs.RecordCount > 0 Then
      mlngCodigo = Nulo(rs!Codigo, adRetornaNumeroMenosUm)
   End If
   
   rs.Close
   Set rs = Nothing

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método BuscarCodigo.Módulo de classe.clsProduto" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Private Sub BuscarCodigoPorCodigoExterno()
   Dim strSQL As String
   Dim rs As New ADODB.Recordset

   On Error GoTo CATCH

   mlngCodigo = -1
   strSQL = "SELECT Codigo" & vbCr & _
            "FROM Produto" & vbCr & _
            "WHERE CodigoExterno = '" & mstrCodigoExterno & "'"

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic
   
   If rs.RecordCount > 0 Then
      mlngCodigo = Nulo(rs!Codigo, adRetornaNumeroMenosUm)
   End If
   
   rs.Close
   Set rs = Nothing

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método BuscarCodigoPorCodigoExterno.Módulo de classe.clsProduto" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Public Function VerificarSeJaExiste() As Long
   On Error GoTo CATCH

   Call BuscarCodigo
   VerificarSeJaExiste = mlngCodigo

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método VerificarSeJaExiste.Módulo de classe.clsProduto" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Public Function Gravar() As Boolean
   Dim strSQL As String
   On Error GoTo CATCH

   Gravar = False
   
   If mlngCodigo = -1 Then
      mlngCodigo = GerarNovoCodigo("Produto", "Codigo")
      strSQL = "INSERT INTO Produto " & vbCr & _
               "(Codigo,Nome,Preco,CodigoExterno)" & vbCr & _
               "VALUES (" & vbCr & _
               "" & mlngCodigo & "" & vbCr & _
               ",'" & mstrNome & "'" & vbCr & _
               "," & FormatarNumeroSQL(Str(mdblPreco)) & "" & vbCr & _
               ",'" & mstrCodigoExterno & "'" & vbCr & _
               ")"
   Else
      strSQL = "UPDATE Produto SET " & vbCr & _
               "  Nome = '" & mstrNome & "'" & vbCr & _
               " ,Preco = " & FormatarNumeroSQL(Str(mdblPreco)) & "" & vbCr & _
               " ,CodigoExterno = '" & mstrCodigoExterno & "'" & vbCr & _
               "WHERE Codigo = " & mlngCodigo
   End If
   Call ConexaoBanco.Execute(strSQL)
   Gravar = True

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método Gravar.Módulo de classe.clsProduto" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Public Function Excluir() As Boolean
   Dim strSQL As String
   On Error GoTo CATCH

   Excluir = False
   
   strSQL = "DELETE FROM Produto " & vbCr & _
            "WHERE Codigo = " & mlngCodigo
   Call ConexaoBanco.Execute(strSQL)

   Excluir = True
   
   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método Excluir.Módulo de classe.clsProduto" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function


Public Sub BuscarProduto()
   Dim strSQL As String
   Dim rs As New ADODB.Recordset

   On Error GoTo CATCH

   mstrNome = ""
   mdblPreco = 0
   mstrCodigoExterno = ""
   
   strSQL = "SELECT Nome, Preco, CodigoExterno" & vbCr & _
            "FROM Produto" & vbCr & _
            "WHERE Codigo = " & mlngCodigo & ""

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic
   
   If rs.RecordCount > 0 Then
      mstrNome = Nulo(rs!Nome, adEntradaTexto)
      mdblPreco = Nulo(rs!Preco, adEntradaNumero)
      mstrCodigoExterno = Nulo(rs!CodigoExterno, adEntradaTexto)
   End If
   
   rs.Close
   Set rs = Nothing

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método BuscarPrecoProduto.Módulo de classe.clsProduto" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub


