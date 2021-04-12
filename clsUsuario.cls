VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Implements IPessoa
Private objPessoa As New clsPessoa

Private mlngCodigo As Long
Private mstrLogin As String
Private mstrSenha As String
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
Private Property Let Telefone(ByVal strTelefone_ As String)
   objPessoa.Telefone = strTelefone_
End Property
Private Property Get Telefone() As String
   Telefone = objPessoa.Telefone
End Property
Private Property Let Celular(ByVal strCelular_ As String)
   objPessoa.Celular = strCelular_
End Property
Private Property Get Celular() As String
   Celular = objPessoa.Celular
End Property
Public Property Let Codigo(ByVal lngCodigo_ As Long)
   mlngCodigo = lngCodigo_
End Property
Public Property Get Codigo() As Long
   Codigo = mlngCodigo
End Property
Public Property Let Senha(ByVal strSenha_ As String)
   mstrSenha = strSenha_
End Property
Public Property Get Senha() As String
   Senha = mstrSenha
End Property
Public Property Let Login(ByVal strLogin_ As String)
   mstrLogin = strLogin_
End Property
Public Property Get Login() As String
   Login = mstrLogin
End Property
Private Function Gravar() As Boolean
   'TODO
   Gravar = True
End Function
Private Function Excluir() As Boolean
   'TODO
   Excluir = True
End Function
Public Function rsUsuario(Optional strWhere_ As String = "") As ADODB.Recordset
   Dim strSQL As String
   Dim rs As New ADODB.Recordset

   On Error GoTo CATCH

   strSQL = "SELECT U.Codigo,U.Login,U.Senha," & vbCr & _
          "       P.Nome, P.Telefone, P.Celular" & vbCr & _
            "FROM Usuario as U INNER JOIN Pessoas as P ON P.Codigo = U.CodigoPessoa" & vbCr
   If strWhere_ <> "" Then
      strSQL = strSQL & "WHERE " & strWhere_
   End If

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic
   Set rsUsuario = rs

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método rsUsuario.Módulo de classe.clsUsuario" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Public Function ValidarUsuario() As Boolean
   Dim strSQL As String
   Dim rs As New ADODB.Recordset

   On Error GoTo CATCH

   ValidarUsuario = False

   strSQL = "SELECT U.Codigo,U.Login,U.Senha, " & vbCr & _
            "       P.Nome, P.Telefone" & vbCr & _
            "FROM Usuario as U LEFT JOIN Pessoa as P ON P.Codigo = U.Codigo" & vbCr & _
            "WHERE Login = '" & mstrLogin & "'" & vbCr & _
            "  AND Senha   = '" & Criptografar(mstrSenha) & "'"

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenKeyset, adLockOptimistic

   If rs.RecordCount = 1 Then
      objPessoa.Nome = "" & rs!Nome
      ValidarUsuario = True
   End If

   rs.Close
   Set rs = Nothing

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método ValidarUsuario.Módulo de classe.clsUsuario" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function

