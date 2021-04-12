Attribute VB_Name = "modBancoDados"
Option Explicit

Public Function AbrirBanco() As Boolean
   Dim strConexao As String
   On Error GoTo CATCH

   AbrirBanco = False

   Set ConexaoBanco = New ADODB.Connection
   strConexao = "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
                "SERVER=localhost;" & _
                "DATABASE=VendasLinear;" & _
                "USER=Randsolo;" & _
                "PASSWORD=SqL@PassW0rd;" & _
                "OPTION=3;"
   With ConexaoBanco
      .ConnectionString = strConexao
      .CursorLocation = adUseClient
      .CommandTimeout = 0
      .Open
   End With

   AbrirBanco = True

   Exit Function
CATCH:
   If Err.Number = -2147467259 Then
      MsgBox "Não foi possível estabelecer uma conexão com o banco de dados." & vbCrLf & "Verifique se os dados digitados estão corretos ou se sua rede está funcionando corretamente.", vbExclamation
   ElseIf Err.Number = -2147217843 Then
      MsgBox "Não foi possível abir o banco de dados." & vbCrLf & "Verifique se os dados digitados estão corretos ou se sua rede está funcionando corretamente.", vbExclamation
   Else
      Call MsgBox("Erro ao abrir o banco de dados " & vbCrLf & "Erro: " & Err.Number & " (" & Err.Description & ") na rotina AbrirBanco de Módulo.BancoDados", vbExclamation)
   End If
End Function
Public Sub FecharBanco()
   On Error Resume Next
   ConexaoBanco.Close
   Set ConexaoBanco = Nothing
End Sub
Function GerarNovoCodigo(Tabela As String, Campo As String) As Long
   Dim strSQL As String
   Dim rs As New ADODB.Recordset
   On Error GoTo CATCH

   GerarNovoCodigo = 1

   strSQL = "SELECT Max(" & Tabela & "." & Campo & ") AS MaxReg " & vbCr & _
            "FROM " & Tabela & vbCr & _
            ""

   rs.CursorLocation = adUseClient
   rs.Open strSQL, ConexaoBanco, adOpenForwardOnly, adLockReadOnly
   
   If Not rs.EOF Then
      If Not IsNull(rs!MaxReg) Then
         GerarNovoCodigo = rs!MaxReg + 1
      Else
         GerarNovoCodigo = 1
      End If
   End If

   rs.Close
   Set rs = Nothing

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método GerarNovoCodigo.Módulo.modBancoDados" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Public Function FormatarNumeroSQL(strNumero_ As String) As String
   FormatarNumeroSQL = Replace(strNumero_, ",", ".")
End Function
