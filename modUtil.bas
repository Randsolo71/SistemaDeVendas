Attribute VB_Name = "modUtil"
Option Explicit
Const CHAVE = "DEATHSTAR"
Public Function Criptografar(strTexto_ As String) As String
   Dim lngContador As Long
   Dim intIndiceChave As Integer
   Dim intCharTemporario As Integer
   Dim strStringTemporaria As String
   Dim strChaveMinimizada As String

   On Error GoTo CATCH

   strChaveMinimizada = LCase(CHAVE)

   Criptografar = ""
   intIndiceChave = 1

   For lngContador = 1 To Len(strTexto_)
      intCharTemporario = Asc(Mid(strTexto_, lngContador, 1)) + Asc(Mid(strChaveMinimizada, intIndiceChave, 1))
      strStringTemporaria = strStringTemporaria & Chr(intCharTemporario)
      intIndiceChave = intIndiceChave + 1
      If intIndiceChave > Len(strChaveMinimizada) Then
         intIndiceChave = 1
      End If
   Next lngContador

   Criptografar = strStringTemporaria

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método Criptografar.Módulo.Util" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Public Function Descriptografar(strTexto_ As String)
   Dim lngContador As Integer
   Dim intIndiceChave As Integer
   Dim intCharTemporario As Integer
   Dim strStringTemporaria As String
   Dim strChaveMinimizada As String

   On Error GoTo CATCH

   strChaveMinimizada = LCase(CHAVE)

   Descriptografar = ""
   intIndiceChave = 1

   For lngContador = 1 To Len(strTexto_)
      intCharTemporario = Asc(Mid(strTexto_, lngContador, 1)) - Asc(Mid(strChaveMinimizada, intIndiceChave, 1))
      If intCharTemporario > -1 Then
         strStringTemporaria = strStringTemporaria & Chr(intCharTemporario)
         intIndiceChave = intIndiceChave + 1
      End If
      If intIndiceChave > Len(strChaveMinimizada) Then
         intIndiceChave = 1
      End If
   Next lngContador

   Descriptografar = strStringTemporaria

   Exit Function
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método Descriptografar.Módulo.Util" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Function
Public Sub CarregaForm(frmForm_ As Form, Optional blnModal_ As Boolean)
On Error Resume Next
   With frmForm_
      If blnModal_ Then
         .Show vbModal, frmPrincipal
      Else
         .Show
         .SetFocus
      End If
    End With
End Sub
Public Function Nulo(varCampo_ As Variant, Optional TipoDado_ As enuTipoEntradaDados, _
                     Optional blnTestarVazio_ As Boolean) As Variant
   On Error Resume Next

   If IsNull(varCampo_) Or (varCampo_ = "" And blnTestarVazio_) Then
      Select Case TipoDado_
      Case enuTipoEntradaDados.adEntradaData
         Nulo = Date
      Case enuTipoEntradaDados.adEntradaNumero
         Nulo = 0
      Case enuTipoEntradaDados.adEntradaBoolean
         Nulo = False
      Case enuTipoEntradaDados.adRetornaNumeroMenosUm
         Nulo = -1
      Case Else
         If IsNull(varCampo_) Then
            Nulo = ""
         Else
            Nulo = varCampo_
         End If
      End Select
   Else
      Nulo = varCampo_
   End If
End Function
Public Sub SimularEnter(ByRef intKeyAscii_ As Integer)
   Select Case intKeyAscii_
     Case vbKeyReturn
         intKeyAscii_ = 0
         Sendkeys "{Tab}"
      Case vbKeyEscape
         intKeyAscii_ = 0
         Sendkeys "+{Tab}"
   End Select
End Sub
Private Sub Sendkeys(strTexto_ As String, Optional blnWait_ As Boolean = False)
    On Error Resume Next
    Dim WshShell As Object
    Set WshShell = CreateObject("wscript.shell")
    WshShell.Sendkeys strTexto_, blnWait_
    Set WshShell = Nothing
End Sub
Public Sub ConfigurarSetFocus(objObjeto_ As Object)
On Error Resume Next
   With objObjeto_
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = COR_SELECAO
   End With
End Sub
Public Sub ConfigurarLostFocus(objObjeto_ As Object)
On Error Resume Next
   With objObjeto_
      .BackColor = &H80000005
   End With
End Sub
