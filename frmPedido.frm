VERSION 5.00
Object = "{77B8E0B4-8FF5-452F-AD50-BF9AEF6C73F3}#1.1#0"; "randcontrols.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Pedido"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9015
   Icon            =   "frmPedido.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9015
   Begin VB.Frame fraClientes 
      Height          =   795
      Left            =   0
      TabIndex        =   13
      Top             =   -60
      Width           =   9015
      Begin VB.ComboBox cmbCliente 
         Height          =   315
         Left            =   90
         TabIndex        =   0
         Text            =   "Combo"
         Top             =   345
         Width           =   7425
      End
      Begin RandControls.lblControl lblCodigoCliente 
         Height          =   285
         Left            =   7530
         TabIndex        =   10
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Tipo            =   2
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   7530
         TabIndex        =   17
         Top             =   150
         Width           =   540
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   150
         Width           =   525
      End
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Itens do Pedido"
      Height          =   4245
      Left            =   0
      TabIndex        =   14
      Top             =   720
      Width           =   9015
      Begin VB.ComboBox cmbProduto 
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Text            =   "Combo"
         Top             =   465
         Width           =   5985
      End
      Begin VB.CommandButton cmdExcluitTodosItens 
         Height          =   315
         Left            =   8610
         Picture         =   "frmPedido.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluit todos os itens do Pedido"
         Top             =   2130
         Width           =   345
      End
      Begin VB.CommandButton cmdExcluirItem 
         Height          =   315
         Left            =   8610
         Picture         =   "frmPedido.frx":1054
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Excluir item Selecionado do pedido"
         Top             =   1785
         Width           =   345
      End
      Begin VB.CommandButton cmdGravarItem 
         Height          =   315
         Left            =   8610
         Picture         =   "frmPedido.frx":13DE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Incluir Item do Pedido"
         Top             =   1440
         Width           =   345
      End
      Begin MSFlexGridLib.MSFlexGrid gridItemPedido 
         Height          =   2715
         Left            =   60
         TabIndex        =   9
         Top             =   1440
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   4789
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin RandControls.lblControl lblCodigoProduto 
         Height          =   285
         Left            =   6090
         TabIndex        =   11
         Top             =   480
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Tipo            =   2
      End
      Begin RandControls.mskControl txtQuantidade 
         Height          =   285
         Left            =   1530
         TabIndex        =   2
         Top             =   1080
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
         CasasDecimais   =   2
      End
      Begin RandControls.lblControl lblPrecoProduto 
         Height          =   285
         Left            =   90
         TabIndex        =   12
         Top             =   1080
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "0"
         Tipo            =   2
      End
      Begin RandControls.lblControl lblValorTotalPedido 
         Height          =   285
         Left            =   7170
         TabIndex        =   22
         Top             =   1110
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "0"
         Tipo            =   2
      End
      Begin RandControls.lblControl lblValorTotalProduto 
         Height          =   285
         Left            =   2970
         TabIndex        =   24
         Top             =   1080
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "0"
         Tipo            =   2
      End
      Begin RandControls.lblControl lblCodigoExterno 
         Height          =   285
         Left            =   7530
         TabIndex        =   26
         Top             =   480
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Tipo            =   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código Externo:"
         Height          =   195
         Left            =   7530
         TabIndex        =   27
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total Produto:"
         Height          =   195
         Left            =   2970
         TabIndex        =   25
         Top             =   870
         Width           =   1410
      End
      Begin VB.Label lblTotalPedido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total Pedido:"
         Height          =   195
         Left            =   7170
         TabIndex        =   23
         Top             =   870
         Width           =   1350
      End
      Begin VB.Label lblPreco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço:"
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   870
         Width           =   465
      End
      Begin VB.Label lblQuantidade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade:"
         Height          =   195
         Left            =   1530
         TabIndex        =   20
         Top             =   870
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   6090
         TabIndex        =   19
         Top             =   270
         Width           =   540
      End
      Begin VB.Label lblProduto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   270
         Width           =   600
      End
   End
   Begin VB.Frame fraGravarPedido 
      Height          =   525
      Left            =   0
      TabIndex        =   16
      Top             =   4890
      Width           =   9015
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         Height          =   315
         Left            =   6525
         TabIndex        =   7
         Top             =   150
         Width           =   1185
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   315
         Left            =   7710
         TabIndex        =   8
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravarPedido 
         Caption         =   "Gravar &Pedido"
         Height          =   315
         Left            =   5250
         TabIndex        =   6
         Top             =   150
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objCliente As New clsCliente
Private objProduto As New clsProduto
Private objPedido As New clsPedido
Private objPedidoItens As New clsPedidoItens
Private Sub cmbCliente_Click()
   lblCodigoCliente.Caption = cmbCliente.ItemData(cmbCliente.ListIndex)
   objCliente.Codigo = lblCodigoCliente.Caption
   If objCliente.Codigo > 0 Then
      Call objCliente.BuscarCliente
   Else
      lblCodigoCliente.Caption = ""
   End If
End Sub
Private Sub cmbCliente_GotFocus()
   Call ConfigurarSetFocus(cmbCliente)
End Sub
Private Sub cmbCliente_LostFocus()
   Call ConfigurarLostFocus(cmbCliente)
End Sub
Private Sub cmbProduto_Click()
   lblCodigoProduto.Caption = cmbProduto.ItemData(cmbProduto.ListIndex)
   objProduto.Codigo = lblCodigoProduto.Caption
   If objProduto.Codigo > 0 Then
      Call objProduto.BuscarProduto
      lblPrecoProduto.Caption = Format(Nulo(objProduto.Preco, adEntradaNumero), "#,##0.00")
      lblCodigoExterno.Caption = Nulo(objProduto.CodigoExterno, adEntradaTexto)
   Else
      lblCodigoProduto.Caption = ""
      lblPrecoProduto.Caption = Format(0, "#,##0.00")
      lblCodigoExterno.Caption = ""
   End If
   txtQuantidade.Text = Format(0, "#,##0.00")
   lblValorTotalProduto.Caption = Format(0, "#,##0.00")
End Sub
Private Sub cmbProduto_GotFocus()
   Call ConfigurarSetFocus(cmbProduto)
End Sub
Private Sub cmbProduto_LostFocus()
   Call ConfigurarLostFocus(cmbProduto)
   If cmbProduto.Text = "" And gridItemPedido.Rows > 1 Then
      cmdGravarPedido.SetFocus
   End If
End Sub
Private Sub cmdExcluirItem_Click()
   On Error GoTo CATCH

   If MsgBox("Deseja realmente excluir este item do pedido?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
      With gridItemPedido
         If .Row > 0 Then
            If .Row = 1 And .Rows = 2 Then
               Call ConfigurarGrid
            Else
               .RemoveItem (.Row)
            End If
         End If
      End With
      Call CalcularTotalPedido
   End If
   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método cmdExcluirItem_Click.Formulário.frmCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Private Function Validar() As Boolean
   Validar = False
   If lblCodigoCliente.Caption = "" Then
      Call MsgBox("Informe o cliente!", vbExclamation, "Atenção")
      cmbCliente.SetFocus
      Exit Function
   ElseIf gridItemPedido.Rows <= 1 Then
      Call MsgBox("Nenhum produto informado!", vbExclamation, "Atenção")
      txtQuantidade.SetFocus
      Exit Function
   ElseIf CDbl(lblValorTotalPedido.Caption) > objCliente.LimiteCredito Then
      Call MsgBox("Limite de crédito insuficiente para realizar este pedido!" & vbCr & _
                  "O cliente possui um limite de R$" & Format(objCliente.LimiteCredito, "#,##0.00") & ".", vbExclamation, "Atenção")
      cmbCliente.SetFocus
      Exit Function
   End If
   Validar = True
End Function
Private Function ValidarItem() As Boolean
   ValidarItem = False
   If lblCodigoProduto.Caption = "" Then
      Call MsgBox("Informe o Produto!", vbExclamation, "Atenção")
      cmbProduto.SetFocus
      Exit Function
   ElseIf CDbl(Nulo(lblPrecoProduto.Caption, adEntradaNumero)) <= 0 Then
      Call MsgBox("Preço do Produto inválido!", vbExclamation, "Atenção")
      cmbProduto.SetFocus
      Exit Function
   ElseIf CDbl(Nulo(txtQuantidade.Text, adEntradaNumero)) <= 0 Then
      Call MsgBox("Quantidade do Produto inválido!", vbExclamation, "Atenção")
      txtQuantidade.SetFocus
      Exit Function
   ElseIf VerificarItemJaIncluido Then
      Call MsgBox("Item já incluido no pedido!", vbExclamation, "Atenção")
      cmbProduto.SetFocus
      Exit Function
   End If
   ValidarItem = True
End Function
Private Sub cmdExcluitTodosItens_Click()
   On Error GoTo CATCH

   If MsgBox("Deseja realmente excluir TODOS os item do pedido?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
      Call LimparItens
      Call ConfigurarGrid
      Call CalcularTotalPedido
      cmbProduto.SetFocus
   End If
   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método cmdExcluitTodosItens_Click.Formulário.frmPedido" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Private Sub cmdGravarItem_Click()
   On Error GoTo CATCH

   If ValidarItem Then
      gridItemPedido.AddItem Nulo(lblCodigoExterno.Caption, adEntradaTexto) & Chr$(9) & _
                             Nulo(cmbProduto.Text, adEntradaTexto) & Chr$(9) & _
                             Format(Nulo(lblPrecoProduto.Caption, adEntradaNumero), "#,##0.00") & Chr$(9) & _
                             Nulo(txtQuantidade.Text, adEntradaTexto) & Chr$(9) & _
                             Format(Nulo(lblValorTotalProduto.Caption, adEntradaNumero), "#,##0.00") & Chr$(9) & _
                             Nulo(lblCodigoProduto.Caption, adEntradaNumero)
   
   Call LimparItens
   cmbProduto.SetFocus
   End If
   Call CalcularTotalPedido

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método cmdGravarItem_Click.Formulário.frmPedido" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub

Private Sub cmdGravarPedido_Click()
Dim lngContador As Long
   On Error GoTo CATCH

   If Validar Then
      ConexaoBanco.BeginTrans
      With objPedido
         .Codigo = -1
         .CodigoCliente = lblCodigoCliente.Caption
         .ValorTotal = lblValorTotalPedido.Caption
         Call .Gravar
      End With
      
      With objPedidoItens
         Call .Excluir
         For lngContador = 1 To gridItemPedido.Rows - 1
            .CodigoPedido = objPedido.Codigo
            .CodigoProduto = gridItemPedido.TextMatrix(lngContador, 5)
            .Preco = gridItemPedido.TextMatrix(lngContador, 2)
            .Quantidade = gridItemPedido.TextMatrix(lngContador, 3)
            .ValorTotal = gridItemPedido.TextMatrix(lngContador, 4)
            .Gravar
         Next
      End With
      
      objCliente.LimiteCredito = Format(objCliente.LimiteCredito - CDbl(lblValorTotalPedido.Caption), "#,##0.00")
      objCliente.AtualizarLimiteCredito
      
      ConexaoBanco.CommitTrans
      
      Call MsgBox("Pedido cadastrado com sucesso!", vbInformation, "Atenção")
      Call LimparTela
      
   End If

   Exit Sub
CATCH:
   ConexaoBanco.RollbackTrans
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método cmdGravarPedido_Click.Formulário.frmPedido" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Private Sub cmdLimpar_Click()
   Call LimparTela
End Sub
Private Sub cmdSair_Click()
   Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Call SimularEnter(KeyAscii)
End Sub
Private Sub Form_Load()
   Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 4
   
   Call CarregarComboCliente
   Call CarregarComboProduto
   Call cmdLimpar_Click

End Sub
Private Sub ConfigurarGrid()
   With gridItemPedido
      .Cols = 6
      .Rows = 1
      .ColWidth(0) = 1200
      .ColWidth(1) = 3500
      .ColWidth(2) = 1200
      .ColWidth(3) = 1200
      .ColWidth(4) = 1200
      .ColWidth(5) = 0

      .FixedAlignment(0) = 2
      .FixedAlignment(1) = 2
      .FixedAlignment(2) = 2
      .FixedAlignment(3) = 2
      .FixedAlignment(4) = 2
      .FixedAlignment(5) = 2

      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .ColAlignment(3) = 6
      .ColAlignment(4) = 6
      .ColAlignment(5) = 0

      .Row = 0
      .Col = 0
      .Text = "Codigo Externo"
      .Col = 1
      .Text = "Nome"
      .Col = 2
      .Text = "Preço"
      .Col = 3
      .Text = "Quantidade"
      .Col = 4
      .Text = "Total"
      .Col = 5
      .Text = "CodigoProduto"
   End With
End Sub
Private Sub LimparTela()
   cmbCliente.ListIndex = 0
   cmbProduto.ListIndex = 0
   lblPrecoProduto.Caption = 0
   txtQuantidade.Text = 0
   lblValorTotalPedido.Caption = 0
   Call ConfigurarGrid
End Sub
Private Sub LimparItens()
   cmbProduto.ListIndex = 0
   lblPrecoProduto.Caption = 0
   txtQuantidade.Text = 0
End Sub
Private Sub CarregarComboCliente()
   Dim rs As New ADODB.Recordset
   
   cmbCliente.Clear
   cmbCliente.AddItem ""
   cmbCliente.ItemData(cmbCliente.NewIndex) = -1
   Set rs = objCliente.rsCliente
   With rs
   Do While Not .EOF
      cmbCliente.AddItem Nulo(!Nome, adEntradaTexto)
      cmbCliente.ItemData(cmbCliente.NewIndex) = Nulo(!Codigo, adEntradaTexto)
      .MoveNext
   Loop
   End With

End Sub
Private Sub CarregarComboProduto()
   Dim rs As New ADODB.Recordset
   
   cmbProduto.Clear
   cmbProduto.AddItem ""
   cmbProduto.ItemData(cmbProduto.NewIndex) = -1
   Set rs = objProduto.rsProduto
   With rs
   Do While Not .EOF
      cmbProduto.AddItem Nulo(!Nome, adEntradaTexto)
      cmbProduto.ItemData(cmbProduto.NewIndex) = Nulo(!Codigo, adEntradaTexto)
      .MoveNext
   Loop
   End With

End Sub
Private Sub txtQuantidade_Change()
   Dim dblValorProduto As Double
   If txtQuantidade.Text <> "" And lblPrecoProduto.Caption <> "" Then
      dblValorProduto = CDbl(txtQuantidade.Text) * CDbl(lblPrecoProduto.Caption)
   Else
      dblValorProduto = 0
   End If
   lblValorTotalProduto.Caption = Format(dblValorProduto, "#,##0.00")
End Sub
Private Function VerificarItemJaIncluido() As Boolean
Dim lngContador As Long
   VerificarItemJaIncluido = False
   With gridItemPedido
      For lngContador = 1 To .Rows - 1
         If gridItemPedido.TextMatrix(lngContador, 0) = lblCodigoExterno.Caption Then
            VerificarItemJaIncluido = True
            Exit For
         End If
      Next
   End With
End Function
Private Sub CalcularTotalPedido()
Dim lngContador As Long
Dim dblValorTotalPedido As Double

   dblValorTotalPedido = 0
   With gridItemPedido
      For lngContador = 1 To .Rows - 1
         dblValorTotalPedido = dblValorTotalPedido + gridItemPedido.TextMatrix(lngContador, 4)
      Next
   End With
   lblValorTotalPedido.Caption = Format(dblValorTotalPedido, "#,##0.00")
   
End Sub
