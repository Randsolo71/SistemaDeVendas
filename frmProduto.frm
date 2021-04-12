VERSION 5.00
Object = "{77B8E0B4-8FF5-452F-AD50-BF9AEF6C73F3}#1.1#0"; "randcontrols.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmProduto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6945
   Icon            =   "frmProduto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6945
   Begin VB.Frame fraClientes 
      Height          =   2625
      Left            =   0
      TabIndex        =   8
      Top             =   -60
      Width           =   6945
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         Height          =   315
         Left            =   4440
         TabIndex        =   5
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   315
         Left            =   5655
         TabIndex        =   6
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   315
         Left            =   3225
         TabIndex        =   4
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   315
         Left            =   2010
         TabIndex        =   3
         Top             =   2220
         Width           =   1215
      End
      Begin RandControls.mskControl txtPreco 
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   930
         Width           =   3345
         _ExtentX        =   5900
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
         MaxLength       =   15
         CasasDecimais   =   2
      End
      Begin RandControls.txtControl txtNome 
         Height          =   285
         Left            =   90
         TabIndex        =   0
         Top             =   390
         Width           =   6795
         _ExtentX        =   11986
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
         MaxLength       =   100
      End
      Begin RandControls.txtControl txtCodigoExterno 
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   1500
         Width           =   3345
         _ExtentX        =   5900
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
         Maiusculo       =   1
      End
      Begin VB.Label lblLimteCredito 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código Externo:"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   1290
         Width           =   1125
      End
      Begin VB.Label lblPreco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço:"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   720
         Width           =   465
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Width           =   465
      End
   End
   Begin VB.Frame fraGrid 
      Height          =   2445
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   6945
      Begin MSFlexGridLib.MSFlexGrid gridProduto 
         Height          =   2235
         Left            =   60
         TabIndex        =   7
         Top             =   150
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   3942
         _Version        =   393216
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objProduto As New clsProduto
Private Sub cmdExcluir_Click()
   On Error GoTo CATCH

   If MsgBox("Deseja realmente excluir este produto?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
      If objProduto.Codigo > 0 Then
         If objProduto.Excluir Then
            Call LimparTela
            Call CarregarGrid
            Call MsgBox("Registro excluído com sucesso!", vbInformation, "Atenção")
         End If
      End If
   End If
   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método cmdExcluir_Click.Formulário.frmProduto" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Private Sub cmdGravar_Click()
   On Error GoTo CATCH

   If Validar Then
      With objProduto
         .Nome = txtNome.Text
         .Preco = txtPreco.Text
         .CodigoExterno = txtCodigoExterno.Text
         Call .Gravar
      End With
      Call LimparTela
      Call CarregarGrid
      Call MsgBox("Registro salvo com sucesso!", vbInformation, "Atenção")
   End If

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método cmdGravar_Click.Formulário.frmProduto" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Private Function Validar() As Boolean
   Validar = False
   If txtNome.Text = "" Then
      Call MsgBox("Informe o nome do Produto!", vbExclamation, "Atenção")
      txtNome.SetFocus
      Exit Function
   ElseIf Nulo(txtPreco.Text, enuTipoEntradaDados.adEntradaNumero, True) <= 0 Then
      Call MsgBox("Informe o preço do Produto!", vbExclamation, "Atenção")
      txtPreco.SetFocus
      Exit Function
   ElseIf Nulo(txtCodigoExterno.Text, enuTipoEntradaDados.adEntradaNumero, True) <= 0 Then
      Call MsgBox("Informe o código externo do Produto!", vbExclamation, "Atenção")
      txtCodigoExterno.SetFocus
      Exit Function
   End If
   Validar = True
End Function

Private Sub cmdLimpar_Click()
   Call LimparTela
   Call CarregarGrid
End Sub
Private Sub cmdSair_Click()
   Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Call SimularEnter(KeyAscii)
End Sub
Private Sub Form_Load()
   Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 4
   Call cmdLimpar_Click
   
End Sub
Private Sub CarregarGrid()
   Dim rs As New ADODB.Recordset
   
   On Error GoTo CATCH

   Call ConfigurarGrid
   
   With gridProduto
      Set rs = objProduto.rsProduto
      If rs.RecordCount > 0 Then
         rs.MoveFirst
         Do While Not rs.EOF
            .AddItem Nulo(rs!Codigo, adEntradaNumero) & Chr$(9) & _
                     Nulo(rs!Nome, adEntradaTexto) & Chr$(9) & _
                     Format(Nulo(rs!Preco, adEntradaNumero), "#,##0.00") & Chr$(9) & _
                     Nulo(rs!CodigoExterno, adEntradaNumero)
            rs.MoveNext
         Loop
      End If
   End With

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método CarregarGrid.Formulário.frmProduto" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Private Sub ConfigurarGrid()
   With gridProduto
      .Cols = 4
      .Rows = 1
      .ColWidth(0) = 600
      .ColWidth(1) = 3400
      .ColWidth(2) = 1100
      .ColWidth(3) = 1800

      .FixedAlignment(0) = 2
      .FixedAlignment(1) = 2
      .FixedAlignment(2) = 2
      .FixedAlignment(3) = 2

      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 6
      .ColAlignment(3) = 1

      .Row = 0
      .Col = 0
      .Text = "Codigo"
      .Col = 1
      .Text = "Nome"
      .Col = 2
      .Text = "Preço"
      .Col = 3
      .Text = "Código Externo"
   End With
End Sub
Private Sub LimparTela()
   objProduto.Codigo = -1
   txtNome.Text = ""
   txtPreco.Text = 0
   txtCodigoExterno.Text = ""
End Sub
Private Sub gridProduto_DblClick()
   On Error GoTo CATCH
   With gridProduto
      .Col = 0
      .Row = .RowSel
      objProduto.Codigo = .TextMatrix(.Row, 0)
      txtNome.Text = .TextMatrix(.Row, 1)
      txtPreco.Text = .TextMatrix(.Row, 2)
      txtCodigoExterno.Text = .TextMatrix(.Row, 3)
   End With

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método gridProduto_DblClick.Formulário.frmProduto" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
