VERSION 5.00
Object = "{77B8E0B4-8FF5-452F-AD50-BF9AEF6C73F3}#1.1#0"; "randcontrols.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6945
   Icon            =   "frmCliente.frx":0000
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
      TabIndex        =   9
      Top             =   -60
      Width           =   6945
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         Height          =   315
         Left            =   4440
         TabIndex        =   6
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   315
         Left            =   2010
         TabIndex        =   4
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   315
         Left            =   3225
         TabIndex        =   5
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   315
         Left            =   5655
         TabIndex        =   7
         Top             =   2220
         Width           =   1215
      End
      Begin RandControls.mskControl txtLimiteCredito 
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   1530
         Width           =   1785
         _ExtentX        =   3149
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
      Begin RandControls.mskControl txtTelefone 
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
         Mask            =   "(99)9999-9999"
         MaxLength       =   13
         TipoDados       =   10
         Text            =   "(__)____-____"
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
      Begin RandControls.mskControl txtCelular 
         Height          =   285
         Left            =   3540
         TabIndex        =   2
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
         Mask            =   "(99)99999-9999"
         MaxLength       =   14
         TipoDados       =   11
         Text            =   "(__)_____-____"
      End
      Begin VB.Label lblCelular 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Celular:"
         Height          =   195
         Left            =   3540
         TabIndex        =   14
         Top             =   720
         Width           =   525
      End
      Begin VB.Label lblLimteCredito 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite de crédito:"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   1290
         Width           =   1200
      End
      Begin VB.Label lblTelefone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   720
         Width           =   675
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   180
         Width           =   465
      End
   End
   Begin VB.Frame fraGrid 
      Height          =   2445
      Left            =   0
      TabIndex        =   10
      Top             =   2520
      Width           =   6945
      Begin MSFlexGridLib.MSFlexGrid gridCliente 
         Height          =   2235
         Left            =   60
         TabIndex        =   8
         Top             =   150
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   3942
         _Version        =   393216
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objCliente As New clsCliente
Private Sub cmdExcluir_Click()
   On Error GoTo CATCH

   If MsgBox("Deseja realmente excluir este cliente?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
      If objCliente.Codigo > 0 Then
         If objCliente.Excluir Then
            Call LimparTela
            Call CarregarGrid
            Call MsgBox("Registro excluído com sucesso!", vbInformation, "Atenção")
         End If
      End If
   End If
   
   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método cmdExcluir_Click.Formulário.frmCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Private Sub cmdGravar_Click()
   On Error GoTo CATCH

   If Validar Then
      With objCliente
         .Nome = txtNome.Text
         .Tipo = "C"
         If .Codigo <= 0 Then Call .VerificarSeJaExiste
         .Telefone = txtTelefone.Text
         .Celular = txtCelular.Text
         .LimiteCredito = txtLimiteCredito.Text
         Call .Gravar
      End With
      Call LimparTela
      Call CarregarGrid
      Call MsgBox("Registro salvo com sucesso!", vbInformation, "Atenção")
   End If

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método cmdGravar_Click.Formulário.frmCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Private Function Validar() As Boolean
   Validar = False
   If txtNome.Text = "" Then
      Call MsgBox("Informe o nome do cliente!", vbExclamation, "Atenção")
      txtNome.SetFocus
      Exit Function
   ElseIf Nulo(txtLimiteCredito.Text, enuTipoEntradaDados.adEntradaNumero, True) <= 0 Then
      Call MsgBox("Informe o limite de crédito do cliente!", vbExclamation, "Atenção")
      txtLimiteCredito.SetFocus
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
   
   With gridCliente
      Set rs = objCliente.rsCliente
      If rs.RecordCount > 0 Then
         rs.MoveFirst
         Do While Not rs.EOF
            .AddItem Nulo(rs!Codigo, adEntradaNumero) & Chr$(9) & _
                     Nulo(rs!Nome, adEntradaTexto) & Chr$(9) & _
                     Nulo(rs!Telefone, adEntradaTexto) & Chr$(9) & _
                     Nulo(rs!Celular, adEntradaTexto) & Chr$(9) & _
                     Format(Nulo(rs!LimiteCredito, adEntradaNumero), "#,##0.00") & Chr$(9) & _
                     Nulo(rs!CodigoPessoa, adEntradaNumero)

            
            rs.MoveNext
         Loop
      End If
   End With

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método CarregarGrid.Formulário.frmCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
Private Sub ConfigurarGrid()
   With gridCliente
      .Cols = 6
      .Rows = 1
      .ColWidth(0) = 600
      .ColWidth(1) = 2000
      .ColWidth(2) = 1200
      .ColWidth(3) = 1300
      .ColWidth(4) = 1300
      .ColWidth(5) = 0

      '.RowHeight(0) = 460

      .FixedAlignment(0) = 2
      .FixedAlignment(1) = 2
      .FixedAlignment(2) = 2
      .FixedAlignment(3) = 2
      .FixedAlignment(4) = 2
      .FixedAlignment(5) = 2

      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 0

      .Row = 0
      .Col = 0
      .Text = "Codigo"
      .Col = 1
      .Text = "Nome"
      .Col = 2
      .Text = "Telefone"
      .Col = 3
      .Text = "Celular"
      .Col = 4
      .Text = "Limite de Crédito"
      .Col = 5
      .Text = "CodigoPessoa"
   End With
End Sub
Private Sub LimparTela()
   objCliente.Codigo = -1
   txtNome.Text = ""
   txtTelefone.Text = "(__)____-____"
   txtCelular.Text = "(__)____-____"
   txtLimiteCredito.Text = 0
End Sub
Private Sub gridCliente_DblClick()
   On Error GoTo CATCH
   With gridCliente
      .Col = 0
      .Row = .RowSel
      objCliente.Codigo = .TextMatrix(.Row, 0)
      txtNome.Text = .TextMatrix(.Row, 1)
      txtTelefone.Text = .TextMatrix(.Row, 2)
      txtCelular.Text = .TextMatrix(.Row, 3)
      txtLimiteCredito.Text = .TextMatrix(.Row, 4)
      objCliente.CodigoPessoa = .TextMatrix(.Row, 5)
   End With

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método gridCliente_DblClick.Formulário.frmCliente" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
