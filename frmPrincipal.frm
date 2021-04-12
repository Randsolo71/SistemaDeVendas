VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.MDIForm frmPrincipal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Controle de Vendas"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   10545
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "MDIForm"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList 
      Left            =   90
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":0CCA
            Key             =   "cliente"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":19A4
            Key             =   "pedido"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":267E
            Key             =   "produto"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":3358
            Key             =   "usuario"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":4032
            Key             =   "vendas"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":4D0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cliente"
            Description     =   "Cliente"
            Object.ToolTipText     =   "Cadastro de Clientes"
            Object.Tag             =   "cliente"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "produto"
            Description     =   "Produtos"
            Object.ToolTipText     =   "Cadastro de Produtos"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pedido"
            Description     =   "Pedido"
            Object.ToolTipText     =   "Pedido de Vendas"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sair"
            Description     =   "Sair"
            Object.ToolTipText     =   "Sair do sistema"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "Data:"
            TextSave        =   "Data:"
            Key             =   "Data"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   10372
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Picture         =   "frmPrincipal.frx":59E6
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "Usuario"
            TextSave        =   "Usuario"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCAdastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnuCadstroCliente 
         Caption         =   "&Cliente"
      End
      Begin VB.Menu mnuCadastroProduto 
         Caption         =   "&Produto"
      End
      Begin VB.Menu mnuCadastroUsuario 
         Caption         =   "&Usuario"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPedido 
      Caption         =   "&Pedido"
   End
   Begin VB.Menu mnuSair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
   StatusBar.Panels.Item(1).Text = "Data: " & Format(Now, "dd/mm/yyyy")
   StatusBar.Panels.Item(4).Text = "Usuario: " & pstrUsuarioLogado
End Sub
Private Sub mnuCadastroProduto_Click()
   Call CarregaForm(frmProduto)
End Sub
Private Sub mnuCadstroCliente_Click()
   Call CarregaForm(frmCliente)
End Sub
Private Sub mnuPedido_Click()
   Call CarregaForm(frmPedido)
End Sub
Private Sub mnuSair_Click()
   Call FecharBanco
   End
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   On Error GoTo CATCH

   MousePointer = vbHourglass
   
   Select Case Button.Key
      Case "cliente": Call CarregaForm(frmCliente)
      Case "produto": Call CarregaForm(frmProduto)
      Case "pedido": Call CarregaForm(frmPedido)
      Case "sair": Call mnuSair_Click
   End Select
   
   MousePointer = vbDefault

   Exit Sub
CATCH:
   Call MsgBox("Ocorreu o erro " & Err.Number & " (" & Err.Description & ") no método Toolbar_ButtonClick.Formulário.frmPrincipal" & vbCr & _
               "Por favor, consulte o suporte técnico!")
End Sub
