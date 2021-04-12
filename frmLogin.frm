VERSION 5.00
Object = "{77B8E0B4-8FF5-452F-AD50-BF9AEF6C73F3}#1.1#0"; "randcontrols.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   9.419
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   13.309
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   285
      Left            =   6180
      TabIndex        =   5
      Top             =   3180
      Width           =   1035
   End
   Begin VB.CommandButton cmdLogar 
      Caption         =   "&Logar"
      Height          =   285
      Left            =   5100
      TabIndex        =   4
      Top             =   3180
      Width           =   1035
   End
   Begin RandControls.txtControl txtUsuario 
      Height          =   315
      Left            =   4530
      TabIndex        =   0
      Top             =   2160
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RandControls.txtControl txtSenha 
      Height          =   315
      Left            =   4530
      TabIndex        =   2
      Top             =   2760
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "*"
   End
   Begin VB.Label lblBemVindo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seja bem-vindo!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   4440
      TabIndex        =   6
      Top             =   1440
      Width           =   1980
   End
   Begin VB.Image imgChave 
      Height          =   480
      Left            =   3870
      Picture         =   "frmLogin.frx":B77D
      Top             =   1890
      Width           =   480
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      Height          =   195
      Left            =   4530
      TabIndex        =   3
      Top             =   2550
      Width           =   510
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      Height          =   195
      Left            =   4530
      TabIndex        =   1
      Top             =   1920
      Width           =   585
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FEE9EA&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1815
      Left            =   4410
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   2925
   End
   Begin VB.Shape ShpSombra 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1815
      Left            =   4470
      Shape           =   4  'Rounded Rectangle
      Top             =   1860
      Width           =   2925
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdLogar_Click()
   Dim objUsuario As New clsUsuario

   If Validar Then
      objUsuario.Login = txtUsuario.Text
      objUsuario.Senha = txtSenha.Text

      If objUsuario.ValidarUsuario Then
         pstrUsuarioLogado = objUsuario.Nome
         Call CarregaForm(frmPrincipal)
         Unload Me
      Else
         Call MsgBox("Usuário inválido!", vbExclamation, "Atenção")
      End If
   End If
End Sub
Private Sub cmdSair_Click()
   End
End Sub
Private Function Validar() As Boolean
   Validar = False
   If txtUsuario.Text = "" Then
      Call MsgBox("Informe o login de usuário!", vbExclamation, "Atenção")
      txtUsuario.SetFocus
      Exit Function
   ElseIf txtSenha.Text = "" Then
      Call MsgBox("Informe a senha do usuário!", vbExclamation, "Atenção")
      txtSenha.SetFocus
      Exit Function
   End If
   Validar = True
   
End Function
Private Sub Form_KeyPress(KeyAscii As Integer)
   Call SimularEnter(KeyAscii)
End Sub
Private Sub Form_Load()
   Call AbrirBanco
End Sub
