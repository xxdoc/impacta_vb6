VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustoProduto 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custo do Produto - Valores Unitários"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11670
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustoProduto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   11670
   Begin VB.Frame fraResultado 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7404&
      Height          =   1530
      Left            =   30
      TabIndex        =   48
      Top             =   4920
      Width           =   11655
      Begin VB.TextBox txtMaximoCustoPorAquisicao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3870
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   53
         Text            =   "0,00"
         Top             =   1095
         Width           =   2265
      End
      Begin VB.TextBox txtLucroFinalPorcentagem 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3870
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   51
         Text            =   "0,00"
         Top             =   675
         Width           =   2265
      End
      Begin VB.TextBox txtLucroFinal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3870
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   49
         Text            =   "0,00"
         Top             =   270
         Width           =   2265
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Máximo Custo Por Aquisição (CPA)"
         Height          =   240
         Index           =   17
         Left            =   105
         TabIndex        =   54
         Top             =   1095
         Width           =   3405
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lucro Final %"
         Height          =   240
         Index           =   16
         Left            =   90
         TabIndex        =   52
         Top             =   675
         Width           =   1335
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lucro Final"
         Height          =   240
         Index           =   15
         Left            =   105
         TabIndex        =   50
         Top             =   285
         Width           =   1065
      End
   End
   Begin VB.Frame fraBotoes 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7404&
      Height          =   3795
      Left            =   10545
      TabIndex        =   40
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton cmdConsultar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consultar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   45
         Picture         =   "frmCustoProduto.frx":1486
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir Registro"
         Top             =   180
         Width           =   950
      End
      Begin VB.CommandButton cmdSalvar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Salvar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   45
         Picture         =   "frmCustoProduto.frx":1F90
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salvar Registro"
         Top             =   1950
         Width           =   950
      End
      Begin VB.CommandButton cmdExcluir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Excluir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   45
         Picture         =   "frmCustoProduto.frx":2A9A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir Registro"
         Top             =   2835
         Width           =   950
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   45
         Picture         =   "frmCustoProduto.frx":35A4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar Campos"
         Top             =   1065
         Width           =   950
      End
   End
   Begin VB.Frame fraTotalCustos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total (Produto + Taxas + Marketing)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7404&
      Height          =   750
      Left            =   5070
      TabIndex        =   38
      Top             =   4185
      Width           =   6600
      Begin VB.TextBox txtTotalCustoOperacional 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3885
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   26
         Text            =   "0,00"
         Top             =   285
         Width           =   1725
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Custo Operacional"
         Height          =   240
         Index           =   11
         Left            =   165
         TabIndex        =   39
         Top             =   285
         Width           =   2385
      End
   End
   Begin VB.Frame fraMarketing 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marketing"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7404&
      Height          =   750
      Left            =   30
      TabIndex        =   35
      Top             =   4185
      Width           =   4950
      Begin VB.TextBox txtPorcentagemMarketingAPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3225
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   25
         Text            =   "0,00"
         Top             =   285
         Width           =   1560
      End
      Begin VB.TextBox txtPorcentagemMarketing 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "0,00"
         Top             =   285
         Width           =   1410
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "% Marketing"
         Height          =   240
         Index           =   14
         Left            =   105
         TabIndex        =   36
         Top             =   285
         Width           =   1230
      End
   End
   Begin VB.Frame fraTaxa 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Custos e Taxas Fixas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7404&
      Height          =   3120
      Left            =   5070
      TabIndex        =   34
      Top             =   1050
      Width           =   5445
      Begin VB.TextBox txtCustosFixos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3855
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   24
         Text            =   "0,00"
         Top             =   2685
         Width           =   1455
      End
      Begin VB.TextBox txtOutrosCustos 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   7
         Text            =   "0,00"
         Top             =   1731
         Width           =   1080
      End
      Begin VB.TextBox txtParcelamento 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   1254
         Width           =   1080
      End
      Begin VB.TextBox txtIOF 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   5
         Text            =   "0,00"
         Top             =   777
         Width           =   1080
      End
      Begin VB.TextBox txtImposto 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   2208
         Width           =   1080
      End
      Begin VB.TextBox txtImpostoAPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3855
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   23
         Text            =   "0,00"
         Top             =   2208
         Width           =   1455
      End
      Begin VB.TextBox txtIOFAPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3855
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   20
         Text            =   "0,00"
         Top             =   777
         Width           =   1455
      End
      Begin VB.TextBox txtParcelamentoAPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3855
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   21
         Text            =   "0,00"
         Top             =   1254
         Width           =   1455
      End
      Begin VB.TextBox txtOutrosCustosAPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3855
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   22
         Text            =   "0,00"
         Top             =   1731
         Width           =   1455
      End
      Begin VB.TextBox txtGatewayAPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3855
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   19
         Text            =   "0,00"
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox txtGateway 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Custos Fixos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   120
         TabIndex        =   47
         Top             =   2685
         Width           =   1995
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Outros Custos %"
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   1725
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Parcelamento %"
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   45
         Top             =   1260
         Width           =   1605
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "IOF %"
         Height          =   240
         Index           =   8
         Left            =   150
         TabIndex        =   44
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Imposto %"
         Height          =   240
         Index           =   10
         Left            =   105
         TabIndex        =   43
         Top             =   2205
         Width           =   1065
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gateway %"
         Height          =   240
         Index           =   9
         Left            =   120
         TabIndex        =   37
         Top             =   300
         Width           =   1140
      End
   End
   Begin VB.Frame fraProduto 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dados do Produto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7404&
      Height          =   1050
      Left            =   30
      TabIndex        =   31
      Top             =   0
      Width           =   10485
      Begin VB.ComboBox cmbProduto 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   165
         TabIndex        =   0
         Text            =   "cmdProduto"
         Top             =   525
         Width           =   8430
      End
      Begin VB.TextBox txtPrecoPraticado 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   8670
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "0,00"
         Top             =   525
         Width           =   1665
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Produto"
         Height          =   240
         Index           =   7
         Left            =   165
         TabIndex        =   33
         Top             =   255
         Width           =   765
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Preço Praticado"
         Height          =   240
         Index           =   6
         Left            =   8670
         TabIndex        =   32
         Top             =   255
         Width           =   1560
      End
   End
   Begin VB.Frame fraCusto 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Custo do Produto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7404&
      Height          =   3120
      Left            =   30
      TabIndex        =   27
      Top             =   1050
      Width           =   4950
      Begin VB.TextBox txtMarkup 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   18
         Text            =   "0,00"
         Top             =   2625
         Width           =   1710
      End
      Begin VB.TextBox txtPrecoFornecedor 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   3030
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "0,00"
         Top             =   285
         Width           =   1710
      End
      Begin VB.TextBox txtFreteFornecedor 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   3030
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "0,00"
         Top             =   870
         Width           =   1710
      End
      Begin VB.TextBox txtTotalCustoFornecedor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "0,00"
         Top             =   1455
         Width           =   1710
      End
      Begin VB.TextBox txtPrecoRecomendado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   17
         Text            =   "0,00"
         Top             =   2040
         Width           =   1710
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Markup"
         Height          =   240
         Index           =   12
         Left            =   135
         TabIndex        =   42
         Top             =   2625
         Width           =   705
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Preço Fornecedor"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   41
         Top             =   285
         Width           =   1740
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Preço Recomendado"
         Height          =   240
         Index           =   3
         Left            =   150
         TabIndex        =   30
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Custo Fornecedor"
         Height          =   240
         Index           =   2
         Left            =   165
         TabIndex        =   29
         Top             =   1455
         Width           =   2340
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frete Fornecedor"
         Height          =   240
         Index           =   1
         Left            =   165
         TabIndex        =   28
         Top             =   870
         Width           =   1710
      End
   End
   Begin MSComctlLib.ListView lvwLucro 
      Height          =   2145
      Left            =   15
      TabIndex        =   14
      Top             =   6480
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   3784
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwConsulta 
      Height          =   1920
      Left            =   0
      TabIndex        =   15
      Top             =   8685
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   3387
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmCustoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   Private RsConsulta               As New ADODB.Recordset
   Private Item                     As ListItem
   Private IL                       As Integer
   
   Private Const colProjeto         As Integer = 1
   Private Const colProduto         As Integer = 2
   Private Const colPrecoPraticado  As Integer = 3
   Private Const colPrecoFornecedor As Integer = 4
   Private Const colFreteFornecedor As Integer = 5
   
   Private Const colGateway         As Integer = 6
   Private Const colIOF             As Integer = 7
   Private Const colParcelamento    As Integer = 8
   Private Const colOutros_Custos   As Integer = 9
   Private Const colImposto         As Integer = 10
   Private Const colMarketing       As Integer = 11
   
   Private Const colTotalCusto      As Integer = 12
   Private Const colPrecoRecomend   As Integer = 13
   Private Const colStatus          As Integer = 14

Private Sub MontaLvwUnidades()
   lvwLucro.ListItems.Clear
   lvwLucro.ColumnHeaders.Clear
   lvwLucro.ColumnHeaders.Add , , "", 0
   lvwLucro.ColumnHeaders.Add , , "Qtde Unidades Vendidas", 7300, lvwColumnRight
   lvwLucro.ColumnHeaders.Add , , "Lucro", 2500, lvwColumnRight
   lvwLucro.View = lvwReport
End Sub

Private Sub MontaLvw()
   lvwConsulta.ListItems.Clear
   lvwConsulta.ColumnHeaders.Clear
   lvwConsulta.ColumnHeaders.Add , , "", 0
   lvwConsulta.ColumnHeaders.Add , , "Projeto", 0
   lvwConsulta.ColumnHeaders.Add , , "Produto", 2800
   lvwConsulta.ColumnHeaders.Add , , "Preço Praticado", 1600, lvwColumnRight
   lvwConsulta.ColumnHeaders.Add , , "Preço Fornecedor", 1700, lvwColumnRight
   lvwConsulta.ColumnHeaders.Add , , "Frete Fornecedor", 1700, lvwColumnRight
   
   lvwConsulta.ColumnHeaders.Add , , "Gateway", 0
   lvwConsulta.ColumnHeaders.Add , , "IOF", 0
   lvwConsulta.ColumnHeaders.Add , , "Parcelamento", 0
   lvwConsulta.ColumnHeaders.Add , , "Outros Custos", 0
   lvwConsulta.ColumnHeaders.Add , , "Imposto", 0
   lvwConsulta.ColumnHeaders.Add , , "Marketing", 0
   
   lvwConsulta.ColumnHeaders.Add , , "Total Custo", 1200, lvwColumnRight
   lvwConsulta.ColumnHeaders.Add , , "Preço Recomendado", 2000, lvwColumnRight
   lvwConsulta.ColumnHeaders.Add , , "", 350
   lvwConsulta.View = lvwReport
End Sub

Private Sub cmbProduto_KeyPress(KeyAscii As Integer)

   Dim CB As Long
   Dim FindString As String
   Const CB_ERR = (-1)
   Const CB_FINDSTRING = &H14C
  
   With cmbProduto
      If KeyAscii = 68 Or KeyAscii = 100 Then cmbProduto.SetFocus
      
      If KeyAscii = 13 Then
         If KeyAscii = vbKeyReturn Then
            Sendkeys "{tab}"
         End If
      End If
         
      If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
      
      If .SelLength = 0 Then
         FindString = .Text & Chr$(KeyAscii)
      Else
         FindString = Left$(.Text, .SelStart) & Chr$(KeyAscii)
      End If
      
      CB = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal FindString)
      If CB <> CB_ERR Then
         .ListIndex = CB
         .SelStart = Len(FindString)
         .SelLength = Len(.Text) - .SelStart
      End If
   End With
   KeyAscii = 0
   
End Sub

Private Sub cmbProduto_Validate(Cancel As Boolean)
   If Len(cmbProduto) > 75 Then cmbProduto = Mid(cmbProduto, 1, 75)
End Sub

Private Sub cmdCancelar_Click()

   MontaLvw
   MontaLvwUnidades
   
   CARREGA_COMBOS

   cmbProduto.Enabled = True
   If cmbProduto.ListCount > 0 Then cmbProduto.ListIndex = 0
   txtPrecoPraticado = "0,00"
   
   txtPrecoFornecedor = "0,00"
   txtFreteFornecedor = "0,00"
   txtTotalCustoFornecedor = "0,00"
   txtPrecoRecomendado = "0,00"
   txtMarkup = "0,00"
   
   txtGateway = "0,00"
   txtGatewayAPagar = "0,00"
   txtIOF = "0,00"
   txtIOFAPagar = "0,00"
   txtParcelamento = "0,00"
   txtParcelamentoAPagar = "0,00"
   txtOutrosCustos = "0,00"
   txtOutrosCustosAPagar = "0,00"
   txtImposto = "0,00"
   txtImpostoAPagar = "0,00"
   txtCustosFixos = "0,00"
   
   txtPorcentagemMarketing = "0,00"
   txtPorcentagemMarketingAPagar = "0,00"
   
   txtTotalCustoOperacional = "0,00"
   
   txtLucroFinal = "0,00"
   txtLucroFinalPorcentagem = "0,00"
   txtMaximoCustoPorAquisicao = "0,00"
      
   CONSULTAR_REGISTROS
   PREENCHER_CUSTOS_FIXOS
   
End Sub

Private Sub cmdConsultar_Click()
   CONSULTAR_REGISTROS
End Sub

Private Sub cmdExcluir_Click()

 If lvwConsulta.ListItems.Count = 0 Then Exit Sub
   
   If MsgBox("Deseja Realmente Excluir Dado Selecionado?", vbYesNo + vbQuestion + vbDefaultButton2, "Exclusão") = vbNo Then
      Exit Sub
   End If
   
   cmdExcluir.Enabled = False
    
   With lvwConsulta
      For IL = 1 To .ListItems.Count
         Select Case .ListItems(IL).SubItems(colStatus)
         Case "D"
            CmdSql = "DELETE FROM CUSTO_PRODUTO" & vbCr
            CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName)) & vbCr
            CmdSql = CmdSql & "  AND PRODUTO = " & PoeAspas(.ListItems(IL).SubItems(colProduto))
            CMySql.Executa CmdSql, True
         End Select
      Next IL
   End With
   
Cancelar:
   cmdExcluir.Enabled = True
      
   Form_Load
   
   If cmbProduto.ListCount > 0 Then cmbProduto.ListIndex = 0
   cmbProduto.SetFocus

End Sub

Private Sub cmdSalvar_Click()

   If Trim(cmbProduto) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O PRODUTO"
      cmbProduto.SetFocus
      Exit Sub
   End If

   If Val(txtPrecoPraticado) = 0 Then
      MsgBoxTabum Me, "FAVOR PREENCHER O PREÇO PRATICADO"
      txtPrecoPraticado.SetFocus
      Exit Sub
   End If
   
   If Val(txtPrecoFornecedor) = 0 Then
      MsgBoxTabum Me, "FAVOR PREENCHER O PREÇO DO FORNECEDOR"
      txtPrecoFornecedor.SetFocus
      Exit Sub
   End If
   
   If Val(txtFreteFornecedor) = 0 Then
      MsgBoxTabum Me, "FAVOR PREENCHER O PREÇO DO FRETE"
      txtFreteFornecedor.SetFocus
      Exit Sub
   End If
               
   cmdSalvar.Enabled = False
           
   CmdSql = "SELECT * FROM CUSTO_PRODUTO" & vbCr
   CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName)) & vbCr
   CmdSql = CmdSql & "  AND PRODUTO = " & PoeAspas(cmbProduto) & vbCr
   CmdSql = CmdSql & "ORDER BY PROJETO,PRODUTO"
   CMySql.Consulta CmdSql, RsConsulta
      
   If RsConsulta.EOF Then
      CmdSql = "INSERT INTO CUSTO_PRODUTO(PROJETO,PRODUTO,PRECO_PRATICADO,PRECO_FORNECEDOR,FRETE_FORNECEDOR,"
      CmdSql = CmdSql & "GATEWAY,IOF,PARCELAMENTO,OUTROS_CUSTOS,IMPOSTO,MARKETING)" & vbCr
      CmdSql = CmdSql & "VALUES("
      CmdSql = CmdSql & PoeAspas(UCase(App.EXEName)) & ","
      CmdSql = CmdSql & PoeAspas(UCase(cmbProduto)) & ","
      CmdSql = CmdSql & Str(txtPrecoPraticado) & ","
      CmdSql = CmdSql & Str(txtPrecoFornecedor) & ","
      CmdSql = CmdSql & Str(txtFreteFornecedor) & ","
      CmdSql = CmdSql & Str(txtGateway) & ","
      CmdSql = CmdSql & Str(txtIOF) & ","
      CmdSql = CmdSql & Str(txtParcelamento) & ","
      CmdSql = CmdSql & Str(txtOutrosCustos) & ","
      CmdSql = CmdSql & Str(txtImposto) & ","
      CmdSql = CmdSql & Str(txtPorcentagemMarketing) & ")"
      CMySql.Executa CmdSql, True
   
      MsgBoxTabum Me, "REGISTRO INCLUÍDO COM SUCESSO"
   Else
      
      CmdSql = "UPDATE CUSTO_PRODUTO SET PRECO_PRATICADO  = " & Str(txtPrecoPraticado) & "," & vbCr
      CmdSql = CmdSql & "                         PRECO_FORNECEDOR = " & Str(txtPrecoFornecedor) & "," & vbCr
      CmdSql = CmdSql & "                         FRETE_FORNECEDOR = " & Str(txtFreteFornecedor) & "," & vbCr
      CmdSql = CmdSql & "                         GATEWAY          = " & Str(txtGateway) & "," & vbCr
      CmdSql = CmdSql & "                         IOF              = " & Str(txtIOF) & "," & vbCr
      CmdSql = CmdSql & "                         PARCELAMENTO     = " & Str(txtParcelamento) & "," & vbCr
      CmdSql = CmdSql & "                         OUTROS_CUSTOS    = " & Str(txtOutrosCustos) & "," & vbCr
      CmdSql = CmdSql & "                         IMPOSTO          = " & Str(txtImposto) & "," & vbCr
      CmdSql = CmdSql & "                         MARKETING        = " & Str(txtPorcentagemMarketing) & vbCr
      CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName)) & vbCr
      CmdSql = CmdSql & "  AND PRODUTO = " & PoeAspas(UCase(cmbProduto))
      CMySql.Executa CmdSql, True
      
      MsgBoxTabum Me, "REGISTRO ALTERADO COM SUCESSO"
   End If
         
   cmdCancelar_Click
   cmdSalvar.Enabled = True
   cmbProduto.Enabled = True
   cmbProduto.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub Form_Load()
   CARREGA_COMBOS
   
   If cmbProduto.ListCount > 0 Then cmbProduto.ListIndex = 0
   
   PREENCHER_CUSTOS_FIXOS
   
   CONSULTAR_REGISTROS
End Sub

Private Sub lvwConsulta_DblClick()
   
   If lvwConsulta.ListItems.Count = 0 Then Exit Sub
   
   If cmbProduto.ListCount > 0 Then cmbProduto.ListIndex = 0
   txtPrecoPraticado = "0,00"
   
   txtPrecoFornecedor = "0,00"
   txtFreteFornecedor = "0,00"
   txtTotalCustoFornecedor = "0,00"
   txtPrecoRecomendado = "0,00"
   
   txtGateway = "0,00"
   txtGatewayAPagar = "0,00"
   txtIOF = "0,00"
   txtIOFAPagar = "0,00"
   txtParcelamento = "0,00"
   txtParcelamentoAPagar = "0,00"
   txtOutrosCustos = "0,00"
   txtOutrosCustosAPagar = "0,00"
   txtImposto = "0,00"
   txtImpostoAPagar = "0,00"
   
   txtPorcentagemMarketing = "0,00"
   txtPorcentagemMarketingAPagar = "0,00"
   
   txtTotalCustoOperacional = "0,00"
         
   cmbProduto = UCase(lvwConsulta.SelectedItem.SubItems(colProduto))
   cmbProduto.Enabled = False
         
   txtPrecoPraticado = Format(lvwConsulta.SelectedItem.SubItems(colPrecoPraticado), "###,##0.00")
   txtPrecoFornecedor = Format(lvwConsulta.SelectedItem.SubItems(colPrecoFornecedor), "###,##0.00")
   txtFreteFornecedor = Format(lvwConsulta.SelectedItem.SubItems(colFreteFornecedor), "###,##0.00")
   txtTotalCustoFornecedor = Format(lvwConsulta.SelectedItem.SubItems(colTotalCusto), "###,##0.00")
   txtPrecoRecomendado = Format(lvwConsulta.SelectedItem.SubItems(colPrecoRecomend), "###,##0.00")
   
   txtGateway = Format(lvwConsulta.SelectedItem.SubItems(colGateway), "###,##0.00")
   txtIOF = Format(lvwConsulta.SelectedItem.SubItems(colIOF), "###,##0.00")
   txtParcelamento = Format(lvwConsulta.SelectedItem.SubItems(colParcelamento), "###,##0.00")
   txtOutrosCustos = Format(lvwConsulta.SelectedItem.SubItems(colOutros_Custos), "###,##0.00")
   txtImposto = Format(lvwConsulta.SelectedItem.SubItems(colImposto), "###,##0.00")
   
   txtPorcentagemMarketing = Format(lvwConsulta.SelectedItem.SubItems(colMarketing), "###,##0.00")
         
   txtPrecoPraticado.SetFocus
   
   CALCULA_CUSTOS

End Sub

Private Sub lvwConsulta_KeyDown(KeyCode As Integer, Shift As Integer)

   If lvwConsulta.ListItems.Count = 0 Then Exit Sub
   
   Select Case KeyCode
   Case "46"
      Select Case lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(colStatus)
      Case ""
         lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(colStatus) = "D"
      Case "D"
         lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(colStatus) = ""
      End Select
   End Select
   
End Sub

Private Sub txtFreteFornecedor_GotFocus()
   Marca Me
End Sub

Private Sub txtFreteFornecedor_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtFreteFornecedor, KeyAscii
End Sub

Private Sub txtFreteFornecedor_LostFocus()
   If Not IsNumeric(txtFreteFornecedor.Text) Then txtFreteFornecedor.Text = 0
   txtFreteFornecedor.Text = Format((txtFreteFornecedor.Text), "###,##0.00")
   
   CALCULA_CUSTOS
End Sub

Private Sub txtGateway_GotFocus()
   Marca Me
End Sub

Private Sub txtGateway_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtGateway, KeyAscii
End Sub

Private Sub txtGateway_LostFocus()
   If Not IsNumeric(txtGateway.Text) Then txtGateway.Text = 0
   txtGateway.Text = Format((txtGateway.Text), "###,##0.00")
   
   CALCULA_CUSTOS
End Sub

Private Sub txtImposto_GotFocus()
   Marca Me
End Sub

Private Sub txtImposto_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtImposto, KeyAscii
End Sub

Private Sub txtImposto_LostFocus()
   If Not IsNumeric(txtImposto.Text) Then txtImposto.Text = 0
   txtImposto.Text = Format((txtImposto.Text), "###,##0.00")
   
   CALCULA_CUSTOS
End Sub

Private Sub txtIOF_GotFocus()
   Marca Me
End Sub

Private Sub txtIOF_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtIOF, KeyAscii
End Sub

Private Sub txtIOF_LostFocus()
   If Not IsNumeric(txtIOF.Text) Then txtIOF.Text = 0
   txtIOF.Text = Format((txtIOF.Text), "###,##0.00")
   
   CALCULA_CUSTOS
End Sub

Private Sub txtOutrosCustos_GotFocus()
   Marca Me
End Sub

Private Sub txtOutrosCustos_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtOutrosCustos, KeyAscii
End Sub

Private Sub txtOutrosCustos_LostFocus()
   If Not IsNumeric(txtOutrosCustos.Text) Then txtOutrosCustos.Text = 0
   txtOutrosCustos.Text = Format((txtOutrosCustos.Text), "###,##0.00")
   
   CALCULA_CUSTOS
End Sub

Private Sub txtParcelamento_GotFocus()
   Marca Me
End Sub

Private Sub txtParcelamento_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtParcelamento, KeyAscii
End Sub

Private Sub txtParcelamento_LostFocus()
   If Not IsNumeric(txtParcelamento.Text) Then txtParcelamento.Text = 0
   txtParcelamento.Text = Format((txtParcelamento.Text), "###,##0.00")
   
   CALCULA_CUSTOS
End Sub

Private Sub txtPorcentagemMarketing_GotFocus()
   Marca Me
End Sub

Private Sub txtPorcentagemMarketing_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtPorcentagemMarketing, KeyAscii
End Sub

Private Sub txtPorcentagemMarketing_LostFocus()
   If Not IsNumeric(txtPorcentagemMarketing.Text) Then txtPorcentagemMarketing.Text = 0
   txtPorcentagemMarketing.Text = Format((txtPorcentagemMarketing.Text), "###,##0.00")
   
   CALCULA_CUSTOS
End Sub

Private Sub txtPrecoFornecedor_GotFocus()
   Marca Me
End Sub

Private Sub txtPrecoFornecedor_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtPrecoFornecedor, KeyAscii
End Sub

Private Sub txtPrecoFornecedor_LostFocus()
   If Not IsNumeric(txtPrecoFornecedor.Text) Then txtPrecoFornecedor.Text = 0
   txtPrecoFornecedor.Text = Format((txtPrecoFornecedor.Text), "###,##0.00")
   
   CALCULA_CUSTOS
End Sub

Private Sub txtPrecoPraticado_GotFocus()
   Marca Me
End Sub

Private Sub txtPrecoPraticado_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtPrecoPraticado, KeyAscii
End Sub


Private Sub txtPrecoPraticado_LostFocus()
   If Not IsNumeric(txtPrecoPraticado.Text) Then txtPrecoPraticado.Text = 0
   txtPrecoPraticado.Text = Format((txtPrecoPraticado.Text), "###,##0.00")
   
   CALCULA_CUSTOS
End Sub

Private Sub CONSULTAR_REGISTROS()

   MontaLvw
   MontaLvwUnidades
   
   cmbProduto.Enabled = True
       
   CmdSql = "SELECT *" & vbCr
   CmdSql = CmdSql & "FROM CUSTO_PRODUTO" & vbCr
   CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName)) & vbCr
   If Trim(cmbProduto) <> "" Then CmdSql = CmdSql & "  AND PRODUTO LIKE " & PoeAspas("%" & Trim(cmbProduto) & "%") & vbCr
   CmdSql = CmdSql & "ORDER BY PROJETO,PRODUTO"
   CMySql.Consulta CmdSql, RsConsulta
      
   Do While Not RsConsulta.EOF
      Set Item = lvwConsulta.ListItems.Add(, , "")
      Item.SubItems(colProjeto) = Trim(RsConsulta("PROJETO"))
      Item.SubItems(colProduto) = Trim(RsConsulta("PRODUTO"))
      Item.SubItems(colPrecoPraticado) = Format(RsConsulta("PRECO_PRATICADO"), "###,##0.00")
      Item.SubItems(colPrecoFornecedor) = Format(RsConsulta("PRECO_FORNECEDOR"), "###,##0.00")
      Item.SubItems(colFreteFornecedor) = Format(RsConsulta("FRETE_FORNECEDOR"), "###,##0.00")
      Item.SubItems(colGateway) = Format(RsConsulta("GATEWAY"), "###,##0.00")
      Item.SubItems(colIOF) = Format(RsConsulta("IOF"), "###,##0.00")
      Item.SubItems(colParcelamento) = Format(RsConsulta("PARCELAMENTO"), "###,##0.00")
      Item.SubItems(colOutros_Custos) = Format(RsConsulta("OUTROS_CUSTOS"), "###,##0.00")
      Item.SubItems(colImposto) = Format(RsConsulta("IMPOSTO"), "###,##0.00")
      Item.SubItems(colMarketing) = Format(RsConsulta("MARKETING"), "###,##0.00")
      Item.SubItems(colTotalCusto) = Format((RsConsulta("PRECO_FORNECEDOR")) + Trim(RsConsulta("FRETE_FORNECEDOR")), "###,##0.00")
      Item.SubItems(colPrecoRecomend) = Format((RsConsulta("MARKETING") + RsConsulta("FRETE_FORNECEDOR")) * 3, "###,##0.00")
      txtTotalCustoFornecedor = Format(Cdblx(txtPrecoFornecedor) + Cdblx(txtFreteFornecedor), "###,##0.00")
      txtPrecoRecomendado = Format(Cdblx(txtTotalCustoFornecedor) * 3, "###,##0.00")
      Item.SubItems(colStatus) = ""
   RsConsulta.MoveNext
   Loop

End Sub

Private Function CALCULA_CUSTOS()
   txtTotalCustoFornecedor = Format(Cdblx(txtPrecoFornecedor) + Cdblx(txtFreteFornecedor), "###,##0.00")
   txtPrecoRecomendado = Format(Cdblx(txtTotalCustoFornecedor) * 3, "###,##0.00")
   
   If Cdblx(txtTotalCustoFornecedor) > 0 Then txtMarkup = Format(Cdblx(txtPrecoPraticado) / Cdblx(txtTotalCustoFornecedor), "###,##0.00")
   
   txtGatewayAPagar = Format((Cdblx(txtPrecoPraticado) * Cdblx(txtGateway)) / 100, "###,##0.00")
   txtIOFAPagar = Format((Cdblx(txtTotalCustoFornecedor) * Cdblx(txtIOF)) / 100, "###,##0.00")
   txtParcelamentoAPagar = Format((Cdblx(txtPrecoPraticado) * Cdblx(txtParcelamento)) / 100, "###,##0.00")
   txtOutrosCustosAPagar = Format((Cdblx(txtPrecoPraticado) * Cdblx(txtOutrosCustos)) / 100, "###,##0.00")
   txtImpostoAPagar = Format((Cdblx(txtPrecoPraticado) * Cdblx(txtImposto)) / 100, "###,##0.00")
   txtCustosFixos = Format(Cdblx(txtGatewayAPagar) + Cdblx(txtIOFAPagar) + Cdblx(txtParcelamentoAPagar) + Cdblx(txtOutrosCustosAPagar) + Cdblx(txtImpostoAPagar), "###,##0.00")
   
   txtPorcentagemMarketingAPagar = Format((Cdblx(txtPorcentagemMarketing) * Cdblx(txtPrecoPraticado)) / 100, "###,##0.00")
   
   txtTotalCustoOperacional = Format(Cdblx(txtCustosFixos) + Cdblx(txtTotalCustoFornecedor) + Cdblx(txtPorcentagemMarketingAPagar), "###,##0.00")
   
   txtLucroFinal = Format(Cdblx(txtPrecoPraticado) - Cdblx(txtTotalCustoOperacional), "###,##0.00")
   
   If Cdblx(txtPrecoPraticado) > 0 Then txtLucroFinalPorcentagem = Format((Cdblx(txtLucroFinal) / Cdblx(txtPrecoPraticado)) * 100, "###,##0.00") & "%"
   
   txtMaximoCustoPorAquisicao = Format(Cdblx(txtLucroFinal) + Cdblx(txtPorcentagemMarketingAPagar), "###,##0.00")
      
   MontaLvwUnidades
   
   If Cdblx(txtLucroFinal) > 0 Then
      Set Item = lvwLucro.ListItems.Add(, , "")
      Item.SubItems(1) = 10
      Item.SubItems(2) = "R$ " & Format(Cdblx(txtLucroFinal) * 10, "###,##0.00")
      
      Set Item = lvwLucro.ListItems.Add(, , "")
      Item.SubItems(1) = 25
      Item.SubItems(2) = "R$ " & Format(Cdblx(txtLucroFinal) * 25, "###,##0.00")
      
      Set Item = lvwLucro.ListItems.Add(, , "")
      Item.SubItems(1) = 50
      Item.SubItems(2) = "R$ " & Format(Cdblx(txtLucroFinal) * 50, "###,##0.00")
      
      Set Item = lvwLucro.ListItems.Add(, , "")
      Item.SubItems(1) = 100
      Item.SubItems(2) = "R$ " & Format(Cdblx(txtLucroFinal) * 100, "###,##0.00")
      
      Set Item = lvwLucro.ListItems.Add(, , "")
      Item.SubItems(1) = 250
      Item.SubItems(2) = "R$ " & Format(Cdblx(txtLucroFinal) * 250, "###,##0.00")
      
      Set Item = lvwLucro.ListItems.Add(, , "")
      Item.SubItems(1) = 500
      Item.SubItems(2) = "R$ " & Format(Cdblx(txtLucroFinal) * 500, "###,##0.00")
      
      Set Item = lvwLucro.ListItems.Add(, , "")
      Item.SubItems(1) = 1000
      Item.SubItems(2) = "R$ " & Format(Cdblx(txtLucroFinal) * 1000, "###,##0.00")
   End If
   
End Function

Private Sub PREENCHER_CUSTOS_FIXOS()

   If cmbProduto.Enabled = True Then
      CmdSql = "SELECT *" & vbCr
      CmdSql = CmdSql & "FROM TIPOCOMBO" & vbCr
      CmdSql = CmdSql & "WHERE PROJETO = 'MTABUM'" & vbCr
      CmdSql = CmdSql & "  AND FORMULARIO = 'CUSTO DO PRODUTO'"
      CMySql.Consulta CmdSql, RsConsulta
      
      Do While Not RsConsulta.EOF
         Select Case Trim(RsConsulta("TIPO"))
         Case "GATEWAY"
            txtGateway = Format(Cdblx(RsConsulta("DESCRICAO")), "###,##0.00")
         Case "IOF"
            txtIOF = Format(Cdblx(RsConsulta("DESCRICAO")), "###,##0.00")
         Case "PARCELAMENTO"
            txtParcelamento = Format(Cdblx(RsConsulta("DESCRICAO")), "###,##0.00")
         Case "OUTROS CUSTOS"
            txtOutrosCustos = Format(Cdblx(RsConsulta("DESCRICAO")), "###,##0.00")
         Case "IMPOSTO"
            txtImposto = Format(Cdblx(RsConsulta("DESCRICAO")), "###,##0.00")
         Case "MARKETING"
            txtPorcentagemMarketing = Format(Cdblx(RsConsulta("DESCRICAO")), "###,##0.00")
         End Select
      RsConsulta.MoveNext
      Loop
   End If
   
End Sub

Private Sub CARREGA_COMBOS()

   CmdSql = "SELECT PRODUTO" & vbCr
   CmdSql = CmdSql & "FROM CUSTO_PRODUTO" & vbCr
   CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName))
   CMySql.Consulta CmdSql, RsConsulta
   
   cmbProduto.Clear
   cmbProduto.AddItem ""
   
   Do While Not RsConsulta.EOF
      cmbProduto.AddItem Trim(RsConsulta("PRODUTO"))
   RsConsulta.MoveNext
   Loop
End Sub
