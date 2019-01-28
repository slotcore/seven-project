VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPedido2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas - Ingreso de Pedidos"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fraseldoc 
      BorderStyle     =   0  'None
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2280
      Left            =   12180
      TabIndex        =   53
      Top             =   1080
      Visible         =   0   'False
      Width           =   5565
      Begin VB.CommandButton CmdBusTipDocGen 
         Height          =   240
         Left            =   5160
         Picture         =   "FrmPedido2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   840
         Width           =   240
      End
      Begin VB.CommandButton CmdBusSerGen 
         Height          =   240
         Left            =   2490
         Picture         =   "FrmPedido2.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1485
         Width           =   240
      End
      Begin VB.TextBox TxtNumDocGen 
         Height          =   300
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   57
         Text            =   "TxtNumDocGen"
         Top             =   1755
         Width           =   1335
      End
      Begin VB.CommandButton CmdBusAlmacen2 
         Height          =   240
         Left            =   3180
         Picture         =   "FrmPedido2.frx":0264
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   525
         Width           =   240
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmiAnul 
         Height          =   300
         Left            =   1425
         TabIndex        =   56
         Top             =   1125
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin VB.TextBox TxtAlmacen2 
         Height          =   300
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "TxtAlmacen2"
         Top             =   495
         Width           =   2025
      End
      Begin VB.TextBox TxtIdDocGen 
         Height          =   300
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "TxtIdDocGen"
         Top             =   810
         Width           =   4005
      End
      Begin VB.TextBox TxtNumSerGen 
         Height          =   300
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "TxtNumSerGen"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Frame Frame7 
         Height          =   1020
         Left            =   3030
         TabIndex        =   62
         Top             =   1065
         Width           =   2400
         Begin VB.CommandButton cmdokseldoc 
            Height          =   600
            Left            =   450
            Picture         =   "FrmPedido2.frx":0396
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   270
            Width           =   720
         End
         Begin VB.CommandButton cmdsalirseldoc 
            Height          =   600
            Left            =   1200
            Picture         =   "FrmPedido2.frx":06A0
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   270
            Width           =   720
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   72
         Top             =   840
         Width           =   1185
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   5550
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   2235
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   5550
         X2              =   5550
         Y1              =   15
         Y2              =   2220
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emisión de Documentos Anulados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   135
         TabIndex        =   71
         Top             =   105
         Width           =   2880
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nº Serie"
         Height          =   195
         Left            =   165
         TabIndex        =   70
         Top             =   1470
         Width           =   585
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   5565
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Label LblIdDocumentoGen 
         AutoSize        =   -1  'True
         Caption         =   "LblIdDocumentoGen"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3765
         TabIndex        =   69
         Top             =   390
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   165
         TabIndex        =   68
         Top             =   1785
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Documento"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   67
         Top             =   1170
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   66
         Top             =   510
         Width           =   615
      End
      Begin VB.Label LblidAlmacen2 
         AutoSize        =   -1  'True
         Caption         =   "LblidAlmacen2"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3765
         TabIndex        =   65
         Top             =   585
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Left            =   45
         Top             =   45
         Width           =   5475
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7755
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido2.frx":09AA
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido2.frx":0EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido2.frx":1280
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido2.frx":1404
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido2.frx":1858
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido2.frx":1970
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido2.frx":1EB4
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido2.frx":23F8
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido2.frx":250C
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido2.frx":2620
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido2.frx":2A74
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido2.frx":2BE0
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido2.frx":3128
            Key             =   "IMG12"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7470
      Left            =   -15
      TabIndex        =   16
      Top             =   360
      Width           =   11895
      _cx             =   20981
      _cy             =   13176
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   8388608
      Caption         =   "  &Consulta  |   &Detalle  "
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   0
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7050
         Left            =   12540
         TabIndex        =   21
         Top             =   375
         Width           =   11805
         Begin VB.Frame FrmModFecha 
            Caption         =   "[ Modificar Fecha de Entrega ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   3645
            Left            =   480
            TabIndex        =   74
            Top             =   2760
            Visible         =   0   'False
            Width           =   5985
            Begin VB.CommandButton CmdAceptar 
               Caption         =   "&Aceptar"
               Height          =   420
               Left            =   1650
               TabIndex        =   77
               Top             =   3120
               Width           =   1170
            End
            Begin VB.CommandButton CmdCancelar 
               Caption         =   "&Cancelar"
               Height          =   420
               Left            =   2880
               TabIndex        =   76
               Top             =   3120
               Width           =   1170
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg3 
               Height          =   1965
               Left            =   90
               TabIndex        =   75
               Top             =   1080
               Width           =   5790
               _cx             =   10213
               _cy             =   3466
               _ConvInfo       =   1
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   128
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   7
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmPedido2.frx":3442
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
            End
            Begin VB.Label LblIdpeddet 
               AutoSize        =   -1  'True
               BackColor       =   &H000000C0&
               BackStyle       =   0  'Transparent
               Caption         =   "LblIdpeddet"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   4830
               TabIndex        =   84
               Top             =   150
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Cantidad"
               Height          =   195
               Index           =   10
               Left            =   3200
               TabIndex        =   83
               Top             =   735
               Width           =   630
            End
            Begin VB.Label LblCantidad 
               BackColor       =   &H80000009&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblCantidad"
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   4230
               TabIndex        =   82
               Top             =   690
               Width           =   1425
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Fecha"
               Height          =   195
               Index           =   9
               Left            =   90
               TabIndex        =   81
               Top             =   735
               Width           =   450
            End
            Begin VB.Label LblFecha 
               BackColor       =   &H80000009&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblFecha"
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   1000
               TabIndex        =   80
               Top             =   690
               Width           =   1425
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Producto"
               Height          =   195
               Index           =   8
               Left            =   90
               TabIndex        =   79
               Top             =   390
               Width           =   645
            End
            Begin VB.Label LblProducto 
               BackColor       =   &H80000009&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblProducto"
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   1005
               TabIndex        =   78
               Top             =   360
               Width           =   4845
            End
         End
         Begin VB.TextBox TxtOC 
            Height          =   300
            Left            =   6150
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   7
            Text            =   "TxtOC"
            Top             =   1800
            Width           =   3330
         End
         Begin VB.CommandButton CmdBusCondicion 
            Height          =   240
            Left            =   2205
            Picture         =   "FrmPedido2.frx":34BA
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   2160
            Width           =   240
         End
         Begin VB.CommandButton CmdBusPtoVta 
            Height          =   240
            Left            =   2205
            Picture         =   "FrmPedido2.frx":35EC
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1170
            Width           =   240
         End
         Begin VB.CommandButton CmdBusAlm 
            Height          =   240
            Left            =   6570
            Picture         =   "FrmPedido2.frx":371E
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   480
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.TextBox TxtIdAlm 
            Height          =   300
            Left            =   6150
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "TxtIdAlm"
            Top             =   450
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtglosa 
            Height          =   315
            Left            =   1560
            TabIndex        =   11
            Text            =   "TxtGlosa"
            Top             =   2460
            Width           =   10185
         End
         Begin VB.Frame Frame6 
            Height          =   3495
            Left            =   7845
            TabIndex        =   30
            Top             =   2820
            Width           =   3915
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   3105
               Left            =   60
               TabIndex        =   15
               Top             =   330
               Width           =   3720
               _cx             =   6562
               _cy             =   5477
               _ConvInfo       =   1
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   128
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   12
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmPedido2.frx":3850
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Entrega"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   60
               TabIndex        =   31
               Top             =   120
               Width           =   1530
            End
         End
         Begin VB.Frame Frame10 
            Height          =   765
            Left            =   9600
            TabIndex        =   32
            Top             =   360
            Width           =   2115
            Begin VB.Label LblPeriodo2 
               Alignment       =   2  'Center
               Caption         =   "LblPeriodo2"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   150
               TabIndex        =   33
               Top             =   300
               Width           =   1860
            End
         End
         Begin VB.CommandButton CmdBusNumSer 
            Height          =   240
            Left            =   2205
            Picture         =   "FrmPedido2.frx":3917
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1830
            Width           =   240
         End
         Begin VB.CommandButton CmdBusCli 
            Height          =   240
            Left            =   3045
            Picture         =   "FrmPedido2.frx":3A49
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   825
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   2205
            Picture         =   "FrmPedido2.frx":3B7B
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1500
            Width           =   240
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2670
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "TxtNumDoc"
            Top             =   1800
            Width           =   1440
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   5
            Text            =   "TxtNumSer"
            Top             =   1800
            Width           =   915
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   4
            Text            =   "TxtTipDoc"
            Top             =   1470
            Width           =   915
         End
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   2
            Text            =   "TxtNumRuc"
            Top             =   795
            Width           =   1770
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchDoc 
            Height          =   300
            Left            =   1560
            TabIndex        =   0
            Top             =   450
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Valor           =   "03/01/2004"
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   3375
            Left            =   45
            TabIndex        =   14
            Top             =   2925
            Width           =   7650
            _cx             =   13494
            _cy             =   5953
            _ConvInfo       =   1
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   128
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   13
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmPedido2.frx":3CAD
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
         End
         Begin VB.Frame Frame4 
            Height          =   780
            Left            =   60
            TabIndex        =   22
            Top             =   6255
            Width           =   11700
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "&Eliminar Item"
               Enabled         =   0   'False
               Height          =   420
               Left            =   1245
               TabIndex        =   13
               Top             =   210
               Width           =   1170
            End
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "&Agregar Item"
               Enabled         =   0   'False
               Height          =   420
               Left            =   45
               TabIndex        =   12
               Top             =   210
               Width           =   1170
            End
         End
         Begin VB.TextBox TxtPtoVta 
            Height          =   300
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   3
            Text            =   "TxtPtoVta"
            Top             =   1140
            Width           =   915
         End
         Begin VB.TextBox TxtConPag 
            Height          =   300
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   8
            Text            =   "TxtConPag"
            Top             =   2130
            Width           =   915
         End
         Begin VB.CommandButton CmdTipPed 
            Height          =   240
            Left            =   6570
            Picture         =   "FrmPedido2.frx":3D8A
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   2160
            Width           =   240
         End
         Begin VB.TextBox TxtTipPed 
            Height          =   300
            Left            =   6150
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   9
            Text            =   "TxtTipPed"
            Top             =   2130
            Width           =   705
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEnt 
            Height          =   300
            Left            =   9780
            TabIndex        =   10
            Top             =   2130
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Valor           =   "03/01/2004"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Orden de Compra"
            Height          =   195
            Left            =   4875
            TabIndex        =   73
            Top             =   1875
            Width           =   1245
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Ent."
            Height          =   195
            Index           =   3
            Left            =   9060
            TabIndex        =   52
            ToolTipText     =   "Fecha de Vencimiento"
            Top             =   2205
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label LblCondPag 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblCondPag"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2505
            TabIndex        =   48
            Top             =   2130
            Width           =   2655
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Condición de Pago"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   47
            Top             =   2205
            Width           =   1350
         End
         Begin VB.Label LblPtoVta 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblPtoVta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2505
            TabIndex        =   45
            Top             =   1140
            Width           =   6975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Punto Venta"
            Height          =   195
            Index           =   6
            Left            =   90
            TabIndex        =   44
            Top             =   1209
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   195
            Index           =   11
            Left            =   5505
            TabIndex        =   41
            Top             =   525
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label LblAlmacen 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblAlmacen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   6855
            TabIndex        =   40
            Top             =   450
            Visible         =   0   'False
            Width           =   2625
         End
         Begin VB.Label lblglosa 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Left            =   90
            TabIndex        =   38
            Top             =   2580
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Documento"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   37
            Top             =   1893
            Width           =   1275
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   36
            Top             =   1551
            Width           =   1185
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   35
            Top             =   867
            Width           =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emisión"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   34
            Top             =   525
            Width           =   1260
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Pedido"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   105
            TabIndex        =   29
            Top             =   45
            Width           =   11595
         End
         Begin VB.Label LblNomCli 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomCli"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   3345
            TabIndex        =   28
            Top             =   795
            Width           =   6135
         End
         Begin VB.Label LblNomDoc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomDoc"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2505
            TabIndex        =   27
            Top             =   1470
            Width           =   3615
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2520
            Top             =   1905
            Width           =   105
         End
         Begin VB.Label LblIdCliente 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCliente"
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   2865
            TabIndex        =   26
            Top             =   510
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label LblTipPed 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipPed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   6855
            TabIndex        =   51
            Top             =   2130
            Width           =   2085
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Pedido"
            Height          =   195
            Index           =   5
            Left            =   5265
            TabIndex        =   50
            Top             =   2205
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7050
         Left            =   45
         TabIndex        =   18
         Top             =   375
         Width           =   11805
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6645
            Left            =   0
            TabIndex        =   42
            Top             =   360
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11721
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "TD"
            Columns(1).DataField=   "abrev"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numerodoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Emi"
            Columns(3).DataField=   "fchemi1"
            Columns(3).NumberFormat=   "Short Date"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Tipo"
            Columns(4).DataField=   "tipo"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Cliente"
            Columns(5).DataField=   "nombre"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Cond. Pago"
            Columns(6).DataField=   "condpago"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Nº Orden Compra"
            Columns(7).DataField=   "oc"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=794"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=714"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2566"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2487"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1773"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1693"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=4815"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=4736"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=4233"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=4154"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   0
            DefColWidth     =   0
            HeadLines       =   1.5
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0FEFE&,.fgcolor=&H0&,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Named:id=33:Normal"
            _StyleDefs(69)  =   ":id=33,.parent=0"
            _StyleDefs(70)  =   "Named:id=34:Heading"
            _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(72)  =   ":id=34,.wraptext=-1"
            _StyleDefs(73)  =   "Named:id=35:Footing"
            _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(75)  =   "Named:id=36:Selected"
            _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=37:Caption"
            _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(79)  =   "Named:id=38:HighlightRow"
            _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=39:EvenRow"
            _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(83)  =   "Named:id=40:OddRow"
            _StyleDefs(84)  =   ":id=40,.parent=33"
            _StyleDefs(85)  =   "Named:id=41:RecordSelector"
            _StyleDefs(86)  =   ":id=41,.parent=34"
            _StyleDefs(87)  =   "Named:id=42:FilterBar"
            _StyleDefs(88)  =   ":id=42,.parent=33"
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label LblMes 
            AutoSize        =   -1  'True
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   9120
            TabIndex        =   19
            Top             =   30
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Pedido"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   45
            Width           =   11565
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   12840
         X2              =   24645
         Y1              =   375
         Y2              =   7425
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Pedido"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Restaurar Pedido"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Saldo del Pedido"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Anular Pedido"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar Pedido"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Emitir Pedido Anulado"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Gasto"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar Item            "
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar Item                "
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu menu2_1 
         Caption         =   "Agregar Documento"
      End
      Begin VB.Menu menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu2_3 
         Caption         =   "Eliminar Documento"
      End
   End
End
Attribute VB_Name = "FrmPedido2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMPEDIDO.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : EN ESTE FORMULARIO SE REGISTRAN LOS PEDIDOS DE LOS CLIENTES, DE INGRESO DE ESTOS
'*                    PEDIDOS SE GENERARA EL CRONOGRAMA DE ENTREGAS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 28/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstVent As New ADODB.Recordset     ' RECORDSET QUE ALMACENA LOS DATOS DE LA TABLA ped_pedido
Dim QueHace As Integer                 ' VARIABLE QUE ESPECIFICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim CaracteresNumericos As String      ' ESPECIFICA LOS CARACTERES NUMERICOS QUE SE UTILIZARAN EN LOS CONTROLES TEXTBOX
Dim SeEjecuto As Boolean               ' VARIABLE QUE CONTROLA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim Mostrando As Boolean               '
Dim agregando As Boolean               ' para saber cuando se este agregando FILAS AL CONTROL grid de productos
Dim xHorIni As Date                    ' ESPECIFICA LA HORA DE INICIO
Dim fOrdenLista As Boolean             ' especfica el orden de la lista de la consulta
Dim mIdRegistro&                       ' identificador del registro
Dim mMesActivo As Integer              ' indica el mes activo
Dim RstEntr As New ADODB.Recordset
Dim mCorrelativo As Long               ' para diferenciar la fecha de entrega del pedido cuando se necesite modificar

Dim fCierrePeriodo As Boolean          ' --indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer            ' INDICA EL CODIGO DEL MENU ACTIVO


'*****************************************************************************************************
'* Nombre           : ActivarEntorno
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA EL CONTROL TABONE Y TOOLBAR DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivarEntorno()
    TabOne1.Enabled = Not TabOne1.Enabled
    Toolbar1.Enabled = Not Toolbar1.Enabled
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA ped_pedido
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim rpta As Integer
    On Error Resume Next
    If RstVent.RecordCount = 0 Then
        MsgBox "No hay documentos para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    TabOne1.CurrTab = 0
    rpta = MsgBox("¿ Esta seguro de eliminar el Registro seleccionado ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If rpta = vbYes Then
        xCon.Execute "DELETE * FROM ped_pedidodetent WHERE idped = " & RstVent("id") & ""
        xCon.Execute "DELETE * FROM ped_pedidodet WHERE idped = " & RstVent("id") & ""
        xCon.Execute "DELETE * FROM ped_pedido WHERE id = " & RstVent("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstVent("id") & " AND idform = " & IdMenuActivo
        
        MsgBox RstVent("nomdoc") & " se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstVent.Requery
        Dg1.Refresh
        If RstVent.RecordCount = 0 Then
            rpta = MsgBox("No se han registrado movimientos en el periodo especificado, ¿ Desea agregar uno ahora ?", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo)
            If rpta = vbYes Then Nuevo
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : Anular
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ANULA UN PEDIDO REGISTRAO PARA ELLO ACTUALIZA EL CAMPO Anulado = -1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Anular()
    Dim rpta As Integer
    Dim A As Integer
    Dim Rst As New ADODB.Recordset
    rpta = MsgBox("¿Esta seguro de anular " & RstVent("nomdoc") & " Nº " & RstVent("numser") & "-" & RstVent("numdoc") + "?", vbYesNo + vbDefaultButton1 + vbQuestion, Me.Caption)
    On Error Resume Next
    If rpta = vbYes Then
        xCon.Execute "UPDATE ped_pedido  SET ped_pedido.Anulado = -1 " _
            & " WHERE ped_pedido.id = " & RstVent("id") & " "
        ' eliminando los registros del detalle
        xCon.Execute "DELETE * FROM ped_pedidodetent WHERE ped_pedidodetent.idped  = " & RstVent("id") & ""
        xCon.Execute "DELETE * FROM ped_pedidodet WHERE ped_pedidodet.idped  = " & RstVent("id") & ""
        
        ' Grabamos el movimiento en la tabla var_edicion
        GrabarOperacion xIdUsuario, IdMenuActivo, 2, Time, Time, Date, xCon, NulosN(RstVent("id"))
        
        MsgBox RstVent("nomdoc") & " se anuló con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstVent.Requery
        Dg1.Refresh
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL INGRESO O MODIFICACION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    Dim X As Integer
    Bloquea
    Fg1.ColComboList(1) = ""
    Label5.Caption = "Detalle de Pedido"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
     
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    
    Label5.Caption = "Agregando Pedido "
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    Fg2.SelectionMode = flexSelectionFree

    Fg1.Rows = 1
    Fg2.Rows = 12
    GRID_COLOR_FONDO Fg2, Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1, &HD0FDFA
    
    TxtFchEnt.Visible = False
    Label3(3).Visible = False
    
    mCorrelativo = 1
    PreparaRSTEntrega
    
    ' Se Inicializa la decah del documento a la fecha Actual
    TxtFchDoc.Valor = Format(Date, "dd/mm/yyyy")
    TxtTipDoc.Text = "107"
    TxtTipDoc_Validate True
    xHorIni = Time
    TxtFchDoc.SetFocus
    ' Se Inicializa la fecha de Entrega a la fecha Actual
    TxtFchEnt.Valor = Format(Date, "dd/mm/yyyy")
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA LA EDICION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    If RstVent.RecordCount = 0 Then
        MsgBox "No hay Registros para Modificar", vbInformation, Me.Caption
        Exit Sub
    End If
    If NulosC(RstVent("nombre")) = "ANULADO" Then
        MsgBox "El Documento de Venta esta Anulado" & vbCr & "No se Puede Modificar", vbInformation, Me.Caption
        Exit Sub
    End If
   
    QueHace = 2
    xHorIni = Time
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    mCorrelativo = 1
    MuestraSegundoTab
    Label5.Caption = "Modificando Pedido"
    
    Fg1.SelectionMode = flexSelectionFree
    Fg2.SelectionMode = flexSelectionFree
    Fg3.SelectionMode = flexSelectionFree
    xHorIni = Time
    TxtFchDoc.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTAÑA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'* Modificacion     :02/02/11 JOSE CHACON
                            'Utilizacion de la tabla ped_pedidosdetent como unica tabla de consulta
                            'descartandose la tabla ped_pedidodet
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim cSQL As String
    
    Blanquea
    
    If RstVent.RecordCount = 0 Then Exit Sub
    If RstVent.EOF = True Or RstVent.BOF = True Then Exit Sub
    
    If IsDate(RstVent("fchemi")) = True Then TxtFchDoc.Valor = CDate(RstVent("fchemi"))
    If IsDate(RstVent("fchent")) = True Then TxtFchEnt.Valor = CDate(RstVent("fchent"))
    
    TxtTipDoc.Text = NulosN(RstVent("tipdoc"))
    LblNomDoc.Caption = NulosC(RstVent("nomdoc"))
    TxtNumRuc.Text = NulosC(RstVent("numruc"))
    LblNomCli.Caption = NulosC(RstVent("nombre"))
    LblIdCliente.Caption = RstVent("idcli")
    TxtPtoVta.Text = NulosC(RstVent("idpunvecli"))
    LblPtoVta.Caption = NulosC(RstVent("ptovta"))
    TxtNumSer.Text = NulosC(RstVent("numser"))
    TxtNumDoc.Text = NulosC(RstVent("numdoc"))
    TxtConPag.Text = NulosC(RstVent("idconpag"))
    LblCondPag.Caption = NulosC(RstVent("desccond"))
    TxtTipPed.Text = NulosN(RstVent("idtipped"))
    LblTipPed.Caption = NulosC(RstVent("tipped"))
    
    TxtOC.Text = NulosC(RstVent("oc"))
    ' mostrando la fecha de entrega
    If NulosN(RstVent("idtipped")) = 1 Then
        Label3(3).Visible = True
        TxtFchEnt.Visible = True
    Else
        Label3(3).Visible = False
        TxtFchEnt.Visible = False
    End If
    
    txtglosa.Text = NulosC(RstVent("glosa"))
    
    ' Detalle del Documento
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
     
    ' CARGAMOS LOS ITEMS DE LA FACTURA
    Set RstDet = Nothing
    agregando = True
    
    On Error GoTo cargar2tipo

    cSQL = "SELECT DISTINCT ped_pedidodet.iditem, alm_inventario.codpro, alm_inventario.descripcion, Sum(ped_pedidodet.canpro) AS SumaDecanpro, mae_unidades.abrev, ped_pedidodet.estado, ped_pedidodet.idunimed, ped_pedidodet.observacion, IIf([ped_pedidodet].[estado]=1,[ped_pedidodet].[canpro],0) AS entregado " _
        + vbCr + "FROM (ped_pedidodet LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id " _
        + vbCr + "GROUP BY ped_pedidodet.iditem, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodet.estado, ped_pedidodet.idunimed, ped_pedidodet.observacion, IIf([ped_pedidodet].[estado]=1,[ped_pedidodet].[canpro],0), ped_pedidodet.idped " _
        + vbCr + "HAVING (((ped_pedidodet.idped)=" & RstVent("id") & "));"

    RST_Busq RstDet, cSQL, xCon
            
    Fg1.Rows = Fg1.FixedRows
    Fg2.Rows = Fg2.FixedRows
        
    If RstDet.State = 1 Then
        If RstDet.RecordCount <> 0 Then
            RstDet.MoveFirst
            Do While Not RstDet.EOF
                Fg1.Rows = Fg1.Rows + 1
                With Me.Fg1
                    .TextMatrix(.Rows - 1, 1) = NulosC(RstDet("codpro"))
                    .TextMatrix(.Rows - 1, 2) = NulosC(RstDet("Descripcion"))
                    .TextMatrix(.Rows - 1, 3) = NulosC(RstDet("abrev"))
                    .TextMatrix(.Rows - 1, 4) = NulosC(RstDet("SumaDecanpro"))
                    .TextMatrix(.Rows - 1, 5) = NulosC(RstDet("iditem"))
                    .TextMatrix(.Rows - 1, 6) = NulosC(RstDet("idunimed"))
                End With
                RstDet.MoveNext
            Loop
            Fg1.Row = 1
        End If
    End If
    
    Set RstDet = Nothing
    ' Cargar las entregas programadas [rst temporal]
    PreparaRSTEntrega
    
'SELECT DISTINCT ped_pedidodet.iditem, ped_pedidodet.idpeddet, alm_inventario.codpro, alm_inventario.descripcion, ped_pedidodet.fchent, ped_pedidodet.canpro, ped_pedidodet.canproent, mae_unidades.abrev, ped_pedidodet.estado, ped_pedidodet.idunimed, ped_pedidodet.observacion, 100 AS porcen
'FROM (ped_pedidodet LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id
'Where (((ped_pedidodet.idPed) = 2905))
'GROUP BY ped_pedidodet.iditem, ped_pedidodet.idpeddet, alm_inventario.codpro, alm_inventario.descripcion, ped_pedidodet.fchent, ped_pedidodet.canpro, ped_pedidodet.canproent, mae_unidades.abrev, ped_pedidodet.estado, ped_pedidodet.idunimed, ped_pedidodet.observacion, 100;
    
    cSQL = "SELECT DISTINCT ped_pedidodet.iditem, ped_pedidodet.idpeddet, alm_inventario.codpro, alm_inventario.descripcion, ped_pedidodet.fchent, ped_pedidodet.canpro, ped_pedidodet.canproent, mae_unidades.abrev, ped_pedidodet.estado, ped_pedidodet.idunimed, ped_pedidodet.observacion, 100 AS porcen " _
        + vbCr + "FROM (ped_pedidodet LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id " _
        + vbCr + "Where (((ped_pedidodet.idped) =" & RstVent("id") & ")) " _
        + vbCr + "GROUP BY ped_pedidodet.iditem, ped_pedidodet.idpeddet, alm_inventario.codpro, alm_inventario.descripcion, ped_pedidodet.fchent, ped_pedidodet.canpro, ped_pedidodet.canproent, mae_unidades.abrev, ped_pedidodet.estado, ped_pedidodet.idunimed, ped_pedidodet.observacion, 100;"

    RST_Busq RstDet, cSQL, xCon
    
    If RstDet.State = 1 Then
        If RstDet.RecordCount <> 0 Then RstDet.MoveFirst
        While Not RstDet.EOF
            mCorrelativo = mCorrelativo + 1
            RstEntr.AddNew
            RstEntr("iditem") = RstDet("iditem")
            RstEntr("fchent") = Format(CDate(RstDet("fchent")), FORMAT_DATE)
            RstEntr("porcen") = NulosN(RstDet("porcen"))
            RstEntr("canpro") = NulosN(RstDet("canpro"))
            RstEntr("canproent") = NulosN(RstDet("canproent"))
            RstEntr("corr") = mCorrelativo
            
            RstEntr("idpeddet") = NulosN(RstDet("idpeddet"))
            RstEntr.Update
            RstDet.MoveNext
        Wend
    End If
    Set RstDet = Nothing
    
    agregando = False
    ' se activa la vision del detalle del primer pedido
    Fg1_RowColChange
    Exit Sub
    
' para el formato de lectura antiguo
cargar2tipo:
'    cSQL = "SELECT DISTINCT ped_pedidodetent.iditem, alm_inventario.codpro, alm_inventario.descripcion, Sum(ped_pedidodetent.canpro) AS SumaDecanpro, mae_unidades.abrev, ped_pedidodetent.estado, ped_pedidodetent.idunimed, ped_pedidodetent.observacion, IIf([ped_pedidodetent].[estado]=1,[ped_pedidodetent].[canpro],0) AS entregado " _
'        + vbCr + "FROM (ped_pedidodetent LEFT JOIN mae_unidades ON ped_pedidodetent.idunimed = mae_unidades.id) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id " _
'        + vbCr + "GROUP BY ped_pedidodetent.iditem, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodetent.estado, ped_pedidodetent.idunimed, ped_pedidodetent.observacion, IIf([ped_pedidodetent].[estado]=1,[ped_pedidodetent].[canpro],0), ped_pedidodetent.idped " _
'        + vbCr + "HAVING (((ped_pedidodetent.idped)=" & RstVent("id") & "));"
        
    cSQL = "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodet.canpro,  ped_pedidodet.estado, ped_pedidodet.iditem, ped_pedidodet.idunimed " _
        + vbCr + "FROM mae_unidades RIGHT JOIN (alm_inventario INNER JOIN ped_pedidodet ON alm_inventario.id = ped_pedidodet.iditem) ON mae_unidades.id = ped_pedidodet.idunimed " _
        + vbCr + "WHERE (((ped_pedidodet.idped)=" & RstVent("id") & "));"

    RST_Busq RstDet, cSQL, xCon
            
    Fg1.Rows = Fg1.FixedRows
    Fg2.Rows = Fg2.FixedRows
        
    If RstDet.State = 1 Then
        If RstDet.RecordCount <> 0 Then
            RstDet.MoveFirst
            Do While Not RstDet.EOF
                Fg1.Rows = Fg1.Rows + 1
                
                With Me.Fg1
                    .TextMatrix(.Rows - 1, 1) = NulosC(RstDet("codpro"))
                    .TextMatrix(.Rows - 1, 2) = NulosC(RstDet("Descripcion"))
                    .TextMatrix(.Rows - 1, 3) = NulosC(RstDet("abrev"))
                    .TextMatrix(.Rows - 1, 4) = NulosN(RstDet("canpro"))
                    .TextMatrix(.Rows - 1, 5) = NulosC(RstDet("iditem"))
                    .TextMatrix(.Rows - 1, 6) = NulosC(RstDet("idunimed"))
                End With
                RstDet.MoveNext
            Loop
            Fg1.Row = 1
        End If
    End If
    
    Set RstDet = Nothing
    ' Cargar las entregas programadas [rst temporal]
    PreparaRSTEntrega
    
'    cSQL = "SELECT DISTINCT ped_pedidodetent.iditem, alm_inventario.codpro, alm_inventario.descripcion, ped_pedidodetent.fchent, ped_pedidodetent.canpro, mae_unidades.abrev, ped_pedidodetent.estado, ped_pedidodetent.idunimed, ped_pedidodetent.observacion, 100 AS porcen, IIf([ped_pedidodetent].[estado]=1,[ped_pedidodetent].[canpro],0) AS entregado " _
'        + vbCr + "FROM (ped_pedidodetent LEFT JOIN mae_unidades ON ped_pedidodetent.idunimed = mae_unidades.id) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id " _
'        + vbCr + "Where (((ped_pedidodetent.idped) =" & RstVent("id") & ")) " _
'        + vbCr + "GROUP BY ped_pedidodetent.iditem, alm_inventario.codpro, alm_inventario.descripcion, ped_pedidodetent.fchent, ped_pedidodetent.canpro, mae_unidades.abrev, ped_pedidodetent.estado, ped_pedidodetent.idunimed, ped_pedidodetent.observacion, 100, IIf([ped_pedidodetent].[estado]=1,[ped_pedidodetent].[canpro],0);"
        
    cSQL = "SELECT ped_pedidodetent.idped, ped_pedidodetent.iditem, ped_pedidodetent.fchent, ped_pedidodetent.canpro, ped_pedidodetent.estado, ped_pedidodetent.observacion, [ped_pedidodetent].[canpro]/[ped_pedidodet].[canpro]*100 AS porcen, IIf([ped_pedidodetent].[estado]=1,[ped_pedidodetent].[canpro],0) AS entregado " _
        + vbCr + "FROM ped_pedidodet INNER JOIN ped_pedidodetent ON (ped_pedidodet.iditem = ped_pedidodetent.iditem) AND (ped_pedidodet.idped = ped_pedidodetent.idped) " _
        + vbCr + "WHERE (((ped_pedidodetent.idped)=" & RstVent("id") & ")) ORDER BY ped_pedidodetent.fchent;"

    RST_Busq RstDet, cSQL, xCon
    
    If RstDet.State = 1 Then
        If RstDet.RecordCount <> 0 Then RstDet.MoveFirst
        
        Do While Not RstDet.EOF
            mCorrelativo = mCorrelativo + 1
            RstEntr.AddNew
            RstEntr("iditem") = RstDet("iditem")
            RstEntr("fchent") = Format(CDate(RstDet("fchent")), FORMAT_DATE)
            RstEntr("porcen") = NulosN(RstDet("porcen"))
            RstEntr("canpro") = NulosN(RstDet("canpro"))
            RstEntr("canproent") = NulosN(RstDet("entregado"))
            
            RstEntr("idpeddet") = 0
            RstEntr("corr") = mCorrelativo
            RstEntr.Update
            RstDet.MoveNext
        Loop
    End If
    Set RstDet = Nothing
    
    agregando = False
    Fg1_RowColChange
End Sub

'Sub MuestraSegundoTab2()
'    Dim Rst As New ADODB.Recordset
'    Dim xRs As New ADODB.Recordset
'    Dim cSQL As String
'
'    Blanquea
'
'    If RstVent.RecordCount = 0 Then Exit Sub
'    If RstVent.EOF = True Or RstVent.BOF = True Then Exit Sub
'
'    If IsDate(RstVent("fchemi")) = True Then TxtFchDoc.Valor = CDate(RstVent("fchemi"))
'    If IsDate(RstVent("fchent")) = True Then TxtFchEnt.Valor = CDate(RstVent("fchent"))
'
'    TxtTipDoc.Text = NulosN(RstVent("tipdoc"))
'    LblNomDoc.Caption = NulosC(RstVent("nomdoc"))
'    TxtNumRuc.Text = NulosC(RstVent("numruc"))
'    LblNomCli.Caption = NulosC(RstVent("nombre"))
'    LblIdCliente.Caption = RstVent("idcli")
'    TxtPtoVta.Text = NulosC(RstVent("idpunvecli"))
'    LblPtoVta.Caption = NulosC(RstVent("ptovta"))
'    TxtNumSer.Text = NulosC(RstVent("numser"))
'    TxtNumDoc.Text = NulosC(RstVent("numdoc"))
'    TxtConPag.Text = NulosC(RstVent("idconpag"))
'    LblCondPag.Caption = NulosC(RstVent("desccond"))
'    TxtTipPed.Text = NulosN(RstVent("idtipped"))
'    LblTipPed.Caption = NulosC(RstVent("tipped"))
'
'    TxtOC.Text = NulosC(RstVent("oc"))
'    ' mostrando la fecha de entrega
'    If NulosN(RstVent("idtipped")) = 1 Then
'        Label3(3).Visible = True
'        TxtFchEnt.Visible = True
'    Else
'        Label3(3).Visible = False
'        TxtFchEnt.Visible = False
'    End If
'
'    txtglosa.Text = NulosC(RstVent("glosa"))
'
'    ' Detalle del Documento
'    Dim RstDet As New ADODB.Recordset
'    Dim A As Integer
'
'    ' CARGAMOS LOS ITEMS DE LA FACTURA
'    Set RstDet = Nothing
'    agregando = True
'
''SELECT DISTINCT ped_pedidodet.iditem, alm_inventario.codpro, alm_inventario.descripcion, Sum(ped_pedidodet.canpro) AS SumaDecanpro, mae_unidades.abrev, ped_pedidodet.estado, ped_pedidodet.idunimed, ped_pedidodet.observacion, IIf([ped_pedidodet].[estado]=1,[ped_pedidodet].[canpro],0) AS entregado
''FROM (ped_pedidodet LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id
''GROUP BY ped_pedidodet.iditem, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodet.estado, ped_pedidodet.idunimed, ped_pedidodet.observacion, IIf([ped_pedidodet].[estado]=1,[ped_pedidodet].[canpro],0), ped_pedidodet.idped
''HAVING (((ped_pedidodet.idped)=2902));
'
'    cSQL = "SELECT DISTINCT ped_pedidodetent.iditem, alm_inventario.codpro, alm_inventario.descripcion, Sum(ped_pedidodetent.canpro) AS SumaDecanpro, mae_unidades.abrev, ped_pedidodetent.estado, ped_pedidodetent.idunimed, ped_pedidodetent.observacion, IIf([ped_pedidodetent].[estado]=1,[ped_pedidodetent].[canpro],0) AS entregado " _
'        + vbCr + "FROM (ped_pedidodetent LEFT JOIN mae_unidades ON ped_pedidodetent.idunimed = mae_unidades.id) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id " _
'        + vbCr + "GROUP BY ped_pedidodetent.iditem, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodetent.estado, ped_pedidodetent.idunimed, ped_pedidodetent.observacion, IIf([ped_pedidodetent].[estado]=1,[ped_pedidodetent].[canpro],0), ped_pedidodetent.idped " _
'        + vbCr + "HAVING (((ped_pedidodetent.idped)=" & RstVent("id") & "));"
'
'    RST_Busq RstDet, cSQL, xCon
'
''    RST_Busq RstDet, " SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodet.canpro,  ped_pedidodet.estado, ped_pedidodet.iditem, ped_pedidodet.idunimed " _
''        & " FROM mae_unidades RIGHT JOIN (alm_inventario INNER JOIN ped_pedidodet ON alm_inventario.id = ped_pedidodet.iditem) ON mae_unidades.id = ped_pedidodet.idunimed " _
''        & " WHERE (((ped_pedidodet.idped)=" & RstVent("id") & "));", xCon
'
'    Fg1.Rows = Fg1.FixedRows
'    Fg2.Rows = Fg2.FixedRows
'
'    If RstDet.State = 1 Then
'        If RstDet.RecordCount <> 0 Then
'            RstDet.MoveFirst
'            Do While Not RstDet.EOF
'                Fg1.Rows = Fg1.Rows + 1
'
'                With Me.Fg1
'                    .TextMatrix(.Rows - 1, 1) = NulosC(RstDet("codpro"))
'                    .TextMatrix(.Rows - 1, 2) = NulosC(RstDet("Descripcion"))
'                    .TextMatrix(.Rows - 1, 3) = NulosC(RstDet("abrev"))
'                    .TextMatrix(.Rows - 1, 4) = NulosC(RstDet("SumaDecanpro"))
'                    '.TextMatrix(.Rows - 1, 4) = NulosN(RstDet("canpro"))
'                    .TextMatrix(.Rows - 1, 5) = NulosC(RstDet("iditem"))
'                    .TextMatrix(.Rows - 1, 6) = NulosC(RstDet("idunimed"))
'                End With
'                RstDet.MoveNext
'            Loop
'            Fg1.Row = 1
'        End If
'    End If
'
'    Set RstDet = Nothing
'    ' Cargar las entregas programadas [rst temporal]
'    PreparaRSTEntrega
'
''SELECT DISTINCT ped_pedidodetent.iditem, alm_inventario.codpro, alm_inventario.descripcion, ped_pedidodetent.fchent, ped_pedidodetent.canpro, mae_unidades.abrev, ped_pedidodetent.estado, ped_pedidodetent.idunimed, ped_pedidodetent.observacion, 100 AS porcen, IIf([ped_pedidodetent].[estado]=1,[ped_pedidodetent].[canpro],0) AS entregado
''FROM (ped_pedidodetent LEFT JOIN mae_unidades ON ped_pedidodetent.idunimed = mae_unidades.id) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id
''Where (((ped_pedidodetent.idped) = 67))
''GROUP BY ped_pedidodetent.iditem, alm_inventario.codpro, alm_inventario.descripcion, ped_pedidodetent.fchent, ped_pedidodetent.canpro, mae_unidades.abrev, ped_pedidodetent.estado, ped_pedidodetent.idunimed, ped_pedidodetent.observacion, 100, IIf([ped_pedidodetent].[estado]=1,[ped_pedidodetent].[canpro],0);
'
'    cSQL = "SELECT DISTINCT ped_pedidodetent.iditem, alm_inventario.codpro, alm_inventario.descripcion, ped_pedidodetent.fchent, ped_pedidodetent.canpro, mae_unidades.abrev, ped_pedidodetent.estado, ped_pedidodetent.idunimed, ped_pedidodetent.observacion, 100 AS porcen, IIf([ped_pedidodetent].[estado]=1,[ped_pedidodetent].[canpro],0) AS entregado " _
'        + vbCr + "FROM (ped_pedidodetent LEFT JOIN mae_unidades ON ped_pedidodetent.idunimed = mae_unidades.id) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id " _
'        + vbCr + "Where (((ped_pedidodetent.idped) =" & RstVent("id") & ")) " _
'        + vbCr + "GROUP BY ped_pedidodetent.iditem, alm_inventario.codpro, alm_inventario.descripcion, ped_pedidodetent.fchent, ped_pedidodetent.canpro, mae_unidades.abrev, ped_pedidodetent.estado, ped_pedidodetent.idunimed, ped_pedidodetent.observacion, 100, IIf([ped_pedidodetent].[estado]=1,[ped_pedidodetent].[canpro],0);"
'
'    RST_Busq RstDet, cSQL, xCon
'
''    RST_Busq RstDet, "SELECT ped_pedidodetent.idped, ped_pedidodetent.iditem, ped_pedidodetent.fchent, ped_pedidodetent.canpro, ped_pedidodetent.estado, ped_pedidodetent.observacion, [ped_pedidodetent].[canpro]/[ped_pedidodet].[canpro]*100 AS porcen, IIf([ped_pedidodetent].[estado]=1,[ped_pedidodetent].[canpro],0) AS entregado " _
''        & " FROM ped_pedidodet INNER JOIN ped_pedidodetent ON (ped_pedidodet.iditem = ped_pedidodetent.iditem) AND (ped_pedidodet.idped = ped_pedidodetent.idped) WHERE (((ped_pedidodetent.idped)=" & RstVent("id") & ")) ORDER BY ped_pedidodetent.fchent;", xCon
'
'    If RstDet.State = 1 Then
'        If RstDet.RecordCount <> 0 Then RstDet.MoveFirst
'
'        Do While Not RstDet.EOF
'            mCorrelativo = mCorrelativo + 1
'            RstEntr.AddNew
'            RstEntr("iditem") = RstDet("iditem")
'            RstEntr("fchent") = Format(CDate(RstDet("fchent")), FORMAT_DATE)
'            RstEntr("porcen") = NulosN(RstDet("porcen"))
'            RstEntr("canpro") = NulosN(RstDet("canpro"))
'            RstEntr("entregado") = NulosN(RstDet("entregado"))
'            RstEntr("corr") = mCorrelativo
'            RstEntr.Update
'            RstDet.MoveNext
'        Loop
'    End If
'    Set RstDet = Nothing
'
'    agregando = False
'    Fg1_RowColChange
'End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TEXTBOX
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea()
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    
    TxtIdAlm.Locked = Not TxtIdAlm.Locked
    
    TxtPtoVta.Locked = Not TxtPtoVta.Locked
    TxtFchDoc.Locked = Not TxtFchDoc.Locked
    TxtFchEnt.Locked = Not TxtFchEnt.Locked
    TxtTipPed.Locked = Not TxtTipPed.Locked
    TxtConPag.Locked = Not TxtConPag.Locked
    
    TxtOC.Locked = Not TxtOC.Locked
    
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : INICIALIZA LOS CONTROLES TEXTBOX PARA EL INGRESO DE DATOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    TxtIdAlm = ""
    TxtFchDoc.Valor = ""
    TxtFchEnt.Valor = ""
    TxtTipDoc.Text = ""
    TxtNumRuc.Text = ""
    TxtPtoVta.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtFchDoc.Valor = ""
    TxtTipPed.Text = ""
    TxtConPag.Text = ""
    txtglosa.Text = ""
    TxtOC.Text = ""
    LblPtoVta.Caption = ""
    LblAlmacen = ""
    LblNomDoc.Caption = ""
    LblNomCli.Caption = ""
    LblAlmacen.Caption = ""
    LblIdCliente.Caption = ""
    LblCondPag.Caption = ""
    LblTipPed.Caption = ""

    Fg2.Rows = Fg2.FixedRows
    Fg1.Rows = Fg1.FixedRows
End Sub

Private Sub CmdAceptar_Click()
    FrmModFecha.Visible = False
End Sub

Private Sub CmdAddItem_Click()
    If QueHace = 3 Then Exit Sub
    agregando = True
   
    If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 2)) = "" Then
        Fg1.Row = Fg1.Rows - 1
        Fg1.Col = 2
        Fg1_CellButtonClick Fg1.Rows - 1, 2
        Fg1.SetFocus
        agregando = False
        Exit Sub
    End If
    
    Fg1.Rows = Fg1.Rows + 1
    Fg1.Row = Fg1.Rows - 1
    Fg1.Col = 2
   
    agregando = False
    Fg1.SetFocus
End Sub

Private Sub CmdBusAlm_Click()
    ' EJECUTA LA BUSQUEDA DE UN ALMACEN
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT alm_almacenes.* FROM alm_almacenes"
    
    xform.Titulo = "Buscando Almacenes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblAlmacen.Caption = xRs("descripcion")
        TxtIdAlm.Text = xRs("id")
        TxtNumRuc.SetFocus
        
        If TxtTipDoc.Text <> "" Then
            Dim Rst As New ADODB.Recordset
            Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc.Text) & " AND idalm = " & NulosN(TxtIdAlm.Text) & "", xCon)
            If Rst.RecordCount <> 0 Then
                TxtNumSer.Text = Rst("numser")
                TxtNumSer_Validate True
            End If
            
            Set Rst = Nothing
        Else
            TxtNumSer.Text = ""
            TxtNumDoc.Text = ""
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusNumSer_Click()
    ' EJECUTA LA BUSQUEDA DE NUMERO DE SERIE
    If QueHace = 3 Then Exit Sub

    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Sub
    End If

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "idtipdoc":       xCampos(0, 2) = "1500":    xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion": xCampos(1, 2) = "2500":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Serie":          xCampos(2, 1) = "numser":      xCampos(2, 2) = "1500":    xCampos(2, 3) = "C"

    xform.SQLCad = " SELECT mae_documento.*,  alm_numseries.idtipdoc, alm_numseries.numser " & _
        " FROM mae_documento LEFT JOIN alm_numseries ON mae_documento.id = alm_numseries.idtipdoc " & _
        " WHERE alm_numseries.idalm =" & NulosN(Me.TxtIdAlm) & " AND alm_numseries.idtipdoc = " & NulosN(TxtTipDoc) & ""
        
    xform.Titulo = "Buscando Series"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numser"
    xform.CampoBusca = "numser"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumSer.Text = Format(xRs("numser"), "0000")
            
            Dim Rst As New ADODB.Recordset
            RST_Busq Rst, "SELECT top 1 numdoc AS numero from ped_pedido  WHERE numser ='" & NulosC(TxtNumSer.Text) & "' AND tipdoc =" & NulosN(TxtTipDoc) & " ORDER BY numdoc DESC ", xCon

            If Rst.RecordCount = 0 Then
                TxtNumDoc.Text = "0000000001"
            Else
                Rst.MoveFirst
                TxtNumDoc.Text = Format(NulosN(Rst("numero")) + 1, "0000000000")
            End If
            Set Rst = Nothing
        End If
        
        TxtNumDoc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusCli_Click()
    ' EJECUTA LA BUSQUEDA DE UN CLIENTE
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Cliente":      xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_cliente.nombre, mae_cliente.numruc, mae_cliente.id, mae_cliente.idven From mae_cliente"
    
    xform.Titulo = "Buscando Cliente"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            If xRs.RecordCount <> 0 Then
                TxtNumRuc.Text = xRs("numruc")
                LblNomCli.Caption = xRs("nombre")
                LblIdCliente.Caption = xRs("id")
               TxtPtoVta.SetFocus
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDoc_Click()
    ' EJECUTA LA BUSQUEDA DE UN DOCUMENTO CONTABLE
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"

    xform.SQLCad = "SELECT mae_documento.id, mae_documento.descripcion, mae_documento.abrev From mae_documento WHERE (((mae_documento.id)=107))"

    xform.Titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipDoc.Text = xRs("id")
            LblNomDoc.Caption = xRs("descripcion")
            TxtNumSer.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdCancelar_Click()
    FrmModFecha.Visible = False
End Sub

Private Sub CmdDelItem_Click()
    If QueHace = 3 Then Exit Sub
    If Fg1.Row < 1 Or Fg1.Rows < 1 Then Exit Sub
        
    ' eliminar rst temporal
    RstRegistroEliminar RstEntr, "iditem", NulosN(Fg1.TextMatrix(Fg1.Row, 5)), True
    Fg1.RemoveItem Fg1.Row
    
    If Fg1.Rows <> 1 Then Fg1.Select Fg1.Rows - 1, 1
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstVent
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DESCENDETE LAS COLUMNAS DEL CONTROL Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstVent.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstVent("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col <> 2 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(4, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "4800":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Unid.":         xCampos(1, 1) = "abrev":         xCampos(1, 2) = "500":     xCampos(1, 3) = "C"
    xCampos(2, 0) = "Stock":        xCampos(2, 1) = "stckact":        xCampos(2, 2) = "800":     xCampos(2, 3) = "N"
    xCampos(3, 0) = "Código":       xCampos(3, 1) = "codpro":         xCampos(3, 2) = "2000":    xCampos(3, 3) = "C"
    
    Dim nSQLId As String
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 5, "alm_inventario.id", " NOT IN ", True)
    If nSQLId <> "" Then nSQLId = " AND " & nSQLId
    
    ' obs. apareceran solo items de ventas que tengan cuenta contable
    xform.SQLCad = "SELECT alm_inventario.*, mae_unidades.descripcion AS descuni, mae_unidades.abrev, mae_percepcion.tasa " _
        & " FROM mae_unidades RIGHT JOIN (mae_percepcion RIGHT JOIN alm_inventario ON mae_percepcion.id = alm_inventario.idper) " _
        & " ON mae_unidades.id = alm_inventario.idunimed Where alm_inventario.activo=-1 and  (((alm_inventario.tippro) in (1,3)  )) " & nSQLId & " AND alm_inventario.idcuentaven <>0 ORDER BY alm_inventario.descripcion"
    
    xform.Titulo = "Buscando Productos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.FormaBusca = CualquierParte
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    Dim A As Integer
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 1) = NulosC(xRs("codpro"))
            Fg1.TextMatrix(Fg1.Row, 2) = NulosC(xRs("descripcion"))
            Fg1.TextMatrix(Fg1.Row, 3) = NulosC(xRs("abrev"))
            Fg1.TextMatrix(Fg1.Row, 4) = 0
            Fg1.TextMatrix(Fg1.Row, 5) = xRs("id")
            Fg1.TextMatrix(Fg1.Row, 6) = NulosN(xRs("idunimed"))
        End If
    End If
    Fg1.Col = 4
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim xTot As Double
    If agregando = True Then Exit Sub
    ' verificar que la cantidad no supere al detalle
    If Col = 4 Then
        xTot = RstRegistroSumar(RstEntr, "canpro", "iditem", Fg1.TextMatrix(Row, 5), "N", True)
        
        If xTot > NulosN(Fg1.TextMatrix(Row, 4)) Then
            MsgBox "La cantidad del detalle de entrega es superior a la cantidad del pedido ", vbExclamation, xTitulo
        End If
    End If
End Sub

Private Sub Fg1_EnterCell()
    If agregando = True Then Exit Sub
    
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    
    If Fg1.Col = 2 Or Fg1.Col = 4 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Then KeyAscii = 0
    If KeyAscii = 13 Then Exit Sub
    
    Select Case Col
        Case 4
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 46 Then CmdDelItem_Click
    If KeyCode = 45 Then CmdAddItem_Click
End Sub

Private Sub Fg1_RowColChange()
    Dim xCol As Long
    If agregando = True Then Exit Sub
    If RstEntr.State = 0 Then Exit Sub
    If Fg1.Row < 1 Then Exit Sub
    ' si no has item =>> salir
    If NulosN(Fg1.TextMatrix(Fg1.Row, 5)) = 0 Then Exit Sub
    
    Fg2.Rows = 1
    RstEntr.Filter = "iditem=" & NulosN(Fg1.TextMatrix(Fg1.Row, 5))
    
    If RstEntr.RecordCount <> 0 Then
    
        RstEntr.MoveFirst
        xCol = 1
        agregando = True
        Do While Not RstEntr.EOF
            Fg2.Rows = Fg2.Rows + 1
            If IsDate(RstEntr("fchent")) = True Then
                Fg2.TextMatrix(xCol, 1) = Format(CDate(RstEntr("fchent")), FORMAT_DATE)
            End If
            
            Fg2.TextMatrix(xCol, 2) = Format(NulosN(RstEntr("porcen")), "0.00")
            Fg2.TextMatrix(xCol, 3) = NulosN(RstEntr("canpro"))
            Fg2.TextMatrix(xCol, 4) = RstEntr("corr")
            Fg2.TextMatrix(xCol, 5) = RstEntr("canproent")
            Fg2.TextMatrix(xCol, 6) = RstEntr("idpeddet")
            xCol = xCol + 1
            RstEntr.MoveNext
        Loop
        
        If Fg2.Rows < 12 Then
            Fg2.Rows = 12
        Else
            Fg2.Rows = Fg2.Rows + 2
        End If
        
        Fg2.TextMatrix(Fg2.Rows - 2, 1) = "Total"
        Fg2.TextMatrix(Fg2.Rows - 2, 2) = GRID_SUMAR_COL(Fg2, 2)
        Fg2.TextMatrix(Fg2.Rows - 2, 3) = GRID_SUMAR_COL(Fg2, 3)
        Fg2.TextMatrix(Fg2.Rows - 2, 5) = GRID_SUMAR_COL(Fg2, 5)
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = "Por Entregar"
        Fg2.TextMatrix(Fg2.Rows - 1, 5) = NulosN(Fg2.TextMatrix(Fg2.Rows - 2, 3)) - NulosN(Fg2.TextMatrix(Fg2.Rows - 2, 5))
        GRID_COLOR_FONDO Fg2, Fg2.Rows - 2, 1, Fg2.Rows - 1, Fg2.Cols - 1, &HD0FDFA
        agregando = False
    Else
        Fg2.Rows = 12
        GRID_COLOR_FONDO Fg2, Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1, &HD0FDFA
    End If
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If QueHace = 3 Then Exit Sub
    If agregando = True Then Exit Sub
    
    If Fg1.Row < 1 Then Exit Sub
    If (Fg2.TextMatrix(Row, 1) = "  /  /    " Or IsDate(Fg2.TextMatrix(Row, 1)) = False) And NulosN(Fg2.TextMatrix(Row, 2)) = 0 Then
        RstEntr.Filter = "corr=" & NulosN(Fg2.TextMatrix(Row, 4))
        If RstEntr.RecordCount <> 0 Then
            RstEntr.Delete
            RstEntr.Filter = ""
        End If
        Fg2.TextMatrix(Row, 1) = ""
        Fg2.TextMatrix(Row, 2) = ""
        Fg2.TextMatrix(Row, 3) = ""
        Exit Sub
    End If
    
    ' validando datos
    If NulosN(TxtTipPed.Text) = 0 Then
        MsgBox "Falta seleccionar el tipo de Pedido", vbExclamation, xTitulo
        Fg2.TextMatrix(Row, Col) = ""
        TxtTipPed.SetFocus
        Exit Sub
    End If
    
    If Col = 1 Then
        If IsDate(Fg2.TextMatrix(Row, Col)) = True Then
            Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, 1), FORMAT_DATE)
        Else
            MsgBox "Fecha Incorrecta", vbExclamation, xTitulo
            Fg2.TextMatrix(Row, Col) = ""
        End If
    ElseIf Col = 2 Then           ' porcentaje
        If IsNumeric(Fg2.TextMatrix(Row, Col)) = False Or NulosN(Fg2.TextMatrix(Row, Col)) > 100 Then
            MsgBox "Porcentaje Incorrecto", vbExclamation, xTitulo
            Fg2.TextMatrix(Row, Col) = ""
            Fg2.TextMatrix(Row, 3) = ""
        Else
            If NulosN(Fg1.TextMatrix(Fg1.Row, 4)) <> 0 Then
                Fg2.TextMatrix(Row, 3) = NulosN(Fg1.TextMatrix(Fg1.Row, 4)) * NulosN(Fg2.TextMatrix(Row, 2)) / 100
            End If
        End If
        Fg2.TextMatrix(Row, 2) = Format(Fg2.TextMatrix(Row, 2), "0.00")
    
    ElseIf Col = 3 Then           ' cantidad
        If IsNumeric(Fg2.TextMatrix(Row, Col)) = False Then
            MsgBox "Cantidad Incorrecta", vbExclamation, xTitulo
            Fg2.TextMatrix(Row, Col) = ""
            Fg2.TextMatrix(Row, 2) = ""
        Else
            If NulosN(Fg1.TextMatrix(Fg1.Row, 4)) <> 0 And NulosN(Fg2.TextMatrix(Row, 3)) <> 0 Then
                Fg2.TextMatrix(Row, 2) = (NulosN(Fg2.TextMatrix(Row, 3)) / NulosN(Fg1.TextMatrix(Fg1.Row, 4))) * 100
                Fg2.TextMatrix(Row, 2) = Format(Fg2.TextMatrix(Row, 2), "0.00")
            Else
                Fg2.TextMatrix(Row, 2) = ""
            End If
        End If
    End If
    
    ' agregando datos al rst
    If NulosN(Fg2.TextMatrix(Row, 4)) = 0 Then
        RstEntr.AddNew
        mCorrelativo = mCorrelativo + 1
        Fg2.TextMatrix(Row, 4) = mCorrelativo
    Else
        RstEntr.Filter = ""
        RstEntr.MoveFirst
        RstEntr.Filter = "corr=" & NulosN(Fg2.TextMatrix(Row, 4))
        If RstEntr.RecordCount = 0 Or RstEntr.EOF = True Or RstEntr.BOF = True Then Exit Sub
    End If
    
    RstEntr("iditem") = NulosN(Fg1.TextMatrix(Fg1.Row, 5))
    
    If IsDate(Fg2.TextMatrix(Row, 1)) = True Then
        RstEntr("fchent") = CDate(Fg2.TextMatrix(Row, 1))
    Else
'        RstEntr("fchent") = Null
    End If
    
    RstEntr("porcen") = NulosN(Fg2.TextMatrix(Row, 2))
    RstEntr("canpro") = NulosN(Fg2.TextMatrix(Row, 3))
    RstEntr("corr") = Fg2.TextMatrix(Row, 4)
    RstEntr.Update
    
    ' totalizar
    generarTotales Col
End Sub

Private Sub generarTotales(COLUMNA_ As Long)
    
    If COLUMNA_ = 2 Or COLUMNA_ = 3 Then
        agregando = True
        
        RstEntr.Filter = adFilterNone
        
        Fg2.TextMatrix(Fg2.Rows - 2, 1) = "Total"
        Fg2.TextMatrix(Fg2.Rows - 2, 2) = RstRegistroSumar(RstEntr, "porcen", "iditem", NulosN(Fg1.TextMatrix(Fg1.Row, 5)), "N", True)
        Fg2.TextMatrix(Fg2.Rows - 2, 3) = RstRegistroSumar(RstEntr, "canpro", "iditem", NulosN(Fg1.TextMatrix(Fg1.Row, 5)), "N", True)
        Fg2.TextMatrix(Fg2.Rows - 2, 5) = RstRegistroSumar(RstEntr, "canproent", "iditem", NulosN(Fg1.TextMatrix(Fg1.Row, 5)), "N", True)
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = "Por Entregar"
        Fg2.TextMatrix(Fg2.Rows - 1, 5) = NulosN(Fg2.TextMatrix(Fg2.Rows - 2, 3)) - NulosN(Fg2.TextMatrix(Fg2.Rows - 2, 5))
        
        GRID_COLOR_FONDO Fg2, Fg2.Rows - 2, 1, Fg2.Rows - 1, Fg2.Cols - 1, &HD0FDFA
        agregando = False
    End If
End Sub

Private Sub Fg2_EnterCell()
    If QueHace = 3 Then
        Fg2.Editable = flexEDNone
        Exit Sub
    End If
    
    If (Fg2.Col = 1 Or Fg2.Col = 3 Or Fg2.Col = 2) And Fg2.Row < Fg2.Rows - 2 Then
        If NulosN(TxtTipPed.Text) = 1 Or NulosN(Fg2.TextMatrix(Fg2.Row, 5)) <> 0 Then
            Fg2.Editable = flexEDNone
        Else
            Fg2.Editable = flexEDKbdMouse
        End If
    Else
        Fg2.Editable = flexEDNone
    End If
End Sub

Private Sub Fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Then KeyAscii = 0
    If KeyAscii = 13 Then Exit Sub
    
    Select Case Col
        Case 1, 2, 3
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        'cmdagregardocs_Click
    End If
    
    If KeyCode = 46 Then
        If Fg2.Rows = 12 Then Exit Sub
        Fg2.RemoveItem Fg2.Row
        
    End If
End Sub

Private Sub Fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    
    If Button = 2 Then
        PopupMenu menu2
    End If
End Sub

Private Sub Fg3_EnterCell()
    If QueHace = 3 Then
        Fg3.Editable = flexEDNone
        Exit Sub
    End If
    
    If (Fg3.Col = 1 Or Fg3.Col = 3 Or Fg3.Col = 2) Then
        Fg3.Editable = flexEDKbdMouse
    Else
        Fg3.Editable = flexEDNone
    End If
End Sub

Private Sub Fg3_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Then KeyAscii = 0
    If KeyAscii = 13 Then Exit Sub
    
    Select Case Col
        Case 1, 2, 3
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDOS E CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        
        mMesActivo = xMes
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        pCargarGrid
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then '--F3 Nuevo
        If fCierrePeriodo = False Then Exit Sub
        If QueHace <> 3 Then Exit Sub
        Nuevo
    End If
    
    If KeyCode = 115 Then '--F4 Modificar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace <> 3 Then Exit Sub
        Modificar
    End If
    
    If KeyCode = 113 Then '--F2 Grabar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace = 3 Then Exit Sub
        If Grabar = True Then
            QueHace = 3
            Set RstVent = Nothing
            Unload Me
        End If
    End If
    
    If KeyCode = 116 Then '--F5 actualizar
    End If
    
    If KeyCode = 117 Then '--F6 '--cancelar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace = 3 Then Exit Sub
        Cancelar
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    QueHace = 3
    TabOne1.CurrTab = 0
    SeEjecuto = False
    iniciarCampos
End Sub

Private Sub iniciarCampos()
    Dg1.Columns("fchemi1").NumberFormat = FORMAT_DATE
    CaracteresNumericos = "0123456789." & Chr(8)
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    Fg1.ColWidth(5) = 0
    Fg1.ColWidth(6) = 0
    GRID_COMBOLIST Fg1, 2
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.ColWidth(4) = 0
    Fg2.ColWidth(6) = 0
    Fg2.ColEditMask(1) = "##/##/####"
    TxtFchDoc.Valor = Date
    TxtFchDoc.Valor = ""
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
'    Fg3.ColWidth(4) = 0
'    Fg3.ColWidth(5) = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando una compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub menu1_1_Click()
    CmdAddItem_Click
End Sub

Private Sub menu1_3_Click()
    CmdDelItem_Click
End Sub

Private Sub menu2_1_Click()
    Fg2.AddItem "", Fg2.Rows - 2
End Sub

Private Sub menu2_3_Click()
    Dim fil As Long
    
    fil = Fg2.Row
    
    RstEntr.Filter = "corr=" & NulosN(Fg2.TextMatrix(fil, 4))
    If RstEntr.RecordCount <> 0 Then
        RstEntr.Delete
        RstEntr.Filter = adFilterNone
    End If
    Fg2.RemoveItem fil
    If Fg2.Rows < 12 Then Fg2.AddItem "", Fg2.Rows - 2
    
    generarTotales 3
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If mMesActivo = 0 Then Cancel = 1: Exit Sub
        ' Validamos si la cuadricula tiene datos
        If QueHace = 3 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No existe información para visualizar", vbInformation, Me.Caption
                Blanquea
                Exit Sub
            ElseIf NulosC(RstVent("nombre")) = "ANULADO" Then
                MsgBox "El Documento de Venta esta Anulado", vbInformation, Me.Caption
                Cancel = 1
                Exit Sub
            Else
                MuestraSegundoTab
            End If
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then
        If RstVent.RecordCount = 0 Then
            MsgBox "No se han registardos ventas para realizar esta opcion", vbInformation, Me.Caption
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.Index = 3 Then
        If RstVent.RecordCount = 0 Then
            MsgBox "No se han registardos ventas para realizar esta opcion", vbInformation, Me.Caption
            Exit Sub
        End If
        ' Validamos si el documento esta anulado
        If RstVent("Anulado") = -1 Then
            MsgBox RstVent("nomdoc") & " ya fue anulado, seleccione otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Anular
    End If
        
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstVent.Requery
            Dg1.Refresh
            If RstVent.RecordCount <> 0 Then
                RstVent.MoveFirst
                RstVent.Find "id=" & mIdRegistro
                If RstVent.EOF = True Then RstVent.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstVent.Filter = ""
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 11 Then CambiarMes
    
    If Button.Index = 13 Then
        Imprimir
    End If
    
    If Button.Index = 15 Then
        Set RstVent = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : OpcionesPeriodo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LAS OPCIONES DE EDICION DEL FORMULARIO EN FUNCION AL MES DE TRABAJO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub OpcionesPeriodo()
    Dim NomMes As String
    Dim Cerrado As Boolean
    Dim xFechaMes As String
    Dim xFchIni, xFchFin As Date
    Dim rpta As Integer
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    Cerrado = Busca_Codigo(mMesActivo, "id", "cerrado", "con_meses", "N", xCon)
    
    If Cerrado = True Then
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = False
        Toolbar1.Buttons(3).Visible = False
        Toolbar1.Buttons(4).Visible = False
    Else
        Toolbar1.Buttons(1).Visible = True
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(3).Visible = True
        Toolbar1.Buttons(4).Visible = True
    End If
    If mMesActivo <> 0 And mMesActivo <> 13 Then
        xFechaMes = "01/" + Trim(Format(mMesActivo, "00")) + "/" + Trim(Format(Year(Date), "0000"))
        xFchIni = xFechaMes
        xFchFin = Format(HallaDiasMes(CDate(xFechaMes)), "00") + "/" + Mid(xFechaMes, 4, 7)
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        ' MODIFICACION DE DOCUMENTOS
        If ButtonMenu.Index = 1 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados Pedidos para realizar esta opción", vbInformation, Me.Caption
                Exit Sub
            End If
            
            If RstVent("anulado") = -1 Then
                MsgBox "No puede modificar " & RstVent("nomdoc") & " anulado proceda a restaurarlo", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
                Exit Sub
            Else
                Modificar
            End If
        End If
        
        ' RESTAURAR DOCUMENTOS
        If ButtonMenu.Index = 2 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados Pedidos para realizar esta opción", vbInformation, Me.Caption
                Exit Sub
            End If
            If RstVent("anulado") = -1 Then ' SI EL DOCUMENTO ESTA ANULADO
                RestaurarFactura
            End If
        End If
    End If
  
    If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados Pedidos para realizar esta opción", vbInformation, Me.Caption
                Exit Sub
            End If
            Anular
        End If
        If ButtonMenu.Index = 2 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados Pedidos para realizar esta opción", vbInformation, Me.Caption
                Exit Sub
            End If
            
            Eliminar
        End If
        
        If ButtonMenu.Index = 3 Then EmitirAnulada
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : EmitirAnulada
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EMITIR UN PEDIDO COMO ANULADO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub EmitirAnulada()
    TabOne1.CurrTab = 0
    ActivarEntorno
    
    Fraseldoc.Left = 3315
    Fraseldoc.Top = 2505
    TxtAlmacen2.Text = ""
    TxtIdDocGen.Text = ""
    TxtNumSerGen.Text = ""
    TxtNumDocGen.Text = ""
    LblIdDocumentoGen.Caption = ""
    Fraseldoc.Visible = True
    TxtAlmacen2.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : RestaurarFactura
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : RESTAURA UNA FACTURA ANULADA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub RestaurarFactura()
    ' Se restaura una factura anulada
    Dim rpta As Integer
    
    rpta = MsgBox("Esta seguro de restaurar el Documento Nº " + RstVent("numser") & "-" & RstVent("numdoc"), vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption)
    If rpta = vbYes Then
        xCon.Execute "UPDATE ped_pedido SET ped_pedido.Anulado = 0 " _
            & " WHERE ped_pedido.id =" & RstVent("id") & ""

        ' Grabamos el movimiento en la tabla var_edicion
        GrabarOperacion xIdUsuario, IdMenuActivo, 2, Time, Time, Date, xCon, NulosN(RstVent("id"))
        
        RstVent.Requery
        Dg1.Refresh
        MsgBox "El documento se restauro con exito", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
    End If
End Sub

Private Sub TxtFchEnt_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosN(TxtTipPed.Text) = 1 Then
        
    End If
End Sub

Private Sub txtglosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub TxtIdAlm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdAlm_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusAlm_Click
    End If
End Sub

Private Sub TxtIdAlm_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtIdAlm.Text) <> "" Then
        LblAlmacen.Caption = Busca_Codigo(NulosN(TxtIdAlm.Text), "id", "descripcion", "alm_almacenes", "N", xCon)
        If LblAlmacen.Caption = "" Then
            TxtIdAlm.Text = ""
        End If
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtNumDoc.Text) <> "" Then
        TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
        
        Dim Rst As New ADODB.Recordset
        Dim nSQL As String
        ' ver si existe el numero de doc
        If QueHace <> 1 Then nSQL = " and ped_pedido.id <> " & NulosN(RstVent("id"))
        
        RST_Busq Rst, "SELECT ped_pedido.numser, ped_pedido.numdoc, ped_pedido.fchemi, mae_cliente.nombre, Left([ped_pedido].[numreg],2) " _
            & " & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & Right([ped_pedido].[numreg],4) AS registro " _
            & " FROM mae_cliente RIGHT JOIN (ped_pedido LEFT JOIN mae_libros ON ped_pedido.idlib = mae_libros.id) ON mae_cliente.id = ped_pedido.idcli " _
            & " WHERE (((ped_pedido.numser)='" & Trim(TxtNumSer.Text) & "') AND ((ped_pedido.numdoc)='" & TxtNumDoc.Text & "') AND ((ped_pedido.tipdoc)=" & NulosN(TxtTipDoc.Text) & "))", xCon

        If Rst.RecordCount <> 0 Then
            ' poner el nuevo numero doc
            MsgBox "El número de documento ya existe " & vbCr & "Nº Registro: " & NulosC(Rst("registro")) & vbCr & "Fecha Doc.   " & NulosC(Rst("fchemi")) & vbCr & "Cliente:         " & NulosC(Rst("nombre")) & vbCr & "Será reemplazado por " + Trim(TxtNumSer.Text) + "-" + Trim(TxtNumDoc.Text), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
        Set Rst = Nothing
    End If
End Sub

Private Sub TxtNumRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumRuc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCli_Click
    End If
End Sub

Private Sub TxtNumRuc_Validate(Cancel As Boolean)
    If NulosC(TxtNumRuc.Text) = "" Then
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    RST_Busq xRs1, "SELECT * FROM mae_cliente WHERE numruc like '" & TxtNumRuc.Text & "%' ORDER BY numruc", xCon
    If xRs1.RecordCount <> 0 Then
        TxtNumRuc.Text = xRs1("numruc")
        LblNomCli.Caption = xRs1("nombre")
        LblIdCliente.Caption = xRs1("id")
    Else
        TxtNumRuc.Text = ""
        LblNomCli.Caption = ""
        LblIdCliente.Caption = ""
    End If
    Set xRs1 = Nothing
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If NulosC(TxtTipDoc.Text) = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.Text = ""
        TxtTipDoc.SetFocus
        TxtNumSer.Text = ""
        Exit Sub
    End If
        
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumSer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusNumSer_Click
    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    Dim Rstdoc As New ADODB.Recordset
    If NulosC(TxtNumSer.Text) = "" Then
        Exit Sub
    Else
        If QueHace <> 1 Then Exit Sub
        
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
        Dim Rst As New ADODB.Recordset
        
        RST_Busq Rst, "SELECT top 1 numdoc AS numero from ped_pedido  WHERE numser ='" & NulosC(TxtNumSer.Text) & "' AND tipdoc =" & NulosN(TxtTipDoc) & " ORDER BY numdoc DESC ", xCon

        If Rst.RecordCount = 0 Then
            TxtNumDoc.Text = "0000000001"
        Else
            Rst.MoveFirst
            TxtNumDoc.Text = Format(NulosN(Rst("numero")) + 1, "0000000000")
        End If
        Set Rst = Nothing
    End If
End Sub

Private Sub TxtOC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtPtoVta_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    Dim xSql As String
    Dim Rst As New ADODB.Recordset
    
    If NulosN(TxtPtoVta.Text) = 0 Then
        TxtPtoVta.Text = ""
        LblPtoVta.Caption = ""
        Exit Sub
    End If
    If NulosN(LblIdCliente.Caption) = 0 Then
        TxtPtoVta.Text = ""
        MsgBox "No ha especificado el cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    xSql = "SELECT vta_puntoVenta.id, vta_puntoVenta.idcli, mae_cliente.nombre, mae_cliente.numruc,  vta_puntoVenta.descripcion FROM vta_puntoVenta LEFT JOIN mae_cliente ON vta_puntoVenta.idcli = mae_cliente.id " _
        & " WHERE (((vta_puntoVenta.id)=" & NulosN(TxtPtoVta.Text) & ") AND ((vta_puntoVenta.idcli)=" & NulosN(LblIdCliente.Caption) & "))"

    Set Rst = BuscaConCriterio(xSql, xCon)
    If Rst.RecordCount <> 0 Then
        LblPtoVta.Caption = Rst("descripcion")
    Else
        TxtPtoVta.Text = ""
        LblPtoVta.Caption = ""
    End If
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDoc_Click
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA ped_pedido, ESTA FUNCION DEVUELVE VERDADERO CUANDO
'*                    TIENE EXITO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim A As Integer
    Dim xTot As Long
    
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Function
    End If
    
    If TxtNumRuc.Text = "" Then
        MsgBox "No ha especificado cliente de la venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Function
    End If
    
    If TxtNumSer.Text = "" Or TxtNumDoc.Text = "" Then
        MsgBox "No ha especificado el numero del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        Exit Function
    End If
    
    If TxtFchDoc.Valor = "" Then
        MsgBox "No ha especificado la fecha de emision del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchDoc.SetFocus
        Exit Function
    End If
    
    If Month(TxtFchDoc.Valor) > mMesActivo Then
        MsgBox "No se puede grabar este documento en el periodo actual la fecha de emision es mayor al Perido actual", vbInformation + vbOKOnly + vbDefaultButton1
        TxtFchDoc.SetFocus
        Exit Function
    End If

    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado items para la liquidacion de gastos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    ' verificando el detalle
    RstEntr.Filter = "" ' quitando el filtro al rst para hacer las evaluaciones
    
    For A = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(A, 4)) = 0 Then
            MsgBox "No se le ha asignado una cantidad para al item : " & Chr(13) _
                & Fg1.TextMatrix(A, 2) & Chr(13) _
                & "Asignele una cantidad, luego proceda a asignale una fecha de entrega", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        Else
            If NulosN(TxtTipPed.Text) = 2 Then
                xTot = RstRegistroSumar(RstEntr, "canpro", "iditem", Fg1.TextMatrix(A, 5), "N", True)
                
                If xTot <> NulosN(Fg1.TextMatrix(A, 4)) Then
                    MsgBox "La cantidad del detalle de entrega es diferente a la cantidad del pedido para el item :" & Chr(13) _
                    & Fg1.TextMatrix(A, 2) & Chr(13) _
                    & "Corrija las cantidades ", vbExclamation, xTitulo
                    Exit Function
                End If
            End If
        End If
    Next A
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDet1 As New ADODB.Recordset
    Dim xId As Double
    Dim nSQL As String
    
    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        ' SI ES UN NUEVO REGISTRO OBTENEMOS EL ULTIMO ID DE LA TABLA ped_pedido
        xId = HallaCodigoTabla("ped_pedido", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM ped_pedido", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        ' SI SE ESTA MOFIGICANDO UN REGISTRO OBTENEMOS EL ID DEL REGISTRO ACTUAL
        xId = RstVent("id")
        RST_Busq RstCab, "SELECT * FROM ped_pedido WHERE id = " & xId & "", xCon
        ' Eliminamos el detalle
        xCon.Execute "DELETE * FROM ped_pedidodetent WHERE idped  = " & xId & ""
        xCon.Execute "DELETE * FROM ped_pedidodet WHERE idped  = " & xId & ""
    End If
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM ped_pedidodet", xCon
    RST_Busq RstDet1, "SELECT TOP 1 * FROM ped_pedidodetent", xCon
    
    mIdRegistro = xId
    
    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("idcli") = NulosN(LblIdCliente.Caption)
    RstCab("idpunvecli") = NulosN(TxtPtoVta.Text)
    RstCab("numser") = TxtNumSer.Text
    RstCab("numdoc") = TxtNumDoc.Text
    RstCab("fchemi") = CDate(TxtFchDoc.Valor)
    RstCab("idtipped") = NulosN(TxtTipPed.Text)
    If NulosN(TxtTipPed.Text) = 1 And IsDate(TxtFchEnt.Valor) = True Then
        RstCab("fchent") = CDate(TxtFchEnt.Valor)
    End If
    RstCab("idconpag") = NulosN(TxtConPag.Text)
    RstCab("glosa") = NulosC(txtglosa.Text)
    RstCab("oc") = NulosC(TxtOC.Text)
    RstCab.Update
    
    ' Detalle
    Dim idPed As Integer
    Dim idpeddet As Integer
    Dim idItem As Integer
    For A = 1 To Fg1.Rows - 1
        idPed = xId
        idpeddet = HallaCodigoTabla("ped_pedidodet", xCon, "idpeddet")
        idItem = NulosN(Fg1.TextMatrix(A, 5))
        If NulosN(TxtTipPed.Text) = 2 Then
            RstEntr.Filter = "iditem=" & idItem
            If RstEntr.RecordCount <> 0 Then
                RstEntr.MoveFirst
                While Not RstEntr.EOF
                    If IsDate(RstEntr("fchent")) = True And RstEntr("canpro") <> 0 Then
                        RstDet.AddNew
                        RstDet("idped") = idPed
                        RstDet("idpeddet") = idpeddet
                        RstDet("iditem") = idItem
                        RstDet("idunimed") = NulosN(Fg1.TextMatrix(A, 6))
                        RstDet("canpro") = NulosN(RstEntr("canpro"))
                        RstDet("canproent") = 0
                        RstDet("fchent") = CDate(RstEntr("fchent"))
                        If NulosN(RstEntr("canproent")) <> 0 Then
                            RstDet("estado") = 1 ' entregado
                        Else
                            RstDet("estado") = 2 ' pendiente
                        End If
                        RstDet("observacion") = ""
                        
                        RstDet.Update
                    End If
                    idpeddet = idpeddet + 1
                    RstEntr.MoveNext
                Wend
            End If
        Else
            RstDet.AddNew
            RstDet("idped") = xId
            RstDet("idpeddet") = idpeddet
            RstDet("iditem") = idItem
            RstDet("idunimed") = NulosN(Fg1.TextMatrix(A, 6))
            RstDet("canpro") = NulosN(Fg1.TextMatrix(A, 4))
            RstDet("canproent") = 0
            RstDet("fchent") = CDate(TxtFchEnt.Valor)
            RstDet("estado") = 2
            RstDet.Update
        End If
    idpeddet = idpeddet + 1
    Next A
        
    ' Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
   
    xCon.CommitTrans
    MsgBox "La " & Trim(LblNomDoc) & " se registró con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDet1 = Nothing
    Grabar = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDet1 = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function


'Function Grabar2() As Boolean
'    Dim A As Integer
'    Dim xTot As Long
'
'    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
'    If TxtTipDoc.Text = "" Then
'        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtTipDoc.SetFocus
'        Exit Function
'    End If
'
'    If TxtNumRuc.Text = "" Then
'        MsgBox "No ha especificado cliente de la venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtNumRuc.SetFocus
'        Exit Function
'    End If
'
'    If TxtNumSer.Text = "" Or TxtNumDoc.Text = "" Then
'        MsgBox "No ha especificado el numero del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtNumSer.SetFocus
'        Exit Function
'    End If
'
'    If TxtFchDoc.Valor = "" Then
'        MsgBox "No ha especificado la fecha de emision del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtFchDoc.SetFocus
'        Exit Function
'    End If
'
'    If Month(TxtFchDoc.Valor) > mMesActivo Then
'        MsgBox "No se puede grabar este documento en el periodo actual la fecha de emision es mayor al Perido actual", vbInformation + vbOKOnly + vbDefaultButton1
'        TxtFchDoc.SetFocus
'        Exit Function
'    End If
'
'    If Fg1.Rows = 1 Then
'        MsgBox "No ha especificado items para la liquidacion de gastos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Fg1.SetFocus
'        Exit Function
'    End If
'
'    ' verificando el detalle
'    RstEntr.Filter = "" ' quitando el filtro al rst para hacer las evaluaciones
'
'    For A = 1 To Fg1.Rows - 1
'        If NulosN(Fg1.TextMatrix(A, 4)) = 0 Then
'            MsgBox "No se le ha asignado una cantidad para al item : " & Chr(13) _
'                & Fg1.TextMatrix(A, 2) & Chr(13) _
'                & "Asignele una cantidad, luego proceda a asignale una fecha de entrega", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            Exit Function
'        Else
'            If NulosN(TxtTipPed.Text) = 2 Then
'                xTot = RstRegistroSumar(RstEntr, "canpro", "iditem", Fg1.TextMatrix(A, 5), "N", True)
'
'                If xTot <> NulosN(Fg1.TextMatrix(A, 4)) Then
'                    MsgBox "La cantidad del detalle de entrega es diferente a la cantidad del pedido para el item :" & Chr(13) _
'                    & Fg1.TextMatrix(A, 2) & Chr(13) _
'                    & "Corrija las cantidades ", vbExclamation, xTitulo
'                    Exit Function
'                End If
'            End If
'        End If
'    Next A
'
'    Dim RstCab As New ADODB.Recordset
'    Dim RstDet As New ADODB.Recordset
'    Dim RstDet1 As New ADODB.Recordset
'    Dim xId As Double
'    Dim nSQL As String
'
'    On Error GoTo LaCague
'
'    xCon.BeginTrans
'
'    If QueHace = 1 Then
'        ' SI ES UN NUEVO REGISTRO OBTENEMOS EL ULTIMO ID DE LA TABLA ped_pedido
'        xId = HallaCodigoTabla("ped_pedido", xCon, "id")
'        RST_Busq RstCab, "SELECT TOP 1 * FROM ped_pedido", xCon
'        RstCab.AddNew
'        RstCab("id") = xId
'    Else
'        ' SI SE ESTA MOFIGICANDO UN REGISTRO OBTENEMOS EL ID DEL REGISTRO ACTUAL
'        xId = RstVent("id")
'        RST_Busq RstCab, "SELECT * FROM ped_pedido WHERE id = " & xId & "", xCon
'        ' Eliminamos el detalle
'        xCon.Execute "DELETE * FROM ped_pedidodetent WHERE idped  = " & xId & ""
'        xCon.Execute "DELETE * FROM ped_pedidodet WHERE idped  = " & xId & ""
'    End If
'
'    RST_Busq RstDet, "SELECT TOP 1 * FROM ped_pedidodet", xCon
'    RST_Busq RstDet1, "SELECT TOP 1 * FROM ped_pedidodetent", xCon
'
'    mIdRegistro = xId
'
'    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
'    RstCab("idcli") = NulosN(LblIdCliente.Caption)
'    RstCab("idpunvecli") = NulosN(TxtPtoVta.Text)
'    RstCab("numser") = TxtNumSer.Text
'    RstCab("numdoc") = TxtNumDoc.Text
'    RstCab("fchemi") = CDate(TxtFchDoc.Valor)
'    RstCab("idtipped") = NulosN(TxtTipPed.Text)
'    If NulosN(TxtTipPed.Text) = 1 And IsDate(TxtFchEnt.Valor) = True Then
'        RstCab("fchent") = CDate(TxtFchEnt.Valor)
'    End If
'    RstCab("idconpag") = NulosN(TxtConPag.Text)
'    RstCab("glosa") = NulosC(txtglosa.Text)
'    RstCab("oc") = NulosC(TxtOC.Text)
'    RstCab.Update
'
'    ' Detalle
'    Dim idPed As Integer
'    Dim idItem As Integer
'
'    For A = 1 To Fg1.Rows - 1
'        RstDet.AddNew
'        RstDet("idped") = xId
'        RstDet("iditem") = NulosN(Fg1.TextMatrix(A, 5))
'        RstDet("idunimed") = NulosN(Fg1.TextMatrix(A, 6))
'        RstDet("canpro") = NulosN(Fg1.TextMatrix(A, 4))
'        RstDet("observacion") = ""
'        RstDet.Update
'
'        ' agregando las entregas
'        ' si es compuesto
'        If NulosN(TxtTipPed.Text) = 2 Then
'            RstEntr.Filter = "iditem=" & Fg1.TextMatrix(A, 5)
'            If RstEntr.RecordCount <> 0 Then
'                RstEntr.MoveFirst
'                Do While Not RstEntr.EOF
'                    If IsDate(RstEntr("fchent")) = True And RstEntr("canpro") <> 0 Then
'                        RstDet1.AddNew
'                        RstDet1("idped") = xId
'                        RstDet1("iditem") = NulosN(Fg1.TextMatrix(A, 5))
'                        RstDet1("idunimed") = NulosN(Fg1.TextMatrix(A, 6))
'                        RstDet1("canpro") = NulosN(RstEntr("canpro"))
'                        RstDet1("fchent") = CDate(RstEntr("fchent"))
'                        If NulosN(RstEntr("entregado")) <> 0 Then
'                            RstDet1("estado") = 1 ' entregado
'                        Else
'                            RstDet1("estado") = 2 ' pendiente
'                        End If
'
'                        RstDet1.Update
'                    End If
'                    RstEntr.MoveNext
'                Loop
'            End If
'        Else
'            ' si es simple
'            RstDet1.AddNew
'            RstDet1("idped") = xId
'            RstDet1("iditem") = NulosN(Fg1.TextMatrix(A, 5))
'            RstDet1("idunimed") = NulosN(Fg1.TextMatrix(A, 6))
'            RstDet1("canpro") = NulosN(Fg1.TextMatrix(A, 4))
'            RstDet1("fchent") = CDate(TxtFchEnt.Valor)
'            RstDet1("estado") = 2
'            RstDet1.Update
'        End If
'    Next
'    ' Grabamos el movimiento en la tabla var_edicion
'    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
'
'    xCon.CommitTrans
'    MsgBox "La " & Trim(LblNomDoc) & " se registró con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'    Set RstCab = Nothing
'    Set RstDet = Nothing
'    Set RstDet1 = Nothing
'    Grabar = True
'    Exit Function
'LaCague:
'    xCon.RollbackTrans
'    Set RstCab = Nothing
'    Set RstDet = Nothing
'    Set RstDet1 = Nothing
'    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
'    Grabar = False
'    Exit Function
'End Function

'*****************************************************************************************************
'* Nombre           : HallaNumAsiento
'* Tipo             : FUNCION
'* Descripcion      :
'* Paranetros       : NOMBRE    |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Mes       |  Mes         |  ESPECIFICA EL MES DE TRABAJO ACTUAL
'* Devuelve         :
'*****************************************************************************************************
Function HallaNumAsiento(Mes As Integer) As String
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT con_diario.idmes, con_diario.idlib, con_diario.numasi From con_diario " _
        & " WHERE (((con_diario.idmes)=" & Mes & ") AND ((con_diario.idlib)=41)) ORDER BY numasi", xCon
    
    If Rst.RecordCount = 0 Then
        HallaNumAsiento = "0001"
    Else
        Rst.MoveLast
        HallaNumAsiento = Format(NulosN(Rst("numasi")) + 1, "0000")
    End If
    Exit Function
End Function

Private Sub TxtTipDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtTipDoc.Text) = "" Then
        Exit Sub
    End If
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
        
    RST_Busq xRs, "SELECT mae_documento.* FROM MAE_documento WHERE id = " & NulosN(Me.TxtTipDoc) & "", xCon
        
    If xRs.RecordCount = 0 Then
        TxtTipDoc.Text = ""
        LblNomDoc.Caption = ""
    Else
        TxtTipDoc.Text = xRs("id")
        LblNomDoc.Caption = xRs("descripcion")
        Set xRs2 = Nothing
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Filtrar
'* Tipo             : FUNCION
'* Descripcion      : EJECUTA UN FILTRO EN EL RECORDSET RstVent
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'* Devuelve         :
'*****************************************************************************************************
Sub Filtrar()
    TabOne1.CurrTab = 0
    Dim xform As New eps_librerias.FormFiltrar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(7, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Cliente":         xCampos(0, 1) = "nombre":        xCampos(0, 2) = "C":         xCampos(0, 3) = "1500"
    xCampos(1, 0) = "Fch. Emision":    xCampos(1, 1) = "fchdoc":        xCampos(1, 2) = "F":         xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Nº Documento":    xCampos(2, 1) = "numerodoc":     xCampos(2, 2) = "C":         xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Tipo Documento":  xCampos(3, 1) = "abrev":         xCampos(3, 2) = "C":         xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Forma de Pago":   xCampos(4, 1) = "desccond":      xCampos(4, 2) = "C":         xCampos(4, 3) = "1500"
    xCampos(5, 0) = "Moneda":          xCampos(5, 1) = "simbolo":       xCampos(5, 2) = "C":         xCampos(5, 3) = "1500"
    xCampos(6, 0) = "Estado":          xCampos(6, 1) = "estadoventa":   xCampos(6, 2) = "C":         xCampos(6, 3) = "1500"
    xCampos(7, 0) = "importe":         xCampos(7, 1) = "imptotdoc":     xCampos(7, 2) = "N":         xCampos(7, 3) = "1500"
    
    Set xform.Coneccion = xCon
    Set xform.Rst = RstVent
    Set RstVent = xform.FiltrarReg(xCampos)
    Set Dg1.DataSource = RstVent
    Dg1.Refresh
End Sub

Sub Imprimir()
    
End Sub

'*****************************************************************************************************
'* Nombre           : CambiarMes
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITE CAMBIAR EL MES DE TRABAJO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub CambiarMes()
    TabOne1.CurrTab = 0
    mMesActivo = SeleccionaMes(xCon)
    pCargarGrid
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS DE LA TABLA ped_pedido EN EL CONTROL Dg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarGrid()
    Dim nSQL  As String
    Dim rpta As Integer
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    LblPeriodo2.Caption = LblMes.Caption
    
    TDB_FiltroLimpiar Dg1
    Set RstVent = Nothing
    
    ' bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
        
    If mMesActivo = 0 Or mMesActivo = 13 Then
        MsgBox "Ha selecionado el mes de Cierre, Seleccione meses comprendidos entre Enero y Diciembre", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstVent = Nothing
        Set Dg1.DataSource = Nothing
        Dg1.Refresh
        Exit Sub
    Else
        nSQL = "SELECT ped_pedido.*, IIf(ped_pedido.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, " _
            & " IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc) AS numerodoc, mae_documento.descripcion AS nomdoc, " _
            & " mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, mae_condpago.abrev AS conpagabre, ped_pedido.fchemi & '' AS fchemi1, " _
            & " vta_puntoVenta.descripcion AS ptovta, ped_tipo.descripcion AS tipped FROM ped_tipo RIGHT JOIN (mae_documento INNER JOIN (mae_cliente " _
            & " RIGHT JOIN (vta_puntoVenta RIGHT JOIN (ped_pedido LEFT JOIN mae_condpago ON ped_pedido.idconpag = mae_condpago.id) ON vta_puntoVenta.id = ped_pedido.idpunvecli) " _
            & " ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) ON ped_tipo.id = ped_pedido.idtipped " _
            & " Where (((Month([ped_pedido].[fchemi])) = " & mMesActivo & ") And ((Year([ped_pedido].[fchemi])) = " & Val(AP_AÑODAT) & ")) " _
            & " ORDER BY ped_pedido.fchemi DESC , IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc)"
    End If
    
    ' cargando datos
    Me.MousePointer = vbHourglass
    RST_Busq RstVent, nSQL, xCon
    Set Dg1.DataSource = RstVent
    Me.MousePointer = vbDefault
    OpcionesPeriodo
    TabOne1.CurrTab = 0
    
    If RstVent.State = 0 Then Exit Sub
    If RstVent.RecordCount = 0 Then
        rpta = MsgBox("No se ha registrado ninguna operacion, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
        If rpta = vbYes Then
            Nuevo
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UNA BUSQUEDA EN EL RECORDSET RstVent
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    TabOne1.CurrTab = 0
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim xCampos(8, 4) As String
    
    xCampos(0, 0) = "T.D.":           xCampos(0, 1) = "abretipdoc":  xCampos(0, 2) = "400":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "N°. Documento":  xCampos(1, 1) = "numerodoc":   xCampos(1, 2) = "1400":  xCampos(1, 3) = "C"
    xCampos(2, 0) = "N°. OC Cliente": xCampos(2, 1) = "oc":          xCampos(2, 2) = "1300":  xCampos(2, 3) = "C"
    xCampos(3, 0) = "FchEmi":         xCampos(3, 1) = "fchemi":      xCampos(3, 2) = "800":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "Cliente":        xCampos(4, 1) = "nombre":      xCampos(4, 2) = "2000":  xCampos(4, 3) = "C"
    xCampos(5, 0) = "Producto":       xCampos(5, 1) = "descripcion": xCampos(5, 2) = "3500":  xCampos(5, 3) = "C"
    xCampos(6, 0) = "Uni. Med":       xCampos(6, 1) = "abreunimed":  xCampos(6, 2) = "1200":  xCampos(6, 3) = "C"
    xCampos(7, 0) = "Cantidad":       xCampos(7, 1) = "canpro":      xCampos(7, 2) = "1200":  xCampos(7, 3) = "N"
    
    nSQL = "SELECT ped_pedido.id, IIf(ped_pedido.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, ped_pedido!numser+'-'+ped_pedido!numdoc AS numerodoc, " _
        & " ped_pedido.oc, mae_documento.abrev AS abretipdoc, Format(ped_pedido.fchemi,'dd/mm/yy') AS fchemi, alm_inventario.descripcion, ped_pedidodet.canpro, " _
        & " mae_unidades.abrev AS abreunimed FROM (((mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN ped_pedido ON mae_cliente.id = ped_pedido.idcli) " _
        & " ON mae_documento.id = ped_pedido.tipdoc) LEFT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN alm_inventario " _
        & " ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id " _
        & " WHERE (((Month([ped_pedido].[fchemi]))=" & mMesActivo & ") AND ((ped_pedidodet.estado)<>1))"
  
    
    'SELECT ped_pedido.id, IIf(ped_pedido.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, ped_pedido!numser+'-'+ped_pedido!numdoc AS numerodoc, mae_documento.abrev, Format(ped_pedido.fchemi,'dd/mm/yy') AS fchemi " _
        + vbCr + " FROM mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN ped_pedido ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc " _
        + vbCr + " WHERE (((Month([ped_pedido].[fchemi]))=" & mMesActivo & ")); "

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Pedidos", "nombre", "nombre", Principio
    If xRs.State = 1 Then
        RstVent.MoveFirst
        RstVent.Find "id = " & xRs("id") & ""
    End If
    Set xRs = Nothing
End Sub

Sub ActualizaSaldoDoc(idDocumento As Integer, Tabla As Integer, ImporteRestar As Double)
'    '1 = compras
'    '2 = Ventas
'    '3 = honorarios
'
'    Dim Rst As New ADODB.Recordset
'    Dim Total As Double
'
'    If Tabla = 2 Then
'        RST_Busq Rst, "SELECT Sum(tes_cajadestinodet.acuenta) AS total FROM tes_caja LEFT JOIN tes_cajadestinodet ON tes_caja.id = tes_cajadestinodet.idtes " _
'            & " GROUP BY tes_cajadestinodet.iddoc, tes_caja.tipmov HAVING (((tes_cajadestinodet.iddoc)=" & idDocumento & ") AND ((tes_caja.tipmov)=1))", xCon
'
'        Total = BuscaImporteDocumento(idDocumento, 1)
'    End If
'
'    If Rst.RecordCount <> 0 Then
'        Total = ((Total - Rst("total")) - ImporteRestar)
'    Else
'        Total = (Total - ImporteRestar)
'    End If
'
'    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = " & Total & " WHERE (((vta_ventas.id)=" & idDocumento & "))"
'    Set Rst = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : BuscaImporteDocumento
'* Tipo             : FUNCION
'* Descripcion      :
'* Paranetros       : NOMBRE      |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    idDocumento |  Integer    |  ESPECIFICA EL ID DEL PEDIDO QUE SE DESEA CONSULTAR
'*                    Tabla       |  Integer    |  ESPECIFICA EL TIPO DE TABLA QUE SE CONSULTARA
'* Devuelve         :
'*****************************************************************************************************
Function BuscaImporteDocumento(idDocumento As Integer, Tabla As Integer) As Double
    Dim Rst As New ADODB.Recordset
    
    If Tabla = 1 Then RST_Busq Rst, "SELECT * FROM ped_pedido WHERE id = " & idDocumento & "", xCon
    
    If Rst.RecordCount <> 0 Then
        BuscaImporteDocumento = Rst("imptot")
    Else
        BuscaImporteDocumento = 0
    End If
    
    Set Rst = Nothing
End Function

Private Sub CmdBusCondicion_Click()
    ' EJECUTA LA BUSQUEDA DE UNA CONDICION DE PAGO
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_condpago ORDER BY descripcion"
    
    xform.Titulo = "Buscando Condicion de Pago"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtConPag.Text = xRs("id")
            LblCondPag.Caption = xRs("descripcion")
            TxtTipPed.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub TxtConPag_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtConPag_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCondicion_Click
    End If
End Sub

Private Sub TxtConPag_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtConPag.Text) = "" Then
        Exit Sub
    End If
    Dim xRs1 As New ADODB.Recordset

    RST_Busq xRs1, "SELECT * FROM mae_condpago WHERE id = " & NulosN(TxtConPag.Text) & "", xCon

    If xRs1.RecordCount = 0 Then
        TxtConPag.Text = ""
        LblCondPag.Caption = ""
        TxtFchEnt.Valor = ""
    Else
        If TxtFchDoc.Valor = "" Then
            MsgBox "No ha especificado la fecha de emision del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtConPag.Text = ""
            LblCondPag.Caption = ""
            Exit Sub
        End If
        LblCondPag.Caption = Trim(xRs1("descripcion"))
    End If
    Set xRs1 = Nothing
End Sub

Private Sub CmdTipPed_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM ped_tipo ORDER BY descripcion"
    
    xform.Titulo = "Buscando Tipo de Pedido"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipPed.Text = xRs("id")
            LblTipPed.Caption = NulosC(xRs("descripcion"))
            If xRs("id") = 1 Then
                Label3(3).Visible = True
                TxtFchEnt.Visible = True
                TxtFchEnt.SetFocus
            Else
                TxtFchEnt.Valor = ""
                Label3(3).Visible = False
                TxtFchEnt.Visible = False
                CmdAddItem.SetFocus
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub TxtTipPed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipPed_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdTipPed_Click
    End If
End Sub

Private Sub TxtTipPed_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosN(TxtTipPed.Text) = 0 Then
        Label3(3).Visible = False
        TxtFchEnt.Visible = False
        Exit Sub
    End If
    Dim xRs1 As New ADODB.Recordset

    RST_Busq xRs1, "SELECT * FROM ped_tipo WHERE id = " & NulosN(TxtTipPed.Text) & "", xCon
    If xRs1.State = 1 Then
        If xRs1.RecordCount = 0 Then
            TxtTipPed.Text = ""
            LblTipPed.Caption = ""
            TxtFchEnt.Valor = ""
        Else
            LblTipPed.Caption = Trim(xRs1("descripcion"))
            If xRs1("id") = 1 Then
                Label3(3).Visible = True
                TxtFchEnt.Visible = True
                TxtFchEnt.SetFocus
            Else
                TxtFchEnt.Valor = ""
                Label3(3).Visible = False
                TxtFchEnt.Visible = False
            End If
        End If
    End If
    Set xRs1 = Nothing
End Sub

Private Sub CmdBusPtoVta_Click()
    ' EJECUTA LA BUSQUEDA DE UN PUNTO DE VENTA
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 3) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":  xCampos(1, 1) = "id":            xCampos(1, 2) = "1400":    xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT VTA_PuntoVenta.idcli, VTA_PuntoVenta.codcen, VTA_PuntoVenta.descripcion, VTA_PuntoVenta.id, VTA_PuntoVenta.dir " _
        & " From VTA_PuntoVenta Where (((VTA_PuntoVenta.idcli) = " & NulosN(LblIdCliente.Caption) & " )) ORDER BY VTA_PuntoVenta.descripcion"

    xform.Titulo = "Buscando Punto de Venta"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            LblPtoVta.Caption = xRs("descripcion")
            TxtPtoVta.Text = xRs("id")
            TxtTipDoc.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub
  
Private Sub TxtPtoVta_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtPtoVta_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        LblPtoVta.Caption = ""
        TxtPtoVta.Text = ""
    End If
    
    If KeyCode = 116 Then
        CmdBusPtoVta_Click
    End If
End Sub

Sub PreparaRSTEntrega()
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(7, 3) As String

    xCampos(0, 0) = "iditem":         xCampos(0, 1) = "C":      xCampos(0, 2) = "10" ' codigo de item
    xCampos(1, 0) = "fchent":         xCampos(1, 1) = "F":      xCampos(1, 2) = "10" ' fecha de entrega
    xCampos(2, 0) = "porcen":        xCampos(2, 1) = "N":      xCampos(2, 2) = "5" ' porcentaje de entrega
    xCampos(3, 0) = "canpro":         xCampos(3, 1) = "N":      xCampos(3, 2) = "5" ' cantidad para entregar
    xCampos(4, 0) = "corr":           xCampos(4, 1) = "N":      xCampos(4, 2) = "2" ' correlativo(identificador de registros temporal)
    xCampos(5, 0) = "canproent":        xCampos(5, 1) = "N":      xCampos(5, 2) = "5" ' cantidad entregada
    xCampos(6, 0) = "idpeddet":        xCampos(6, 1) = "N":      xCampos(6, 2) = "5" ' identificador unico del pedido
        
    Set RstEntr = xFun.CrearRstTMP(xCampos)
    RstEntr.Open
End Sub

Private Sub cmdokseldoc_Click()
    If TxtIdDocGen.Text = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdDocGen.SetFocus
        Exit Sub
    End If
    
    If TxtNumSerGen.Text = "" Then
        MsgBox "No ha especificado el numero de serie del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSerGen.SetFocus
        Exit Sub
    End If
    
    If TxtNumDocGen.Text = "" Then
        MsgBox "No ha especificado el numero del documeto a generar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDocGen.SetFocus
        Exit Sub
    End If

    If Month(TxtFchEmiAnul.Valor) < mMesActivo Then
        MsgBox "La fecha del documento no corresponde la periodo especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If

    Dim RstCab As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim xId As Double
    Dim xNumAsiento As String

    RST_Busq xRs, "SELECT ped_pedido.tipdoc FROM ped_pedido " _
        & " WHERE (((ped_pedido.tipdoc)=" & NulosN(LblIdDocumentoGen.Caption) & ") AND ((ped_pedido.numser)='" & TxtNumSerGen.Text & "') AND " _
        & " ((ped_pedido.numdoc)='" & TxtNumDocGen.Text & "'))", xCon
    
    If xRs.RecordCount = 1 Then
        Set xRs = Nothing
        MsgBox "El numero de documento que quiere emitir ya existe", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDocGen.SetFocus
        Exit Sub
    End If
    
On Error GoTo LaCague
    xCon.BeginTrans
    ' Validar si el nro de documento existe solo en modo adicionar documento
    RST_Busq RstCab, "SELECT TOP 1 * FROM ped_pedido", xCon
    
    xId = HallaCodigoTabla("ped_pedido", xCon, "id")
    RstCab.AddNew
    RstCab("id") = xId
    RstCab("tipdoc") = NulosN(LblIdDocumentoGen.Caption)
    RstCab("idcli") = 1
    RstCab("numser") = TxtNumSerGen.Text
    RstCab("numdoc") = TxtNumDocGen.Text
    RstCab("Fchemi") = TxtFchEmiAnul.Valor
    RstCab("anulado") = -1
    RstCab("idconpag") = -1
    RstCab("idtipped") = 2
    RstCab.Update
            
    xCon.CommitTrans
        
    MsgBox "El documento anulado se genero con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    RstVent.Requery
    Dg1.Refresh
    cmdsalirseldoc_Click
    Exit Sub
    
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set xRs = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Private Sub cmdsalirseldoc_Click()
    ActivarEntorno
    Fraseldoc.Visible = False
End Sub

Private Sub CmdBusAlmacen2_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT alm_almacenes.* FROM alm_almacenes"
    
    xform.Titulo = "Buscando Almacenes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblidAlmacen2.Caption = xRs("id")
        TxtAlmacen2.Text = xRs("descripcion")
        TxtIdDocGen.SetFocus
        
        If TxtIdDocGen.Text <> "" Then
            Dim Rst As New ADODB.Recordset
            Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(LblIdDocumentoGen.Caption) & " AND idalm = " & NulosN(LblidAlmacen2.Caption) & "", xCon)
            If Rst.RecordCount <> 0 Then
                TxtNumSerGen.Text = Rst("numser")
            End If
            Set Rst = Nothing
        Else
            TxtNumSerGen.Text = ""
            TxtNumDocGen.Text = ""
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub TxtAlmacen2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtAlmacen2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusAlmacen2_Click
    End If
End Sub

Private Sub CmdBusTipDocGen_Click()
    ' EJECUTA LA BUSQUEFA DE UN DOCUMENTO CONTABLE
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"

    xform.SQLCad = " SELECT mae_documento.*, alm_numseries.numser" & _
        " FROM mae_documento LEFT JOIN alm_numseries ON mae_documento.id = alm_numseries.idtipdoc " & _
        " WHERE alm_numseries.idalm =" & NulosN(LblidAlmacen2.Caption)

    xform.Titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            LblIdDocumentoGen = xRs("id")
            TxtIdDocGen.Text = xRs("descripcion")
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub TxtIdDocGen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdDocGen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDocGen_Click
    End If
End Sub

Private Sub CmdBusSerGen_Click()
    ' EJECUTA LA BUSQUEDA DE UN NUMERO DE SERIE
    If TxtIdDocGen.Text = "" Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(4, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":         xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Abreviatura":    xCampos(1, 1) = "abrev":            xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Cod. Sunat":     xCampos(2, 1) = "codsun":           xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nº Serie":       xCampos(3, 1) = "numser":           xCampos(3, 2) = "1000":         xCampos(3, 3) = "C"
    
    xform.SQLCad = "SELECT alm_numseries.idalm, alm_numseries.idtipdoc, alm_numseries.numser, mae_documento.abrev, mae_documento.codsun, " _
        & " mae_documento.descripcion " _
        & " FROM alm_numseries LEFT JOIN mae_documento ON alm_numseries.idtipdoc = mae_documento.id " _
        & " WHERE (((alm_numseries.idalm)=" & NulosN(LblidAlmacen2.Caption) & ") AND ((alm_numseries.idtipdoc)=" & NulosN(LblIdDocumentoGen.Caption) & "))"
    
    xform.Titulo = "Buscando Series de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numser"
    xform.CampoBusca = "numser"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtNumSerGen.Text = xRs("numser")
        TxtNumDocGen.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub TxtNumSerGen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumSerGen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusSerGen_Click
    End If
End Sub

Private Sub TxtNumDocGen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        TxtNumDocGen_Validate True
    End If
End Sub

Private Sub TxtNumDocGen_Validate(Cancel As Boolean)
    If TxtNumDocGen.Text <> "" Then
        TxtNumDocGen.Text = Format(TxtNumDocGen.Text, "0000000000")
    End If
End Sub
