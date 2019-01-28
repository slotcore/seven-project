VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLetrascobrar 
   Caption         =   "Letras por Cobrar"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11940
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
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
            Picture         =   "FrmLetrascobrar.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLetrascobrar.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLetrascobrar.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLetrascobrar.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLetrascobrar.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLetrascobrar.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLetrascobrar.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLetrascobrar.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLetrascobrar.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLetrascobrar.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLetrascobrar.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLetrascobrar.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLetrascobrar.frx":277E
            Key             =   "IMG12"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   1005
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
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Letra"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Anular Letra"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar Letra"
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
            Object.ToolTipText     =   "Imprimir Guia"
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
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7620
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   11895
      _cx             =   20981
      _cy             =   13441
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
      CurrTab         =   1
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7200
         Left            =   45
         TabIndex        =   49
         Top             =   375
         Width           =   11805
         Begin VB.Frame Frame1 
            Height          =   6990
            Left            =   0
            TabIndex        =   5
            Top             =   240
            Width           =   11820
            Begin VB.CommandButton CmdBusProCli 
               Height          =   240
               Left            =   3270
               Picture         =   "FrmLetrascobrar.frx":2A98
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   3135
               Width           =   240
            End
            Begin VB.TextBox TxtNumDoc 
               Height          =   300
               Left            =   1905
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   29
               Text            =   "TxtNumDoc"
               Top             =   2040
               Width           =   1440
            End
            Begin VB.CommandButton CmdBuscaMovimiento 
               Height          =   240
               Left            =   2565
               Picture         =   "FrmLetrascobrar.frx":2BCA
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   1350
               Width           =   240
            End
            Begin VB.CommandButton CmdBusTipDoc 
               Height          =   240
               Left            =   2565
               Picture         =   "FrmLetrascobrar.frx":2CFC
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   1725
               Width           =   240
            End
            Begin VB.CommandButton CmdBusCondicion 
               Height          =   240
               Left            =   2565
               Picture         =   "FrmLetrascobrar.frx":2E2E
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   990
               Width           =   240
            End
            Begin VB.TextBox TxtConPag 
               Height          =   300
               Left            =   1905
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   17
               Text            =   "TxtConPag"
               Top             =   960
               Width           =   915
            End
            Begin VB.CommandButton CmdBusMon 
               Height          =   240
               Left            =   2565
               Picture         =   "FrmLetrascobrar.frx":2F60
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   615
               Width           =   240
            End
            Begin VB.TextBox TxtIdMoneda 
               Height          =   300
               Left            =   1905
               Locked          =   -1  'True
               TabIndex        =   11
               Text            =   "TxtIdMoneda"
               Top             =   600
               Width           =   915
            End
            Begin VB.TextBox TxtGlosa 
               Height          =   300
               Left            =   1905
               Locked          =   -1  'True
               TabIndex        =   31
               Text            =   "TxtGlosa"
               Top             =   2400
               Width           =   5850
            End
            Begin VB.TextBox TxtImporte 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1905
               Locked          =   -1  'True
               TabIndex        =   33
               Text            =   "TxtImporte"
               Top             =   2760
               Width           =   1275
            End
            Begin VB.TextBox TxtIdMov 
               Height          =   300
               Left            =   1905
               Locked          =   -1  'True
               TabIndex        =   21
               Text            =   "TxtIdMov"
               Top             =   1320
               Width           =   915
            End
            Begin VB.TextBox TxtNumRuc 
               Height          =   300
               Left            =   1905
               Locked          =   -1  'True
               MaxLength       =   11
               TabIndex        =   35
               Text            =   "TxtNumRuc"
               Top             =   3120
               Width           =   1620
            End
            Begin VB.CommandButton CmdAgregar 
               Caption         =   "&Agregar"
               Height          =   570
               Left            =   9405
               TabIndex        =   41
               Top             =   3840
               Width           =   1380
            End
            Begin VB.CommandButton CmdEliminar 
               Caption         =   "&Eliminar"
               Height          =   570
               Left            =   9360
               TabIndex        =   44
               Top             =   5640
               Width           =   1380
            End
            Begin VB.TextBox TxtTipDoc 
               Height          =   300
               Left            =   1905
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   25
               Text            =   "TxtTipDoc"
               Top             =   1680
               Width           =   915
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   1530
               Left            =   240
               TabIndex        =   40
               Top             =   3840
               Width           =   8985
               _cx             =   15849
               _cy             =   2699
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
               BackColorSel    =   -2147483635
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmLetrascobrar.frx":3092
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
            Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
               Height          =   300
               Left            =   1905
               TabIndex        =   7
               Top             =   240
               Width           =   1275
               _ExtentX        =   2249
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
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   1260
               Left            =   240
               TabIndex        =   43
               Top             =   5640
               Width           =   8985
               _cx             =   15849
               _cy             =   2222
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
               BackColorSel    =   -2147483635
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   11
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmLetrascobrar.frx":3195
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
            Begin AspaTextBoxFecha.TextBoxFecha Txtfechaven 
               Height          =   300
               Left            =   6480
               TabIndex        =   9
               Top             =   240
               Width           =   1275
               _ExtentX        =   2249
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
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nº de Documento"
               Height          =   195
               Index           =   15
               Left            =   360
               TabIndex        =   28
               Top             =   2040
               Width           =   1275
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Condicion de Pago"
               Height          =   195
               Index           =   16
               Left            =   360
               TabIndex        =   16
               Top             =   960
               Width           =   1350
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
               Left            =   2880
               TabIndex        =   19
               Top             =   960
               Width           =   4875
            End
            Begin VB.Label LblMoneda 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblMoneda"
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
               Left            =   2880
               TabIndex        =   13
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Index           =   4
               Left            =   360
               TabIndex        =   10
               Top             =   600
               Width           =   585
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Emision"
               Height          =   195
               Index           =   0
               Left            =   360
               TabIndex        =   6
               Top             =   240
               Width           =   1035
            End
            Begin VB.Label LblMovimiento 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblMovimiento"
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
               Left            =   2880
               TabIndex        =   23
               Top             =   1320
               Width           =   4875
            End
            Begin VB.Label lblDestipocambio 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Cambio"
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
               Height          =   195
               Left            =   5025
               TabIndex        =   14
               Top             =   660
               Width           =   1335
            End
            Begin VB.Label LblTipoCambio 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblTipoCambio"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   6480
               TabIndex        =   15
               Top             =   600
               Width           =   1275
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
               Left            =   3540
               TabIndex        =   38
               Top             =   3120
               Width           =   4215
            End
            Begin VB.Label LblidCli 
               Caption         =   "LblidCli"
               ForeColor       =   &H000000FF&
               Height          =   180
               Left            =   9240
               TabIndex        =   37
               Top             =   2880
               Visible         =   0   'False
               Width           =   690
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
               Left            =   2880
               TabIndex        =   27
               Top             =   1680
               Width           =   4875
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Venc."
               Height          =   195
               Index           =   1
               Left            =   5040
               TabIndex        =   8
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Movimiento"
               Height          =   195
               Index           =   2
               Left            =   360
               TabIndex        =   20
               Top             =   1320
               Width           =   810
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Documento"
               Height          =   195
               Index           =   3
               Left            =   360
               TabIndex        =   24
               Top             =   1680
               Width           =   825
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Glosa"
               Height          =   195
               Index           =   5
               Left            =   360
               TabIndex        =   30
               Top             =   2400
               Width           =   405
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Importe de Letra"
               Height          =   195
               Index           =   6
               Left            =   360
               TabIndex        =   32
               Top             =   2760
               Width           =   1155
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Cliente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   240
               Index           =   7
               Left            =   330
               TabIndex        =   34
               Top             =   3120
               Width           =   735
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Documentos Pendientes x Cobrar"
               Height          =   195
               Index           =   8
               Left            =   240
               TabIndex        =   39
               Top             =   3600
               Width           =   2370
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Documentos Canjeados"
               Height          =   195
               Index           =   9
               Left            =   240
               TabIndex        =   42
               Top             =   5400
               Width           =   1695
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Letras por cobrar"
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
            Left            =   45
            TabIndex        =   4
            Top             =   0
            Width           =   11745
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7200
         Left            =   -12450
         TabIndex        =   47
         Top             =   375
         Width           =   11805
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6570
            Left            =   0
            TabIndex        =   2
            Top             =   240
            Width           =   11805
            _ExtentX        =   20823
            _ExtentY        =   11589
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "T.D."
            Columns(0).DataField=   "abrev"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Moneda"
            Columns(1).DataField=   "simbolo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numerodoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Emi"
            Columns(3).DataField=   "fchdoc"
            Columns(3).NumberFormat=   "Short Date"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Cliente"
            Columns(4).DataField=   "nombre"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Forma Pago"
            Columns(5).DataField=   "desccond"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Importe"
            Columns(6).DataField=   "imptotdoc"
            Columns(6).NumberFormat=   "0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Fch. Ven."
            Columns(7).DataField=   "fchven"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Saldo"
            Columns(8).DataField=   "impsal"
            Columns(8).NumberFormat=   "0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Estado"
            Columns(9).DataField=   "EstadoVenta"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=900"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=820"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1402"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1323"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2566"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2487"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1773"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1693"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=4128"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=4048"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2090"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2011"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=1640"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1561"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=1773"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1693"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=1667"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1588"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(55)=   "Column(9).Width=1561"
            Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1482"
            Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=516"
            Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
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
            _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
            _StyleDefs(76)  =   "Named:id=33:Normal"
            _StyleDefs(77)  =   ":id=33,.parent=0"
            _StyleDefs(78)  =   "Named:id=34:Heading"
            _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(80)  =   ":id=34,.wraptext=-1"
            _StyleDefs(81)  =   "Named:id=35:Footing"
            _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(83)  =   "Named:id=36:Selected"
            _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=37:Caption"
            _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(87)  =   "Named:id=38:HighlightRow"
            _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=39:EvenRow"
            _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(91)  =   "Named:id=40:OddRow"
            _StyleDefs(92)  =   ":id=40,.parent=33"
            _StyleDefs(93)  =   "Named:id=41:RecordSelector"
            _StyleDefs(94)  =   ":id=41,.parent=34"
            _StyleDefs(95)  =   "Named:id=42:FilterBar"
            _StyleDefs(96)  =   ":id=42,.parent=33"
         End
         Begin VB.Label LblMes 
            AutoSize        =   -1  'True
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   10500
            TabIndex        =   48
            Top             =   0
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Documentos de Referencia"
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
            Height          =   345
            Index           =   0
            Left            =   -15
            TabIndex        =   3
            Top             =   6840
            Width           =   11820
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Letras por cobrar"
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
            Left            =   -60
            TabIndex        =   1
            Top             =   0
            Width           =   12000
         End
      End
   End
   Begin VB.Label lbldocsrefs 
      Caption         =   "Documentos de Referencia"
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
      Height          =   255
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "FrmLetrascobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xMes As Integer         'numero de mes en el que se realiza la operacion

Dim CaracteresNumericos As String
Dim Mostrando As Boolean
Dim CaracteresNumericos2 As String
Dim QueHace As Integer
Dim xRs1 As New ADODB.Recordset
Dim xCuentaDoc As Integer   'codigo de la cuenta contable del documento
Dim ValTipCam As Double
Dim xlibro As Integer
Dim rstvent As New ADODB.Recordset
Dim seEjecuto As Boolean

Sub Anular()

            
                
    Dim Rpta As Integer
    Dim A As Integer
    Rpta = MsgBox("¿Esta seguro de anular " & rstvent("nomdoc") & " Nº " & rstvent("numser") & "-" & rstvent("numdoc") + "?", vbYesNo + vbDefaultButton1 + vbQuestion, Me.Caption)
    
    If Rpta = vbYes Then
        xCon.Execute "UPDATE vta_ventas SET vta_ventas.Anulado = -1, " _
            & " vta_ventas.impbru = 0, vta_ventas.impinaf = 0, vta_ventas.impigv = 0,  vta_ventas.impisc = 0,  " _
            & " vta_ventas.impotr = 0, vta_ventas.imptotdoc = 0,  vta_ventas.impsal = 0  " _
            & " WHERE vta_ventas.id = " & rstvent("id") & " "
        
        
        
        xCon.Execute "DELETE * FROM vta_ventasdet WHERE vta_ventasdet.idvta = " & rstvent("id") & ""
                
        MsgBox rstvent("nomdoc") & " se anulo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        rstvent.Requery
        Dg1.Refresh
    End If
End Sub

Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    Toolbar1.Buttons(11).Enabled = Not Toolbar1.Buttons(11).Enabled
    
    Toolbar1.Buttons(13).Enabled = Not Toolbar1.Buttons(13).Enabled
    Toolbar1.Buttons(15).Enabled = Not Toolbar1.Buttons(15).Enabled
End Sub

Sub Cancelar()

    Bloquea
    Label5.Caption = "Detalle de Letras por cobrar"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
     
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
            
End Sub
Sub Nuevo()
    QueHace = 1
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Label5.Caption = "Agregando Letras por cobrar Venta"
    TxtFecha.SetFocus
End Sub

Sub Modificar()
    
    
    QueHace = 2
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    MuestraSegundoTab
    Label5.Caption = "Modificando Letras por cobrar"
    
    Fg1.SelectionMode = flexSelectionFree
    
    TxtFecha.SetFocus
End Sub

Function Grabar() As Boolean
        
    If TxtFecha.Valor = "" Then
        MsgBox "No ha especificado la fecha de emision del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFecha.SetFocus
        Exit Function
    End If
    If Txtfechaven.Valor = "" Then
        MsgBox "No ha especificado la fecha de vencimiento del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Txtfechaven.SetFocus
        Exit Function
    End If
    
    If TxtIdMoneda.Text = "" Then
        MsgBox "No ha especificado la moneda del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMoneda.SetFocus
        Exit Function
    End If
    
    If TxtConPag.Text = "" Then
        MsgBox "No ha especificado la condicion de pago del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtConPag.SetFocus
        Exit Function
    End If
    
    If TxtIdMov.Text = "" Then
        MsgBox "No ha especificado el movimiento ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMov.SetFocus
        Exit Function
    End If
    
    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Function
    End If
        
    If TxtNumDoc.Text = "" Then
        MsgBox "No ha especificado Nro de documento ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        Exit Function
    End If
        
    If TxtGlosa.Text = "" Then
        MsgBox "No ha especificado glosa para el documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtGlosa.SetFocus
        Exit Function
    End If
    
    If TxtImporte.Text = "" Then
        MsgBox "No ha especificado importe para la letra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtImporte.SetFocus
        Exit Function
    End If
    
    If Val(TxtImporte.Text) = 0 Then
        MsgBox "No ha especificado importe para la letra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtImporte.SetFocus
        Exit Function
    End If
           
    If TxtNumRuc.Text = "" Then
        MsgBox "No ha especificado cliente para el documento ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Function
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado items para el canje", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    'If QueHace = 1 Then 'Validamos si existe el numero del documento en modo adicion
    'Dim RstCab As New ADODB.Recordset
    
    'RST_Busq RstCab, " Select * from vta_ventas where VAL(tipdoc) =" & Val(TxtTipDoc) & " and Clng(Numser) =" & CLng(TxtNumSer) & " and VAL(numdoc) = " & Val(Me.TxtNumDoc) & " ", xCon
    
    'If RstCab.RecordCount > 0 Then
    '    MsgBox "El Nro de documento ha sido registrado por otro usuario se grabara con otro numero", vbInformation, Me.Caption
    '    TxtNumDoc = HallaNumDoc(Val(TxtTipDoc), Val(Trim(TxtNumSer.Text)))
    'End If
    'Set RstCab = Nothing
    'End If
    
    
    
    
    
    Dim RstCab As New ADODB.Recordset
    Dim Rstdet As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim xIdCuen As Integer
    Dim xTotal As Double
    Dim xNumAsiento As String
    
    Dim xId As Integer
    Dim A As Integer
    Dim X As Integer
    Dim P As Integer
    
    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("con_letras", xCon, "id")
        xNumAsiento = HallaNumAsiento(xMes)
        
        RST_Busq RstCab, "SELECT * FROM con_letras", xCon
        RST_Busq Rstdet, "SELECT * FROM con_letrasdet", xCon
        RST_Busq RstDia, "SELECT * FROM con_diario", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = rstvent("id")
        RST_Busq RstCab, "SELECT * FROM con_letras WHERE id = " & xId & "", xCon
        
        'DEVOLVEMOS EL SALDO ACTUAL DEL DOCUMENTO

'        RST_Busq RstDeta2, "SELECT vta_ventasdet.* From vta_ventasdet WHERE (((vta_ventasdet.idvta)= " & xId & "))", xCon
'        If RstDeta2.RecordCount <> 0 Then
'            RstDeta2.MoveFirst
'            For A = 1 To RstDeta2.RecordCount
'                RST_Busq RstActPro, "SELECT alm_inventario.id, alm_inventario.stckact  From alm_inventario WHERE ((alm_inventario.id=" & RstDeta2("iditem") & "))", xCon
'                If RstActPro.RecordCount = 1 Then
'                    RstActPro("stckact") = RstActPro("stckact") + RstDeta2("canpro")
'                    RstActPro.Update
'                End If
'                Set RstActPro = Nothing
'            Next A
'
'        End If
'
'        Set RstDeta2 = Nothing
        
        'Eliminamos el detalle de la letra
        xCon.Execute "DELETE * FROM con_letrasdet WHERE id = " & xId & ""
        
        RST_Busq Rstdet, "SELECT * FROM con_letrasdet ", xCon
        
        RST_Busq RstDia, "SELECT * FROM con_diario WHERE idmes = " & Format(CDate(Me.TxtFecha.Valor), "mm") & " AND " _
                         & " idlib =" & xlibro & " AND idmov = " & xId & " And iddoc = " & Val(TxtTipDoc) & "", xCon
            
        If RstDia.RecordCount <> 0 Then
            xNumAsiento = RstDia("numasi")
        End If
        
        Set RstDia = Nothing
        
       'Eliminamos el asiento contable
        xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & Format(CDate(TxtFecha.Valor), "mm") & " AND " _
            & " idlib = 8 AND idmov = " & xId & " And iddoc = " & Val(TxtTipDoc) & ""
            
        RST_Busq RstDia, "SELECT * FROM con_diario", xCon
    End If
    
    RstCab("tipmov") = 2 'ingreso
    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("idcliprov") = NulosN(LblidCli.Caption)
    RstCab("numdoc") = TxtNumDoc.Text
    RstCab("fchdoc") = TxtFecha.Valor
    RstCab("fchven") = Txtfechaven.Valor
    RstCab("idconpag") = NulosN(TxtConPag.Text)
    RstCab("idmon") = NulosN(TxtIdMoneda.Text)
    RstCab("imptotdoc") = NulosN(TxtImporte.Text)
    RstCab("impsal") = NulosN(TxtImporte.Text)
    RstCab("numreg") = Trim(Str(xMes)) + xNumAsiento
    RstCab("glosa") = Trim(TxtGlosa)
    RstCab("idmov") = Val(TxtIdMov) 'Id de asiento donde esta la cuenta  del documento de canje
    RstCab("estado") = 0
    RstCab.Update
    
    For A = 1 To Fg2.Rows - 1
        Rstdet.AddNew
        Rstdet("id") = xId
        Rstdet("iddoc") = Val(Fg2.TextMatrix(A, 9))
        Rstdet("impabo") = Val(Fg2.TextMatrix(A, 7))
        Rstdet("salant") = 0
        Rstdet.Update
    Next A
   
    
    'grabamos el documento de canje en la tabla diario
    'Grabamos el libro diario del movimiento
    RstDia.AddNew
    RstDia("año") = xaño
    RstDia("idmes") = xMes
    RstDia("idlib") = 8
    RstDia("iddoc") = Val(TxtTipDoc)
    RstDia("idmov") = xId
    RstDia("numasi") = xNumAsiento
    RstDia("tc") = ValTipCam
    RstDia("idcue") = xCuentaDoc
    If TxtIdMoneda.Text = "1" Then
        RstDia("impdebsol") = Val(TxtImporte.Text)
        RstDia("impdebdol") = 0
    Else
        RstDia("impdebsol") = Val(TxtImporte.Text) * Val(LblTipoCambio.Caption)
        RstDia("impdebdol") = Val(TxtImporte.Text)
    End If
    RstDia.Update
     
      'Grabamos los documentos que intervienen en el canje
          For X = 1 To Fg2.Rows - 1
                    RstDia.AddNew
                    RstDia("año") = xaño
                    RstDia("idmes") = xMes                  'LLAVE - CODIGO DEL MES
                    RstDia("idlib") = xlibro                'LLAVE - CODIGO DEL LIBRO
                    RstDia("iddoc") = Val(TxtTipDoc)        'LLAVE - CODIGO DEL DOCUMENTO
                    RstDia("idmov") = xId                   'LLAVE - CODIGO DEL MOVIMIENTO
                    RstDia("numasi") = xNumAsiento          'LLAVE - NUMERO DE ASIENTO
                    RstDia("tc") = ValTipCam
                    RstDia("idcue") = Fg2.TextMatrix(X, 10)
                    
                    If TxtIdMoneda.Text = "1" Then
                        RstDia("imphabsol") = Val(Fg2.TextMatrix(X, 7))
                        RstDia("imphabdol") = 0
                    Else
                        RstDia("imphabsol") = Val(Fg2.TextMatrix(X, 7)) * Val(LblTipoCambio.Caption)
                        RstDia("imphabdol") = Val(Fg2.TextMatrix(X, 7))
                    End If
                    RstDia.Update
           Next
                         
    xCon.CommitTrans
    MsgBox " Letra se registro con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Set RstCab = Nothing
    Set Rstdet = Nothing
    Set RstDia = Nothing
    Grabar = True
    Exit Function
    
LaCague:
    
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set Rstdet = Nothing
    Set RstDia = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
    
End Function


Private Sub CmdAgregar_Click()
    AgregarParaPago
End Sub

Private Sub CmdBuscaMovimiento_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Id":                         xCampos(0, 1) = "id":            xCampos(0, 2) = "350":          xCampos(0, 3) = "N"
    xCampos(1, 0) = "Movimiento":                 xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3700":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Cuenta":                  xCampos(2, 1) = "cuenta":        xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Descripcion de la Cuenta":   xCampos(3, 1) = "descue":        xCampos(3, 2) = "2900":         xCampos(3, 3) = "C"
    
    xform.SQLCad = "SELECT con_cajabanmovi.*, con_planctas.cuenta, con_cajabanmovi.descripcion AS descue, " _
        & " con_cajabanmovi.tipmov, con_cajabanmovi.tipope FROM con_planctas RIGHT JOIN con_cajabanmovi " _
        & " ON con_planctas.id = con_cajabanmovi.idcue WHERE  con_cajabanmovi.tipope =3 AND con_cajabanmovi.tipMOV = 1 "

    xform.Titulo = "Buscando Movimientos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdMov.Text = xRs("id")
        LblMovimiento.Caption = Trim(xRs("descripcion"))
        TxtTipDoc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub CmdBusCondicion_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
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
        TxtConPag.Text = xRs("id")
        LblCondPag.Caption = xRs("descripcion")
        TxtIdMov.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_moneda ORDER BY descripcion"
    
    xform.Titulo = "Buscando Moneda"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdMoneda.Text = xRs("id")
        LblMoneda.Caption = xRs("descripcion")
        Fg1.SetFocus
        
        If Trim(TxtIdMoneda.Text) = "1" Then
            lblDestipocambio.Visible = False
            LblTipoCambio.Visible = False
        Else
            If TxtFecha.Valor = "" Then
                MsgBox "No ha especificado la fecha del documento, no se puede determinar " & Chr(13) _
                    & "la fecha del tipo de cambio para este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                
                TxtIdMoneda.Text = ""
                TxtFecha.SetFocus
                Exit Sub
            End If
            
            
            lblDestipocambio.Visible = True
            LblTipoCambio.Visible = True
            Set xRs = Nothing
            Set xRs = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = CDATE('" & TxtFecha.Valor & "')", xCon)
            If xRs.RecordCount = 1 Then
                LblTipoCambio.Caption = Format(xRs("impven"), "0.000")
            End If
            
        End If
    TxtConPag.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdBusProCli_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Nombre":          xCampos(0, 1) = "nombre":   xCampos(0, 2) = "6000":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":       xCampos(1, 1) = "numruc":   xCampos(1, 2) = "1500":     xCampos(1, 3) = "C"
        

    
        xform.SQLCad = "SELECT mae_cliente.nombre, mae_cliente.numruc, mae_cliente.id " _
            & " From mae_cliente ORDER BY mae_cliente.nombre"
        xform.Titulo = "Buscando Clientes"
           
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtNumRuc.Text = xRs("numruc")
        LblNomCli.Caption = xRs("nombre")
        LblidCli.Caption = xRs("id")
        MuestraDocumentos
        Fg1.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Function HallaNumAsiento(Mes As Integer) As String
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT con_diario.idmes, con_diario.idlib, con_diario.numasi From con_diario " _
        & " WHERE (((con_diario.idmes)=" & Mes & ") AND ((con_diario.idlib)= " & xlibro & " )) ORDER BY numasi", xCon
    
    If Rst.RecordCount = 0 Then
        HallaNumAsiento = "0001"
    Else
        Rst.MoveLast
        HallaNumAsiento = Format(Val(Rst("numasi")) + 1, "0000")
    End If
    Exit Function
End Function

Sub MuestraDocumentos()
    Dim Rst As New ADODB.Recordset
    Dim Rstdet As New ADODB.Recordset
    
    
    
    With Fg1
     
     .Cols = 9
     .Rows = 1
     
     .TextMatrix(0, 1) = "TD"
     .TextMatrix(0, 2) = "Fch Emi"
     .TextMatrix(0, 3) = "Moneda"
     .TextMatrix(0, 4) = "Nº Documento"
     .TextMatrix(0, 5) = "Importe"
     .TextMatrix(0, 6) = "Saldo"
     .TextMatrix(0, 7) = "Idmov"
     .TextMatrix(0, 8) = "Idcuenta"
     
     .ColWidth(2) = 1000
     .ColWidth(4) = 1500
     .ColWidth(7) = 1000
     .ColWidth(8) = 1000
     
    End With
    
   'Fusionamos la tabla de letras y venta para canjear documentos
    
'    RST_Busq Rst, "SELECT vta_ventas.idcli, mae_documento.abrev, vta_ventas.fchdoc, mae_moneda.simbolo, " _
'        & " [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.imptotdoc, vta_ventas.impsal, " _
'        & "  vta_ventas.id, vta_ventas.idmon, vta_ventas.tipdoc " _
'        & " FROM mae_moneda INNER JOIN (mae_documento INNER JOIN (mae_cliente INNER JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) " _
'        & " ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon " _
'        & " WHERE  ((vta_ventas.idcli=" & Val(LblidCli.Caption) & ") AND (vta_ventas.impsal<>0))", xCon
    
    
    
    
    RST_Busq Rst, "   SELECT con_diario.iddoc, mae_documento.abrev, vta_ventas.fchdoc, mae_moneda.simbolo, [vta_ventas]![numser]+ '-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.imptotdoc, vta_ventas.impsal, con_diario.idcue, vta_ventas.id, vta_ventas.idmon " _
                  & " FROM mae_moneda INNER JOIN (mae_documento INNER JOIN (con_planctas INNER JOIN (vta_ventas INNER JOIN con_diario ON vta_ventas.id = con_diario.idmov) ON con_planctas.id = con_diario.idcue) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon " _
                  & " WHERE (vta_ventas.idcli = 1) AND (con_diario.idlib = 2) AND ((Not (con_diario.iddoc)=7 And Not (con_diario.iddoc)=8)) AND (con_planctas.cuenta  Like '121%') AND vta_ventas.impsal <> 0 " _
                  & " Union " _
                  & " SELECT con_diario.iddoc, mae_documento.abrev, vta_notascreabo.fchdoc, mae_moneda.simbolo, [vta_notascreabo]![numser]+ '-'+[vta_notascreabo]![numdoc] AS numdoc, vta_notascreabo.imptotdoc, vta_notascreabo.impsal, con_diario.idcue, vta_notascreabo.id, vta_notascreabo.idmon " _
                  & " FROM mae_moneda INNER JOIN ((mae_documento INNER JOIN vta_notascreabo ON mae_documento.id = vta_notascreabo.tipdoc) INNER JOIN (con_planctas INNER JOIN con_diario ON con_planctas.id = con_diario.idcue) ON vta_notascreabo.id = con_diario.idmov) ON mae_moneda.id = vta_notascreabo.idmon " _
                  & " WHERE vta_notascreabo.idcli = 1 AND con_diario.idlib = 2 AND con_diario.iddoc = 8  AND con_planctas.cuenta  Like '121%' AND vta_notascreabo.impsal <> 0 " _
                  & " Union " _
                  & " SELECT con_diario.iddoc,mae_documento.abrev, con_letras.fchdoc, mae_moneda.simbolo, [con_letras]![numdoc] AS numdoc, con_letras.imptotdoc, con_letras.impsal,  con_diario.idcue, con_letras.id, con_letras.idmon " _
                  & " FROM (mae_moneda INNER JOIN (mae_documento INNER JOIN con_letras ON mae_documento.id = con_letras.tipdoc) ON mae_moneda.id = con_letras.idmon) INNER JOIN (con_planctas INNER JOIN con_diario ON con_planctas.id = con_diario.idcue) ON con_letras.id = con_diario.idmov " _
                  & " WHERE con_letras.idcliprov = 1 AND con_diario.idlib = 8 AND con_planctas.cuenta Like '123%' AND con_letras.impsal <> 0 ", xCon
                  
    If Rst.RecordCount <> 0 Then
        With Fg1
        Fg1.Rows = 1
                
        Do While Not Rst.EOF
                        
                .AddItem ""
                Fg1.TextMatrix(.Rows - 1, 1) = Rst("abrev")
                Fg1.TextMatrix(.Rows - 1, 2) = Rst("fchdoc")
                Fg1.TextMatrix(.Rows - 1, 3) = Rst("simbolo")
                Fg1.TextMatrix(.Rows - 1, 4) = Rst("numdoc")
                Fg1.TextMatrix(.Rows - 1, 5) = Format(Rst("imptotdoc"), "0.00")
                Fg1.TextMatrix(.Rows - 1, 6) = Format(Rst("impsal"), "0.00")
                Fg1.TextMatrix(.Rows - 1, 7) = Rst("id")
                Fg1.TextMatrix(.Rows - 1, 8) = Rst("idcue")
            
                'Obtenemos la cuenta del documento a canjear segun como ha sido provisionado
                'Exit Sub
                'Set Rstdet = BuscaConCriterio(" SELECT vta_ventas.idcli, con_diario.idlib, con_diario.idmov, con_diario.iddoc, con_diario.idcue, con_planctas.cuenta " _
                '              & " FROM con_planctas INNER JOIN (vta_ventas INNER JOIN con_diario ON vta_ventas.id = con_diario.idmov) ON con_planctas.id = con_diario.idcue " _
                '              & " WHERE vta_ventas.idcli = " & Rst("idlci") & "AND con_diario.idlib = 2  AND con_diario.idmov =" & Rst("id") & " AND con_diario.iddoc =" & Rst("iddoc") & "", xCon)
                                            
                'ó en la tabla de provision mae_documentocta donde indica que cuenta esta realcionada con el documento
                            
                'Set Rstdet = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & Rst("tipdoc") & " and mae_documentocta.idmon =" & Rst("IdMon") & " and tipope = -1", xCon)
                'If Rstdet.RecordCount = 1 Then
                '    Fg1.TextMatrix(.Rows - 1, 8) = Rstdet("idcuen")
                'End If
                
            
            Rst.MoveNext
        Loop
        
        End With
    Else
        MsgBox "El Cliente seleccionado no tiene documentos pendientes de cobro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    Set Rst = Nothing
    Set Rstdet = Nothing
End Sub

Private Sub CmdBusTipDoc_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
        & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuenvta  as cuentaimp" _
        & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuenvta = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id"
    
    Dim xImpuesto As Double
    
    xform.Titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        
        TxtTipDoc.Text = xRs("id")
        LblNomDoc.Caption = xRs("descripcion")
        
        
        
        
        Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & Val(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & Val(TxtIdMoneda.Text) & " and tipope = -1", xCon)
            If xRs2.RecordCount > 0 Then
                xCuentaDoc = NulosN(xRs2("idcuen"))
            End If
            Set xRs2 = Nothing
                        
        
        TxtNumDoc.SetFocus
        
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdEliminar_Click()
    Fg2.RemoveItem Fg2.Row
End Sub



Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Fg2.Col = 7 Then
        Fg2.TextMatrix(Fg2.Row, 7) = Format(Fg2.TextMatrix(Fg2.Row, 7), "0.00")
        Fg2.TextMatrix(Fg2.Row, 8) = Val(Fg2.TextMatrix(Fg2.Row, 6)) - Val(Fg2.TextMatrix(Fg2.Row, 7))
        Fg2.TextMatrix(Fg2.Row, 8) = Format(Fg2.TextMatrix(Fg2.Row, 8), "0.00")
    End If
End Sub

Private Sub Fg2_EnterCell()
    If Fg2.Col = 7 Then
        Fg2.Editable = flexEDKbdMouse
    Else
        Fg2.Editable = flexEDNone
    End If
End Sub

Private Sub Form_Activate()










If seEjecuto = False Then
        seEjecuto = True
        Dim Rpta As Integer
        Dim Rst As New ADODB.Recordset
        
        
        Set Rst = Nothing
        
        
        RST_Busq rstvent, " SELECT con_letras.*, [con_letras]![numdoc] AS numerodoc, IIf(con_letras.Anulado=0,'Generado','Anulado') AS EstadoVenta, " _
        & " mae_documento.descripcion AS Nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.nombre, mae_cliente.numruc, mae_moneda.simbolo, mae_moneda.descripcion AS descmon, con_tc.impven " _
        & " FROM (mae_documento INNER JOIN (((con_letras INNER JOIN mae_cliente ON con_letras.id = mae_cliente.id) LEFT JOIN mae_condpago ON con_letras.idconpag = mae_condpago.id) INNER JOIN mae_moneda ON con_letras.idmon = mae_moneda.id) ON mae_documento.id = con_letras.tipdoc) LEFT JOIN con_tc ON con_letras.fchdoc = con_tc.fecha ", xCon

        'RST_Busq rstvent, " SELECT con_letras.*, [con_letras]![numdoc] AS numerodoc, IIf(con_letras.Anulado = 0, 'Generado', 'Anulado') AS [EstadoVenta], " _
        '& " mae_documento.descripcion as [Nomdoc], mae_condpago.descripcion AS desccond, mae_documento.abrev , mae_cliente.nombre,mae_cliente.numruc, mae_moneda.simbolo,mae_moneda.descripcion AS descmon " _
        '& " FROM mae_documento INNER JOIN (((con_letras INNER JOIN mae_cliente ON con_letras.id = mae_cliente.id) LEFT JOIN mae_condpago ON con_letras.idconpag = mae_condpago.id) INNER JOIN mae_moneda ON con_letras.idmon = mae_moneda.id) ON mae_documento.id = con_letras.tipdoc ", xCon

        
        'RST_Busq rstvent, "SELECT Vta_Ventas.*, mae_cliente.nombre, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numerodoc, IIf(vta_ventas.Anulado = 0, 'Facturado', 'Anulado') AS [EstadoVenta], " _
        '    & " mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, " _
        '    & " mae_moneda.descripcion AS descmon, mae_moneda.simbolo, mae_impuestos.idcuenvta, con_tc.impcom, mae_tipoproducto.descripcion AS desctipcom" _
        '    & " FROM (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN ((mae_documento LEFT JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) " _
        '    & " RIGHT JOIN (mae_condpago RIGHT JOIN (vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON " _
        '    & " mae_condpago.id = vta_ventas.idconpag) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) " _
        '    & " ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_tipoproducto ON vta_ventas.idtipo = mae_tipoproducto.id " _
        '    & " WHERE (((vta_ventas.numreg) Like '" & Format(xMes, "00") & "%')) ORDER BY vta_ventas.fchdoc, vta_ventas.NumSer, vta_ventas.Numdoc DESC ", xCon
        
        Set Dg1.DataSource = rstvent
        If rstvent.RecordCount = 0 Then
            Rpta = MsgBox("No se ha registrado ninguna venta, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            If Rpta = vbYes Then
                Nuevo
            Else
                Unload Me
            End If
        Else
            Dg1.SetFocus
        End If
        
    End If
End Sub

Private Sub Form_Load()
    CaracteresNumericos = "0123456789." & Chr(8)
    CaracteresNumericos2 = "12" & Chr(8)
    
    QueHace = 3
    TabOne1.CurrTab = 0
    seEjecuto = False

    Fg1.ColWidth(7) = 0
    Fg1.ColWidth(8) = 0
    
    Fg2.ColWidth(9) = 0
    Fg2.ColWidth(10) = 0
    
    Fg1.Rows = 1
    Fg2.Rows = 1
    xlibro = 8
    xMes = 12
    xaño = 2006
    
    
        

End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
If OldTab = 0 Then
        
        'Validamos si la cuadricula tiene datos
        
            If QueHace = 3 Then
                If rstvent.RecordCount = 0 Then
                    MsgBox "No existe información para visualizar", vbInformation, Me.Caption
                    Blanquea
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
        If rstvent.RecordCount = 0 Then
            MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
            Exit Sub
        End If
        
        Modificar
    End If
    
    If Button.Index = 3 Then
        If rstvent.RecordCount = 0 Then
                MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
                Exit Sub
            End If
      
        'Validamos si el documento esta anulado
        If rstvent("Anulado") = -1 Then
            MsgBox rstvent("nomdoc") & " ya fue anulado, seleccione otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Anular
    End If
        
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            rstvent.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    
    'If Button.Index = 11 Then CambiarMes
    
    If Button.Index = 15 Then
        Set rstvent = Nothing
        Unload Me
    End If

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  If ButtonMenu.Parent.Index = 2 Then
        
       'MODIFICACION DE DOCUMENTOS
        If ButtonMenu.Index = 1 Then
            If rstvent.RecordCount = 0 Then
                MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
                Exit Sub
            End If
            If rstvent("anulado") = -1 Then
                        MsgBox "No puede modificar " & rstvent("nomdoc") & " anulado proceda a restaurarlo", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
                        Exit Sub
            Else
                        Modificar
            End If
        End If
        
        'RESTAURAR DOCUMENTOS
        If ButtonMenu.Index = 2 Then
            If rstvent.RecordCount = 0 Then
                MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
                Exit Sub
            End If
                If rstvent("anulado") = -1 Then ' SI EL DOCUMENTO ESTA ANULADO
                    'RestaurarFactura
                End If
        End If
    
    End If
  
  If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then Anular
        If ButtonMenu.Index = 2 Then Eliminar
        
    End If

End Sub

Private Sub TxtConPag_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        
        If NulosC(TxtConPag.Text) = "" Then Exit Sub
        Dim xRs1 As New ADODB.Recordset
        
        RST_Busq xRs1, "SELECT * FROM mae_condpago WHERE id = " & Val(TxtConPag.Text) & "", xCon
        
        If xRs1.RecordCount = 0 Then
            TxtConPag.Text = ""
            LblCondPag.Caption = ""
        Else
            LblCondPag.Caption = Trim(xRs1("descripcion"))
        End If
        Set xRs1 = Nothing
        TxtIdMov.SetFocus
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If


End Sub

Private Sub TxtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtImporte.SetFocus
End Sub



Private Sub TxtIdMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        
        If NulosC(TxtIdMoneda.Text) = "" Then Exit Sub
        Dim xRs1 As New ADODB.Recordset
        
        'buscamos el codigo de la moneda         digitada
        RST_Busq xRs1, "SELECT * FROM mae_moneda WHERE id = " & Val(TxtIdMoneda.Text) & "", xCon
        
        If xRs1.RecordCount = 0 Then
            TxtIdMoneda.Text = ""
            LblMoneda.Caption = ""
        Else
            LblMoneda.Caption = Trim(xRs1("descripcion"))
            
            If Trim(TxtIdMoneda.Text) = "1" Then
                lblDestipocambio.Visible = False
                LblTipoCambio.Visible = False
            Else
                If TxtFecha.Valor = "" Then
                    MsgBox "No ha especificado la fecha del documento, no se puede determinar " & Chr(13) _
                        & "la fecha del tipo de cambio para este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    
                    TxtIdMoneda.Text = ""
                    LblMoneda.Caption = ""
                    TxtFecha.SetFocus
                    Exit Sub
                End If
                lblDestipocambio.Visible = True
                LblTipoCambio.Visible = True
                Set xRs1 = Nothing
                Set xRs1 = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = CDATE('" & TxtFecha.Valor & "')", xCon)
                If xRs1.RecordCount = 1 Then
                    LblTipoCambio.Caption = Format(xRs1("impven"), "0.000")
                    ValTipCam = xRs1("impven")
                Else
                    LblTipoCambio.Caption = "0.00"
                    ValTipCam = 0
                    
                End If
            End If
        End If
        Set xRs1 = Nothing
        TxtConPag.SetFocus
        
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If


End Sub

Private Sub TxtIdMoneda_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    CmdBusMon_Click
End If
End Sub

Private Sub TxtIdMov_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
    
        If NulosC(TxtIdMov.Text) = "" Then Exit Sub
        
        RST_Busq xRs1, "SELECT con_cajabanmovi.*, con_planctas.cuenta, con_cajabanmovi.descripcion AS descue, " _
        & " con_cajabanmovi.tipmov, con_cajabanmovi.tipope FROM con_planctas RIGHT JOIN con_cajabanmovi " _
        & " ON con_planctas.id = con_cajabanmovi.idcue WHERE con_cajabanmovi.id = " & Val(TxtIdMov) & "", xCon

        LblMovimiento.Caption = Trim(xRs1("descripcion"))
        TxtTipDoc.SetFocus
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    
End Sub

Private Sub TxtIdMov_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
      CmdBuscaMovimiento_Click
    End If

End Sub

Private Sub TxtImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtNumRuc.SetFocus
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub
Sub MuestraSegundoTab()

    
    TxtIdMoneda.Text = rstvent("idmon")
    TxtConPag.Text = rstvent("idconpag")
    TxtIdMov.Text = rstvent("idmov")
    TxtTipDoc.Text = rstvent("tipdoc")
    TxtNumDoc.Text = rstvent("numdoc")
    TxtNumRuc.Text = rstvent("numruc")
    TxtFecha.Valor = rstvent("fchdoc")
    Txtfechaven.Valor = rstvent("fchven")

    LblNomDoc.Caption = rstvent("nomdoc")
    LblNomCli.Caption = rstvent("nombre")
    LblCondPag.Caption = NulosC(rstvent("desccond"))
    LblMoneda.Caption = rstvent("descmon")
    LblidCli.Caption = rstvent("idcliprov")
  
    TxtGlosa = rstvent("glosa")
    TxtImporte = rstvent("imptotdoc")


    If rstvent("idmon") = 1 Then
        LblTipoCambio.Visible = False
    Else
        LblTipoCambio.Visible = True
        LblTipoCambio.Caption = rstvent("impven")
    End If

    Dim Rst As New ADODB.Recordset
    Dim Rstdet As New ADODB.Recordset

    Mostrando = True

    RST_Busq Rst, " SELECT con_diario.iddoc, mae_documento.abrev, vta_ventas.fchdoc, mae_moneda.simbolo, [vta_ventas]![numser]+ '-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.imptotdoc, vta_ventas.impsal, con_diario.idcue, vta_ventas.id, vta_ventas.idmon " _
                  & " FROM mae_moneda INNER JOIN (mae_documento INNER JOIN (con_planctas INNER JOIN (vta_ventas INNER JOIN con_diario ON vta_ventas.id = con_diario.idmov) ON con_planctas.id = con_diario.idcue) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon " _
                  & " WHERE (vta_ventas.idcli = 1) AND (con_diario.idlib = 2) AND ((Not (con_diario.iddoc)=7 And Not (con_diario.iddoc)=8)) AND (con_planctas.cuenta  Like '121%') AND vta_ventas.impsal <> 0 " _
                  & " Union " _
                  & " SELECT con_diario.iddoc, mae_documento.abrev, vta_notascreabo.fchdoc, mae_moneda.simbolo, [vta_notascreabo]![numser]+ '-'+[vta_notascreabo]![numdoc] AS numdoc, vta_notascreabo.imptotdoc, vta_notascreabo.impsal, con_diario.idcue, vta_notascreabo.id, vta_notascreabo.idmon " _
                  & " FROM mae_moneda INNER JOIN ((mae_documento INNER JOIN vta_notascreabo ON mae_documento.id = vta_notascreabo.tipdoc) INNER JOIN (con_planctas INNER JOIN con_diario ON con_planctas.id = con_diario.idcue) ON vta_notascreabo.id = con_diario.idmov) ON mae_moneda.id = vta_notascreabo.idmon " _
                  & " WHERE vta_notascreabo.idcli = 1 AND con_diario.idlib = 2 AND con_diario.iddoc = 8  AND con_planctas.cuenta  Like '121%' AND vta_notascreabo.impsal <> 0 " _
                  & " Union " _
                  & " SELECT con_diario.iddoc,mae_documento.abrev, con_letras.fchdoc, mae_moneda.simbolo, [con_letras]![numdoc] AS numdoc, con_letras.imptotdoc, con_letras.impsal,  con_diario.idcue, con_letras.id, con_letras.idmon " _
                  & " FROM (mae_moneda INNER JOIN (mae_documento INNER JOIN con_letras ON mae_documento.id = con_letras.tipdoc) ON mae_moneda.id = con_letras.idmon) INNER JOIN (con_planctas INNER JOIN con_diario ON con_planctas.id = con_diario.idcue) ON con_letras.id = con_diario.idmov " _
                  & " WHERE con_letras.idcliprov = 1 AND con_diario.idlib = 8 AND con_planctas.cuenta Like '123%' AND con_letras.impsal <> 0 ", xCon
                  
    If Rst.RecordCount <> 0 Then
        With Fg2
        Fg2.Rows = 1
                
        Do While Not Rst.EOF
                        
                .AddItem ""
                Fg2.TextMatrix(.Rows - 1, 1) = Rst("abrev")
                Fg2.TextMatrix(.Rows - 1, 2) = Rst("fchdoc")
                Fg2.TextMatrix(.Rows - 1, 3) = Rst("simbolo")
                Fg2.TextMatrix(.Rows - 1, 4) = Rst("numdoc")
                Fg2.TextMatrix(.Rows - 1, 5) = Format(Rst("imptotdoc"), "0.00")
                Fg2.TextMatrix(.Rows - 1, 6) = Format(Rst("impsal"), "0.00")
                Fg2.TextMatrix(.Rows - 1, 7) = Rst("id")
                Fg2.TextMatrix(.Rows - 1, 8) = Rst("idcue")
                                                       
                Rst.MoveNext
        Loop
        
        End With
    End If
    Set Rstdet = Nothing
    Mostrando = False

    Set Rstdet = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & Val(TxtTipDoc.Text) & " AND mae_documentocta.idmon =" & Val(TxtIdMoneda.Text) & " and tipope = -1", xCon)
    If Rstdet.RecordCount = 1 Then
        xCuentaDoc = Rstdet("idcuen")
    End If

    Set Rstdet = Nothing
    Set Rst = Nothing
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    
    If rstvent.RecordCount = 0 Then
                MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
                Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de eliminar " & rstvent("nomdoc") & " seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM con_diario WHERE idmov = " & rstvent("id") & " AND idlib = " & xlibro & " AND Iddoc = " & rstvent("tipdoc") & ""
               
        xCon.Execute "DELETE * FROM con_letras WHERE id = " & rstvent("id") & ""
        xCon.Execute "DELETE * FROM con_letrasdet WHERE id = " & rstvent("id") & ""
        
        MsgBox "Letra por cobrar se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        rstvent.Requery
        Dg1.Refresh
    End If
    
    
    
End Sub





Sub Blanquea()
Dim X As Control
 For Each X In FrmLetrascobrar.Controls
    If TypeOf X Is TextBox Then
       X.Text = ""
    End If
 Next X
 
 LblMovimiento.Caption = ""
 LblNomDoc.Caption = ""
 LblMoneda.Caption = ""
 LblNomCli.Caption = ""
 LblidCli.Caption = ""
 Me.LblTipoCambio = ""
 Set X = Nothing
End Sub

Sub Bloquea()
    TxtFecha.Locked = Not TxtFecha.Locked
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    
    TxtIdMov.Locked = Not TxtIdMov.Locked
    TxtIdMoneda.Locked = Not TxtIdMoneda.Locked
    TxtGlosa.Locked = Not TxtGlosa.Locked
    TxtImporte.Locked = Not TxtImporte.Locked
    
    
End Sub

Sub AgregarParaPago()
    Dim A As Integer
    
    If Fg1.Rows = 1 Then
        MsgBox "No hay documentos para agregar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    For A = 1 To Fg2.Rows - 1
        If Val(Fg2.TextMatrix(A, 7)) = Val(Fg1.TextMatrix(Fg1.Row, 7)) Then
            MsgBox "El documento seleccionado ya esta agregado para cancelacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
    Next A
    
    Fg2.Rows = Fg2.Rows + 1
    Fg2.TextMatrix(Fg2.Rows - 1, 1) = Fg1.TextMatrix(Fg1.Row, 1)
    Fg2.TextMatrix(Fg2.Rows - 1, 2) = Fg1.TextMatrix(Fg1.Row, 2)
    Fg2.TextMatrix(Fg2.Rows - 1, 3) = Fg1.TextMatrix(Fg1.Row, 3)
    Fg2.TextMatrix(Fg2.Rows - 1, 4) = Fg1.TextMatrix(Fg1.Row, 4)
    
    Fg2.TextMatrix(Fg2.Rows - 1, 5) = Fg1.TextMatrix(Fg1.Row, 5)
    Fg2.TextMatrix(Fg2.Rows - 1, 6) = Fg1.TextMatrix(Fg1.Row, 6)
    Fg2.TextMatrix(Fg2.Rows - 1, 9) = Fg1.TextMatrix(Fg1.Row, 7)
    Fg2.TextMatrix(Fg2.Rows - 1, 10) = Fg1.TextMatrix(Fg1.Row, 8)
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtGlosa.SetFocus
End Sub

Private Sub TxtNumRuc_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
 CmdBusProCli_Click
End If

End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
                
        If NulosC(TxtTipDoc.Text) = "" Then Exit Sub
        
            Dim xRs As New ADODB.Recordset
            Dim xRs2 As New ADODB.Recordset
    
            RST_Busq xRs, "SELECT * FROM Mae_documento WHERE mae_documento.id  = " & Val(TxtTipDoc.Text) & "", xCon
    
            If xRs.RecordCount = 0 Then
                TxtTipDoc.Text = ""
                LblNomDoc.Caption = ""
            Else
                TxtTipDoc.Text = xRs("id")
                LblNomDoc.Caption = xRs("descripcion")
    
                
                Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & Val(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & Val(TxtIdMoneda) & " and mae_documentocta.tipope = -1", xCon)
                If xRs2.RecordCount > 0 Then
                   xCuentaDoc = NulosN(xRs2("idcuen"))
                End If
            End If
                Set xRs2 = Nothing
                TxtNumDoc.SetFocus
        Else
            If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If

End Sub

Private Sub TxtTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
 CmdBusTipDoc_Click
End If
End Sub
