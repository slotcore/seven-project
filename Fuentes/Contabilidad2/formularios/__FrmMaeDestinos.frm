VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmMaeDestinos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Destinos"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   11685
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7065
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaeDestinos.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaeDestinos.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaeDestinos.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaeDestinos.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaeDestinos.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaeDestinos.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaeDestinos.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaeDestinos.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaeDestinos.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaeDestinos.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaeDestinos.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaeDestinos.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
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
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6780
      Left            =   15
      TabIndex        =   8
      Top             =   375
      Width           =   11670
      _cx             =   20585
      _cy             =   11959
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
      Appearance      =   2
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6360
         Left            =   45
         TabIndex        =   12
         Top             =   375
         Width           =   11580
         Begin VB.Frame Frame3 
            Height          =   4830
            Left            =   840
            TabIndex        =   15
            Top             =   870
            Width           =   10050
            Begin VB.OptionButton OptNo 
               Caption         =   "&No"
               Height          =   195
               Left            =   7005
               TabIndex        =   36
               Top             =   2400
               Width           =   600
            End
            Begin VB.OptionButton OptSi 
               Caption         =   "&Si"
               Height          =   195
               Left            =   6270
               TabIndex        =   35
               Top             =   2400
               Width           =   480
            End
            Begin VB.Frame Frame4 
               Height          =   1800
               Left            =   7695
               TabIndex        =   31
               Top             =   2820
               Width           =   2085
               Begin VB.CommandButton CmdDelDoc 
                  Caption         =   "Eliminar Documento"
                  Height          =   465
                  Left            =   195
                  TabIndex        =   33
                  Top             =   975
                  Width           =   1695
               End
               Begin VB.CommandButton CmdAdd 
                  Caption         =   "Agregar Documento"
                  Height          =   465
                  Left            =   195
                  TabIndex        =   32
                  Top             =   495
                  Width           =   1695
               End
            End
            Begin VB.CommandButton CmdEnt 
               Height          =   240
               Left            =   2325
               Picture         =   "FrmMaeDestinos.frx":277E
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   1980
               Width           =   240
            End
            Begin VB.CommandButton CmdMon 
               Height          =   240
               Left            =   2325
               Picture         =   "FrmMaeDestinos.frx":28B0
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   1665
               Width           =   240
            End
            Begin VB.CommandButton CmdTipMov 
               Height          =   240
               Left            =   2325
               Picture         =   "FrmMaeDestinos.frx":29E2
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   1350
               Width           =   240
            End
            Begin VB.CommandButton CmdBusMon 
               Height          =   240
               Left            =   3090
               Picture         =   "FrmMaeDestinos.frx":2B14
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   1035
               Width           =   240
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   1710
               Left            =   240
               TabIndex        =   6
               Top             =   2910
               Width           =   7365
               _cx             =   12991
               _cy             =   3016
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmMaeDestinos.frx":2C46
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
            Begin VB.TextBox TxtEntGen 
               Height          =   300
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   5
               Text            =   "TxtEntGen"
               Top             =   1950
               Width           =   800
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   4
               Text            =   "TxtIdMon"
               Top             =   1635
               Width           =   800
            End
            Begin VB.TextBox TxtTipMov 
               Height          =   300
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   3
               Text            =   "TxtTipMov"
               Top             =   1320
               Width           =   800
            End
            Begin VB.TextBox TxtCuenta 
               Height          =   300
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   2
               Text            =   "TxtCuenta"
               Top             =   1005
               Width           =   1560
            End
            Begin VB.TextBox TxtDescripcion 
               Height          =   300
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   1
               Text            =   "TxtDescripcion"
               Top             =   690
               Width           =   7785
            End
            Begin VB.TextBox TxtId 
               Height          =   300
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   0
               Text            =   "TxtId"
               Top             =   375
               Width           =   800
            End
            Begin VB.Label LblIdCuenta 
               AutoSize        =   -1  'True
               Caption         =   "LblIdCuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   8070
               TabIndex        =   37
               Top             =   1380
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Se mostrara este destino en el modulo cuentas por rendir"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   34
               Top             =   2385
               Width           =   4860
            End
            Begin VB.Label LblDesEntGen 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDesEntGen"
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
               Left            =   2625
               TabIndex        =   30
               Top             =   1950
               Width           =   3045
            End
            Begin VB.Label LblDesMon 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDesMon"
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
               Left            =   2625
               TabIndex        =   29
               Top             =   1635
               Width           =   3045
            End
            Begin VB.Label LblDesTipMov 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDesTipMov"
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
               Left            =   2625
               TabIndex        =   28
               Top             =   1320
               Width           =   3045
            End
            Begin VB.Label LblDesCta 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDesCta"
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
               Left            =   3390
               TabIndex        =   27
               Top             =   1005
               Width           =   6195
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Documentos Asignados"
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   22
               Top             =   2670
               Width           =   1680
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Descripcion"
               Height          =   195
               Index           =   8
               Left            =   240
               TabIndex        =   21
               Top             =   720
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nº Cta Contable"
               Height          =   195
               Index           =   7
               Left            =   240
               TabIndex        =   20
               Top             =   1050
               Width           =   1140
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Movimiento"
               Height          =   195
               Index           =   6
               Left            =   240
               TabIndex        =   19
               Top             =   1365
               Width           =   1395
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Index           =   5
               Left            =   240
               TabIndex        =   18
               Top             =   1680
               Width           =   585
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Entidad Generadora"
               Height          =   195
               Index           =   4
               Left            =   240
               TabIndex        =   17
               Top             =   1980
               Width           =   1425
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Codigo"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   16
               Top             =   435
               Width           =   495
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Destino"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   15
            TabIndex        =   13
            Top             =   75
            Width           =   11550
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6360
         Left            =   -12225
         TabIndex        =   9
         Top             =   375
         Width           =   11580
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   5970
            Left            =   0
            TabIndex        =   10
            Top             =   390
            Width           =   11580
            _ExtentX        =   20426
            _ExtentY        =   10530
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Moneda"
            Columns(2).DataField=   "simbolo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Cuenta"
            Columns(3).DataField=   "cuenta"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Descripcion de la Cuenta"
            Columns(4).DataField=   "desccta"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Entidad"
            Columns(5).DataField=   "desent"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1270"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1191"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=6509"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6429"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1349"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1270"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2064"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1984"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=6615"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=6535"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=1720"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1640"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
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
            FootLines       =   1.5
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HDBFDFD&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80&"
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
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=78,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=75,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=76,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=77,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
         Begin VB.Label LblTipo 
            AutoSize        =   -1  'True
            Caption         =   "LblTipo"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   8580
            TabIndex        =   14
            Top             =   30
            Width           =   1050
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Mantenimiento de Destinos"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   90
            TabIndex        =   11
            Top             =   45
            Width           =   11400
         End
      End
   End
End
Attribute VB_Name = "FrmMaeDestinos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstDes As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim Quehace As Integer
Public TipoMovimmiento As Integer

Private Sub CmdAdd_Click()
    If Quehace = 3 Then Exit Sub
    Fg1.Rows = Fg1.Rows + 1
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(3, 4) As String
    Dim A As Integer
    Dim Encontrado As Boolean
    
    xCampos2(0, 0) = "Descripcion":    xCampos2(0, 1) = "descripcion":        xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Abrev.":         xCampos2(1, 1) = "abrev":              xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"
    xCampos2(2, 0) = "Codigo":         xCampos2(2, 1) = "id":                 xCampos2(2, 2) = "1000":         xCampos2(2, 3) = "N"
        
    xform.SQLCad = "SELECT mae_documento.id, mae_documento.descripcion, mae_documento.abrev From mae_documento ORDER BY mae_documento.descripcion"

    xform.Titulo = "Buscando Documentos"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        'buscamos que el documento no haya sido agregado
        For A = 1 To Fg1.Rows - 1
            If NulosN(Fg1.TextMatrix(A, 3)) = xRs("id") Then
                Encontrado = True
            End If
        Next A
        
        If Encontrado = False Then
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = xRs("descripcion")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRs("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosN(xRs("id"))
        Else
            MsgBox "El documento seleccionado ya fue agregado al destino", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Fg1.RemoveItem Fg1.Rows - 1
        End If
        
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    If Quehace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Cuenta":        xCampos2(0, 1) = "cuenta":            xCampos2(0, 2) = "1000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Descripcion":   xCampos2(1, 1) = "descripcion":       xCampos2(1, 2) = "5000":         xCampos2(1, 3) = "C"
        
    xform.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id From con_planctas ORDER BY con_planctas.cuenta"

    xform.Titulo = "Buscando Cuenta Contable"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtCuenta.Text = xRs("cuenta")
        LblDesCta.Caption = xRs("descripcion")
        LblIdCuenta.Caption = xRs("id")
        TxtTipMov.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Command1_Click()
    If Quehace = 3 Then Exit Sub

End Sub

Private Sub Command2_Click()
    If Quehace = 3 Then Exit Sub
    
End Sub

Private Sub Command3_Click()

End Sub

Private Sub CmdDelDoc_Click()
    If Quehace = 3 Then Exit Sub
    
    If Fg1.Rows = 1 Then Exit Sub
    
    Fg1.RemoveItem Fg1.Row
End Sub

Private Sub CmdEnt_Click()
    If Quehace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Descripcion":    xCampos2(0, 1) = "descripcion":        xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Codigo":         xCampos2(1, 1) = "id":                 xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"
        
    xform.SQLCad = "SELECT  mae_entidades.id,  mae_entidades.descripcion From  mae_entidades ORDER BY  mae_entidades.descripcion"

    xform.Titulo = "Buscando Entidad Generadora"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtEntGen.Text = xRs("id")
        LblDesEntGen.Caption = xRs("descripcion")
        Fg1.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdMon_Click()
    If Quehace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Descripcion":    xCampos2(0, 1) = "descripcion":        xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Codigo":         xCampos2(1, 1) = "id":                 xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"
        
    xform.SQLCad = "SELECT mae_moneda.id, mae_moneda.descripcion From mae_moneda ORDER BY mae_moneda.descripcion"

    xform.Titulo = "Buscando Monedas"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtIdMon.Text = xRs("id")
        LblDesMon.Caption = xRs("descripcion")
        TxtEntGen.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdTipMov_Click()
     Exit Sub
'    If Quehace = 3 Then Exit Sub

'    Dim xform As New eps_librerias.FormBuscar
'    Dim xRs As New ADODB.Recordset
'    Dim xCampos2(2, 4) As String
'
'    xCampos2(0, 0) = "Descripcion":    xCampos2(0, 1) = "descripcion":        xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
'    xCampos2(1, 0) = "Codigo":         xCampos2(1, 1) = "id":                 xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"
'
'    xform.SQLCad = "SELECT mae_tipomov.id, mae_tipomov.descripcion From mae_tipomov ORDER BY mae_tipomov.descripcion"
'
'    xform.Titulo = "Buscando Tipo de Movimiento"
'
'    xform.FormaBusca = Principio
'    xform.Criterio = ""
'    xform.Ordenado = "cuenta"
'    xform.CampoBusca = "cuenta"
'    Set xform.Coneccion = xCon
'    Set xRs = xform.BuscarReg(xCampos2)
'
'    If xRs.State = 1 Then
'        TxtTipMov.Text = xRs("id")
'        LblDesTipMov.Caption = xRs("descripcion")
'        TxtIdMon.SetFocus
'    End If
'    Set xform = Nothing
'    Set xRs = Nothing
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
    MuestraSegundoTab
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim Rpta As Integer
        
        SeEjecuto = True

        LblTipo.Caption = Busca_Codigo(TipoMovimmiento, "id", "descripcion", "mae_tipomov", "N", xCon)
        
        If TipoMovimmiento = 1 Then Label1.Caption = "Detalle Destino del Ingreso"
        If TipoMovimmiento = 2 Then Label1.Caption = "Detalle Destino del Egreso"
        
        RST_Busq RstDes, "SELECT tes_destino.*, con_planctas.cuenta, con_planctas.descripcion AS desccta, mae_tipomov.descripcion AS destipmov, mae_moneda.simbolo, mae_moneda.descripcion AS desmon, " _
            & " mae_entidades.descripcion AS desent FROM (mae_tipomov RIGHT JOIN (mae_moneda RIGHT JOIN (con_planctas RIGHT JOIN tes_destino ON con_planctas.id = tes_destino.idcuen) " _
            & " ON mae_moneda.id = tes_destino.idmon) ON mae_tipomov.id = tes_destino.tipmov) LEFT JOIN mae_entidades ON tes_destino.entgen = mae_entidades.id " _
            & " Where (((tes_destino.tipmov) = " & TipoMovimmiento & ")) ORDER BY tes_destino.descripcion", xCon
        
        Set Dg1.DataSource = RstDes
        If RstDes.RecordCount = 0 Then
            Rpta = MsgBox("No se han registrado registros para los destinos, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstDes = Nothing
                Unload Me
            End If
        End If
        
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    Quehace = 3
    TabOne1.CurrTab = 0
    Fg1.ColWidth(3) = 0
    Fg1.SelectionMode = flexSelectionByRow
End Sub

Sub Blanquea()
    TxtId.Text = ""
    TxtDescripcion.Text = ""
    TxtCuenta.Text = ""
    TxtTipMov.Text = ""
    TxtIdMon.Text = ""
    TxtEntGen.Text = ""
    
    LblDesCta.Caption = ""
    LblDesTipMov.Caption = ""
    LblDesMon.Caption = ""
    LblDesEntGen.Caption = ""
End Sub

Sub Bloquea()
    'TxtId.Locked = Not TxtId.Locked
    TxtDescripcion.Locked = Not TxtDescripcion.Locked
    'TxtCuenta.Locked = Not TxtCuenta.Locked
    'TxtTipMov.Locked = Not TxtTipMov.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtEntGen.Locked = Not TxtEntGen.Locked
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
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
End Sub

Sub Nuevo()
    Quehace = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    If TipoMovimmiento = 1 Then Label1.Caption = "Agregando Destino del Ingreso"
    If TipoMovimmiento = 2 Then Label1.Caption = "Agregando Destino del Egreso"
    
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea
    Blanquea
    Fg1.Rows = 1
    TxtTipMov.Text = TipoMovimmiento
    TxtTipMov_Validate True
    
    TxtId.Text = HallaCodigoTabla("tes_destino", xCon, "id")
    TxtDescripcion.SetFocus
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If Quehace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    If Button.Index = 5 Then
        If Grabar = True Then
            RstDes.Requery
            Dg1.Refresh
            Cancelar
        End If
    End If
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 12 Then
        Set RstDes = Nothing
        Unload Me
    End If
End Sub

Sub Cancelar()
    If TipoMovimmiento = 1 Then Label1.Caption = "Detalle Destino del Ingreso"
    If TipoMovimmiento = 2 Then Label1.Caption = "Detalle Destino del Egreso"

    Quehace = 3
    Bloquea
    ActivaTool
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

Function Grabar() As Boolean
    Grabar = False
    
    If NulosC(TxtDescripcion.Text) = "" Then
        MsgBox "No ha especificado la descripcion del destino", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtDescripcion.SetFocus
        Exit Function
    End If

    If NulosC(TxtCuenta.Text) = "" Then
        MsgBox "No ha especificado la cuenta contable para el destino", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCuenta.SetFocus
        Exit Function
    End If

    If NulosC(TxtIdMon.Text) = "" Then
        MsgBox "No ha especificado la moneda del destino", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If

    If NulosC(TxtEntGen.Text) = "" Then
        MsgBox "No ha especificado la entidad generadora para el destino", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtEntGen.SetFocus
        Exit Function
    End If

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId, A As Integer
    
On Error GoTo LaCague

    xCon.BeginTrans
    If Quehace = 1 Then
        RST_Busq RstCab, "SELECT * FROM tes_destino", xCon
        RST_Busq RstDet, "SELECT * FROM tes_destinodoc", xCon
        xId = HallaCodigoTabla("tes_destino", xCon, "id")
        RstCab.AddNew
        
        RstCab("id") = xId
    Else
        RST_Busq RstCab, "SELECT * FROM tes_destino WHERE id = " & NulosN(TxtId.Text) & "", xCon
        xCon.Execute "DELETE * FROM tes_destinodoc WHERE id = " & NulosN(TxtId.Text) & ""
        RST_Busq RstDet, "SELECT * FROM tes_destinodoc", xCon
         xId = RstDes("id")
    End If

    RstCab("descripcion") = NulosC(TxtDescripcion.Text)
    RstCab("idcuen") = LblIdCuenta.Caption
    RstCab("tipmov") = NulosN(TxtTipMov.Text)
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    RstCab("entgen") = NulosN(TxtEntGen.Text)
    If OptSi.Value = True Then RstCab("rendir") = -1
    If OptNo.Value = True Then RstCab("rendir") = 0
    
    RstCab.Update
    
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("id") = xId
        RstDet("iddoc") = Fg1.TextMatrix(A, 3)
        RstDet.Update
    Next A
    
    xCon.CommitTrans
    Grabar = True
    MsgBox "El destino se grabo con exito ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Exit Function

LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing:    Set RstDet = Nothing
    MsgBox "No se pudo guardar el destino por el siguiente motivo : " + Err.Description
    Grabar = False
End Function

Sub Eliminar()
    Dim Rpta As Integer
    Dim Rst As New ADODB.Recordset
    Dim CadSql As String
    Rpta = MsgBox("¿Esta seguro de eliminar el destino seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        'buscamos que el destino no tenga movimientos en la tabla con_cajabanco, con_ctasrendir, con_devoluciones
        
        'buscamos en caja y bancos
        CadSql = "SELECT con_cajabancoorides.idorides From con_cajabancoorides Where (((con_cajabancoorides.idorides) = " & RstDes("id") & ")) ORDER BY con_cajabancoorides.idorides"
        Set Rst = BuscaConCriterio(CadSql, xCon)
        If Rst.RecordCount <> 0 Then
            MsgBox "No se puede eliminar el destino porque tiene datos relacionados con el modulo de Caja y Bancos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set Rst = Nothing
            Exit Sub
        End If
        Set Rst = Nothing
        
        'buscamos en cuentas por rendir
        CadSql = "SELECT con_ctasrendir.iddes From con_ctasrendir WHERE (((con_ctasrendir.iddes)=" & RstDes("id") & "))"

        Set Rst = BuscaConCriterio(CadSql, xCon)
        If Rst.RecordCount <> 0 Then
            MsgBox "No se puede eliminar el destino porque tiene datos relacionados con el modulo de Cargas a Rendir", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set Rst = Nothing
            Exit Sub
        End If
        Set Rst = Nothing

        'buscamos en Devoluciones
        CadSql = "SELECT con_devoluciones.[imp], con_devoluciones.idope From con_devoluciones WHERE (((con_devoluciones.[imp])>0) AND ((con_devoluciones.idope)=" & RstDes("id") & "))"
        Set Rst = BuscaConCriterio(CadSql, xCon)
        If Rst.RecordCount <> 0 Then
            MsgBox "No se puede eliminar el destino porque tiene datos relacionados con el modulo de Rendir Cuentas ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set Rst = Nothing
            Exit Sub
        End If
        Set Rst = Nothing

        xCon.Execute "DELETE * FROM tes_destino WHERE id = " & RstDes("id") & ""
        RstDes.Requery
        Dg1.Refresh
        MsgBox "El destino se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Sub Modificar()
    Quehace = 2
    ActivaTool
    TabOne1.TabEnabled(0) = False
    
    If TipoMovimmiento = 1 Then Label1.Caption = "Modificando Destino del Ingreso"
    If TipoMovimmiento = 2 Then Label1.Caption = "Modificando Destino del Egreso"
    
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea
    
    TxtTipMov.Text = TipoMovimmiento
    TxtTipMov_Validate True
    
    TxtDescripcion.SetFocus
End Sub

Private Sub TxtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtCuenta_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtEntGen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtEntGen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdEnt_Click
    End If
End Sub

Private Sub TxtEntGen_Validate(Cancel As Boolean)
    If NulosN(TxtEntGen.Text) = 0 Then Exit Sub
    
    LblDesEntGen.Caption = Busca_Codigo(NulosN(TxtIdMon.Text), "id", "descripcion", " mae_entidades", "N", xCon)
    If NulosC(LblDesEntGen.Caption) = "" Then
        TxtEntGen.Text = ""
    End If
End Sub

Private Sub TxtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If NulosN(TxtIdMon.Text) = 0 Then Exit Sub
    
    LblDesMon.Caption = Busca_Codigo(NulosN(TxtIdMon.Text), "id", "descripcion", "mae_moneda", "N", xCon)
    If NulosC(LblDesMon.Caption) = "" Then
        TxtTipMov.Text = ""
    End If
End Sub

Private Sub TxtTipMov_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtTipMov_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdTipMov_Click
    End If
End Sub

Private Sub TxtTipMov_Validate(Cancel As Boolean)
    If NulosN(TxtTipMov.Text) = 0 Then Exit Sub
    
    LblDesTipMov.Caption = Busca_Codigo(NulosN(TxtTipMov.Text), "id", "descripcion", "mae_tipomov", "N", xCon)
    If NulosC(LblDesTipMov.Caption) = "" Then
        TxtTipMov.Text = ""
    End If
End Sub

Sub MuestraSegundoTab()
    TxtId.Text = RstDes("id")
    TxtDescripcion.Text = RstDes("descripcion")
    TxtCuenta.Text = NulosC(RstDes("cuenta"))
    LblDesCta.Caption = NulosC(RstDes("desccta"))
    LblIdCuenta.Caption = NulosN(RstDes("idcuen"))
    TxtTipMov.Text = RstDes("tipmov")
    LblDesTipMov.Caption = RstDes("destipmov")
    TxtIdMon.Text = RstDes("idmon")
    LblDesMon.Caption = RstDes("desmon")
    TxtEntGen.Text = NulosN(RstDes("entgen"))
    LblDesEntGen.Caption = NulosC(RstDes("desent"))
    
    If RstDes("rendir") = -1 Then
        OptSi.Value = True
    Else
        OptNo.Value = True
    End If
    'Mostramos los documentos adjuntos
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    RST_Busq xRs, "SELECT mae_documento.id AS iddoc, mae_documento.descripcion, mae_documento.abrev, tes_destinodoc.id FROM tes_destinodoc LEFT JOIN mae_documento " _
        & " ON tes_destinodoc.iddoc = mae_documento.id WHERE (((tes_destinodoc.id) = " & NulosN(TxtId.Text) & "))", xCon

    Fg1.Rows = 1
    
    If xRs.RecordCount <> 0 Then
        xRs.MoveFirst
        For A = 1 To xRs.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = xRs("descripcion")
            Fg1.TextMatrix(A, 2) = NulosC(xRs("abrev"))
            Fg1.TextMatrix(A, 3) = xRs("iddoc")
            
            xRs.MoveNext
            If xRs.EOF = True Then Exit For
        Next A
    End If
    
    Set xRs = Nothing
End Sub
