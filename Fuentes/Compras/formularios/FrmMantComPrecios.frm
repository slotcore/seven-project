VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMantComPrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras - Asignación de Precios"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantComPrecios.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantComPrecios.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantComPrecios.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantComPrecios.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantComPrecios.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantComPrecios.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantComPrecios.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantComPrecios.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantComPrecios.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantComPrecios.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantComPrecios.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantComPrecios.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7020
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   11880
      _cx             =   20955
      _cy             =   12382
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
         Height          =   6600
         Left            =   12525
         TabIndex        =   10
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdQuitarPrecio 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   405
            Left            =   5790
            TabIndex        =   27
            Top             =   6150
            Width           =   1755
         End
         Begin VB.CommandButton CmdAgregarPrec 
            Caption         =   "Agregar Precios"
            Height          =   405
            Left            =   3990
            TabIndex        =   26
            Top             =   6150
            Width           =   1755
         End
         Begin VB.TextBox TxtObs 
            Height          =   2895
            Left            =   7650
            MultiLine       =   -1  'True
            TabIndex        =   25
            Tag             =   "a"
            Text            =   "FrmMantComPrecios.frx":277E
            Top             =   3675
            Width           =   4080
         End
         Begin VSFlex7Ctl.VSFlexGrid fg1 
            Height          =   2430
            Left            =   60
            TabIndex        =   24
            Top             =   3675
            Width           =   7575
            _cx             =   13361
            _cy             =   4286
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
            Rows            =   30
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmMantComPrecios.frx":2787
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
         Begin VB.Frame Frame3 
            Height          =   3135
            Left            =   60
            TabIndex        =   13
            Top             =   525
            Width           =   11670
            Begin VB.CommandButton CmdBusProd 
               Height          =   240
               Left            =   3975
               Picture         =   "FrmMantComPrecios.frx":289D
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   390
               Width           =   240
            End
            Begin VB.TextBox TxtCodProd 
               Height          =   300
               Left            =   2250
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   0
               Text            =   "TxtCodProd"
               Top             =   360
               Width           =   1995
            End
            Begin VB.TextBox TxtPreTope 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2250
               Locked          =   -1  'True
               TabIndex        =   1
               Tag             =   "a"
               Text            =   "TxtPreTope"
               Top             =   1320
               Width           =   1320
            End
            Begin VB.TextBox TxtStocMax 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2250
               Locked          =   -1  'True
               TabIndex        =   3
               Tag             =   "a"
               Text            =   "TxtStocMax"
               Top             =   1635
               Width           =   1320
            End
            Begin VB.TextBox TxtTopeMax 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   5445
               Locked          =   -1  'True
               TabIndex        =   2
               Tag             =   "a"
               Text            =   "TxtTopeMax"
               Top             =   1320
               Width           =   1320
            End
            Begin VB.TextBox TxtStocMin 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2250
               Locked          =   -1  'True
               TabIndex        =   4
               Tag             =   "a"
               Text            =   "TxtStocMin"
               Top             =   1980
               Width           =   1320
            End
            Begin VB.CommandButton CmdVerPrecHist 
               Caption         =   "&Ver Precio Histórico"
               Height          =   405
               Left            =   705
               TabIndex        =   5
               Top             =   2580
               Width           =   1755
            End
            Begin VB.Label LblIdprod 
               Caption         =   "LblIdprod"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   4545
               TabIndex        =   28
               Tag             =   "a"
               Top             =   390
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Stock Maximo"
               Height          =   195
               Index           =   5
               Left            =   705
               TabIndex        =   23
               Top             =   1770
               Width           =   1005
            End
            Begin VB.Label LblUM 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblUM"
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
               Left            =   2250
               TabIndex        =   22
               Tag             =   "a"
               Top             =   1005
               Width           =   1320
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tope Max"
               Height          =   195
               Index           =   4
               Left            =   4590
               TabIndex        =   21
               Top             =   1395
               Width           =   720
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Stock Mínimo"
               Height          =   195
               Index           =   2
               Left            =   705
               TabIndex        =   20
               Top             =   2100
               Width           =   990
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Descripción"
               Height          =   195
               Index           =   0
               Left            =   705
               TabIndex        =   19
               Top             =   780
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Uni. Medida"
               Height          =   195
               Index           =   1
               Left            =   705
               TabIndex        =   18
               Top             =   1110
               Width           =   855
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               Caption         =   "Precio Tope"
               Height          =   195
               Left            =   705
               TabIndex        =   17
               Top             =   1440
               Width           =   870
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Producto"
               Height          =   195
               Index           =   3
               Left            =   705
               TabIndex        =   16
               Top             =   450
               Width           =   645
            End
            Begin VB.Label LblDesProd 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDesProd"
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
               Left            =   2250
               TabIndex        =   15
               Tag             =   "a"
               Top             =   690
               Width           =   8955
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Asignación de Precios"
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
            Height          =   255
            Left            =   60
            TabIndex        =   11
            Top             =   30
            Width           =   11700
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6600
         Left            =   45
         TabIndex        =   7
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6210
            Left            =   15
            TabIndex        =   8
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   10954
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Producto"
            Columns(0).DataField=   "descripcion"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "U.M."
            Columns(1).DataField=   "abrev"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Prec. Tope"
            Columns(2).DataField=   "pretop"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tope Max"
            Columns(3).DataField=   "topmax"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Stock Max"
            Columns(4).DataField=   "stockmax"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Stock Min"
            Columns(5).DataField=   "stockmin"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=7064"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6985"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=979"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=900"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2408"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2328"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=514"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2408"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2328"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=514"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2408"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2328"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=514"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2408"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2328"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(37)  =   ":id=62,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(38)  =   ":id=62,.fontname=MS Sans Serif"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(62)  =   "Named:id=33:Normal"
            _StyleDefs(63)  =   ":id=33,.parent=0"
            _StyleDefs(64)  =   "Named:id=34:Heading"
            _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(66)  =   ":id=34,.wraptext=-1"
            _StyleDefs(67)  =   "Named:id=35:Footing"
            _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(69)  =   "Named:id=36:Selected"
            _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(71)  =   "Named:id=37:Caption"
            _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(73)  =   "Named:id=38:HighlightRow"
            _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(75)  =   "Named:id=39:EvenRow"
            _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(77)  =   "Named:id=40:OddRow"
            _StyleDefs(78)  =   ":id=40,.parent=33"
            _StyleDefs(79)  =   "Named:id=41:RecordSelector"
            _StyleDefs(80)  =   ":id=41,.parent=34"
            _StyleDefs(81)  =   "Named:id=42:FilterBar"
            _StyleDefs(82)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Asignación de Precios"
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
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   9
            Top             =   30
            Width           =   11640
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmMantComPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANTCOMPRECIOS.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA EL CONTROL Y DEFINICION DE PRECIOS DE COMPRA PARA LOS ITEMS DEL
'*                    ALMACEN
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 18/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim Rst_ComPre As New ADODB.Recordset              ' VARIABLE PRINCIPAL QUE ALMACENA LA LISTA DE ITEMS CON PRECIO DEFINIDO
Dim QueHace As Integer                             ' INDICA EN QUE MODO SE ENCUENTRA EL FORMULARIO 1 = NUEVO; 2 = MODIFICA; 3 = SOLO LECTURA
Dim SeEjecuto As Boolean                           ' VARIABLE PARA CONTROLAR QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim Mostrando As Boolean                           ' VARIABLE UTILIZADA PARA INFORMAR A LOS CONTROLES FlexGrid QUE SE ESTAN AGREGANDO FILAS
Dim CaracteresNumericos As String                  ' VARIABLE QUE ALMACENA LOS CARACTERES NUMERICOS QUE UTILIZARAN ALGUNOS CONTROLES TextBox
Dim CaracteresNumericos2 As String, vStr As String ' VARIABLE QUE ALMACENA LOS CARACTERES NUMERICOS QUE UTILIZARAN ALGUNOS CONTROLES TextBox
Dim vFormatString As String                        ' VARIABLE QUE ALMACENA EL FORMATO DE NUMERO PARA MOSTRAR  LOS DATOS NUMERICOS
Dim vVerifPreDet As Integer                        ' varible para determinar si se verifica el precio del detalle
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO


'*****************************************************************************************************
'* Nombre           : fVerifSiTieneDet
'* Tipo             : FUNCION
'* Descripcion      : VERIFICA SI UN ITEM TIENE PRECIOS ASIGNADOS, ESTA FUNCION DEVUELVE VERDADERO SI
'*                    ENCUENTRA UN ITEM CON PRECIO ASIGNADO
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pIdPro    |  LONG       |  ESPECIFICA EL ID DEL ITEM QUE SE ESTA CONSULTANDO
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Private Function fVerifSiTieneDet(pIdPro As Long) As Boolean
    Dim RsVerif As New ADODB.Recordset
    vStr = "SELECT idpro, fecreg FROM com_preciosdet WHERE idpro = " & pIdPro & ""
    RST_Busq RsVerif, vStr, xCon
    If RsVerif.RecordCount > 0 Then
        fVerifSiTieneDet = True
    Else
        fVerifSiTieneDet = False
    End If
    Set RsVerif = Nothing
End Function

'*****************************************************************************************************
'* Nombre           : PintarBacColorGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ASIGNA COLOR AL CONTROL FG1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub PintarBacColorGrid()
    'BACKCOLO PARA COLUMNAS
    With Fg1
        .Select 1, 4, .Rows - 1, 4
        .FillStyle = flexFillRepeat
        .CellBackColor = &H8000000F
        .Select 1, 1, 1, 1
        .Select 1, 6, .Rows - 1, 6
        .FillStyle = flexFillRepeat
        .CellBackColor = &H8000000F
        .Select 1, 1, 1, 1
    End With
End Sub

'*****************************************************************************************************
'* Nombre           : GrabarDetPrec
'* Tipo             : FUNCION
'* Descripcion      : GRABA LOS PRECIOS ASIGNADOS AL ITEM, ESTA FUNCION DEVUELVE VERDADERO SI LLEGA A
'*                    TENER EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Private Function GrabarDetPrec() As Boolean
    ' VERIFICAMOS QUE LOS DATOS REQUERIDOS SEAN LOS CORRECTOS
    If fVerifDatOblig_Det = True Then
        GrabarDetPrec = False
        Exit Function
    End If
    If fValidacPrecDetalle = True Then
        GrabarDetPrec = False
        Exit Function
    End If
    
    Dim i_grab As Integer
    Dim RstDetPrec As New ADODB.Recordset
    
    xCon.Execute ("DELETE FROM com_preciosdet WHERE idpro = " & NulosN(LblIdprod.Caption) & "")
    RST_Busq RstDetPrec, "SELECT * FROM com_preciosdet", xCon
    
    ' GRABAMOS EL DETALLE
    For i_grab = 1 To Fg1.Rows - 1
        RstDetPrec.AddNew
        RstDetPrec("idpro") = NulosN(LblIdprod.Caption)
        RstDetPrec("fecreg") = CDate(Fg1.TextMatrix(i_grab, 3))
        RstDetPrec("idprov") = NulosN(Fg1.TextMatrix(i_grab, 2))
        RstDetPrec("pretope") = NulosN(Fg1.TextMatrix(i_grab, 4))
        RstDetPrec("precio") = NulosN(Fg1.TextMatrix(i_grab, 5))
        RstDetPrec("obs") = Fg1.TextMatrix(i_grab, 8) '
        RstDetPrec.Update
    Next
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, 2, xHorIni, Time, Date, xCon, NulosN(LblIdprod.Caption)
    
    
    MsgBox "El registro se guardo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    GrabarDetPrec = True
    TabOne1.CurrTab = 0
    vVerifPreDet = 0
End Function

'*****************************************************************************************************
'* Nombre           : fVerifSiYaAgregoItemDet
'* Tipo             : FUNCION
'* Descripcion      : VERIFICA EL QUE PRECIO QUE SE LE ESTE ASIGNADO AL PRODUCTO EN EL DIA ESPECIFICADO
'*                    NO SE HAYA AGREGADO YA, ESTA FUNCION DEVUELVE VERDADERO SI TIENE EXITO EN LA
'*                    BUSQUEDA
'* Paranetros       : NOMBRE    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pIdProd   |  LONG      |  ID DEL PRODUCTO QUE SE VA A CONSULTAR
'*                    pFecReg   |  STRING    |  FECHA DE REGISTRO
'*                    pIdProv   |  LONG      |  ID DEL PROVEEDOR
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Private Function fVerifSiYaAgregoItemDet(pIdProd As Long, pFecReg As String, pIdProv As Long) As Boolean
    Dim i_verif As Integer, ii As Integer
    
    ' RECORREMOS EL CONTROL  FlexGrid Fg1 EN BUSCA DE COINCIDENCIAS
    For i_verif = 1 To Fg1.Rows - 2
        ii = 1
        With Fg1
            If Val(.TextMatrix(i_verif, 2)) = pIdProv And Format(.TextMatrix(i_verif, 3), "dd/mm/yyyy") = Format(pFecReg, "dd/mm/yyyy") Then
                fVerifSiYaAgregoItemDet = True
                MsgBox "Los datos ingresados ya existe.", vbInformation, xTitulo
                Exit For
            Else
                fVerifSiYaAgregoItemDet = False
            End If
        End With
    Next
    If ii = 0 Then
        fVerifSiYaAgregoItemDet = False
    End If
End Function

'*****************************************************************************************************
'* Nombre           : fVerifDatOblig_Det
'* Tipo             : FUNCION
'* Descripcion      : VERICA QUE SE HAYAN INGRESADO LOS DATOS NECESARIOS EL CONTROL FlexGrid Fg1,
'*                    DEVUELVE VERDADERO SI LOS DATOS SON LOS CORRECTOS
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Private Function fVerifDatOblig_Det() As Boolean
    Dim i_veri As Integer
    With Fg1
        For i_veri = 1 To Fg1.Rows - 1
            If Trim(Fg1.TextMatrix(i_veri, 1)) = "NUE" Or Trim(Fg1.TextMatrix(i_veri, 1)) = "" Then
                If Trim(.TextMatrix(i_veri, 3)) = "" Then
                    fVerifDatOblig_Det = True
                    MsgBox "Ingrese una fecha válida en el detalle.", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
                    Fg1.Row = i_veri: Fg1.Col = 3 '
                    Fg1.SetFocus
                    Exit Function
                Else
                    fVerifDatOblig_Det = False
                End If
                If IsDate(.TextMatrix(i_veri, 3)) = True Then
                    If Year(Format(.TextMatrix(i_veri, 3), "dd/mm/yyyy")) < 1990 Then
                        fVerifDatOblig_Det = True
                        MsgBox "Ingrese una fecha válida en el detalle.", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
                        Fg1.Row = i_veri: Fg1.Col = 3 '
                        Fg1.SetFocus
                        Exit Function
                    Else
                        fVerifDatOblig_Det = False
                    End If
                Else
                    fVerifDatOblig_Det = True
                    MsgBox "Ingrese una fecha válida en el detalle.", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
                    Fg1.Row = i_veri: Fg1.Col = 3 '
                    Fg1.SetFocus
                    Exit Function
                End If
                If Year(Format(.TextMatrix(i_veri, 3), "dd/mm/yyyy")) > 1990 Then
                    If Trim(.TextMatrix(i_veri, 3)) = "" Or IsDate(Format(.TextMatrix(i_veri, 3), "dd/mm/yyyy")) = False Then 'FECHA
                        fVerifDatOblig_Det = True
                        MsgBox "Ingrese una fecha válida en el detalle.", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
                        Fg1.Row = i_veri: Fg1.Col = 3 '
                        Fg1.SetFocus
                        Exit Function
                    Else
                        fVerifDatOblig_Det = False
                    End If
                Else
                    fVerifDatOblig_Det = True
                    MsgBox "Verifique bien la fecha en especial el año.", vbInformation, xTitulo
                    Fg1.SetFocus
                    Exit Function
                End If
                .TextMatrix(i_veri, 4) = Format(Val(TxtPreTope.Text), "####.00")
                If Val(.TextMatrix(i_veri, 5)) = 0 Or Trim(.TextMatrix(i_veri, 5)) = "" Then 'PRECIO
                    fVerifDatOblig_Det = True
                    MsgBox "Falta especificar el precio en el detalle.", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
                    Fg1.Row = i_veri: Fg1.Col = 5
                    Fg1.SetFocus
                    Exit Function
                Else
                    fVerifDatOblig_Det = False
                End If
                If Trim(.TextMatrix(i_veri, 2)) = "" Then  ' 'ID PROVEEDOR
                    fVerifDatOblig_Det = True
                    MsgBox "Falta especificar el proveedor en el detalle de precios.", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
                    Fg1.Row = i_veri: Fg1.Col = 7 '
                    Fg1.SetFocus
                    Exit Function
                Else
                    fVerifDatOblig_Det = False
                End If
                If fVerifSiYaAgregoItemDet(Val(TxtCodProd.Text), Trim(.TextMatrix(.Row, 3)), .TextMatrix(.Row, 2)) = True Then
                    fVerifDatOblig_Det = True
                    Fg1.Row = i_veri
                    Fg1.SetFocus
                    Exit Function
                Else
                    fVerifDatOblig_Det = False
                End If
            End If
        Next
    End With
End Function

'*****************************************************************************************************
'* Nombre           : BloqueDetalle
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES Fg1, CmdAgregarPre, TxtObs
'* Paranetros       : NOMBRE    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pBool     |  BOOLEAN   |  INDICA SI SE ACTIVA O DESACTIVA EL CONTTROL
'* Devuelve         :
'*****************************************************************************************************
Private Sub BloqueDetalle(pBool As Boolean)
    Fg1.Enabled = pBool
    CmdAgregarPrec.Enabled = pBool
    TxtObs.Locked = Not pBool
End Sub

'*****************************************************************************************************
'* Nombre           : fValidacPrecDetalle
'* Tipo             : FUNCION
'* Descripcion      : VALIDA QUE EL PRECIO INGRESADO NO SEA MENOR AL PRECIO TOPE, DEVUELVE VERDADERO SI
'*                    EL PRECIO INGRESADO ES MAYOR AL PRECIO TOPE
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Private Function fValidacPrecDetalle() As Boolean
    Dim i_valid As Integer
    ' VALIDACION DE PRECIO
    With Fg1
        For i_valid = 1 To .Rows - 1
            If Trim(Fg1.TextMatrix(i_valid, 1)) = "NUE" Or Trim(Fg1.TextMatrix(i_valid, 1)) = "" Then
                If Val(.TextMatrix(i_valid, 5)) < Val(TxtPreTope.Text) Then '
                    fValidacPrecDetalle = True
                    MsgBox "El precio ingresado en el detalle es menor al precio tope" & vbCrLf _
                        & "No puede ingresar un precio menor a precio tope.", vbInformation, xTitulo
                    .Row = i_valid: .Col = 5 '
                    .SetFocus
                    Exit Function
                Else
                    fValidacPrecDetalle = False
                End If
            End If
        Next
    End With
End Function

'*****************************************************************************************************
'* Nombre           : fJalarDatProv
'* Tipo             : FUNCION
'* Descripcion      : MUESTRA EL NOMBRE DEL PROVEEDOR ESPECIFICADO, DEVUELVE UNA CADENA CON EL NOMBRE
'*                    DEL PROVEEDOR
'* Paranetros       : NOMBRE    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pIdProv   |  LONG      |  ESPECIFICA EL ID DEL PROVEEDOR
'* Devuelve         : STRING
'*****************************************************************************************************
Private Function fJalarDatProv(pIdProv As Long) As String
    Dim RstProv As New ADODB.Recordset
    vStr = "SELECT id, nombre FROM mae_prov WHERE id = " & pIdProv & ""
    RST_Busq RstProv, vStr, xCon
    If RstProv.RecordCount > 0 Then
        fJalarDatProv = NulosC(RstProv("nombre"))
    End If
    Set RstProv = Nothing
End Function

'*****************************************************************************************************
'* Nombre           : LlenarGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLENA EL CONTROL FlexGrid Fg1 CON LOS PRECIOS ASIGNADOS AL PROVEEDOR ESPECIFICADO
'* Paranetros       : NOMBRE    |  TIPO     |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pIdProd   |  LONG     |  ESPECIFICA EL ID DEL PROVEEDOR
'* Devuelve         :
'*****************************************************************************************************
Private Sub LlenarGrid(pIdProd As Long)
    LimpiarGrid
    Dim Rst_Grid As New ADODB.Recordset
    ' BUSCAMOS LOS PRECIOS ASIGNADOS AL PROVEEDOR ESPECIFICADO
    vStr = "SELECT * FROM com_preciosdet WHERE idpro = " & pIdProd & " ORDER BY fecreg DESC"
    RST_Busq Rst_Grid, vStr, xCon
    If Rst_Grid.RecordCount > 0 Then
        Rst_Grid.MoveFirst
        Fg1.AddItem ""
        Do While Not Rst_Grid.EOF
            With Fg1
                If .TextMatrix(1, 3) <> "" Then .AddItem ""
                .TextMatrix(.Rows - 1, 1) = "GRA"
                .TextMatrix(.Rows - 1, 2) = NulosN(Rst_Grid("idprov"))   'ID PROVEEDOR
                .TextMatrix(.Rows - 1, 3) = Format(NulosC(Rst_Grid("fecreg")), "dd/mm/yyyy")
                .TextMatrix(.Rows - 1, 4) = NulosN(Rst_Grid("pretope"))
                .TextMatrix(.Rows - 1, 4) = Format(.TextMatrix(.Rows - 1, 4), "0.0000")
                .TextMatrix(.Rows - 1, 5) = NulosN(Rst_Grid("precio"))
                .TextMatrix(.Rows - 1, 5) = Format(.TextMatrix(.Rows - 1, 5), "0.0000")
                'DIFERENCIA DE PRECIO
                .TextMatrix(.Rows - 1, 6) = Abs(Format(NulosN(Rst_Grid("pretope")) - NulosN(Rst_Grid("precio")), "####0.00"))
                .TextMatrix(.Rows - 1, 6) = Format(.TextMatrix(.Rows - 1, 6), "0.0000")
                .TextMatrix(.Rows - 1, 7) = fJalarDatProv(NulosN(Rst_Grid("idprov"))) '
                .TextMatrix(.Rows - 1, 8) = Rst_Grid("obs") '
            End With
            Rst_Grid.MoveNext
        Loop
    End If
    Set Rst_Grid = Nothing
    If Fg1.Rows >= 2 Then
        PintarBacColorGrid
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : LimpiarGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LIMPIA EL CONTROL FlexGrid Fg1 Y LO PREPARA PARA EL INGRESO DE NUEVOS DATOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub LimpiarGrid()
    Fg1.Clear
    Fg1.Rows = 1
    Fg1.FormatString = vFormatString
    
    Fg1.ColComboList(7) = "|..."  ' 'prov
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
        
    confiGrid
End Sub

'*****************************************************************************************************
'* Nombre           : confiGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL CONTROL FlexGrid Fg1 PARA EL INGRESO DE DATOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub confiGrid()
    Fg1.ColWidth(1) = 0 ' 'pa verif
    Fg1.ColWidth(2) = 0 ' 'idprov
    Fg1.ColWidth(8) = 0 ' 'obs
End Sub

'*****************************************************************************************************
'* Nombre           : actualizarStoc_alminvent
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pIdProd   |  LONG       |  ID DEL PRODUCTO QUE SE VA A CONSULTAR
'*                    pStocMax  |  DOUBLE     |  INDICA EL STOCK MAXIMO DEL PRODUCTO
'*                    pStocMin  |  DOUBLE     |  INDICA EL STOCK MINIMO DEL PRODUCTO
'* Devuelve         :
'*****************************************************************************************************
Sub actualizarStoc_alminvent(pIdProd As Long, pStocMax As Double, pStocMin As Double)
    vStr = "UPDATE alm_inventario SET stckmax = " & pStocMax & ", stckmin = " & pStocMin & "" _
        & " WHERE id = " & pIdProd & ""
    xCon.Execute (vStr)
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO ACTUAL EN LA PESTAÑA DETALLE DEL
'*                    FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    If Rst_ComPre.RecordCount = 0 Then Exit Sub
    LblIdprod.Caption = Rst_ComPre("idpro")
    TxtCodProd.Text = Rst_ComPre("codpro")
    LblDesProd.Caption = NulosC(Rst_ComPre("descripcion"))
    LblUM.Caption = NulosC(Rst_ComPre("abrev"))
    TxtPreTope.Text = Format(NulosN(Rst_ComPre("pretop")), "####0.00")
    TxtTopeMax.Text = Format(NulosN(Rst_ComPre("topmax")), "####0.00")
    TxtStocMax.Text = Format(NulosN(Rst_ComPre("stockmax")), "####0.00")
    TxtStocMin.Text = Format(NulosN(Rst_ComPre("stockmin")), "####0.00")
    
    ' LLENA EL CONTROL FlexGrid Fg1
    LlenarGrid Val(LblIdprod.Caption)
End Sub

'*****************************************************************************************************
'* Nombre           : fVerifSiSeRegProd
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : VERIFICA SI EL ITEM ESPECIFICADO TIENE LISTA DE PRECIOS ASIGNADA, ESTA FUNCION
'*                    DEVUELVE VERDADERO SI ENCUENTRA PRECIOS ASIGNADOS
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Private Function fVerifSiSeRegProd() As Boolean
    Dim rst_verif As New ADODB.Recordset
    vStr = "SELECT idpro from com_precios WHERE idpro = " & Val(TxtCodProd.Text) & ""
    RST_Busq rst_verif, vStr, xCon
    If rst_verif.RecordCount > 0 Then
        If NulosN(rst_verif.Fields(0)) > 0 Then
            fVerifSiSeRegProd = True
        Else
            fVerifSiSeRegProd = False
        End If
    Else
        fVerifSiSeRegProd = False
    End If
    Set rst_verif = Nothing
End Function

'*****************************************************************************************************
'* Nombre           : fVerifDatosObligat
'* Tipo             : FUNCION
'* Descripcion      : VERIFICA QUE SE HAYAN INGRESADO LOS DATOS OBLIGATORIOS PARA PODER GUARDAR EL
'*                    REGISTRO, ESTA FUNCION DEVUELVE VERDADERO SI TODOS LOS DATOS SON CORRECTOS
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Private Function fVerifDatosObligat() As Boolean
    Dim i_veri As Integer
    If Trim(TxtCodProd.Text) = "" Or Trim(LblDesProd.Caption) = "" Then
        fVerifDatosObligat = True
        MsgBox "Falta Seleccionar el producto.", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
        TxtCodProd.SetFocus
        Exit Function
    Else
        fVerifDatosObligat = False
    End If
    If NulosN(TxtPreTope.Text) = 0 Then
        fVerifDatosObligat = True
        MsgBox "Falta especificar el precio tope.", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
        TxtPreTope.SetFocus
        Exit Function
    Else
        fVerifDatosObligat = False
    End If
    If NulosN(TxtTopeMax.Text) = 0 Then
        fVerifDatosObligat = True
        MsgBox "Falta especificar el valor del tope maximo.", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
        TxtTopeMax.SetFocus
        Exit Function
    Else
        fVerifDatosObligat = False
    End If
End Function

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE NUEVOS DATOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Precios"
    Bloquea True
    Blanquea
    
    BloqueDetalle False
    TxtCodProd.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCERO DE AGREGAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Bloquea False
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    ActivaTool
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA com_precios, DEVUELVE VERDADERO SI TIENE EXITO AL
'*                    GRABAR
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    ' VERIFICAMOS QUE SE HAYAN INGRESADO LOS DATOS NECESARIOS
    If fVerifDatosObligat = True Then
        Exit Function
    End If
    
    Dim xId As Double
    Dim i_grab As Integer
    Dim RstCab As New ADODB.Recordset
    Dim RstDetPrec As New ADODB.Recordset
            
    On Error GoTo LaCague
    xCon.BeginTrans
    
    If QueHace = 1 Then ' SI SE ESTA AGREGANDO UN NUEVO REGISTRO
        ' VERIFICAMOS QUE NO SE HAYAN REGISTRADO PRECIOS PARA EL ITEM SELECCIONADO
        If fVerifSiSeRegProd = True Then
            MsgBox "El producto seleccionado ya está registrado.", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtCodProd.SetFocus
            Exit Function
        End If
        RST_Busq RstCab, "SELECT * FROM com_precios", xCon
        RST_Busq RstDetPrec, "SELECT * FROM com_preciosdet", xCon
        RstCab.AddNew
        RstCab("idpro") = NulosN(LblIdprod.Caption)
        xId = NulosN(LblIdprod.Caption)
    Else                ' SI SE ESTA MODIFICANDO UN REGISTRO
        xId = NulosN(Rst_ComPre("idpro"))
        RST_Busq RstCab, "SELECT * FROM com_precios WHERE idpro = " & xId & "", xCon
    End If
    RstCab("pretop") = NulosN(TxtPreTope.Text)
    RstCab("topmax") = NulosN(TxtTopeMax.Text)
    RstCab("stockmax") = NulosN(TxtStocMax.Text)
    RstCab("stockmin") = NulosN(TxtStocMin.Text)
    RstCab.Update
    
    ' ACTUALIZA EL STOCK MAX Y MIN DE LA TABAL ALM_INVENTARIO
    actualizarStoc_alminvent NulosN(LblIdprod.Caption), NulosN(TxtStocMax.Text), NulosN(TxtStocMin.Text)
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    Grabar = True
    MsgBox "El registro se guardo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Exit Function

LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA com_precios
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    If Rst_ComPre.RecordCount = 0 Then
        MsgBox "No hay registro para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de eliminar el precio registrado ", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
       
    If Rpta = vbYes Then
        ' CONFIRMAMOS QUE EL USUARIO DESEA ELIMINAR LOS DATOS
        
        xCon.Execute "DELETE FROM com_preciosdet WHERE idpro = " & NulosN(Rst_ComPre("idpro")) & ""
        
        xCon.Execute "DELETE FROM com_precios WHERE idpro = " & NulosN(Rst_ComPre("idpro")) & ""
        
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & NulosN(Rst_ComPre("idpro")) & " AND idform = " & IdMenuActivo

        Rst_ComPre.Requery
        Dg1.Refresh
        MsgBox "El precio seleccionado se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    If Rst_ComPre.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbInformation, xTitulo
        Exit Sub
    End If

    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Precios"
    QueHace = 2
    xHorIni = Time
    Bloquea True
    TxtCodProd.Locked = True
    MuestraSegundoTab
    BloqueDetalle False
    TxtCodProd.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITE EFECTUAR UNA BUSQUEDA EN EL RECORSET Rst_ComPre
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    TabOne1.CurrTab = 0
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Producto":      xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "2500":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id Prod.":      xCampos(1, 1) = "idpro":        xCampos(1, 2) = "1200":    xCampos(1, 3) = "N"
        
    xform.SQLCad = "SELECT alm_inventario.descripcion, com_precios.idpro " _
        & " FROM alm_inventario RIGHT JOIN com_precios ON alm_inventario.id = com_precios.idpro "
    
    xform.Titulo = "Buscando precios"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        Rst_ComPre.MoveFirst
        Rst_ComPre.Find "idpro = " & xRs("idpro") & ""
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA DESACTIVA LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
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
    Toolbar1.Buttons(14).Enabled = Not Toolbar1.Buttons(14).Enabled
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BLANQUEA LOS CONTROLES TextBox DEL FORMULARIO PARA EL INGRESO DE DATOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    Dim obj As Object
    TxtCodProd.Text = ""
    ' RECORRE LOS CONTROLES DE FORMULARIO CUYO IDENTIFICADOR tag = "a"
    For Each obj In Me.Controls
        If obj.Tag = "a" Then
            obj = ""
        End If
    Next
    LimpiarGrid
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA LOS CONTROLES TextBox DEL FORMULARIO PARA EL INGRESO DE DATOS
'* Paranetros       : NOMBRE    |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pBool     |  BOOLEAN     |  ESPECIFICA SI SE ACTIVA O DESACTIVA LOS CONTROLES
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea(pBool As Boolean)
    Dim obj As Object
    ' RECORRE LOS CONTROLES DE FORMULARIO CUYO IDENTIFICADOR tag = "a"
    For Each obj In Me.Controls
        If obj.Tag = "a" And TypeName(obj) = "TextBox" Then
            obj.Locked = Not pBool
        End If
    Next
End Sub

Private Sub CmdAgregarPrec_Click()
    If Fg1.Rows >= 2 Then
        If Fg1.TextMatrix(1, 3) <> "" Then
            If CmdAgregarPrec.Caption = "Agregar Precios" Then
            
                xHorIni = Time
            
                CmdAgregarPrec.Caption = "Grabar Detalle Precios"
                CmdQuitarPrecio.Enabled = True
                Fg1.AddItem "" '
                Fg1.TextMatrix(Fg1.Rows - 1, 1) = "NUE"
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = Val(TxtPreTope.Text)
                Fg1.Row = Fg1.Rows - 1: Fg1.Col = 3
                PintarBacColorGrid
                Fg1.SetFocus
            Else 'grabar
                If GrabarDetPrec = True Then
                    CmdAgregarPrec.Caption = "Agregar Precios"
                    CmdQuitarPrecio.Enabled = False
                End If
                TabOne1.TabEnabled(0) = True
                Toolbar1.Enabled = True
            End If
        End If
    ElseIf Fg1.Rows = 1 Then
        TabOne1.TabEnabled(0) = False
        Toolbar1.Enabled = False
        If CmdAgregarPrec.Caption = "Agregar Precios" Then
            
            xHorIni = Time
            
            CmdAgregarPrec.Caption = "Grabar Detalle Precios"
            CmdQuitarPrecio.Enabled = True
            
            Fg1.AddItem "" '
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = "NUE"
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Val(TxtPreTope.Text)
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(Fg1.TextMatrix(Fg1.Rows - 1, 4), "0.0000")
            Fg1.Row = Fg1.Rows - 1: Fg1.Col = 3
            vVerifPreDet = 1
            Fg1.SetFocus
        Else 'grabar
            If GrabarDetPrec = True Then
                CmdAgregarPrec.Caption = "Agregar Precios"
                CmdQuitarPrecio.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub CmdBusProd_Click()
    ' EJECUTA LA BUSQUEDA DE ITEM
    If QueHace = 3 Then Exit Sub
    If QueHace = 2 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    vStr = "SELECT alm_inventario.codpro, alm_inventario.id, alm_inventario.descripcion, mae_unidades.abrev" _
        & " FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed where alm_inventario.activo=-1"

    xform.SQLCad = vStr
    
    xform.Titulo = "Buscando precios"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCodProd.Text = xRs("codpro")
        LblIdprod.Caption = xRs("id")
        LblDesProd.Caption = xRs("descripcion")
        LblUM.Caption = xRs("abrev")
        TxtPreTope.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdQuitarPrecio_Click()
    On Error Resume Next
    If Fg1.TextMatrix(Fg1.Rows - 1, 1) = "NUE" Then
        Fg1.Row = Fg1.Rows - 1
        Fg1.RemoveItem Fg1.Row
        CmdAgregarPrec.Caption = "Agregar Precios"
        CmdQuitarPrecio.Enabled = False
        Toolbar1.Enabled = True
        TabOne1.TabEnabled(0) = True
        vVerifPreDet = 0
    End If
End Sub

Private Sub CmdVerPrecHist_Click()
    ' INVOCA AL FORMULARIO HISTORICO DE PRECIOS PARA VER LA HISTORIA DE LOS PRECIOS DEL ITEM
    If NulosN(LblIdprod.Caption) = 0 Then
        MsgBox "Seleccione el Item", vbExclamation, xTitulo
        Exit Sub
    End If
    
    Dim xfrm As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    xfrm.PreciosHistoricos xCon, CStr(LblIdprod.Caption), True, ""
    Set xfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            TabOne1.CurrTab = 1
        Case 45 'INSERTAR
            TabOne1.CurrTab = 1
            Nuevo
        Case 46 'ELIMINAR
            Eliminar
    End Select
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(Rst_ComPre("idpro")), xCon
    End If
End Sub

Private Sub fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    ' EJECUTA LA BUSQUEDA DE UN PROVEEDOR EN EL CONTROL FlexGrid Fg1
    If Fg1.TextMatrix(Row, 1) = "GRA" Then
        MsgBox "No puede modificar el registro seleccionado por que es un dato histórico.", vbInformation, xTitulo
        Exit Sub
    End If
    If IsDate(Fg1.TextMatrix(Fg1.Row, 3)) = False Then
        MsgBox "Ingrese una fecha valida.", vbInformation, xTitulo
        Fg1.Row = Fg1.Row: Fg1.Col = 3
        Exit Sub
    End If
    On Error GoTo ERROR
    Dim xCampos() As String
    Dim N_TITULO As String
    Dim N_ORDENADO As String
    Dim N_CAMPO_BUSCA As String
    Dim T_SQL As String
    If Col <> 7 Then Exit Sub
    
    ReDim xCampos(2, 3) As String
    xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "R.U.C.":   xCampos(1, 1) = "numruc":    xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "Codigo":   xCampos(2, 1) = "id":        xCampos(2, 2) = "800":   xCampos(2, 3) = "N"
    
    T_SQL = "SELECT id, nombre,numruc FROM mae_prov order by nombre asc"
    
    N_TITULO = "Buscando Proveedores"
    N_ORDENADO = "nombre"
    N_CAMPO_BUSCA = "nombre"

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    xform.Titulo = N_TITULO
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = N_ORDENADO
    xform.CampoBusca = N_CAMPO_BUSCA
    xform.SQLCad = T_SQL
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount > 0 Then
            If fVerifSiYaAgregoItemDet(Val(TxtCodProd.Text), Trim(Fg1.TextMatrix(Fg1.Row, 3)), NulosC(xRs("id"))) = False Then
                Fg1.TextMatrix(Fg1.Row, 2) = xRs("id")
                Fg1.TextMatrix(Fg1.Row, 7) = NulosC(xRs("nombre")) '
            End If
        End If
    End If
salir:
    Set xform = Nothing
    Set xRs = Nothing

    Exit Sub
ERROR:
    Set xform = Nothing
    Set xRs = Nothing
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    ' VERIFICA QUE EL PRECIO INGRESADO SEA EL CORRECTO
    If Fg1.TextMatrix(Row, 1) = "GRA" Then
        Exit Sub
    End If
    If vVerifPreDet = 0 Then Exit Sub
    If Col <> 5 Then Exit Sub
    If Val(TxtPreTope.Text) > 0 Then
        If Val(Fg1.TextMatrix(Fg1.Row, Fg1.Col)) < Val(TxtPreTope.Text) Then
            MsgBox "El precio ingresado en el detalle es menor al precio tope" & vbCrLf _
                & "No puede ingresar un precio menor a precio tope.", vbInformation, xTitulo
            Fg1.TextMatrix(Fg1.Row, Fg1.Col) = Val(TxtPreTope.Text)
        Else
            Fg1.TextMatrix(Row, 6) = Abs(Val(Fg1.TextMatrix(Row, 4)) - Val(Fg1.TextMatrix(Row, 5)))
        End If
        Fg1.TextMatrix(Row, 5) = Format(Fg1.TextMatrix(Row, 5), "0.0000")
        Fg1.TextMatrix(Row, 6) = Format(Fg1.TextMatrix(Row, 6), "0.0000")
    End If
End Sub

Private Sub fg1_Click()
    If Fg1.Row < 1 Then Exit Sub
    vVerifPreDet = 1
    TxtObs.Text = Fg1.TextMatrix(Fg1.Row, 8) '
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Fg1.Col = 4 Or Fg1.Col = 6 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon

        
        vStr = "SELECT com_precios.idpro, alm_inventario.descripcion, mae_unidades.abrev, com_precios.pretop, com_precios.topmax, com_precios.stockmax, com_precios.stockmin, " _
            & " alm_inventario.codpro FROM mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN com_precios ON alm_inventario.id = com_precios.idpro) ON mae_unidades.id = alm_inventario.idunimed"

        RST_Busq Rst_ComPre, vStr, xCon
        
        Set Dg1.DataSource = Rst_ComPre
        
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJCUTARSE CUANDO SE CARGUE EL FORMULARIO
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    CaracteresNumericos = "0123456789." & Chr(8)
    CaracteresNumericos2 = "0123456789." & Chr(8) & Chr(13)
    Mostrando = False
    vFormatString = Fg1.FormatString
    Blanquea
    confiGrid
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        vVerifPreDet = 0
        If Rst_ComPre.State = 0 Then Exit Sub
        If Rst_ComPre.RecordCount = 0 And QueHace <> 1 Then
            Cancel = 1
            Exit Sub
        End If
        If QueHace = 3 Then MuestraSegundoTab
        BloqueDetalle True
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    vVerifPreDet = 0
    
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Rst_ComPre.Requery
            Dg1.Refresh
            Cancelar
        End If
    End If
    If Button.Index = 6 Then
        Cancelar
    End If
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 12 Then
        Dim vArr(1 To 6, 1 To 4) As String
        'DES DEL PROD
        vArr(1, 1) = Trim(LblDesProd.Caption)
        vArr(1, 2) = LblDesProd.Width
        vArr(1, 3) = Label3(0).Caption
        vArr(1, 4) = 2000
        'UM
        vArr(2, 1) = Trim(LblUM.Caption)
        vArr(2, 2) = LblUM.Width
        vArr(2, 3) = Label3(1).Caption
        vArr(2, 4) = 2000
        'PRECIO TOPE
        vArr(3, 1) = Trim(TxtPreTope.Text)
        vArr(3, 2) = TxtPreTope.Width
        vArr(3, 3) = Label33.Caption
        vArr(3, 4) = 2000
        'TOPE MAXIMO
        vArr(4, 1) = Trim(TxtTopeMax.Text)
        vArr(4, 2) = TxtTopeMax.Width
        vArr(4, 3) = Label3(4).Caption
        vArr(4, 4) = 2000
        'STOCK MAXIMO
        vArr(5, 1) = Trim(TxtStocMax.Text)
        vArr(5, 2) = TxtStocMax.Width
        vArr(5, 3) = Label3(5).Caption
        vArr(5, 4) = 2000
        'STOC MINIMO
        vArr(6, 1) = Trim(TxtStocMin.Text)
        vArr(6, 2) = TxtStocMin.Width
        vArr(6, 3) = Label3(2).Caption
        vArr(6, 4) = 2000
        FrmPrintAsignacPrec.propTitulo1 = "Reporte de Asignación de Precios"
        FrmPrintAsignacPrec.Detalle Fg1, vArr
        FrmPrintAsignacPrec.Show
    End If
    If Button.Index = 14 Then
        Unload Me
        Set Rst_ComPre = Nothing
    End If
End Sub

Private Sub TxtCodProd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then 'F5
        CmdBusProd_Click
    End If
End Sub

Private Sub TxtObs_Change()
    If Fg1.TextMatrix(Fg1.Row, 1) = "GRAB" Then Exit Sub
    Fg1.TextMatrix(Fg1.Row, 8) = TxtObs.Text '
End Sub

Private Sub TxtPreTope_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtPreTope_Validate(Cancel As Boolean)
    If Trim(TxtTopeMax.Text) = "" Then
        TxtTopeMax.Text = Val(TxtPreTope.Text)
    ElseIf Trim(TxtTopeMax.Text) <> "" And Val(TxtTopeMax.Text) > 0 Then
        If Val(TxtPreTope.Text) > Val(TxtTopeMax.Text) Then
            MsgBox "Ingrese un valor menor o ingual a tope maximo.", vbOKCancel + vbInformation + vbDefaultButton1, xTitulo
            TxtPreTope.Text = Val(TxtTopeMax.Text)
        End If
    End If
End Sub

Private Sub TxtStocMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtStocMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtStocMin_Validate(Cancel As Boolean)
    If Val(TxtStocMin.Text) > Val(TxtStocMax.Text) Then
        MsgBox "El stock mínimo debe ser menor al stock máximo.", vbInformation, xTitulo
        TxtStocMin.Text = 0
    End If
End Sub

Private Sub TxtTopeMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTopeMax_Validate(Cancel As Boolean)
    If Val(TxtTopeMax.Text) < Val(TxtPreTope.Text) Then
        MsgBox "Ingrese un valor mayor o igual al precio tope.", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
        TxtTopeMax.Text = Val(TxtPreTope.Text)
    End If
End Sub
