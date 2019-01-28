VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmManControlPers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producción - Control de Personal"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6570
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
            Picture         =   "FrmManControlPers.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlPers.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlPers.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlPers.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlPers.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlPers.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlPers.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlPers.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlPers.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlPers.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlPers.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlPers.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlPers.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7890
      _ExtentX        =   13917
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
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Listado"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
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
      Height          =   6630
      Left            =   15
      TabIndex        =   1
      Top             =   375
      Width           =   7875
      _cx             =   13891
      _cy             =   11695
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
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6210
         Left            =   8520
         TabIndex        =   5
         Top             =   375
         Width           =   7785
         Begin VB.Frame Frame3 
            Height          =   4350
            Left            =   345
            TabIndex        =   7
            Top             =   555
            Width           =   7005
            Begin VB.CommandButton cb 
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   1980
               Picture         =   "FrmManControlPers.frx":2B10
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   510
               Width           =   225
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   0
               Left            =   1020
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   9
               Text            =   "txt_cb(0)"
               ToolTipText     =   "Ingrese DNI del Personal"
               Top             =   480
               Width           =   1215
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   2280
               Left            =   540
               TabIndex        =   13
               Top             =   1290
               Width           =   5355
               _cx             =   9446
               _cy             =   4022
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   1
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManControlPers.frx":2C42
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
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Seleccionar"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Left            =   540
               TabIndex        =   14
               Top             =   1065
               Width           =   1155
            End
            Begin VB.Label lbl_cb_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb_cod(0)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Index           =   0
               Left            =   5295
               TabIndex        =   12
               Top             =   180
               Visible         =   0   'False
               Width           =   1290
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(0)"
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
               Index           =   0
               Left            =   2235
               TabIndex        =   11
               Top             =   480
               Width           =   4035
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nombres"
               Height          =   195
               Index           =   0
               Left            =   270
               TabIndex        =   10
               Top             =   585
               Width           =   630
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Control de Personal"
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
            Left            =   75
            TabIndex        =   6
            Top             =   75
            Width           =   7605
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6210
         Left            =   45
         TabIndex        =   2
         Top             =   375
         Width           =   7785
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   5775
            Left            =   0
            TabIndex        =   3
            Top             =   345
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   10186
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "IdPer"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "IdEmp"
            Columns(1).DataField=   "idemp"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Empleado"
            Columns(2).DataField=   "nomemp"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "T.D."
            Columns(3).DataField=   "abrev"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Nº Documento"
            Columns(4).DataField=   "numdoc"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Nº Funciones"
            Columns(5).DataField=   "totalfunc"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1111"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1032"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(15)=   "Column(2).Width=7250"
            Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=7170"
            Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(21)=   "Column(3).Width=979"
            Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=900"
            Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(27)=   "Column(4).Width=2328"
            Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=2249"
            Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(31)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(33)=   "Column(5).Width=2223"
            Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=2143"
            Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Control de Personal"
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
            TabIndex        =   4
            Top             =   45
            Width           =   7905
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
      End
   End
End
Attribute VB_Name = "FrmManControlPers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANCONTROLPERS.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : PERMITE CONFIGURAR LOS CARGOS DEL TRABAJADOR EN EL SISTEMA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 02/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer                  ' ESPECIFICA EL ESTADO ACTUAL DEL FORMULARIO
Dim RstFrm As New ADODB.Recordset       ' RECORDSET QUE ALMACENARA LOS DATOS DE LA TABLA pro_emp
Dim Mostrando As Boolean                ' CONTROLA EL INGRESO DE UNA FILA EN EL CONTROL FLEX GRID
Dim SeEjecuto As Boolean                ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim mIdRegistro&                        ' identificador del registro
Dim fOrdenLista As Boolean              ' especfica el orden de la lista de la consulta
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA LA BUSQUEDA DE UN EMPLEADO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Empleado":     xCampos(0, 1) = "nomemp":     xCampos(0, 2) = "4000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "T.D.":         xCampos(1, 1) = "abrev":      xCampos(1, 2) = "600":     xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Documento": xCampos(2, 1) = "numdoc":     xCampos(2, 2) = "1200":    xCampos(2, 3) = "C"
    
    nSQL = "SELECT pro_emp.id, pro_emp.idemp,  [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nomemp, Count(pro_empdet.idfun) AS totalfunc, mae_dociden.abrev, pla_empleados.numdoc " _
            + vbCr + " FROM mae_dociden RIGHT JOIN ((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) ON mae_dociden.id = pla_empleados.idtipdoc " _
            + vbCr + " GROUP BY pro_emp.id, pro_emp.idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom], mae_dociden.abrev, pla_empleados.numdoc; "
        
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscanco Personal", "nomemp", "nomemp", Principio
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(xRs("id"))

SALIR:
    Set xRs = Nothing
    Exit Sub
    
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTAÑA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    If RstFrm.RecordCount = 0 Then Exit Sub
    Mostrando = True
    
    txt_cb(0).Text = NulosC(RstFrm("numdoc"))
    lbl_cb_cod(0).Caption = NulosN(RstFrm("idemp"))
    lbl_cb(0).Caption = NulosC(RstFrm("nomemp"))
    Fg1.Rows = 1
    pCargarDetalle
    Mostrando = False
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS DEL FORMULARIO
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
'* Nombre           : Habilitar_Obj
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TEXTBOX DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Habilitar_Obj(band As Boolean)
    habilitar_Locked txt_cb, Not band
    habilitar Me.cb, band
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    LimpiaText txt_cb
    LimpiaText lbl_cb
    LimpiaText lbl_cb_cod
    Fg1.Rows = 1
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE INGRESAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Habilitar_Obj False
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    ActivaTool
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA pro_emp, ESTA FUNCION DEVUELVE VERDADERO CUANDO
'*                    TIENE EXITO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim xId As Double
    Dim A As Integer
    
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el Registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xFila&
    
    On Error GoTo LaCague
    xCon.BeginTrans
    
    If QueHace = 1 Then ' NUEVO
        xId = HallaCodigoTabla("pro_emp", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pro_emp", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else                ' MODIFICAR
        xId = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM pro_emp WHERE id = " & xId & "", xCon
        xCon.Execute "DELETE FROM pro_empdet WHERE idper = " & xId & ""
    End If
    
    mIdRegistro = xId
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM pro_empdet", xCon
    RstCab("idemp") = NulosN(lbl_cb_cod(0).Caption)
    RstCab.Update
    
    ' Agregando las funciones
    For xFila = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(xFila, 2)) = -1 Then
            RstDet.AddNew
            RstDet("idper") = xId
            RstDet("idfun") = NulosN(Fg1.TextMatrix(xFila, 3))
            RstDet.Update
        End If
    Next xFila
    
   'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    Grabar = True
    MsgBox "El registro se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    Set RstCab = Nothing
    Exit Function

LaCague:
    Set RstCab = Nothing
    xCon.RollbackTrans
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo "
End Function

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA pro_emp
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    On Error GoTo error
    Dim Rpta As Integer
    Dim xId&
    Dim nSQL As String
    Dim RstBusca As New ADODB.Recordset
    xId = NulosN(RstFrm("id"))
    
    nSQL = "SELECT TOP 1  'Programación Semanal' AS origen, 'Del ' & Format([pro_programa].[fchfin],'dd/mm/yy') & ' Al ' & Format([pro_programa].[fchfin],'dd/mm/yy') AS nombre FROM pro_programa WHERE (((pro_programa.idprog)=" & xId & ")); " _
    + vbCr + " UNION " _
    + vbCr + " SELECT TOP 1 'Producción' AS origen, 'Fecha Producción: ' & Format([pro_produccion].[dia],'dd/mm/yy') AS nombre FROM pro_produccion WHERE (((pro_produccion.idsup)=" & xId & ")); " _
    + vbCr + " UNION " _
    + vbCr + " SELECT TOP 1 'Producción' AS origen, 'Fecha Producción: ' & Format([pro_produccion].[dia],'dd/mm/yy') & '  Producto: ' & [alm_inventario].[descripcion] AS nombre FROM (pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id WHERE (((pro_producciondet.idres)=" & xId & ")); " _
    + vbCr + " UNION " _
    + vbCr + " SELECT TOP 1 'Mantenimiento de Grupos de Trabajo' AS origen, 'Grupo : ' & [pro_grupo].[num] AS nombre FROM pro_grupo INNER JOIN pro_grupodet ON pro_grupo.id = pro_grupodet.idgrupo WHERE (((pro_grupodet.idper)=" & xId & ")); " _
    + vbCr + " UNION " _
    + vbCr + " SELECT TOP 1 'Control de Tarea de Producción' AS origen, 'Fecha Trabajo: ' & Format([pro_controltar].[fchtra],'dd/mm/yy') & '    Area:  ' & [mae_area].[descripcion] AS nombre FROM mae_area INNER JOIN (pro_controltar INNER JOIN pro_controltardet ON pro_controltar.id = pro_controltardet.idctr) ON mae_area.id = pro_controltar.idarea WHERE (((pro_controltardet.idref)=" & xId & ") AND ((pro_controltardet.tipo)=1)); " _
    + vbCr + " UNION " _
    + vbCr + " SELECT TOP 1 'Control de Tarea de Producción en Grupo' AS origen, 'Fecha Trabajo: ' & Format([pro_controltar].[fchtra],'dd/mm/yy') & '    Area:  ' & [mae_area].[descripcion] AS nombre FROM (mae_area INNER JOIN pro_controltar ON mae_area.id = pro_controltar.idarea) INNER JOIN (pro_controltardet INNER JOIN pro_controltardetgr ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON pro_controltar.id = pro_controltardet.idctr WHERE (((pro_controltardet.tipo)=2) AND ((pro_controltardetgr.idper)=" & xId & ")); " _
    + vbCr + " UNION " _
    + vbCr + " SELECT TOP 1 'Control de Tarea de Producción' AS origen, 'Fecha Trabajo: ' & Format([pro_controltar].[fchtra],'dd/mm/yy') & '    Area:  ' & [mae_area].[descripcion] AS nombre FROM mae_area INNER JOIN pro_controltar ON mae_area.id = pro_controltar.idarea WHERE (((pro_controltar.idres)=" & xId & ")); "

    ' si el registro tiene relaciones mostrara un menaje
    RST_Busq RstBusca, nSQL, xCon
    If RstBusca.EOF = False Or RstBusca.BOF = False Or RstBusca.RecordCount <> 0 Then
        MsgBox "El registro no se puede eliminar" + vbCr + "Esta asociado a " & RstBusca("origen") & vbCr & RstBusca("nombre"), vbExclamation, xTitulo
        Set RstBusca = Nothing
        Exit Sub
    End If
    Set RstBusca = Nothing
    
    Rpta = MsgBox("¿Esta seguro de eliminar el registro seleccionado?", vbQuestion + vbYesNo, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_empdet WHERE idper = " & xId & ""
        xCon.Execute "DELETE * FROM pro_emp WHERE id = " & xId & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & IdMenuActivo

        
        RstFrm.Requery
        Dg3.Refresh
        MsgBox "Registro fue eliminado con éxito", vbInformation + vbOKOnly, xTitulo
    End If
    
    TabOne1.CurrTab = 0
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Eliminar", True, "Error al eliminar..."
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA LA EDICION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    TabOne1.CurrTab = 1
    xHorIni = Time
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Control de Personal"
    QueHace = 2
    Habilitar_Obj True
    Blanquea
    MuestraSegundoTab
    txt_cb(0).SetFocus
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
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Control de Personal"
    Habilitar_Obj True
    Blanquea
    pCargarDetalle
    txt_cb(0).SetFocus
End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg3_FilterChange()
    TDB_FiltroGenerar Dg3, RstFrm
End Sub

Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDENTE LAS COLUMNAS DEL CONTROL Dg3
On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg3.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstFrm("id")), xCon
    End If
End Sub

Private Sub Fg1_DblClick()
    If QueHace = 3 Then Exit Sub
    If Fg1.Row < 1 Then Exit Sub
    Fg1.TextMatrix(Fg1.Row, 2) = Not NulosN(Fg1.TextMatrix(Fg1.Row, 2))
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    If Fg1.Col = 2 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
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
        
        Fg1.ColWidth(3) = 0
        Dim nSQL As String
        nSQL = "SELECT pro_emp.id, pro_emp.idemp,  [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nomemp, Count(pro_empdet.idfun) AS totalfunc, mae_dociden.abrev, pla_empleados.numdoc " _
            + vbCr + " FROM mae_dociden RIGHT JOIN ((pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) ON mae_dociden.id = pla_empleados.idtipdoc " _
            + vbCr + " WHERE (((pla_empleados.fchcese) Is Null)) " _
            + vbCr + " GROUP BY pro_emp.id, pro_emp.idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom], mae_dociden.abrev, pla_empleados.numdoc " _
            + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]"

        RST_Busq RstFrm, nSQL, xCon
        
        Set Dg3.DataSource = RstFrm

    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Mostrando = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RstFrm = Nothing
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstFrm.Requery
            Dg3.Refresh
            If RstFrm.RecordCount <> 0 Then
                RstFrm.MoveFirst
                RstFrm.Find "id=" & mIdRegistro
                If RstFrm.EOF = True Then RstFrm.MoveFirst
            End If
            Cancelar
        End If
    End If
    
    If Button.Index = 6 Then
        Cancelar
    End If
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg3
        RstFrm.Filter = ""
        RstFrm.Requery
        
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 12 Then pExportar
    
    If Button.Index = 13 Then TDB_IMPRIMIR Dg3, "IMPRESIÓN", "LISTADO DE PERSONAL"
        
    If Button.Index = 15 Then
        Unload Me
        Set RstFrm = Nothing
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : FUNCION
'* Descripcion      : VERIFICA QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS, ESTA FUNCION DEVUELVE
'*                    VERDADERO SI LOS DATOS SON LOS CORRECTOS
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    If Trim(lbl_cb_cod(0).Caption) = "" Then
        MsgBox "Falta especificar el empleado.", vbInformation, xTitulo
        txt_cb(0).SetFocus
        Exit Function
    End If
    
    fValidarDatos = True
End Function

Private Sub cb_Click(Index As Integer)
    ' EJECUTA LA BUSQUEDA DE UN EMPLEADO
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    Dim nSQL As String
    
    xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nomemp":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "DNI":      xCampos(1, 1) = "numdoc":   xCampos(1, 2) = "1000":    xCampos(1, 3) = "C"
        
    nSQL = "SELECT pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nomemp, pla_empleados.id " _
        + vbCr + " FROM pla_empleados " _
        + vbCr + " WHERE (((pla_empleados.id) Not In (select idemp from pro_emp))) " _
        + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Personal", "nomemp", "nomemp", Principio

    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR

    txt_cb(Index) = xRs.Fields(0) & ""             ' TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & ""     ' NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" ' CODIGO

SALIR:
    Set xRs = Nothing
    Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cb_cod(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If

    If txt_cb(Index).Text = "" Then Exit Sub
    If KeyCode <> 13 Then Exit Sub
    Dim RST_TMP As New ADODB.Recordset
    Dim nSQL As String

    nSQL = "SELECT pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nomemp, pla_empleados.id " _
        + vbCr + " FROM pla_empleados " _
        + vbCr + " WHERE (((pla_empleados.id) Not In (select idemp from pro_emp))) AND pla_empleados.numdoc ='" + Trim(txt_cb(Index).Text) + "'" _
        + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"

    If xCon.State = 0 Then Exit Sub
    RST_Busq RST_TMP, nSQL, xCon
    
    If RST_TMP.State = 0 Then Exit Sub
    If RST_TMP.RecordCount > 0 Then
        txt_cb(Index) = RST_TMP.Fields(0) & ""             ' TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RST_TMP.Fields(1) & ""     ' NOMBRE
        lbl_cb_cod(Index).Caption = RST_TMP.Fields(2) & "" ' CODIGO
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cb_cod(Index).Caption = ""
    End If
    Set RST_TMP = Nothing
    Exit Sub

error:
    Set RST_TMP = Nothing
    SHOW_ERROR Me.Name, "txt_cb_KeyDown(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    Select Case Index
        Case 3: If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else: If validar_numero(KeyAscii) = True And KeyAscii = 46 Then KeyAscii = 0
    End Select
End Sub

Private Sub pCargarDetalle()
    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
    If QueHace <> 1 Then
        nSQL = "SELECT pro_funcion.id, pro_funcion.descripcion, -1 AS activo, pro_funcion.orden FROM pro_funcion INNER JOIN pro_empdet ON pro_funcion.id = pro_empdet.idfun WHERE (((pro_empdet.idper)=" & NulosN(RstFrm("id")) & " )); " _
            + vbCr + " UNION " _
            + vbCr + " SELECT pro_funcion.id, pro_funcion.descripcion, 0 AS activo, pro_funcion.orden FROM pro_funcion WHERE (((pro_funcion.id) Not In (select pro_empdet.idfun from pro_empdet  where pro_empdet.idper=" & NulosN(RstFrm("id")) & "))); "
    Else
        nSQL = "SELECT pro_funcion.id, pro_funcion.descripcion, 0 AS activo, pro_funcion.orden FROM pro_funcion ; "
    End If
    RST_Busq Rst, nSQL, xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.Sort = "orden asc"
        Rst.MoveFirst
    End If
    Do While Not Rst.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst("descripcion"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosN(Rst("activo"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosN(Rst("id"))
        Rst.MoveNext
    Loop
    Set Rst = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : pExportar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTAR A MS EXCEL LOS DATOS DEL RECORDSET RstTmp
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportar()
    TabOne1.CurrTab = 0
    
    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset
    Dim xCampos(3, 3) As String
    
    ' 0 Nombre a Mostrar;
    ' 1 nombre de Campo del Rst;
    ' 2 alineacion(0::derecha, 1::centro, 2::izquierda);
    ' 3 ancho de columna
    ' obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Personal":         xCampos(0, 1) = "nomemp":    xCampos(0, 2) = 0:  xCampos(0, 3) = "4000"
    xCampos(1, 0) = "T.D.":             xCampos(1, 1) = "abrev":     xCampos(1, 2) = 0:  xCampos(1, 3) = "500"
    xCampos(2, 0) = "Número":           xCampos(2, 1) = "numdoc":    xCampos(2, 2) = 0:  xCampos(2, 3) = "1100"
    xCampos(3, 0) = "Total Funciones":  xCampos(3, 1) = "totalfunc": xCampos(3, 2) = 0:  xCampos(3, 3) = "1300"
    Set RstTmp = RstFrm.Clone
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Lista de Personal de Producción", "", "", "Personal de Producción", RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
End Sub
