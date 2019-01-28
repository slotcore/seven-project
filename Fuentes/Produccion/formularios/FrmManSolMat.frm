VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManSolMat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produccion - Solicitud de Materiales"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11880
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
            Picture         =   "FrmManSolMat.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManSolMat.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManSolMat.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManSolMat.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManSolMat.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManSolMat.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManSolMat.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManSolMat.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManSolMat.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManSolMat.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManSolMat.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManSolMat.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManSolMat.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   9
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
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Anular Registro"
               EndProperty
            EndProperty
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Materiales"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Linea"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7080
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   11895
      _cx             =   20981
      _cy             =   12488
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6660
         Left            =   45
         TabIndex        =   13
         Top             =   375
         Width           =   11805
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6135
            Left            =   30
            TabIndex        =   16
            Top             =   480
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   10821
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
            Columns(1).Caption=   "Fecha"
            Columns(1).DataField=   "fchdoc"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numdoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Responsable"
            Columns(3).DataField=   "desresp"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Almacén"
            Columns(4).DataField=   "desalm"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "T.D.Ref."
            Columns(5).DataField=   "destipdocref"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Nº Doc. Ref."
            Columns(6).DataField=   "numdocref"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Anexo"
            Columns(7).DataField=   "anexo"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Estado"
            Columns(8).DataField=   "desestado"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1402"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1323"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2540"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2461"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=3625"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=3545"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=3413"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=3334"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1217"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1138"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=2434"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2355"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=3122"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=3043"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=1958"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1879"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=3"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14,.alignment=2"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15,.alignment=3"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14,.alignment=3"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=74,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Named:id=33:Normal"
            _StyleDefs(73)  =   ":id=33,.parent=0"
            _StyleDefs(74)  =   "Named:id=34:Heading"
            _StyleDefs(75)  =   ":id=34,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(76)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1"
            _StyleDefs(77)  =   "Named:id=35:Footing"
            _StyleDefs(78)  =   ":id=35,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(79)  =   ":id=35,.fgcolor=&H80000012&"
            _StyleDefs(80)  =   "Named:id=36:Selected"
            _StyleDefs(81)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(82)  =   "Named:id=37:Caption"
            _StyleDefs(83)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(84)  =   "Named:id=38:HighlightRow"
            _StyleDefs(85)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(86)  =   "Named:id=39:EvenRow"
            _StyleDefs(87)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(88)  =   "Named:id=40:OddRow"
            _StyleDefs(89)  =   ":id=40,.parent=33"
            _StyleDefs(90)  =   "Named:id=41:RecordSelector"
            _StyleDefs(91)  =   ":id=41,.parent=34"
            _StyleDefs(92)  =   "Named:id=42:FilterBar"
            _StyleDefs(93)  =   ":id=42,.parent=33"
         End
         Begin VB.Label lblperiodo 
            AutoSize        =   -1  'True
            Caption         =   "lblperiodo"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   10020
            TabIndex        =   15
            Top             =   90
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Solicitud"
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
            Left            =   45
            TabIndex        =   14
            Top             =   45
            Width           =   11685
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6660
         Left            =   12540
         TabIndex        =   11
         Top             =   375
         Width           =   11805
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   1
            Left            =   7750
            Picture         =   "FrmManSolMat.frx":2B10
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   450
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Frame FrameItem 
            Caption         =   "[ Detalle de Solicitud ]"
            Height          =   4965
            Left            =   30
            TabIndex        =   29
            Top             =   1700
            Width           =   11780
            Begin VB.Frame Frame3 
               Height          =   4620
               Left            =   10290
               TabIndex        =   30
               Top             =   210
               Width           =   1410
               Begin VB.CommandButton cmd 
                  Caption         =   "&Seleccionar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   3
                  Left            =   60
                  TabIndex        =   34
                  Top             =   510
                  Width           =   1275
               End
               Begin VB.CommandButton cmd 
                  Caption         =   "Eliminar &Todos"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   9
                  Left            =   60
                  TabIndex        =   33
                  Top             =   1530
                  Width           =   1275
               End
               Begin VB.CommandButton cmd 
                  Caption         =   "&Eliminar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   8
                  Left            =   60
                  TabIndex        =   32
                  Top             =   1140
                  Width           =   1275
               End
               Begin VB.CommandButton cmd 
                  Caption         =   "&Agregar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   0
                  Left            =   60
                  TabIndex        =   31
                  Top             =   150
                  Width           =   1275
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   4560
               Index           =   0
               Left            =   60
               TabIndex        =   8
               Top             =   300
               Width           =   10200
               _cx             =   17992
               _cy             =   8043
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
               Rows            =   2
               Cols            =   11
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManSolMat.frx":2C42
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   0   'False
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
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2550
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   3
            Text            =   "TxtNumDoc"
            Top             =   1020
            Width           =   3260
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   2
            Text            =   "TxtNumSer"
            Top             =   1020
            Width           =   1155
         End
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   2
            Left            =   1830
            Picture         =   "FrmManSolMat.frx":2D79
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1350
            Width           =   240
         End
         Begin VB.ComboBox cbEstado 
            Height          =   315
            Left            =   1170
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   390
            Width           =   4665
         End
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   5
            Left            =   11490
            Picture         =   "FrmManSolMat.frx":2EAB
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1380
            Width           =   240
         End
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   4
            Left            =   7740
            Picture         =   "FrmManSolMat.frx":2FDD
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1035
            Width           =   240
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchPro 
            Height          =   300
            Left            =   1170
            TabIndex        =   1
            Top             =   720
            Width           =   1320
            _ExtentX        =   2328
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
         Begin VB.TextBox TxtIdTipDocRef 
            Height          =   300
            Left            =   7095
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "TxtIdTipDocRef"
            Top             =   1005
            Width           =   915
         End
         Begin VB.TextBox txtNumDocRef 
            Height          =   300
            Left            =   7095
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   7
            Text            =   "txtNumDocRef"
            Top             =   1350
            Width           =   4665
         End
         Begin VB.TextBox TxtIdResp 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   4
            Text            =   "TxtIdResp"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtIdAlm 
            Height          =   300
            Left            =   7095
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   5
            Text            =   "txtIdAlm"
            Top             =   420
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblAlmacen 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblAlmacen"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   8055
            TabIndex        =   37
            Top             =   435
            Visible         =   0   'False
            Width           =   3690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   195
            Index           =   3
            Left            =   6060
            TabIndex        =   36
            Top             =   465
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Num. Doc."
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   1065
            Width           =   765
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2400
            Top             =   1140
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Responsable"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1365
            Width           =   930
         End
         Begin VB.Label lblResponsable 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblResponsable"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2115
            TabIndex        =   26
            Top             =   1320
            Width           =   3690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   24
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Doc. Ref."
            Height          =   195
            Index           =   7
            Left            =   6030
            TabIndex        =   23
            Top             =   1395
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Doc. Ref."
            Height          =   195
            Index           =   5
            Left            =   6030
            TabIndex        =   21
            Top             =   1050
            Width           =   1005
         End
         Begin VB.Label lbliddocref 
            AutoSize        =   -1  'True
            Caption         =   "lbliddocref"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   9930
            TabIndex        =   20
            Top             =   1050
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Doc."
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   765
            Width           =   705
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Solicitud"
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
            TabIndex        =   12
            Top             =   75
            Width           =   11685
         End
         Begin VB.Label LblTipDocRef 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipDocRef"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   8055
            TabIndex        =   22
            Top             =   1020
            Width           =   3690
         End
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Insertar"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Ver Receta"
      End
   End
End
Attribute VB_Name = "FrmManSolMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim RstSol As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean               ' para saber cuando se este agregando FILAS AL CONTROL grid de productos
Dim IdMenuActivo As Integer            ' INDICA EL CODIGO DEL MENU ACTIVO
Dim agregados As Integer
Dim xHorIni As Date                    ' ESPECIFICA LA HORA DE INICIO
Dim mCorrelativo As Long               ' para diferenciar la fecha de entrega del pedido cuando se necesite modificar
Dim mIdRegistro&                       ' identificador del registro
Dim mMesActivo As Integer              ' indica el mes activo
Dim fCierrePeriodo As Boolean                        ' --indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim cSQL As String
Dim RstValores As New ADODB.Recordset
Dim CAMBIOGRABAR_ As Double
Dim ESTADOANTERIOR_ As Double
Dim fOrdenLista As Boolean              ' especfica el orden de la lista de la consulta
' ----------------------DEFINICION DE COLUMNAS
Private Enum COLUMNA_
    COLUMNACODIGO_ = 1
    COLUMNAITEM
    COLUMNAUNIMED
    COLUMNASTOCK
    COLUMNACANTIDAD
    COLUMNALOTE
    COLUMNAIDITEM
    COLUMNAIDLOTE
    COLUMNAIDLOTEDET
    COLUMNAIDUNIMED
End Enum
' ----------------------DEFINICION DE ESTADOS
Const ESTADOPENDIENTE_ = 1
Const ESTADOPROCESADO_ = 2
Const ESTADOAPROBADO_ = 3
Const ESTADOANULADO_ = 4

Sub preparaRST(ByRef RST_ As ADODB.Recordset)
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(10, 3) As String
    ' N: Numerico
    ' D: Double
    ' F: Fecha
    ' C: Caracter
    ' L: Logico
    xCampos(0, 0) = "iditem":           xCampos(0, 1) = "N":      xCampos(0, 2) = ""
    xCampos(1, 0) = "cantidad":         xCampos(1, 1) = "D":      xCampos(1, 2) = ""
    xCampos(2, 0) = "cantteo":         xCampos(2, 1) = "D":      xCampos(2, 2) = ""
    xCampos(3, 0) = "idtipo":           xCampos(3, 1) = "N":      xCampos(3, 2) = ""
    xCampos(4, 0) = "idlote":           xCampos(4, 1) = "N":      xCampos(4, 2) = ""
    xCampos(5, 0) = "idlotedet":        xCampos(5, 1) = "N":      xCampos(5, 2) = ""
    xCampos(6, 0) = "canant":           xCampos(6, 1) = "D":      xCampos(6, 2) = ""
    xCampos(7, 0) = "idalm":            xCampos(7, 1) = "N":      xCampos(7, 2) = ""
    xCampos(8, 0) = "idloteant":        xCampos(8, 1) = "N":      xCampos(8, 2) = ""
    xCampos(9, 0) = "idlotedetant":     xCampos(9, 1) = "N":      xCampos(9, 2) = ""
    
    Set RST_ = xFun.CrearRstTMP(xCampos)
    RST_.Open
End Sub

Private Function GrabarIngreso() As Boolean
    Dim A As Integer
    Dim FCHMOV_ As String
    Dim TIPDOC_ As Integer
    Dim NUMSER_ As String
    Dim IDRESP_ As Integer
    Dim IDPROV_ As Integer
    Dim DESPROV_ As String
    Dim IDESTADO_ As Integer
    Dim IDTIPMOV_ As Integer
    Dim IDTIPDOCREF_ As Integer
    Dim IDDOCREF_ As Integer
    Dim IDING_ As Integer
    Dim IDALM_ As Integer
    Dim NUMDOC_ As String
    Dim xRs As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
    
    ' Se llenan los detalles
    IDING_ = 0
    FCHMOV_ = Format(TxtFchPro.valor, "dd/mm/yyyy")
    TIPDOC_ = 71
    NUMSER_ = NulosC(TxtNumSer.Text)
    NUMDOC_ = Format(hallarNumDoc("alm_ingreso", "'" & NulosC(TxtNumSer.Text) & "'", "numser"), "0000000000") 'NulosC(TxtNumDoc.Text)
    IDRESP_ = NulosN(TxtIdResp.Text)
    IDPROV_ = 0
    DESPROV_ = ""
    IDESTADO_ = ESTADOPENDIENTE_
    IDTIPMOV_ = 0
    IDTIPDOCREF_ = 110
    IDDOCREF_ = NulosN(RstSol("id"))
    IDALM_ = NulosN(txtIdAlm.Text)
    ' Se prepara el Recordset
    If xRs.State = 0 Then preparaRST xRs
    limpiarRST xRs
    ' Se llena el recordset
    For A = 1 To fg(0).Rows - 1
        xRs.AddNew
        xRs("iditem") = NulosN(fg(0).TextMatrix(A, COLUMNAIDITEM))
        xRs("cantidad") = NulosN(fg(0).TextMatrix(A, COLUMNACANTIDAD))
        xRs("idlote") = NulosN(fg(0).TextMatrix(A, COLUMNAIDLOTE))
        xRs("idlotedet") = NulosN(fg(0).TextMatrix(A, COLUMNAIDLOTEDET))
        xRs("canant") = 0
        xRs("idloteant") = 0
        xRs("idlotedetant") = 0
        xRs.Update
    Next A
    
    ' Se graba el movimiento
    GrabarIngreso = GrabarMovimiento(FCHMOV_, TIPDOC_, NUMSER_, IDRESP_, IDPROV_, DESPROV_, IDESTADO_, _
                                IDTIPMOV_, IDTIPDOCREF_, IDDOCREF_, IDALM_, xRs, IDING_, NUMDOC_, 6)
End Function

Private Sub cbEstado_Click()
    Dim Rpta As Integer
    Dim MENSAJE_ As String
    Dim xRs As New ADODB.Recordset
    Dim RSTSOL_ As New ADODB.Recordset
    Dim IDRECETA_ As Integer
    Dim CANTIDAD_ As Double
    
    Dim IDSOL_ As Integer
    Dim FCHSOL_ As String
    Dim NUMSER_ As String
    Dim NUMERODOCUMENTO_ As Integer
    Dim NUMDOC_ As String
    Dim IDRESP_ As Integer
    Dim IDTIPDOCREF_ As Integer
    Dim IDDOCREF_ As Integer
    Dim IDITEM_ As Integer
    Dim IDALM_ As Integer
    Dim IDESTADO_ As Integer
    Dim A As Integer

    If Agregando Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    If Not verificarCampos Then
        Agregando = True
        llenarEstado 0, 1, , cbEstado, ESTADOANTERIOR_
        Agregando = False
        Exit Sub
    End If

    IDSOL_ = NulosN(RstSol("id"))
    
    Select Case cbEstado.ItemData(cbEstado.ListIndex)
        Case ESTADOPENDIENTE_ ' Pendiente
            If ESTADOANTERIOR_ > ESTADOPENDIENTE_ Then
                MsgBox "No se puede cambiar el estado a " & cbEstado.Text, vbInformation, xTitulo
                llenarEstado 0, 1, , cbEstado, ESTADOANTERIOR_
            End If
            Exit Sub

        Case ESTADOPROCESADO_ ' Procesado
            If ESTADOANTERIOR_ < ESTADOPROCESADO_ Then
                
                Rpta = MsgBox("Cambiar el estado a " & cbEstado.Text & " bloqueara el registro para su posterior modificación " _
                                    + vbCr + "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
                
                If Rpta = vbYes Then
                    Grabar
                    RstSol.Requery
                    Dg1.Refresh
                    RstSol.Find "id=" & IDSOL_

'                    Rpta = MsgBox("¿Desea generar el registro de salida de almacén?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
'                    If Rpta = vbNo Then Exit Sub
'
'                    If GrabarIngreso Then
'                        Grabar
'                        RstSol.Requery
'                        Dg1.Refresh
'                        RstSol.Find "id=" & IDSOL_
'                    Else
'                        Agregando = True
'                        llenarEstado 0, 1, , cbEstado, ESTADOANTERIOR_
'                        Agregando = False
'                    End If
                Else
                    Agregando = True
                    llenarEstado 0, 1, , cbEstado, ESTADOANTERIOR_
                    Agregando = False
                End If
            Else
                MsgBox "No se puede pasar a un estado " & cbEstado.Text, vbInformation, xTitulo
                Agregando = True
                llenarEstado 0, 1, , cbEstado, ESTADOANTERIOR_
                Agregando = False
            End If
            Exit Sub

        Case ESTADOANULADO_ ' Anulada
            If ESTADOANTERIOR_ = ESTADOPROCESADO_ Then
                If Not verificarCambioEstado(NulosN(RstSol("id")), MENSAJE_) Then
                    MsgBox "No se puede pasar a un estado " & cbEstado.Text, vbInformation, xTitulo
                    Agregando = True
                    llenarEstado 0, 1, , cbEstado, ESTADOANTERIOR_
                    Agregando = False
                Else
                    ' -------------SE CAMBIA DE ESTADO A LA SOLICITUD DE MATERIALES
                    cSQL = "UPDATE alm_ingreso SET alm_ingreso.estado = 2 " _
                        + vbCr + "WHERE (((alm_ingreso.idtipdocref)=110) AND ((alm_ingreso.iddocref)=" & NulosN(RstSol("id")) & "));"
                    ' --------------EJECUTA COMANDO
                    xCon.Execute cSQL
                    ' --------------ACTUALIZA VAR_EDICION
                    cSQL = "SELECT alm_ingreso.id " _
                        + vbCr + "FROM alm_ingreso " _
                        + vbCr + "WHERE (((alm_ingreso.idtipdocref)=110) And ((alm_ingreso.iddocref)=" & NulosN(RstSol("id")) & "))"
                    
                    Set xRs = Nothing
                    RST_Busq xRs, cSQL, xCon
                    If xRs.State = 0 Then Exit Sub
                    If xRs.RecordCount = 0 Then Exit Sub
                    xRs.MoveFirst
                    While Not xRs.EOF
                        GrabarOperacion xIdUsuario, 8, 7, xHorIni, Time, Date, xCon, NulosN(xRs("id"))
                        xRs.MoveNext
                    Wend
                End If
            Else
                MsgBox "No se puede cambiar el estado a " & cbEstado.Text, vbInformation, xTitulo
                Agregando = True
                llenarEstado 0, 1, , cbEstado, ESTADOANTERIOR_
                Agregando = False
            End If
            Exit Sub

    End Select
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim A As Integer
    Dim num As Integer
    Dim Rpta As Integer
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim MENSAJE_ As String
    Dim IDRECETA_ As Integer
    Dim CANTIDAD_ As Double
    
    If QueHace = 3 Then Exit Sub
            
    Select Case Index
        Case 0 ' AGREGAR
            ReDim xCampos(3, 4) As String
            Dim nSQLId As String
            Dim nSQLId2 As String
            
            If NulosN(cbEstado.ItemData(cbEstado.ListIndex)) > ESTADOPENDIENTE_ Then
                MsgBox "El registro esta en un estado en el que no se puede modificar", vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
                Exit Sub
            End If
            
'            If NulosN(txtIdAlm.Text) = 0 Then
'                MsgBox "No se ha especificado el Almacén", vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
'                txtIdAlm.SetFocus
'                Exit Sub
'            End If
            
            xCampos(0, 0) = "Ítem":        xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":      xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Uni. Med":    xCampos(2, 1) = "abrev":         xCampos(2, 2) = "1200":         xCampos(2, 3) = "C"
                    
            nTitulo = "Buscando Ítems"
                       
            nSQLId = GENERAR_SQL_ID(fg(0), COLUMNAIDITEM, " AND alm_almacenesdet.iditem", "NOT IN", True)
            nSQLId2 = Replace(nSQLId, "alm_almacenesdet.iditem", "alm_inventario.id")
            
'            cSQL = "SELECT alm_almacenesdet.iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
'                + vbCr + "FROM ((alm_almacenes INNER JOIN alm_almacenesdet ON alm_almacenes.id = alm_almacenesdet.idalm) INNER JOIN alm_inventario ON alm_almacenesdet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
'                + vbCr + "WHERE (((alm_almacenes.id) = " & NulosN(txtIdAlm.Text) & ") And ((alm_almacenes.idtippro) = 0)) " & nSQLId _
'                + vbCr + "UNION " _
'                + vbCr + "SELECT alm_inventario.id AS iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
'                + vbCr + "FROM (alm_almacenes INNER JOIN alm_inventario ON alm_almacenes.idtippro = alm_inventario.tippro) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
'                + vbCr + "WHERE (((alm_almacenes.id) = " & NulosN(txtIdAlm.Text) & "))" & nSQLId2
                
            cSQL = "SELECT alm_inventario.id AS iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
                + vbCr + "FROM alm_inventario INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                + vbCr + "WHERE (((alm_inventario.activo)=True)) " & nSQLId2
            
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            fg(0).Rows = fg(0).Rows + 1
            
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNAITEM) = UCase(NulosC(xRs("descripcion")))
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNAIDITEM) = NulosN(xRs("iditem"))
            fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("STOCK")) = Format(SaldoActual(NulosN(xRs("iditem")), "01/01/" & AnoTra, Date, xCon), FORMAT_CANTIDAD)
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACODIGO_) = Busca_Codigo(NulosN(xRs("iditem")), "id", "codpro", "alm_inventario", "N", xCon)
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNAIDUNIMED) = Busca_Codigo(NulosN(xRs("iditem")), "id", "idunimed", "alm_inventario", "N", xCon)
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNAUNIMED) = Busca_Codigo(NulosN(fg(0).TextMatrix(fg(0).Rows - 1, COLUMNAIDUNIMED)), "id", "abrev", "mae_unidades", "N", xCon)
            fg(0).Col = COLUMNACANTIDAD
                        
            fg(0).Select fg(0).Rows - 1, 1
            fg(0).SetFocus
            
            
        Case 1 ' ALMACEN
            If QueHace = 3 Then Exit Sub
            
            'Dim xform As New eps_librerias.FormBuscar
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
            
            nTitulo = "Buscando Almacenes"
            cSQL = "SELECT alm_almacenes.* FROM alm_almacenes"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            txtIdAlm.Text = NulosN(xRs("id"))
            lblAlmacen.Caption = NulosC(xRs("descripcion"))
            TxtIdTipDocRef.SetFocus
            Set xRs = Nothing
            
        Case 2 ' RESPONSABLE
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "apenom":     xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":               xCampos(1, 1) = "id":         xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
            
            cSQL = "SELECT pla_empleados.nombre AS apenom, pla_empleados.id " _
                + vbCr + "FROM pla_empleados;"
            
            nTitulo = "Buscando Responsable"
                   
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "apenom", "apenom", Principio
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            TxtIdResp.Text = NulosN(xRs("id"))
            lblResponsable.Caption = NulosC(xRs("apenom"))
            TxtIdTipDocRef.SetFocus
            
        Case 3 ' SELECCIONAR
            'Dim xform As New eps_librerias.FormSeleccion
            ReDim xCampos(3, 4) As String
            
            If NulosN(cbEstado.ItemData(cbEstado.ListIndex)) > ESTADOPENDIENTE_ Then
                MsgBox "El registro esta en un estado en el que no se puede modificar", vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
                Exit Sub
            End If
            
'            If NulosN(txtIdAlm.Text) = 0 Then
'                MsgBox "No se ha especificado el Almacén", vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
'                txtIdAlm.SetFocus
'                Exit Sub
'            End If
            
            xCampos(0, 0) = "Ítem":        xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":      xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Uni. Med":    xCampos(2, 1) = "abrev":         xCampos(2, 2) = "1200":         xCampos(2, 3) = "C"
                    
            nTitulo = "Seleccionando Ítems"
                       
            nSQLId = GENERAR_SQL_ID(fg(0), COLUMNAIDITEM, " AND alm_almacenesdet.iditem", "NOT IN", True)
            nSQLId2 = Replace(nSQLId, "alm_almacenesdet.iditem", "alm_inventario.id")
                            
'            cSQL = "SELECT 0 AS xsel, alm_almacenesdet.iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
'                + vbCr + "FROM ((alm_almacenes INNER JOIN alm_almacenesdet ON alm_almacenes.id = alm_almacenesdet.idalm) INNER JOIN alm_inventario ON alm_almacenesdet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
'                + vbCr + "WHERE (((alm_almacenes.id) = " & NulosN(txtIdAlm.Text) & ") And ((alm_almacenes.idtippro) = 0)) " & nSQLId _
'                + vbCr + "UNION " _
'                + vbCr + "SELECT 0 AS xsel, alm_inventario.id AS iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
'                + vbCr + "FROM (alm_almacenes INNER JOIN alm_inventario ON alm_almacenes.idtippro = alm_inventario.tippro) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
'                + vbCr + "WHERE (((alm_almacenes.id) = " & NulosN(txtIdAlm.Text) & "))" & nSQLId2
                   
            cSQL = "SELECT 0 AS xsel, alm_inventario.id AS iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
                + vbCr + "FROM alm_inventario INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                + vbCr + "WHERE (((alm_inventario.activo)=True)) " & nSQLId2
                        
            xform.SQLCad = cSQL
            xform.titulo = nTitulo
            Set xform.Coneccion = xCon
            Set xRs = Nothing
            Set xRs = xform.seleccionar(xCampos)
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            xRs.MoveFirst
            While Not xRs.EOF
                fg(0).Rows = fg(0).Rows + 1
                
                fg(0).TextMatrix(fg(0).Rows - 1, COLUMNAITEM) = UCase(NulosC(xRs("descripcion")))
                fg(0).TextMatrix(fg(0).Rows - 1, COLUMNAIDITEM) = NulosN(xRs("iditem"))
                fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACODIGO_) = Busca_Codigo(NulosN(xRs("iditem")), "id", "codpro", "alm_inventario", "N", xCon)
                fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("STOCK")) = Format(SaldoActual(NulosN(xRs("iditem")), "01/01/" & AnoTra, Date, xCon), FORMAT_CANTIDAD)
                fg(0).TextMatrix(fg(0).Rows - 1, COLUMNAIDUNIMED) = Busca_Codigo(NulosN(xRs("iditem")), "id", "idunimed", "alm_inventario", "N", xCon)
                fg(0).TextMatrix(fg(0).Rows - 1, COLUMNAUNIMED) = Busca_Codigo(NulosN(fg(0).TextMatrix(fg(0).Rows - 1, COLUMNAIDUNIMED)), "id", "abrev", "mae_unidades", "N", xCon)
                fg(0).Col = COLUMNACANTIDAD
                
                xRs.MoveNext
            Wend
                        
            fg(0).Select fg(0).Rows - 1, 1
            fg(0).SetFocus
                    
        Case 4 ' TIPO DE DOCUMENTO DE REFERENCIA
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
                  
            nTitulo = "Buscando Tipo de Documento de Referencia"
            
            cSQL = "SELECT mae_documento.* FROM mae_documento WHERE (id In (110,115))"
            
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            TxtIdTipDocRef.Text = NulosN(xRs("id"))
            LblTipDocRef.Caption = NulosC(xRs("descripcion"))
            txtNumDocRef.SetFocus
        
        Case 5 ' DOCUMENTO DE REFERENCIA
            ReDim xCampos(6, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Fch.Doc.":     xCampos(0, 1) = "fchpro":          xCampos(0, 2) = "900":          xCampos(0, 3) = "C"
            xCampos(1, 0) = "Num.Doc":      xCampos(1, 1) = "numdoc":          xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Ítem":         xCampos(2, 1) = "desitem":         xCampos(2, 2) = "3200":         xCampos(2, 3) = "C"
            xCampos(3, 0) = "Receta":       xCampos(3, 1) = "codrec":          xCampos(3, 2) = "1000":         xCampos(3, 3) = "C"
            xCampos(4, 0) = "Cantidad":     xCampos(4, 1) = "cantidad":        xCampos(4, 2) = "900":         xCampos(4, 3) = "N"
            xCampos(5, 0) = "Hor.Ini.":     xCampos(5, 1) = "horini":          xCampos(5, 2) = "800":         xCampos(5, 3) = "C"
                  
            nTitulo = "Buscando Tipo de Documento de Referencia"
            cSQL = "SELECT pro_ordenprod.id, Format([pro_ordenprod].[fchpro],'dd/mm/yy') AS fchpro, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numdoc, alm_inventario.descripcion AS desitem, pro_receta.codrec, pro_ordenprod.cantidad, Format([pro_ordenprod].[horini],'Short Time') AS horini, Format([pro_ordenprod].[horfin],'Short Time') AS horfin, pro_ordenprod.estado, UCase([mae_estados].[descripcion]) AS desestado " _
                + vbCr + "FROM ((pro_ordenprod LEFT JOIN pro_receta ON pro_ordenprod.idrec = pro_receta.id) LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) LEFT JOIN mae_estados ON pro_ordenprod.estado = mae_estados.id " _
                + vbCr + "WHERE (((pro_ordenprod.ano) = " & AnoTra & ") And ((pro_ordenprod.idmes) in (" & mMesActivo & ", " & mMesActivo - 1 & ")));"
            
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "fchpro", "fchpro", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            lbliddocref.Caption = NulosN(xRs("id"))
            txtNumDocRef.Text = NulosC(xRs("numdoc"))
            
            fg(0).SetFocus
            ' -------------------------------------SE BUSCAN LOS INSUMOS
            IDRECETA_ = Busca_Codigo(NulosN(xRs("id")), "id", "idrec", "pro_ordenprod", "N", xCon)
            CANTIDAD_ = Busca_Codigo(NulosN(xRs("id")), "id", "cantidad", "pro_ordenprod", "N", xCon)
            
            nSQLId = GENERAR_SQL_ID(fg(0), COLUMNAIDITEM, " AND alm_inventario.id", "NOT IN", True)
                        
            cSQL = "SELECT alm_inventario.tippro AS idtippro, pro_recetains.iditem, [pro_recetains]![canpro]*" & CANTIDAD_ & " AS cantidad, pro_recetains.idunimed " _
                + vbCr + "FROM pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id " _
                + vbCr + "WHERE (((pro_recetains.idrec)=" & IDRECETA_ & ")) " & nSQLId
            Set xRs = Nothing
            RST_Busq xRs, cSQL, xCon
            
            fg(0).Rows = fg(0).FixedRows
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            ' SE LLENA INSUMOS
            xRs.MoveFirst
            While Not xRs.EOF
                fg(0).Rows = fg(0).Rows + 1
                With fg(0)
                    .TextMatrix(.Rows - 1, COLUMNACODIGO_) = UCase(Busca_Codigo(NulosN(xRs("iditem")), "id", "codpro", "alm_inventario", "N", xCon))
                    .TextMatrix(.Rows - 1, COLUMNAITEM) = UCase(Busca_Codigo(NulosN(xRs("iditem")), "id", "descripcion", "alm_inventario", "N", xCon))
                    .TextMatrix(.Rows - 1, COLUMNAUNIMED) = UCase(Busca_Codigo(NulosN(xRs("idunimed")), "id", "abrev", "mae_unidades", "N", xCon))
                    .TextMatrix(.Rows - 1, .ColIndex("STOCK")) = Format(SaldoActual(NulosN(xRs("iditem")), "01/01/" & AnoTra, Date, xCon), FORMAT_CANTIDAD)
                    .TextMatrix(.Rows - 1, COLUMNALOTE) = ""
                    .TextMatrix(.Rows - 1, COLUMNAIDITEM) = NulosN(xRs("iditem"))
                    .TextMatrix(.Rows - 1, COLUMNAIDLOTE) = ""
                    .TextMatrix(.Rows - 1, COLUMNAIDLOTEDET) = ""
                    .TextMatrix(.Rows - 1, COLUMNAIDUNIMED) = NulosN(xRs("idunimed"))
                End With
                xRs.MoveNext
            Wend
            
        Case 6 '
        
        Case 7 '
            
        Case 8 ' ELIMINAR
            If fg(0).Rows <= 0 Then Exit Sub
            
            If NulosN(cbEstado.ItemData(cbEstado.ListIndex)) > ESTADOPENDIENTE_ Then
                MsgBox "El registro esta en un estado en el que no se puede modificar", vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
                Exit Sub
            End If
            
            Rpta = MsgBox("¿Esta seguro de Eliminar esta Solicitud?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            
            If Rpta = vbYes Then fg(0).RemoveItem fg(0).Row
            
        Case 9 ' ELIMINAR TODOS
            If fg(0).Rows <= 0 Then Exit Sub
            
            If NulosN(cbEstado.ItemData(cbEstado.ListIndex)) > ESTADOPENDIENTE_ Then
                MsgBox "El registro esta en un estado en el que no se puede modificar", vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
                Exit Sub
            End If
            
            Rpta = MsgBox("¿Está seguro de eliminar todos los registros?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
        
            If Rpta = vbYes Then fg(0).Rows = fg(0).FixedRows
        
    End Select
End Sub

Sub CrearCabeceraVS(numPag As Integer)
    Dim xCad As String

    FrmVsPrinter.Vs.TextAlign = taLeftTop
    FrmVsPrinter.Vs.FontName = "Courier New"
    FrmVsPrinter.Vs.FontBold = True
    FrmVsPrinter.Vs.FontSize = 9

    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 200
    FrmVsPrinter.Vs.Paragraph = "EMPRESA   : " & NomEmp

    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 200
    FrmVsPrinter.Vs.Paragraph = "FECHA        : " & Format(Date, FORMAT_DATE)

    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 400
    FrmVsPrinter.Vs.Paragraph = "Nº R.U.C. : " & NumRUC

    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 400
    FrmVsPrinter.Vs.Paragraph = "Nº PAG.      : " & Format(numPag, "0000")

    FrmVsPrinter.Vs.DrawLine 1000, 650, 11000, 650
End Sub

Private Function ImprimirSolicitud(RSTCAB_ As ADODB.Recordset, RSTDET_ As ADODB.Recordset) As Boolean
    Dim A As Integer
    Dim numPag As Integer
    Dim Rst As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim B As Integer
    Dim FILA_ As Integer
    Dim COLUMNA_ As Integer
    Dim numper As Double
    Dim xFila As Integer
    Dim Nombre As String
    
    If RSTCAB_.State = 0 Then ImprimirSolicitud = False: Exit Function
    If RSTDET_.State = 0 Then ImprimirSolicitud = False: Exit Function
    RSTCAB_.Filter = adFilterNone
    RSTDET_.Filter = adFilterNone
    If RSTCAB_.RecordCount = 0 Then ImprimirSolicitud = False: Exit Function
    If RSTDET_.RecordCount = 0 Then ImprimirSolicitud = False: Exit Function
        
    FrmVsPrinter.Vs.ExportFormat = vpxRTF
    FrmVsPrinter.Vs.ExportFile = "c:\report2.xls"
    With FrmVsPrinter.Vs
        numPag = 0
        .BrushColor = &H80000005
        .FontSize = 11
        .TextAlign = taCenterMiddle
        .StartDoc
            
        '*************************************************************
        ' -----------------------------------------------PIE DE PAGINA
        '*************************************************************
        FILA_ = 800
        COLUMNA_ = 1000
        numPag = numPag + 1
        CrearCabeceraVS numPag
        
        RSTCAB_.MoveFirst
        While Not RSTCAB_.EOF
            '******************************************************
            ' -----------------------------------------------TITULO
            '******************************************************
            If FILA_ >= 13000 Then
                .NewPage
                FILA_ = 800
                numPag = numPag + 1
                CrearCabeceraVS numPag
            End If
            .FontSize = 12
            .FontBold = True
            .TextAlign = taCenterMiddle
            .TextBox "SOLICITUD", COLUMNA_, FILA_, 8000, 500, True, False, True
            .FontSize = 10
            .TextAlign = taCenterTop
            .TextBox "Nº ", COLUMNA_ + 8100, FILA_, 1900, 250, True, False, True
            FILA_ = FILA_ + 240
            .TextBox NulosC(RSTCAB_("numser")) & "-" & NulosC(RSTCAB_("numdoc")), COLUMNA_ + 8100, FILA_, 1900, 250, True, False, True
            
            '********************************************************
            ' -----------------------------------------------CABECERA
            '********************************************************
            ' ------------------------DETALLE DE LA ORDEN COMO ANEXO
            If NulosN(RSTCAB_("idtipdocref")) = 115 Then
                
                cSQL = "SELECT alm_inventario.descripcion AS anexo, pro_ordenprod.cantidad, mae_unidades.abrev " _
                    + vbCr + "FROM (alm_inventario RIGHT JOIN (pro_ordenprod LEFT JOIN pro_receta ON pro_ordenprod.idrec = pro_receta.id) ON alm_inventario.id = pro_receta.iditem) LEFT JOIN mae_unidades ON pro_ordenprod.idunimed = mae_unidades.id " _
                    + vbCr + "WHERE (((pro_ordenprod.id)=" & NulosN(RSTCAB_("iddocref")) & "));"
                
'                cSQL = "SELECT alm_inventario.descripcion AS anexo " _
'                    + vbCr + "FROM alm_inventario RIGHT JOIN (pro_ordenprod LEFT JOIN pro_receta ON pro_ordenprod.idrec = pro_receta.id) ON alm_inventario.id = pro_receta.iditem " _
'                    + vbCr + "WHERE (((pro_ordenprod.id)=" & NulosN(RSTCAB_("iddocref")) & "));"
                
                Set xRs = Nothing
                RST_Busq xRs, cSQL, xCon
                
                If xRs.State = 0 Then GoTo SALIRANEXO_
                If xRs.RecordCount = 0 Then GoTo SALIRANEXO_
                
                FILA_ = FILA_ + 300
                .FontBold = True
                .TextAlign = taLeftMiddle
                .TextBox "Producto   :", COLUMNA_, FILA_, 1500, 250, True, False, False
                .FontBold = False
                .TextBox UCase(NulosC(xRs("anexo"))), COLUMNA_ + 1500, FILA_, 10000, 250, True, False, False
                
                FILA_ = FILA_ + 300
                .FontBold = True
                .TextAlign = taLeftMiddle
                .TextBox "Cantidad   :", COLUMNA_, FILA_, 1500, 250, True, False, False
                .FontBold = False
                .TextBox NulosC(xRs("cantidad")) & " " & NulosC(xRs("abrev")), COLUMNA_ + 1500, FILA_, 10000, 250, True, False, False
SALIRANEXO_:
            End If
            
            .TextAlign = taLeftMiddle
            .FontSize = 9
            FILA_ = FILA_ + 300
            .FontBold = True
            .TextBox "T. Doc. Ref.:", COLUMNA_, FILA_, 1500, 250, True, False, False
            .FontBold = False
            .TextBox UCase(NulosC(RSTCAB_("destipdocref"))), COLUMNA_ + 1500, FILA_, 7000, 250, True, False, False
            .FontBold = True
            .TextBox "Nº Doc. Ref.:", COLUMNA_ + 6000, FILA_, 1500, 250, True, False, False
            .FontBold = False
            .TextBox NulosC(RSTCAB_("numdocref")), COLUMNA_ + 7500, FILA_, 6000, 250, True, False, False
            FILA_ = FILA_ + 250
            .FontBold = True
            .TextBox "Lote     :", COLUMNA_, FILA_, 1500, 250, True, False, False
            .FontBold = False
            .TextBox Busca_Codigo(NulosC(RSTCAB_("iddocref")), "id", "lote", "pro_ordenprod", "N", xCon), COLUMNA_ + 1500, FILA_, 6000, 250, True, False, False
            .FontBold = True
            .TextBox "Fecha Doc.  :", COLUMNA_ + 6000, FILA_, 1500, 250, True, False, False
            .FontBold = False
            .TextBox Format(NulosC(RSTCAB_("fchdoc")), FORMAT_DATE), COLUMNA_ + 7500, FILA_, 6000, 250, True, False, False
                        
            '*******************************************************
            '------------------------------------------------DETALLE
            '*******************************************************
            RSTDET_.Filter = "idsol = " & NulosN(RSTCAB_("id"))
            If RSTDET_.RecordCount = 0 Then ImprimirSolicitud = False: Exit Function
            
            FILA_ = FILA_ + 350
            .TextAlign = taCenterMiddle
            .TextBox "Código", COLUMNA_, FILA_, 1750, 500, True, False, True
            .TextBox "Ítem", COLUMNA_ + 1750, FILA_, 5900, 500, True, False, True
            .TextBox "U.M.", COLUMNA_ + 7650, FILA_, 800, 500, True, False, True
            .TextBox "Cantidad", COLUMNA_ + 8450, FILA_, 1550, 500, True, False, True
    
            FILA_ = FILA_ + 250
                
            xFila = FILA_
            While Not RSTDET_.EOF
                FILA_ = FILA_ + 250
                If FILA_ >= 16200 Then
                    FILA_ = 800
                    numPag = numPag + 1
                    .NewPage
                    CrearCabeceraVS numPag
                End If
                
                .FontSize = 8
                .FontBold = False
                .TextAlign = taLeftMiddle
                .TextBox " " & RSTDET_("coditem"), COLUMNA_, FILA_, 1750, 250, True, False, True
                .FontSize = 7
                .TextBox " " & RSTDET_("desitem"), COLUMNA_ + 1750, FILA_, 5900, 250, True, False, True
                .FontSize = 8
                .TextAlign = taCenterMiddle
                .TextBox NulosC(RSTDET_("desunimed")), COLUMNA_ + 7650, FILA_, 800, 250, True, False, True
                .TextAlign = taRightMiddle
                .TextBox Format(RSTDET_("cantidad"), FORMAT_CANTIDADDECIMAL), COLUMNA_ + 8450, FILA_, 1550, 250, True, False, True
                
                RSTDET_.MoveNext
            Wend
    
            FILA_ = FILA_ + 400
            If FILA_ >= 16000 Then
                FILA_ = 2000
                .NewPage
            End If
    
            .TextBox "_______________________________", COLUMNA_ + 700, FILA_, 3500, 250, True, False, False
            .TextBox "_______________________________", COLUMNA_ + 5700, FILA_, 3500, 250, True, False, False
    
            FILA_ = FILA_ + 200
    
            .FontSize = 7
            .TextAlign = taCenterMiddle
    
            .TextBox "PRODUCCION", COLUMNA_ + 1000, FILA_, 3500, 250, True, False, False
    
            .TextBox "ALMACEN", COLUMNA_ + 6000, FILA_, 3500, 250, True, False, False
            .FontSize = 8
    
            FILA_ = FILA_ + 400
            
            RSTCAB_.MoveNext
        Wend
        .EndDoc
    End With
    'Muestra la preimagen de la impresion
    FrmVsPrinter.WindowState = 2
    FrmVsPrinter.Show
End Function

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstSol("id")), xCon
    End If
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstSol
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDETE LAS COLUMNAS DEL CONTROL Dg3
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstSol.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub fg_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Index
        Case 0:
            If NulosN(cbEstado.ItemData(cbEstado.ListIndex)) > ESTADOPENDIENTE_ Then Cancel = True
            
    End Select
End Sub

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Index = 0 Then
        Dim xRs As New ADODB.Recordset
        Dim nTitulo As String
        Dim xCampos() As String
        Dim TIPOPRODUCTO_ As Integer
        Dim IDITEM_ As Integer
        Dim nSQLId As String
        Dim nSQLId2 As String
        
        If QueHace = 3 Then Exit Sub
        
        Select Case Col
            Case COLUMNACODIGO_ ' CODIGO DE PRODUCTO
                ReDim xCampos(3, 4) As String
                
                xCampos(0, 0) = "Ítem":        xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
                xCampos(1, 0) = "Codigo":      xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
                xCampos(2, 0) = "Uni. Med":    xCampos(2, 1) = "abrev":         xCampos(2, 2) = "1200":         xCampos(2, 3) = "C"
                        
                nTitulo = "Buscando " & NulosC(fg(0).TextMatrix(Row, COLUMNACODIGO_))
                       
                nSQLId = GENERAR_SQL_ID(fg(0), COLUMNAIDITEM, " AND alm_almacenesdet.iditem", "NOT IN", True)
                nSQLId2 = Replace(nSQLId, "alm_almacenesdet.iditem", "alm_inventario.id")
                                
                cSQL = "SELECT alm_almacenesdet.iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
                    + vbCr + "FROM ((alm_almacenes INNER JOIN alm_almacenesdet ON alm_almacenes.id = alm_almacenesdet.idalm) INNER JOIN alm_inventario ON alm_almacenesdet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                    + vbCr + "WHERE (((alm_almacenes.id) = " & NulosN(txtIdAlm.Text) & ") And ((alm_almacenes.idtippro) = 0)) " & nSQLId _
                    + vbCr + "UNION " _
                    + vbCr + "SELECT alm_inventario.id AS iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
                    + vbCr + "FROM (alm_almacenes INNER JOIN alm_inventario ON alm_almacenes.idtippro = alm_inventario.tippro) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                    + vbCr + "WHERE (((alm_almacenes.id) = " & NulosN(txtIdAlm.Text) & "))" & nSQLId2
                
                Set xRs = Nothing
                CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                "descripcion", "descripcion", Principio, ""
                
                If xRs.State = 0 Then Exit Sub
                If xRs.RecordCount = 0 Then Exit Sub
                
                fg(0).TextMatrix(Row, COLUMNAITEM) = UCase(NulosC(xRs("descripcion")))
                fg(0).TextMatrix(Row, COLUMNAIDITEM) = NulosN(xRs("iditem"))
                
                fg(0).TextMatrix(Row, COLUMNACODIGO_) = Busca_Codigo(NulosN(xRs("iditem")), "id", "codpro", "alm_inventario", "N", xCon)
                fg(0).TextMatrix(Row, COLUMNAIDUNIMED) = Busca_Codigo(NulosN(xRs("iditem")), "id", "idunimed", "alm_inventario", "N", xCon)
                fg(0).TextMatrix(Row, COLUMNAUNIMED) = Busca_Codigo(NulosN(fg(0).TextMatrix(Row, COLUMNAIDUNIMED)), "id", "abrev", "mae_unidades", "N", xCon)
                fg(0).Col = COLUMNACANTIDAD
                    
                
            Case COLUMNAUNIMED
                ReDim xCampos(2, 4) As String
                
                ' Se verifica si se escogio el producto
                If NulosN(fg(0).TextMatrix(Row, COLUMNAIDITEM)) = 0 Then
                    MsgBox "Seleccione el Ítem para el registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    fg(0).Col = COLUMNAITEM
                    Exit Sub
                End If
                
                'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
                xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "2500":         xCampos(0, 3) = "C"
                xCampos(1, 0) = "Abrev.":           xCampos(1, 1) = "abrev":            xCampos(1, 2) = "1000":         xCampos(1, 3) = "D"
                        
                nTitulo = "Buscando Unidades"

                cSQL = "SELECT * " _
                    + vbCr + "FROM mae_unidades;"
                
                Set xRs = Nothing
                CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                "descripcion", "descripcion", Principio, ""
                
                If xRs.State = 0 Then Exit Sub
                If xRs.RecordCount = 0 Then Exit Sub
                
                fg(0).TextMatrix(Row, COLUMNAIDUNIMED) = NulosN(xRs("id"))
                fg(0).TextMatrix(Row, COLUMNAUNIMED) = NulosC(xRs("abrev"))
                                        
            Case COLUMNALOTE ' LOTE
                ReDim xCampos(4, 4) As String
                
                ' Se verifica si se escogio el producto
                If NulosN(fg(0).TextMatrix(Row, COLUMNAIDITEM)) = 0 Then
                    MsgBox "Seleccione el Ítem para el registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    fg(0).Col = COLUMNAITEM
                    Exit Sub
                End If
                
                'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
                xCampos(0, 0) = "Lote":         xCampos(0, 1) = "deslote":      xCampos(0, 2) = "2000":         xCampos(0, 3) = "C"
                xCampos(1, 0) = "Fch. Ing.":    xCampos(1, 1) = "fching":       xCampos(1, 2) = "1000":         xCampos(1, 3) = "D"
                xCampos(2, 0) = "Almacen":      xCampos(2, 1) = "desalm":       xCampos(2, 2) = "3000":         xCampos(2, 3) = "C"
                xCampos(3, 0) = "Cantidad":     xCampos(3, 1) = "cantidad":     xCampos(3, 2) = "1000":         xCampos(3, 3) = "N"
                        
                nTitulo = "Buscando Lotes de " & NulosC(fg(0).TextMatrix(Row, COLUMNAITEM))

                cSQL = "SELECT alm_inventariolotedet.idlote, alm_inventariolotedet.id AS idlotedet, alm_inventariolote.iditem, alm_inventariolotedet.idalm, alm_inventariolote.fching, alm_almacenes.descripcion AS desalm, alm_inventariolotedet.cantidad, alm_inventariolote.descripcion AS deslote " _
                    + vbCr + "FROM (alm_inventariolote LEFT JOIN alm_inventariolotedet ON alm_inventariolote.id = alm_inventariolotedet.idlote) LEFT JOIN alm_almacenes ON alm_inventariolotedet.idalm = alm_almacenes.id " _
                    + vbCr + "WHERE (((alm_inventariolote.iditem)=" & NulosN(fg(0).TextMatrix(Row, COLUMNAIDITEM)) & ") AND ((alm_inventariolotedet.idalm)=" & NulosN(txtIdAlm.Text) & "))"
                
                Set xRs = Nothing
                CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                "deslote", "deslote", Principio, ""
                
                If xRs.State = 0 Then Exit Sub
                If xRs.RecordCount = 0 Then Exit Sub
                
                If xRs("cantidad") < NulosN(fg(0).TextMatrix(Row, COLUMNACANTIDAD)) Then
                    MsgBox "El lote seleccionado no contiene stock suficiente", vbCritical + vbOKOnly + vbDefaultButton1, xTitulo
                    ' LOTE
                    fg(0).TextMatrix(Row, COLUMNAIDLOTE) = 0
                    fg(0).TextMatrix(Row, COLUMNAIDLOTEDET) = 0
                    fg(0).TextMatrix(Row, COLUMNALOTE) = ""
                    Exit Sub
                End If
                
                ' LOTE
                fg(0).TextMatrix(Row, COLUMNAIDLOTE) = NulosN(xRs("idlote"))
                fg(0).TextMatrix(Row, COLUMNAIDLOTEDET) = NulosN(xRs("idlotedet"))
                fg(0).TextMatrix(Row, COLUMNALOTE) = NulosC(xRs("deslote"))
                
        End Select
    End If
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Index = 0 Then
        If Agregando = True Then Exit Sub

        If Col = COLUMNACANTIDAD Then 'Cambiar cantidad
            fg(0).TextMatrix(Row, Col) = Format(fg(0).TextMatrix(Row, Col), FORMAT_CANTIDADDECIMAL)
        End If
    End If
End Sub

Private Function cambiarEstadoRelacionados(IDORDDET_ As Double, ESTADO_ As Double) As Boolean
    Dim ID_ As Double
    
    On Error GoTo ERROR_
    ' Salidas de Almacen
    cSQL = "UPDATE alm_ingreso SET alm_ingreso.estado = " & ESTADO_ & " " _
        + vbCr + "WHERE (((alm_ingreso.idorddet)=" & IDORDDET_ & "));"

    xCon.Execute cSQL
    
    ' GRABAMOS LOS MOVIMIENTOS
    ' INGRESOS Y SALIDAS DE ALMACEN
    ID_ = Busca_Codigo(IDORDDET_, "idorddet", "id", "alm_ingreso", "N", xCon)
    GrabarOperacion xIdUsuario, 8, 7, xHorIni, Time, Date, xCon, ID_
        
    cambiarEstadoRelacionados = True
    Exit Function
    
ERROR_:
    MsgBox "Ha ocurrido un error al tratar de cambiar de estado", vbInformation, xTitulo
    cambiarEstadoRelacionados = False
End Function

Private Sub Fg_EnterCell(Index As Integer)
    If QueHace = 3 Then
        fg(Index).Editable = flexEDNone
        fg(Index).SelectionMode = flexSelectionByRow
        Exit Sub
    End If
    fg(Index).Editable = flexEDKbdMouse
    fg(Index).SelectionMode = flexSelectionFree
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Index = 0 Then
        Select Case Col
            Case COLUMNACODIGO_, COLUMNALOTE, COLUMNAITEM, COLUMNAUNIMED
                KeyAscii = 0
            
            Case COLUMNACANTIDAD
                If IsNumeric(KeyAscii) = False Then KeyAscii = 0
                
        End Select
    End If
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        If QueHace = 3 Then Exit Sub
        If Button = 2 Then
            PopupMenu Menu1
        End If
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    TabOne1.CurrTab = 0
    SeEjecuto = False
    Agregando = False
    iniciarCampos
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDOS E CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        mMesActivo = xMes
        
        pCargarGrid
    End If
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub

    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 4000 Then Me.Height = 4000

    ' Se dimensiona el Contenido
    TabOne1.Width = Me.Width - 150
    TabOne1.Height = Me.Height - 750
    
    Label4(0).Width = Me.Width - 100
    lblperiodo.Left = TabOne1.Width - 1200
    Dg1.Width = TabOne1.Width - 135
    Dg1.Height = TabOne1.Height - 945
    
    ' Se dimensiona el Detalle
    Label5.Width = Me.Width - 100
    
    fg(0).Width = TabOne1.Width - 1665
    fg(0).Height = TabOne1.Height - 2730
        
    Frame3.Left = TabOne1.Width - 1545
    Frame3.Height = TabOne1.Height - 2640
    
End Sub

Private Sub iniciarCampos()
    fg(0).AllowUserResizing = flexResizeColumns
    fg(0).ExplorerBar = flexExSortShowAndMove
    fg(0).SelectionMode = flexSelectionByRow
    fg(0).BackColorSel = &H80&
    fg(0).ForeColorSel = &H80000005
    
    fg(0).Rows = 1
    fg(0).ColWidth(COLUMNAIDITEM) = 0
    fg(0).ColWidth(COLUMNAIDLOTE) = 0
    fg(0).ColWidth(COLUMNAIDLOTEDET) = 0
    fg(0).ColWidth(COLUMNAIDUNIMED) = 0
    fg(0).ColWidth(COLUMNALOTE) = 0
    
    GRID_COMBOLIST fg(0), COLUMNACODIGO_
    GRID_COMBOLIST fg(0), COLUMNAUNIMED
    GRID_COMBOLIST fg(0), COLUMNALOTE
    
    Dg1.Columns("numdoc").Alignment = dbgCenter
    Dg1.Columns("destipdocref").Alignment = dbgCenter
    
    ' Se agrega el mes Activo
    mMesActivo = xMes
    lblperiodo.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
        
    CAMBIOGRABAR_ = 0
    ESTADOANTERIOR_ = ESTADOPENDIENTE_
End Sub

Private Sub llenarEstados()
    Dim xRs As New ADODB.Recordset
    
    cSQL = "SELECT * FROM mae_estados"
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    cbEstado.Clear
    xRs.MoveFirst
    While Not xRs.EOF
        cbEstado.AddItem UCase(NulosC(xRs("descripcion")))
        cbEstado.ItemData(cbEstado.NewIndex) = NulosN(xRs("id"))
        xRs.MoveNext
    Wend
    
    cbEstado.ListIndex = 0
End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub pCargarGrid()
    Dim cSQL  As String
    Dim Rpta As Integer
    
    TDB_FiltroLimpiar Dg1
    
    cSQL = "SELECT pro_solicitudmat.id, Format([pro_solicitudmat].[fchdoc],'dd/mm/yy') As fchdoc, [pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc] AS numdoc, pla_empleados.nombre AS desresp, alm_almacenes.descripcion AS desalm, mae_documento.abrev AS destipdocref, IIf([pro_solicitudmat].[idtipdocref]=115,[pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc],'') AS numdocref, UCase([mae_estados].[descripcion]) AS desestado, pro_solicitudmat.estado, alm_inventario.descripcion AS anexo " _
        + vbCr + "FROM ((((((pro_solicitudmat LEFT JOIN pla_empleados ON pro_solicitudmat.idresp = pla_empleados.id) LEFT JOIN mae_estados ON pro_solicitudmat.estado = mae_estados.id) LEFT JOIN mae_documento ON pro_solicitudmat.idtipdocref = mae_documento.id) LEFT JOIN alm_almacenes ON pro_solicitudmat.idalm = alm_almacenes.id) LEFT JOIN pro_ordenprod ON pro_solicitudmat.iddocref = pro_ordenprod.id) LEFT JOIN pro_receta ON pro_ordenprod.idrec = pro_receta.id) LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((pro_solicitudmat.ano) = " & AnoTra & ") And ((pro_solicitudmat.idmes) = " & mMesActivo & ")) " _
        + vbCr + "ORDER BY Format([pro_solicitudmat].[fchdoc],'dd/mm/yy') DESC , [pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc] DESC;"

'    cSQL = "SELECT pro_solicitudmat.id, pro_solicitudmat.fchdoc, [pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc] AS numdoc, pla_empleados.nombre AS desresp, alm_almacenes.descripcion AS desalm, mae_documento.abrev AS destipdocref, IIf([pro_solicitudmat].[idtipdocref]=115,[pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc],'') AS numdocref, UCase([mae_estados].[descripcion]) AS desestado, pro_solicitudmat.estado " _
'        + vbCr + "FROM ((((pro_solicitudmat LEFT JOIN pla_empleados ON pro_solicitudmat.idresp = pla_empleados.id) LEFT JOIN mae_estados ON pro_solicitudmat.estado = mae_estados.id) LEFT JOIN mae_documento ON pro_solicitudmat.idtipdocref = mae_documento.id) LEFT JOIN alm_almacenes ON pro_solicitudmat.idalm = alm_almacenes.id) LEFT JOIN pro_ordenprod ON pro_solicitudmat.iddocref = pro_ordenprod.id " _
'        + vbCr + "WHERE (((pro_solicitudmat.ano)=" & AnoTra & ") AND ((pro_solicitudmat.idmes)=" & mMesActivo & ")) " _
'        + vbCr + "ORDER BY pro_solicitudmat.fchdoc DESC, [pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc] DESC;"
    
    ' cargando datos
    Me.MousePointer = vbHourglass
    
    RST_Busq RstSol, cSQL, xCon
    Set Dg1.DataSource = RstSol
    
    Me.MousePointer = vbDefault
       
            
    '********************************************************************************************
    lblperiodo.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    '********************************************************************************************

    '------------------------------------------------------------------------------------------
    ' bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    
    If RstSol.State = 0 Then Exit Sub
End Sub

Private Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim Rpta As Integer
    
    Agregando = True
    Blanquea
    'llenarEstados
    If QueHace = 3 Then llenarEstado 1, 1, , cbEstado, , , True
    
    If RstSol.RecordCount = 0 Then Exit Sub
    If RstSol.EOF = True Then Exit Sub
     
    Set xRs = Nothing
    Agregando = True
    
    ' CABECERA
    cSQL = "SELECT * " _
        + vbCr + "FROM pro_solicitudmat " _
        + vbCr + "WHERE (((pro_solicitudmat.id)=" & NulosN(RstSol("id")) & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    seleccionarIndiceCombo NulosN(xRs("estado")), cbEstado
    
    TxtFchPro.valor = NulosC(xRs("fchdoc"))
    TxtNumSer.Text = NulosC(xRs("numser"))
    TxtNumDoc.Text = NulosC(xRs("numdoc"))
    TxtIdResp.Text = NulosN(xRs("idresp"))
    lblResponsable.Caption = UCase(Busca_Codigo(NulosN(xRs("idresp")), "id", "nombre", "pla_empleados", "N", xCon))
    txtIdAlm.Text = NulosN(xRs("idalm"))
    If NulosN(txtIdAlm.Text) = 0 Then txtIdAlm.Text = ""
    lblAlmacen.Caption = UCase(Busca_Codigo(NulosN(xRs("idalm")), "id", "descripcion", "alm_almacenes", "N", xCon))
    TxtIdTipDocRef.Text = NulosN(xRs("idtipdocref"))
    If NulosN(TxtIdTipDocRef.Text) = 0 Then TxtIdTipDocRef.Text = ""
    LblTipDocRef.Caption = UCase(Busca_Codigo(NulosN(xRs("idtipdocref")), "id", "descripcion", "mae_documento", "N", xCon))
    lbliddocref.Caption = NulosN(xRs("iddocref"))
    If NulosN(xRs("idtipdocref")) = 115 Then ' ORDEN DE PRODUCCION
        txtNumDocRef.Text = Busca_Codigo(NulosN(xRs("iddocref")), "id", "numser", "pro_ordenprod", "N", xCon)
        txtNumDocRef.Text = txtNumDocRef.Text & "-" & Busca_Codigo(NulosN(xRs("iddocref")), "id", "numdoc", "pro_ordenprod", "N", xCon)
    End If
        
    ' DETALLE
    cSQL = "SELECT pro_solicitudmatdet.*, alm_inventario.descripcion AS desitem, mae_unidades.abrev AS desunimed, alm_inventariolote.descripcion AS deslote " _
        + vbCr + "FROM ((pro_solicitudmatdet LEFT JOIN alm_inventario ON pro_solicitudmatdet.iditem = alm_inventario.id)  LEFT JOIN mae_unidades ON pro_solicitudmatdet.idunimed = mae_unidades.id) LEFT JOIN alm_inventariolote ON pro_solicitudmatdet.idlote = alm_inventariolote.id " _
        + vbCr + "WHERE (((pro_solicitudmatdet.idsol)=" & NulosN(RstSol("id")) & "));"

    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon

    fg(0).Rows = fg(0).FixedRows
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub

    xRs.MoveFirst
    While Not xRs.EOF
        fg(0).Rows = fg(0).Rows + 1
        With fg(0)
            .TextMatrix(.Rows - 1, COLUMNACODIGO_) = UCase(Busca_Codigo(NulosN(xRs("iditem")), "id", "codpro", "alm_inventario", "N", xCon))
            .TextMatrix(.Rows - 1, COLUMNAITEM) = UCase(NulosC(xRs("desitem")))
            .TextMatrix(.Rows - 1, COLUMNAUNIMED) = UCase(NulosC(xRs("desunimed")))
            .TextMatrix(.Rows - 1, .ColIndex("STOCK")) = Format(SaldoActual(NulosN(xRs("iditem")), "01/01/" & AnoTra, Date, xCon), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, COLUMNACANTIDAD) = Format(NulosN(xRs("cantidad")), FORMAT_CANTIDADDECIMAL)
            .TextMatrix(.Rows - 1, COLUMNALOTE) = UCase(NulosC(xRs("deslote")))
            .TextMatrix(.Rows - 1, COLUMNAIDITEM) = NulosN(xRs("iditem"))
            .TextMatrix(.Rows - 1, COLUMNAIDLOTE) = NulosN(xRs("idlote"))
            .TextMatrix(.Rows - 1, COLUMNAIDLOTEDET) = NulosN(xRs("idlotedet"))
            .TextMatrix(.Rows - 1, COLUMNAIDUNIMED) = NulosN(xRs("idunimed"))
        End With
        xRs.MoveNext
    Wend

    fg(0).Row = 1
    Set xRs = Nothing
    Agregando = False
End Sub

Sub Cancelar()
    If CAMBIOGRABAR_ = -1 Then
        MsgBox "No se puede Cancelar la operación; Grabe los registros para continuar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Bloquea
    Label5.Caption = "Detalle de Solicitud de Materiales"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    fg(0).SelectionMode = flexSelectionByRow
     
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
    limpiarRST RstValores
End Sub

Sub Nuevo()
    'llenarEstados
    llenarEstado 1, 1, , cbEstado, , , False, ESTADOPENDIENTE_ & "," & ESTADOPROCESADO_
    
    QueHace = 1
    xHorIni = Time
    Bloquea
    Blanquea
    agregados = 0
    fg(0).Rows = 1
    'fg(0).Rows = fg(0).Rows + 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    
    Label5.Caption = "Agregando Solicitud de Materiales"
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
    TxtFchPro.valor = Date
End Sub

Sub Bloquea()
    cbEstado.Locked = Not cbEstado.Locked
    TxtFchPro.Locked = Not TxtFchPro.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtIdResp.Locked = Not TxtIdResp.Locked
    txtIdAlm.Locked = Not txtIdAlm.Locked
    TxtIdTipDocRef.Locked = Not TxtIdTipDocRef.Locked
    txtNumDocRef.Locked = Not txtNumDocRef.Locked
    habilitar Cmd, Not TxtFchPro.Locked
End Sub

Sub Blanquea()
    TxtFchPro.valor = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtIdResp.Text = ""
    lblResponsable.Caption = ""
    txtIdAlm.Text = ""
    lblAlmacen.Caption = ""
    TxtIdTipDocRef.Text = ""
    LblTipDocRef.Caption = ""
    lbliddocref.Caption = ""
    txtNumDocRef.Text = ""
End Sub

Private Function verificarCampos() As Boolean
    If Agregando Then Exit Function
    
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If TxtFchPro.valor = "" Then
        MsgBox "No ha especificado fecha de solicitud", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchPro.SetFocus
        verificarCampos = False
        Exit Function
    End If
    
    If TxtIdResp.Text = "" Then
        MsgBox "No ha especificado un encargado para la solicitud", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdResp.SetFocus
        verificarCampos = False
        Exit Function
    End If
    
    If TxtNumSer.Text = "" Then
        MsgBox "No ha especificado el número de serie", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        verificarCampos = False
        Exit Function
    End If
    
    If TxtNumDoc.Text = "" Then
        MsgBox "No ha especificado el número de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        verificarCampos = False
        Exit Function
    End If
    
'    If txtIdAlm.Text = "" Then
'        MsgBox "No ha especificado el tipo de solicitud", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        txtIdAlm.SetFocus
'        verificarCampos = False
'        Exit Function
'    End If
    
    If fg(0).Rows = fg(0).FixedRows Then
        MsgBox "No ha especificado items para la Solicitud de Materiales", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).SetFocus
        verificarCampos = False
        Exit Function
    End If
    
    verificarCampos = True
End Function

Function Grabar() As Boolean
    Dim IDSOL_ As Integer
    Dim FCHSOL_ As String
    Dim NUMSER_ As String
    Dim NUMDOC_ As String
    Dim IDRESP_ As Integer
    Dim IDTIPDOCREF_ As Integer
    Dim IDDOCREF_ As Integer
    Dim IDITEM_ As Integer
    Dim IDALM_ As Integer
    Dim IDESTADO_ As Integer
    Dim xRs As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
    Dim A As Integer
    
    If Not verificarCampos Then Grabar = False: Exit Function
    
    ' Se llenan los detalles
    If QueHace = 1 Then IDSOL_ = 0 Else IDSOL_ = NulosN(RstSol("id"))
    FCHSOL_ = TxtFchPro.valor
    NUMSER_ = NulosC(TxtNumSer.Text)
    NUMDOC_ = NulosC(TxtNumDoc.Text)
    IDRESP_ = NulosN(TxtIdResp.Text)
    IDALM_ = NulosN(txtIdAlm.Text)
    IDTIPDOCREF_ = NulosN(TxtIdTipDocRef.Text)
    IDDOCREF_ = NulosN(lbliddocref.Caption)
    IDESTADO_ = NulosN(cbEstado.ItemData(cbEstado.ListIndex))
    ' Se prepara el Recordset
    If xRs.State = 0 Then
        cSQL = "SELECT TOP 1 * FROM pro_solicitudmatdet"
        Set xRsAux = Nothing
        RST_Busq xRsAux, cSQL, xCon
        If xRsAux.State = 0 Then Grabar = False: Exit Function
        DEFINIR_RST_TMP xRs, xRsAux
        'preparaRST xRs, 2
    End If
    limpiarRST xRs
    ' Se llena el recordset
    For A = 1 To fg(0).Rows - 1
        xRs.AddNew
        xRs("iditem") = NulosN(fg(0).TextMatrix(A, COLUMNAIDITEM))
        xRs("idunimed") = NulosN(fg(0).TextMatrix(A, COLUMNAIDUNIMED))
        xRs("cantidad") = NulosN(fg(0).TextMatrix(A, COLUMNACANTIDAD))
        xRs("idlote") = NulosN(fg(0).TextMatrix(A, COLUMNAIDLOTE))
        xRs("idlotedet") = NulosN(fg(0).TextMatrix(A, COLUMNAIDLOTEDET))
        xRs.Update
    Next A
    
    ' Se graba el movimiento
    Grabar = grabarSolicitud(FCHSOL_, IDTIPDOCREF_, IDDOCREF_, IDRESP_, NUMDOC_, IDALM_, _
                                    xRs, NUMSER_, IDSOL_, IDESTADO_, CInt(AnoTra), mMesActivo, QueHace)

    mIdRegistro = IDSOL_
End Function

Sub Modificar()
    llenarEstado 1, 1, , cbEstado, , , False, ESTADOPENDIENTE_ & "," & ESTADOPROCESADO_ 'llenarEstados
    
    If NulosN(RstSol("estado")) > ESTADOPENDIENTE_ Then
        MsgBox "El registro está en un estado no modificable", vbInformation, Me.Caption
        Exit Sub
    End If
            
    If RstSol.RecordCount = 0 Then
        MsgBox "No hay Registros para Modificar", vbInformation, Me.Caption
        Exit Sub
    End If
   
    QueHace = 2
    xHorIni = Time
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    
    Label5.Caption = "Modificando Solicitud de Materiales"
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
    
    xHorIni = Time
    TxtIdResp.SetFocus
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    Dim xRs As New ADODB.Recordset
    
    If RstSol.RecordCount = 0 Then
        MsgBox "No hay documentos para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    TabOne1.CurrTab = 0
    Rpta = MsgBox("¿ Esta seguro de eliminar el Registro seleccionado ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_solicitudmatdet WHERE idsol = " & RstSol("id")
        xCon.Execute "DELETE * FROM pro_solicitudmat WHERE id = " & RstSol("id")
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstSol("id") & " AND idform = " & IdMenuActivo
        
        MsgBox "El registro se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstSol.Requery
        Dg1.Refresh
    End If
End Sub

Sub ExportarExcel()
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    Dim xCad As String
    Dim objExcel As Object
    
    Dim oExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    
    Set oExcel = New Excel.Application
    Set oWBook = oExcel.Workbooks.Add
    Screen.MousePointer = vbHourglass

    oExcel.WindowState = 2
    For A = 1 To fg(0).Rows - 1
        oWBook.ActiveSheet.Name = fg(0).TextMatrix(A, 10)
        With oWBook.ActiveSheet
            'Se llena cabecera
            .Cells(1, 2) = "SOLICITUD DE MATERIALES Nº:   " + "0001" & "-" & fg(0).TextMatrix(A, 10)
            .Range("B1", "H1").Merge
            .Cells(1, 2).HorizontalAlignment = xlHAlignCenterAcrossSelection
            .Cells(1, 2).Font.Bold = True
            .Cells(1, 2).Rows(1).Font.Size = 12
            
            .Cells(3, 2) = "Producción    Nº " + fg(0).TextMatrix(A, 6)
            .Cells(3, 2).Font.Bold = True
            .Cells(4, 2) = "Fch. Prod. :   " + TxtFchPro.valor
            .Cells(4, 2).Font.Bold = True
            .Cells(5, 2) = "Producto :   " + fg(0).TextMatrix(A, 2)
            .Cells(5, 2).Font.Bold = True
            .Cells(6, 2) = "Receta :   " + fg(0).TextMatrix(A, 5)
            .Cells(6, 2).Font.Bold = True
            
            .Cells(7, 2) = "Cantidad  :   " + fg(0).TextMatrix(A, 4)
            .Cells(7, 2).Font.Bold = True
            
            .Cells(9, 2) = "Item"
            .Cells(9, 2).Font.Bold = True
            .Cells(9, 3) = "INSUMO / PRODUCTO / MP"
            .Cells(9, 3).Font.Bold = True
            .Cells(9, 4) = "Uni. Med."
            .Cells(9, 4).Font.Bold = True
            .Cells(9, 5) = "Cantidad Teorica"
            .Cells(9, 5).Font.Bold = True
            .Cells(9, 6) = "Cantidad Real"
            .Cells(9, 6).Font.Bold = True
            .Cells(9, 7) = "Adicional"
            .Cells(9, 7).Font.Bold = True
            .Cells(9, 8) = "Devolucion"
            .Cells(9, 8).Font.Bold = True
            
            Dim Rst As New ADODB.Recordset
            
            RST_Busq Rst, "SELECT pro_recetains.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, [pro_recetains]![canpro]*" & NulosN(fg(0).TextMatrix(A, 3)) & " AS canreq " _
                + vbCr + "FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
                + vbCr + "WHERE (((pro_recetains.idrec)=" & NulosN(fg(0).TextMatrix(A, 8)) & "))", xCon
        
            If Rst.RecordCount <> 0 Then
                Dim xFila As Integer
                xFila = 10
                For B = 1 To Rst.RecordCount
                    .Cells(xFila, 2) = Format(B, "00")
                    .Cells(xFila, 3) = Rst("descripcion")
                    .Cells(xFila, 4) = Rst("abrev")
                    .Cells(xFila, 5) = Format(Rst("canreq"), FORMAT_CANTIDADDECIMAL)
    
                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                    xFila = xFila + 1
                Next B
            End If
            .Cells(xFila + 5, 5) = "VºBº Ger. Prod. "
            .Cells(xFila + 5, 5).Font.Bold = True
            .Cells(xFila + 5, 7) = "Entregado Por "
            .Cells(xFila + 5, 7).Font.Bold = True
        End With
        If A < fg(0).Rows - 1 Then oWBook.Sheets.Add
    Next A
    
    oExcel.Visible = True
    
    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Reporte de Pedidos"
    oExcel.WindowState = 1
    
    Set oExcel = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub limpiarRST(Rst As ADODB.Recordset, Optional TODO As Boolean = True)
    With Rst
        If .State <> 0 Then
            If TODO Then .Filter = adFilterNone
            If .RecordCount <> 0 Then
                .MoveFirst
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
            End If
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CAMBIOGRABAR_ = -1 Then
        MsgBox "No se puede Cancelar la operación; Grabe los registros para continuar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
    End If
End Sub

Private Sub lblAlmacen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAlmacen.ToolTipText = lblAlmacen.Caption
End Sub

Private Sub lblResponsable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblResponsable.ToolTipText = lblResponsable.Caption
End Sub

Private Sub LblTipDocRef_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LblTipDocRef.ToolTipText = LblTipDocRef.Caption
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 1 Then Exit Sub
        MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then
        If RstSol.RecordCount = 0 Then
            MsgBox "No se han registardos ventas para realizar esta opcion", vbInformation, Me.Caption
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.Index = 3 Then
        If RstSol.RecordCount = 0 Then
            MsgBox "No se han registrados Pedidos para realizar esta opción", vbInformation, Me.Caption
            Exit Sub
        End If
        Eliminar
    End If
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstSol.Requery
            Dg1.Refresh
            If RstSol.RecordCount <> 0 Then
                RstSol.MoveFirst
                RstSol.Find "id=" & mIdRegistro
                If RstSol.EOF = True Then RstSol.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        If TabOne1.CurrTab = 0 Then RstSol.Filter = "": TDB_FiltroLimpiar Dg1
    End If
    
    If Button.Index = 12 Then CambiarMes
    If Button.Index = 14 Then ExportarExcel
    If Button.Index = 15 Then imprimir
    If Button.Index = 17 Then Unload Me
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        If ButtonMenu.Index = 1 Then
            If TabOne1.CurrTab = 1 Then TabOne1.CurrTab = 0
            anular
        End If
    End If
End Sub

Private Function verificarCambioEstado(IDSOL_ As Integer, ByRef MENSAJE_ As String) As Boolean
    Dim xRs As New ADODB.Recordset
    
    ' -------------------------------------INGRESOS Y SALIDAS DE ALMACEN
    cSQL = "SELECT * " _
        + vbCr + "FROM alm_ingreso " _
        + vbCr + "WHERE (((alm_ingreso.idtipdocref)=110) AND ((alm_ingreso.iddocref)=" & IDSOL_ & ") AND ((alm_ingreso.estado)=" & ESTADOPROCESADO_ & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    MENSAJE_ = "Registros de Ingresos y Salidas"
    
    If xRs.State = 0 Then verificarCambioEstado = False: GoTo SALIR_
    If xRs.RecordCount > 0 Then verificarCambioEstado = False: GoTo SALIR_
    
    verificarCambioEstado = True
    Exit Function
    
SALIR_:
    MENSAJE_ = "Se han encontrado " & MENSAJE_ & " que se encuentran en un estado no modificable; " _
    & vbCr & "verifique la condición de dichos Registros para completar esta acción."
End Function

Private Sub anular()
    Dim MENSAJE_ As String
    Dim xRs As New ADODB.Recordset
    
    If verificarCambioEstado(NulosN(RstSol("id")), MENSAJE_) Then
        ' ----------------------------------------SE CAMBIA DE ESTADO AL REGISTRO DE INGRESO Y SALIDA
        cSQL = "UPDATE alm_ingreso SET alm_ingreso.estado = " & ESTADOANULADO_ & " " _
            + vbCr + "WHERE (((pro_solicitudmat.idtipdocref)=110) AND ((pro_solicitudmat.iddocref)=" & NulosN(RstSol("id")) & "));"
        ' --------------EJECUTA COMANDO
        xCon.Execute cSQL
        ' --------------ACTUALIZA VAR_EDICION
        cSQL = "SELECT alm_ingreso.id " _
            + vbCr + "FROM alm_ingreso " _
            + vbCr + "WHERE (((alm_ingreso.idtipdocref)=110) And ((alm_ingreso.iddocref)=" & NulosN(RstSol("ID")) & "))"
        
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        xRs.MoveFirst
        While Not xRs.EOF
            GrabarOperacion xIdUsuario, 8, 7, xHorIni, Time, Date, xCon, NulosN(xRs("id"))
            xRs.MoveNext
        Wend
        ' ----------------------------------------SE CAMBIA DE ESTADO AL REGISTRO
        xCon.Execute "UPDATE pro_solicitudmat SET pro_solicitudmat.estado = " & ESTADOANULADO_ & " WHERE (((pro_solicitudmat.id) = " & NulosN(RstSol("id")) & "))"
        MsgBox "El registro se anuló con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstSol.Requery
        Dg1.Refresh
    Else
        MsgBox MENSAJE_, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Sub imprimir()
    Dim RSTCAB_ As New ADODB.Recordset
    Dim RSTDET_ As New ADODB.Recordset
    Dim IDSOL_ As Integer
    Dim xform As New eps_librerias.FormSeleccion
    Dim nSQLId As String
    Dim nSQLId2 As String
    Dim xCampos() As String
    Dim xRs As New ADODB.Recordset
    
    If TabOne1.CurrTab = 0 Then
        ReDim xCampos(7, 5) As String
        
        xCampos(0, 0) = "Fecha":            xCampos(0, 1) = "fchdoc":           xCampos(0, 2) = "1000":      xCampos(0, 3) = "C":   xCampos(0, 4) = "C"
        xCampos(1, 0) = "Nº Documento":     xCampos(1, 1) = "numdoc":           xCampos(1, 2) = "1200":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
        xCampos(2, 0) = "Responsable":      xCampos(2, 1) = "desresp":          xCampos(2, 2) = "2300":     xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
        xCampos(3, 0) = "Almacén":          xCampos(3, 1) = "desalm":           xCampos(3, 2) = "2300":     xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
        xCampos(4, 0) = "TD Ref.":          xCampos(4, 1) = "destipdocref":     xCampos(4, 2) = "700":      xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
        xCampos(5, 0) = "NºDoc.Ref.":       xCampos(5, 1) = "numdocref":        xCampos(5, 2) = "1500":     xCampos(5, 3) = "C":    xCampos(5, 4) = "N"
        xCampos(6, 0) = "Anexo":            xCampos(6, 1) = "anexo":            xCampos(6, 2) = "1500":     xCampos(6, 3) = "C":    xCampos(6, 4) = "C"

        cSQL = "SELECT 0 AS xsel, pro_solicitudmat.id, Format(pro_solicitudmat.fchdoc,'Short Date') AS fchdoc, [pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc] AS numdoc, pla_empleados.nombre AS desresp, alm_almacenes.descripcion AS desalm, mae_documento.abrev AS destipdocref, IIf([pro_solicitudmat].[idtipdocref]=115,[pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc],'') AS numdocref, UCase([mae_estados].[descripcion]) AS desestado, pro_solicitudmat.estado, alm_inventario.descripcion AS anexo " _
            + vbCr + "FROM ((((((pro_solicitudmat LEFT JOIN pla_empleados ON pro_solicitudmat.idresp = pla_empleados.id) LEFT JOIN mae_estados ON pro_solicitudmat.estado = mae_estados.id) LEFT JOIN mae_documento ON pro_solicitudmat.idtipdocref = mae_documento.id) LEFT JOIN alm_almacenes ON pro_solicitudmat.idalm = alm_almacenes.id) LEFT JOIN pro_ordenprod ON pro_solicitudmat.iddocref = pro_ordenprod.id) LEFT JOIN pro_receta ON pro_ordenprod.idrec = pro_receta.id) LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id " _
            + vbCr + "WHERE (((pro_solicitudmat.ano) = " & AnoTra & ") And ((pro_solicitudmat.idmes) = " & mMesActivo & ")) " _
            + vbCr + "ORDER BY Format([fchdoc],'Short Date') DESC, [pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc] DESC;"
        
        xform.SQLCad = cSQL
        xform.titulo = "Operaciones a Imprimir"
        Set xform.Coneccion = xCon
        
        Set xRs = Nothing
        Set xRs = xform.seleccionar(xCampos)
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        nSQLId = GENERAR_SQL_ID_RST(xRs, "id", " AND pro_solicitudmat.id", "IN", True)
        nSQLId2 = GENERAR_SQL_ID_RST(xRs, "id", " AND pro_solicitudmatdet.idsol", "IN", True)
    Else
        If NulosN(RstSol("estado")) = ESTADOPENDIENTE_ Then
            MsgBox "El registro actual no se puede imprimir debido a su estado", vbInformation, xTitulo
            Exit Sub
        End If
        nSQLId = " AND pro_solicitudmat.id=" & NulosN(RstSol("id"))
        nSQLId2 = " AND pro_solicitudmatdet.idsol=" & NulosN(RstSol("id"))
    End If
    
    ' SE CREA CABECERA
    cSQL = "SELECT pro_solicitudmat.id, pro_solicitudmat.fchdoc, pro_solicitudmat.numser, pro_solicitudmat.numdoc, pro_solicitudmat.idtipdocref, pro_solicitudmat.iddocref, alm_almacenes.descripcion AS desalm, mae_documento.descripcion AS destipdocref, pla_empleados.nombre AS desresp, IIf([pro_solicitudmat].[idtipdocref]=115,[pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc],'') AS numdocref " _
        + vbCr + "FROM (((pro_solicitudmat LEFT JOIN mae_documento ON pro_solicitudmat.idtipdocref = mae_documento.id) LEFT JOIN pla_empleados ON pro_solicitudmat.idresp = pla_empleados.id) LEFT JOIN alm_almacenes ON pro_solicitudmat.idalm = alm_almacenes.id) LEFT JOIN pro_ordenprod ON pro_solicitudmat.iddocref = pro_ordenprod.id " _
        + vbCr + "Where (((pro_solicitudmat.ano) = " & AnoTra & ") And ((pro_solicitudmat.idmes) = " & mMesActivo & ")) " & nSQLId _
        + vbCr + "ORDER BY pro_solicitudmat.numdoc;"
        
    Set RSTCAB_ = Nothing
    RST_Busq RSTCAB_, cSQL, xCon
    
    ' SE CREA EL DETALLE
    cSQL = "SELECT pro_solicitudmatdet.idsol, pro_solicitudmatdet.iditem, alm_inventariolote.descripcion AS deslote, alm_inventario.descripcion AS desitem, alm_inventario.codpro AS coditem, mae_unidades.abrev AS desunimed, pro_solicitudmatdet.cantidad " _
        + vbCr + "FROM ((pro_solicitudmatdet LEFT JOIN alm_inventariolote ON pro_solicitudmatdet.idlote = alm_inventariolote.id) LEFT JOIN alm_inventario ON pro_solicitudmatdet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_solicitudmatdet.idunimed = mae_unidades.id " _
        + vbCr + "WHERE (((pro_solicitudmatdet.cantidad) <> 0)) " & nSQLId2 _
        + vbCr + "ORDER BY alm_inventario.descripcion;"

    Set RSTDET_ = Nothing
    RST_Busq RSTDET_, cSQL, xCon
        
    ImprimirSolicitud RSTCAB_, RSTDET_
End Sub

Private Sub CambiarMes()
    TabOne1.CurrTab = 0
    mMesActivo = SeleccionaMes(xCon)
    lblperiodo.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    pCargarGrid
End Sub

Private Sub TxtIdResp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 2
    End If
End Sub

Private Sub TxtIdTipDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 4
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If NulosC(TxtNumSer.Text) = "" Then
        MsgBox "Ingrese un número de serie", vbInformation, Me.Caption
        TxtNumDoc.Text = ""
        TxtNumSer.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtNumDocRef_KeyPress(KeyAscii As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub txtNumDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 5
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtNumSer.Text) <> "" Then
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
        TxtNumDoc.Text = hallarNumDoc("pro_solicitudmat", "'" & NulosC(TxtNumSer.Text) & "'", "numser", , , , , , "0000000000")
    End If
End Sub

Private Sub txtIdAlm_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 1
    End If
End Sub
