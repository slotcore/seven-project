VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmManRutasRutas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEVEN - Mantenimiento de Rutas de Acceso"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9330
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
            Picture         =   "FrmManRutasRutas.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRutasRutas.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRutasRutas.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRutasRutas.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRutasRutas.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRutasRutas.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRutasRutas.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRutasRutas.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRutasRutas.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRutasRutas.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRutasRutas.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRutasRutas.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   4980
      Left            =   15
      TabIndex        =   5
      Top             =   375
      Width           =   9330
      _cx             =   16457
      _cy             =   8784
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4560
         Left            =   -9885
         TabIndex        =   11
         Top             =   375
         Width           =   9240
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   4185
            Left            =   30
            TabIndex        =   16
            Top             =   345
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   7382
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripcion"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Archivo"
            Columns(2).DataField=   "archivo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tipo de Ruta"
            Columns(3).DataField=   "tipo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=767"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=688"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=7488"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7408"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=4577"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4498"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2302"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2223"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Named:id=33:Normal"
            _StyleDefs(53)  =   ":id=33,.parent=0"
            _StyleDefs(54)  =   "Named:id=34:Heading"
            _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   ":id=34,.wraptext=-1"
            _StyleDefs(57)  =   "Named:id=35:Footing"
            _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   "Named:id=36:Selected"
            _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=37:Caption"
            _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(63)  =   "Named:id=38:HighlightRow"
            _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=39:EvenRow"
            _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(67)  =   "Named:id=40:OddRow"
            _StyleDefs(68)  =   ":id=40,.parent=33"
            _StyleDefs(69)  =   "Named:id=41:RecordSelector"
            _StyleDefs(70)  =   ":id=41,.parent=34"
            _StyleDefs(71)  =   "Named:id=42:FilterBar"
            _StyleDefs(72)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consultando Rutas de Acceso"
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
            Left            =   75
            TabIndex        =   13
            Top             =   30
            Width           =   9105
         End
         Begin VB.Label LblPeriodo 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   9810
            TabIndex        =   12
            Top             =   0
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4560
         Left            =   45
         TabIndex        =   6
         Top             =   375
         Width           =   9240
         Begin VB.Frame Frame3 
            Height          =   3600
            Left            =   420
            TabIndex        =   7
            Top             =   675
            Width           =   8460
            Begin VB.OptionButton OptNo 
               Caption         =   "&No"
               Height          =   195
               Left            =   5100
               TabIndex        =   4
               Top             =   3270
               Width           =   555
            End
            Begin VB.OptionButton OptSi 
               Caption         =   "&Si"
               Height          =   195
               Left            =   4455
               TabIndex        =   3
               Top             =   3270
               Width           =   555
            End
            Begin VB.CommandButton CmdCrearIni 
               Caption         =   "Crear Archivo INI y establecer como ruta de Acceso"
               Height          =   345
               Left            =   4290
               TabIndex        =   17
               Top             =   255
               Width           =   3930
            End
            Begin VB.TextBox TxtId 
               Height          =   300
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   0
               Text            =   "TxtId"
               Top             =   300
               Width           =   915
            End
            Begin VB.TextBox TxtArch 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2220
               Left            =   1200
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   2
               Text            =   "FrmManRutasRutas.frx":277E
               Top             =   930
               Width           =   7020
            End
            Begin VB.TextBox TxtDescripcion 
               Height          =   300
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   1
               Text            =   "TxtDescripcion"
               Top             =   615
               Width           =   7035
            End
            Begin VB.Label Label1 
               Caption         =   "Esta configuracion sera utilizada como ruta del servidor"
               Height          =   210
               Left            =   195
               TabIndex        =   18
               Top             =   3255
               Width           =   6330
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Codigo"
               Height          =   195
               Index           =   1
               Left            =   195
               TabIndex        =   15
               Top             =   330
               Width           =   495
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Ruta"
               Height          =   195
               Index           =   0
               Left            =   195
               TabIndex        =   9
               Top             =   960
               Width           =   345
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Descripcion"
               Height          =   195
               Index           =   10
               Left            =   195
               TabIndex        =   8
               Top             =   645
               Width           =   840
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Ruta de Acceso"
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
            Left            =   60
            TabIndex        =   10
            Top             =   30
            Width           =   9105
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Desactivar Usuario"
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
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmManRutasRutas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo    : FRMMANRUTASRUTAS
'* Tipo              : FORMULARIO
'* Descripcion       : FORMULARIO PARA LA CREACION DE ARCHIVOS INI
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 03/09/09
'* VERSION           : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstRut As New ADODB.Recordset       ' RECORSET PRINCIPAL PARA CARGAR TODOS LOS REGISTROS
Dim QueHace As Integer                  ' VARIABLE PARA IDENTIFICAR LAS ACCIONES SOBRE EL FORMULARIO (1 = NUEVO,2 = MODIFICAR, 3 = SOLOLECTURA)
Dim SeEjecuto As Boolean                ' VARIABLE QUE INDICARA SI EL FORMULARIO YA EJECUTO EL EVENTO LOAD
Dim xConRuta As New ADODB.Connection

'*****************************************************************************************************
'* Nombre Modulo  : Grabar()
'* Tipo           : FUNCCION
'* Descripcion    : PERMITE GUARDAR LOS DATOS EDITADOS EN EL FORMULARIO, RETORANA UN VALOR VERDADERO
'*                  CUANDO EL REGISTRO SE GUARDA CON EXITO
'* Paranetros     : NULL
'* Retorna        : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SON LOS CORRECTOS
    If NulosC(TxtDescripcion.Text) = "" Then
        MsgBox "No ha especificado la descripcion para la ruta de acceso", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtDescripcion.SetFocus
        Exit Function
    End If

    If NulosC(TxtArch.Text) = "" Then
        MsgBox "No ha especificado la ruta de acceso", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtArch.SetFocus
        Exit Function
    End If
    
    If OptSi.Value = False And OptNo.Value = False Then
        MsgBox "No ha especificado si la ruta de acceso se usara como ruta de servidor", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtId.SetFocus
        Exit Function
    End If
    
    Dim Rst As New ADODB.Recordset
    Dim RstBus As New ADODB.Recordset
    Dim xId, Rpta As Integer
    
    ' PREGUNTAMOS QUE ES LO QUE HACE
    If QueHace = 1 Then
        ' SI SE ESTA AGREGANDO UN REGISTRO, OBTENEMOS EL NUMERO ID PARA EL REGISTRO
        xId = HallaCodigoTabla("mae_inis", xConRuta, "id")
        RST_Busq Rst, "SELECT * FROM mae_inis", xConRuta
        ' CREAMOS UN NUEVO REGISTRO
        Rst.AddNew
        Rst("id") = xId
    Else
        ' BUSCAMOS EL REGISTRO Y TRAEMOS LOS DATOS
        RST_Busq Rst, "SELECT * FROM mae_inis WHERE id = " & RstRut("id") & "", xConRuta
    End If
    
    '' ASIGNAMOS LOS DATOS A CADA CAMPO
    Rst("descripcion") = NulosC(TxtDescripcion.Text)
    Rst("archivo") = NulosC(TxtArch.Text)
    If OptSi.Value = True Then
        RST_Busq RstBus, "SELECT * FROM mae_inis WHERE servidor = -1", xConRuta
        If RstBus.RecordCount <> 0 Then
            Rpta = MsgBox("Ya existe una ruta definida como ruta de ssrvidor ¿Desea que la ruta que se esta guardando sea la ruta del servidor?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                xConRuta.Execute "UPDATE mae_inis SET mae_inis.activo = 0 WHERE (((mae_inis.id)=" & RstBus("id") & "))"
                Rst("servidor") = -1
            Else
                Rst("servidor") = 0
            End If
        Else
            Rst("servidor") = -1
        End If
        Set RstBus = Nothing
    Else
        Rst("servidor") = 0
    End If
    Rst.Update
End Function

Private Sub CmdCrearIni_Click()
    Dim xArch As New FileSystemObject
    Dim Rpta As Integer
    Dim Ruta, Ruta2, xRutaServ As String
    
    Rpta = MsgBox("Esta seguro de reemplazar el archivo INI actual", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        If xArch.FileExists(Trim(App.Path) + "\seven.ini") = True Then
            xArch.DeleteFile Trim(App.Path) + "\seven.ini"
        End If
        
        Set xArch = Nothing
        Open Trim(App.Path) + "\seven.ini" For Output As #1
        Print #1, RstRut("archivo")
        Close #1
        
        xConRuta.Execute "UPDATE mae_inis SET mae_inis.activo = -1 WHERE (((mae_inis.id)=" & RstRut("id") & "));"
        
        If RstRut("servidor") = 0 Then
            'si la ruta que se esta estableciendo no es el servidor, se pregunta si se quiere copiar la BD a la nuevar ruta
            Rpta = MsgBox("Ha decidido que la ruta de acceso no sea una ruta establecida como servidor" & Chr(13) _
                & "¿Desea hacer una copia de la BD a la nueva ruta de acceso? ", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Ruta = LeerLineaINI(Trim(App.Path) & "\seven.ini", "RUTABD", "RUTAS")
                MsgBox "Se compiarara la base de datos a la siguiente direccion " + Trim(Ruta), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                
                Dim RstTmp As New ADODB.Recordset
                
                Set RstTmp = BuscaConCriterio("SELECT * FROM mae_inis WHERE servidor=-1", xConRuta)
                'escribimos la ruta del servidor en el ini
                Open Trim(App.Path) + "\seven.ini" For Output As #1
                Print #1, RstTmp("archivo")
                Close #1
                xRutaServ = LeerLineaINI(Trim(App.Path) & "\seven.ini", "RUTABD ", "RUTAS")
                
                If xArch.FolderExists(Ruta) = True Then
                    xArch.DeleteFolder Mid(Trim(Ruta), 1, Len(Trim(Ruta)) - 1)
                End If
                
                xArch.CopyFolder Mid(xRutaServ, 1, Len(Trim(xRutaServ)) - 1), Mid(Trim(Ruta), 1, Len(Trim(Ruta)) - 1)
                Set RstTmp = Nothing
                MsgBox "Los archivos de datos se transfirieron con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Else
                MsgBox "Ha especificado una nueva ruta de acceso, pero no se ha copiado la base de datos, es posible que" & Chr(13) _
                    & "Seven no pueda volver a ejecutarse la proxima vez que lo inicie", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            End If
            Open Trim(App.Path) + "\seven.ini" For Output As #1
            Print #1, RstRut("archivo")
            Close #1
            
            MsgBox "El SEVEN requiere reiniciarse ¿Haga clic en aceptar par reiniciar el SEVEN?", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set xCon = Nothing
            Set xConRuta = Nothing
            Unload Me
            End
        End If
        TabOne1.CurrTab = 0
        MsgBox "El archivo INI se reemplazo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Sub Form_Activate()
    ' ES EL SEGUNDO EVENTO QUE SE EJECUTAR AL CARGAR EL FORMULARIO
    Dim Rpta As Integer
    
    If SeEjecuto = False Then
        
        ' ABRIMOS LA CONECCION A LA BD DE ENLACE PARA PODER REALIZARLAS OPERACIONES
        Dim xFun As New eps_librerias.FuncionesData
        
        xFun.F_BASEDATOS = AP_RUTABD + "data.mdb"                                           ' PASAMOS LA RUTA DE LA BASE DE DATOS PARA ABRIR LA CONECCION
        xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"                                       ' PASAMOS LA RUTA DEL ARCHIVO DE TRABJO DE LA BASE DE DATOS
        xFun.F_PASSWORD = Eps_Pass                                                          ' PASAMOS EL PASWORD DE LA BASE DE DATOS
        xFun.F_USUARIO = Eps_User                                                           ' PASAMOS EL USUARIO DE LA BASE DE DATOS
        xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"                                        ' PASAMOS EL NOMBRE DEL PROVEEDORE DE DATOS PARA ADO 2.5
        
        Set xConRuta = xFun.AbrirConeccion                                                      ' ABRIMOS LA CONECCION DE DATOS
        Set xFun = Nothing
               
        SeEjecuto = True
        ' CARGAR TODO LOS REGISTROS EXISTENTE
        RST_Busq RstRut, "SELECT mae_inis.*, IIf([servidor]=-1,'Servidor','Local') AS tipo FROM mae_inis", xConRuta
        ' MUESTRA LA INFORMACION EN EL DATAGRID
        Set Dg1.DataSource = RstRut
        
        ' PREGUNTAMOS SI HAY DATOS
        If RstRut.RecordCount = 0 Then
            ' SIN O HAY DATOS PREGUNTAMOS SI SE AGREGARA UNO
            Rpta = MsgBox("No se ha registrado ninguna ruta, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                ' SI ES SI SE AGREGA UN NUEVO REGISTRO
                Nuevo
            Else
                ' SI ES NO SE SALE DEL FORMULARIO
                Set RstRut = Nothing
                Unload Me
            End If
        End If
        
        Dg1.SetFocus
    End If
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Nuevo()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : LE DICE AL FORMULARIO QUE SE AGREGARA UN NUEVO REGISTRO, PARA ELLO ACTUALIZA EL
'*                  EL VALOR DE LA VARIABLE QUEHACE = 1
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Nuevo()
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Blanquea
    Bloquea
    TxtId.Text = HallaCodigoTabla("mae_inis", xConRuta, "id")
    TxtDescripcion.SetFocus
End Sub

Private Sub Form_Load()
    ' ES EL PRIMER EVENTO QUE SE EJECUTARA AL CARGAR EL FORMULARIO
    TabOne1.CurrTab = 0
    QueHace = 3
    SeEjecuto = False
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Bloquea()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : ACTIVA O DESACTIVA LOS CONTROLES DEL FORMULARIO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Bloquea()
    TxtDescripcion.Locked = Not TxtDescripcion.Locked
    TxtArch.Locked = Not TxtArch.Locked
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Blanquea()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : BLANQUEA LOS CONTROLES DEL FORMULARIO PARA EL INGRESO DE DATOS
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Blanquea()
    TxtId.Text = ""
    TxtDescripcion.Text = ""
    TxtArch.Text = ""
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : MuestraSegundoTab()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : MUESTRA LOS DATOS AL DETALLE DEL REGISTRO SELECCIONADO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub MuestraSegundoTab()
    TxtId.Text = RstRut("id")
    TxtDescripcion.Text = RstRut("descripcion")
    TxtArch.Text = RstRut("archivo")
    If RstRut("servidor") = -1 Then
        OptSi.Value = True
    Else
        OptNo.Value = True
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 8 Then
        Set RstRut = Nothing
        Unload Me
    End If
End Sub
