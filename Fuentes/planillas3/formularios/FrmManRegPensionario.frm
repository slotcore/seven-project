VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmManRegPensionario 
   Caption         =   "Planillas - Régimen Pensionario"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8235
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegPensionario.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegPensionario.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegPensionario.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegPensionario.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegPensionario.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegPensionario.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegPensionario.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegPensionario.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegPensionario.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegPensionario.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegPensionario.frx":1EA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   609
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
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5610
      Left            =   30
      TabIndex        =   8
      Top             =   360
      Width           =   7890
      _cx             =   13917
      _cy             =   9895
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   12632256
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   12632256
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   "   Consulta   |   Detalles   "
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
         Height          =   5190
         Left            =   45
         TabIndex        =   11
         Top             =   375
         Width           =   7800
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   4710
            Left            =   45
            TabIndex        =   12
            Top             =   390
            Width           =   7740
            _ExtentX        =   13653
            _ExtentY        =   8308
            _LayoutType     =   4
            _RowHeight      =   14
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Código"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Abrev"
            Columns(2).DataField=   "abrev"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=8811"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8731"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1852"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1773"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bgcolor=&HDBFDFD&,.fgcolor=&H0&,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.fgcolor=&HFF0000&,.bold=0"
            _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(25)  =   ":id=13,.strikethrough=0,.charset=0"
            _StyleDefs(26)  =   ":id=13,.fontname=MS Sans Serif"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.namedParent=33,.fgcolor=&H800000&"
            _StyleDefs(29)  =   ":id=14,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(30)  =   ":id=14,.fontname=MS Sans Serif"
            _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H80&,.bold=0"
            _StyleDefs(34)  =   ":id=18,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(35)  =   ":id=18,.fontname=MS Sans Serif"
            _StyleDefs(36)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(37)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(38)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(39)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(40)  =   ":id=21,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(41)  =   ":id=21,.fontname=MS Sans Serif"
            _StyleDefs(42)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(43)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.bold=0,.fontsize=825"
            _StyleDefs(45)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(46)  =   ":id=28,.fontname=MS Sans Serif"
            _StyleDefs(47)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
            _StyleDefs(48)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(51)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(52)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(58)  =   "Named:id=33:Normal"
            _StyleDefs(59)  =   ":id=33,.parent=0"
            _StyleDefs(60)  =   "Named:id=34:Heading"
            _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(62)  =   ":id=34,.wraptext=-1"
            _StyleDefs(63)  =   "Named:id=35:Footing"
            _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(65)  =   "Named:id=36:Selected"
            _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(67)  =   "Named:id=37:Caption"
            _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(69)  =   "Named:id=38:HighlightRow"
            _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(71)  =   "Named:id=39:EvenRow"
            _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(73)  =   "Named:id=40:OddRow"
            _StyleDefs(74)  =   ":id=40,.parent=33"
            _StyleDefs(75)  =   "Named:id=41:RecordSelector"
            _StyleDefs(76)  =   ":id=41,.parent=34"
            _StyleDefs(77)  =   "Named:id=42:FilterBar"
            _StyleDefs(78)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Régimen Pensionario"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   105
            TabIndex        =   13
            Top             =   60
            Width           =   7830
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5190
         Left            =   8535
         TabIndex        =   9
         Top             =   375
         Width           =   7800
         Begin VB.TextBox txt 
            BackColor       =   &H0000FFFF&
            Height          =   285
            Index           =   3
            Left            =   6840
            TabIndex        =   22
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   75
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Caption         =   "Frame12"
            Height          =   570
            Index           =   0
            Left            =   45
            TabIndex        =   21
            Top             =   4605
            Width           =   7680
            Begin VB.CommandButton CmdDet 
               Caption         =   "&Agregar"
               Enabled         =   0   'False
               Height          =   435
               Index           =   0
               Left            =   120
               TabIndex        =   3
               Top             =   60
               Width           =   1395
            End
            Begin VB.CommandButton CmdDet 
               Caption         =   "&Eliminar"
               Enabled         =   0   'False
               Height          =   435
               Index           =   1
               Left            =   3165
               TabIndex        =   6
               Top             =   60
               Width           =   1395
            End
            Begin VB.CommandButton CmdDet 
               Caption         =   "&Seleccionar"
               Enabled         =   0   'False
               Height          =   435
               Index           =   2
               Left            =   1530
               TabIndex        =   4
               Top             =   60
               Width           =   1395
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   2
               X1              =   15
               X2              =   15
               Y1              =   0
               Y2              =   1000
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               Index           =   3
               X1              =   7665
               X2              =   7665
               Y1              =   -15
               Y2              =   970
            End
            Begin VB.Line lin 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               Index           =   3
               X1              =   15
               X2              =   13045
               Y1              =   555
               Y2              =   555
            End
            Begin VB.Line lin 
               BorderColor     =   &H80000009&
               BorderWidth     =   2
               Index           =   2
               X1              =   -15
               X2              =   13000
               Y1              =   15
               Y2              =   15
            End
         End
         Begin VB.Frame Frame4 
            Height          =   1725
            Left            =   45
            TabIndex        =   16
            Top             =   360
            Width           =   7680
            Begin VB.Frame FraCta 
               Caption         =   "¿Cta Contable Resumen?"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   1740
               TabIndex        =   24
               Top             =   975
               Width           =   5835
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   0
                  Left            =   2520
                  Picture         =   "FrmManRegPensionario.frx":23EC
                  Style           =   1  'Graphical
                  TabIndex        =   27
                  ToolTipText     =   "Seleccione el Tipo de Concepto (Primero seleccione la categoría)"
                  Top             =   255
                  Width           =   210
               End
               Begin VB.OptionButton OptCta 
                  Caption         =   "Si"
                  Height          =   285
                  Index           =   1
                  Left            =   720
                  TabIndex        =   26
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   435
               End
               Begin VB.OptionButton OptCta 
                  Caption         =   "No"
                  Height          =   285
                  Index           =   0
                  Left            =   165
                  TabIndex        =   25
                  Top             =   270
                  Width           =   510
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   0
                  Left            =   1515
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   28
                  Text            =   "txt_cb(0)"
                  Top             =   225
                  Width           =   1260
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cta"
                  Height          =   195
                  Index           =   0
                  Left            =   1230
                  TabIndex        =   30
                  Top             =   315
                  Width           =   240
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(0)"
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
                  Height          =   285
                  Index           =   0
                  Left            =   4725
                  TabIndex        =   29
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   975
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
                  Height          =   285
                  Index           =   0
                  Left            =   2760
                  TabIndex        =   31
                  Top             =   225
                  Width           =   3000
               End
            End
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   1
               Left            =   1110
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   7
               Text            =   "txt(1)"
               Top             =   240
               Width           =   5430
            End
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   2
               Left            =   1110
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   1
               Text            =   "txt(2)"
               Top             =   570
               Width           =   915
            End
            Begin VB.Frame FraSpp 
               Caption         =   "¿Es SPP ?"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   105
               TabIndex        =   17
               Top             =   975
               Width           =   1605
               Begin VB.OptionButton opt 
                  Caption         =   "No"
                  Height          =   285
                  Index           =   0
                  Left            =   165
                  TabIndex        =   2
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   585
               End
               Begin VB.OptionButton opt 
                  Caption         =   "Si"
                  Height          =   285
                  Index           =   1
                  Left            =   825
                  TabIndex        =   18
                  Top             =   270
                  Width           =   555
               End
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   20
               Top             =   330
               Width           =   840
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cod.Sunat"
               Height          =   195
               Index           =   2
               Left            =   105
               TabIndex        =   19
               Top             =   660
               Width           =   750
            End
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0000FFFF&
            Height          =   285
            Index           =   0
            Left            =   10665
            TabIndex        =   14
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   -15
            Visible         =   0   'False
            Width           =   945
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2370
            Left            =   45
            TabIndex        =   5
            Top             =   2160
            Width           =   7680
            _cx             =   13547
            _cy             =   4180
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
            ForeColorSel    =   16777215
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
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManRegPensionario.frx":251E
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   3
            Left            =   6315
            TabIndex        =   23
            Top             =   165
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   10125
            TabIndex        =   15
            Top             =   75
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Régimen Pensionario"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   75
            TabIndex        =   10
            Top             =   60
            Width           =   8295
         End
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu2 
         Caption         =   "&Agregar"
      End
      Begin VB.Menu Menu3 
         Caption         =   "&Seleccionar"
      End
      Begin VB.Menu Menu4 
         Caption         =   "-"
      End
      Begin VB.Menu Menu5 
         Caption         =   "&Eliminar"
      End
   End
End
Attribute VB_Name = "FrmManRegPensionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstFrm As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim Agregando As Boolean
Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim mIdRegistro& '--identificador del registro


Sub Cancelar()
    ActivaTool
    Label5.Caption = "Detalle del Régimen Pensionario"
    QueHace = 3
    Bloquea False
    Fg1.SelectionMode = flexSelectionByRow
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub


Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstFrm("id")), xCon
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 7 Then KeyAscii = 0
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If QueHace <> 2 Then PopupMenu Menu1
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
  
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        pConfigurarGrilla
        
        pCargarGrid
    
    End If

    
End Sub


Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    CentrarFrm Me

    TabOne1.CurrTab = 0

End Sub


Sub Blanquea()
    LimpiaText txt
    LimpiaText txt_cb
    LimpiaText lbl_cb
    LimpiaText lbl_cod
    
    Fg1.Rows = 1
End Sub

Sub Bloquea(band As Boolean)

    habilitar_Locked txt, Not band
    habilitar CmdDet, band
    FraSpp.Enabled = band
    FraCta.Enabled = band
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set RstFrm = Nothing
End Sub

Private Sub Menu2_Click()
    pRegistroAdd False
End Sub

Private Sub Menu3_Click()
    pRegistroAdd True
End Sub

Private Sub Menu5_Click()
    pRegistroDel
End Sub

Private Sub opt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then CmdDet(0).SetFocus
End Sub

Private Sub OptCta_Click(Index As Integer)
    If OptCta(1).Value = True Then '--es resumen
        Fg1.ColWidth(6) = 0
        Fg1.ColWidth(7) = 0
        If QueHace = 3 Then Exit Sub
        habilitar cb, True
        habilitar_Locked txt_cb, False
    Else '--cuenta es en el detalle
        Fg1.ColWidth(6) = 1000
        Fg1.ColWidth(7) = 3200
        LimpiaText txt_cb
        If QueHace = 3 Then Exit Sub
        habilitar cb, False
        habilitar_Locked txt_cb, True
        
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Function Grabar() As Boolean

    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " el Régimen Pensionario", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
   
    Dim xId As Double
    Dim A&

    On Error GoTo LaCague

    xCon.BeginTrans

    If QueHace = 1 Then

        xId = HallaCodigoTabla("mae_regimenpen", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM mae_regimenpen", xCon

        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstFrm("id")
        
        RST_Busq RstCab, "SELECT * FROM mae_regimenpen WHERE id = " & xId & "", xCon
        '--eliminando concepto relacionado al regimen pensionario
        xCon.Execute "Delete from pla_conceptoregpen where idregpen =  " & xId & " ;"
        
    End If

    mIdRegistro = xId

    RST_Busq RstDet, "SELECT TOP 1 * FROM pla_conceptoregpen ; ", xCon

    RstCab("descripcion") = NulosC(txt(1).Text)
    RstCab("codsun") = Mid(NulosC(txt(2).Text), 1, RstCab("codsun").DefinedSize)
    If opt(1).Value = True Then
        RstCab("cuspp") = -1
    Else
        RstCab("cuspp") = 0
    End If
    
    If OptCta(0).Value = True Then
        RstCab("ctaresumen") = 0
    Else
        RstCab("ctaresumen") = -1
    End If
    
    RstCab("idcuenta") = NulosN(lbl_cod(0).Caption) '--cuenta contable
    
    RstCab.Update
    
    '--detalle de conceptos
    For A = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(A, 1)) <> 0 Then
            RstDet.AddNew
            RstDet("idregpen") = xId
            RstDet("idcpto") = NulosN(Fg1.TextMatrix(A, 1))
            '--verificar si se escribira en el detalle
            If OptCta(0).Value = True Then
                RstDet("idcuenta") = NulosN(Fg1.TextMatrix(A, 8)) '--cuenta contable
            Else
                RstDet("idcuenta") = 0
            End If
            RstDet.Update
        End If
    Next A
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    
    xCon.CommitTrans
    
    MsgBox "Los datos del Régimen Pensionario se " + IIf(QueHace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo

    Set RstCab = Nothing
    Set RstDet = Nothing
    Grabar = True
    Exit Function
LaCague:
    Set RstCab = Nothing
    Set RstDet = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar registro por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea True
    Blanquea
    Label5.Caption = "Agregando Régimen Pensionario"
    txt(1).SetFocus
End Sub

Sub Modificar()
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    Label5.Caption = "Modificando Régimen Pensionario"
    
    ActivaTool
    
    Bloquea True
    
    If TabOne1.CurrTab = 0 Then TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    QueHace = 2
    xHorIni = Time
    Agregando = False
    txt(1).SetFocus
    
End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Eliminar()
    TabOne1.CurrTab = 0
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros para Eliminar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    '--ver si hay personal con este regimen
    Dim RstBus As New ADODB.Recordset
    Dim nSQL As String
    nSQL = "SELECT pla_categoria1.idregpen , 'Trabajador' as categoria FROM pla_categoria1 WHERE (((pla_categoria1.idregpen)=1)) " _
        + vbCr + " Union " _
        + vbCr + " SELECT pla_categoria2.idregpen , 'Pensionista' as categoria  FROM pla_categoria2 WHERE (((pla_categoria2.idregpen)=1))"
    RST_Busq RstBus, nSQL, xCon
    If RstFrm.RecordCount <> 0 Then
        MsgBox "No se puede eliminar el Régimen Pensionario, pues hay Personal con Categoría: " & RstBus.Fields("categoria") & "con este Régimen", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstBus = Nothing
        Exit Sub
    End If
    Set RstBus = Nothing
    
    Dim Rpta As Integer
    Rpta = MsgBox("Esta seguro de eliminar el Régimen Pensionario seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE FROM pla_conceptoregpen WHERE idregpen=" & RstFrm.Fields("id") & " ;"
        xCon.Execute "DELETE FROM mae_regimenpen WHERE id = " & RstFrm("id") & ""
        MsgBox "El Régimen Pensionario se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstFrm.Requery
        Dg1.Refresh
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo

    If Button.Index = 2 Then Modificar

    If Button.Index = 3 Then Eliminar

    If Button.Index = 5 Then Cancelar

    If Button.Index = 6 Then
        If Grabar = True Then
            RstFrm.Requery
            Dg1.Refresh
            Cancelar

            If RstFrm.RecordCount <> 0 Then
                RstFrm.MoveFirst
                RstFrm.Find "id=" & mIdRegistro
                If RstFrm.EOF = True Then RstFrm.MoveFirst
            End If
            
        End If
    End If

    If Button.Index = 10 Then Buscar

    If Button.Index = 14 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub


Sub Buscar()
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    Dim nSQL As String
    xCampos(0, 0) = "Descripción":   xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "7000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":            xCampos(1, 1) = "id":           xCampos(1, 2) = "700":     xCampos(1, 3) = "n"

    nSQL = "SELECT mae_regimenpen.* FROM mae_regimenpen ORDER BY mae_regimenpen.codsun; "
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Régimen Pensoinario", "descripcion", "descripcion", Principio
    
    If xRs.State = 1 Then
        RstFrm.MoveFirst
        RstFrm.Find "id = " & xRs("id") & ""
    End If
    Set xRs = Nothing
End Sub

Sub MuestraSegundoTab()
    On Error GoTo error
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros para Mostrar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    Blanquea
    txt(0).Text = NulosN(RstFrm("id"))
    txt(1).Text = NulosC(RstFrm("descripcion"))
    txt(2).Text = NulosC(RstFrm("codsun"))
    If NulosN(RstFrm.Fields("cuspp")) = 0 Then
        opt(0).Value = True
    Else
        opt(1).Value = True
    End If
    '------------------
    If NulosN(RstFrm("ctaresumen")) = -1 Then
        OptCta(1).Value = True
        If NulosN(RstFrm("idcuenta")) <> 0 Then
            txt_cb(0).Text = Busca_Codigo(NulosN(RstFrm("idcuenta")), "id", "cuenta", "con_planctas", "N", xCon)
            lbl_cb(0).Caption = Busca_Codigo(NulosN(RstFrm("idcuenta")), "id", "descripcion", "con_planctas", "N", xCon)
            lbl_cod(0).Caption = NulosN(RstFrm("idcuenta"))
        End If
    Else
        OptCta(0).Value = True
    End If
    '------------------
    pCargarDatosDet
    Exit Sub
error:
    SHOW_ERROR Me.Name, "MuestraSegundoTab"
End Sub

Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL As String
    
    nSQL = "SELECT mae_regimenpen.*, con_planctas.id AS ctaid, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom " _
        + vbCr + " FROM mae_regimenpen LEFT JOIN con_planctas ON mae_regimenpen.idcuenta = con_planctas.id " _
        + vbCr + " ORDER BY mae_regimenpen.codsun; "

    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg1.DataSource = RstFrm
    TabOne1.CurrTab = 0
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub


Private Function fValidarDatos() As Boolean
    Dim band As Integer
    band = Validar(txt)
    If band <> -1 Then
       MsgBox "Llene el Campo de " & lbl(band).Caption, vbInformation, xTitulo
       txt(band).SetFocus
       Exit Function
    End If
    
    '--
    fValidarDatos = True
    
End Function

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub pCargarDatosDet()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    nSQL = "SELECT pla_concepto.id, pla_concepto.codsun, pla_conceptocat.descripcion AS catnombre, pla_conceptotipo.descripcion AS tipnombre, pla_concepto.descripcion, pla_concepto.variable, pla_conceptoregpen.idcuenta, con_planctas.cuenta, con_planctas.descripcion AS nomcuenta " _
        + vbCr + " FROM con_planctas RIGHT JOIN ((pla_conceptocat RIGHT JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) RIGHT JOIN (pla_conceptoregpen LEFT JOIN pla_concepto ON pla_conceptoregpen.idcpto = pla_concepto.id) ON pla_conceptotipo.id = pla_concepto.idtipo) ON con_planctas.id = pla_conceptoregpen.idcuenta " _
        + vbCr + " WHERE (((pla_conceptoregpen.idregpen) = " & RstFrm.Fields("id") & ")) " _
        + vbCr + " ORDER BY pla_conceptocat.descripcion DESC , pla_conceptotipo.descripcion, pla_concepto.descripcion;"

    RST_Busq RstTmp, nSQL, xCon
    Fg1.Rows = 1
    Agregando = True
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosN(RstTmp("id"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstTmp("codsun"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(RstTmp("catnombre"))
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(RstTmp("tipnombre"))
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(RstTmp("descripcion"))
        
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(RstTmp("cuenta"))
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(RstTmp("nomcuenta"))
        Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(RstTmp("idcuenta"))
        
        RstTmp.MoveNext
    Loop
    Agregando = False
    Set RstTmp = Nothing
    Exit Sub
error:
    Agregando = False
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "pCargarDatosDet"
End Sub

'**************************

Private Sub pConfigurarGrilla()
    With Fg1 '--
        .Rows = 1
        .Cols = 9
        .FixedRows = 1
        .RowHeight(0) = 250
        
        .TextMatrix(0, 1) = "ID":            .ColWidth(1) = 0:
        .TextMatrix(0, 2) = "CodSunat":      .ColWidth(2) = 0:      .ColAlignment(2) = flexAlignCenterCenter:  .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Categoría":     .ColWidth(3) = 0:      .ColAlignment(3) = flexAlignLeftCenter:    .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 4) = "Tipo":          .ColWidth(4) = 0:      .ColAlignment(4) = flexAlignLeftCenter:    .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 5) = "Descripción":   .ColWidth(5) = 5000:   .ColAlignment(5) = flexAlignLeftCenter:    .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftCenter
        '--para la cuenta contable
        .TextMatrix(0, 6) = "Cuenta":        .ColWidth(6) = 0:      .ColAlignment(6) = flexAlignLeftCenter:    .Row = 0: .Col = 6: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 7) = "Nombre Cuenta": .ColWidth(7) = 0:      .ColAlignment(7) = flexAlignLeftCenter:    .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 8) = "IdCta":         .ColWidth(8) = 0:
        
        .SelectionMode = flexSelectionByRow
        
        GRID_COMBOLIST Fg1, 6
        
    End With
    DoEvents
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Or Fg1.Row < 1 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    
    If Fg1.Col = 6 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub


Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    If Col <> 6 Then Exit Sub
    If Fg1.TextMatrix(Row, Col) = "" Then
        Fg1.TextMatrix(Row, 7) = ""
        Fg1.TextMatrix(Row, 8) = ""
        Exit Sub
    End If
    
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT * FROM con_planctas WHERE cuenta = '" & NulosC(Fg1.TextMatrix(Row, 6)) & "'", xCon
    If Rst.State = 1 Then
        If Rst.RecordCount = 1 Then
            Fg1.TextMatrix(Row, 7) = NulosC(Rst("descripcion"))
            Fg1.TextMatrix(Row, 8) = NulosN(Rst("id"))
        Else
            Fg1.TextMatrix(Row, 6) = ""
            Fg1.TextMatrix(Row, 7) = ""
            Fg1.TextMatrix(Row, 8) = ""
        End If
    End If
    Set Rst = Nothing
    
End Sub


Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If QueHace = 3 Then Exit Sub
  
    If Col <> 6 Then Exit Sub

    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
      
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "cuenta":       xCampos(0, 2) = "2000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "5000":  xCampos(1, 3) = "C"
 
    nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id " _
        + vbCr + " From con_planctas ORDER BY con_planctas.cuenta "
    
    CARGAR_DLL_EPSBUSCAR xCon, Rst, nSQL, xCampos(), "Buscando Cuentas Contables", "cuenta", "cuenta", Principio
    
    If Rst.State = 0 Then GoTo salir
    If Rst.RecordCount = 0 Then GoTo salir
       
    Agregando = True

    Fg1.TextMatrix(Fg1.Row, 6) = NulosC(Rst("cuenta"))
    Fg1.TextMatrix(Fg1.Row, 7) = NulosC(Rst("descripcion"))
    Fg1.TextMatrix(Fg1.Row, 8) = NulosN(Rst("id"))

salir:
    Set Rst = Nothing
    
    Agregando = False
    Exit Sub
error:
    Set Rst = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "Fg1_CellButtonClick"
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then pRegistroAdd False
    If KeyCode = 46 Then pRegistroDel
End Sub

Private Sub pRegistroAdd(Optional fSeleccionVarios As Boolean = True)
    On Error GoTo error
    Dim xCampos(4, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String '--almacenara los codigos de documentos ya seleccionados
    Dim nSQLNotInDocumentos As String
    Dim nSQL As String
    Dim nTitulo As String
    xCampos(0, 0) = "CodSun":       xCampos(0, 1) = "codsun":       xCampos(0, 2) = "800":   xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
    If fSeleccionVarios = True Then
        xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "5000":  xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
        xCampos(2, 0) = "Tipo":         xCampos(2, 1) = "tipnombre":    xCampos(2, 2) = "2800":  xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
        xCampos(3, 0) = "Categoría":    xCampos(3, 1) = "catnombre":    xCampos(3, 2) = "1700":  xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
    Else
        xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "4200":  xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
        xCampos(2, 0) = "Tipo":         xCampos(2, 1) = "tipnombre":    xCampos(2, 2) = "2000":  xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
        xCampos(3, 0) = "Categoría":    xCampos(3, 1) = "catnombre":    xCampos(3, 2) = "1200":  xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
    End If
    '*************************************************************
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 1, "pla_concepto.id", " NOT IN ")
    If nSQLId <> "" Then nSQLId = " WHERE " & nSQLId
    '*************************************************************
    nSQL = "SELECT pla_concepto.id, pla_conceptocat.descripcion AS catnombre, pla_conceptotipo.descripcion AS tipnombre, pla_concepto.codsun, pla_concepto.descripcion, pla_concepto.variable, pla_concepto.formula " _
        + vbCr + " FROM (pla_conceptocat INNER JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
        + nSQLId _
        + vbCr + " ORDER BY pla_conceptocat.descripcion DESC,pla_concepto.id ; "

    nTitulo = "Buscando Conceptos"
    '*************************************************************
    If fSeleccionVarios = True Then
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), nTitulo
    Else
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "descripcion", "descripcion", CualquierParte
    End If
    If xRs.State = 0 Then GoTo salir
    If xRs.RecordCount = 0 Then GoTo salir
    If fSeleccionVarios = True Then xRs.MoveFirst
    Agregando = True
    Do While Not xRs.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("id"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRs("codsun"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(xRs("catnombre"))
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(xRs("tipnombre"))
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(xRs("descripcion"))
        '-----------
        If fSeleccionVarios = False Then Exit Do
        xRs.MoveNext
    Loop
    Agregando = False
    Fg1.Row = Fg1.Rows - 1: Fg1.Col = 5:  Fg1.SetFocus
salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "Fg1_CellButtonClick"
End Sub

Private Sub pRegistroDel()
    If Fg1.Rows = 1 Then Exit Sub
    If Fg1.Row < 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar el Registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    Fg1.RemoveItem Fg1.Row
End Sub

Private Sub CmdDet_Click(Index As Integer)
    Select Case Index
        Case 0 '--agregar
            pRegistroAdd False
        Case 1 '--eliminar
            pRegistroDel
        Case 2 '--seleccionar
            pRegistroAdd True
    End Select
End Sub


'*******************************************************************************************

Private Sub cb_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--cta contable
        
            ReDim xCampos(2, 4) As String
    
            xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "cuenta":             xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":        xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
            
            nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id " _
            + vbCr + " From con_planctas " _
            + vbCr + " ORDER BY con_planctas.cuenta "

    End Select

    Dim xRs As New ADODB.Recordset
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "cuenta", "cuenta", Principio
    
    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO

    CmdDet(0).SetFocus
salir:
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
        lbl_cb(Index).Tag = ""
    End If

End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
       
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then Exit Sub
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--cuenta contable
            nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id " _
                + vbCr + " FROM con_planctas " _
                + vbCr + " WHERE con_planctas.cuenta= '" & NulosC(txt_cb(Index).Text) & "' ;"

    End Select

    If xCon.State = 0 Then GoTo salir
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cod(Index).Caption = RstTmp.Fields(2) & "" '--CODIGO
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    
    '--------------
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(Index) + ")"
    Exit Sub
salir:
    Set RstTmp = Nothing
    txt_cb(Index).Text = ""
End Sub


'****************************************************************************************

