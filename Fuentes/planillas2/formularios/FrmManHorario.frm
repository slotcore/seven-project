VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManHorario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Horario"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11685
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7065
      Top             =   -15
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
            Picture         =   "FrmManHorario.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManHorario.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManHorario.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManHorario.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManHorario.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManHorario.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManHorario.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManHorario.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManHorario.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManHorario.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManHorario.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManHorario.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6630
      Left            =   15
      TabIndex        =   9
      Top             =   360
      Width           =   11670
      _cx             =   20585
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6210
         Left            =   45
         TabIndex        =   10
         Top             =   375
         Width           =   11580
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   5550
            Left            =   15
            TabIndex        =   18
            Top             =   360
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   9790
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
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Tolerancia"
            Columns(2).DataField=   "tolerancia"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Estado"
            Columns(3).DataField=   "estado"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1005"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=926"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=8255"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8176"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1984"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1905"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
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
            HeadLines       =   1
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HDBFDFD&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
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
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Horario"
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
            Left            =   15
            TabIndex        =   11
            Top             =   30
            Width           =   11550
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6210
         Left            =   12315
         TabIndex        =   12
         Top             =   375
         Width           =   11580
         Begin VB.CommandButton cb 
            Height          =   240
            Index           =   0
            Left            =   2520
            Picture         =   "FrmManHorario.frx":277E
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1095
            Width           =   255
         End
         Begin MSComCtl2.DTPicker dtpk 
            Height          =   300
            Index           =   0
            Left            =   1215
            TabIndex        =   1
            Top             =   1065
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   529
            _Version        =   393216
            Format          =   57344002
            CurrentDate     =   39534
         End
         Begin VB.Frame Frame4 
            Caption         =   "[ Personal ]"
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
            Height          =   4530
            Left            =   75
            TabIndex        =   23
            Top             =   1575
            Width           =   6615
            Begin VB.Frame fra 
               Height          =   675
               Index           =   0
               Left            =   120
               TabIndex        =   24
               Top             =   3690
               Width           =   6315
               Begin VB.CommandButton cmd 
                  Caption         =   "Seleccionar"
                  Enabled         =   0   'False
                  Height          =   345
                  Index           =   1
                  Left            =   1305
                  TabIndex        =   3
                  ToolTipText     =   "Seleccionar Personal"
                  Top             =   210
                  Width           =   1200
               End
               Begin VB.CommandButton cmd 
                  Caption         =   "Agregar"
                  Enabled         =   0   'False
                  Height          =   345
                  Index           =   0
                  Left            =   60
                  TabIndex        =   2
                  ToolTipText     =   "Agregar Personal"
                  Top             =   210
                  Width           =   1200
               End
               Begin VB.CommandButton cmd 
                  Caption         =   "Eliminar"
                  Enabled         =   0   'False
                  Height          =   345
                  Index           =   2
                  Left            =   2805
                  TabIndex        =   4
                  ToolTipText     =   "Eliminar Personal"
                  Top             =   210
                  Width           =   1200
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   3400
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Top             =   300
               Width           =   6315
               _cx             =   11139
               _cy             =   5997
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
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManHorario.frx":28B0
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
         End
         Begin VB.Frame Frame3 
            Caption         =   "[ Tipo de Hora ]"
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
            Height          =   4530
            Left            =   6765
            TabIndex        =   21
            Top             =   1575
            Width           =   4815
            Begin VB.Frame fra 
               Height          =   675
               Index           =   1
               Left            =   150
               TabIndex        =   22
               Top             =   3645
               Width           =   4530
               Begin VB.CommandButton cmd 
                  Caption         =   "Agregar"
                  Enabled         =   0   'False
                  Height          =   345
                  Index           =   3
                  Left            =   60
                  TabIndex        =   6
                  ToolTipText     =   "Agregar Tipo de Hora"
                  Top             =   180
                  Width           =   1200
               End
               Begin VB.CommandButton cmd 
                  Caption         =   "Eliminar"
                  Enabled         =   0   'False
                  Height          =   345
                  Index           =   4
                  Left            =   1605
                  TabIndex        =   7
                  ToolTipText     =   "Eliminar Tipo de Hora"
                  Top             =   180
                  Width           =   1200
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   3400
               Index           =   1
               Left            =   150
               TabIndex        =   8
               Top             =   300
               Width           =   4530
               _cx             =   7990
               _cy             =   5997
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
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManHorario.frx":2934
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
               Ellipsis        =   1
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
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   1215
            MaxLength       =   100
            TabIndex        =   0
            Text            =   "txt(1)"
            Top             =   720
            Width           =   5895
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0000FFFF&
            Height          =   315
            Index           =   0
            Left            =   10350
            TabIndex        =   15
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   60
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Tolerancia"
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   19
            Top             =   1170
            Width           =   750
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   16
            Top             =   795
            Width           =   555
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "codigo"
            Height          =   195
            Index           =   0
            Left            =   9750
            TabIndex        =   14
            Top             =   150
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Horario"
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
            Height          =   210
            Left            =   15
            TabIndex        =   13
            Top             =   30
            Width           =   11550
         End
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu Menu3 
      Caption         =   "Menu3"
      Visible         =   0   'False
      Begin VB.Menu Menu3_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu3_2 
         Caption         =   "Seleccionar"
      End
      Begin VB.Menu Menu3_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu3_4 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu Menu3_5 
         Caption         =   "Eliminar Todo"
      End
   End
End
Attribute VB_Name = "FrmManHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim RstFrm As New ADODB.Recordset
Dim Agregando As Boolean
Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Private Sub pRegistroAdd(Index As Integer, Optional fSeleccionVarios As Boolean = False)
    '--
    If QueHace = 3 Then Exit Sub
    If txt(1).Text = "" Then
        MsgBox "Ingrese la Descripción del Horario", vbExclamation, xTitulo
        txt(1).SetFocus
        Exit Sub
    End If
    
    Dim nSQLId As String
    Dim nSQL As String
    Dim nTitulo As String
    
    Agregando = True
    Select Case Index
        Case 0 '--
            
            ReDim xCampos(3, 5) As String
            
            xCampos(0, 0) = "Personal":         xCampos(0, 1) = "nombres":   xCampos(0, 2) = "6100":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
            xCampos(1, 0) = "Fch. Ingreso":     xCampos(1, 1) = "fching1":    xCampos(1, 2) = "1400":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
            xCampos(2, 0) = "Id":               xCampos(2, 1) = "id":        xCampos(2, 2) = "500":      xCampos(2, 3) = "c":    xCampos(2, 4) = "N"
            '---------
            nSQLId = GRID_GENERAR_SQL_ID(fg(0), 1, "pla_empleados.id", " NOT IN ", True)
            If nSQLId <> "" Then nSQLId = " AND " & nSQLId
            '---------
            nTitulo = "Buscando Personal"
            nSQL = "SELECT 0 AS xsel, pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres,format(pla_empleados.fching,'dd/mm/yyyy') as fching1 " _
                + vbCr + " FROM pla_empleados " _
                + vbCr + " WHERE (((pla_empleados.id) Not In (select  idemp  from  mae_horarioemp where vigencia = -1)) AND ((pla_empleados.fchcese) Is Null)); "
            
        Case 1
        
            ReDim xCampos(1, 5) As String
            
            xCampos(0, 0) = "Tipo de Hora":     xCampos(0, 1) = "nombres":   xCampos(0, 2) = "6500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
            '---------
            nSQLId = GRID_GENERAR_SQL_ID(fg(1), 1, "mae_tipohora.id", " NOT IN ", True)
            If nSQLId <> "" Then nSQLId = " AND " & nSQLId
            '---------
            nSQL = "SELECT mae_tipohora.id, mae_tipohora.descripcion as nombres FROM mae_tipohora " _
                + vbCr + " WHERE (((mae_tipohora.horario)=-1)) " & nSQLId & " ORDER BY mae_tipohora.prioridad;"
            
    End Select
    
    '-------------------------------
    On Error GoTo error
    Dim xRs  As New ADODB.Recordset
            
    If fSeleccionVarios = True Then
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), nTitulo
    Else
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombres", "nombres", Principio
    End If
    
    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir
    '---
    If fSeleccionVarios = True Then xRs.MoveFirst
    Do While Not xRs.EOF
        With fg(Index)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = NulosC(xRs.Fields("id"))
            .TextMatrix(.Rows - 1, 2) = NulosC(xRs.Fields("nombres"))
            If Index = 0 Then
                .TextMatrix(.Rows - 1, 3) = NulosC(xRs.Fields("sexo"))
            End If
            '---
        End With
        If fSeleccionVarios = False Then Exit Do
        xRs.MoveNext
    Loop
    fg(Index).Row = fg(Index).Rows - 1
    fg(Index).Col = 2
salir:
    Agregando = False
    Set xRs = Nothing
    '----
    fg(Index).SetFocus
    Exit Sub
error:
    Agregando = False
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "pRegistroAdd"
End Sub

Private Sub pRegistroDel(Index As Integer)
    If QueHace = 3 Then Exit Sub
    If fg(Index).Row <= 0 Then
        MsgBox "La fila no se puede eliminar" + vbCr + "Seleccione una fila correcta", vbExclamation, xTitulo
        Exit Sub
    End If
    If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    
    fg(Index).RemoveItem (fg(Index).Row)
    
End Sub


Private Sub cb_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    Dim obj As New SGI2_funciones.formularios
    obj.HoraSeleccionar dtpk(Index), -1, -1, dtpk(Index).Value
    Set obj = Nothing
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--ADD PERSONAL
            pRegistroAdd 0
        Case 1 '--SEL REG
            pRegistroAdd 0, True
        Case 2 '--DEL
            pRegistroDel 0
        '--DE LOS TIPOS DE HORAS
        Case 3 '--ADD
            pRegistroAdd 1
        Case 4 '--DEL
            pRegistroDel 1
    End Select
End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
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

Private Sub dtpk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If dtpk(Index).Enabled = False Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
    ElseIf KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub fg_EnterCell(Index As Integer)
    If QueHace = 3 Or Index = 0 Then
        fg(Index).Editable = flexEDNone
        Exit Sub
    End If
    If fg(Index).Col < 3 Then
        fg(Index).Editable = flexEDNone
    Else
        fg(Index).Editable = flexEDKbdMouse
    End If
End Sub
Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Or Index = 2 Or Index = 3 Then
        KeyAscii = 0
        Exit Sub
    End If
'    Select Case Col
'        Case 3, 7
'            If validar_numero(KeyAscii) = False Then KeyAscii = 0
'        Case 4
'        Case Else
'            KeyAscii = 0
'    End Select
End Sub

Private Sub fg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then 'F3 = Agregar Item
        pRegistroAdd Index
    End If
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then
        pRegistroDel Index  'F4 = Eliminar Item
    End If
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If QueHace = 3 Then Exit Sub
    If Agregando = True Then Exit Sub
    If Index = 0 Then Exit Sub
    
    Select Case Col
        Case 3
            If IsDate(fg(Index).TextMatrix(Row, 3)) = True Then
                If IsDate(fg(Index).TextMatrix(Row, 4)) = True Then
                    If CDate(fg(Index).TextMatrix(Row, 4)) < CDate(fg(Index).TextMatrix(Row, 3)) Then
                        MsgBox "La Hora Inicial es Superior a la Hora Final", vbExclamation, xTitulo
                        fg(Index).TextMatrix(Row, Col) = ""
                    End If
                End If
            Else
                fg(Index).TextMatrix(Row, Col) = ""
            End If
        Case 4
            If IsDate(fg(Index).TextMatrix(Row, 4)) = True Then
                If IsDate(fg(Index).TextMatrix(Row, 3)) = True Then
                    If CDate(fg(Index).TextMatrix(Row, 4)) < CDate(fg(Index).TextMatrix(Row, 3)) Then
                        MsgBox "La Hora Final es Inferior a la Hora Inicial", vbExclamation, xTitulo
                        fg(Index).TextMatrix(Row, Col) = ""
                    End If
                End If
            Else
                fg(Index).TextMatrix(Row, Col) = ""
            End If
        End Select
    Exit Sub
error:
    SHOW_ERROR Me.Name, "fg_CellChanged (" + CStr(Index) + ")"
End Sub

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Index = 0 Then Exit Sub
    If Col = 3 Or Col = 4 Then
        '--invocar al formulario de horas
        Dim obj As New SGI2_funciones.formularios
        obj.HoraSeleccionar fg(1), Row, Col, fg(1).TextMatrix(Row, Col)
        Set obj = Nothing
    End If
    Exit Sub
salir:
    Agregando = False
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
            Select Case Index
            Case 0: PopupMenu Menu3
            Case 1: PopupMenu menu1
            End Select
        End If
    End If
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

Private Sub pCargarGrid()
    Dim nSQL  As String
    nSQL = "SELECT mae_horario.*, IIf([mae_horario].[vigencia]=-1,'Vigente','De Baja') AS estado " _
        + vbCr + " FROM mae_horario " _
        + vbCr + " ORDER BY mae_horario.descripcion;"

    '--cargando datos
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg3.DataSource = RstFrm
    Me.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    Agregando = False
    QueHace = 3
        
    Dg3.BatchUpdates = False
    
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    '--
    Habilitar_Obj False
    '----
    Dg3.Columns("tolerancia").NumberFormat = FORMAT_HORA_LARGO

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set Dg3.DataSource = Nothing
End Sub


Private Sub Menu1_1_Click()
    cmd_Click 3
End Sub

Private Sub Menu1_3_Click()
    cmd_Click 4
End Sub

Private Sub Menu3_1_Click()
    cmd_Click 0
End Sub

Private Sub Menu3_2_Click()
    cmd_Click 1
End Sub

Private Sub Menu3_4_Click()
    cmd_Click 2
End Sub

Private Sub Menu3_5_Click()
    Dim mRow&
    If fg(0).Rows <= 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar todos los registros", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    Agregando = True
    Do While fg(0).Rows > 1
        fg(0).Row = 1
        cmd_Click 2
    Loop
    Agregando = False

End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstFrm.Requery
            Dg3.Refresh
        End If
    End If
    If Button.Index = 6 Then Cancelar
    If Button.Index = 8 Then Filtrar
    If Button.Index = 9 Then RstFrm.Filter = ""
    If Button.Index = 11 Then Buscar
    If Button.Index = 13 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Sub Eliminar()

    On Error GoTo error
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    

    If MsgBox("¿Esta seguro de eliminar el registro?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
        xCon.Execute "DELETe * FROM mae_horarioemp WHERE idhor = " & RstFrm("id") & ""
        xCon.Execute "DELETe * FROM mae_horariohora WHERE idhor = " & RstFrm("id") & ""
        xCon.Execute "DELETe * FROM mae_horario WHERE id = " & RstFrm("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstFrm("id") & " AND idform = " & IdMenuActivo
        
        
        MsgBox "El registro fue eliminado con éxito", vbInformation, xTitulo
        RstFrm.Requery
        Dg3.Refresh
        TabOne1.CurrTab = 0
        If RstFrm.RecordCount = 0 Then
            If MsgBox("No hay registrado ningún Estado, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
                nuevo
            End If
        End If
    End If
    
Exit Sub
error:
    SHOW_ERROR Me.Name, "Eliminar"
End Sub

Sub Cancelar()
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    ActivaTool
    Habilitar_Obj False
    Label1.Caption = "Detalle del Estado"
    TabOne1.CurrTab = 0
    fg(1).ColFormat(3) = FORMAT_HORA_AL_SEGUNDO
    fg(1).ColFormat(4) = FORMAT_HORA_AL_SEGUNDO
    Dg3.SetFocus
End Sub

Private Sub Modificar()
   '------
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If
    
    QueHace = 2
    xHorIni = Time
    ActivaTool
    Habilitar_Obj True
    MuestraSegundoTab
    
    fg(1).Editable = flexEDKbdMouse
    fg(1).SelectionMode = flexSelectionFree
    
    Label1.Caption = "Modificando Horario"
    
    fg(1).ColFormat(3) = FORMAT_HORA_LARGO
    fg(1).ColFormat(4) = FORMAT_HORA_LARGO
    
    txt(1).SetFocus
    
End Sub

Sub MuestraSegundoTab()
'    On Error GoTo error
    With RstFrm
        Blanquea
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Then Exit Sub
        
        txt(0).Text = NulosN(.Fields("id")) '--CODIGO
        txt(1).Text = NulosC(.Fields("descripcion"))
        If IsDate(.Fields("tolerancia")) = True Then
            dtpk(0).Value = CDate(.Fields("tolerancia"))
        End If
        '---
       
        MuestraDetalle
        
    End With
    
    Exit Sub
error:
    
    SHOW_ERROR
End Sub

Private Sub MuestraDetalle()
    Dim xRs As New ADODB.Recordset
    Dim A&, xCol&, xFil&, xFila&
    Dim nSQL As String
    
    On Error GoTo error
    '--del personal
    nSQL = "SELECT  pla_empleados.id,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_sexo.abrev AS sexo " _
        + vbCr + " FROM mae_sexo RIGHT JOIN (pla_empleados INNER JOIN mae_horarioemp ON pla_empleados.id = mae_horarioemp.idemp) ON mae_sexo.id = pla_empleados.idsex " _
        + vbCr + " WHERE (((mae_horarioemp.idhor)=" & NulosN(RstFrm.Fields("id")) & " ) AND ((mae_horarioemp.vigencia)=-1));"
       
    RST_Busq xRs, nSQL, xCon
    If xRs.RecordCount <> 0 Then
        Agregando = True
        If xRs.RecordCount <> 0 Then
            fg(0).Rows = 1
            xRs.MoveFirst
            Do While Not xRs.EOF
                With fg(0)
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = NulosC(xRs.Fields("id"))
                    .TextMatrix(.Rows - 1, 2) = NulosC(xRs.Fields("nombres"))
                    .TextMatrix(.Rows - 1, 3) = NulosC(xRs.Fields("sexo"))
                    xRs.MoveNext
                End With
            Loop
        End If
    End If
    '--del tipo de hora
    nSQL = "SELECT mae_tipohora.id, mae_tipohora.descripcion, mae_horariohora.hingreso, mae_horariohora.hsalida " _
        + vbCr + " FROM mae_tipohora INNER JOIN mae_horariohora ON mae_tipohora.id = mae_horariohora.idhora " _
        + vbCr + " WHERE ((mae_horariohora.idhor) = " & NulosN(RstFrm.Fields("id")) & ") " _
        + vbCr + " ORDER BY mae_tipohora.prioridad;"
        
    RST_Busq xRs, nSQL, xCon
    If xRs.RecordCount <> 0 Then
        Agregando = True
        If xRs.RecordCount <> 0 Then
            fg(1).Rows = 1
            xRs.MoveFirst
            Do While Not xRs.EOF
                With fg(1)
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = NulosN(xRs.Fields("id"))
                    .TextMatrix(.Rows - 1, 2) = NulosC(xRs.Fields("descripcion"))
                    .TextMatrix(.Rows - 1, 3) = NulosC(xRs.Fields("hingreso"))
                    .TextMatrix(.Rows - 1, 4) = NulosC(xRs.Fields("hsalida"))
                    xRs.MoveNext
                End With
            Loop
        End If
    End If
    '----------------------------------------------------
    Set xRs = Nothing
    Agregando = False
    Exit Sub
error:
    Set xRs = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "MuestraDetalle"
End Sub


Private Sub Habilitar_Obj(band As Boolean)
    habilitar_Locked txt, Not band
    habilitar cmd, band
    habilitar dtpk, band
    
    TabOne1.CurrTab = IIf(band = False, 0, 1)
    TabOne1.TabEnabled(0) = Not band
    
    If band = False Then
        fg(0).SelectionMode = flexSelectionByRow
        fg(1).SelectionMode = flexSelectionByRow
    End If
    
End Sub

Private Sub Blanquea()
    LimpiaText txt
    
    fg(0).Rows = 1
    fg(1).Rows = 1
    
End Sub

Private Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub nuevo()
    QueHace = 1
    xHorIni = Time
    ActivaTool
    Blanquea
    Habilitar_Obj True
    Label1.Caption = "Agregando Horario"
    '------------
    
    fg(1).Editable = flexEDKbdMouse
    fg(1).SelectionMode = flexSelectionFree
    '------------
    txt(1).SetFocus
    fg(1).ColFormat(3) = FORMAT_HORA_LARGO
    fg(1).ColFormat(4) = FORMAT_HORA_LARGO

    dtpk(0).Value = CDate("12:00:00 AM")

End Sub


Function Grabar() As Boolean
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el Registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo salir
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstHora As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim xId As Double
    Dim xCol&, xFil&
    Dim nSQL As String
    
    On Error GoTo LaCague
    Me.MousePointer = vbHourglass
    xCon.BeginTrans
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM mae_horario ", xCon
        xId = HallaCodigoTabla("mae_horario", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstFrm("id")
        
        RST_Busq RstCab, "SELECT * FROM mae_horario WHERE id =" & xId & "", xCon
        '--Eliminando los tipos de horas asignados al horario
        xCon.Execute "DELETE * FROM mae_horariohora WHERE idhor = " & xId & ""
        '--DESACTIVANDO LAS PERSONAS ASOCIADAS AL HORARIO
        xCon.Execute "UPDATE mae_horarioemp SET mae_horarioemp.vigencia = 0 WHERE (((mae_horarioemp.idhor)=" & xId & "));"
        
    End If
    
    RST_Busq RstDet, "SELECT top 1 * FROM mae_horarioemp", xCon
    RST_Busq RstHora, "SELECT top 1 * FROM mae_horariohora", xCon
    
    RstCab("descripcion") = Trim(txt(1).Text) & ""
    If IsDate(dtpk(0).Value) = True Then
        RstCab("tolerancia") = CDate(dtpk(0).Value)
    Else
        RstCab("tolerancia") = Null
    End If
    RstCab.Update
    '--del personal
    With fg(0)
        For xFil = 1 To .Rows - 1
            nSQL = "SELECT mae_horarioemp.idemp FROM mae_horarioemp " _
                + vbCr + " WHERE (((mae_horarioemp.idhor)=" & xId & ") AND ((mae_horarioemp.idemp)=" & NulosN(.TextMatrix(xFil, 1)) & "));"
            RST_Busq RstTmp, nSQL, xCon
            If RstTmp.RecordCount <> 0 Then
                xCon.Execute "UPDATE mae_horarioemp SET vigencia=-1 WHERE idhor=" & xId & " AND idemp=" & NulosN(.TextMatrix(xFil, 1)) & " ;"
            Else
                RstDet.AddNew
                RstDet("idhor") = xId
                RstDet("idemp") = NulosN(.TextMatrix(xFil, 1))
                RstDet("vigencia") = -1
                RstDet.Update
            End If
        Next xFil
    End With
    '--del tipo de hora
    With fg(1)
        For xFil = 1 To .Rows - 1
            RstHora.AddNew
            RstHora("idhor") = xId
            RstHora("idhora") = NulosN(.TextMatrix(xFil, 1))
            RstHora("hingreso") = CDate(.TextMatrix(xFil, 3))
            RstHora("hsalida") = CDate(.TextMatrix(xFil, 4))
            RstHora.Update
        Next xFil
    End With
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    
    MsgBox "El registro se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    Grabar = True
salir:
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstHora = Nothing:    Set RstTmp = Nothing
    Me.MousePointer = vbDefault
    Exit Function
LaCague:
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstHora = Nothing:    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function


Private Function fValidarDatos() As Boolean
    Dim mRow&, QGrid&
    
    If Trim(txt(1).Text) = "" Then
        MsgBox "Ingrese la descripción del Horario", vbExclamation, xTitulo
        Exit Function
    End If
    
    '--------------------------------
    '--VALIDAR QUE NO ESTE REGISTRADO
    Dim RstTmp As New ADODB.Recordset
    If QueHace = 1 Then
        RST_Busq RstTmp, "SELECT descripcion FROM mae_horario WHERE ucase(descripcion)='" + UCase(Trim(txt(1).Text)) + "';", xCon
    Else
        RST_Busq RstTmp, "SELECT descripcion FROM mae_horario WHERE ucase(descripcion)='" + UCase(Trim(txt(1).Text)) + "' AND id <> " + CStr(RstFrm.Fields("id")) + ";", xCon
    End If
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then
        MsgBox "El registro " + IIf(QueHace = 1, " ya fue ingresado", "ya existe"), vbExclamation, xTitulo
        Set RstTmp = Nothing
        Exit Function
    End If
    Set RstTmp = Nothing
    '--------------------------------
    If fg(0).Rows = 1 Then
        MsgBox "Ingrese al menos un Personal para este Horario", vbExclamation, xTitulo
        cmd(0).SetFocus
        Exit Function
    End If
    If fg(1).Rows = 1 Then
        MsgBox "Ingrese al menos un Tipo de Hora para este Horario", vbExclamation, xTitulo
        cmd(3).SetFocus
        Exit Function
    End If
    '--------------------------------
    With fg(1)
        For mRow = 1 To .Rows - 1
            If IsDate(.TextMatrix(mRow, 3)) = False Then
                MsgBox "Ingrese la Hora de Inicio" + vbCr + "Tipo de Hora: " + .TextMatrix(mRow, 2), vbExclamation, xTitulo
                Agregando = True:  .Row = mRow:     .Col = 3: Agregando = False
                fg(1).SetFocus
                Exit Function
            ElseIf IsDate(.TextMatrix(mRow, 4)) = False Then
                MsgBox "Ingrese la Hora Final" + vbCr + "Tipo de Hora: " + .TextMatrix(mRow, 2), vbExclamation, xTitulo
                Agregando = True:  .Row = mRow:     .Col = 4: Agregando = False
                fg(1).SetFocus
                Exit Function
            End If
        Next mRow
    End With
    '--------------------------------
    fValidarDatos = True
End Function
 


Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
   
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripción":        xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "Tolerancia":         xCampos(1, 1) = "tolerancia":       xCampos(1, 2) = "1500":     xCampos(1, 3) = "F"
    
        
    nSQL = "SELECT mae_horario.*, IIf([mae_horario].[vigencia]=-1,'Vigente','De Baja') AS estado " _
        + vbCr + " FROM mae_horario " _
        + vbCr + " ORDER BY mae_horario.descripcion;"
        
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Horario", "descripcion", "descripcion", Principio
    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(xRs("id"))
salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

Private Sub Filtrar()
    
    Dim xCampos(2, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    xCampos(0, 0) = "Descripción":        xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "C":     xCampos(0, 3) = "5000"
    xCampos(1, 0) = "Tolerancia":               xCampos(1, 1) = "tolerancia":       xCampos(0, 2) = "F":     xCampos(1, 3) = "1500"
    
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg3

    TabOne1.CurrTab = 0
End Sub

Private Function fGenerarConsulta(X_ID As String) As String
    
    Dim nSQL As String

    nSQL = "SELECT mae_horarioemp.orden, mae_horariohora.idest, mae_horariohora.idcuenta, con_planctas.cuenta, con_planctas.descripcion " _
        + vbCr + " FROM con_planctas INNER JOIN (mae_horarioemp INNER JOIN mae_horariohora ON (mae_horarioemp.id = mae_horariohora.idest) AND (mae_horarioemp.idcab = mae_horariohora.idcab)) ON con_planctas.id = mae_horariohora.idcuenta " _
        + vbCr + " WHERE (((mae_horariohora.idcab)=" + X_ID + "));"

    fGenerarConsulta = nSQL
    
End Function


Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
End Sub


Private Sub pConfigurarGrilla()
    With fg(0) '--del personal
        OCULTAR_COL fg(0), 1, 1 '--idemp
    
        .Rows = 1
        .Cols = 4
        .FixedRows = 1
        .RowHeight(0) = 300
        .ColWidth(1) = 0:
        
        .TextMatrix(0, 1) = "Idemp":                .ColWidth(1) = 0:
        
        .TextMatrix(0, 2) = "Apellidos y Nombres":  .ColWidth(2) = 5000:    .ColAlignment(2) = flexAlignLeftCenter:     .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 3) = "Sexo":                 .ColWidth(3) = 550:    .ColAlignment(3) = flexAlignCenterCenter:   .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterCenter
        
        .SelectionMode = flexSelectionByRow
    End With
    With fg(1) '--tipos de horas
        .Rows = 1
        .Cols = 5
        .ColWidth(1) = 200
        .FixedRows = 1
        .RowHeight(0) = 300
        .ColWidth(1) = 0:
        .TextMatrix(0, 1) = "Idemp":                .ColWidth(1) = 0:
        
        .TextMatrix(0, 2) = "Descripción":  .ColWidth(2) = 1500:    .ColAlignment(2) = flexAlignLeftCenter:  .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 3) = "Ingreso":  .ColWidth(3) = 1350:    .ColAlignment(3) = flexAlignLeftCenter:  .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 4) = "Salida":     .ColWidth(4) = 1350:    .ColAlignment(4) = flexAlignLeftCenter:  .Row = 0: .Col = 4: .CellAlignment = flexAlignCenterCenter
        
        .ColFormat(3) = FORMAT_HORA_AL_SEGUNDO
        .ColFormat(4) = FORMAT_HORA_AL_SEGUNDO
        .ColEditMask(3) = "##:##:##"
        .ColEditMask(4) = "##:##:##"
        .SelectionMode = flexSelectionByRow
    End With
    GRID_COMBOLIST fg(1), 3
    GRID_COMBOLIST fg(1), 4
    '*****************************************
    DoEvents
End Sub

