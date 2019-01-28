VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmManConceptoVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Otros Conceptos"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7155
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
            Picture         =   "FrmManConceptoVarios.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoVarios.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoVarios.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoVarios.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoVarios.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoVarios.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoVarios.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoVarios.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoVarios.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoVarios.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoVarios.frx":1EA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
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
      Height          =   6750
      Left            =   30
      TabIndex        =   12
      Top             =   360
      Width           =   7125
      _cx             =   12568
      _cy             =   11906
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
         Height          =   6330
         Left            =   45
         TabIndex        =   15
         Top             =   375
         Width           =   7035
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   5895
            Left            =   45
            TabIndex        =   16
            Top             =   360
            Width           =   6930
            _ExtentX        =   12224
            _ExtentY        =   10398
            _LayoutType     =   4
            _RowHeight      =   14
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Descripción"
            Columns(0).DataField=   "descripcion"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Variable"
            Columns(1).DataField=   "variable"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Valor"
            Columns(2).DataField=   "formula"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=5345"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5265"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=4419"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4339"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=1482"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1402"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=2"
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
            _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(46)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(56)  =   "Named:id=33:Normal"
            _StyleDefs(57)  =   ":id=33,.parent=0"
            _StyleDefs(58)  =   "Named:id=34:Heading"
            _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(60)  =   ":id=34,.wraptext=-1"
            _StyleDefs(61)  =   "Named:id=35:Footing"
            _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(63)  =   "Named:id=36:Selected"
            _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=37:Caption"
            _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(67)  =   "Named:id=38:HighlightRow"
            _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=39:EvenRow"
            _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(71)  =   "Named:id=40:OddRow"
            _StyleDefs(72)  =   ":id=40,.parent=33"
            _StyleDefs(73)  =   "Named:id=41:RecordSelector"
            _StyleDefs(74)  =   ":id=41,.parent=34"
            _StyleDefs(75)  =   "Named:id=42:FilterBar"
            _StyleDefs(76)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Otros Conceptos"
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
            TabIndex        =   17
            Top             =   60
            Width           =   7830
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6330
         Left            =   7770
         TabIndex        =   13
         Top             =   375
         Width           =   7035
         Begin VB.TextBox txt 
            BackColor       =   &H0000FFFF&
            Height          =   285
            Index           =   3
            Left            =   6015
            TabIndex        =   24
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   15
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Frame Frame4 
            Height          =   2010
            Left            =   45
            TabIndex        =   20
            Top             =   285
            Width           =   6930
            Begin VB.TextBox txt 
               BackColor       =   &H8000000F&
               Height          =   315
               Index           =   2
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   1
               Text            =   "txt(2)"
               Top             =   585
               Width           =   3645
            End
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   0
               Left            =   1200
               Picture         =   "FrmManConceptoVarios.frx":23EC
               Style           =   1  'Graphical
               TabIndex        =   26
               ToolTipText     =   "Seleccione el Tipo de Concepto (Primero seleccione la categoría)"
               Top             =   1650
               Width           =   210
            End
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Index           =   1
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   0
               Text            =   "txt(1)"
               Top             =   240
               Width           =   5430
            End
            Begin VB.Frame FraValor 
               Caption         =   "¿El Valor varía en función al Periodo?"
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
               Left            =   90
               TabIndex        =   21
               Top             =   960
               Width           =   5490
               Begin VB.TextBox txt_Valor 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   4170
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   4
                  Tag             =   "null"
                  Text            =   "txt_Valor"
                  Top             =   195
                  Width           =   1155
               End
               Begin VB.OptionButton opt 
                  Caption         =   "No"
                  Height          =   285
                  Index           =   0
                  Left            =   615
                  TabIndex        =   3
                  Top             =   255
                  Value           =   -1  'True
                  Width           =   585
               End
               Begin VB.OptionButton opt 
                  Caption         =   "Si"
                  Height          =   285
                  Index           =   1
                  Left            =   1875
                  TabIndex        =   22
                  Top             =   255
                  Width           =   555
               End
               Begin VB.Label lbl_Valor 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valor"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   3690
                  TabIndex        =   31
                  ToolTipText     =   "Es nombre abreviado del concepto, con posibilidad de que aparesca en el reporte"
                  Top             =   300
                  Width           =   360
               End
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   0
               Left            =   735
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   5
               Text            =   "txt_cb(0)"
               Top             =   1620
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Variable"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   30
               Top             =   705
               Width           =   570
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
               Left            =   3945
               TabIndex        =   28
               Top             =   1635
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Entidad"
               Height          =   195
               Index           =   0
               Left            =   105
               TabIndex        =   27
               Top             =   1725
               Width           =   540
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   23
               Top             =   360
               Width           =   840
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
               Left            =   1485
               TabIndex        =   29
               Top             =   1620
               Width           =   4095
            End
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0000FFFF&
            Height          =   285
            Index           =   0
            Left            =   10665
            TabIndex        =   18
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   -15
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Frame Frame3 
            Height          =   4065
            Left            =   45
            TabIndex        =   32
            Top             =   2235
            Width           =   6930
            Begin VB.Frame fra 
               BorderStyle     =   0  'None
               Caption         =   "Frame12"
               Height          =   480
               Index           =   1
               Left            =   135
               TabIndex        =   34
               Top             =   150
               Width           =   6570
               Begin MSComCtl2.UpDown UpDown1 
                  Height          =   300
                  Left            =   1321
                  TabIndex        =   35
                  Top             =   90
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   529
                  _Version        =   393216
                  Value           =   2999
                  AutoBuddy       =   -1  'True
                  BuddyControl    =   "txt_ano"
                  BuddyDispid     =   196626
                  OrigLeft        =   1290
                  OrigTop         =   90
                  OrigRight       =   1545
                  OrigBottom      =   390
                  Max             =   2999
                  Min             =   2000
                  Wrap            =   -1  'True
                  Enabled         =   0   'False
               End
               Begin MSDataListLib.DataCombo dtcb_periodo 
                  Height          =   315
                  Left            =   2595
                  TabIndex        =   7
                  Top             =   75
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Style           =   2
                  Text            =   "dtcb_periodo"
               End
               Begin VB.TextBox txt_ano 
                  Height          =   300
                  Left            =   540
                  Locked          =   -1  'True
                  TabIndex        =   6
                  Text            =   "txt_ano"
                  Top             =   90
                  Width           =   780
               End
               Begin VB.Line lin 
                  BorderColor     =   &H00808080&
                  BorderWidth     =   2
                  Index           =   1
                  X1              =   -30
                  X2              =   13000
                  Y1              =   465
                  Y2              =   465
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00808080&
                  BorderWidth     =   2
                  Index           =   0
                  X1              =   6555
                  X2              =   6555
                  Y1              =   15
                  Y2              =   1000
               End
               Begin VB.Line Line2 
                  BorderColor     =   &H00FFFFFF&
                  BorderWidth     =   2
                  X1              =   15
                  X2              =   15
                  Y1              =   0
                  Y2              =   1000
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Periodo"
                  Height          =   195
                  Index           =   0
                  Left            =   1950
                  TabIndex        =   37
                  Top             =   195
                  Width           =   540
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Año"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   36
                  Top             =   195
                  Width           =   285
               End
               Begin VB.Line lin 
                  BorderColor     =   &H80000009&
                  BorderWidth     =   2
                  Index           =   0
                  X1              =   -15
                  X2              =   13000
                  Y1              =   15
                  Y2              =   15
               End
            End
            Begin VB.Frame fra 
               BorderStyle     =   0  'None
               Caption         =   "Frame12"
               Height          =   570
               Index           =   0
               Left            =   135
               TabIndex        =   33
               Top             =   3405
               Width           =   6570
               Begin VB.CommandButton CmdDetRestablece 
                  Caption         =   "Duplicar de Último Periodo"
                  Enabled         =   0   'False
                  Height          =   435
                  Left            =   4440
                  TabIndex        =   11
                  Top             =   60
                  Width           =   2040
               End
               Begin VB.CommandButton CmdDet 
                  Caption         =   "&Eliminar"
                  Enabled         =   0   'False
                  Height          =   435
                  Index           =   1
                  Left            =   1545
                  TabIndex        =   10
                  Top             =   60
                  Width           =   1395
               End
               Begin VB.CommandButton CmdDet 
                  Caption         =   "&Agregar"
                  Enabled         =   0   'False
                  Height          =   435
                  Index           =   0
                  Left            =   120
                  TabIndex        =   8
                  Top             =   60
                  Width           =   1395
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
               Begin VB.Line lin 
                  BorderColor     =   &H00808080&
                  BorderWidth     =   2
                  Index           =   3
                  X1              =   15
                  X2              =   13045
                  Y1              =   555
                  Y2              =   555
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00808080&
                  BorderWidth     =   2
                  Index           =   3
                  X1              =   6555
                  X2              =   6555
                  Y1              =   -30
                  Y2              =   955
               End
               Begin VB.Line Line5 
                  BorderColor     =   &H00FFFFFF&
                  BorderWidth     =   2
                  X1              =   15
                  X2              =   15
                  Y1              =   0
                  Y2              =   1000
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   2685
               Left            =   135
               TabIndex        =   9
               Top             =   675
               Width           =   6570
               _cx             =   11589
               _cy             =   4736
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
               FormatString    =   $"FrmManConceptoVarios.frx":251E
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   3
            Left            =   5490
            TabIndex        =   25
            Top             =   105
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
            TabIndex        =   19
            Top             =   75
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Otros Conceptos"
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
            Left            =   30
            TabIndex        =   14
            Top             =   60
            Width           =   6885
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
Attribute VB_Name = "FrmManConceptoVarios"
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


Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If Col = 2 Then KeyAscii = 0
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If QueHace <> 2 Then PopupMenu Menu1
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
   
    SeEjecuto = False
    pConfigurarGrilla
    pCargarGrid
        
    '--agregando los meses
    Dim RsCons As New ADODB.Recordset
    RST_Busq RsCons, "SELECT id, descripcion From con_meses WHERE (((con_meses.id) Not In (0,13))) ORDER BY id", xCon
    Set dtcb_periodo.RowSource = RsCons
    dtcb_periodo.ListField = "descripcion"
    dtcb_periodo.BoundColumn = "id"
    Set RsCons = Nothing
    
    '--asignado los valores por defecto
'    dtcb_periodo.BoundText = 1
    txt_ano.Text = AnoTra
    dtcb_periodo.BoundText = xMes
    '--------
    SeEjecuto = True
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado Otros Conceptos, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            nuevo
        End If
    End If
    
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    QueHace = 3

    TabOne1.CurrTab = 0

End Sub

Sub Blanquea()
    LimpiaText txt
    LimpiaText txt_cb
    LimpiaText lbl_cb
    LimpiaText lbl_cod
    txt_Valor.Text = ""
    Fg1.Rows = Fg1.FixedRows
End Sub

Sub Bloquea(band As Boolean)

    habilitar_Locked txt, Not band
    habilitar CmdDet, band
    FraValor.Enabled = band
    habilitar_Locked txt, Not band
    habilitar_Locked txt_cb, Not band

    dtcb_periodo.Enabled = band
    txt_ano.Enabled = band
    Me.UpDown1.Enabled = band


    If (QueHace = 1) Or (QueHace = 2 And NulosC(txt(2).Text) = "") Then
        txt(2).Enabled = True
        txt(2).BackColor = vbWhite
    Else
        Dim RstTmp As New ADODB.Recordset
        Dim nSQL As String
        
        nSQL = "SELECT pla_concepto.id FROM pla_concepto WHERE ucase(pla_concepto.formula) Like '%" & UCase(NulosC(RstFrm.Fields("variable"))) & "%' ;"
        RST_Busq RstTmp, nSQL, xCon
        If RstTmp.RecordCount <> 0 Then
            txt(2).Enabled = False
            txt(2).BackColor = &H8000000F
        Else
            txt(2).Enabled = True
            txt(2).BackColor = vbWhite
        End If
        Set RstTmp = Nothing
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set RstFrm = Nothing
End Sub

Private Sub lbl_cb_Change(Index As Integer)
    If Index <> 0 Then Exit Sub
    If NulosC(lbl_cb(0).Caption) = "" Then
        Fg1.TextMatrix(0, 2) = "Descripción"
    Else
        Fg1.TextMatrix(0, 2) = "Descripción - " & lbl_cb(0).Caption:
    End If
End Sub


Private Sub opt_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    If opt(0).Value = True Then '--es fijo
        Fg1.Rows = Fg1.FixedRows '--limpiar grilla
        habilitar CmdDet, False
        dtcb_periodo.Enabled = False
        txt_ano.Enabled = False
        Me.UpDown1.Enabled = False
        txt_Valor.Locked = False
    Else
        habilitar CmdDet, True
        dtcb_periodo.Enabled = True
        txt_ano.Enabled = True
        Me.UpDown1.Enabled = True
        txt_Valor.Text = ""
        txt_Valor.Locked = True
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
   
    Dim xId&, A&

    On Error GoTo LaCague

    xCon.BeginTrans

    If QueHace = 1 Then

        xId = HallaCodigoTabla("pla_conceptovarios", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pla_conceptovarios", xCon

        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstFrm("id")
        
        RST_Busq RstCab, "SELECT * FROM pla_conceptovarios WHERE id = " & xId & "", xCon
        '--eliminando los registros del detalle de Otros Conceptos
        xCon.Execute "Delete from pla_conceptovariosdet where idcptov =  " & xId & " and anno= " & NulosN(txt_ano.Text) & " and idmes = " & dtcb_periodo.BoundText & " ;"
    End If

    RST_Busq RstDet, "SELECT TOP 1 * FROM pla_conceptovariosdet ; ", xCon

    RstCab("descripcion") = NulosC(txt(1).Text)
    RstCab("variable") = NulosC(txt(2).Text)
    '--sivaria en fucion al periodo
    If opt(0).Value = True Then
        RstCab("esfijo") = -1
    Else
        RstCab("esfijo") = 0
    End If
    
    RstCab("formula") = NulosN(txt_Valor.Text)
    RstCab("entgen") = NulosN(lbl_cod(0).Caption)
    
    RstCab.Update
    
    '--detalle otros conceptos
    If opt(1).Value = True Then '--si elige la opcion si
        For A = 1 To Fg1.Rows - 1
            If NulosN(Fg1.TextMatrix(A, 1)) <> 0 Then
                RstDet.AddNew
                RstDet("idcptov") = xId
                RstDet("idref") = NulosN(Fg1.TextMatrix(A, 1))
                RstDet("anno") = NulosN(txt_ano.Text)
                RstDet("idmes") = NulosN(dtcb_periodo.BoundText)
                RstDet("imptot") = NulosN(Fg1.TextMatrix(A, 3))
                RstDet.Update
            End If
        Next A
    End If
    xCon.CommitTrans
    
    MsgBox "Los datos se " + IIf(QueHace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo

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

Sub nuevo()
    QueHace = 1
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea True
    Blanquea
    
    If opt(0).Value = False Then
        opt(0).Value = True
    Else
        opt_Click 0
    End If
    
    Label5.Caption = "Agregando Otros Conceptos"
    txt(1).SetFocus
End Sub

Sub Modificar()
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    Label5.Caption = "Modificando Otros Conceptos"
    
    ActivaTool
    
    Bloquea True
    
    If TabOne1.CurrTab = 0 Then TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    QueHace = 2
    Fg1.SelectionMode = flexSelectionFree
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
    If Button.Index = 1 Then nuevo

    If Button.Index = 2 Then Modificar

    If Button.Index = 3 Then Eliminar

    If Button.Index = 5 Then Cancelar

    If Button.Index = 6 Then
        If Grabar = True Then
            RstFrm.Requery
            Dg1.Refresh
            Cancelar
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
    xCampos(0, 0) = "Concepto":   xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "5200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Variable":   xCampos(1, 1) = "variable":           xCampos(1, 2) = "2800":     xCampos(1, 3) = "n"

    nSQL = "SELECT pla_conceptovarios.* FROM pla_conceptovarios ORDER BY pla_conceptovarios.descripcion; "
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Otros Conceptos", "descripcion", "descripcion", Principio
    
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
    txt(2).Text = NulosC(RstFrm("variable"))
    
    If NulosN(RstFrm.Fields("esfijo")) = -1 Then '--es fijo
        opt(0).Value = True
    Else '--es variable
        opt(1).Value = True
    End If
    txt_Valor.Text = NulosN(RstFrm("formula"))
    '------------------
    If NulosN(RstFrm("entgen")) <> -1 Then
        txt_cb(0).Text = NulosN(RstFrm("entgen"))
        lbl_cb(0).Caption = Busca_Codigo(NulosN(RstFrm("entgen")), "id", "descripcion", "pla_entidades", "N", xCon)
        lbl_cod(0).Caption = NulosN(RstFrm("entgen"))
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
    
    nSQL = "SELECT pla_conceptovarios.* FROM pla_conceptovarios ORDER BY pla_conceptovarios.descripcion; "

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


Public Sub pCargarDatosDet()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    If SeEjecuto = False Then Exit Sub

    If NulosN(txt_ano.Text) = 0 Or NulosN(dtcb_periodo.BoundText) = 0 Then
        MsgBox "Vuelva a seleccionar El Año o Periodo", vbExclamation, xTitulo
        Exit Sub
    End If
    nSQL = "SELECT pla_conceptovariosdet.idref, IIf(pla_conceptovarios.entgen=1,mae_regimenpen.descripcion,mae_cargo.descripcion) AS descripcion, pla_conceptovariosdet.imptot " _
        + vbCr + " FROM pla_conceptovarios INNER JOIN ((pla_conceptovariosdet LEFT JOIN mae_regimenpen ON pla_conceptovariosdet.idref = mae_regimenpen.id) LEFT JOIN mae_cargo ON pla_conceptovariosdet.idref = mae_cargo.id) ON pla_conceptovarios.id = pla_conceptovariosdet.idcptov " _
        + vbCr + " WHERE (((pla_conceptovariosdet.anno) = " & NulosN(txt_ano.Text) & ") And ((pla_conceptovariosdet.idmes) = " & NulosN(dtcb_periodo.BoundText) & ") And ((pla_conceptovarios.id) = " & RstFrm("id") & ")) " _
        + vbCr + " ORDER BY IIf(pla_conceptovarios.entgen=1,mae_regimenpen.descripcion,mae_cargo.descripcion), pla_conceptovariosdet.imptot;"
    
    RST_Busq RstTmp, nSQL, xCon
    Fg1.Rows = 1
    Agregando = True
    If RstTmp.RecordCount <> 0 Then
        RstTmp.MoveFirst
        CmdDetRestablece.Enabled = False
    Else
        CmdDetRestablece.Enabled = True
    End If
    Do While Not RstTmp.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosN(RstTmp("idref"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstTmp("descripcion"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosN(RstTmp("imptot"))
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
        .Cols = 4
        .FixedRows = 1
        .RowHeight(0) = 350
        .TextMatrix(0, 1) = "ID":           .ColWidth(1) = 0:
        .TextMatrix(0, 2) = "Descripción":  .ColWidth(2) = 4500: .ColAlignment(2) = flexAlignLeftCenter:   .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 3) = "  Valor ":        .ColWidth(3) = 1100: .ColAlignment(3) = flexAlignRightCenter:  .Row = 0: .Col = 3: .CellAlignment = flexAlignRightCenter
        .ColFormat(3) = "#.#####"
        
        .ColEditMask(3) = "#.#####"
        
        .SelectionMode = flexSelectionByRow
        GRID_COMBOLIST Fg1, 2
    End With
    
    '*****************************************
    DoEvents
End Sub

Private Sub Fg1_EnterCell()
     If QueHace = 3 Or Fg1.Row < 1 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    If Fg1.Col >= 2 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    Select Case Col
        Case 3
            If NulosN(Fg1.TextMatrix(Row, Col)) = 0 Then Fg1.TextMatrix(Row, Col) = 0
            Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), "0.00000")
    End Select
    Exit Sub
error:
    SHOW_ERROR Me.Name, "fg1_CellChanged"
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Row < 1 Then Exit Sub
    If Col <> 2 Then Exit Sub
    
    If NulosN(txt_cb(0).Text) = 0 Then
        MsgBox "Falta seleccionar la Entidad", vbExclamation, xTitulo
        txt_cb(0).SetFocus
        Exit Sub
    End If
    
    On Error GoTo error
    Dim xCampos(2, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String '--almacenara los codigos de documentos ya seleccionados
    Dim nSQL As String

    xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":  xCampos(0, 2) = "4500":  xCampos(0, 3) = "C":     xCampos(0, 4) = "S"
    xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":      xCampos(1, 2) = "700":   xCampos(1, 3) = "N":     xCampos(1, 4) = "N"
        
    '*************************************************************
    Select Case NulosN(txt_cb(0).Text)
        Case 1 '--regimen pensionario
            nSQLId = GRID_GENERAR_SQL_ID(Fg1, 1, "mae_regimenpen.id", " NOT IN ")
            nSQL = "SELECT mae_regimenpen.id, mae_regimenpen.descripcion AS nombre, mae_regimenpen.id AS cod FROM mae_regimenpen " _
                + vbCr + " WHERE mae_regimenpen.cuspp=-1 " _
                + vbCr + IIf(nSQLId = "", "", " AND " + nSQLId)
        Case 2 '--cargo del personal
            nSQLId = GRID_GENERAR_SQL_ID(Fg1, 1, "mae_cargo.id", " NOT IN ")
            
            nSQL = "SELECT mae_cargo.id, mae_cargo.descripcion AS nombre, mae_cargo.id AS cod " _
                + vbCr + " FROM mae_cargo " _
                + vbCr + IIf(nSQLId = "", "", "WHERE " + nSQLId)
        Case Else
        
            Exit Sub
    End Select
    '*************************************************************
    '*************************************************************
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando" & lbl_cb(0).Caption, "nombre", "nombre", Principio
    If xRs.State = 0 Then GoTo salir
    If xRs.RecordCount = 0 Then GoTo salir
    Agregando = True
    Do While Not xRs.EOF
        Fg1.TextMatrix(Row, 1) = NulosN(xRs("id"))
        Fg1.TextMatrix(Row, 2) = NulosC(xRs("nombre"))
        xRs.MoveNext
    Loop
    Agregando = False
    Fg1.Row = Row: Fg1.Col = 2:  Fg1.SetFocus
salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "Fg1_CellButtonClick"
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then pRegistroAdd
    If KeyCode = 46 Then pRegistroDel
End Sub

'*******************************************************************************************

Private Sub cb_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--entidad
        
            ReDim xCampos(2, 4) As String
    
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "600":    xCampos(1, 3) = "N"
            
            nSQL = "SELECT pla_entidades.id, pla_entidades.descripcion AS nombre, pla_entidades.id AS cod " _
                + vbCr + " FROM pla_entidades;"

    End Select

    Dim xRs As New ADODB.Recordset
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
    
    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO

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


Private Sub CmdDet_Click(Index As Integer)
    Select Case Index
        Case 0 '--agregar
            If NulosN(txt_cb(0).Text) = 0 Then Exit Sub
            pRegistroAdd
        Case 1 '--eliminar
            pRegistroDel
    End Select
End Sub

Private Sub txt_ano_Change()
    If SeEjecuto = False Then Exit Sub
    If NulosN(txt_ano.Text) < 2005 Then Exit Sub
    pCargarDatosDet
End Sub

Private Sub UpDown1_Change()
    txt_ano.Text = UpDown1.Value
End Sub

Private Sub dtcb_periodo_Change()
    If SeEjecuto = False Then Exit Sub
    If dtcb_periodo.MatchedWithList = False Then Exit Sub
    pCargarDatosDet
End Sub

Private Sub CmdDetRestablece_Click()
    Dim nPeriodo As String
    Dim nSQL As String
    Dim mMes&, mAnno&
    If NulosN(txt_ano.Text) = 0 Or dtcb_periodo.MatchedWithList = False Then Exit Sub
    
    mMes = dtcb_periodo.BoundText
    mAnno = NulosN(txt_ano.Text)
    
    nPeriodo = "Año: " & IIf(mMes = 1, mAnno - 1, mAnno) & "   Mes: " & Busca_Codigo(IIf(mMes = 1, 12, mMes - 1), "id", "descripcion", "con_meses", "N", xCon)

'    '--eliminando registros si esta en bd
    
    Dim RstTmp As New ADODB.Recordset
    nSQL = "SELECT pla_conceptovariosdet.idref, mae_regimenpen.descripcion, IIf(idmes=12,anno+1,anno) AS ano, IIf(idmes=12,1,idmes+1) AS mes, pla_conceptovariosdet.imptot " _
        + vbCr + " FROM mae_regimenpen INNER JOIN pla_conceptovariosdet ON mae_regimenpen.id = pla_conceptovariosdet.idref WHERE idcptov =" & RstFrm("id") & " and anno=" & IIf(mMes = 1, mAnno - 1, mAnno) & " AND idmes= " & IIf(mMes = 1, 12, mMes - 1) & ""
    
    RST_Busq RstTmp, nSQL, xCon
    
    Fg1.Rows = 1
    
    If RstTmp.RecordCount <> 0 Then
        RstTmp.MoveFirst
        CmdDetRestablece.Enabled = False
    Else
        CmdDetRestablece.Enabled = True
    End If
    Agregando = True
    Do While Not RstTmp.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosN(RstTmp("idref"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstTmp("descripcion"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosN(RstTmp("imptot"))
        RstTmp.MoveNext
    Loop
    
    Set RstTmp = Nothing
    Agregando = False
    If Fg1.Rows = 1 Then
        MsgBox "No hay información en el " & vbCr & nPeriodo, vbInformation, xTitulo
    Else
        MsgBox "La Información que se muestra es Duplicado de..." + vbCr + nPeriodo, vbInformation, xTitulo
    End If
End Sub




Private Sub pRegistroAdd()
    Dim mCol%
    If QueHace = 3 Then Exit Sub
    Agregando = True
    If Fg1.Rows > 1 Then
        If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 1)) = 0 Then
            MsgBox "Falta Completar...", vbExclamation, xTitulo
        Else
            Fg1.AddItem ""
        End If
    Else
        Fg1.AddItem ""
    End If
    Fg1.Row = Fg1.Rows - 1
    Fg1.Col = 2
    Fg1.SetFocus
    Agregando = False
End Sub

Private Sub pRegistroDel()
    If Fg1.Rows = 1 Then Exit Sub
    If Fg1.Row < 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar el Registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    Fg1.RemoveItem Fg1.Row
    If Fg1.Rows = 1 Then
        CmdDet(0).SetFocus
    Else
        Fg1.Row = Fg1.Rows - 1
        Fg1.Col = 2
        Fg1.SetFocus
    End If
End Sub
