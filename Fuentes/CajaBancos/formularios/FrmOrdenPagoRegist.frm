VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOrdenPagoRegist 
   Caption         =   "Caja Y Bancos - Orden de Pago"
   ClientHeight    =   7620
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   11940
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
            Picture         =   "FrmOrdenPagoRegist.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoRegist.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoRegist.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoRegist.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoRegist.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoRegist.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoRegist.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoRegist.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoRegist.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoRegist.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoRegist.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoRegist.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   15
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
         NumButtons      =   16
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
            Object.Visible         =   0   'False
            ImageIndex      =   2
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
            Object.ToolTipText     =   "Opciones de Impresion"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Comprobante de Retencion"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Registro de Retenciones"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7260
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   11940
      _cx             =   21061
      _cy             =   12806
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
         Height          =   6840
         Left            =   45
         TabIndex        =   12
         Top             =   375
         Width           =   11850
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6495
            Left            =   30
            TabIndex        =   13
            Top             =   345
            Width           =   11805
            _ExtentX        =   20823
            _ExtentY        =   11456
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Orden Pago"
            Columns(0).DataField=   "numdoc"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nº Doc."
            Columns(1).DataField=   "candoc"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Emi."
            Columns(2).DataField=   "fchemi"
            Columns(2).NumberFormat=   "dd/mm/yy"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Pago"
            Columns(3).DataField=   "fchpago"
            Columns(3).NumberFormat=   "dd/mm/yy"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Importe"
            Columns(4).DataField=   "imptot"
            Columns(4).NumberFormat=   "#,###0.00"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Programado Por"
            Columns(5).DataField=   "nomprog2"
            Columns(5).NumberFormat=   "0.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Autorizado Por"
            Columns(6).DataField=   "nomaut2"
            Columns(6).NumberFormat=   "0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Estado"
            Columns(7).DataField=   "desest"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1958"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1879"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1349"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1270"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2037"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1958"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2117"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2037"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2037"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1958"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=514"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=4260"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=4180"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=3810"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=3731"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=2196"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2117"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=0"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Programación de Orden de Pagos"
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
            TabIndex        =   14
            Top             =   30
            Width           =   11610
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Detalle de la Cuenta"
         Height          =   6840
         Left            =   12585
         TabIndex        =   9
         Top             =   375
         Width           =   11850
         Begin VB.Frame Frame3 
            Height          =   2220
            Left            =   10185
            TabIndex        =   10
            Top             =   3165
            Width           =   1590
            Begin VB.CommandButton CmdDel 
               Caption         =   "Eliminar Documento"
               Height          =   540
               Left            =   210
               Style           =   1  'Graphical
               TabIndex        =   6
               Tag             =   "b"
               Top             =   1170
               Width           =   1170
            End
            Begin VB.CommandButton CmdAdd 
               Caption         =   "Agregar Documentos"
               Height          =   540
               Left            =   210
               Style           =   1  'Graphical
               TabIndex        =   5
               Tag             =   "b"
               Top             =   600
               Width           =   1170
            End
         End
         Begin VB.TextBox TxtTotACuentaRec 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   7110
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "TxtTotACuentaRec"
            Top             =   6525
            Width           =   885
         End
         Begin VB.TextBox TxtTotSalAntRec 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   6180
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "TxtTotSalAntRec"
            Top             =   6525
            Width           =   915
         End
         Begin VB.TextBox TxtTotImpRec 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "TxtTotImpRec"
            Top             =   6525
            Width           =   1155
         End
         Begin VB.TextBox TxtTotNueSalRec 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   8010
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "TxtTotNueSalRec"
            Top             =   6525
            Width           =   945
         End
         Begin VB.TextBox TxtTotNueSal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   8010
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "TxtTotNueSal"
            Top             =   6195
            Width           =   945
         End
         Begin VB.Frame Frame6 
            Caption         =   "[ Tipo Movimiento ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   600
            Left            =   6315
            TabIndex        =   33
            Top             =   360
            Width           =   5460
            Begin VB.OptionButton OptBanco 
               Caption         =   "Banco"
               Height          =   195
               Left            =   2940
               TabIndex        =   3
               Tag             =   "b"
               Top             =   300
               Width           =   1125
            End
            Begin VB.OptionButton OptCaja 
               Caption         =   "Caja"
               Height          =   195
               Left            =   1740
               TabIndex        =   2
               Tag             =   "b"
               Top             =   300
               Value           =   -1  'True
               Width           =   1080
            End
         End
         Begin VB.Frame Frame4 
            Height          =   840
            Left            =   10200
            TabIndex        =   30
            Top             =   5325
            Width           =   1575
            Begin VB.CheckBox ChkAutorizar 
               Caption         =   "Aprobar Todos"
               Height          =   195
               Left            =   105
               TabIndex        =   31
               Top             =   375
               Width           =   1365
            End
         End
         Begin VB.Frame Frame8 
            Height          =   1080
            Left            =   6315
            TabIndex        =   25
            Top             =   870
            Width           =   5460
            Begin VB.CommandButton CmdAprobada 
               Caption         =   "Aprobar"
               Height          =   300
               Left            =   1230
               TabIndex        =   27
               Top             =   690
               Width           =   1605
            End
            Begin VB.CommandButton CmdRecha 
               Caption         =   "Rechazar"
               Height          =   300
               Left            =   2865
               TabIndex        =   26
               Top             =   690
               Width           =   1605
            End
            Begin VB.Label LblIdEstado 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               Caption         =   "LblIdEstado"
               Height          =   195
               Left            =   120
               TabIndex        =   28
               Top             =   315
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label LblEstado 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Pendiente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   465
               Left            =   1215
               TabIndex        =   29
               Top             =   195
               Width           =   3240
            End
         End
         Begin VB.TextBox TxtObs 
            Height          =   720
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Tag             =   "a"
            Text            =   "FrmOrdenPagoRegist.frx":277E
            Top             =   2235
            Width           =   11685
         End
         Begin VB.TextBox TxtTotalImp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "TxtTotalImp"
            Top             =   6195
            Width           =   1155
         End
         Begin VB.TextBox TxtTotSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   6180
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "TxtTotSaldo"
            Top             =   6195
            Width           =   915
         End
         Begin VB.TextBox TxtTotalACuenta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   7110
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "TxtTotalACuenta"
            Top             =   6195
            Width           =   885
         End
         Begin VB.Frame Frame5 
            Height          =   1590
            Left            =   90
            TabIndex        =   16
            Top             =   360
            Width           =   6180
            Begin VB.OptionButton OptDol 
               Caption         =   "Dolares"
               Height          =   195
               Left            =   2625
               TabIndex        =   47
               Tag             =   "b"
               Top             =   1290
               Width           =   1080
            End
            Begin VB.OptionButton OptSol 
               Caption         =   "Soles"
               Height          =   195
               Left            =   1380
               TabIndex        =   46
               Tag             =   "b"
               Top             =   1290
               Value           =   -1  'True
               Width           =   1080
            End
            Begin VB.TextBox TxtCodigo 
               Height          =   300
               Left            =   4575
               TabIndex        =   24
               Tag             =   "a"
               Text            =   "TxtCodigo"
               Top             =   1200
               Visible         =   0   'False
               Width           =   1155
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFecEmisOrdPag 
               Height          =   300
               Left            =   1365
               TabIndex        =   0
               Tag             =   "b"
               Top             =   885
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
               Valor           =   "31/10/2007"
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFecPago 
               Height          =   300
               Left            =   3975
               TabIndex        =   1
               Tag             =   "b"
               Top             =   885
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
               Valor           =   "31/10/2007"
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Index           =   1
               Left            =   150
               TabIndex        =   48
               Top             =   1260
               Width           =   585
            End
            Begin VB.Label LblIdProg 
               BackColor       =   &H00C0FFC0&
               Caption         =   "LblIdProg"
               Height          =   285
               Left            =   3180
               TabIndex        =   35
               Top             =   225
               Visible         =   0   'False
               Width           =   885
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Autorizado Por"
               Height          =   195
               Left            =   150
               TabIndex        =   38
               Top             =   585
               Width           =   1035
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Programado Por"
               Height          =   195
               Left            =   150
               TabIndex        =   37
               Top             =   255
               Width           =   1140
            End
            Begin VB.Label LblProg 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblProg"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1365
               TabIndex        =   36
               Top             =   240
               Width           =   4695
            End
            Begin VB.Label LblIdPersona 
               BackColor       =   &H00C0FFC0&
               Caption         =   "LblIdPersona"
               Height          =   255
               Left            =   3615
               TabIndex        =   34
               Top             =   585
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.Label LblNombre 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblNombre"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1365
               TabIndex        =   32
               Top             =   555
               Width           =   4695
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Pago"
               Height          =   195
               Index           =   0
               Left            =   2760
               TabIndex        =   18
               Top             =   915
               Width           =   1095
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Emis."
               Height          =   195
               Index           =   8
               Left            =   150
               TabIndex        =   17
               Top             =   915
               Width           =   1095
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2910
            Left            =   90
            TabIndex        =   7
            Top             =   3255
            Width           =   10020
            _cx             =   17674
            _cy             =   5133
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
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmOrdenPagoRegist.frx":2787
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
            Caption         =   "Documentos a Pagar"
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
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Top             =   3030
            Width           =   1785
         End
         Begin VB.Label LblTotRechaz 
            AutoSize        =   -1  'True
            Caption         =   "Total Rechazados"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3285
            TabIndex        =   44
            Top             =   6555
            Width           =   1500
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Observación"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   2010
            Width           =   900
         End
         Begin VB.Label LblTotAprob 
            AutoSize        =   -1  'True
            Caption         =   "Total Aprobados"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3285
            TabIndex        =   22
            Top             =   6225
            Width           =   1395
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle Orden de Pagos"
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
            Left            =   120
            TabIndex        =   11
            Top             =   60
            Width           =   11640
         End
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar_Item"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar Item"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "aprodesa"
      Visible         =   0   'False
      Begin VB.Menu menu2_1 
         Caption         =   "&Aprobar"
      End
      Begin VB.Menu menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu2_3 
         Caption         =   "&Rechazar"
      End
   End
End
Attribute VB_Name = "FrmOrdenPagoRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vFormatString As String
Dim CaracteresNumericos As String
Dim QueHace  As Integer
Dim vStrSql As String
Dim Mostrando As Boolean
'--VARIABLE PARA SABER SI ES PROGRAMADOR O AUTORIZADOR
'--0 NO ES NINGUNO, 1=PROGRAMADOR, 2=AUTORIZADOR
Dim vProgra_o_Autoriz As Integer
Dim vArrProg() As String, vArrAutor() As String
'-----------------------------------------------------
Dim vMoneda As Integer
Dim SeEjecuto As Boolean
Dim RstOrdPago As New ADODB.Recordset
'--VARIABLES PARA EL CAMBIO DE MES
Dim xFchIni As String, xFchFin As String
'----------------------------------------

'--VARIABLES PARA EXTRAER LOS COLORES DE ESTADO
Dim vArrColorEstado(1 To 4) As String
'-------------------------------------
Dim vCalcularTotales As Integer

Sub ColorEstado()
    Dim RsColor As New ADODB.Recordset
    Dim i_color As Integer
    vStrSql = "SELECT mae_estados.id, mae_estados.color" _
        & " FROM mae_estados" _
        & " ORDER BY mae_estados.id"

    RST_Busq RsColor, vStrSql, xCon
    If RsColor.RecordCount > 0 Then
        RsColor.MoveFirst
        For i_color = 1 To RsColor.RecordCount
            vArrColorEstado(i_color) = NulosC(RsColor("color"))
            RsColor.MoveNext
        Next
    End If
    Set RsColor = Nothing
End Sub

Sub Buscar()
    TabOne1.CurrTab = 0
     
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(5, 4) As String
    
    xCampos(0, 0) = "Nº Doc. Pago":     xCampos(0, 1) = "numdoc":     xCampos(0, 2) = "1400":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fec. Emis.":       xCampos(1, 1) = "fchemi":     xCampos(1, 2) = "1100":    xCampos(1, 3) = "F"
    xCampos(2, 0) = "Fec. Pago":        xCampos(2, 1) = "fchpago":    xCampos(2, 2) = "1100":    xCampos(2, 3) = "F"
    xCampos(3, 0) = "Programado por":   xCampos(3, 1) = "nomprog":    xCampos(3, 2) = "2300":    xCampos(3, 3) = "C"
    xCampos(4, 0) = "Autorizador por":  xCampos(4, 1) = "nomaut":     xCampos(4, 2) = "2300":    xCampos(4, 3) = "C"
        
    xForm.SQLCad = "SELECT con_ordenpago.numdoc, pla_empleados.nom & ' ' & pla_empleados.ape AS nomprog, pla_empleados_1.nom & ' ' & pla_empleados_1.ape AS nomaut, con_ordenpago.fchemi, con_ordenpago.fchpago" _
        & " FROM pla_empleados AS pla_empleados_1 RIGHT JOIN (con_emptes AS con_emptes_1 RIGHT JOIN (pla_empleados RIGHT JOIN (con_emptes RIGHT JOIN con_ordenpago ON con_emptes.id = con_ordenpago.idprog) ON pla_empleados.id = con_emptes.idemp) ON con_emptes_1.id = con_ordenpago.idaut) ON pla_empleados_1.id = con_emptes_1.idemp" _
        & " WHERE con_ordenpago.fchemi BETWEEN CDATE('" & xFchIni & "') AND CDATE('" & xFchFin & "')" _
        & " ORDER BY pla_empleados.nom & ' ' & pla_empleados.ape, con_ordenpago.numdoc"
                
    xForm.Titulo = "Buscando Orden de Pago"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "nomprog"
    xForm.CampoBusca = "nomprog"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        RstOrdPago.MoveFirst
        RstOrdPago.Find "id = " & Val(xRs("numdoc")) & ""
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Function fDevolRsParaCons_AlActivarVentana(pFecha As String) As ADODB.Recordset
    Dim RsDevolver As New ADODB.Recordset
    Dim xFECHA As String
    TabOne1.CurrTab = 0
'    xFecha = CDate(Date)
    xFECHA = CDate(pFecha)
        
    xFchIni = "01/" + Mid(Format(CDate(xFECHA), "dd/mm/yy"), 4, 5)
    xFchFin = Trim(Format(HallaDiasMes(CDate(xFECHA)), "00")) + "/" + Mid(Format(CDate(xFECHA), "dd/mm/yy"), 4, 5)
        
    vStrSql = "SELECT DISTINCT con_ordenpago.id, con_ordenpago.numdoc, con_ordenpago.candoc, con_ordenpago.fchemi, con_ordenpago.fchpago, con_ordenpago.imptot, (SELECT pla_empleados.ape & ', ' & pla_empleados.nom AS nomprog FROM pla_empleados RIGHT JOIN con_emptes ON pla_empleados.id = con_emptes.idemp WHERE (((con_emptes.prog)=-1)) AND  con_ordenpago.idprog = con_emptes.id) AS nomprog2, (SELECT pla_empleados.ape & ', ' & pla_empleados.nom AS nomaut FROM pla_empleados RIGHT JOIN con_emptes ON pla_empleados.id = con_emptes.idemp WHERE (((con_emptes.aut)=-1)) AND con_ordenpago.idaut = con_emptes.id) AS nomaut2," _
            & " mae_estados.descripcion AS desest, mae_moneda.simbolo, con_ordenpago.tipmov, con_ordenpago.obs, con_ordenpago.idest, con_ordenpago.idprog, con_ordenpago.idaut, mae_moneda.id AS idmon, mae_estados.color" _
            & " FROM pla_empleados RIGHT JOIN (mae_moneda RIGHT JOIN (mae_estados RIGHT JOIN ((con_emptes RIGHT JOIN con_ordenpago ON con_emptes.id = con_ordenpago.idprog) LEFT JOIN con_ordenpagodet ON con_ordenpago.id = con_ordenpagodet.idord) ON mae_estados.id = con_ordenpago.idest) ON mae_moneda.id = con_ordenpago.idmon) ON pla_empleados.id = con_emptes.idemp" _
            & " WHERE con_ordenpago.fchemi BETWEEN CDATE('" & xFchIni & "') AND CDATE('" & xFchFin & "')" _
            & " ORDER BY con_ordenpago.id, con_ordenpago.fchemi"
    
    RST_Busq RsDevolver, vStrSql, xCon
    Set fDevolRsParaCons_AlActivarVentana = RsDevolver
    
    Set RsDevolver = Nothing
End Function

Sub CambiarMes()
    Dim i_cammes As Integer
    i_cammes = SeleccionaMes(xCon)
    If i_cammes <> 0 And i_cammes <> 13 Then
        Set RstOrdPago = fDevolRsParaCons_AlActivarVentana("01/" + Trim(Format(i_cammes, "00")) + "/" + Trim(Format(Year(Date), "0000")))
        Set Dg1.DataSource = RstOrdPago
        'CmdSalir_Click
    End If
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    
    Rpta = MsgBox("¿Esta seguro de eliminar la orden de pago seleccionada?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        TabOne1.CurrTab = 0
        xCon.Execute "DELETE * FROM con_ordenpago WHERE id = " & NulosN(RstOrdPago("id")) & ""
        xCon.Execute "DELETE FROM con_ordenpagodet WHERE idord = " & NulosN(RstOrdPago("id")) & ""
        xCon.Execute "DELETE FROM con_ordpagodetrechaz WHERE idord = " & NulosN(RstOrdPago("id")) & ""
        
        MsgBox "El orden de pago se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstOrdPago.Requery
        Dg1.Refresh
        Dg1.SetFocus
    End If
End Sub

Private Function fVerifSiLosAutorSonIguales() As Boolean
    If (NulosN(RstOrdPago("idaut")) = Val(LblIdPersona.Caption)) Or NulosN(RstOrdPago("idaut")) = 0 Then
        fVerifSiLosAutorSonIguales = True
    Else
        fVerifSiLosAutorSonIguales = False
    End If
End Function

Private Sub LlenarGridDetalle(pRs As ADODB.Recordset)
    pRs.MoveFirst
    Do While Not pRs.EOF
        If AddItemGrid(Fg1) = True Then
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(pRs("nombre"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(pRs("numdoccompra"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(pRs("fchemi"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(pRs("simbolo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(pRs("imptot")), "#,###0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosC(pRs("saldo")), "#,###0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosC(pRs("acuenta")), "#,###0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosC(pRs("salrest")), "#,###0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(pRs("autorizado"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosC(pRs("idcomp"))
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosC(pRs("idprov"))
        End If
        pRs.MoveNext
    Loop
    Set pRs = Nothing
End Sub

Private Sub HabDeshabControles_de_Autor()
    If Dg1.ApproxCount > 0 Then
        If NulosN(RstOrdPago("idest")) = 3 Then 'PROCESADO
            CmdAprobada.Enabled = False
            CmdRecha.Enabled = False
            ChkAutorizar.Enabled = False
            Fg1.ColWidth(9) = 0
            Exit Sub
        End If
        
        If vProgra_o_Autoriz = 1 Or vProgra_o_Autoriz = 0 Then
            CmdAprobada.Enabled = False
            CmdRecha.Enabled = False
            ChkAutorizar.Enabled = False
            If RstOrdPago("idaut") <> 0 Then
                Fg1.ColWidth(9) = 1000
            Else
                Fg1.ColWidth(9) = 0
            End If
        ElseIf Dg1.ApproxCount <= 0 Then
            CmdAprobada.Enabled = False
            CmdRecha.Enabled = False
            ChkAutorizar.Enabled = False
            Fg1.ColWidth(9) = 0
        Else
            CmdAprobada.Enabled = True
            CmdRecha.Enabled = True
            ChkAutorizar.Enabled = True
            If Val(vArrAutor(1)) = 1 Then
                Fg1.Editable = flexEDKbdMouse
            End If
            Fg1.ColWidth(9) = 1065
        End If
        '---------
        If vProgra_o_Autoriz = 2 And RstOrdPago("idest") = 1 Then 'PENDIENTE
            ChkAutorizar.Enabled = True
            Fg1.ColWidth(9) = 1065
'            HabDeshabControles_de_Autor
        ElseIf vProgra_o_Autoriz = 2 And RstOrdPago("idest") = 2 Then 'APROBADO
            ChkAutorizar.Enabled = False
            CmdAprobada.Enabled = False
            'Fg1.ColWidth(9) = 0
        ElseIf vProgra_o_Autoriz = 2 And RstOrdPago("idest") = 3 Then 'PROCESADO
            CmdAprobada.Enabled = False
            CmdRecha.Enabled = False
            ChkAutorizar.Enabled = False
            Fg1.ColWidth(9) = 1065
        ElseIf vProgra_o_Autoriz = 2 And RstOrdPago("idest") = 4 Then 'RECHAZADO
            CmdAprobada.Enabled = True
            CmdRecha.Enabled = False
            ChkAutorizar.Enabled = True
            Fg1.ColWidth(9) = 1065
        End If
        '---------
    Else
        If vProgra_o_Autoriz = 1 Or vProgra_o_Autoriz = 0 Then
            CmdAprobada.Enabled = False
            CmdRecha.Enabled = False
            ChkAutorizar.Enabled = False
        ElseIf Dg1.ApproxCount <= 0 Then
            CmdAprobada.Enabled = False
            CmdRecha.Enabled = False
            ChkAutorizar.Enabled = False
        Else
            CmdAprobada.Enabled = True
            CmdRecha.Enabled = True
            ChkAutorizar.Enabled = True
        End If
        
        If vProgra_o_Autoriz = 2 Then 'PENDIENTE
            ChkAutorizar.Enabled = True
        ElseIf vProgra_o_Autoriz = 2 Then 'APROBADO
            ChkAutorizar.Enabled = False
            CmdAprobada.Enabled = False
        ElseIf vProgra_o_Autoriz = 2 Then 'PROCESADO
            CmdAprobada.Enabled = False
            CmdAprobada.Enabled = False
            ChkAutorizar.Enabled = False
        ElseIf vProgra_o_Autoriz = 2 Then 'RECHAZADO
            CmdAprobada.Enabled = True
            CmdRecha.Enabled = False
            ChkAutorizar.Enabled = True
        End If
        
        If vProgra_o_Autoriz = 3 Then
            If QueHace = 1 Or QueHace = 2 Then
                CmdAprobada.Enabled = False
                CmdRecha.Enabled = False
                ChkAutorizar.Enabled = False
            End If
        End If
    End If
    
    If vProgra_o_Autoriz = 3 Then 'CUANDO ES PROGRAMA Y AUTORIZA
        If QueHace = 1 Or QueHace = 2 Then
            CmdAprobada.Enabled = False
            CmdRecha.Enabled = False
            ChkAutorizar.Enabled = False
            Fg1.ColWidth(9) = 0
        ElseIf QueHace = 3 Then
            If Val(LblIdEstado.Caption) = 1 Then 'PENDI
                CmdAprobada.Enabled = True
                CmdRecha.Enabled = True
                ChkAutorizar.Enabled = True
                Fg1.ColWidth(9) = 1065
            ElseIf Val(LblIdEstado.Caption) = 2 Then 'APROBADO
                CmdAprobada.Enabled = False
                CmdRecha.Enabled = True
                ChkAutorizar.Enabled = False
                Fg1.ColWidth(9) = 0
            ElseIf Val(LblIdEstado.Caption) = 3 Then 'PROCESADO
                CmdAprobada.Enabled = False
                CmdRecha.Enabled = False
                ChkAutorizar.Enabled = False
                Fg1.ColWidth(9) = 1065
            ElseIf Val(LblIdEstado.Caption) = 4 Then 'RECHAZADA
                CmdAprobada.Enabled = True
                CmdRecha.Enabled = False
                ChkAutorizar.Enabled = True
                Fg1.ColWidth(9) = 1065
            ElseIf Val(LblIdEstado.Caption) = 0 Then 'VACIO
                CmdAprobada.Enabled = False
                CmdRecha.Enabled = False
                ChkAutorizar.Enabled = False
                Fg1.ColWidth(9) = 0
            End If
        End If
    End If
End Sub

Private Function fVerifSiSelecAutPaAprob() As Boolean
    Dim i_verif As Long, verif As Long
    For i_verif = 1 To Fg1.Rows - 1
'--FORMATO DEL GRID DETALLE
'   1            2             3        4         5        6        7           8        9
'proveedor, Nro Ord Compra, Fec. Emis, Moneda, Imp. Total, Saldo, A Cuenta, Restante, Autorizado,
'      10          11
'IdOrdenCompra, IdProv.
'-----------------------------
        If Abs(Val(Fg1.TextMatrix(i_verif, 9))) = 1 Then
            verif = verif + 1
        End If
    Next
    If verif = 0 Then
        MsgBox "Falta especificar el documento o los documentos que desea aprobar.", vbInformation, xTitulo
        fVerifSiSelecAutPaAprob = False
    Else
        fVerifSiSelecAutPaAprob = True
    End If
End Function

Sub MuestraSegundoTab()
    If Dg1.ApproxCount <= 0 Then
        MsgBox "No hay datos para mostrar.", vbInformation, xTitulo
        Exit Sub
    End If
    
    Dim vColorEst As Variant
    Blanquea
    'Bloquea False
    
    LblIdProg.Caption = RstOrdPago("idprog")
    LblProg.Caption = RstOrdPago("nomprog2")
    
    If vProgra_o_Autoriz = 2 Or vProgra_o_Autoriz = 3 Then
        LblIdPersona.Caption = vArrAutor(2)
        LblNombre.Caption = vArrAutor(3)
    Else
        LblIdPersona.Caption = NulosN(RstOrdPago("idaut"))
        LblNombre.Caption = NulosC(RstOrdPago("nomaut2"))
    End If
    
    
    TxtCodigo.Text = NulosC(RstOrdPago("numdoc"))
    TxtFecEmisOrdPag.Valor = RstOrdPago("fchemi")
    TxtFecPago.Valor = RstOrdPago("fchpago")
    If NulosN(RstOrdPago("idmon")) = 1 Then
        OptSol.Value = True
    Else
        OptDol.Value = True
    End If
    
    LblIdEstado.Caption = RstOrdPago("idest")
    LblEstado.Caption = RstOrdPago("desest")
    vColorEst = RstOrdPago("color").Value
    vColorEst = Val(vColorEst)
    LblEstado.ForeColor = vColorEst

    TxtObs.Text = NulosC(RstOrdPago("obs"))
    
    'Mostramos el detalle de la orden de compra
    Dim xCad As String
    Dim A As Integer
    Dim RstTmp As New ADODB.Recordset
'--FORMATO DEL GRID DETALLE
'   1            2             3        4         5        6        7           8        9
'proveedor, Nro Ord Compra, Fec. Emis, Moneda, Imp. Total, Saldo, A Cuenta, Restante, Autorizado,
'      10          11
'IdOrdenCompra, IdProv.
'-----------------------------
    '--LOS APROBADOS
    If RstOrdPago("idaut") <> 0 Then
        Fg1.ColWidth(9) = 1000
    Else
        Fg1.ColWidth(9) = 0
    End If
    xCad = "SELECT mae_prov.nombre, com_compras.numser & '-' & com_compras.numdoc AS numdoccompra, con_ordenpago.fchemi, mae_moneda.simbolo, com_compras.imptot, con_ordenpagodet.saldo, con_ordenpagodet.acuenta, con_ordenpagodet.saldo-con_ordenpagodet.acuenta AS salrest, iif(con_ordenpago.idest = 1, 0, -1) AS autorizado, com_compras.id as idcomp, mae_prov.id as idprov" _
        & " FROM mae_moneda RIGHT JOIN (con_ordenpago LEFT JOIN ((mae_prov RIGHT JOIN com_compras ON mae_prov.id = com_compras.idpro) RIGHT JOIN con_ordenpagodet ON com_compras.id = con_ordenpagodet.idcom) ON con_ordenpago.id = con_ordenpagodet.idord) ON mae_moneda.id = con_ordenpago.idmon" _
        & " WHERE con_ordenpagodet.idord = " & NulosN(RstOrdPago("id")) & ""

    Set RstTmp = BuscaConCriterio(xCad, xCon)
    Mostrando = True
    
    If RstTmp.RecordCount > 0 Then
        LlenarGridDetalle RstTmp
    End If
    '------------------------------
    
    '--LOS RECHAZADOS
    xCad = "SELECT mae_prov.nombre, com_compras.numser & '-' & com_compras.numdoc AS numdoccompra, con_ordenpago.fchemi, mae_moneda.simbolo, com_compras.imptot, con_ordpagodetrechaz.saldo, con_ordpagodetrechaz.acuenta, con_ordpagodetrechaz.saldo-con_ordpagodetrechaz.acuenta AS salrest, 0 AS autorizado, com_compras.id AS idcomp, mae_prov.id AS idprov" _
        & " FROM mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (com_compras RIGHT JOIN (con_ordenpago LEFT JOIN con_ordpagodetrechaz ON con_ordenpago.id = con_ordpagodetrechaz.idord) ON com_compras.id = con_ordpagodetrechaz.idcom) ON mae_prov.id = com_compras.idpro) ON mae_moneda.id = con_ordenpago.idmon" _
        & " WHERE con_ordpagodetrechaz.idord = " & NulosN(RstOrdPago("id")) & ""
    Set RstTmp = BuscaConCriterio(xCad, xCon)
    If RstTmp.RecordCount > 0 Then
        LlenarGridDetalle RstTmp
    End If
    '---------------------------
        
    Mostrando = False
    Dim vArray(1 To 4) As Integer
    vArray(1) = 5: vArray(2) = 6: vArray(3) = 7: vArray(4) = 8
    '--SUMA LOS APROBADOS
    SumarCol vArray, 1
    '--SUMA LOS RECHAZADOS
    SumarCol vArray, 0
    
    '--CONTROLAR HABILITACINES DE BOTONES DE AUTORIZADOR
    HabDeshabControles_de_Autor
End Sub

Private Sub fBuscaEstado(pIdEst As Integer)
    Dim RsEstado As New ADODB.Recordset
    vStrSql = "SELECT mae_estados.id, mae_estados.descripcion" _
        & " FROM mae_estados" _
        & " WHERE mae_estados.id = " & pIdEst & ""
    RST_Busq RsEstado, vStrSql, xCon
    If RsEstado.RecordCount > 0 Then
        LblIdEstado.Caption = NulosN(RsEstado("id"))
        LblEstado.Caption = NulosC(RsEstado("descripcion"))
    End If
    Set RsEstado = Nothing
End Sub

Sub BloqueBoton(pBool As Boolean)
'1: nuevo, 2: modificar, 3: eliminar, 5: grabar, 6: cancelar
    Toolbar1.Buttons(1).Enabled = pBool
    Toolbar1.Buttons(2).Enabled = pBool
    Toolbar1.Buttons(3).Enabled = pBool
    Toolbar1.Buttons(5).Enabled = pBool
    Toolbar1.Buttons(6).Enabled = pBool
End Sub
 
Private Function fVerifAutorizador() As String()
    Dim vArr(1 To 3) As String
    Dim RsVerif As New ADODB.Recordset
    vStrSql = "SELECT con_emptes.id, [pla_empleados].[ape] & ', ' & [pla_empleados].[nom] AS nombre FROM (pla_empleados INNER JOIN con_emptes " _
        & " ON pla_empleados.id = con_emptes.idemp) LEFT JOIN mae_usuarios ON pla_empleados.id = mae_usuarios.idemp " _
        & " WHERE (((mae_usuarios.id) = " & xIdUsuario & ") AND ((con_emptes.aut)=-1))"

    
    '"SELECT  con_emptes.id, [pla_empleados].[nom] & ' ' & [pla_empleados].[ape] AS nombre" _
        & " FROM (pla_empleados INNER JOIN mae_usuarios ON pla_empleados.id = mae_usuarios.id) INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp" _
        & " WHERE mae_usuarios.id= " & CStr(xIdUsuario) & " AND con_emptes.aut = -1"
        
        
    RST_Busq RsVerif, vStrSql, xCon
    If RsVerif.RecordCount > 0 Then
        vArr(1) = "1"
        vArr(2) = RsVerif("id")
        vArr(3) = RsVerif("nombre")
    Else
        vArr(1) = "0"
        vArr(2) = ""
        vArr(3) = ""
    End If
    fVerifAutorizador = vArr
    Set RsVerif = Nothing
End Function

Private Function fVerifProgramador() As String()
    Dim vArr(1 To 3) As String
    Dim RsVerif As New ADODB.Recordset
    vStrSql = "SELECT con_emptes.id, [pla_empleados].[ape] & ', ' & [pla_empleados].[nom] AS nombre FROM (pla_empleados INNER JOIN con_emptes " _
        & " ON pla_empleados.id = con_emptes.idemp) LEFT JOIN mae_usuarios ON pla_empleados.id = mae_usuarios.idemp " _
        & " WHERE ((mae_usuarios.id = " & xIdUsuario & ") AND (con_emptes.prog=-1))"

    RST_Busq RsVerif, vStrSql, xCon
    If RsVerif.RecordCount > 0 Then
        vArr(1) = "1"
        vArr(2) = RsVerif("id")
        vArr(3) = RsVerif("nombre")
    Else
        vArr(1) = "0"
        vArr(2) = ""
        vArr(3) = ""
    End If
    fVerifProgramador = vArr
    Set RsVerif = Nothing
End Function

Private Function VERIFICAR_PROG_AUT(BUSCAPROGRAMADOR As Boolean, OBJ_ID As Label, OBJ_NOMBRE As Label) As Boolean
    Dim RST_TMP As New ADODB.Recordset
    Dim N_SQL As String
    Dim N_SQL_PROG As String
    If BUSCAPROGRAMADOR = True Then
        N_SQL_PROG = " AND con_emptes.prog)=-1;"
    Else
        N_SQL_PROG = " AND con_emptes.aut)=-1;"
    End If
    
    N_SQL = "SELECT  con_emptes.id, [pla_empleados].[nom] & ' ' & [pla_empleados].[ape] AS nombre " _
    + vbCr + " FROM (pla_empleados INNER JOIN mae_usuarios ON pla_empleados.id = mae_usuarios.id) INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp " _
    + vbCr + " WHERE mae_usuarios.id= " + CStr(xIdUsuario) + N_SQL_PROG
    
    
    RST_Busq RST_TMP, N_SQL, xCon
    If RST_TMP.State = 0 Then GoTo SALIR
    If RST_TMP.EOF = True Or RST_TMP.BOF = True Then
        OBJ_ID.Caption = "0"
        OBJ_NOMBRE.Caption = "NO ES " + IIf(BUSCAPROGRAMADOR = True, "PROGRAMADOR", "AUTORIZADOR")
        VERIFICAR_PROG_AUT = False
    Else
        VERIFICAR_PROG_AUT = True
        OBJ_ID.Caption = RST_TMP.Fields(0) & ""
        OBJ_NOMBRE.Caption = IIf(BUSCAPROGRAMADOR = True, "PROGRAMADOR", "AUTORIZADOR") + ":  " + RST_TMP.Fields(1) & ""
    End If
SALIR:
    Set RST_TMP = Nothing
End Function

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

Function HallaNumDoc() As Long
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT id, numdoc FROM con_ordenpago ORDER BY id desc", xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        HallaNumDoc = Format(NulosN(Rst("id")) + 1, "0000")
    Else
        HallaNumDoc = Format("1", "0000")
    End If
    Set Rst = Nothing
End Function

Sub Bloquea(pBool As Boolean)
    Dim obj As Object
    For Each obj In Me.Controls
        If obj.Tag = "a" And TypeName(obj) = "TextBox" Then
            obj.Locked = Not pBool
        ElseIf (obj.Tag = "b" And TypeName(obj) = "TextBoxFecha") Then
            obj.Locked = Not pBool
        ElseIf obj.Tag = "b" Then
            obj.Enabled = pBool
        End If
    Next
End Sub

Private Function fVerifSiExistId(pId As Long) As Boolean
    Dim RsVerif As New ADODB.Recordset
    vStrSql = "SELECT id, numdoc FROM con_ordenpago WHERE id = " & pId & ""
    RST_Busq RsVerif, vStrSql, xCon
    If RsVerif.RecordCount > 0 Then
        fVerifSiExistId = True
    Else
        fVerifSiExistId = False
    End If
    Set RsVerif = Nothing
End Function

Private Function fVerifDatosObligat() As Boolean
    If vProgra_o_Autoriz = 1 Or vProgra_o_Autoriz = 3 Then 'PROGRAMADOR
        'If val(LblIdPersona.Caption) = 0 Or Trim(LblNombre.Caption) = "" Then
        If Val(LblIdProg.Caption) = 0 Or Trim(LblProg.Caption) = "" Then
            fVerifDatosObligat = True
            MsgBox "Falta especificar el programador de la orden de pago.", vbInformation, xTitulo
            Exit Function
        Else
            fVerifDatosObligat = False
        End If
        If Fg1.Rows = 1 Then
            MsgBox "Falta especificar los documentos de la orden de pago en detalle.", vbInformation, xTitulo
            fVerifDatosObligat = True
            Exit Function
        Else
            fVerifDatosObligat = False
        End If
        If Val(TxtTotalACuenta.Text) = 0 Then
            MsgBox "Falta ingresar el importe de a cuenta en detalle.", vbInformation, xTitulo
            fVerifDatosObligat = True
            Exit Function
        Else
            fVerifDatosObligat = False
        End If
        If IsDate(TxtFecEmisOrdPag.Valor) = False Then
            fVerifDatosObligat = True
            MsgBox "Ingrese una fecha de emisión válida.", vbInformation, xTitulo
            Exit Function
        Else
            fVerifDatosObligat = False
        End If
        If IsDate(TxtFecPago.Valor) = False Then
            fVerifDatosObligat = True
            MsgBox "Ingrese una fecha de pago válida.", vbInformation, xTitulo
            Exit Function
        Else
            fVerifDatosObligat = False
        End If
    ElseIf vProgra_o_Autoriz = 2 Then 'AUTORIZADOR
        If Trim(LblIdPersona.Caption) = "" Then
            fVerifDatosObligat = True
            MsgBox "Falta especificar la persona que autoriza la orden de pago.", vbInformation, xTitulo
            Exit Function
        Else
            fVerifDatosObligat = False
        End If
    End If
End Function

Function Grabar() As Boolean
    Grabar = False
        
    '--AQUI VERIFICAR DATOS OBLIGATORIOS
    If fVerifDatosObligat = True Then
        Exit Function
    End If
    
    Dim A As Integer, xId As Long
    Dim Rst As New ADODB.Recordset
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
        
On Error GoTo LaCague
    
    xCon.BeginTrans
    If QueHace = 1 Then
        '--AQUI VERIFICAR SI EXISTE SI EXISTE EL ID
'        If fVerifSiExistId(val(TxtCodigo.Text)) = True Then
'            MsgBox "El numero del documento ya existe.", vbInformation, xTitulo
'            Exit Function
'        End If
        '-----
        xId = HallaCodigoTabla("con_ordenpago", xCon, "id")
        TxtCodigo.Text = Format(xId, "0000")
        '----
        
        RST_Busq RstCab, "SELECT * FROM con_ordenpago", xCon
        RST_Busq RstDet, "SELECT * FROM con_ordenpagodet", xCon
        RstCab.AddNew
        RstCab("id") = Val(TxtCodigo.Text)
    ElseIf QueHace = 2 Then 'PARA MODIFICAR
        xCon.Execute "DELETE FROM con_ordenpagodet WHERE idord = " & NulosN(RstOrdPago("id")) & ""
        
        RST_Busq RstCab, "SELECT * FROM con_ordenpago WHERE id = " & NulosN(RstOrdPago("id")) & "", xCon
        RST_Busq RstDet, "SELECT * FROM con_ordenpagodet", xCon
    End If
    RstCab("fchemi") = NulosC(TxtFecEmisOrdPag.Valor)
    RstCab("fchpago") = NulosC(TxtFecPago.Valor)
    RstCab("idprog") = NulosN(LblIdProg.Caption)
    'RstCab("idaut") = NulosN(LblIdAutor.Caption)
    RstCab("idest") = NulosN(LblIdEstado.Caption)
    If OptCaja.Value = True Then
        RstCab("tipmov") = 1
    Else
        RstCab("tipmov") = 2
    End If
    RstCab("numdoc") = Format(Trim(TxtCodigo.Text), "0000")
    RstCab("idmon") = vMoneda
    RstCab("candoc") = Fg1.Rows - 1
    RstCab("imptot") = Val(Format(TxtTotalACuenta.Text, "#####0.00"))
    RstCab("obs") = TxtObs.Text
    RstCab.Update
        
    For A = 1 To Fg1.Rows - 1
'   1            2             3        4         5        6        7            8             9
'producto, Nro Ord Compra, Fec. Emis, Moneda, Imp. Total, Saldo, A Cuenta, Saldo Restante, Autorizado,
'      10          11
'IdOrdenCompra, IdProv.
        RstDet.AddNew
        RstDet("idord") = Val(TxtCodigo.Text)
        RstDet("idcom") = Val(Fg1.TextMatrix(A, 10))
        RstDet("saldo") = Val(Format(Fg1.TextMatrix(A, 6), "#####0.00"))
        RstDet("acuenta") = Val(Format(Fg1.TextMatrix(A, 7), "#####0.00"))
        RstDet("nuevosaldo") = Val(Format(Fg1.TextMatrix(A, 8), "#####0.00"))
        RstDet.Update
    Next
    
    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    
    RstOrdPago.Requery
    
    MsgBox "La Orden de Pago se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
LaCague:
'    Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
End Function

Sub Cancelar()
    QueHace = 3
    ActivaTool
    Bloquea False
    Fg1.Editable = flexEDNone
    Fg1.SelectionMode = flexSelectionByRow
    
    Label5.Caption = "Detalle de la Solicitud"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

Sub Modificar()
    QueHace = 2
    ActivaTool
'    Blanquea
    Bloquea True
    Label5.Caption = "Modificando Orden de Pago"
    TabOne1.CurrTab = 1
'    Fg1.Rows = 1
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    
    TabOne1.TabEnabled(0) = False
    TxtFecEmisOrdPag.SetFocus
    
    LblNombre.Caption = ""
    LblIdPersona.Caption = ""
'    PreparaRST
End Sub

Sub Nuevo()
    fBuscaEstado 1
    QueHace = 1
    ActivaTool
    Blanquea
    Bloquea True
    Label5.Caption = "Agregando Ordenes de Pago"
    TabOne1.CurrTab = 1
    Fg1.Rows = 1
    TabOne1.TabEnabled(0) = False
    TxtCodigo.Text = Format(HallaCodigoTabla("con_ordenpago", xCon, "id"), "0000")
    
    TxtFecEmisOrdPag.SetFocus
    HabDeshabControles_de_Autor
    
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    
    LblIdProg.Caption = vArrProg(2)
    LblProg.Caption = vArrProg(3)
    Fg1.ColWidth(9) = 0
    LblIdPersona.Caption = ""
    LblNombre.Caption = ""
    
    LblEstado.ForeColor = Val(vArrColorEstado(1))
End Sub

Private Sub ActualSaldoRestan(pCol1 As Integer, pCol2 As Integer)
    Fg1.TextMatrix(Fg1.Row, 8) = Val(Format(Fg1.TextMatrix(Fg1.Row, pCol1), "#####0.00")) - Val(Format(Fg1.TextMatrix(Fg1.Row, pCol2), "#####0.00"))
    Fg1.TextMatrix(Fg1.Row, 8) = Format(Fg1.TextMatrix(Fg1.Row, 8), "#,###0.00")
End Sub

Private Sub Blanquea()
    Dim obj As Object
    For Each obj In Me.Controls
        If obj.Tag = "a" Then
            obj = ""
        End If
    Next
    TxtFecEmisOrdPag.Valor = Date
    TxtFecPago.Valor = Date
    
    TxtTotalImp.Text = ""
    TxtTotSaldo.Text = ""
    TxtTotalACuenta.Text = ""
    TxtTotNueSal.Text = ""
    
    TxtTotImpRec.Text = ""
    TxtTotSalAntRec.Text = ""
    TxtTotACuentaRec.Text = ""
    TxtTotNueSalRec.Text = ""
    
    Fg1.Rows = 1
End Sub

Sub SumarCol(pArray() As Integer, pQSuma As Integer)
    Dim i_sum As Long
    Dim i_row As Long
    Dim vSumTotal As Double
'   1            2             3        4         5        6        7          8          9
'producto, Nro Ord Compra, Fec. Emis, Moneda, Imp. Total, Saldo, A Cuenta, Nuevo Sal, Autorizado,
'      10          11
'IdOrdenCompra, IdProv.
    For i_sum = LBound(pArray) To UBound(pArray)
        vSumTotal = 0
        For i_row = 1 To Fg1.Rows - 1
            With Fg1
                If pQSuma = Abs(Val(.TextMatrix(i_row, 9))) Then
                    vSumTotal = vSumTotal + Val(Format(.TextMatrix(i_row, pArray(i_sum)), "######0.00"))
                End If
            End With
        Next
        If pQSuma = 1 Then 'INDICA LOS APROBADOS
            Select Case UCase(Trim(Fg1.TextMatrix(0, pArray(i_sum))))
                Case "IMP. TOTAL"
                    TxtTotalImp.Text = Format(vSumTotal, "#,###0.00")
                Case "SALDO ANT."
                    TxtTotSaldo.Text = Format(vSumTotal, "#,###0.00")
                Case "A CUENTA"
                    TxtTotalACuenta.Text = Format(vSumTotal, "#,###0.00")
                Case "NUEVO SAL."
                    TxtTotNueSal.Text = Format(vSumTotal, "#,###0.00")
            End Select
        ElseIf pQSuma = 0 Then 'LOS RECHAZADOS
            If vCalcularTotales = 0 Then
                If Val(LblIdEstado.Caption) <> 1 Then
                    LblTotAprob.Caption = "Total Aprobados"
                    Select Case UCase(Trim(Fg1.TextMatrix(0, pArray(i_sum))))
                        Case "IMP. TOTAL"
                            TxtTotImpRec.Text = Format(vSumTotal, "#,###0.00")
                        Case "SALDO ANT."
                            TxtTotSalAntRec.Text = Format(vSumTotal, "#,###0.00")
                        Case "A CUENTA"
                            TxtTotACuentaRec.Text = Format(vSumTotal, "#,###0.00")
                        Case "NUEVO SAL."
                            TxtTotNueSalRec.Text = Format(vSumTotal, "#,###0.00")
                    End Select
                ElseIf Val(LblIdEstado.Caption) = 1 Then
                    LblTotAprob.Caption = "Total Pendientes"
                    Select Case UCase(Trim(Fg1.TextMatrix(0, pArray(i_sum))))
                        Case "IMP. TOTAL"
                            TxtTotalImp.Text = Format(vSumTotal, "#,###0.00")
                        Case "SALDO ANT."
                            TxtTotSaldo.Text = Format(vSumTotal, "#,###0.00")
                        Case "A CUENTA"
                            TxtTotalACuenta.Text = Format(vSumTotal, "#,###0.00")
                        Case "NUEVO SAL."
                            TxtTotNueSal.Text = Format(vSumTotal, "#,###0.00")
                    End Select
                End If
            ElseIf vCalcularTotales = 1 Then
                LblTotAprob.Caption = "Total Aprobados"
                Select Case UCase(Trim(Fg1.TextMatrix(0, pArray(i_sum))))
                    Case "IMP. TOTAL"
                        TxtTotImpRec.Text = Format(vSumTotal, "#,###0.00")
                    Case "SALDO ANT."
                        TxtTotSalAntRec.Text = Format(vSumTotal, "#,###0.00")
                    Case "A CUENTA"
                        TxtTotACuentaRec.Text = Format(vSumTotal, "#,###0.00")
                    Case "NUEVO SAL."
                        TxtTotNueSalRec.Text = Format(vSumTotal, "#,###0.00")
                End Select
            End If
        End If
    Next
End Sub

Sub LimpiarGrid()
    Fg1.Clear
    Fg1.Rows = 1
    Fg1.FormatString = vFormatString
    Fg1.Editable = True
    Fg1.SelectionMode = flexSelectionFree
    Fg1.ColDataType(9) = flexDTBoolean
    
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
End Sub

Function AddItemGrid(pGrid As Object) As Boolean
    If pGrid.Rows = 1 Then
        pGrid.AddItem ""
        AddItemGrid = True
    Else
        If pGrid.TextMatrix(1, 1) <> "" Then
            pGrid.AddItem ""
            AddItemGrid = True
        Else
            AddItemGrid = False
        End If
    End If
End Function

Sub CargarFacturasPorPagar()
    Dim vArray(1 To 4) As Integer
    Dim xForm As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(8, 5) As String
    
    xCampos(0, 0) = "Proveedor":    xCampos(0, 1) = "nombre":     xCampos(0, 2) = "3000":   xCampos(0, 3) = "C":    xCampos(0, 4) = "N"
    xCampos(1, 0) = "T.D.":         xCampos(1, 1) = "abrev":      xCampos(1, 2) = "600":    xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "Num. Doc.":    xCampos(2, 1) = "numdoc1":    xCampos(2, 2) = "1500":   xCampos(2, 3) = "C":    xCampos(2, 4) = "S"
    xCampos(3, 0) = "Fecha Doc.":   xCampos(3, 1) = "fchdoc":     xCampos(3, 2) = "1000":   xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
    xCampos(4, 0) = "Fecha Venc.":  xCampos(4, 1) = "fchven":     xCampos(4, 2) = "1000":   xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
    xCampos(5, 0) = "Moneda":       xCampos(5, 1) = "simbolo":    xCampos(5, 2) = "800":    xCampos(5, 3) = "C":    xCampos(5, 4) = "N"
    xCampos(6, 0) = "Importe":      xCampos(6, 1) = "imptot":     xCampos(6, 2) = "1200":   xCampos(6, 3) = "N":    xCampos(6, 4) = "N"
    xCampos(7, 0) = "Saldo":        xCampos(7, 1) = "impsal":     xCampos(7, 2) = "1200":   xCampos(7, 3) = "N":    xCampos(7, 4) = "N"
    
    vStrSql = "SELECT com_compras.id, mae_prov.nombre, mae_documento.abrev, com_compras.numser & '-' & com_compras.numdoc AS numdoc1, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, com_compras.imptot, com_compras.impsal, mae_prov.id AS idprov" _
        & " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro" _
        & " WHERE com_compras.id NOT IN (SELECT con_ordenpagodet.idcom" _
        & " FROM con_ordenpago INNER JOIN con_ordenpagodet ON con_ordenpago.id = con_ordenpagodet.idord" _
        & " Where con_ordenpago.fchpago = CDate('" & Trim(TxtFecPago.Valor) & "'))" & "" _
        & " AND com_compras.impsal > 0" _
        & " ORDER BY mae_prov.nombre, com_compras.numser, com_compras.numdoc"
    
    If OptSol.Value = True Then
        vStrSql = vStrSql & " AND com_compras.idmon = 1"
    Else
        vStrSql = vStrSql & " AND com_compras.idmon = 2"
    End If
    
    xForm.SQLCad = vStrSql
    xForm.Titulo = "Buscando Documentos de Compras con Saldo"
    Set xForm.Coneccion = xCon
    Set xRs = Nothing
    Set xRs = xForm.Seleccionar(xCampos)
    If xRs.State = 1 Then
        Dim A As Integer
        Dim xFila As Integer
        
        If xRs.RecordCount > 0 Then
            xRs.MoveFirst
            Do While Not xRs.EOF
                For A = 1 To Fg1.Rows - 1
                    If Fg1.TextMatrix(A, 10) = NulosN(xRs("id")) And Fg1.TextMatrix(A, 11) = xRs("idprov") Then
                        If MsgBox("Ya esta agregado el doc. de compra: " & xRs("numdoc1") & " del proveedor: " & xRs("nombre") & "" & vbCrLf & " Desea reemplazarlo.", vbYesNo + vbQuestion, xTitulo) = vbYes Then
                            Fg1.TextMatrix(A, 1) = NulosC(xRs("nombre"))
                            Fg1.TextMatrix(A, 2) = NulosC(xRs("numdoc1"))
                            Fg1.TextMatrix(A, 3) = NulosC(xRs("fchdoc"))
                            Fg1.TextMatrix(A, 4) = NulosC(xRs("simbolo"))
                            Fg1.TextMatrix(A, 5) = Format(NulosC(xRs("imptot")), "#,###0.00")
                            Fg1.TextMatrix(A, 6) = Format(NulosC(xRs("impsal")), "#,###0.00")
                            Fg1.TextMatrix(A, 7) = ""
                            '.TextMatrix(.Rows - 1, 8) = ? VALOR BOOLEANO
                            Fg1.TextMatrix(A, 10) = NulosN(xRs("id"))
                            Fg1.TextMatrix(A, 11) = NulosN(xRs("idprov"))
                            Set xRs = Nothing
                            Set xForm = Nothing
                            Exit Sub
                        Else
                            Set xRs = Nothing
                            Set xForm = Nothing
                            Exit Sub
                        End If
                    End If
                Next
                xRs.MoveNext
            Loop
        End If
        
        
        If xRs.RecordCount <> 0 Then
            xRs.MoveFirst
            Do While Not xRs.EOF
                If AddItemGrid(Fg1) = True Then
'   1            2             3        4         5        6        7           8        9
'proveedor, Nro Ord Compra, Fec. Emis, Moneda, Imp. Total, Saldo, A Cuenta, Restante, Autorizado,
'      10          11
'IdOrdenCompra, IdProv.
                    With Fg1
                        .TextMatrix(.Rows - 1, 1) = NulosC(xRs("nombre"))
                        .TextMatrix(.Rows - 1, 2) = NulosC(xRs("numdoc1"))
                        .TextMatrix(.Rows - 1, 3) = Format(NulosC(xRs("fchdoc")), "dd/mm/yy")
                        .TextMatrix(.Rows - 1, 4) = NulosC(xRs("simbolo"))
                        .TextMatrix(.Rows - 1, 5) = Format(NulosC(xRs("imptot")), "#,###0.00")
                        .TextMatrix(.Rows - 1, 6) = Format(NulosC(xRs("impsal")), "#,###0.00")
                        .TextMatrix(.Rows - 1, 7) = ""
                        '.TextMatrix(.Rows - 1, 8) = ? VALOR BOOLEANO
                        .TextMatrix(.Rows - 1, 10) = NulosN(xRs("id"))
                        .TextMatrix(.Rows - 1, 11) = NulosN(xRs("idprov"))
                    End With
                End If
                xRs.MoveNext
            Loop
            vArray(1) = 5: vArray(2) = 6: vArray(3) = 7: vArray(4) = 8
            SumarCol vArray, 1
            SumarCol vArray, 0
        End If
    End If
    Set xRs = Nothing
    Set xForm = Nothing
End Sub

Private Sub CmdAdd_Click()
    CargarFacturasPorPagar
End Sub

Private Sub CmdAprobada_Click()
    If fVerifSiLosAutorSonIguales = False Then
        MsgBox "Lo siento usted no puede aprobar la orden de pago.", vbInformation, xTitulo
        Exit Sub
    End If

    If fVerifSiSelecAutPaAprob = False Then
        Exit Sub
    End If
    
    If MsgBox("Esta seguro de aprobar la orden de pago.", vbYesNo + vbQuestion, xTitulo) = vbNo Then
        Exit Sub
    End If
    
    Dim i_apro As Long
    Dim vCtaAprob As Long, vCtaRechaz As Long
    Dim vImpTotalAprob As Double
        
    '--APROBADOS
    Dim RsLosAprob As New ADODB.Recordset
    vStrSql = "SELECT con_ordenpagodet.*" _
        & " FROM con_ordenpagodet"
    
    RST_Busq RsLosAprob, vStrSql, xCon
    '--FIN APROBADOS
    
    '--RECHAZADOS
    Dim RsLosRechaz As New ADODB.Recordset
    vStrSql = "SELECT con_ordpagodetrechaz.*" _
        & " FROM con_ordpagodetrechaz"
    
    RST_Busq RsLosRechaz, vStrSql, xCon
    '--FIN RECHAZADOS
'   1            2             3        4         5          6        7           8        9
'proveedor, Nro Ord Compra, Fec. Emis, Moneda, Imp. Total, Saldo, A Cuenta, Restante, Autorizado,
'      10          11
'IdOrdenCompra, IdProv.

'    vStrSql = "UPDATE con_ordenpago SET idaut = " & val(LblIdPersona.Caption) & ", idest = 2 WHERE id = " & NulosN(RstOrdPago("id")) & ""
'    xCon.Execute vStrSql
    
    xCon.Execute ("DELETE FROM con_ordenpagodet WHERE idord = " & NulosN(RstOrdPago("id")) & "")
    xCon.Execute ("DELETE FROM con_ordpagodetrechaz WHERE idord = " & NulosN(RstOrdPago("id")) & "")
    
    '--LOS APROBADOS
    For i_apro = 1 To Fg1.Rows - 1
        If Abs(Val(Fg1.TextMatrix(i_apro, 9))) = 1 Then
            RsLosAprob.AddNew
            RsLosAprob("idord") = NulosN(RstOrdPago("id"))
            RsLosAprob("idcom") = Val(Fg1.TextMatrix(i_apro, 10))
            RsLosAprob("saldo") = Val(Fg1.TextMatrix(i_apro, 6))
            RsLosAprob("nuevosaldo") = Val(Fg1.TextMatrix(i_apro, 8))
            RsLosAprob("acuenta") = Val(Fg1.TextMatrix(i_apro, 7))
            RsLosAprob.Update
            
            vCtaAprob = vCtaAprob + 1
            vImpTotalAprob = vImpTotalAprob + Val(Format(Fg1.TextMatrix(i_apro, 7), "#####0.00"))
        End If
    Next
    
    '--LOS RECHAZADOS
    For i_apro = 1 To Fg1.Rows - 1
        If Abs(Val(Fg1.TextMatrix(i_apro, 9))) = 0 Then
            RsLosRechaz.AddNew
            RsLosRechaz("idord") = NulosN(RstOrdPago("id"))
            RsLosRechaz("idcom") = Fg1.TextMatrix(i_apro, 10)
            RsLosRechaz("saldo") = Fg1.TextMatrix(i_apro, 6)
            RsLosRechaz("nuevosaldo") = Fg1.TextMatrix(i_apro, 8)
            RsLosRechaz("acuenta") = Fg1.TextMatrix(i_apro, 7)
            RsLosRechaz.Update
            
            vCtaRechaz = vCtaRechaz + 1
'            xCon.Execute ("DELETE FROM con_ordenpagodet WHERE idord = " & NulosN(RstOrdPago("id")) & " AND idcom = " & val(Fg1.TextMatrix(i_apro, 10)) & "")
        End If
    Next
    
    '--ACTUALIZA LA TABLA DEL ENCABEZADO DE ORDEN DE PAGO
    vStrSql = "UPDATE con_ordenpago SET idaut = " & Val(LblIdPersona.Caption) & ", idest = 2, candoc = " & vCtaAprob & ", imptot = " & vImpTotalAprob & "" _
        & " WHERE id = " & NulosN(RstOrdPago("id")) & ""
    xCon.Execute vStrSql
    '-------------------------------------------------------
    RstOrdPago.Requery
    
    MsgBox "La Orden de pago se aprobó satisfactoriamente.", vbInformation, xTitulo
    
    Set RsLosAprob = Nothing
    Set RsLosRechaz = Nothing
    TabOne1.CurrTab = 0
End Sub

Private Sub CmdDel_Click()
    If Fg1.Rows > 1 Then
        On Error Resume Next
        Fg1.RemoveItem Fg1.Row
        Dim vArray(1 To 4) As Integer
        vArray(1) = 5: vArray(2) = 6: vArray(3) = 7: vArray(4) = 8
        '--SUMA LOS APROBADOS
        SumarCol vArray, 1
        '--SUMA LOS RECHAZADOS
        SumarCol vArray, 0
    End If
End Sub

Private Sub CmdOk_Click()
'    Dim xFecha As String
'    xFecha = Format(MonthView1.Value, "dd/mm/yy")
'
'    xFchIni = "01/" + Mid(Format(CDate(xFecha), "dd/mm/yy"), 4, 5)
'    xFchFin = Trim(Format(HallaDiasMes(CDate(xFecha)), "00")) + "/" + Mid(Format(CDate(xFecha), "dd/mm/yy"), 4, 5)
'
'    RST_Busq RstGui, "SELECT Format(vta_guia.numser,'0000')+'-'+Format(vta_guia.numdoc,'0000000000') AS numguia, VTA_Guia.*, MAE_Cliente.nombre, MAE_Cliente.dir, " _
'        & " VTA_PuntoVenta.descripcion AS despunven, IIf(vta_guia.Anulado=0,'','Anulado') AS Anulado, VTA_PuntoVenta.dir AS Direccion, mae_emptra.nombre AS desemptra, " _
'        & " mae_emptra.numruc, mae_mottra.descripcion AS descmotgui, MAE_Cliente.numruc AS RucCli, UCase([pla_empleados].[ape])+', '+[pla_empleados].[nom] AS apenomcho, " _
'        & " mae_chofer.numbre, mae_vehiculo.marca AS marcacar, mae_vehiculo.numpla, Format([vta_ventas].[numser],'0000')+'-'+Format([vta_ventas].[numdoc],'0000000000') AS numdocref" _
'        & " FROM pla_empleados RIGHT JOIN (vta_ventas RIGHT JOIN ((((((VTA_Guia LEFT JOIN MAE_Cliente ON VTA_Guia.idcli = MAE_Cliente.id) LEFT JOIN VTA_PuntoVenta " _
'        & " ON VTA_Guia.idpunven = VTA_PuntoVenta.id) LEFT JOIN mae_emptra ON VTA_Guia.idemptra = mae_emptra.id) LEFT JOIN mae_mottra ON VTA_Guia.idmottra = mae_mottra.id) " _
'        & " LEFT JOIN mae_chofer ON VTA_Guia.idcho = mae_chofer.id) LEFT JOIN mae_vehiculo ON VTA_Guia.idveh = mae_vehiculo.id) ON vta_ventas.id = VTA_Guia.iddocven) " _
'        & " ON pla_empleados.id = mae_chofer.idper WHERE (((VTA_Guia.fecgiro)>=CDate('" & xFchIni & "') And (VTA_Guia.fecgiro)<=CDate('" & xFchFin & "'))) " _
'        & " ORDER BY VTA_Guia.numser, VTA_Guia.numdoc DESC", xCon

'valia
'    Set RstOrdPago = fDevolRsParaCons_AlActivarVentana("01/" + Trim(Format(xMes, "00")) + "/" + Trim(Format(Year(Date), "0000")))
'    Set Dg1.DataSource = RstOrdPago
'    CmdSalir_Click
End Sub

Private Sub CmdRecha_Click()
    If fVerifSiLosAutorSonIguales = False Then
        MsgBox "Lo siento usted no puede rechazar la orden de pago.", vbInformation, xTitulo
        Exit Sub
    End If

    If MsgBox("Está seguro de rechazar los documentos de la orden de pago.", vbYesNo + vbQuestion, xTitulo) = vbNo Then
        Exit Sub
    End If

    xCon.Execute ("DELETE FROM con_ordpagodetrechaz WHERE idord = " & RstOrdPago("id") & "")
    xCon.Execute ("DELETE FROM con_ordenpagodet WHERE idord = " & RstOrdPago("id") & "")
    Dim i_rec As Long
    
    Dim vCtaAprob As Long, vCtaRechaz As Long
    Dim vImpTotalAprob As Double
    
    Dim RsLosRechaz As New ADODB.Recordset
        
    vStrSql = "SELECT con_ordpagodetrechaz.*" _
        & " FROM con_ordpagodetrechaz" _
    
    RST_Busq RsLosRechaz, vStrSql, xCon
    
'   1            2             3        4         5          6        7           8        9
'proveedor, Nro Ord Compra, Fec. Emis, Moneda, Imp. Total, Saldo, A Cuenta, Restante, Autorizado,
'      10          11
'IdOrdenCompra, IdProv.
    vStrSql = "UPDATE con_ordenpago SET idest = 4 WHERE id = " & NulosN(RstOrdPago("id")) & ""
    xCon.Execute vStrSql
    
    For i_rec = 1 To Fg1.Rows - 1
        RsLosRechaz.AddNew
        RsLosRechaz("idord") = NulosN(RstOrdPago("id"))
        RsLosRechaz("idcom") = Fg1.TextMatrix(i_rec, 10)
        RsLosRechaz("saldo") = Fg1.TextMatrix(i_rec, 6)
        RsLosRechaz("nuevosaldo") = Fg1.TextMatrix(i_rec, 8)
        RsLosRechaz("acuenta") = Fg1.TextMatrix(i_rec, 7)
        RsLosRechaz.Update
    Next
    
    '--ACTUALIZA EL ENCAB DE LA ORDEN DE PAGO
    vStrSql = "UPDATE con_ordenpago SET candoc = 0, imptot = 0" _
        & " WHERE id = " & NulosN(RstOrdPago("id")) & ""
    xCon.Execute vStrSql
    '----------------------------------------
    
    RstOrdPago.Requery
    
    MsgBox "Orden de pago rechazado.", vbInformation, xTitulo
    TabOne1.CurrTab = 0
End Sub

Private Sub ChkAutorizar_Click()
    Dim i_chk As Long
    If ChkAutorizar.Value = 1 Then
        For i_chk = 1 To Fg1.Rows - 1
            Fg1.TextMatrix(i_chk, 9) = -1
        Next
    Else
        For i_chk = 1 To Fg1.Rows - 1
            Fg1.TextMatrix(i_chk, 9) = 0
        Next
    End If
End Sub

Private Sub Dg1_DblClick()
    If Dg1.ApproxCount > 0 Then
        MuestraSegundoTab
        TabOne1.CurrTab = 1
    End If
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    
    If KeyCode = 46 Then 'SUPRIMIR
        Eliminar
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Mostrando = True Then
        Exit Sub
    End If
'    Dim vArray(1 To 2) As Integer
    Dim vArray(1 To 4) As Integer
    
    Select Case Fg1.Col
        Case 7
'   1            2             3        4         5        6        7            8             9
'producto, Nro Ord Compra, Fec. Emis, Moneda, Imp. Total, Saldo, A Cuenta, Saldo Restante, Autorizado,
'      10          11
'IdOrdenCompra, IdProv.
            If Val(Format(Fg1.TextMatrix(Row, 7), "#####0.00")) > Val(Format(Fg1.TextMatrix(Row, 6), "#####0.00")) Then
                MsgBox "El valor ingresado es mayor al saldo anterior.", vbInformation, xTitulo
                Fg1.TextMatrix(Row, 7) = 0
                Exit Sub
            End If
            ActualSaldoRestan 6, 7
            
            vArray(1) = 7: vArray(2) = 8  ': vArray(2) = 6: vArray(3) = 7
            SumarCol vArray, 1
            SumarCol vArray, 0
        Case 9
            If QueHace = 3 And (vProgra_o_Autoriz = 2 Or vProgra_o_Autoriz = 3) Then
                vCalcularTotales = 1
'                ActualSaldoRestan 6, 7
                
                vArray(1) = 5: vArray(2) = 6: vArray(3) = 7: vArray(4) = 8
                'vArray(1) = 7: vArray(2) = 8  ': vArray(2) = 6: vArray(3) = 7
                SumarCol vArray, 1
                SumarCol vArray, 0
                vCalcularTotales = 0
            End If
    End Select
End Sub

Private Sub Fg1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 46
            If CmdDel.Enabled = True Then
                CmdDel_Click
            End If
    End Select
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 0 Then
        KeyAscii = 0
    End If
    If QueHace <> 1 And QueHace <> 2 Then
        KeyAscii = 0
    End If
    Select Case Col
        Case Is <> 7
            KeyAscii = 0
        Case Is = 7
            If vProgra_o_Autoriz = 1 Or vProgra_o_Autoriz = 3 Then
                Select Case KeyAscii
                    Case 13
                        
                    Case Else
                        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then
                            KeyAscii = 0
                        End If
                End Select
            Else
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim Rpta As Integer
        SeEjecuto = True
                
        Set RstOrdPago = fDevolRsParaCons_AlActivarVentana(Date)
            
        Set Dg1.DataSource = RstOrdPago
        If RstOrdPago.RecordCount = 0 Then
            If vProgra_o_Autoriz = 1 Or vProgra_o_Autoriz = 3 Then
                Rpta = MsgBox("No se ha registrado una orden de pago, ¿Desea agregar una ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
                If Rpta = vbYes Then
                    HabDeshabControles_de_Autor
                    Nuevo
                Else
                    Set RstOrdPago = Nothing
                    Unload Me
                    Exit Sub
                End If
            ElseIf vProgra_o_Autoriz = 2 Then
                Bloquea False
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    vFormatString = Fg1.FormatString
    LimpiarGrid
    Blanquea
    '----
    LblProg.Caption = ""
    LblIdProg.Caption = ""
    LblNombre.Caption = ""
    LblIdPersona.Caption = ""
    '----
    CaracteresNumericos = "0123456789." & Chr(8)
    vMoneda = 1
    QueHace = 3
    
    vArrProg = fVerifProgramador
    
    vArrAutor = fVerifAutorizador
    Bloquea False
    
    If Val(vArrProg(1)) = 1 And Val(vArrAutor(1)) = 0 Then 'ES PROGRAMADOR
        LblIdProg.Caption = vArrProg(2)
        LblProg.Caption = vArrProg(3)
        vProgra_o_Autoriz = 1
    ElseIf Val(vArrProg(1)) = 0 And Val(vArrAutor(1)) = 1 Then 'ES AUTORIZADOR
'        LblMensaje.Caption = "Autorizador de Orden de Pago"
        LblIdPersona.Caption = Trim(vArrAutor(2))
        LblNombre.Caption = Trim(vArrAutor(3))
        vProgra_o_Autoriz = 2
        BloqueBoton False
    ElseIf Val(vArrProg(1)) = 1 And Val(vArrAutor(1)) = 1 Then 'ES PROG Y AUTOR
        vProgra_o_Autoriz = 3
    ElseIf Val(vArrProg(1)) = 0 And Val(vArrAutor(1)) = 0 Then 'PERSONA NO AUTORIZADA
'        LblMensaje.Caption = "Persona no Autorizada"
        LblIdProg.Caption = ""
        LblProg.Caption = ""

        LblIdPersona.Caption = ""
        LblNombre.Caption = ""
        vProgra_o_Autoriz = 0
        BloqueBoton False
    End If
    
'    HabDeshabControles_de_Autor
    TabOne1.CurrTab = 0
    Fg1.ColWidth(9) = 0
    Fg1.Editable = flexEDNone
    Fg1.SelectionMode = flexSelectionByRow
    ColorEstado
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SeEjecuto = False
End Sub

Private Sub OptDol_Click()
    If Fg1.Rows > 1 Then
        If MsgBox("Esta seguro de cambiar el tipo de moneda.", vbYesNo + vbQuestion, xTitulo) = vbNo Then
            Exit Sub
        End If
        Fg1.Rows = 1
    End If
    vMoneda = 2
End Sub

Private Sub OptSol_Click()
    If Fg1.Rows > 1 Then
        If MsgBox("Esta seguro de cambiar el tipo de moneda.", vbYesNo + vbQuestion, xTitulo) = vbNo Then
            Exit Sub
        End If
        Fg1.Rows = 1
    End If
    vMoneda = 1
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            MuestraSegundoTab
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        Nuevo
    End If
    If Button.Index = 2 Then 'MODIFICAR
        If Dg1.ApproxCount <= 0 Then
            MsgBox "No hay datos para modificar.", vbInformation, xTitulo
            Exit Sub
        End If
        If NulosN(RstOrdPago("idest")) = 2 Then 'APROBADO
            MsgBox "El registro seleccionado ya está aprobado. No puede modificarlo.", vbInformation, xTitulo
            Exit Sub
        End If
        If NulosN(RstOrdPago("idest")) = 3 Then 'PROCESADO
            MsgBox "El registro seleccionado ya está procesado. No puede modificarlo.", vbInformation, xTitulo
            Exit Sub
        End If
        If Val(vArrProg(2)) <> NulosN(RstOrdPago("idprog")) Then
            MsgBox "Lo siento usted no puede modificar el orden de pago.", vbInformation, xTitulo
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.Index = 3 Then
        If Dg1.ApproxCount <= 0 Then
            MsgBox "No hay datos para eliminar.", vbInformation, xTitulo
            Exit Sub
        End If
        If NulosN(RstOrdPago("idest")) = 2 Then 'APROBADO
            MsgBox "El registro seleccionado ya está aprobado. No puede eliminarlo.", vbInformation, xTitulo
            Exit Sub
        End If
        If NulosN(RstOrdPago("idest")) = 3 Then 'PROCESADO
            MsgBox "El registro seleccionado ya está procesado. No puede eliminarlo.", vbInformation, xTitulo
            Exit Sub
        End If
        If Val(vArrProg(2)) <> NulosN(RstOrdPago("idprog")) Then
            MsgBox "Lo siento usted no puede eliminar la orden de pago.", vbInformation, xTitulo
            Exit Sub
        End If
        Eliminar
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstOrdPago.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then
        Cancelar
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 12 Then 'CAMBIAR MES
        CambiarMes
    End If
    
    If Button.Index = 16 Then 'SALIR
        Set RstOrdPago = Nothing
        Unload Me
    End If
End Sub

