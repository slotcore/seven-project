VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmCuentasRendir_Dev 
   Caption         =   "Caja y Bancos - Devoluciones de Cuentas por Rendir"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11775
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
            Picture         =   "FrmCuentasRendir_Dev.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasRendir_Dev.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasRendir_Dev.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasRendir_Dev.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasRendir_Dev.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasRendir_Dev.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasRendir_Dev.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasRendir_Dev.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasRendir_Dev.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasRendir_Dev.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasRendir_Dev.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasRendir_Dev.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
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
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Listado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Registro"
               EndProperty
            EndProperty
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
      Height          =   7230
      Left            =   15
      TabIndex        =   0
      Top             =   360
      Width           =   11760
      _cx             =   20743
      _cy             =   12753
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   -12315
         TabIndex        =   9
         Top             =   375
         Width           =   11670
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   6435
            Left            =   60
            TabIndex        =   13
            Top             =   345
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   11351
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Num.Reg."
            Columns(0).DataField=   "dev_numreg"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Fch. Dev."
            Columns(1).DataField=   "dev_emi"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Devuelto"
            Columns(2).DataField=   "dev_imp"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Declarado"
            Columns(3).DataField=   "Declarado"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "x Rendir"
            Columns(4).DataField=   "xrendir"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Num.Reg."
            Columns(5).DataField=   "numreg"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Fch. Emi."
            Columns(6).DataField=   "fchemi"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Fch. Pag."
            Columns(7).DataField=   "fchpag"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Fch. Ren."
            Columns(8).DataField=   "fchren"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "M"
            Columns(9).DataField=   "simbolo"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Importe"
            Columns(10).DataField=   "imp"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Tipo Mov."
            Columns(11).DataField=   "tipmov"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "Origen"
            Columns(12).DataField=   "origen"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "Num. Doc"
            Columns(13).DataField=   "numdoc"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "Destino"
            Columns(14).DataField=   "destino"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "Persona"
            Columns(15).DataField=   "tipper"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   0
            Columns(16)._MaxComboItems=   5
            Columns(16).Caption=   "Entregado A."
            Columns(16).DataField=   "benef"
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   17
            Splits(0)._UserFlags=   0
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).SizeMode=   1
            Splits(0).Size  =   3495.118
            Splits(0).Size.vt=   4
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).Caption=   "Rendición de Cuenta"
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=17"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1535"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1455"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1640"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1561"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1535"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1455"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=514"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1640"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1561"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=514"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1931"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1852"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=514"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=1535"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1455"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5).AllowSizing=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(5).Visible=0"
            Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(39)=   "Column(6).Width=1535"
            Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=1455"
            Splits(0)._ColumnProps(42)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(43)=   "Column(6).AllowSizing=0"
            Splits(0)._ColumnProps(44)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(45)=   "Column(6).Visible=0"
            Splits(0)._ColumnProps(46)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(47)=   "Column(7).Width=1720"
            Splits(0)._ColumnProps(48)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(49)=   "Column(7)._WidthInPix=1640"
            Splits(0)._ColumnProps(50)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(51)=   "Column(7).AllowSizing=0"
            Splits(0)._ColumnProps(52)=   "Column(7)._ColStyle=513"
            Splits(0)._ColumnProps(53)=   "Column(7).Visible=0"
            Splits(0)._ColumnProps(54)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(55)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(56)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(58)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(59)=   "Column(8).AllowSizing=0"
            Splits(0)._ColumnProps(60)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(61)=   "Column(8).Visible=0"
            Splits(0)._ColumnProps(62)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(63)=   "Column(9).Width=767"
            Splits(0)._ColumnProps(64)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(65)=   "Column(9)._WidthInPix=688"
            Splits(0)._ColumnProps(66)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(67)=   "Column(9).AllowSizing=0"
            Splits(0)._ColumnProps(68)=   "Column(9)._ColStyle=516"
            Splits(0)._ColumnProps(69)=   "Column(9).Visible=0"
            Splits(0)._ColumnProps(70)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(71)=   "Column(10).Width=1535"
            Splits(0)._ColumnProps(72)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(73)=   "Column(10)._WidthInPix=1455"
            Splits(0)._ColumnProps(74)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(75)=   "Column(10).AllowSizing=0"
            Splits(0)._ColumnProps(76)=   "Column(10)._ColStyle=514"
            Splits(0)._ColumnProps(77)=   "Column(10).Visible=0"
            Splits(0)._ColumnProps(78)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(79)=   "Column(11).Width=1482"
            Splits(0)._ColumnProps(80)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(81)=   "Column(11)._WidthInPix=1402"
            Splits(0)._ColumnProps(82)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(83)=   "Column(11).AllowSizing=0"
            Splits(0)._ColumnProps(84)=   "Column(11)._ColStyle=516"
            Splits(0)._ColumnProps(85)=   "Column(11).Visible=0"
            Splits(0)._ColumnProps(86)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(87)=   "Column(12).Width=5900"
            Splits(0)._ColumnProps(88)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(89)=   "Column(12)._WidthInPix=5821"
            Splits(0)._ColumnProps(90)=   "Column(12)._EditAlways=0"
            Splits(0)._ColumnProps(91)=   "Column(12).AllowSizing=0"
            Splits(0)._ColumnProps(92)=   "Column(12)._ColStyle=516"
            Splits(0)._ColumnProps(93)=   "Column(12).Visible=0"
            Splits(0)._ColumnProps(94)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(95)=   "Column(13).Width=2725"
            Splits(0)._ColumnProps(96)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(97)=   "Column(13)._WidthInPix=2646"
            Splits(0)._ColumnProps(98)=   "Column(13)._EditAlways=0"
            Splits(0)._ColumnProps(99)=   "Column(13).AllowSizing=0"
            Splits(0)._ColumnProps(100)=   "Column(13)._ColStyle=516"
            Splits(0)._ColumnProps(101)=   "Column(13).Visible=0"
            Splits(0)._ColumnProps(102)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(103)=   "Column(14).Width=4075"
            Splits(0)._ColumnProps(104)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(105)=   "Column(14)._WidthInPix=3995"
            Splits(0)._ColumnProps(106)=   "Column(14)._EditAlways=0"
            Splits(0)._ColumnProps(107)=   "Column(14).AllowSizing=0"
            Splits(0)._ColumnProps(108)=   "Column(14)._ColStyle=516"
            Splits(0)._ColumnProps(109)=   "Column(14).Visible=0"
            Splits(0)._ColumnProps(110)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(111)=   "Column(15).Width=1561"
            Splits(0)._ColumnProps(112)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(113)=   "Column(15)._WidthInPix=1482"
            Splits(0)._ColumnProps(114)=   "Column(15)._EditAlways=0"
            Splits(0)._ColumnProps(115)=   "Column(15).AllowSizing=0"
            Splits(0)._ColumnProps(116)=   "Column(15)._ColStyle=516"
            Splits(0)._ColumnProps(117)=   "Column(15).Visible=0"
            Splits(0)._ColumnProps(118)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(119)=   "Column(16).Width=4948"
            Splits(0)._ColumnProps(120)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(121)=   "Column(16)._WidthInPix=4868"
            Splits(0)._ColumnProps(122)=   "Column(16)._EditAlways=0"
            Splits(0)._ColumnProps(123)=   "Column(16).AllowSizing=0"
            Splits(0)._ColumnProps(124)=   "Column(16)._ColStyle=516"
            Splits(0)._ColumnProps(125)=   "Column(16).Visible=0"
            Splits(0)._ColumnProps(126)=   "Column(16).Order=17"
            Splits(1)._UserFlags=   0
            Splits(1).Locked=   -1  'True
            Splits(1).MarqueeStyle=   3
            Splits(1).SizeMode=   1
            Splits(1).Size  =   4500.284
            Splits(1).Size.vt=   4
            Splits(1).RecordSelectors=   0   'False
            Splits(1).RecordSelectorWidth=   503
            Splits(1)._SavedRecordSelectors=   0   'False
            Splits(1).Caption=   "Cuenta x Rendir"
            Splits(1).DividerColor=   12632256
            Splits(1).SpringMode=   0   'False
            Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(1)._ColumnProps(0)=   "Columns.Count=17"
            Splits(1)._ColumnProps(1)=   "Column(0).Width=1191"
            Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
            Splits(1)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(1)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(1)._ColumnProps(6)=   "Column(0)._ColStyle=516"
            Splits(1)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
            Splits(1)._ColumnProps(9)=   "Column(1).Width=1746"
            Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
            Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=1667"
            Splits(1)._ColumnProps(12)=   "Column(1)._EditAlways=0"
            Splits(1)._ColumnProps(13)=   "Column(1).AllowSizing=0"
            Splits(1)._ColumnProps(14)=   "Column(1)._ColStyle=516"
            Splits(1)._ColumnProps(15)=   "Column(1).Visible=0"
            Splits(1)._ColumnProps(16)=   "Column(1).Order=2"
            Splits(1)._ColumnProps(17)=   "Column(2).Width=1535"
            Splits(1)._ColumnProps(18)=   "Column(2).DividerColor=0"
            Splits(1)._ColumnProps(19)=   "Column(2)._WidthInPix=1455"
            Splits(1)._ColumnProps(20)=   "Column(2)._EditAlways=0"
            Splits(1)._ColumnProps(21)=   "Column(2).AllowSizing=0"
            Splits(1)._ColumnProps(22)=   "Column(2)._ColStyle=516"
            Splits(1)._ColumnProps(23)=   "Column(2).Visible=0"
            Splits(1)._ColumnProps(24)=   "Column(2).Order=3"
            Splits(1)._ColumnProps(25)=   "Column(3).Width=2725"
            Splits(1)._ColumnProps(26)=   "Column(3).DividerColor=0"
            Splits(1)._ColumnProps(27)=   "Column(3)._WidthInPix=2646"
            Splits(1)._ColumnProps(28)=   "Column(3)._EditAlways=0"
            Splits(1)._ColumnProps(29)=   "Column(3).AllowSizing=0"
            Splits(1)._ColumnProps(30)=   "Column(3)._ColStyle=516"
            Splits(1)._ColumnProps(31)=   "Column(3).Visible=0"
            Splits(1)._ColumnProps(32)=   "Column(3).Order=4"
            Splits(1)._ColumnProps(33)=   "Column(4).Width=2725"
            Splits(1)._ColumnProps(34)=   "Column(4).DividerColor=0"
            Splits(1)._ColumnProps(35)=   "Column(4)._WidthInPix=2646"
            Splits(1)._ColumnProps(36)=   "Column(4)._EditAlways=0"
            Splits(1)._ColumnProps(37)=   "Column(4).AllowSizing=0"
            Splits(1)._ColumnProps(38)=   "Column(4)._ColStyle=516"
            Splits(1)._ColumnProps(39)=   "Column(4).Visible=0"
            Splits(1)._ColumnProps(40)=   "Column(4).Order=5"
            Splits(1)._ColumnProps(41)=   "Column(5).Width=1535"
            Splits(1)._ColumnProps(42)=   "Column(5).DividerColor=0"
            Splits(1)._ColumnProps(43)=   "Column(5)._WidthInPix=1455"
            Splits(1)._ColumnProps(44)=   "Column(5)._EditAlways=0"
            Splits(1)._ColumnProps(45)=   "Column(5)._ColStyle=513"
            Splits(1)._ColumnProps(46)=   "Column(5).Order=6"
            Splits(1)._ColumnProps(47)=   "Column(6).Width=1535"
            Splits(1)._ColumnProps(48)=   "Column(6).DividerColor=0"
            Splits(1)._ColumnProps(49)=   "Column(6)._WidthInPix=1455"
            Splits(1)._ColumnProps(50)=   "Column(6)._EditAlways=0"
            Splits(1)._ColumnProps(51)=   "Column(6)._ColStyle=513"
            Splits(1)._ColumnProps(52)=   "Column(6).Order=7"
            Splits(1)._ColumnProps(53)=   "Column(7).Width=1720"
            Splits(1)._ColumnProps(54)=   "Column(7).DividerColor=0"
            Splits(1)._ColumnProps(55)=   "Column(7)._WidthInPix=1640"
            Splits(1)._ColumnProps(56)=   "Column(7)._EditAlways=0"
            Splits(1)._ColumnProps(57)=   "Column(7)._ColStyle=513"
            Splits(1)._ColumnProps(58)=   "Column(7).Order=8"
            Splits(1)._ColumnProps(59)=   "Column(8).Width=1826"
            Splits(1)._ColumnProps(60)=   "Column(8).DividerColor=0"
            Splits(1)._ColumnProps(61)=   "Column(8)._WidthInPix=1746"
            Splits(1)._ColumnProps(62)=   "Column(8)._EditAlways=0"
            Splits(1)._ColumnProps(63)=   "Column(8)._ColStyle=513"
            Splits(1)._ColumnProps(64)=   "Column(8).Order=9"
            Splits(1)._ColumnProps(65)=   "Column(9).Width=767"
            Splits(1)._ColumnProps(66)=   "Column(9).DividerColor=0"
            Splits(1)._ColumnProps(67)=   "Column(9)._WidthInPix=688"
            Splits(1)._ColumnProps(68)=   "Column(9)._EditAlways=0"
            Splits(1)._ColumnProps(69)=   "Column(9)._ColStyle=516"
            Splits(1)._ColumnProps(70)=   "Column(9).Order=10"
            Splits(1)._ColumnProps(71)=   "Column(10).Width=1535"
            Splits(1)._ColumnProps(72)=   "Column(10).DividerColor=0"
            Splits(1)._ColumnProps(73)=   "Column(10)._WidthInPix=1455"
            Splits(1)._ColumnProps(74)=   "Column(10)._EditAlways=0"
            Splits(1)._ColumnProps(75)=   "Column(10)._ColStyle=514"
            Splits(1)._ColumnProps(76)=   "Column(10).Order=11"
            Splits(1)._ColumnProps(77)=   "Column(11).Width=1482"
            Splits(1)._ColumnProps(78)=   "Column(11).DividerColor=0"
            Splits(1)._ColumnProps(79)=   "Column(11)._WidthInPix=1402"
            Splits(1)._ColumnProps(80)=   "Column(11)._EditAlways=0"
            Splits(1)._ColumnProps(81)=   "Column(11)._ColStyle=516"
            Splits(1)._ColumnProps(82)=   "Column(11).Order=12"
            Splits(1)._ColumnProps(83)=   "Column(12).Width=5900"
            Splits(1)._ColumnProps(84)=   "Column(12).DividerColor=0"
            Splits(1)._ColumnProps(85)=   "Column(12)._WidthInPix=5821"
            Splits(1)._ColumnProps(86)=   "Column(12)._EditAlways=0"
            Splits(1)._ColumnProps(87)=   "Column(12)._ColStyle=516"
            Splits(1)._ColumnProps(88)=   "Column(12).Order=13"
            Splits(1)._ColumnProps(89)=   "Column(13).Width=2725"
            Splits(1)._ColumnProps(90)=   "Column(13).DividerColor=0"
            Splits(1)._ColumnProps(91)=   "Column(13)._WidthInPix=2646"
            Splits(1)._ColumnProps(92)=   "Column(13)._EditAlways=0"
            Splits(1)._ColumnProps(93)=   "Column(13)._ColStyle=516"
            Splits(1)._ColumnProps(94)=   "Column(13).Order=14"
            Splits(1)._ColumnProps(95)=   "Column(14).Width=4075"
            Splits(1)._ColumnProps(96)=   "Column(14).DividerColor=0"
            Splits(1)._ColumnProps(97)=   "Column(14)._WidthInPix=3995"
            Splits(1)._ColumnProps(98)=   "Column(14)._EditAlways=0"
            Splits(1)._ColumnProps(99)=   "Column(14)._ColStyle=516"
            Splits(1)._ColumnProps(100)=   "Column(14).Order=15"
            Splits(1)._ColumnProps(101)=   "Column(15).Width=1561"
            Splits(1)._ColumnProps(102)=   "Column(15).DividerColor=0"
            Splits(1)._ColumnProps(103)=   "Column(15)._WidthInPix=1482"
            Splits(1)._ColumnProps(104)=   "Column(15)._EditAlways=0"
            Splits(1)._ColumnProps(105)=   "Column(15)._ColStyle=516"
            Splits(1)._ColumnProps(106)=   "Column(15).Order=16"
            Splits(1)._ColumnProps(107)=   "Column(16).Width=4948"
            Splits(1)._ColumnProps(108)=   "Column(16).DividerColor=0"
            Splits(1)._ColumnProps(109)=   "Column(16)._WidthInPix=4868"
            Splits(1)._ColumnProps(110)=   "Column(16)._EditAlways=0"
            Splits(1)._ColumnProps(111)=   "Column(16)._ColStyle=516"
            Splits(1)._ColumnProps(112)=   "Column(16).Order=17"
            Splits.Count    =   2
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
            _StyleDefs(24)  =   "Splits(0).Style:id=95,.parent=1,.bgcolor=&HDBFDFD&,.fgcolor=&H800000&"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=104,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=96,.parent=2,.fgcolor=&H800000&"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=97,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=98,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=100,.parent=6,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=99,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=101,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=102,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=103,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=105,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=106,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=110,.parent=95,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=107,.parent=96"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=108,.parent=97"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=109,.parent=99"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=114,.parent=95,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=111,.parent=96"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=112,.parent=97"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=113,.parent=99"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=118,.parent=95,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=115,.parent=96"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=116,.parent=97"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=117,.parent=99"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=78,.parent=95,.alignment=1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=96"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=97"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=99"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=170,.parent=95,.alignment=1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=167,.parent=96"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=168,.parent=97"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=169,.parent=99"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=126,.parent=95,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=123,.parent=96"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=124,.parent=97"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=125,.parent=99"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=130,.parent=95,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=127,.parent=96"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=128,.parent=97"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=129,.parent=99"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=134,.parent=95,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=131,.parent=96"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=132,.parent=97"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=133,.parent=99"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=178,.parent=95"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=175,.parent=96"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=176,.parent=97"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=177,.parent=99"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=138,.parent=95"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=135,.parent=96"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=136,.parent=97"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=137,.parent=99"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=142,.parent=95,.alignment=1"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=139,.parent=96"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=140,.parent=97"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=141,.parent=99"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=146,.parent=95"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=143,.parent=96"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=144,.parent=97"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=145,.parent=99"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=150,.parent=95"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=147,.parent=96"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=148,.parent=97"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=149,.parent=99"
            _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=154,.parent=95"
            _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=151,.parent=96"
            _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=152,.parent=97"
            _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=153,.parent=99"
            _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=158,.parent=95"
            _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=155,.parent=96"
            _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=156,.parent=97"
            _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=157,.parent=99"
            _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=162,.parent=95"
            _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=159,.parent=96"
            _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=160,.parent=97"
            _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=161,.parent=99"
            _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=166,.parent=95"
            _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=163,.parent=96"
            _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=164,.parent=97"
            _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=165,.parent=99"
            _StyleDefs(104) =   "Splits(1).Style:id=13,.parent=1"
            _StyleDefs(105) =   "Splits(1).CaptionStyle:id=22,.parent=4"
            _StyleDefs(106) =   "Splits(1).HeadingStyle:id=14,.parent=2"
            _StyleDefs(107) =   "Splits(1).FooterStyle:id=15,.parent=3"
            _StyleDefs(108) =   "Splits(1).InactiveStyle:id=16,.parent=5"
            _StyleDefs(109) =   "Splits(1).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(110) =   "Splits(1).EditorStyle:id=17,.parent=7"
            _StyleDefs(111) =   "Splits(1).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(112) =   "Splits(1).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(113) =   "Splits(1).OddRowStyle:id=21,.parent=10"
            _StyleDefs(114) =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(115) =   "Splits(1).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(116) =   "Splits(1).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(117) =   "Splits(1).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(118) =   "Splits(1).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(119) =   "Splits(1).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(120) =   "Splits(1).Columns(1).Style:id=82,.parent=13"
            _StyleDefs(121) =   "Splits(1).Columns(1).HeadingStyle:id=79,.parent=14"
            _StyleDefs(122) =   "Splits(1).Columns(1).FooterStyle:id=80,.parent=15"
            _StyleDefs(123) =   "Splits(1).Columns(1).EditorStyle:id=81,.parent=17"
            _StyleDefs(124) =   "Splits(1).Columns(2).Style:id=94,.parent=13"
            _StyleDefs(125) =   "Splits(1).Columns(2).HeadingStyle:id=91,.parent=14"
            _StyleDefs(126) =   "Splits(1).Columns(2).FooterStyle:id=92,.parent=15"
            _StyleDefs(127) =   "Splits(1).Columns(2).EditorStyle:id=93,.parent=17"
            _StyleDefs(128) =   "Splits(1).Columns(3).Style:id=122,.parent=13"
            _StyleDefs(129) =   "Splits(1).Columns(3).HeadingStyle:id=119,.parent=14"
            _StyleDefs(130) =   "Splits(1).Columns(3).FooterStyle:id=120,.parent=15"
            _StyleDefs(131) =   "Splits(1).Columns(3).EditorStyle:id=121,.parent=17"
            _StyleDefs(132) =   "Splits(1).Columns(4).Style:id=174,.parent=13"
            _StyleDefs(133) =   "Splits(1).Columns(4).HeadingStyle:id=171,.parent=14"
            _StyleDefs(134) =   "Splits(1).Columns(4).FooterStyle:id=172,.parent=15"
            _StyleDefs(135) =   "Splits(1).Columns(4).EditorStyle:id=173,.parent=17"
            _StyleDefs(136) =   "Splits(1).Columns(5).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(137) =   "Splits(1).Columns(5).HeadingStyle:id=29,.parent=14"
            _StyleDefs(138) =   "Splits(1).Columns(5).FooterStyle:id=30,.parent=15"
            _StyleDefs(139) =   "Splits(1).Columns(5).EditorStyle:id=31,.parent=17"
            _StyleDefs(140) =   "Splits(1).Columns(6).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(141) =   "Splits(1).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(142) =   "Splits(1).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(143) =   "Splits(1).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(144) =   "Splits(1).Columns(7).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(145) =   "Splits(1).Columns(7).HeadingStyle:id=43,.parent=14"
            _StyleDefs(146) =   "Splits(1).Columns(7).FooterStyle:id=44,.parent=15"
            _StyleDefs(147) =   "Splits(1).Columns(7).EditorStyle:id=45,.parent=17"
            _StyleDefs(148) =   "Splits(1).Columns(8).Style:id=182,.parent=13,.alignment=2"
            _StyleDefs(149) =   "Splits(1).Columns(8).HeadingStyle:id=179,.parent=14"
            _StyleDefs(150) =   "Splits(1).Columns(8).FooterStyle:id=180,.parent=15"
            _StyleDefs(151) =   "Splits(1).Columns(8).EditorStyle:id=181,.parent=17"
            _StyleDefs(152) =   "Splits(1).Columns(9).Style:id=58,.parent=13"
            _StyleDefs(153) =   "Splits(1).Columns(9).HeadingStyle:id=55,.parent=14"
            _StyleDefs(154) =   "Splits(1).Columns(9).FooterStyle:id=56,.parent=15"
            _StyleDefs(155) =   "Splits(1).Columns(9).EditorStyle:id=57,.parent=17"
            _StyleDefs(156) =   "Splits(1).Columns(10).Style:id=90,.parent=13,.alignment=1"
            _StyleDefs(157) =   "Splits(1).Columns(10).HeadingStyle:id=87,.parent=14"
            _StyleDefs(158) =   "Splits(1).Columns(10).FooterStyle:id=88,.parent=15"
            _StyleDefs(159) =   "Splits(1).Columns(10).EditorStyle:id=89,.parent=17"
            _StyleDefs(160) =   "Splits(1).Columns(11).Style:id=54,.parent=13"
            _StyleDefs(161) =   "Splits(1).Columns(11).HeadingStyle:id=51,.parent=14"
            _StyleDefs(162) =   "Splits(1).Columns(11).FooterStyle:id=52,.parent=15"
            _StyleDefs(163) =   "Splits(1).Columns(11).EditorStyle:id=53,.parent=17"
            _StyleDefs(164) =   "Splits(1).Columns(12).Style:id=66,.parent=13"
            _StyleDefs(165) =   "Splits(1).Columns(12).HeadingStyle:id=63,.parent=14"
            _StyleDefs(166) =   "Splits(1).Columns(12).FooterStyle:id=64,.parent=15"
            _StyleDefs(167) =   "Splits(1).Columns(12).EditorStyle:id=65,.parent=17"
            _StyleDefs(168) =   "Splits(1).Columns(13).Style:id=86,.parent=13"
            _StyleDefs(169) =   "Splits(1).Columns(13).HeadingStyle:id=83,.parent=14"
            _StyleDefs(170) =   "Splits(1).Columns(13).FooterStyle:id=84,.parent=15"
            _StyleDefs(171) =   "Splits(1).Columns(13).EditorStyle:id=85,.parent=17"
            _StyleDefs(172) =   "Splits(1).Columns(14).Style:id=62,.parent=13"
            _StyleDefs(173) =   "Splits(1).Columns(14).HeadingStyle:id=59,.parent=14"
            _StyleDefs(174) =   "Splits(1).Columns(14).FooterStyle:id=60,.parent=15"
            _StyleDefs(175) =   "Splits(1).Columns(14).EditorStyle:id=61,.parent=17"
            _StyleDefs(176) =   "Splits(1).Columns(15).Style:id=70,.parent=13"
            _StyleDefs(177) =   "Splits(1).Columns(15).HeadingStyle:id=67,.parent=14"
            _StyleDefs(178) =   "Splits(1).Columns(15).FooterStyle:id=68,.parent=15"
            _StyleDefs(179) =   "Splits(1).Columns(15).EditorStyle:id=69,.parent=17"
            _StyleDefs(180) =   "Splits(1).Columns(16).Style:id=74,.parent=13"
            _StyleDefs(181) =   "Splits(1).Columns(16).HeadingStyle:id=71,.parent=14"
            _StyleDefs(182) =   "Splits(1).Columns(16).FooterStyle:id=72,.parent=15"
            _StyleDefs(183) =   "Splits(1).Columns(16).EditorStyle:id=73,.parent=17"
            _StyleDefs(184) =   "Named:id=33:Normal"
            _StyleDefs(185) =   ":id=33,.parent=0"
            _StyleDefs(186) =   "Named:id=34:Heading"
            _StyleDefs(187) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(188) =   ":id=34,.wraptext=-1"
            _StyleDefs(189) =   "Named:id=35:Footing"
            _StyleDefs(190) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(191) =   "Named:id=36:Selected"
            _StyleDefs(192) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(193) =   "Named:id=37:Caption"
            _StyleDefs(194) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(195) =   "Named:id=38:HighlightRow"
            _StyleDefs(196) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(197) =   "Named:id=39:EvenRow"
            _StyleDefs(198) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(199) =   "Named:id=40:OddRow"
            _StyleDefs(200) =   ":id=40,.parent=33"
            _StyleDefs(201) =   "Named:id=41:RecordSelector"
            _StyleDefs(202) =   ":id=41,.parent=34"
            _StyleDefs(203) =   "Named:id=42:FilterBar"
            _StyleDefs(204) =   ":id=42,.parent=33"
         End
         Begin VB.Label lblperiodo 
            Caption         =   "lblperiodo(0)"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   0
            Left            =   9555
            TabIndex        =   64
            Top             =   15
            Width           =   1980
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Devolución de Cuentas por Rendir"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   15
            TabIndex        =   10
            Top             =   30
            Width           =   11550
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6810
         Left            =   45
         TabIndex        =   11
         Top             =   375
         Width           =   11670
         Begin VB.Frame Frame6 
            Caption         =   "[ Documento ]"
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
            Height          =   615
            Left            =   165
            TabIndex        =   81
            Top             =   4785
            Width           =   7980
            Begin VB.TextBox txt 
               ForeColor       =   &H00000000&
               Height          =   300
               Index           =   3
               Left            =   6630
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   4
               Text            =   "txt(3)"
               Top             =   240
               Width           =   1230
            End
            Begin VB.CommandButton cb 
               Enabled         =   0   'False
               Height          =   225
               Index           =   2
               Left            =   690
               Picture         =   "FrmCuentasRendir_Dev.frx":277E
               Style           =   1  'Graphical
               TabIndex        =   82
               Top             =   270
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   2
               Left            =   180
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   3
               Text            =   "txt_cb(2)"
               Top             =   240
               Width           =   750
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "N°. Documento"
               Height          =   195
               Index           =   3
               Left            =   5490
               TabIndex        =   84
               Top             =   345
               Width           =   1095
            End
            Begin VB.Label lbl_cb_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb_cod(2)"
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
               Index           =   2
               Left            =   4125
               TabIndex        =   83
               Top             =   270
               Visible         =   0   'False
               Width           =   1230
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(2)"
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
               Index           =   2
               Left            =   930
               TabIndex        =   85
               Top             =   240
               Width           =   4500
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "[ Observación]"
            Height          =   1950
            Left            =   8190
            TabIndex        =   79
            Top             =   4785
            Width           =   3390
            Begin VB.TextBox txt 
               Height          =   1605
               Index           =   2
               Left            =   75
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   8
               Tag             =   "null"
               Text            =   "FrmCuentasRendir_Dev.frx":28B0
               Top             =   225
               Width           =   3210
            End
         End
         Begin VB.Frame fra_datos 
            Caption         =   "[ Restringido ]"
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
            Height          =   1305
            Left            =   165
            TabIndex        =   67
            Top             =   5445
            Width           =   7980
            Begin VB.Frame Frame5 
               Caption         =   "[ Moneda ]"
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
               Height          =   615
               Left            =   2820
               TabIndex        =   74
               Top             =   255
               Width           =   4200
               Begin VB.CommandButton cb 
                  Enabled         =   0   'False
                  Height          =   225
                  Index           =   0
                  Left            =   495
                  Picture         =   "FrmCuentasRendir_Dev.frx":28B9
                  Style           =   1  'Graphical
                  TabIndex        =   75
                  Top             =   270
                  Width           =   210
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   0
                  Left            =   75
                  MaxLength       =   3
                  TabIndex        =   6
                  Text            =   "txt_cb(0)"
                  Top             =   240
                  Width           =   675
               End
               Begin VB.Label lbl_cb_capt 
                  AutoSize        =   -1  'True
                  Height          =   195
                  Index           =   0
                  Left            =   90
                  TabIndex        =   77
                  Top             =   345
                  Width           =   465
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
                  Left            =   2610
                  TabIndex        =   76
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1230
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
                  Left            =   735
                  TabIndex        =   78
                  Top             =   240
                  Width           =   3240
               End
            End
            Begin VB.Frame fr 
               Caption         =   "[ Tipo de Operación ]"
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
               Height          =   615
               Index           =   5
               Left            =   150
               TabIndex        =   69
               Top             =   255
               Width           =   2610
               Begin VB.OptionButton opt_operacion 
                  Caption         =   "Banco"
                  Enabled         =   0   'False
                  Height          =   195
                  Index           =   1
                  Left            =   1485
                  TabIndex        =   70
                  Top             =   300
                  Width           =   840
               End
               Begin VB.OptionButton opt_operacion 
                  Caption         =   "Caja"
                  Enabled         =   0   'False
                  Height          =   195
                  Index           =   0
                  Left            =   300
                  TabIndex        =   5
                  Top             =   300
                  Value           =   -1  'True
                  Width           =   840
               End
            End
            Begin VB.CommandButton cb 
               Enabled         =   0   'False
               Height          =   225
               Index           =   1
               Left            =   2040
               Picture         =   "FrmCuentasRendir_Dev.frx":29EB
               Style           =   1  'Graphical
               TabIndex        =   68
               Top             =   930
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   1
               Left            =   1545
               Locked          =   -1  'True
               MaxLength       =   4
               TabIndex        =   7
               Text            =   "txt_cb(1)"
               Top             =   900
               Width           =   750
            End
            Begin VB.Label lblCtaDevolucion 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblCtaDevolucion"
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
               Left            =   6120
               TabIndex        =   80
               Top             =   900
               Visible         =   0   'False
               Width           =   1710
            End
            Begin VB.Label lbl_cb_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb_cod(1)"
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
               Index           =   1
               Left            =   4035
               TabIndex        =   72
               Top             =   900
               Visible         =   0   'False
               Width           =   1230
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Destino de Ingreso"
               Height          =   300
               Index           =   1
               Left            =   165
               TabIndex        =   71
               Top             =   975
               Width           =   1335
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(1)"
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
               Index           =   1
               Left            =   2295
               TabIndex        =   73
               Top             =   900
               Width           =   4770
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "( Periodo )"
            Height          =   720
            Left            =   9570
            TabIndex        =   65
            Top             =   3975
            Width           =   2010
            Begin VB.Label lblperiodo 
               Alignment       =   2  'Center
               Caption         =   "lblperiodo(1)"
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   66
               Top             =   330
               Width           =   1740
            End
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   3870
            Left            =   90
            TabIndex        =   18
            Top             =   345
            Width           =   11490
            _cx             =   20267
            _cy             =   6826
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FrontTabColor   =   -2147483633
            BackTabColor    =   -2147483633
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   8388608
            Caption         =   "[Información de la Cuenta por Rendir]|           [Información Declarada]           "
            Align           =   0
            CurrTab         =   1
            FirstTab        =   0
            Style           =   0
            Position        =   1
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
            Begin VB.Frame fr 
               BorderStyle     =   0  'None
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
               Height          =   3495
               Index           =   0
               Left            =   45
               TabIndex        =   39
               Top             =   45
               Width           =   11400
               Begin VB.CommandButton cmd 
                  Caption         =   "&Eliminar"
                  Enabled         =   0   'False
                  Height          =   405
                  Index           =   2
                  Left            =   2850
                  TabIndex        =   58
                  Top             =   3030
                  Width           =   1125
               End
               Begin VB.CommandButton cmd 
                  Caption         =   "Seleccionar"
                  Enabled         =   0   'False
                  Height          =   405
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   56
                  Top             =   3030
                  Width           =   1125
               End
               Begin VB.CommandButton cmd 
                  Caption         =   "Agregar"
                  Enabled         =   0   'False
                  Height          =   405
                  Index           =   0
                  Left            =   120
                  TabIndex        =   55
                  Top             =   3030
                  Width           =   1125
               End
               Begin VB.TextBox txt_total 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000040C0&
                  Height          =   315
                  Index           =   2
                  Left            =   10170
                  Locked          =   -1  'True
                  TabIndex        =   42
                  Tag             =   "null"
                  Text            =   "txt_total(2)"
                  Top             =   3045
                  Width           =   1095
               End
               Begin VB.TextBox txt_total 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00808000&
                  Height          =   315
                  Index           =   1
                  Left            =   8235
                  Locked          =   -1  'True
                  TabIndex        =   41
                  Tag             =   "null"
                  Text            =   "txt_total(1)"
                  Top             =   3045
                  Width           =   1095
               End
               Begin VB.TextBox txt_total 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   0
                  Left            =   5970
                  Locked          =   -1  'True
                  TabIndex        =   40
                  Tag             =   "null"
                  Text            =   "txt_total(0)"
                  Top             =   3045
                  Width           =   1095
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg1 
                  Height          =   2760
                  Left            =   120
                  TabIndex        =   59
                  Top             =   120
                  Width           =   11160
                  _cx             =   19685
                  _cy             =   4868
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   13
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmCuentasRendir_Dev.frx":2B1D
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
                  ShowComboButton =   -1  'True
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
               Begin VB.CommandButton cmd 
                  Caption         =   "Nuevo"
                  Enabled         =   0   'False
                  Height          =   405
                  Index           =   3
                  Left            =   2310
                  TabIndex        =   57
                  Top             =   2595
                  Visible         =   0   'False
                  Width           =   1125
               End
               Begin VB.Label lbl_total 
                  AutoSize        =   -1  'True
                  Caption         =   "Por Rendir"
                  Height          =   195
                  Index           =   2
                  Left            =   9390
                  TabIndex        =   45
                  Top             =   3165
                  Width           =   750
               End
               Begin VB.Label lbl_total 
                  AutoSize        =   -1  'True
                  Caption         =   "Tot. Declarado"
                  Height          =   195
                  Index           =   1
                  Left            =   7125
                  TabIndex        =   44
                  Top             =   3165
                  Width           =   1065
               End
               Begin VB.Label lbl_total 
                  AutoSize        =   -1  'True
                  Caption         =   "Tot. Entregado"
                  Height          =   195
                  Index           =   0
                  Left            =   4860
                  TabIndex        =   43
                  Top             =   3165
                  Width           =   1065
               End
            End
            Begin VB.Frame fr 
               BorderStyle     =   0  'None
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
               Height          =   3495
               Index           =   1
               Left            =   -12045
               TabIndex        =   19
               Top             =   45
               Width           =   11400
               Begin VB.Frame fr 
                  Caption         =   "[ Del Beneficiario ]"
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
                  Height          =   720
                  Index           =   2
                  Left            =   3345
                  TabIndex        =   27
                  Top             =   2700
                  Width           =   8010
                  Begin VB.Label lbl_dato 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "lbl_dato(9)"
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
                     Index           =   9
                     Left            =   1605
                     TabIndex        =   29
                     Top             =   285
                     Width           =   5235
                  End
                  Begin VB.Label lbl_dato_x 
                     AutoSize        =   -1  'True
                     Caption         =   "Persona"
                     Height          =   195
                     Index           =   9
                     Left            =   180
                     TabIndex        =   28
                     Top             =   330
                     Width           =   1365
                  End
               End
               Begin VB.Frame fr 
                  Caption         =   "[ Del Destino ]"
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
                  Height          =   720
                  Index           =   3
                  Left            =   3345
                  TabIndex        =   24
                  Top             =   1680
                  Width           =   8010
                  Begin VB.Label lblCtaRendir 
                     BackColor       =   &H0000FF00&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "lblCtaRendir"
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
                     Left            =   6075
                     TabIndex        =   60
                     Top             =   270
                     Visible         =   0   'False
                     Width           =   1290
                  End
                  Begin VB.Label lbl_dato 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "lbl_dato(8)"
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
                     Index           =   8
                     Left            =   1605
                     TabIndex        =   26
                     Top             =   285
                     Width           =   5235
                  End
                  Begin VB.Label lbl_dato_x 
                     AutoSize        =   -1  'True
                     Caption         =   "Destino"
                     Height          =   195
                     Index           =   8
                     Left            =   180
                     TabIndex        =   25
                     Top             =   330
                     Width           =   540
                  End
               End
               Begin VB.Frame fr 
                  Caption         =   "[ Del Origen ]"
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
                  Height          =   720
                  Index           =   4
                  Left            =   3345
                  TabIndex        =   21
                  Top             =   660
                  Width           =   8010
                  Begin VB.Label lbl_dato 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "lbl_dato(7)"
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
                     Index           =   7
                     Left            =   1605
                     TabIndex        =   23
                     Top             =   285
                     Width           =   5235
                  End
                  Begin VB.Label lbl_dato_x 
                     AutoSize        =   -1  'True
                     Caption         =   "Caja"
                     Height          =   195
                     Index           =   7
                     Left            =   180
                     TabIndex        =   22
                     Top             =   330
                     Width           =   315
                  End
               End
               Begin VB.CommandButton cb_rendir 
                  Height          =   240
                  Left            =   2625
                  Picture         =   "FrmCuentasRendir_Dev.frx":2C99
                  Style           =   1  'Graphical
                  TabIndex        =   20
                  Top             =   255
                  Width           =   240
               End
               Begin VB.TextBox txt_rendir 
                  Height          =   300
                  Left            =   1725
                  MaxLength       =   20
                  TabIndex        =   30
                  Text            =   "txt_cb(0)"
                  Top             =   225
                  Width           =   1170
               End
               Begin VB.Label lblIdMoneda 
                  BackColor       =   &H0000FF00&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lblIdMoneda"
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
                  Left            =   10020
                  TabIndex        =   61
                  Top             =   420
                  Visible         =   0   'False
                  Width           =   1275
               End
               Begin VB.Label lbl_dato_x 
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha Rend. Cta."
                  Height          =   195
                  Index           =   5
                  Left            =   210
                  TabIndex        =   54
                  Top             =   1605
                  Width           =   1260
               End
               Begin VB.Label lbl_dato 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_dato(10)"
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
                  Index           =   10
                  Left            =   1725
                  TabIndex        =   53
                  Top             =   1560
                  Width           =   1515
               End
               Begin VB.Label lbl_dato 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_dato(1)"
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
                  Index           =   1
                  Left            =   1725
                  TabIndex        =   51
                  Top             =   780
                  Width           =   1515
               End
               Begin VB.Label lbl_dato_x 
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha Emisión"
                  Height          =   195
                  Index           =   1
                  Left            =   210
                  TabIndex        =   50
                  Top             =   825
                  Width           =   1035
               End
               Begin VB.Label lbl_dato 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_dato(6)"
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
                  Index           =   6
                  Left            =   1725
                  TabIndex        =   49
                  Top             =   3120
                  Width           =   1515
               End
               Begin VB.Label lbl_dato 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_dato(5)"
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
                  Index           =   5
                  Left            =   1725
                  TabIndex        =   48
                  Top             =   2730
                  Width           =   1515
               End
               Begin VB.Label lbl_dato 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_dato(3)"
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
                  Index           =   3
                  Left            =   1725
                  TabIndex        =   47
                  Top             =   2340
                  Width           =   1515
               End
               Begin VB.Label lbl_dato 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_dato(2)"
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
                  Index           =   2
                  Left            =   1725
                  TabIndex        =   46
                  Top             =   1170
                  Width           =   1515
               End
               Begin VB.Label lbl_dato 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_dato(4)"
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
                  Index           =   4
                  Left            =   1725
                  TabIndex        =   38
                  Top             =   1950
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "Importe"
                  Height          =   195
                  Index           =   6
                  Left            =   210
                  TabIndex        =   37
                  Top             =   3165
                  Width           =   525
               End
               Begin VB.Label lbl_dato_x 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo de Documento"
                  Height          =   195
                  Index           =   3
                  Left            =   210
                  TabIndex        =   36
                  Top             =   2385
                  Width           =   1410
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "N° de Documento"
                  Height          =   195
                  Index           =   5
                  Left            =   210
                  TabIndex        =   35
                  Top             =   2775
                  Width           =   1275
               End
               Begin VB.Label lbl_dato_x 
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha Pago"
                  Height          =   195
                  Index           =   2
                  Left            =   210
                  TabIndex        =   34
                  Top             =   1215
                  Width           =   870
               End
               Begin VB.Label lbl_dato_x 
                  AutoSize        =   -1  'True
                  Caption         =   "Moneda"
                  Height          =   195
                  Index           =   4
                  Left            =   210
                  TabIndex        =   33
                  Top             =   1995
                  Width           =   585
               End
               Begin VB.Label lbl_rendir 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_rendir"
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
                  Left            =   2895
                  TabIndex        =   32
                  Top             =   225
                  Visible         =   0   'False
                  Width           =   1680
               End
               Begin VB.Label label10 
                  AutoSize        =   -1  'True
                  Caption         =   "Cta por Rendir"
                  Height          =   195
                  Left            =   210
                  TabIndex        =   31
                  Top             =   270
                  Width           =   1020
               End
            End
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
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
            Index           =   1
            Left            =   4740
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "txt(1)"
            Top             =   4305
            Width           =   1260
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0000FFFF&
            Height          =   315
            Index           =   0
            Left            =   10320
            TabIndex        =   15
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   0
            Visible         =   0   'False
            Width           =   1170
         End
         Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
            Height          =   300
            Index           =   0
            Left            =   1575
            TabIndex        =   1
            Tag             =   "b"
            Top             =   4305
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
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   7515
            TabIndex        =   63
            Top             =   4305
            Width           =   1350
         End
         Begin VB.Label LblTipCam2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cambio"
            Height          =   195
            Left            =   6330
            TabIndex        =   62
            Top             =   4410
            Width           =   1110
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Importe a Devolver"
            Height          =   195
            Index           =   1
            Left            =   3330
            TabIndex        =   17
            Top             =   4410
            Width           =   1350
         End
         Begin VB.Label lblfch 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Devolución"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   16
            Top             =   4410
            Width           =   1305
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "codigo"
            Height          =   195
            Index           =   0
            Left            =   9720
            TabIndex        =   14
            Top             =   120
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Devolución de Cuentas por Rendir"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   15
            Width           =   11550
         End
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar Compra"
      End
      Begin VB.Menu menu1_4 
         Caption         =   "Seleccionar Compra"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar Compra"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu menu2_1 
         Caption         =   "Eliminar Compra"
      End
      Begin VB.Menu menu2_2 
         Caption         =   "Eliminar Todo"
      End
   End
End
Attribute VB_Name = "FrmCuentasRendir_Dev"
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


Private Sub cb_rendir_Click()
    Dim xCampos() As String
    Dim nSQL As String
    If QueHace = 3 Then Exit Sub
    
'    On Error GoTo error
    
    ReDim xCampos(10, 3) As String
    xCampos(0, 0) = "Tip.Mov.":    xCampos(0, 1) = "tipmov":       xCampos(0, 2) = "800":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fch.Emi":     xCampos(1, 1) = "fchemi":       xCampos(1, 2) = "820":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "Fch.Pag.":    xCampos(2, 1) = "fchpag":       xCampos(2, 2) = "820":   xCampos(2, 3) = "C"
    xCampos(3, 0) = "Fch.Ren.":    xCampos(3, 1) = "fchren":       xCampos(3, 2) = "820":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "M":           xCampos(4, 1) = "simbolo":      xCampos(4, 2) = "450":   xCampos(4, 3) = "C"
    xCampos(5, 0) = "Importe":     xCampos(5, 1) = "imp":          xCampos(5, 2) = "700":  xCampos(5, 3) = "N"
    xCampos(6, 0) = "T.Persona":   xCampos(6, 1) = "tipper":       xCampos(6, 2) = "850":   xCampos(6, 3) = "C"
    xCampos(7, 0) = "Entregado A": xCampos(7, 1) = "benef":        xCampos(7, 2) = "1400":  xCampos(7, 3) = "C"
    xCampos(8, 0) = "T.Doc":       xCampos(8, 1) = "abrev":        xCampos(8, 2) = "600":  xCampos(8, 3) = "C"
    xCampos(9, 0) = "N° Doc":      xCampos(9, 1) = "numdoc":       xCampos(9, 2) = "1000":  xCampos(9, 3) = "C"
        
    nSQL = "SELECT con_ctasrendir.id, format(con_ctasrendir.fchemi,'dd/mm/yy') as fchemi, format(con_ctasrendir.fchpag,'dd/mm/yy') as fchpag, format(con_ctasrendir.fchren,'dd/mm/yy') as fchren, con_ctasrendir.numdoc, IIf(con_ctasrendir.tipmov=1,'Caja','Banco') AS tipmov,  mae_moneda.descripcion AS moneda, mae_moneda.simbolo, " _
        + vbCr + " IIf(con_ctasrendir.tipmov=1,(SELECT destino.descripcion FROM con_destino AS destino WHERE (((destino.id)=con_ctasrendir.idori)) ),(SELECT [mae_bancos].[descripcion] & '  Cta. N°: ' & [con_bancocuenta].[numcue] AS origen FROM mae_bancos INNER JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban WHERE (((con_bancocuenta.id)=con_ctasrendir.idori)) )) AS origen, " _
        + vbCr + " con_destino.descripcion AS destino, mae_doccajaban.descripcion AS tipdocnom, mae_doccajaban.abrev, " _
        + vbCr + " IIf(con_ctasrendir.tipper=1,'Persona','Proveedor') AS tipper, IIf(con_ctasrendir.tipper=1,(SELECT [pla_empleados].[nom] & ' ' & [pla_empleados].[ape] AS nombre FROM pla_empleados INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp WHERE (((con_emptes.id)=con_ctasrendir.idper))),(SELECT mae_prov.nombre FROM mae_prov WHERE (((mae_prov.id)=con_ctasrendir.idper)) )) AS benef, con_ctasrendir.[imp], " _
        + vbCr + " IIf((SELECT Sum(con_devolucionesdet.acuenta) AS SumaDeacuenta FROM con_devoluciones INNER JOIN con_devolucionesdet ON con_devoluciones.id = con_devolucionesdet.id GROUP BY con_devoluciones.idren HAVING  (((con_devoluciones.idren)=con_ctasrendir.id)) ) Is Null,0,(SELECT Sum(con_devolucionesdet.acuenta) AS SumaDeacuenta FROM con_devoluciones INNER JOIN con_devolucionesdet ON con_devoluciones.id = con_devolucionesdet.id GROUP BY con_devoluciones.idren HAVING (((con_devoluciones.idren)=con_ctasrendir.id)) )) AS declarado, " _
        + vbCr + " (con_ctasrendir.imp-declarado) AS xrendir  , con_destino.idcuen, con_ctasrendir.idmon " _
        + vbCr + " FROM mae_doccajaban RIGHT JOIN (mae_moneda RIGHT JOIN (con_ctasrendir LEFT JOIN con_destino ON con_ctasrendir.iddes = con_destino.id) ON mae_moneda.id = con_ctasrendir.idmon) ON mae_doccajaban.id = con_ctasrendir.tipdoc " _
        + vbCr + " WHERE (((con_ctasrendir.id) Not In (SELECT con_devoluciones.idren FROM con_devoluciones)) AND ((con_ctasrendir.[imp])>IIf((SELECT Sum(con_devolucionesdet.acuenta) AS SumaDeacuenta FROM con_devoluciones INNER JOIN con_devolucionesdet ON con_devoluciones.id = con_devolucionesdet.id GROUP BY con_devoluciones.idren HAVING  ((con_devoluciones.idren)=con_ctasrendir.id)) Is Null,0,(SELECT Sum(con_devolucionesdet.acuenta) AS SumaDeacuenta FROM con_devoluciones INNER JOIN con_devolucionesdet ON con_devoluciones.id = con_devolucionesdet.id GROUP BY   con_devoluciones.idren HAVING (((con_devoluciones.idren)=con_ctasrendir.id))  ))) AND ((con_ctasrendir.idest)=2));"


    Dim xRs As New ADODB.Recordset

    Me.MousePointer = vbHourglass
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Documentos x Rendir", "benef", "benef", Principio
    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    '----------
    LimpiaText TxtFecha
    LimpiaText lbl_dato, True
    LimpiaText txt_total, True
    TxtFecha(0).Valor = ""
    txt(1).Text = ""
    txt(2).Text = ""
    '------------
    pPonerDatosRendirCta xRs

Salir:
    Set xRs = Nothing
    Me.MousePointer = vbDefault
Exit Sub
error:
    Me.MousePointer = vbDefault
    Set xRs = Nothing
    SHOW_ERROR
End Sub

Private Sub pRegistroAdd(Optional fSeleccionVarios As Boolean = True)
    '--CARGAR LAS COMPRAS PARA LUEGO SELECCIONAR LOS QUE DESEEMOS
    '--SE CARGARAN DE ACUERDO A LA MONEDA DE CUENTAS POR RENDIR
    If lbl_rendir.Caption = "" Then
        MsgBox "Primero Seleccione la Cuenta a Rendir ", vbExclamation, xTitulo
        TabOne2.CurrTab = 0
        txt_rendir.SetFocus
        Exit Sub
    End If
    Dim nSQLIdCompra As String
    '--GENERAR EL WHERE DE LOS ID'S COMPRA PARA QUE NO SE REPITAN
    nSQLIdCompra = GENERAR_SQL_ID(Fg1, 1, "com_compras.id", " NOT IN ")
    If nSQLIdCompra <> "" Then nSQLIdCompra = " AND " + nSQLIdCompra
    '--DE LA MONEDA
    nSQLIdCompra = nSQLIdCompra + " AND com_compras.idmon=" & NulosN(LblIdMoneda.Caption) & " "
    '----
    On Error GoTo error
    Dim xRs  As New ADODB.Recordset
    Dim xCampos(8, 5) As String
    Dim nSQL As String
    
    xCampos(0, 0) = "Num.Reg.":      xCampos(0, 1) = "numreg":      xCampos(0, 2) = "900":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "abrev":        xCampos(1, 2) = "450":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "M":             xCampos(2, 1) = "simbolo":     xCampos(2, 2) = "450":      xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
    xCampos(3, 0) = "N°.Documento":  xCampos(3, 1) = "doc":         xCampos(3, 2) = "1400":      xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
    xCampos(4, 0) = "Fch.Emi.":      xCampos(4, 1) = "fchdoc":      xCampos(4, 2) = "900":      xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
    xCampos(5, 0) = "Proveedor":     xCampos(5, 1) = "nombre":      xCampos(5, 2) = "2500":      xCampos(5, 3) = "C":    xCampos(5, 4) = "N"
    xCampos(6, 0) = "Importe":       xCampos(6, 1) = "imptot":      xCampos(6, 2) = "800":      xCampos(6, 3) = "N":    xCampos(6, 4) = "N"
    xCampos(7, 0) = "Saldo":         xCampos(7, 1) = "impsal":      xCampos(7, 2) = "800":      xCampos(7, 3) = "N":    xCampos(7, 4) = "N"
    '--obtenemos la consulta
    nSQL = fGenerarConsulta(True, -1, nSQLIdCompra)
    
    If fSeleccionVarios = True Then
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Compras"
    Else
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Compras", "doc", "doc", CualquierParte
    End If
    
    Agregando = True
    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    If fSeleccionVarios = True Then xRs.MoveFirst
    Do While Not xRs.EOF
        With Fg1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = xRs.Fields("id") & ""
            .TextMatrix(.Rows - 1, 2) = xRs.Fields("numreg") & ""
            .TextMatrix(.Rows - 1, 3) = xRs.Fields("abrev") & ""
            .TextMatrix(.Rows - 1, 4) = xRs.Fields("simbolo") & ""
            .TextMatrix(.Rows - 1, 5) = xRs.Fields("doc") & ""
            .TextMatrix(.Rows - 1, 6) = xRs.Fields("fchdoc") & ""
            .TextMatrix(.Rows - 1, 7) = xRs.Fields("nombre") & ""
            .TextMatrix(.Rows - 1, 8) = NulosN(xRs.Fields("imptot"))
            .TextMatrix(.Rows - 1, 9) = NulosN(xRs.Fields("impsal"))
            
            .TextMatrix(.Rows - 1, 12) = NulosN(xRs.Fields("idcue"))
            '---
        End With
        If fSeleccionVarios = False Then Exit Do
        xRs.MoveNext
    Loop
Salir:
    Agregando = False
    Set xRs = Nothing
    '----
        
    Exit Sub
error:
    Agregando = False
    Set xRs = Nothing
    SHOW_ERROR
End Sub

Private Sub pRegistroDel()
    If Fg1.Row <= 0 Then
        MsgBox "La fila no se puede eliminar" + vbCr + "Seleccione una correcta", vbExclamation
        Exit Sub
    End If
    If MsgBox("Seguro desea Eliminar el registro seleccionado", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    '--ELIMINAR EL REGISTRO
    Fg1.RemoveItem (Fg1.Row)
    If Fg1.Rows > 1 Then Fg1.Row = 1
    fTotalizarDatos
End Sub

Private Function fTotalizarDatos() As Boolean
    '--
    Dim sACuenta As Double
    sACuenta = GRID_SUMAR_COL(Fg1, 10)
        
    If IsNumeric(Trim(txt_total(0).Text)) = False Then
        MsgBox "El total entregado no es numérico", vbExclamation, xTitulo
        txt_total(0).Text = "":     txt_total(1).Text = "":     txt_total(2).Text = "":     txt(1).Text = ""
        Exit Function
    End If
    
    txt_total(1).Text = Format(sACuenta, FORMAT_MONTO)
    txt_total(2).Text = Format(NulosN(txt_total(0).Text) - sACuenta, FORMAT_MONTO)
    txt(1).Text = txt_total(2).Text
    If sACuenta > NulosN(txt_total(0).Text) Then
        MsgBox "El Total declarado supera al Total entregado" + vbCr + _
               "Elimine algún registro de compra o modifique el importe a pagar" + vbCr + _
               "Importe a Devolver: " + txt(1).Text, vbExclamation, xTitulo
                
        'txt_total(1).Text = "":     txt_total(2).Text = "":     txt(1).Text = ""
        Cmd(2).SetFocus
'        Exit Function
    End If
            
    fTotalizarDatos = True

End Function


Private Function fGenerarConsulta(fAddRegistro As Boolean, Optional mIdCompra As Integer = -1, Optional nSQLNotIn As String = "") As String
    '--mIdCompra <>-1 CUANDO SE CREA EL REGISTRO DE COMPRA
    Dim nSQL As String
    Dim nnSQLIdCompra As String
    If mIdCompra <> -1 Then nnSQLIdCompra = " AND com_compras.ID = " + CStr(mIdCompra) + " "
    
    If fAddRegistro = True Then '--NUEVO
        nSQL = "SELECT com_compras.id, Format([con_diario]![idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',Format([mae_libros].[codsun],'00')) & Trim([con_diario]![numasi]) AS numreg, mae_documento.abrev, mae_moneda.simbolo, com_compras!numser & ' ' & com_compras!numdoc AS doc, format(com_compras.fchdoc,'dd/mm/yy') as fchdoc, mae_prov.nombre, com_compras.imptot, com_compras.impsal, con_diario.idcue " _
            + vbCr + " FROM (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (con_planctas RIGHT JOIN (com_compras LEFT JOIN con_diario ON com_compras.id = con_diario.idmov) ON con_planctas.id = con_diario.idcue) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id " _
            + vbCr + " WHERE con_diario.idlib=1 AND con_planctas.cuenta Like '42%' and com_compras.impsal <> 0 " + nnSQLIdCompra + nSQLNotIn _
            + vbCr + " ORDER BY mae_prov.nombre, com_compras.fchdoc;"

    Else '--CONSULTA O MODIFICAR
        nSQL = " SELECT con_devolucionesdet.id, com_compras.id AS idcom, Format([con_diario]![idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',Format([mae_libros].[codsun],'00')) & Trim([con_diario]![numasi]) AS numreg, mae_documento.abrev, mae_moneda.simbolo, com_compras!numser & ' ' & com_compras!numdoc AS doc, format(com_compras.fchdoc,'dd/mm/yy') as fchdoc, mae_prov.nombre, com_compras.imptot, con_devolucionesdet.saldo, con_devolucionesdet.acuenta, con_devolucionesdet.nuevosaldo, con_diario.idcue " _
            + vbCr + " FROM mae_libros RIGHT JOIN (con_planctas RIGHT JOIN (((mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) INNER JOIN con_devolucionesdet ON com_compras.id = con_devolucionesdet.idcom) LEFT JOIN con_diario ON com_compras.id = con_diario.idmov) ON con_planctas.id = con_diario.idcue) ON mae_libros.id = com_compras.idlib " _
            + vbCr + " WHERE con_diario.idlib=1 AND con_planctas.cuenta Like '42%' and con_devolucionesdet.ID = " + CStr(RstFrm.Fields("dev_id")) + "  " _
            + vbCr + " ORDER BY mae_prov.nombre, com_compras.fchdoc;"

        
    End If
    
    fGenerarConsulta = nSQL
End Function



Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--AGREGAR COMPRAS YA REGISTRADAS
            pRegistroAdd False
        Case 1 '--SELECCCIONAR
            pRegistroAdd True
        Case 2 '--ELIMINAR REGISTROS AGREGADOS
            pRegistroDel
        Case 3 '--NUEVA COMPRA
            pRegistroNew
            
    End Select
End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index <> 2 Then Exit Sub
    If Button = 2 Then
        If QueHace <> 3 Then
            PopupMenu menu2
        End If
    End If
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

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    If Fg1.Col <> 10 Then
        Fg1.Editable = flexEDNone
    Else
        Fg1.Editable = flexEDKbdMouse
    End If
End Sub
Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case Col
        Case 10
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then 'F3 = Agregar Item
        Cmd_Click 1
    End If
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then
        Cmd_Click 2  'F4 = Eliminar Item
    End If
End Sub
Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row = 0 Then Exit Sub
    Select Case Col
        Case 10
            If IsNumeric(Fg1.TextMatrix(Row, Col)) = False Then
                MsgBox "El valor ingresado no es numérico", vbCritical, xTitulo
                Fg1.TextMatrix(Row, 10) = "":       Fg1.TextMatrix(Row, 11) = ""
            Else
                If NulosN(Fg1.TextMatrix(Row, 10)) > NulosN(Fg1.TextMatrix(Row, 9)) Then
                    MsgBox "El valor Ingresado supera al saldo anterior", vbExclamation, xTitulo
                    Fg1.TextMatrix(Row, 10) = "":        Fg1.TextMatrix(Row, 11) = ""
                    Exit Sub
                End If
                If fTotalizarDatos() = False Then
                    Fg1.TextMatrix(Row, 10) = "":       Fg1.TextMatrix(Row, 11) = ""
                    Exit Sub
                End If
                Fg1.TextMatrix(Row, 11) = NulosN(Fg1.TextMatrix(Row, 9)) - NulosN(Fg1.TextMatrix(Row, 10))
            End If
    End Select
    Exit Sub
error:
    SHOW_ERROR
End Sub


Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
            PopupMenu Menu1
        End If
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    Dim Rpta As Integer

    SeEjecuto = False
    pCargarGrid
    SeEjecuto = True
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado ninguna cuenta por rendir, ¿Desea agergar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            Nuevo
        End If
    End If
End Sub

Private Sub pCargarGrid()

    Dim nSQL  As String
    
    LblPeriodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
    LblPeriodo(1).Caption = LblPeriodo(0).Caption

    'nSQL = "SELECT con_devoluciones.id AS dev_id, con_devoluciones.fchemi AS dev_emi, con_devoluciones.[imp] AS dev_imp, con_devoluciones.obs AS dev_obs," _
        + vbCr + " IIf((SELECT Sum(con_devolucionesdet.acuenta) FROM con_devolucionesdet GROUP BY con_devolucionesdet.id HAVING (((con_devolucionesdet.id)=con_devoluciones.id)) ) Is Null,0,(SELECT Sum(con_devolucionesdet.acuenta) FROM con_devolucionesdet GROUP BY con_devolucionesdet.id HAVING (((con_devolucionesdet.id)= con_devoluciones.id)) )) AS declarado, " _
        + vbCr + " (con_ctasrendir.imp-(declarado+con_devoluciones.[imp])) AS xrendir, " _
        + vbCr + " con_ctasrendir.id, con_ctasrendir.fchemi, con_ctasrendir.fchpag, con_ctasrendir.fchren, con_ctasrendir.numdoc, IIf(con_ctasrendir.tipmov=1,'Caja','Banco') AS tipmov, mae_moneda.descripcion as moneda , mae_moneda.simbolo, IIf(con_ctasrendir.tipmov=1,(SELECT destino.descripcion FROM con_destino AS destino WHERE (((destino.id)=con_ctasrendir.idori)) ),(SELECT [mae_bancos].[descripcion] & '  Cta. N°: ' & [con_bancocuenta].[numcue] AS origen FROM mae_bancos INNER JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban WHERE (((con_bancocuenta.id)=con_ctasrendir.idori)) )) AS origen, con_destino.descripcion AS destino, mae_doccajaban.descripcion AS tipdocnom, mae_doccajaban.abrev, " _
        + vbCr + " IIf(con_ctasrendir.tipper=1,'Persona','Proveedor') AS tipper, IIf(con_ctasrendir.tipper=1,(SELECT [pla_empleados].[nom] & ' ' & [pla_empleados].[ape] AS nombre FROM pla_empleados INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp WHERE (((con_emptes.id)=con_ctasrendir.idper))),(SELECT mae_prov.nombre FROM mae_prov WHERE (((mae_prov.id)=con_ctasrendir.idper)) )) AS benef, con_ctasrendir.[imp], con_destino.idcuen,con_ctasrendir.idmon, " _
        + vbCr + " IIf([con_devoluciones].[numreg] Is Null Or [con_devoluciones].[numreg]='','',Format([con_devoluciones].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([con_devoluciones].[numreg],3)) AS dev_numreg, " _
        + vbCr + " IIf([con_ctasrendir].[numreg] Is Null Or [con_ctasrendir].[numreg]='','',Format([con_ctasrendir].[idmes],'00') & IIf([mae_libros1].[codsun] Is Null Or [mae_libros1].[codsun]='','FF',[mae_libros1].[codsun]) & Mid([con_ctasrendir].[numreg],3)) AS numreg " _
        + vbCr + " FROM ((mae_moneda RIGHT JOIN (mae_doccajaban RIGHT JOIN ((con_destino RIGHT JOIN con_ctasrendir ON con_destino.id = con_ctasrendir.iddes) LEFT JOIN con_devoluciones ON con_ctasrendir.id = con_devoluciones.idren) ON mae_doccajaban.id = con_ctasrendir.tipdoc) ON mae_moneda.id = con_ctasrendir.idmon) LEFT JOIN mae_libros ON con_devoluciones.idlib = mae_libros.id) LEFT JOIN mae_libros AS mae_libros1 ON con_ctasrendir.idlib = mae_libros1.id " _
        + vbCr + " WHERE con_devoluciones.ano= " + AnoTra + " AND con_devoluciones.idmes=" & xMes & "" _
        + vbCr + " ORDER BY con_devoluciones.fchemi;"
    
    nSQL = "SELECT con_devoluciones.id AS dev_id, con_devoluciones.fchemi AS dev_emi, con_devoluciones.[imp] AS dev_imp, con_devoluciones.obs AS dev_obs, " _
        & " IIf((SELECT Sum(con_devolucionesdet.acuenta) FROM con_devolucionesdet GROUP BY con_devolucionesdet.id HAVING (((con_devolucionesdet.id)=con_devoluciones.id))) Is Null," _
        & " 0,(SELECT Sum(con_devolucionesdet.acuenta) FROM con_devolucionesdet GROUP BY con_devolucionesdet.id HAVING (((con_devolucionesdet.id)= con_devoluciones.id)) )) AS declarado, " _
        & " (con_ctasrendir.imp-(declarado+con_devoluciones.[imp])) AS xrendir, con_ctasrendir.id, con_ctasrendir.fchemi, con_ctasrendir.fchpag, con_ctasrendir.fchren, con_ctasrendir.numdoc, " _
        & " IIf(con_ctasrendir.tipmov=1,'Caja','Banco') AS tipmov, mae_moneda.descripcion AS moneda, mae_moneda.simbolo, IIf(con_ctasrendir.tipmov=1, " _
        & " (SELECT destino.descripcion FROM con_destino AS destino WHERE (((destino.id)=con_ctasrendir.idori)) ),(SELECT [mae_bancos].[descripcion] & '  Cta. N°: ' & [con_bancocuenta].[numcue] AS origen " _
        & " FROM mae_bancos INNER JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban WHERE (((con_bancocuenta.id)=con_ctasrendir.idori)) )) AS origen, " _
        & " con_destino.descripcion AS destino, mae_doccajaban.descripcion AS tipdocnom, mae_doccajaban.abrev, IIf(con_ctasrendir.tipper=1,'Persona','Proveedor') AS tipper, " _
        & " IIf([con_ctasrendir].[tipper]=1,(SELECT [pla_empleados].[apepat] & ' '&[pla_empleados].[apemat]&', ' & [pla_empleados].[nom] AS nombre FROM pla_empleados " _
        & " INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp WHERE (((con_emptes.id)=con_ctasrendir.idper))),(SELECT mae_prov.nombre FROM mae_prov " _
        & " WHERE (((mae_prov.id)=con_ctasrendir.idper)) )) AS benef, con_ctasrendir.[imp], con_destino.idcuen, con_ctasrendir.idmon, IIf([con_devoluciones].[numreg] Is Null " _
        & " Or [con_devoluciones].[numreg]='','',Format([con_devoluciones].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) " _
        & " & Mid([con_devoluciones].[numreg],3)) AS dev_numreg, IIf([con_ctasrendir].[numreg] Is Null Or [con_ctasrendir].[numreg]='','',Format([con_ctasrendir].[idmes],'00') " _
        & " & IIf([mae_libros1].[codsun] Is Null Or [mae_libros1].[codsun]='','FF',[mae_libros1].[codsun]) & Mid([con_ctasrendir].[numreg],3)) AS numreg " _
        & " FROM mae_doccajaban RIGHT JOIN (mae_moneda RIGHT JOIN (((con_destino RIGHT JOIN con_ctasrendir ON con_destino.id = con_ctasrendir.iddes) LEFT JOIN mae_libros AS mae_libros1 " _
        & " ON con_ctasrendir.idlib = mae_libros1.id) LEFT JOIN (con_devoluciones LEFT JOIN mae_libros ON con_devoluciones.idlib = mae_libros.id) " _
        & " ON con_ctasrendir.id = con_devoluciones.idren) ON mae_moneda.id = con_ctasrendir.idmon) ON mae_doccajaban.id = con_ctasrendir.tipdoc " _
        & " Where (((con_devoluciones.ano) = 2009) And ((con_devoluciones.idmes) = 10)) ORDER BY con_devoluciones.fchemi"
    
    '--CARGANDO_DATOS
    TabOne1.CurrTab = 0
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg3.DataSource = RstFrm
    Me.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    SeEjecuto = False
    Agregando = False
    QueHace = 3
    
    Dg3.Columns("dev_emi").NumberFormat = FORMAT_DATE
    Dg3.Columns("fchemi").NumberFormat = FORMAT_DATE
    Dg3.Columns("fchpag").NumberFormat = FORMAT_DATE
    Dg3.Columns("fchren").NumberFormat = FORMAT_DATE
    
    Dg3.Columns("dev_imp").NumberFormat = FORMAT_MONTO
    Dg3.Columns("Declarado").NumberFormat = FORMAT_MONTO
    Dg3.Columns("xrendir").NumberFormat = FORMAT_MONTO
    Dg3.Columns("imp").NumberFormat = FORMAT_MONTO
    
    Dg3.BatchUpdates = False
    
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    '--
    Habilitar_Obj False
    Dg3.Splits(0).SizeMode = dbgExact
    Dg3.Splits(0).Size = 5000
    Dg3.Splits(1).Size = 12000
    '----
    Fg1.Tag = Fg1.FormatString
    
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
    pRegistroAdd False
End Sub

Private Sub menu1_3_Click()
    pRegistroDel
End Sub

Private Sub menu1_4_Click()
    pRegistroAdd True
End Sub

Private Sub Menu2_1_Click()
    pRegistroDel
End Sub

Private Sub Menu2_2_Click()
    Dim mRow As Long
    If Fg1.Rows <= 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar todos los registros", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    Agregando = True
    Do While Fg1.Rows > 1
        Fg1.Row = 1
        pRegistroDel
    Loop
    Agregando = False
End Sub

Private Sub opt_operacion_Click(Index As Integer)
        txt_cb(1).Text = ""
        lblCtaDevolucion.Caption = ""
End Sub

Private Sub opt_operacion_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_cb(1).SetFocus
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then MuestraSegundoTab
    End If
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
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
    If Button.Index = 9 Then
        If RstFrm.State = 0 Then Exit Sub
        RstFrm.Filter = ""
    End If
    If Button.Index = 10 Then Buscar
    If Button.Index = 11 Then CambiarMes
    
    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Sub Eliminar()
    
    If MsgBox("¿Esta seguro de eliminar el registro?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
        
        Dim RstTmp As New ADODB.Recordset
        Dim xCod As Long
        
        On Error GoTo error
        xCon.BeginTrans
        
        xCod = NulosN(RstFrm("dev_id"))
        RST_Busq RstTmp, "SELECT con_devolucionesdet.acuenta, con_devolucionesdet.idcom FROM con_devolucionesdet WHERE (((con_devolucionesdet.id)= " & xCod & " ));", xCon
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            xCon.Execute "UPDATE com_compras SET impsal = impsal + " & NulosN(RstTmp.Fields("acuenta")) & " WHERE id = " & RstTmp.Fields("idcom") & ";"
            RstTmp.MoveNext
        Loop
        Set RstTmp = Nothing
        
        '--REINICIAR EL SALDO DE LA CUENTA POR RENDIR
        xCon.Execute "UPDATE con_ctasrendir INNER JOIN con_devoluciones ON con_ctasrendir.id = con_devoluciones.idren " _
                    + vbCr + " SET con_ctasrendir.saldo = [con_ctasrendir].[imp] " _
                    + vbCr + " WHERE (((con_devoluciones.id)= " & xCod & "  ));"
        '--
        xCon.Execute "DELETE * FROM con_devolucionesdet WHERE id = " & xCod & ";"
        xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & xMes & " and idlib = 39 AND idmov = " & xCod & " ;"
        
        xCon.Execute "DELETe * FROM con_devoluciones WHERE id = " & xCod & ""
        
        xCon.CommitTrans
        MsgBox "El registro fue eliminado con éxito", vbInformation, xTitulo
        TabOne1.CurrTab = 0
        RstFrm.Requery
        If RstFrm.RecordCount = 0 Then
            If MsgBox("No hay registrado ninguna devolución, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
                Nuevo
            End If
        Else
            RstFrm.MoveFirst
        End If
    End If
    Exit Sub
error:
    xCon.RollbackTrans
    SHOW_ERROR Me.Name, "Eliminar"
End Sub

Sub Cancelar()
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    ActivaTool
    Habilitar_Obj False
    Label1.Caption = "Detalle de Devolución de Cuentas por Rendir"
    Fg1.SelectionMode = flexSelectionByRow
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
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
    TabOne1.CurrTab = 1
    TabOne2.CurrTab = 0
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Habilitar_Obj True
    MuestraSegundoTab
   
    Label1.Caption = "Modificando Cuentas por Rendir"
    
    If NulosN(txt(1).Text) <> 0 Then pHabilitarBotonInfo True
    Fg1.SelectionMode = flexSelectionFree
    
    TxtFecha(0).SetFocus
End Sub

Sub MuestraSegundoTab()
    On Error GoTo error
    With RstFrm
        Blanquea
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Then Exit Sub
        
        txt(0).Text = .Fields("dev_id") & "" '--CODIGO
        TxtFecha(0).Valor = .Fields("dev_emi")  '--FECHA DE EMISION
        txtfecha_Validate 0, False
        txt(1).Text = Format(NulosN(.Fields("dev_imp")), FORMAT_MONTO)
        txt(2).Text = .Fields("dev_obs") & ""
        '---
        pPonerDatosRendirCta RstFrm
        
        MuestraDetalle
        
    End With
    
    Exit Sub
error:
    
    SHOW_ERROR
End Sub

Private Sub MuestraDetalle()
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim xCol, xFil As Integer
    Dim xSQL As String
    Dim xFch As Date
    Dim xFila  As Integer
    On Error GoTo error
    xSQL = fGenerarConsulta(False)
    
    RST_Busq xRs, xSQL, xCon
    If xRs.RecordCount <> 0 Then
        Agregando = True
        With Fg1
            .Rows = 1
            xRs.MoveFirst
            For A = 1 To xRs.RecordCount
                xFila = .Rows
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NulosC(xRs.Fields("idcom"))
                .TextMatrix(.Rows - 1, 2) = NulosC(xRs.Fields("numreg"))
                .TextMatrix(.Rows - 1, 3) = NulosC(xRs.Fields("abrev"))
                .TextMatrix(.Rows - 1, 4) = NulosC(xRs.Fields("simbolo"))
                .TextMatrix(.Rows - 1, 5) = NulosC(xRs.Fields("doc"))
                .TextMatrix(.Rows - 1, 6) = NulosC(xRs.Fields("fchdoc"))
                .TextMatrix(.Rows - 1, 7) = NulosC(xRs.Fields("nombre"))
                .TextMatrix(.Rows - 1, 8) = NulosN(xRs.Fields("imptot"))
                .TextMatrix(.Rows - 1, 9) = NulosN(xRs.Fields("saldo"))
                .TextMatrix(.Rows - 1, 10) = NulosN(xRs.Fields("acuenta"))
                .TextMatrix(.Rows - 1, 11) = NulosN(xRs.Fields("nuevosaldo"))
                .TextMatrix(.Rows - 1, 12) = NulosN(xRs.Fields("idcue"))
                
                xRs.MoveNext
                If xRs.EOF = True Then Exit For
            Next A
        End With
        
    End If
     
    Set xRs = Nothing
    Agregando = False
    Exit Sub
error:
    Set xRs = Nothing
    Agregando = False
    SHOW_ERROR
End Sub


Private Sub Habilitar_Obj(band As Boolean)
    habilitar_Locked TxtFecha, Not band
    habilitar_Locked txt, Not band
    txt_rendir.Locked = Not band
    txt_rendir.Enabled = band
    habilitar_Locked txt_cb, band
    habilitar Cmd, band
    
    txt_cb(2).Locked = Not band
    cb(2).Enabled = band
    
End Sub

Private Sub Blanquea()
    LimpiaText TxtFecha
    LimpiaText txt
    LimpiaText lbl_dato, True
    LimpiaText txt_total, True
    LimpiaText txt_cb
    LblTipoCambio.Caption = ""
    txt_cb(2).Text = ""
    
    LblIdMoneda.Caption = ""
    lblCtaRendir.Caption = ""
    lblCtaDevolucion.Caption = ""
    
    LimpiarGrid Fg1, True, 1
    OCULTAR_COL Fg1, 1, 1
    OCULTAR_COL Fg1, 12, 12
    Fg1.ColFormat(6) = FORMAT_DATE
    Fg1.ColFormat(8) = FORMAT_MONTO
    Fg1.ColFormat(9) = FORMAT_MONTO
    Fg1.ColFormat(10) = FORMAT_MONTO
    Fg1.ColFormat(11) = FORMAT_MONTO
    
    txt_rendir.Text = ""
End Sub

Private Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Blanquea
    Habilitar_Obj True
    Label1.Caption = "Programando Devolución de Cuentas por Rendir"
    
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    
    TabOne2.CurrTab = 0
    txt_rendir.SetFocus
End Sub

Function Grabar() As Boolean
   If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " la Rendición de Cuenta", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo Salir
    
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim xNumAsiento As String
    Dim xFil As Long
    Dim xCod As Integer
    
    
    On Error GoTo LaCague

    xCon.BeginTrans
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM con_devoluciones ", xCon
        
        xNumAsiento = NuevoNumAsiento(39, xMes, xCon)
        
        xCod = HallaCodigoTabla("con_devoluciones", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xCod
        txt(0).Text = xCod
    Else
        xCod = RstFrm("dev_id")
        RST_Busq RstCab, "SELECT * FROM con_devoluciones WHERE id =" & xCod & "", xCon
        
        xNumAsiento = DevuelveNumAsiento(39, NulosN(RstFrm("dev_id")), xMes, xCon)
        If xNumAsiento = "" Then xNumAsiento = NuevoNumAsiento(39, xMes, xCon)
        
        '--AGREGANDO EL SALDO AL DOCUMENTO DEL PROEVEEDOR
        Dim RstTmp As New ADODB.Recordset
        RST_Busq RstTmp, "SELECT con_devolucionesdet.acuenta, con_devolucionesdet.idcom FROM con_devolucionesdet WHERE (((con_devolucionesdet.id)= " & xCod & " )) ;", xCon
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            xCon.Execute "UPDATE com_compras SET impsal = impsal + " & NulosN(RstTmp.Fields("acuenta")) & " WHERE id = " & RstTmp.Fields("idcom") & " ;"
            RstTmp.MoveNext
        Loop
        Set RstTmp = Nothing
        '--
        xCon.Execute "DELETE * FROM con_devolucionesdet WHERE id = " & xCod & ";"
        xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & xMes & " and idlib = 39 AND idmov = " & xCod & " ;"
        
    End If
    '*******************************
    RST_Busq RstDet, "SELECT top 1 * FROM con_devolucionesdet ", xCon
    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    
    '*******************************
    RstCab("ano") = AnoTra
    RstCab("idlib") = 39 '--rendicion de cuentas
    RstCab("idmes") = xMes
    RstCab("numreg") = Format(xMes, "00") + xNumAsiento
    If xMes <> 0 And xMes <> 13 Then
        RstCab("fchreg") = CDate("01/" + Format(xMes, "00") + "/" + AnoTra)
    End If
    RstCab("fchemi") = CDate(TxtFecha(0).Valor)
    RstCab("idren") = NulosN(lbl_rendir.Caption)
    RstCab("idmon") = NulosN(LblIdMoneda.Caption)
    RstCab("imp") = NulosN(Trim(txt(1).Text)) '--IMPORTE
    RstCab("obs") = Trim(txt(2).Text)
    '*******************************************************
    If NulosN(txt(1).Text) <> 0 Then 'SOLO CUANDO EL IMPORTE A DIF A CERO
        If NulosN(txt(1).Text) > 0 Then RstCab("tipmov") = "1"
        If NulosN(txt(1).Text) < 0 Then RstCab("tipmov") = "2"
        If opt_operacion(0).Value = True Then RstCab("tipope") = "1"
        If opt_operacion(1).Value = True Then RstCab("tipope") = "2"
        
        RstCab("idope") = NulosN(txt_cb(1).Text)
    End If
    RstCab("iddoc") = NulosN(txt_cb(2).Text)
    RstCab("numdoc") = Trim(txt(3).Text)
    '*******************************************************
    
    RstCab.Update
    
    For xFil = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("id") = xCod
        RstDet("idcom") = NulosN(Fg1.TextMatrix(xFil, 1))
        RstDet("saldo") = NulosN(Fg1.TextMatrix(xFil, 9))
        RstDet("acuenta") = NulosN(Fg1.TextMatrix(xFil, 10))
        RstDet("nuevosaldo") = NulosN(Fg1.TextMatrix(xFil, 11))
        RstDet.Update
    Next xFil
    
    '************************************************************* ESCRIBIMOS EN EL DIARIO
    '-haber x el total de importe a rendir
    If NulosN(txt(1).Text) >= 0 Then
        pGenerarAsiento RstDia, AnoTra, xMes, 39, xCod, 0, 0, xNumAsiento, NulosN(LblTipoCambio.Caption), CDate(TxtFecha(0).Valor), NulosN(lblCtaRendir.Caption), NulosN(LblIdMoneda.Caption), NulosN(txt_total(1).Text), False
    Else
        pGenerarAsiento RstDia, AnoTra, xMes, 39, xCod, 0, 0, xNumAsiento, NulosN(LblTipoCambio.Caption), CDate(TxtFecha(0).Valor), NulosN(lblCtaRendir.Caption), NulosN(LblIdMoneda.Caption), NulosN(txt_total(0).Text), False
    End If
    '--debe x c/u de los valores acuenta del proveedor
    For xFil = 1 To Fg1.Rows - 1
        '--ACTUALIZAMOS EL SALDO DEL DOCUMENTO DE COMPRA
        xCon.Execute "UPDATE com_compras SET com_compras.impsal = [com_compras]![impsal] - " & NulosN(Fg1.TextMatrix(xFil, 10)) & " WHERE (com_compras.id = " & NulosN(Fg1.TextMatrix(xFil, 1)) & ") ;"
        '--
        pGenerarAsiento RstDia, AnoTra, xMes, 39, xCod, NulosN(Fg1.TextMatrix(xFil, 1)), 0, xNumAsiento, NulosN(LblTipoCambio.Caption), CDate(TxtFecha(0).Valor), NulosN(Fg1.TextMatrix(xFil, 12)), NulosN(LblIdMoneda.Caption), NulosN(Fg1.TextMatrix(xFil, 10)), True
    Next xFil
    
    '************************************************************* SI EL IMPORTE A DEVOLVER ES DIF CERO
    If NulosN(txt(1).Text) <> 0 Then
        If NulosN(txt(1).Text) > 0 Then
            '"pGenerarAsiento RstDia, " & AnoTra & " ," & xMes & ", 39, " & xCod & " , 0, 0, " & xNumAsiento & " , " & NulosN(LblTipoCambio.Caption) & " , " & CDate(txtfecha(0).Valor) & " , " & NulosN(lblCtaDevolucion.Caption) & " , " & NulosN(lblIdMoneda.Caption) & ", " & NulosN(txt(1).Text) & " , True "
            '--debe = CTA QUE SELECCIONA
            pGenerarAsiento RstDia, AnoTra, xMes, 39, xCod, 0, 0, xNumAsiento, NulosN(LblTipoCambio.Caption), CDate(TxtFecha(0).Valor), NulosN(lblCtaDevolucion.Caption), NulosN(LblIdMoneda.Caption), NulosN(txt(1).Text), True
            '-haber x el total de importe a rendir
            pGenerarAsiento RstDia, AnoTra, xMes, 39, xCod, 0, 0, xNumAsiento, NulosN(LblTipoCambio.Caption), CDate(TxtFecha(0).Valor), NulosN(lblCtaRendir.Caption), NulosN(LblIdMoneda.Caption), NulosN(txt(1).Text), False

        Else
            '--haber=CTA QUE SELECCIONA
            pGenerarAsiento RstDia, AnoTra, xMes, 39, xCod, 0, 0, xNumAsiento, NulosN(LblTipoCambio.Caption), CDate(TxtFecha(0).Valor), NulosN(lblCtaDevolucion.Caption), NulosN(LblIdMoneda.Caption), Abs(NulosN(txt(1).Text)), False
        End If
    End If
    '*************************************************************
    '--ACTUALIZANDO EL SALDO EN CUENTAS POR RENDIR
    xCon.Execute "UPDATE con_ctasrendir SET con_ctasrendir.saldo =  [con_ctasrendir].[imp] - " & NulosN(txt_total(1).Text) + NulosN(txt(1).Text) & " WHERE con_ctasrendir.id =" & NulosN(lbl_rendir.Caption) & " ;"
    '*************************************************************
    
    MsgBox "El registro se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito" + vbCr + "Nun.Reg. " + Format(xMes, "00") + xNumAsiento, vbInformation, xTitulo

    xCon.CommitTrans
    Grabar = True
Salir:
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDia = Nothing
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDia = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + vbCr + Trim(Err.Description), vbCritical, xTitulo
End Function

Private Function fValidarDatos() As Boolean
    If IsDate(TxtFecha(0).Valor) = False Then
        MsgBox "No ha especificado la fecha de Devolución", vbInformation, xTitulo
        TxtFecha(0).SetFocus
        Exit Function
    End If

    If lbl_rendir.Caption = "" Then
        MsgBox "Seleccione una cuenta por rendir", vbInformation, xTitulo
        txt_rendir.SetFocus
        Exit Function
    End If
    
    '--VALIDAR EL INGRESO DE LOS IMPORTES A PAGAR
    Dim mRow  As Long
    For mRow = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(mRow, 10)) = 0 Then
            MsgBox "Ingrese un valor a la Compra:" + vbCr + _
            "Proveedor:        " + Fg1.TextMatrix(mRow, 7) & "" + vbCr + _
            "Num.Reg:         " + Fg1.TextMatrix(mRow, 2) & "" + vbCr + _
            "N°.Documento: " + Fg1.TextMatrix(mRow, 5) & "", vbExclamation, xTitulo
            
            Agregando = True:  Fg1.Row = mRow: Fg1.Col = 10: Agregando = False
            Fg1.SetFocus
            Exit Function
        
        ElseIf NulosN(Fg1.TextMatrix(mRow, 12)) = 0 Then
            MsgBox "El documento de Compra no tiene Cta Contable:" + vbCr + _
            "Proveedor:        " + Fg1.TextMatrix(mRow, 7) & "" + vbCr + _
            "Num.Reg:         " + Fg1.TextMatrix(mRow, 2) & "" + vbCr + _
            "N°.Documento: " + Fg1.TextMatrix(mRow, 5) & "", vbExclamation, xTitulo
            
            Agregando = True:  Fg1.Row = mRow: Fg1.Col = 10: Agregando = False
            Fg1.SetFocus
            Exit Function
        End If
    Next mRow
    '-----
    '-----

    If IsNumeric(Trim(txt(1).Text)) = False Then
        MsgBox "El importe a devolver no es correcto", vbExclamation, xTitulo
        txt(1).Text = ""
        txt(1).SetFocus
        Exit Function
    End If
    If IsDate(lbl_dato(1).Caption) = True Then
        If CDate(TxtFecha(0).Valor) < CDate(lbl_dato(1).Caption) Then
        
            MsgBox "La fecha de Devolución es inferior a la fecha de emisión de la Cuenta por Rendir " + vbCr + _
                   "Fecha de Emisión de la Cuenta: " + lbl_dato(1).Caption + vbCr + "Modifique la fecha", vbInformation, xTitulo
            TxtFecha(0).Valor = ""
            TxtFecha(0).SetFocus
            Exit Function
        End If
    End If
    
    If CDbl(txt_total(2).Text) > CDbl(Trim(txt(1).Text)) Then
        If MsgBox("El monto a devolver es inferior al monto por rendir" + vbCr + _
                  "Diferencia: " + Format(CDbl(txt_total(2).Text) - CDbl(Trim(txt(1).Text)), FORMAT_MONTO) + vbCr + _
                  "Desea Continuar", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function
                  
    ElseIf NulosN(txt_total(2).Text) < NulosN(Trim(txt(1).Text)) Then
        MsgBox "El monto a devolver es superior al monto por rendir" + vbCr + _
                  "Diferencia: " + Format(NulosN(txt_total(2).Text) - NulosN(Trim(txt(1).Text)), FORMAT_MONTO) + vbCr + _
                  "Elimine algún registro de compra o modifique el importe a pagar", vbExclamation, xTitulo
                  
        Agregando = True:  Fg1.Row = 1: Fg1.Col = 10: Agregando = False
        
        Exit Function
    End If
    
'    If NulosN(txt_cb(2).Text) = 0 Then
'        MsgBox "Falta especificar el Documento", vbExclamation, xTitulo
'        txt_cb(2).SetFocus
'        Exit Function
'    End If
    If Trim(txt(3).Text) = "" Then
        MsgBox "Falta especificar Número del Documento", vbExclamation, xTitulo
        txt(3).SetFocus
        Exit Function
    End If
    
    
    
    '*************************************************************
    '** SI EL IMPORTE A DEVOLVER ES DIF CERO
    If NulosN(txt(1).Text) <> 0 Then
        If NulosN(txt_cb(1).Text) = 0 Then
            MsgBox "Falta especificar el " + lbl_cb_capt(1).Caption, vbExclamation, xTitulo
            txt_cb(1).SetFocus
            Exit Function
        End If
    End If
    '*************************************************************
    fValidarDatos = True
End Function
 

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then Imprimir True

    If ButtonMenu.Index = 2 Then Imprimir
End Sub




Private Sub txt_rendir_Change()
    If txt_rendir.Text = "" Then Me.lbl_rendir.Caption = ""
End Sub

Private Sub txt_rendir_KeyDown(KeyCode As Integer, Shift As Integer)
    If txt_rendir.Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_rendir_Click
        Exit Sub
    End If

    If txt_rendir.Text = "" Then Exit Sub
    If KeyCode <> 13 Then Exit Sub
    On Error GoTo error
    Dim RST_TMP As New ADODB.Recordset
    Dim nSQL As String

    nSQL = "SELECT con_ctasrendir.id, format(con_ctasrendir.fchemi,'dd/mm/yy') as fchemi, format(con_ctasrendir.fchpag,'dd/mm/yy') as fchpag, format(con_ctasrendir.fchren,'dd/mm/yy') as fchren, con_ctasrendir.numdoc, IIf(con_ctasrendir.tipmov=1,'Caja','Banco') AS tipmov, mae_moneda.descripcion AS moneda, mae_moneda.simbolo, " _
        + vbCr + " IIf(con_ctasrendir.tipmov=1,(SELECT destino.descripcion FROM con_destino AS destino WHERE (((destino.id)=con_ctasrendir.idori)) ),(SELECT [mae_bancos].[descripcion] & '  Cta. N°: ' & [con_bancocuenta].[numcue] AS origen FROM mae_bancos INNER JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban WHERE (((con_bancocuenta.id)=con_ctasrendir.idori)) )) AS origen, " _
        + vbCr + " con_destino.descripcion AS destino, mae_doccajaban.descripcion AS tipdocnom, mae_doccajaban.abrev, " _
        + vbCr + " IIf(con_ctasrendir.tipper=1,'Persona','Proveedor') AS tipper, IIf(con_ctasrendir.tipper=1,(SELECT [pla_empleados].[nom] & ' ' & [pla_empleados].[ape] AS nombre FROM pla_empleados INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp WHERE (((con_emptes.id)=con_ctasrendir.idper))),(SELECT mae_prov.nombre FROM mae_prov WHERE (((mae_prov.id)=con_ctasrendir.idper)) )) AS benef, con_ctasrendir.[imp], " _
        + vbCr + " IIf((SELECT Sum(con_devolucionesdet.acuenta) AS SumaDeacuenta FROM con_devoluciones INNER JOIN con_devolucionesdet ON con_devoluciones.id = con_devolucionesdet.id GROUP BY con_devoluciones.idren HAVING  (((con_devoluciones.idren)=con_ctasrendir.id)) ) Is Null,0,(SELECT Sum(con_devolucionesdet.acuenta) AS SumaDeacuenta FROM con_devoluciones INNER JOIN con_devolucionesdet ON con_devoluciones.id = con_devolucionesdet.id GROUP BY con_devoluciones.idren HAVING (((con_devoluciones.idren)=con_ctasrendir.id)) )) AS declarado, " _
        + vbCr + " (con_ctasrendir.imp-declarado) AS xrendir  , con_destino.idcuen, con_ctasrendir.idmon " _
        + vbCr + " FROM mae_doccajaban RIGHT JOIN (mae_moneda RIGHT JOIN (con_ctasrendir LEFT JOIN con_destino ON con_ctasrendir.iddes = con_destino.id) ON mae_moneda.id = con_ctasrendir.idmon) ON mae_doccajaban.id = con_ctasrendir.tipdoc " _
        + vbCr + " WHERE (((con_ctasrendir.id) Not In (SELECT con_devoluciones.idren FROM con_devoluciones)) AND ((con_ctasrendir.[imp])>IIf((SELECT Sum(con_devolucionesdet.acuenta) AS SumaDeacuenta FROM con_devoluciones INNER JOIN con_devolucionesdet ON con_devoluciones.id = con_devolucionesdet.id GROUP BY con_devoluciones.idren HAVING  ((con_devoluciones.idren)=con_ctasrendir.id)) Is Null,0,(SELECT Sum(con_devolucionesdet.acuenta) AS SumaDeacuenta FROM con_devoluciones INNER JOIN con_devolucionesdet ON con_devoluciones.id = con_devolucionesdet.id GROUP BY   con_devoluciones.idren HAVING (((con_devoluciones.idren)=con_ctasrendir.id))  ))) AND ((con_ctasrendir.idest)=2)) " _
        + vbCr + " AND con_ctasrendir.id =" + CStr(txt_rendir.Text) + " "



    Me.MousePointer = vbHourglass
    If xCon.State = 0 Then GoTo Salir
    RST_Busq RST_TMP, nSQL, xCon
    
    If RST_TMP.State = 0 Then GoTo Salir
    If RST_TMP.RecordCount > 0 Then
        '----------
        LimpiaText TxtFecha
        LimpiaText lbl_dato, True
        LimpiaText txt_total, True
        
        TxtFecha(0).Valor = ""
        txt(1).Text = ""
        txt(2).Text = ""
        '------------
        pPonerDatosRendirCta RST_TMP

    Else
        txt_rendir.Text = ""
    End If
            
Salir:
    Set RST_TMP = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    Set RST_TMP = Nothing
    SHOW_ERROR
End Sub

Private Sub txt_rendir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub




Private Function fVerificarProgAut(BUSCAPROGRAMADOR As Boolean, OBJ_ID As Label, OBJ_NOMBRE As Label) As Boolean
    Dim RST_TMP As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQL_PROG As String
    If BUSCAPROGRAMADOR = True Then
        nSQL_PROG = " AND con_emptes.prog=-1;"
    Else
        nSQL_PROG = " AND con_emptes.aut=-1;"
    End If
    
    nSQL = "SELECT  con_emptes.id, [pla_empleados].[nom] & ' ' & [pla_empleados].[ape] AS nombre " _
    + vbCr + " FROM (pla_empleados INNER JOIN mae_usuarios ON pla_empleados.id = mae_usuarios.id) INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp " _
    + vbCr + " WHERE mae_usuarios.id= " + CStr(xIdUsuario) + nSQL_PROG
    
    
    RST_Busq RST_TMP, nSQL, xCon
    If RST_TMP.State = 0 Then GoTo Salir
    If RST_TMP.EOF = True Or RST_TMP.BOF = True Then
        OBJ_ID.Caption = "0"
        OBJ_NOMBRE.Caption = "NO ES " + IIf(BUSCAPROGRAMADOR = True, "PROGRAMADOR", "AUTORIZADOR")
    Else
        OBJ_ID.Caption = RST_TMP.Fields(0) & ""
        OBJ_NOMBRE.Caption = IIf(BUSCAPROGRAMADOR = True, "PROGRAMADOR", "AUTORIZADOR") + ":  " + RST_TMP.Fields(1) & ""
        fVerificarProgAut = True
    End If
Salir:
    Set RST_TMP = Nothing
End Function



Private Sub pPonerDatosRendirCta(rst As ADODB.Recordset)
'    On Error GoTo ERROR
    With rst
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Then Exit Sub
        txt_rendir.Text = .Fields("id") & "" '--CODIGO
        lbl_rendir.Caption = .Fields("id") & ""  '--CODIGO
        
        lbl_dato(1).Caption = .Fields("fchemi") & "" '--FECHA DE EMISION
        lbl_dato(2).Caption = .Fields("fchpag") & "" '--FECHA DE PAGO
        lbl_dato(10).Caption = .Fields("fchren") & "" '--FECHA DE PAGO
        
        '--DEL TIPO DE DOCUMENTO
        Me.lbl_dato(3).Caption = NulosC(.Fields("tipdocnom")) '--DESCRIPCION DEL TIPO DOC
        '--DE LA MONEDA
        Me.lbl_dato(4).Caption = NulosC(.Fields("moneda"))   '--MONEDA
        LblIdMoneda.Caption = NulosN(.Fields("idmon"))
        
        lbl_dato(5).Caption = NulosN(.Fields("numdoc"))
        lbl_dato(6).Caption = Format(NulosN(.Fields("imp")), FORMAT_MONTO)
        '--DEL ORIGEN
        lbl_dato_x(7).Caption = NulosC(.Fields("tipmov")) '--TIPO DE MOVIENTO
        lbl_dato(7).Caption = NulosC(.Fields("origen"))
        '--DEL DESTINO
        Me.lbl_dato(8).Caption = NulosC(.Fields("destino"))
        
        lblCtaRendir.Caption = NulosN(.Fields("idcuen"))
        '--DEL BENEFICIARIO
        lbl_dato_x(9).Caption = NulosC(.Fields("tipper"))
        lbl_dato(9).Caption = NulosC(.Fields("benef")) '--NOMBRE DE LA PERSONA,NOMBRE DEL PROVEEDOR
        
        '--DE LOS TOTALES
        txt_total(0).Text = Format(NulosN(.Fields("imp")), FORMAT_MONTO)
        txt_total(1).Text = Format(NulosN(.Fields("declarado")), FORMAT_MONTO)
        txt_total(2).Text = Format(NulosN(txt_total(0).Text) - NulosN(txt_total(1).Text), FORMAT_MONTO)
        
            Dim RstDev As New ADODB.Recordset
            RST_Busq RstDev, "select * from con_devoluciones where idren = " & rst.Fields(("id")) & " ;", xCon
            If RstDev.RecordCount <> 0 Then
                If NulosN(RstDev.Fields("tipope")) = 1 Then
                    opt_operacion(0).Value = True
                Else
                    opt_operacion(1).Value = True
                End If
                'moneda
                txt_cb(0).Text = LblIdMoneda.Caption
                lbl_cb(0).Caption = lbl_dato(4).Caption
                lbl_cb_cod(0).Caption = LblIdMoneda.Caption
    
                If NulosN(RstDev("idope")) <> 0 Then
                    txt_cb(1).Text = NulosN(RstDev("idope"))
                    txt_cb_Validate 1, False
                End If
                If NulosN(RstDev("iddoc")) <> 0 Then
                    txt_cb(2).Text = NulosN(RstDev("iddoc"))
                    txt_cb_Validate 2, False
                End If
                txt(3).Text = NulosC(RstDev.Fields("numdoc"))
            End If
            If NulosN(txt(1).Text) <> 0 Then
               pHabilitarBotonInfo False
           End If
        txt_cb(0).Text = LblIdMoneda.Caption
        lbl_cb(0).Caption = lbl_dato(4).Caption
        lbl_cb_cod(0).Caption = LblIdMoneda.Caption
        
        
        Set RstDev = Nothing
        
        End With
    Exit Sub
error:
    Set RstDev = Nothing
    SHOW_ERROR Me.Name, "pPonerDatosRendirCta"

End Sub

'------DEL CAMBIO DE PERIODO

Sub CambiarMes()
    xMes = SeleccionaMes(xCon)
    pCargarGrid
End Sub



Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
   
    Dim xCampos(5, 4) As String
    xCampos(0, 0) = "Num.Reg.":         xCampos(0, 1) = "numreg":   xCampos(0, 2) = "1000":      xCampos(0, 3) = "c"
    xCampos(1, 0) = "Fch.Dev":          xCampos(1, 1) = "fchemi":   xCampos(1, 2) = "1000":      xCampos(1, 3) = "F"
    xCampos(2, 0) = "N°.Documento":     xCampos(2, 1) = "numdoc":   xCampos(2, 2) = "1500":     xCampos(2, 3) = "C"
    xCampos(3, 0) = "M":                xCampos(3, 1) = "simbolo":  xCampos(3, 2) = "550":      xCampos(3, 3) = "C"
    xCampos(4, 0) = "Imp.Devuelto":     xCampos(4, 1) = "imp":      xCampos(4, 2) = "1200":     xCampos(4, 3) = "N"
    
    nSQL = "SELECT con_devoluciones.id, Format([con_devoluciones].[fchemi],'dd/mm/yy') AS fchemi, mae_moneda.descripcion AS moneda, mae_moneda.simbolo, con_devoluciones.numdoc, con_devoluciones.[imp], con_devoluciones.obs, IIf([con_devoluciones].[numreg] Is Null Or [con_devoluciones].[numreg]='','',Format([con_devoluciones].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([con_devoluciones].[numreg],3)) AS numreg " _
        + vbCr + " FROM mae_moneda RIGHT JOIN (con_devoluciones LEFT JOIN mae_libros ON con_devoluciones.idlib = mae_libros.id) ON mae_moneda.id = con_devoluciones.idmon " _
        + vbCr + " WHERE con_devoluciones.ano= " + AnoTra + " AND con_devoluciones.idmes=" & xMes & "" _
        + vbCr + " ORDER BY con_devoluciones.fchemi;"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Devolución de Cuenta", "numdoc", "numdoc", Principio
    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(xRs("id"))
Salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub




Private Sub Filtrar()
    
    Dim xCampos(5, 4) As String
    xCampos(0, 0) = "Num.Reg.":         xCampos(0, 1) = "numreg":   xCampos(0, 2) = "C":      xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Fch.Dev":          xCampos(1, 1) = "fchemi":   xCampos(1, 2) = "F":      xCampos(1, 3) = "1000"
    xCampos(2, 0) = "N°.Documento":     xCampos(2, 1) = "numdoc":   xCampos(2, 2) = "C":     xCampos(2, 3) = "1200"
    xCampos(3, 0) = "M":                xCampos(3, 1) = "simbolo":  xCampos(3, 2) = "C":      xCampos(3, 3) = "550"
    xCampos(4, 0) = "Imp.Devuelto":     xCampos(4, 1) = "imp":      xCampos(4, 2) = "N":     xCampos(4, 3) = "1200"
    
    TabOne1.CurrTab = 0
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg3

End Sub


Private Sub Imprimir(Optional IMP_LISTADO As Boolean = False)

    On Error GoTo error

    Me.MousePointer = vbHourglass
    If IMP_LISTADO = False Then
        If Me.TabOne1.CurrTab = 0 Then
        
        Else
''            MsgBox "Primero muestre el detalle del Registro" + vbCr + _
''                   "Luego inténtelo otra vez", vbExclamation, xTitulo
        End If
    Else
    
        TDB_IMPRIMIR Dg3, "IMPRESIÓN DE DEVOLUCIÓN DE CUENTAS", "LISTADO DE DEVOLUCIÓN DE CUENTAS -  Periodo: " + MonthName(xMes, False)
   
    End If

    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "IMPRIMIR"

End Sub


'*******************
Private Sub pRegistroNew()
''    Dim xForm As New sgi2_compras.Compras
''    Dim xRs As New ADODB.Recordset
''    Dim nSQL As String
''    Dim xIdCompra As Integer
''    On Error GoTo error
''
''    xIdCompra = xForm.RegCompras(xCon, xMes, 1)
''    Set xForm = Nothing
''    If xIdCompra = 0 Then Exit Sub
''
''    '--VALIDAR QUE EL IDCOMPRA NO ESTE EN LA LISTA
''    If VERIFICAR_LISTA(Fg1, 1, CStr(xIdCompra)) = False Then Exit Sub
''    nSQL = fGenerarConsulta(True, xIdCompra)
''    RST_Busq xRs, nSQL, xCon
''
''    If xRs.State = 0 Then GoTo Salir
''    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
''    Agregando = True
''
''    With Fg1
''        .Rows = .Rows + 1
''        .TextMatrix(.Rows - 1, 1) = xRs.Fields("id") & ""
''        .TextMatrix(.Rows - 1, 2) = xRs.Fields("numreg") & ""
''        .TextMatrix(.Rows - 1, 3) = xRs.Fields("abrev") & ""
''        .TextMatrix(.Rows - 1, 4) = xRs.Fields("simbolo") & ""
''        .TextMatrix(.Rows - 1, 5) = xRs.Fields("doc") & ""
''        .TextMatrix(.Rows - 1, 6) = xRs.Fields("fchdoc") & ""
''        .TextMatrix(.Rows - 1, 7) = xRs.Fields("nombre") & ""
''        .TextMatrix(.Rows - 1, 8) = NulosN(xRs.Fields("imptot"))
''        .TextMatrix(.Rows - 1, 9) = NulosN(xRs.Fields("impsal"))
''        '---
''    End With
''Salir:
''    Agregando = False
''    Set xRs = Nothing
''    Exit Sub
''error:
''    Agregando = False
''    Set xRs = Nothing
End Sub
'**************************



Private Sub pGenerarAsiento(RstDiario As ADODB.Recordset, nAnoTrabajo, mMesActivo, IDLibro, IDMov, mIdDocPro, mCorr, nAsiento, mTipoCambio, FchDoc, IDcuenta, IDMoneda, mImporte, Optional EsDEBE As Boolean)
    '--mCorr por le general es igual a 0
    RstDiario.AddNew
    RstDiario("año") = nAnoTrabajo
    RstDiario("idmes") = mMesActivo  'CODIGO DEL MES
    RstDiario("idlib") = IDLibro     'CODIGO DEL LIBRO
    RstDiario("idmov") = IDMov       'CODIGO DEL MOVIMIENTO
    RstDiario("iddocpro") = mIdDocPro
    RstDiario("correlativo") = mCorr
    RstDiario("numasi") = nAsiento
    RstDiario("tc") = mTipoCambio
    If mMesActivo = 0 Then
        RstDiario("fchasi") = CDate("01/01/" + nAnoTrabajo)
    ElseIf mMesActivo = 13 Then
        RstDiario("fchasi") = CDate("31/12/" + nAnoTrabajo)
    Else
        RstDiario("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + nAnoTrabajo)
    End If
    RstDiario("fchdoc") = FchDoc
    RstDiario("idcue") = IDcuenta
    If EsDEBE = False Then
        If IDMoneda = 1 Then '--DE LA MONEDA
            RstDiario("imphabsol") = mImporte
            RstDiario("imphabdol") = 0
        Else
            RstDiario("imphabsol") = mImporte * mTipoCambio
            RstDiario("imphabdol") = mImporte
        End If
    Else
        If IDMoneda = 1 Then '--DE LA MONEDA
            RstDiario("impdebsol") = mImporte
            RstDiario("impdebdol") = 0
        Else
            RstDiario("impdebsol") = mImporte * mTipoCambio
            RstDiario("impdebdol") = mImporte
        End If
    End If

    RstDiario.Update
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    If Index = 1 Then
        '--limpiar
        txt_cb(1).Text = ""
        '--
        If NulosN(txt(1).Text) = 0 Then '--no insertar datos
            txt(1).Text = 0
            fra_datos.Caption = "[ Restringido ]"
            
        ElseIf NulosN(txt(1).Text) > 0 Then
            fra_datos.Caption = "[ Ingreso ]"
            lbl_cb_capt(1).Caption = "Destino del Ingreso"
                
        Else
            fra_datos.Caption = "[ Egreso ]"
            lbl_cb_capt(1).Caption = "Origen del Egreso"
            
        End If
        
        pHabilitarBotonInfo False
        
        If NulosN(txt(1).Text) <> 0 Then
            pHabilitarBotonInfo True
        End If
        
    End If
End Sub

Private Sub txtfecha_Validate(Index As Integer, Cancel As Boolean)
    If Index <> 0 Then Exit Sub
    If IsDate(TxtFecha(0).Valor) = True Then
        LblTipoCambio.Caption = HallaTipoCambio(TxtFecha(0).Valor, 2, Venta, xCon)
    Else
        LblTipoCambio.Caption = ""
    End If
End Sub




'*************************************

'--de los estados
'LblEstado(1).Caption = "2"
'LblEstado(1).Caption = "4"
'

Private Sub cb_Click(Index As Integer)
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nCampoBusca As String
    Dim nSQL As String
    
    If QueHace = 3 Then Exit Sub
    If Index = 0 Then Exit Sub
    On Error GoTo error

    Select Case Index
        Case 0 '--MONEDA
        Case 1 '--
            'INGRESO: DESTINO DEL INGRESO
            'EGRESO: ORIGEN DEL EGRESO
            
            nCampoBusca = "nombre"
            If NulosN(txt(1).Text) > 0 Then
            '--ingreso
                
                ReDim xCampos(4, 3) As String
                xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "3500":   xCampos(0, 3) = "C"
                xCampos(1, 0) = "Cuenta":       xCampos(1, 1) = "cuenta":    xCampos(1, 2) = "2800":   xCampos(1, 3) = "C"
                xCampos(2, 0) = "N°.Cuenta":    xCampos(2, 1) = "numcta":    xCampos(2, 2) = "1300":    xCampos(2, 3) = "C"
                xCampos(3, 0) = "Id":           xCampos(3, 1) = "id":        xCampos(3, 2) = "450":    xCampos(3, 3) = "N"
            
                nSQL = "SELECT con_destino.id, con_destino.descripcion AS nombre, con_destino.id AS cod, con_destino.idcuen as idcta, con_planctas.descripcion AS cuenta, con_planctas.cuenta AS numcta " _
                    + vbCr + " FROM con_planctas RIGHT JOIN con_destino ON con_planctas.id = con_destino.idcuen " _
                    + vbCr + " Where (((con_destino.tipmov) = 1) And ((con_destino.idmon) = " & NulosN(lbl_cb_cod(0).Caption) & ")) " _
                    + vbCr + " ORDER BY con_destino.descripcion;"
                    
                nTitulo = "Buscando Destinos del Egreso"

            Else
            '--egreso
                If opt_operacion(0).Value = True Then '--caja / Origen
                
                    ReDim xCampos(4, 3) As String
                    xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "3500":   xCampos(0, 3) = "C"
                    xCampos(1, 0) = "Cuenta":       xCampos(1, 1) = "cuenta":    xCampos(1, 2) = "1000":   xCampos(1, 3) = "C"
                    xCampos(2, 0) = "N°.Cuenta":    xCampos(2, 1) = "numcta":    xCampos(2, 2) = "800":    xCampos(2, 3) = "C"
                    xCampos(3, 0) = "Id":           xCampos(3, 1) = "id":        xCampos(3, 2) = "450":    xCampos(3, 3) = "N"
                
                    nSQL = "SELECT con_origen.id, con_origen.descripcion AS nombre, con_origen.id AS cod, con_origen.idcue as idcta, con_planctas.descripcion AS cuenta, con_planctas.cuenta AS numcta " _
                        + vbCr + " FROM con_planctas RIGHT JOIN con_origen ON con_planctas.id = con_origen.idcue " _
                        + vbCr + " WHERE (((con_origen.tipmov) = 2) And ((con_origen.idmon) = " & NulosN(lbl_cb_cod(0).Caption) & ")) " _
                        + vbCr + " ORDER BY con_origen.descripcion;"
                        
                    nTitulo = "Buscando Origen del Egreso"
                
                Else '--banco / Origen
                
                    ReDim xCampos(4, 3) As String
                    xCampos(0, 0) = "Banco":            xCampos(0, 1) = "banco":    xCampos(0, 2) = "3500":   xCampos(0, 3) = "C"
                    xCampos(1, 0) = "N° Cuenta":        xCampos(1, 1) = "numcue":   xCampos(1, 2) = "1000":   xCampos(1, 3) = "C"
                    xCampos(2, 0) = "M":                xCampos(2, 1) = "simbolo":  xCampos(2, 2) = "800":    xCampos(2, 3) = "C"
                    xCampos(3, 0) = "N°Cta. Contable":  xCampos(3, 1) = "numcta":   xCampos(3, 2) = "450":    xCampos(3, 3) = "N"
                
                    nSQL = "SELECT con_bancocuenta.id, [mae_bancos].[descripcion] & '  N° Cta. ' & [con_bancocuenta].[numcue] AS nombre, con_bancocuenta.id AS cod, mae_bancos.descripcion AS banco, con_bancocuenta.numcue, mae_moneda.simbolo, con_bancocuenta.idcuen as idcta, con_planctas.cuenta, con_planctas.descripcion " _
                        + vbCr + " FROM con_planctas RIGHT JOIN ((mae_bancos RIGHT JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban) LEFT JOIN mae_moneda ON con_bancocuenta.idmon = mae_moneda.id) ON con_planctas.id = con_bancocuenta.idcuen " _
                        + vbCr + " WHERE (((con_bancocuenta.idmon)=1));"
                        
                    nTitulo = "Buscando Cuentas del Banco"
                    nCampoBusca = "banco"
                End If
                    
            End If
                    
        Case 2 '--DOCUMENTO
            ReDim xCampos(3, 3) As String
            xCampos(0, 0) = "Descripción":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Abrev":         xCampos(1, 1) = "abrev":     xCampos(1, 2) = "500":    xCampos(1, 3) = "C"
            xCampos(2, 0) = "Id":         xCampos(2, 1) = "id":     xCampos(2, 2) = "450":    xCampos(2, 3) = "n"
            nTitulo = "Buscando Documentos"

             nSQL = "SELECT mae_documento.id, mae_documento.descripcion AS nombre, mae_documento.id AS cod, mae_documento.abrev " _
                + vbCr + " FROM mae_documento" _
                + vbCr + " ORDER BY mae_documento.descripcion ASC "
            
            nCampoBusca = "nombre"
                        
    End Select
    Dim xRs As New ADODB.Recordset

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, nCampoBusca, nCampoBusca, Principio
    
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO

    If Index = 1 Then 'obteniendo el id de Cuenta contable
        lblCtaDevolucion.Caption = NulosN(xRs.Fields("idcta"))
    End If
    
    Select Case Index
        Case 1 '--
            txt_cb(2).SetFocus
        Case 2 '--DOCUMENTO
            txt(3).SetFocus
    End Select

Salir:
    Set xRs = Nothing
Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If Trim(txt_cb(Index).Text) = "" Then
        lbl_cb_cod(Index).Caption = ""
        lbl_cb(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
    
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    
    If txt_cb(Index).Text = "" Then Exit Sub
    If Index = 0 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error

    Select Case Index
        Case 0 '--MONEDA
        Case 1 '--
            'INGRESO: DESTINO DEL INGRESO
            'EGRESO: ORIGEN DEL EGRESO
            
            If NulosN(txt(1).Text) > 0 Then
            '--ingreso
                
                nSQL = "SELECT con_destino.id, con_destino.descripcion AS nombre, con_destino.id AS cod, con_destino.idcuen as idcta, con_planctas.descripcion AS cuenta, con_planctas.cuenta AS numcta " _
                    + vbCr + " FROM con_planctas RIGHT JOIN con_destino ON con_planctas.id = con_destino.idcuen " _
                    + vbCr + " WHERE (((con_destino.tipmov) = 1) And ((con_destino.idmon) = " & NulosN(lbl_cb_cod(0).Caption) & "))  and con_destino.id = " & NulosN(txt_cb(Index).Text) & "" _
                    + vbCr + " ORDER BY con_destino.descripcion;"

            Else
            '--egreso
                If opt_operacion(0).Value = True Then '--caja / Origen
                                
                    nSQL = "SELECT con_origen.id, con_origen.descripcion AS nombre, con_origen.id AS cod, con_origen.idcue as idcta, con_planctas.descripcion AS cuenta, con_planctas.cuenta AS numcta " _
                        + vbCr + " FROM con_planctas RIGHT JOIN con_origen ON con_planctas.id = con_origen.idcue " _
                        + vbCr + " WHERE (((con_origen.tipmov) = 2) And ((con_origen.idmon) = " & NulosN(lbl_cb_cod(0).Caption) & "))  and con_origen.id = " & NulosN(txt_cb(Index).Text) & "" _
                        + vbCr + " ORDER BY con_origen.descripcion;"
                                       
                Else '--banco / Origen
                
                    nSQL = "SELECT con_bancocuenta.id, [mae_bancos].[descripcion] & '  N° Cta. ' & [con_bancocuenta].[numcue] AS nombre, con_bancocuenta.id AS cod, mae_bancos.descripcion AS banco, con_bancocuenta.numcue, mae_moneda.simbolo, con_bancocuenta.idcuen as idcta, con_planctas.cuenta, con_planctas.descripcion " _
                        + vbCr + " FROM con_planctas RIGHT JOIN ((mae_bancos RIGHT JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban) LEFT JOIN mae_moneda ON con_bancocuenta.idmon = mae_moneda.id) ON con_planctas.id = con_bancocuenta.idcuen " _
                        + vbCr + " WHERE (((con_bancocuenta.idmon)=1))  and con_bancocuenta.id = " & NulosN(txt_cb(Index).Text) & ";"
                        
                End If
                    
            End If
                    
        Case 2 '--DOCUMENTO

             nSQL = "SELECT mae_documento.id, mae_documento.descripcion AS nombre, mae_documento.id AS cod, mae_documento.abrev " _
                + vbCr + " FROM mae_documento " _
                + vbCr + " WHERE mae_documento.id = " & NulosN(txt_cb(Index).Text) & ";"
                        
    End Select
    If xCon.State = 0 Then Exit Sub
    RST_Busq xRs, nSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount > 0 Then
        txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
        
        If Index = 1 Then 'obteniendo el id de Cuenta contable
            lblCtaDevolucion.Caption = NulosN(xRs.Fields("idcta"))
        End If
        
    Else
        txt_cb(Index).Text = "":
        lblCtaDevolucion.Caption = ""
    End If
    Set xRs = Nothing
       
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR
End Sub


Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    Select Case Index
        Case 1
            If KeyAscii = 45 Then Exit Sub
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
'        Case 2:
'            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub

'**********************
Private Sub pHabilitarBotonInfo(band As Boolean)
    habilitar opt_operacion, band
    
    txt_cb(0).Locked = Not band
    txt_cb(1).Locked = Not band
    
    cb(0).Enabled = band
    cb(1).Enabled = band
    
End Sub
