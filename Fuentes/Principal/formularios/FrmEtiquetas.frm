VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEtiquetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Etiquetas"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7815
      TabIndex        =   36
      Top             =   390
      Width           =   1470
   End
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
            Picture         =   "FrmEtiquetas.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
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
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5640
      Left            =   0
      TabIndex        =   14
      Top             =   360
      Width           =   9285
      _cx             =   16378
      _cy             =   9948
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      FrontTabForeColor=   -2147483630
      Caption         =   "   &Consulta   |    &Detalle    "
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
         Caption         =   "Frame2"
         Height          =   5265
         Left            =   9930
         TabIndex        =   16
         Top             =   330
         Width           =   9195
         Begin VB.TextBox TxtProv 
            Height          =   300
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "TxtProv"
            Top             =   4740
            Width           =   5010
         End
         Begin VB.CommandButton CmdBusTiDocEmp 
            Height          =   240
            Left            =   3840
            Picture         =   "FrmEtiquetas.frx":277E
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   735
            Width           =   240
         End
         Begin VB.TextBox TxtIdFormato 
            Height          =   300
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "TxtIdFormato"
            Top             =   705
            Width           =   930
         End
         Begin VB.TextBox TxtDir 
            Height          =   300
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "TxtDir"
            Top             =   4440
            Width           =   5010
         End
         Begin VB.TextBox TxtIngre 
            Height          =   300
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "TxtIngre"
            Top             =   4140
            Width           =   5010
         End
         Begin VB.TextBox TxtObserva 
            Height          =   300
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "TxtObserva"
            Top             =   3840
            Width           =   5010
         End
         Begin VB.TextBox TxtNumRes 
            Height          =   300
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "TxtNumRes"
            Top             =   3540
            Width           =   5010
         End
         Begin VB.TextBox TxtNumLot 
            Height          =   300
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "TxtNumLot"
            Top             =   2640
            Width           =   2010
         End
         Begin VB.TextBox TxtPesTar 
            Height          =   300
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "TxtPesTar"
            Top             =   2325
            Width           =   2010
         End
         Begin VB.TextBox TxtPesNet 
            Height          =   300
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "TxtPesNet"
            Top             =   2025
            Width           =   2010
         End
         Begin VB.TextBox TxtPesBru 
            Height          =   300
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "TxtPesBru"
            Top             =   1725
            Width           =   2010
         End
         Begin VB.TextBox TxtTitulo2 
            Height          =   300
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   2
            Text            =   "TxtTitulo2"
            Top             =   1425
            Width           =   5010
         End
         Begin VB.TextBox TxtTitulo 
            Height          =   300
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   1
            Text            =   "TxtTitulo"
            Top             =   1140
            Width           =   5010
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchProd 
            Height          =   300
            Left            =   3180
            TabIndex        =   7
            Top             =   2940
            Width           =   1485
            _ExtentX        =   2619
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchVen 
            Height          =   300
            Left            =   3180
            TabIndex        =   8
            Top             =   3240
            Width           =   1485
            _ExtentX        =   2619
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
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
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
            Left            =   1185
            TabIndex        =   37
            Top             =   4785
            Width           =   885
         End
         Begin VB.Label LblFormato 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblFormato"
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
            Left            =   4125
            TabIndex        =   35
            Top             =   705
            Width           =   4065
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Formato"
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
            Left            =   1155
            TabIndex        =   33
            Top             =   750
            Width           =   690
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "Label13"
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
            Height          =   240
            Left            =   105
            TabIndex        =   29
            Top             =   60
            Width           =   9000
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
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
            Left            =   1185
            TabIndex        =   28
            Top             =   4485
            Width           =   825
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Ingredientes"
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
            Left            =   1155
            TabIndex        =   27
            Top             =   4185
            Width           =   1065
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Observación"
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
            Left            =   1155
            TabIndex        =   26
            Top             =   3885
            Width           =   1080
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Nº Resolución"
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
            Left            =   1155
            TabIndex        =   25
            Top             =   3585
            Width           =   1230
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Vencimiento"
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
            Left            =   1155
            TabIndex        =   24
            Top             =   3270
            Width           =   1905
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Producción"
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
            Left            =   1155
            TabIndex        =   23
            Top             =   2970
            Width           =   1410
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nº Lote"
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
            Left            =   1155
            TabIndex        =   22
            Top             =   2685
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Peso Tara"
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
            Left            =   1155
            TabIndex        =   21
            Top             =   2370
            Width           =   885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Peso Neto"
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
            Left            =   1155
            TabIndex        =   20
            Top             =   2070
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Peso Bruto"
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
            Left            =   1155
            TabIndex        =   19
            Top             =   1770
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Título 2"
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
            Left            =   1155
            TabIndex        =   18
            Top             =   1470
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Título 1"
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
            Left            =   1155
            TabIndex        =   17
            Top             =   1185
            Width           =   690
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5265
         Left            =   45
         TabIndex        =   15
         Top             =   330
         Width           =   9195
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   4905
            Left            =   0
            TabIndex        =   31
            Top             =   330
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   8652
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
            Columns(1).Caption=   "Etiqueta"
            Columns(1).DataField=   "titulo1"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Pro."
            Columns(2).DataField=   "fchpro1"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Ven."
            Columns(3).DataField=   "fchven1"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Nº Lote"
            Columns(4).DataField=   "numlot1"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Formato"
            Columns(5).DataField=   "idform1"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=741"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=661"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=6826"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6747"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1799"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1720"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1852"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1773"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2619"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2540"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=1376"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1296"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
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
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "Label13"
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
            Height          =   240
            Left            =   105
            TabIndex        =   30
            Top             =   60
            Width           =   9000
         End
      End
   End
End
Attribute VB_Name = "FrmEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstLista As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim xConEtiqueta As New ADODB.Connection
Dim fOrdenLista As Boolean           ' especfica el orden de la lista de la consulta
Dim mIdRegistro&                     ' identificador del registro

Public xTitulo As String

Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO



Function Grabar() As Boolean
On Error GoTo LaCague
    
    Dim xCampos(14, 5) As String
    
    Dim xId As Double
    Dim xIdNotDebito As Double
    Dim xTipLet, A As Integer
    Dim FchReg As String
    
    xConEtiqueta.BeginTrans
    'ESPECIFICAMOS EL ID DEL MOVIMIENTO
    If QueHace = 1 Then
        xId = HallaCodigoTabla("etiqueta", xConEtiqueta, "id")
    Else
        xId = RstLista("id")
    End If
    
    mIdRegistro = xId
    
    'Columna    | Descripcion
    '------------------------
    '0          | campo
    '1          | Valor
    '2          | requerido
    '3          | tipo
    '4          | msj error
    
    '--------------------------------
    'GRABAMOS LA CABECERA DE LA LETRA
    xCampos(0, 0) = "id":            xCampos(0, 1) = Str(xId):            xCampos(0, 2) = "S":    xCampos(0, 3) = "N":     xCampos(0, 4) = "":  xCampos(0, 5) = "S"
    xCampos(1, 0) = "titulo1":       xCampos(1, 1) = TxtTitulo.Text:      xCampos(1, 2) = "S":    xCampos(1, 3) = "C":     xCampos(1, 4) = "":  xCampos(1, 5) = "N"
    xCampos(2, 0) = "titulo2":       xCampos(2, 1) = TxtTitulo2.Text:     xCampos(2, 2) = "N":    xCampos(2, 3) = "C":     xCampos(2, 4) = "":  xCampos(2, 5) = "N"
    xCampos(3, 0) = "pesbru":        xCampos(3, 1) = TxtPesBru.Text:      xCampos(3, 2) = "N":    xCampos(3, 3) = "C":     xCampos(3, 4) = "":  xCampos(3, 5) = "N"
    xCampos(4, 0) = "pesnet":        xCampos(4, 1) = TxtPesNet.Text:      xCampos(4, 2) = "N":    xCampos(4, 3) = "C":     xCampos(4, 4) = "":  xCampos(4, 5) = "N"
    xCampos(5, 0) = "pestar":        xCampos(5, 1) = TxtPesTar.Text:      xCampos(5, 2) = "N":    xCampos(5, 3) = "C":     xCampos(5, 4) = "":  xCampos(5, 5) = "N"
    xCampos(6, 0) = "numlot":        xCampos(6, 1) = TxtNumLot.Text:      xCampos(6, 2) = "N":    xCampos(6, 3) = "C":     xCampos(6, 4) = "":  xCampos(6, 5) = "N"
    xCampos(7, 0) = "fchpro":        xCampos(7, 1) = TxtFchProd.Valor:    xCampos(7, 2) = "N":    xCampos(7, 3) = "F":     xCampos(7, 4) = "":  xCampos(7, 5) = "N"
    xCampos(8, 0) = "fchven":        xCampos(8, 1) = TxtFchVen.Valor:     xCampos(8, 2) = "N":    xCampos(8, 3) = "F":     xCampos(8, 4) = "":  xCampos(8, 5) = "N"
    xCampos(9, 0) = "numres":        xCampos(9, 1) = TxtNumRes.Text:      xCampos(9, 2) = "N":    xCampos(9, 3) = "C":     xCampos(9, 4) = "":  xCampos(9, 5) = "N"
    xCampos(10, 0) = "abserva":      xCampos(10, 1) = TxtObserva.Text:    xCampos(10, 2) = "N":   xCampos(10, 3) = "C":    xCampos(10, 4) = "": xCampos(10, 5) = "N"
    xCampos(11, 0) = "idform":       xCampos(11, 1) = TxtIdFormato.Text:  xCampos(11, 2) = "N":   xCampos(11, 3) = "N":    xCampos(11, 4) = "": xCampos(11, 5) = "N"
    xCampos(12, 0) = "ingredientes": xCampos(12, 1) = TxtIngre.Text:      xCampos(12, 2) = "N":   xCampos(12, 3) = "C":    xCampos(12, 4) = "": xCampos(12, 5) = "N"
    xCampos(13, 0) = "direccion":    xCampos(13, 1) = TxtDir.Text:        xCampos(13, 2) = "N":   xCampos(13, 3) = "C":    xCampos(13, 4) = "": xCampos(13, 5) = "N"
    xCampos(14, 0) = "proveedor":    xCampos(14, 1) = TxtProv.Text:       xCampos(14, 2) = "N":   xCampos(14, 3) = "C":    xCampos(14, 4) = "": xCampos(14, 5) = "N"
    
    If QueHace = 1 Then
        If EscribirNuevoRegistro(xCampos, "etiqueta", xConEtiqueta) = False Then
            xConEtiqueta.RollbackTrans
            Exit Function
        End If
    Else
        If ModificarRegistro(xCampos, "etiqueta", xConEtiqueta) = False Then
            xConEtiqueta.RollbackTrans
            Exit Function
        End If
    End If
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    MsgBox "La etiqueta se grabó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    xConEtiqueta.CommitTrans
    Grabar = True
    
    Exit Function
LaCague:
    MsgBox "Entro al Error"
    xConEtiqueta.RollbackTrans
    MsgBox "No se pudo guardar la letra por el siguiente motivo : " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Err.Clear
    
    Grabar = False
End Function

Sub MuestraSegundoTab()
    Blanquea
    
    TxtIdFormato.Text = RstLista("idform")
    TxtIdFormato_Validate True
    
    TxtTitulo.Text = RstLista("titulo1")
    TxtTitulo2.Text = NulosC(RstLista("titulo2"))
    
    If NulosN(RstLista("pesbru")) = 0 Then
        TxtPesBru.Text = ""
    Else
        TxtPesBru.Text = RstLista("pesbru")
    End If
    
    TxtPesNet.Text = RstLista("pesnet")
    
    If NulosN(RstLista("pestar")) = 0 Then
        TxtPesTar.Text = ""
    Else
        TxtPesTar.Text = RstLista("pestar")
    End If
    
    TxtNumLot.Text = RstLista("numlot")
    TxtFchProd.Valor = RstLista("fchpro")
    TxtFchVen.Valor = RstLista("fchven")
    TxtNumRes.Text = NulosC(RstLista("numres"))
    TxtObserva.Text = NulosC(RstLista("abserva"))
    TxtIngre.Text = NulosC(RstLista("ingredientes"))
    TxtDir.Text = NulosC(RstLista("direccion"))
    TxtProv.Text = NulosC(RstLista("proveedor"))
End Sub

Private Sub CmdBusTiDocEmp_Click()
    If QueHace = 3 Then Exit Sub
    
    ' CARGAMOS LA LISTA PARA BUSCAR EL NIVEL DE USUARIO
    'Dim xform As New EPS_Buscar.Buscar
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_formatos.* FROM mae_formatos"
    
    xform.Titulo = "Buscando Formatos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xConEtiqueta
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdFormato.Text = xRs("id")
        LblFormato.Caption = xRs("descripcion")
        TxtTitulo.SetFocus
        TxtIdFormato_Validate False
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Sub ImprimirFormato1()
    Dim xFila As Integer
    Dim xAltoTexto As Integer
    Dim xColAdi As Integer
    
    'Printer.PaperSize = 269
'    FrmVsPrinter.Vs.PaperHeight = 2600
'    FrmVsPrinter.Vs.PaperWidth = 4500
'
    FrmVsPrinter.Vs.BrushColor = &H80000005
    FrmVsPrinter.Vs.StartDoc
    
    FrmVsPrinter.Vs.TextAlign = taCenterMiddle
    FrmVsPrinter.Vs.FontName = "Agency FB"
    FrmVsPrinter.Vs.FontSize = 14
    FrmVsPrinter.Vs.FontBold = True
    
    xColAdi = 200
    FrmVsPrinter.Vs.TextAlign = taCenterMiddle
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo.Text), 30 + xColAdi, 100, 4000, 360, True, False, False
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo2.Text), 30 + xColAdi, 450, 4000, 360, True, False, False
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    
    FrmVsPrinter.Vs.FontSize = 12
    
    xAltoTexto = 300
    xFila = 900
    
    FrmVsPrinter.Vs.TextBox "Peso Bruto", 50 + xColAdi, xFila, 900, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextAlign = taRightMiddle
    FrmVsPrinter.Vs.TextBox TxtPesBru.Text, 950 + xColAdi, xFila, 800, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "Peso Tara", 2100 + xColAdi, xFila, 900, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtPesTar.Text, 3000 + xColAdi, xFila, 800, xAltoTexto, True, False, False
    
    xFila = xFila + 300
    FrmVsPrinter.Vs.TextBox "Peso Neto", 50 + xColAdi, xFila, 900, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextAlign = taRightMiddle
    FrmVsPrinter.Vs.TextBox TxtPesNet.Text, 950 + xColAdi, xFila, 800, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "Nº Lote   ", 2100 + xColAdi, xFila, 900, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtNumLot.Text, 2800 + xColAdi, xFila, 2000, xAltoTexto, True, False, False
    
    xFila = xFila + 300
    FrmVsPrinter.Vs.TextBox "Fch. Prod.", 50 + xColAdi, xFila, 900, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox Format(TxtFchProd.Valor, "dd/mm/yy"), 1000 + xColAdi, xFila, 1000, xAltoTexto, True, False, False

    FrmVsPrinter.Vs.TextBox "Fch. Ven. :", 2100 + xColAdi, xFila, 900, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox Format(TxtFchVen.Valor, "dd/mm/yy"), 3000 + xColAdi, xFila, 1000, xAltoTexto, True, False, False
    
    xFila = xFila + 300
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "R.S. Nº ", 50 + xColAdi, xFila, 900, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtNumRes.Text, 950 + xColAdi, xFila, 2000, xAltoTexto, True, False, False
    
    xFila = xFila + 300
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "Almacenar", 50 + xColAdi, xFila, 900, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtObserva.Text, 950 + xColAdi, xFila, 4000, xAltoTexto, True, False, False

    FrmVsPrinter.Vs.EndDoc
    FrmVsPrinter.Show vbModal
End Sub

Sub ImprimirFormato4()
    Dim xFila As Integer
    Dim xAltoTexto As Integer
    Dim xColAdi As Integer
    
    'Printer.PaperSize = 269
'    FrmVsPrinter.Vs.PaperHeight = 1500
'    FrmVsPrinter.Vs.PaperWidth = 4200

    FrmVsPrinter.Vs.BrushColor = &H80000005
    FrmVsPrinter.Vs.StartDoc
    
    FrmVsPrinter.Vs.TextAlign = taCenterMiddle
    FrmVsPrinter.Vs.FontName = "Agency FB"
    FrmVsPrinter.Vs.FontSize = 13
    FrmVsPrinter.Vs.FontBold = True
    
    xColAdi = 100
    FrmVsPrinter.Vs.TextAlign = taCenterMiddle
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo.Text), 30 + xColAdi, 90, 4000, 330, True, False, False
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo2.Text), 30 + xColAdi, 340, 4000, 330, True, False, False
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.FontSize = 10
    
    xAltoTexto = 250
    xFila = 600
    FrmVsPrinter.Vs.TextBox "Peso Neto", 50 + xColAdi, xFila, 750, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextAlign = taRightMiddle
    FrmVsPrinter.Vs.TextBox TxtPesNet.Text, 780 + xColAdi, xFila, 1200, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "Nº Lote   ", 2100 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtNumLot.Text, 2800 + xColAdi, xFila, 2000, xAltoTexto, True, False, False
    
    xFila = xFila + 170
    FrmVsPrinter.Vs.TextBox "Fch. Prod.", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtFchProd.Valor, 800 + xColAdi, xFila, 950, xAltoTexto, True, False, False

    FrmVsPrinter.Vs.TextBox "Fch. Ven. :", 2100 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtFchVen.Valor, 2800 + xColAdi, xFila, 950, xAltoTexto, True, False, False
    
    xFila = xFila + 170
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "R.S. Nº ", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtNumRes.Text, 800 + xColAdi, xFila, 4000, xAltoTexto, True, False, False
    
    xFila = xFila + 170
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "Almacenar", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtObserva.Text, 800 + xColAdi, xFila, 4000, xAltoTexto, True, False, False

    FrmVsPrinter.Vs.EndDoc
    FrmVsPrinter.Show vbModal
End Sub

Sub ImprimirFormato5()
    Dim xFila As Integer
    Dim xAltoTexto As Integer
    Dim xColAdi As Integer
    
    'Printer.PaperSize = 269
'    FrmVsPrinter.Vs.PaperHeight = 1500
'    FrmVsPrinter.Vs.PaperWidth = 6150

    FrmVsPrinter.Vs.StartDoc
    FrmVsPrinter.Vs.FontName = "Agency FB"
    
    FrmVsPrinter.Vs.BrushColor = &H80000005
    FrmVsPrinter.Vs.StartDoc
    
    FrmVsPrinter.Vs.TextAlign = taCenterMiddle
    FrmVsPrinter.Vs.FontName = "Agency FB"
    FrmVsPrinter.Vs.FontSize = 13
    FrmVsPrinter.Vs.FontBold = True
    
    xColAdi = 150
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo.Text), 50 + xColAdi, 70, 2700, 350, True, False, False
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo.Text), 3050 + xColAdi, 70, 2700, 350, True, False, False
    
    FrmVsPrinter.Vs.FontSize = 11
    FrmVsPrinter.Vs.TextBox TxtTitulo2.Text, 50 + xColAdi, 400, 2700, 350, True, False, False
    FrmVsPrinter.Vs.TextBox TxtTitulo2.Text, 3050 + xColAdi, 400, 2700, 350, True, False, False
    
    xAltoTexto = 250
    xFila = 700
    
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.FontSize = 10
    
    FrmVsPrinter.Vs.TextBox "Nº Lote   ", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtNumLot.Text, 750 + xColAdi, xFila, 2000, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextBox "Nº Lote   ", 3050 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtNumLot.Text, 3750 + xColAdi, xFila, 2000, xAltoTexto, True, False, False
    
    xFila = xFila + 200
    FrmVsPrinter.Vs.FontSize = 9
    FrmVsPrinter.Vs.TextBox "Fch. Prod.", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontSize = 8
    FrmVsPrinter.Vs.TextBox Format(TxtFchProd.Valor, "dd/mm/yyyy"), 750 + xColAdi, xFila + 20, 750, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.FontSize = 9
    FrmVsPrinter.Vs.TextBox "Fch. Prod.", 3050 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontSize = 8
    FrmVsPrinter.Vs.TextBox Format(TxtFchProd.Valor, "dd/mm/yyyy"), 3750 + xColAdi, xFila + 20, 750, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.FontSize = 9
    FrmVsPrinter.Vs.TextBox "Fch. Ven.", 1500 + xColAdi, xFila, 650, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontSize = 8
    FrmVsPrinter.Vs.TextBox Format(TxtFchVen.Valor, "dd/mm/yyyy"), 2100 + xColAdi, xFila + 20, 750, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.FontSize = 9
    FrmVsPrinter.Vs.TextBox "Fch. Ven.", 4500 + xColAdi, xFila, 650, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontSize = 8
    FrmVsPrinter.Vs.TextBox Format(TxtFchVen.Valor, "dd/mm/yyyy"), 5100 + xColAdi, xFila + 20, 750, xAltoTexto, True, False, False
    
    xFila = xFila + 200
    FrmVsPrinter.Vs.FontSize = 10
    FrmVsPrinter.Vs.TextBox "Peso Neto ", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextBox TxtPesNet.Text, 750 + xColAdi, xFila, 3000, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextBox "Peso Neto ", 3050 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextBox TxtPesNet.Text, 3750 + xColAdi, xFila, 3000, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.EndDoc
    FrmVsPrinter.Show vbModal
End Sub

Sub ImprimirFormato6()
    '--Se agrega el tamaño de papel
    '--El fomato de la fecha es dd/mm/yyyy antes dd/mm/yy
    '--Se modifica el inicio de los valores de los campos (Peso neto, Fch Prod., R.S.N, Ingrediente)
    Dim xFila As Integer
    Dim xAltoTexto As Integer
    Dim xColAdi As Integer
    
    'Printer.PaperSize = 271  ' tamaño de papel 2.5 x 7.5 cm
'    FrmVsPrinter.Vs.PaperHeight = 1500
'    FrmVsPrinter.Vs.PaperWidth = 5500

    FrmVsPrinter.Vs.StartDoc
    FrmVsPrinter.Vs.BrushColor = &H80000005
    
    FrmVsPrinter.Vs.TextAlign = taCenterMiddle
    FrmVsPrinter.Vs.FontName = "Agency FB"
    FrmVsPrinter.Vs.FontSize = 13
    FrmVsPrinter.Vs.FontBold = True
    
    xColAdi = 200
    FrmVsPrinter.Vs.TextAlign = taCenterMiddle
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo.Text), 30 + xColAdi, 70, 4000, 330, True, False, False
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo2.Text), 30 + xColAdi, 340, 4000, 330, True, False, False
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.FontSize = 10
    
    xAltoTexto = 250
    xFila = 600
    FrmVsPrinter.Vs.TextBox "Peso Neto", 50 + xColAdi, xFila, 750, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox TxtPesNet.Text, 800 + xColAdi, xFila, 1000, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "Nº Lote   ", 2100 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtNumLot.Text, 2800 + xColAdi, xFila, 2000, xAltoTexto, True, False, False
    
    xFila = xFila + 170
    FrmVsPrinter.Vs.TextBox "Fch. Prod.", 50 + xColAdi, xFila, 750, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox Format(TxtFchProd.Valor, "dd/mm/yyyy"), 800 + xColAdi, xFila, 1000, xAltoTexto, True, False, False

    FrmVsPrinter.Vs.TextBox "Fch. Ven. :", 2100 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox Format(TxtFchVen.Valor, "dd/mm/yyyy"), 2800 + xColAdi, xFila, 1000, xAltoTexto, True, False, False
    
    xFila = xFila + 170
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "R.S. Nº ", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtNumRes.Text, 800 + xColAdi, xFila, 1300, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextBox "Almacenar ", 2100 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtObserva.Text, 2800 + xColAdi, xFila, 3000, xAltoTexto, True, False, False
    
    xFila = xFila + 170
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "Ingredientes", 50 + xColAdi, xFila, 900, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtIngre.Text, 970 + xColAdi, xFila, 4000, xAltoTexto, True, False, False

    FrmVsPrinter.Vs.EndDoc
    FrmVsPrinter.Show vbModal
End Sub

Sub ImprimirFormato7()
    Dim xFila As Integer
    Dim xAltoTexto As Integer
    Dim xColAdi As Integer
    
    'Printer.PaperSize = 269
'    FrmVsPrinter.Vs.PaperHeight = 1500
'    FrmVsPrinter.Vs.PaperWidth = 6150

    FrmVsPrinter.Vs.StartDoc
    FrmVsPrinter.Vs.FontName = "Agency FB"
    
    FrmVsPrinter.Vs.BrushColor = &H80000005
    FrmVsPrinter.Vs.StartDoc
    
    FrmVsPrinter.Vs.TextAlign = taCenterMiddle
    FrmVsPrinter.Vs.FontName = "Agency FB"
    FrmVsPrinter.Vs.FontSize = 13
    FrmVsPrinter.Vs.FontBold = True
    
    xColAdi = 150
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo.Text), 50 + xColAdi, 70, 2700, 350, True, False, False
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo.Text), 3050 + xColAdi, 70, 2700, 350, True, False, False
    
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.FontSize = 10
    
    xAltoTexto = 250
    xFila = 500
    FrmVsPrinter.Vs.TextBox "Nº Lote   ", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtNumLot.Text, 750 + xColAdi, xFila, 2000, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextBox "Nº Lote   ", 3050 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtNumLot.Text, 3750 + xColAdi, xFila, 2000, xAltoTexto, True, False, False
    
    xFila = xFila + 200
    FrmVsPrinter.Vs.FontSize = 9
    FrmVsPrinter.Vs.TextBox "Fch. Prod.", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontSize = 8
    FrmVsPrinter.Vs.TextBox Format(TxtFchProd.Valor, "dd/mm/yyyy"), 750 + xColAdi, xFila + 20, 750, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.FontSize = 9
    FrmVsPrinter.Vs.TextBox "Fch. Prod.", 3050 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontSize = 8
    FrmVsPrinter.Vs.TextBox Format(TxtFchProd.Valor, "dd/mm/yyyy"), 3750 + xColAdi, xFila + 20, 750, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.FontSize = 9
    FrmVsPrinter.Vs.TextBox "Fch. Ven.", 1500 + xColAdi, xFila, 650, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontSize = 8
    FrmVsPrinter.Vs.TextBox Format(TxtFchVen.Valor, "dd/mm/yyyy"), 2100 + xColAdi, xFila + 20, 750, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.FontSize = 9
    FrmVsPrinter.Vs.TextBox "Fch. Ven.", 4500 + xColAdi, xFila, 650, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontSize = 8
    FrmVsPrinter.Vs.TextBox Format(TxtFchVen.Valor, "dd/mm/yyyy"), 5100 + xColAdi, xFila + 20, 750, xAltoTexto, True, False, False
    
    xFila = xFila + 200
    FrmVsPrinter.Vs.FontSize = 10
    FrmVsPrinter.Vs.TextBox "Peso Neto ", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextBox TxtPesNet.Text, 750 + xColAdi, xFila, 3000, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextBox "Peso Neto ", 3050 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextBox TxtPesNet.Text, 3750 + xColAdi, xFila, 3000, xAltoTexto, True, False, False
    
    xFila = xFila + 200
    FrmVsPrinter.Vs.TextBox "Proveedor ", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtProv.Text, 750 + xColAdi, xFila, 3000, xAltoTexto, True, False, False

    FrmVsPrinter.Vs.TextBox "Proveedor ", 3050 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtProv.Text, 3750 + xColAdi, xFila, 3000, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.EndDoc
    FrmVsPrinter.Show vbModal
End Sub

Sub ImprimirFormato8()
    Dim xFila As Integer
    Dim xAltoTexto As Integer
    Dim xColAdi As Integer
    
    'Printer.PaperSize = 269   ' tamaño de papel 2.5 x 10.5 cm
'    FrmVsPrinter.Vs.PaperHeight = 1450
'    FrmVsPrinter.Vs.PaperWidth = 6200
    
    FrmVsPrinter.Vs.BrushColor = &H80000005
    FrmVsPrinter.Vs.StartDoc

    FrmVsPrinter.Vs.TextAlign = taCenterMiddle
    FrmVsPrinter.Vs.FontName = "Agency FB"
    FrmVsPrinter.Vs.FontSize = 13
    FrmVsPrinter.Vs.FontBold = True
    
    xColAdi = 150
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo.Text), 50 + xColAdi, 70, 2700, 700, True, False, False
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo.Text), 3150 + xColAdi, 70, 2700, 700, True, False, False
    
    
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.FontSize = 10
    
    
    xAltoTexto = 250
    xFila = 700
    FrmVsPrinter.Vs.TextBox "Nº Lote   ", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtNumLot.Text, 750 + xColAdi, xFila, 2000, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextBox "Nº Lote   ", 3150 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox TxtNumLot.Text, 3800 + xColAdi, xFila, 2000, xAltoTexto, True, False, False
    
    
    xFila = xFila + 200
    FrmVsPrinter.Vs.TextBox "Fch. Prod.", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox Format(TxtFchProd.Valor, "dd/mm/yy"), 750 + xColAdi, xFila, 800, xAltoTexto, True, False, False

    FrmVsPrinter.Vs.TextBox "Fch. Prod.", 3150 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox Format(TxtFchProd.Valor, "dd/mm/yy"), 3850 + xColAdi, xFila, 800, xAltoTexto, True, False, False

    FrmVsPrinter.Vs.TextBox "Fch. Ven. ", 1500 + xColAdi, xFila, 650, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox Format(TxtFchVen.Valor, "dd/mm/yy"), 2200 + xColAdi, xFila, 800, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextBox "Fch. Ven. ", 4650 + xColAdi, xFila, 650, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.TextBox Format(TxtFchVen.Valor, "dd/mm/yy"), 5300 + xColAdi, xFila, 800, xAltoTexto, True, False, False
    
    xFila = xFila + 200
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "Peso Neto ", 50 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextAlign = taRightMiddle
    FrmVsPrinter.Vs.TextBox TxtPesNet.Text, 750 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "Peso Neto ", 3150 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.TextAlign = taRightMiddle
    FrmVsPrinter.Vs.TextBox TxtPesNet.Text, 3800 + xColAdi, xFila, 700, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.EndDoc
    FrmVsPrinter.Show vbModal
End Sub

' * Sub: ImprimirFormato9                        *
' *                                              *
' * Autor: Jose Luis Chacon Manrique             *
' *                                              *
' * Sub para dar formato al tipo de etiqueta 9   *

Sub ImprimirFormato9()
    Dim xFila As Integer
    Dim xColumna As Integer
    Dim vFila As Integer
    Dim vColumna As Integer
    Dim tamanioTitulo As Integer
    Dim tamanioTexto As Integer
    
    Dim xAltoTexto As Integer
    Dim xColAdi As Integer
    
    FrmVsPrinter.Vs.BrushColor = &H80000005
    
    'Tamaños para el papel
'    FrmVsPrinter.Vs.PaperHeight = 2050
'    FrmVsPrinter.Vs.PaperWidth = 6000
    
    FrmVsPrinter.Vs.StartDoc
    FrmVsPrinter.Vs.FontName = "Arial"
    
    xFila = 100 'Posicion inicial de la Fila
    xColumna = 200 'Posicion inicial de la Columna
    vFila = 150 'Variacion por Fila
    vColumna = 900 'Variacion por Columna
    tamanioTitulo = 8 'Tamaño de Letra para titulo
    tamanioTexto = 7 'Tamaño de letra para texto
    
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    'Se llena la primera etiqueta
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    
    'Mermelada de Lucuma
    FrmVsPrinter.Vs.TextAlign = taCenterMiddle
    FrmVsPrinter.Vs.FontSize = tamanioTitulo
    FrmVsPrinter.Vs.FontBold = True
    FrmVsPrinter.Vs.TextBox TxtTitulo.Text, xColumna, xFila, 2700, 200, True, False, False
    
    xFila = xFila + vFila + 75
    'Elaborado por: AGROVADO
    FrmVsPrinter.Vs.FontSize = tamanioTexto
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "Elab. por: ", xColumna, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.FontBold = True
    FrmVsPrinter.Vs.FontSize = tamanioTitulo - 1
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo2.Text), xColumna + vColumna, xFila, 800, 0, True, False, False
    
    FrmVsPrinter.Vs.FontName = "Arial"
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.FontSize = tamanioTexto
    
    'Direccion: Lima 42
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Direccion: ", xColumna, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtDir.Text, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Para: TANTA
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Para: ", xColumna, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtProv.Text, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Ingredientes: xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Ingredientes: ", xColumna, xFila, 1500, 0, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtIngre.Text, xColumna + vColumna, xFila, 2000, 0, True, False, False
    
    'R.S. Nº: xxxxxxxxxxxxxxxxxxx
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila + vFila
    FrmVsPrinter.Vs.TextBox "R.S. Nº:", xColumna, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtNumRes.Text, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Fecha de Produccion: xxxxxxxxxxxxxxxxxxx
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Fch. Prod. : ", xColumna, xFila, 1000, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtFchProd.Valor, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Fecha de Vencimiento: xxxxxxxxxxxxxxxxxxx
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Fch. Ven. :", xColumna, xFila, 1000, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtFchVen.Valor, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Peso Neto: 300
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Peso Neto: ", xColumna, xFila, 1500, 0, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtPesNet.Text & "  gr.", xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Lote: MLU-240910
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Lote:    ", xColumna, xFila, 1500, 0, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtNumLot.Text, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Observacion: Despues de abrir refrigerar
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Observación: ", xColumna, xFila, 1500, 0, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtObserva.Text, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    'Se llena la segunda etiqueta
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xFila = 100
    xColumna = 3300
    
    'Mermelada de Lucuma
    FrmVsPrinter.Vs.TextAlign = taCenterMiddle
    FrmVsPrinter.Vs.FontSize = tamanioTitulo
    FrmVsPrinter.Vs.FontBold = True
    FrmVsPrinter.Vs.TextBox TxtTitulo.Text, xColumna, xFila, 2700, 200, True, False, False
    
    xFila = xFila + vFila + 75
    'Elaborado por: AGROVADO
    FrmVsPrinter.Vs.FontSize = tamanioTexto
    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
    FrmVsPrinter.Vs.TextBox "Elab. por: ", xColumna, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.FontBold = True
    FrmVsPrinter.Vs.FontSize = tamanioTitulo - 1
    FrmVsPrinter.Vs.TextBox UCase(TxtTitulo2.Text), xColumna + vColumna, xFila, 800, 0, True, False, False
    
    FrmVsPrinter.Vs.FontName = "Arial"
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.FontSize = tamanioTexto
    
    'Direccion: Lima 42
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Direccion: ", xColumna, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtDir.Text, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Para: TANTA
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Para: ", xColumna, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtProv.Text, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Ingredientes: xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Ingredientes: ", xColumna, xFila, 1500, 0, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtIngre.Text, xColumna + vColumna, xFila, 2000, 0, True, False, False
    
    'R.S. Nº: xxxxxxxxxxxxxxxxxxx
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila + vFila
    FrmVsPrinter.Vs.TextBox "R.S. Nº:", xColumna, xFila, 700, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtNumRes.Text, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Fecha de Produccion: xxxxxxxxxxxxxxxxxxx
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Fch. Prod. : ", xColumna, xFila, 1000, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtFchProd.Valor, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Fecha de Vencimiento: xxxxxxxxxxxxxxxxxxx
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Fch. Ven. :", xColumna, xFila, 1000, xAltoTexto, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtFchVen.Valor, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Peso Neto: 300
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Peso Neto: ", xColumna, xFila, 1500, 0, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtPesNet.Text & "  gr.", xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Lote: MLU-240910
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Lote:    ", xColumna, xFila, 1500, 0, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtNumLot.Text, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    'Observacion: Despues de abrir refrigerar
    FrmVsPrinter.Vs.FontBold = True
    xFila = xFila + vFila
    FrmVsPrinter.Vs.TextBox "Observación: ", xColumna, xFila, 1500, 0, True, False, False
    FrmVsPrinter.Vs.FontBold = False
    FrmVsPrinter.Vs.TextBox TxtObserva.Text, xColumna + vColumna, xFila, 2000, xAltoTexto, True, False, False
    
    FrmVsPrinter.Vs.EndDoc
    FrmVsPrinter.Show vbModal
End Sub

Private Sub Command2_Click()
    If TabOne1.CurrTab = 0 Then TabOne1.CurrTab = 1
    
    Dim xFila As Integer

    If RstLista("idform") = 1 Then ImprimirFormato1
    If RstLista("idform") = 4 Then ImprimirFormato4
    If RstLista("idform") = 5 Then ImprimirFormato5
    If RstLista("idform") = 6 Then ImprimirFormato6
    If RstLista("idform") = 7 Then ImprimirFormato7
    If RstLista("idform") = 8 Then ImprimirFormato8
    If RstLista("idform") = 9 Then ImprimirFormato9
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstLista
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LA COLUMNAS DEL DtaGrid Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstLista.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstLista("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim xFun As New eps_librerias.FuncionesData
        Dim xCad As String
        Dim xRutaData As String
        Dim xRst As New ADODB.Recordset
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = 84
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        
        xTitulo = "Sistema Gestion Informatica"                                             ' CARGAMOS LOS TITULOS PARA LOS MSGBOX Y OTROS MENSAJES QUE EMITIERA EL SISTEMA
        
        xFun.F_BASEDATOS = Trim(AP_RUTABD) & "etiquetas.mdb"                                ' PASAMOS LA RUTA DE LA BASE DE DATOS PARA ABRIR LA CONECCION
        xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"                                       ' PASAMOS LA RUTA DEL ARCHIVO DE TRABJO DE LA BASE DE DATOS
        xFun.F_PASSWORD = Eps_Pass                                                          ' PASAMOS EL PASWORD DE LA BASE DE DATOS
        xFun.F_USUARIO = Eps_User                                                           ' PASAMOS EL USUARIO DE LA BASE DE DATOS
        xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"                                        ' PASAMOS EL NOMBRE DEL PROVEEDORE DE DATOS PARA ADO 2.5
        
        Set xConEtiqueta = xFun.AbrirConeccion                                                      ' ABRIMOS LA CONECCION DE DATOS
        Set xFun = Nothing
        
        Dim Rpta As Integer
        SeEjecuto = True
        
        RST_Busq RstLista, "SELECT etiqueta.*, etiqueta.fchpro & '' as fchpro1, etiqueta.fchven & '' as fchven1, etiqueta.numlot & '' as numlot1, etiqueta.idform & '' as idform1 " _
                         & " FROM etiqueta WHERE desactiva = 0 ORDER BY titulo1", xConEtiqueta
                         
        Set Dg1.DataSource = RstLista
        
    End If
End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Nuevo()
    xHorIni = Time
    QueHace = 1
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label13.Caption = "Agregando Etiqueta"
    PonerGris
'    Bloquea
    Blanquea
    TxtIdFormato.Locked = Not TxtIdFormato.Locked
    LblFormato.Caption = ""
    TxtIdFormato.SetFocus
End Sub

Sub Modificar()
    xHorIni = Time
    QueHace = 2
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label13.Caption = "Modificando Etiqueta"
    PonerGris
    Blanquea
    MuestraSegundoTab
    TxtIdFormato.Locked = Not TxtIdFormato.Locked
    TxtIdFormato_Validate True
    LblFormato.Caption = ""
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    
    Dg1.Columns("fchpro1").NumberFormat = FORMAT_DATE
    Dg1.Columns("fchven1").NumberFormat = FORMAT_DATE
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    TabOne1.CurrTab = 0
    Label14.Caption = "Lista de Etiquetas"
    Label13.Caption = "Detalle de la Etiqueta"
    QueHace = 3

End Sub

Sub Blanquea()
    TxtIdFormato.Text = ""
    
    TxtTitulo.Text = ""
    TxtTitulo2.Text = ""
    TxtPesBru.Text = ""
    TxtPesNet.Text = ""
    TxtPesTar.Text = ""
    TxtNumLot.Text = ""
    TxtFchProd.Valor = ""
    TxtFchVen.Valor = ""
    TxtNumRes.Text = ""
    TxtObserva.Text = ""
    TxtIngre.Text = ""
    TxtDir.Text = ""
    TxtProv.Text = ""
End Sub

Sub Bloquea(idForm As Integer)
    If idForm = 1 Then
        TxtTitulo.Locked = Not TxtTitulo.Locked
        TxtTitulo2.Locked = Not TxtTitulo2.Locked
        TxtPesBru.Locked = Not TxtPesBru.Locked
        TxtPesNet.Locked = Not TxtPesNet.Locked
        TxtPesTar.Locked = Not TxtPesTar.Locked
        TxtNumLot.Locked = Not TxtNumLot.Locked
        TxtFchProd.Locked = Not TxtFchProd.Locked
        TxtFchVen.Locked = Not TxtFchVen.Locked
        TxtNumRes.Locked = Not TxtNumRes.Locked
        TxtObserva.Locked = Not TxtObserva.Locked
    End If
    
    If idForm = 2 Then
        TxtTitulo.Locked = Not TxtTitulo.Locked
        TxtTitulo2.Locked = Not TxtTitulo2.Locked
        TxtPesNet.Locked = Not TxtPesNet.Locked
        TxtNumLot.Locked = Not TxtNumLot.Locked
        TxtFchProd.Locked = Not TxtFchProd.Locked
        TxtFchVen.Locked = Not TxtFchVen.Locked
        TxtNumRes.Locked = Not TxtNumRes.Locked
        TxtObserva.Locked = Not TxtObserva.Locked
    End If
    
    If idForm = 3 Then
        TxtTitulo.Locked = Not TxtTitulo.Locked
        TxtTitulo2.Locked = Not TxtTitulo2.Locked
        TxtPesNet.Locked = Not TxtPesNet.Locked
        TxtNumLot.Locked = Not TxtNumLot.Locked
        TxtFchProd.Locked = Not TxtFchProd.Locked
        TxtFchVen.Locked = Not TxtFchVen.Locked
        TxtNumRes.Locked = Not TxtNumRes.Locked
        TxtObserva.Locked = Not TxtObserva.Locked
        TxtIngre.Locked = Not TxtIngre.Locked
        TxtDir.Locked = Not TxtDir.Locked
    End If
    
    If idForm = 4 Then
        TxtTitulo.Locked = Not TxtTitulo.Locked
        TxtTitulo2.Locked = Not TxtTitulo2.Locked
        TxtPesNet.Locked = Not TxtPesNet.Locked
        TxtNumLot.Locked = Not TxtNumLot.Locked
        TxtFchProd.Locked = Not TxtFchProd.Locked
        TxtFchVen.Locked = Not TxtFchVen.Locked
        TxtNumRes.Locked = Not TxtNumRes.Locked
        TxtObserva.Locked = Not TxtObserva.Locked
    End If
    
    If idForm = 5 Then
        TxtTitulo.Locked = Not TxtTitulo.Locked
        TxtTitulo2.Locked = Not TxtTitulo2.Locked
        TxtPesNet.Locked = Not TxtPesNet.Locked
        TxtNumLot.Locked = Not TxtNumLot.Locked
        TxtFchProd.Locked = Not TxtFchProd.Locked
        TxtFchVen.Locked = Not TxtFchVen.Locked
    End If

    If idForm = 6 Then
        TxtTitulo.Locked = Not TxtTitulo.Locked
        TxtTitulo2.Locked = Not TxtTitulo2.Locked
        TxtPesNet.Locked = Not TxtPesNet.Locked
        TxtNumLot.Locked = Not TxtNumLot.Locked
        TxtFchProd.Locked = Not TxtFchProd.Locked
        TxtFchVen.Locked = Not TxtFchVen.Locked
        TxtNumRes.Locked = Not TxtNumRes.Locked
        TxtObserva.Locked = Not TxtObserva.Locked
        TxtIngre.Locked = Not TxtIngre.Locked
        TxtDir.Locked = Not TxtDir.Locked
    End If

    If idForm = 7 Then
        TxtTitulo.Locked = Not TxtTitulo.Locked
        TxtPesNet.Locked = Not TxtPesNet.Locked
        TxtNumLot.Locked = Not TxtNumLot.Locked
        TxtFchProd.Locked = Not TxtFchProd.Locked
        TxtFchVen.Locked = Not TxtFchVen.Locked
        TxtProv.Locked = Not TxtProv.Locked
    End If

    If idForm = 8 Then
        TxtTitulo.Locked = Not TxtTitulo.Locked
        TxtPesNet.Locked = Not TxtPesNet.Locked
        TxtNumLot.Locked = Not TxtNumLot.Locked
        TxtFchProd.Locked = Not TxtFchProd.Locked
        TxtFchVen.Locked = Not TxtFchVen.Locked
        'TxtProv.Locked = Not TxtProv.Locked
    End If

    If idForm = 0 Then
        TxtTitulo.Locked = True
        TxtTitulo2.Locked = True
        TxtPesBru.Locked = True
        TxtPesNet.Locked = True
        TxtPesTar.Locked = True
        TxtNumLot.Locked = True
        TxtFchProd.Locked = True
        TxtFchVen.Locked = True
        TxtNumRes.Locked = True
        TxtObserva.Locked = True
        TxtIngre.Locked = True
        TxtDir.Locked = True
        TxtProv.Locked = True
    End If
    
    If idForm = 9 Then
        TxtTitulo.Locked = Not TxtTitulo.Locked
        TxtTitulo2.Locked = Not TxtTitulo2.Locked
        TxtPesNet.Locked = Not TxtPesNet.Locked
        TxtNumLot.Locked = Not TxtNumLot.Locked
        TxtFchProd.Locked = Not TxtFchProd.Locked
        TxtFchVen.Locked = Not TxtFchVen.Locked
        TxtNumRes.Locked = Not TxtNumRes.Locked
        TxtObserva.Locked = Not TxtObserva.Locked
        TxtProv.Locked = Not TxtProv.Locked
        TxtDir.Locked = Not TxtDir.Locked
        TxtIngre.Locked = Not TxtIngre.Locked
    End If

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

    If Button.Index = 2 Then
        Modificar
    End If

    If Button.Index = 3 Then
        Dim Rpta As Integer
        Rpta = MsgBox("¿Esta seguro de eliminar esta etiqueta?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbYes Then
            RstLista.Delete
            RstLista.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstLista.Requery
            Dg1.Refresh
            If RstLista.RecordCount <> 0 Then
                RstLista.MoveFirst
                RstLista.Find "id=" & mIdRegistro
                If RstLista.EOF = True Then RstLista.MoveFirst
            End If
            
            Cancelar
        End If
    End If

    If Button.Index = 6 Then
        Cancelar
    End If
    
    If Button.Index = 8 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstLista.Filter = ""
    End If

    If Button.Index = 10 Then
        RstLista.Close
        Set RstLista = Nothing
        xConEtiqueta.Close
        Set xConEtiqueta = Nothing
        Unload Me
    End If
End Sub

Sub Cancelar()
    QueHace = 3
    ActivaTool
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    TxtIdFormato.Locked = Not TxtIdFormato.Locked
    Bloquea 0
    Label13.Caption = "Detalle de la Etiqueta"
End Sub

Private Sub TxtDir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdFormato_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdFormato_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTiDocEmp_Click
    End If
    If KeyCode = 117 Then
        TxtIdFormato.Text = ""
        LblFormato.Caption = ""
    End If
End Sub

Private Sub TxtIdFormato_Validate(Cancel As Boolean)
    If NulosN(TxtIdFormato.Text) = 0 Then
        LblFormato.Caption = ""
        Bloquea 0
        PonerGris
        Exit Sub
    End If
    PonerBlanco
    
    LblFormato.Caption = Busca_Codigo(NulosN(TxtIdFormato.Text), "id", "descripcion", "mae_formatos", "N", xConEtiqueta)
    
    Bloquea NulosN(TxtIdFormato.Text)
    
    If NulosN(TxtIdFormato.Text) = 1 Then
        TxtIngre.BackColor = &H8000000F
        TxtDir.BackColor = &H8000000F
        
        TxtIngre.Locked = True
        TxtDir.Locked = True
    End If
    
    If NulosN(TxtIdFormato.Text) = 2 Then
        TxtPesBru.BackColor = &H8000000F
        TxtPesTar.BackColor = &H8000000F
        TxtIngre.BackColor = &H8000000F
        TxtDir.BackColor = &H8000000F
        
        TxtPesBru.Locked = True
        TxtPesTar.Locked = True
        TxtIngre.Locked = True
        TxtDir.Locked = True
    End If

    If NulosN(TxtIdFormato.Text) = 3 Then
        TxtPesBru.BackColor = &H8000000F
        TxtPesTar.BackColor = &H8000000F
        
        TxtPesBru.Locked = True
        TxtPesTar.Locked = True
    End If

    If NulosN(TxtIdFormato.Text) = 4 Then
        TxtPesBru.BackColor = &H8000000F
        TxtPesTar.BackColor = &H8000000F
        TxtIngre.BackColor = &H8000000F
        TxtDir.BackColor = &H8000000F
        TxtProv.BackColor = &H8000000F
        
        TxtPesBru.Locked = True
        TxtPesTar.Locked = True
        TxtIngre.Locked = True
        TxtDir.Locked = True
        TxtProv.Locked = True
    End If

    If NulosN(TxtIdFormato.Text) = 5 Then
        'TxtTitulo2.BackColor = &H8000000F
        TxtPesBru.BackColor = &H8000000F
        TxtPesTar.BackColor = &H8000000F
        TxtNumRes.BackColor = &H8000000F
        TxtObserva.BackColor = &H8000000F
        TxtIngre.BackColor = &H8000000F
        TxtDir.BackColor = &H8000000F
        TxtProv.BackColor = &H8000000F
        
        'TxtTitulo2.Locked = True
        TxtPesBru.Locked = True
        TxtPesTar.Locked = True
        TxtObserva.Locked = True
        TxtIngre.Locked = True
        TxtDir.Locked = True
        TxtProv.Locked = True
    End If

    If NulosN(TxtIdFormato.Text) = 6 Then
        TxtPesBru.BackColor = &H8000000F
        TxtPesTar.BackColor = &H8000000F
        TxtProv.BackColor = &H8000000F
        
        TxtPesBru.Locked = True
        TxtPesTar.Locked = True
        TxtProv.Locked = True
    End If

    If NulosN(TxtIdFormato.Text) = 7 Then
        TxtTitulo2.BackColor = &H8000000F
        TxtPesBru.BackColor = &H8000000F
        TxtPesTar.BackColor = &H8000000F
        TxtNumRes.BackColor = &H8000000F
        TxtObserva.BackColor = &H8000000F
        TxtIngre.BackColor = &H8000000F
        TxtDir.BackColor = &H8000000F
        
        TxtTitulo2.Locked = True
        TxtPesBru.Locked = True
        TxtPesTar.Locked = True
        TxtObserva.Locked = True
        TxtIngre.Locked = True
        TxtDir.Locked = True
    End If

    If NulosN(TxtIdFormato.Text) = 8 Then
        TxtTitulo2.BackColor = &H8000000F
        TxtPesBru.BackColor = &H8000000F
        TxtPesTar.BackColor = &H8000000F
        TxtNumRes.BackColor = &H8000000F
        TxtObserva.BackColor = &H8000000F
        TxtIngre.BackColor = &H8000000F
        TxtDir.BackColor = &H8000000F
        TxtProv.BackColor = &H8000000F
        
        TxtTitulo2.Locked = True
        TxtPesBru.Locked = True
        TxtPesTar.Locked = True
        TxtObserva.Locked = True
        TxtIngre.Locked = True
        TxtDir.Locked = True
        TxtProv.Locked = True
    End If
    
    If NulosN(TxtIdFormato.Text) = 9 Then
'        TxtTitulo2.BackColor = &H8000000F
        TxtPesBru.BackColor = &H8000000F
        TxtPesTar.BackColor = &H8000000F
'        TxtNumRes.BackColor = &H8000000F
'        TxtObserva.BackColor = &H8000000F
'        TxtIngre.BackColor = &H8000000F
'        TxtDir.BackColor = &H8000000F
'        TxtProv.BackColor = &H8000000F
'
'        TxtTitulo2.Locked = True
'        TxtPesBru.Locked = True
'        TxtPesTar.Locked = True
'        TxtObserva.Locked = True
'        TxtIngre.Locked = True
'        TxtDir.Locked = True
'        TxtProv.Locked = True
    End If

End Sub

Sub PonerGris()
    TxtTitulo.BackColor = &H8000000F
    TxtTitulo2.BackColor = &H8000000F
    TxtPesBru.BackColor = &H8000000F
    TxtPesNet.BackColor = &H8000000F
    TxtPesTar.BackColor = &H8000000F
    TxtNumLot.BackColor = &H8000000F
    TxtFchProd.BackColor = &H8000000F
    TxtFchVen.BackColor = &H8000000F
    TxtNumRes.BackColor = &H8000000F
    TxtObserva.BackColor = &H8000000F
    TxtIngre.BackColor = &H8000000F
    TxtDir.BackColor = &H8000000F
    TxtProv.BackColor = &H8000000F
End Sub

Sub PonerBlanco()
    TxtTitulo.BackColor = &H80000005
    TxtTitulo2.BackColor = &H80000005
    TxtPesBru.BackColor = &H80000005
    TxtPesNet.BackColor = &H80000005
    TxtPesTar.BackColor = &H80000005
    TxtNumLot.BackColor = &H80000005
    TxtFchProd.BackColor = &H80000005
    TxtFchVen.BackColor = &H80000005
    TxtNumRes.BackColor = &H80000005
    TxtObserva.BackColor = &H80000005
    TxtIngre.BackColor = &H80000005
    TxtDir.BackColor = &H80000005
    TxtProv.BackColor = &H80000005
End Sub

Private Sub TxtIngre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumLot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumRes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtObserva_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtPesBru_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtPesNet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtPesTar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTitulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTitulo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub




