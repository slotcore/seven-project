VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmVerInforme 
   Caption         =   "Contabilidad - Informes Financieros"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   12465
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5760
      Left            =   60
      TabIndex        =   27
      Top             =   1110
      Width           =   11760
      _cx             =   20743
      _cy             =   10160
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
      Appearance      =   0
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "   Informe   | Detalle por Rubros "
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
      Begin VSFlex7Ctl.VSFlexGrid fg1 
         Height          =   5400
         Left            =   -12345
         TabIndex        =   28
         Top             =   15
         Width           =   11730
         _cx             =   20690
         _cy             =   9525
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
         BackColorFixed  =   -2147483643
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmVerInforme.frx":0000
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
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   5400
         Left            =   15
         TabIndex        =   29
         Top             =   15
         Width           =   11730
         _cx             =   20690
         _cy             =   9525
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmVerInforme.frx":0214
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
   Begin VB.Frame Frame6 
      Caption         =   "[ Seleccionar Informe ]"
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
      Height          =   735
      Left            =   4620
      TabIndex        =   22
      Top             =   345
      Width           =   4155
      Begin VB.CommandButton CmdBusInforme 
         Height          =   225
         Left            =   3870
         Picture         =   "FrmVerInforme.frx":02D6
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   300
         Width           =   210
      End
      Begin VB.TextBox TxtInforme 
         Height          =   300
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "TxtInforme"
         Top             =   270
         Width           =   3480
      End
      Begin VB.Label LblNumCol 
         Caption         =   "LblNumCol"
         Height          =   225
         Left            =   1800
         TabIndex        =   26
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label LblIdInforme 
         AutoSize        =   -1  'True
         Caption         =   "LblIdInforme"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3870
         TabIndex        =   25
         Top             =   750
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Informe"
         Height          =   195
         Left            =   60
         TabIndex        =   24
         Top             =   360
         Width           =   525
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "[ Expresado en ]"
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
      Height          =   735
      Left            =   8790
      TabIndex        =   14
      Top             =   345
      Width           =   3015
      Begin VB.CommandButton CmdBusMon 
         Height          =   230
         Left            =   945
         Picture         =   "FrmVerInforme.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   315
         Width           =   210
      End
      Begin VB.TextBox TxtIdMon 
         Height          =   300
         Left            =   690
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "TxtIdMon"
         Top             =   285
         Width           =   495
      End
      Begin VB.Label LblMoneda 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblMoneda"
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
         Left            =   1200
         TabIndex        =   17
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label LblTipCam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda"
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   16
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "[ Consulta ]"
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
      Height          =   735
      Left            =   30
      TabIndex        =   18
      Top             =   345
      Width           =   1245
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Por Fecha"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Por Periodo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   19
         Top             =   480
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Seleccionar ]"
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
      Height          =   735
      Left            =   1290
      TabIndex        =   6
      Top             =   345
      Width           =   3315
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   405
         Left            =   60
         TabIndex        =   9
         Top             =   210
         Visible         =   0   'False
         Width           =   3180
         Begin VB.CommandButton cmd_periodo1 
            Height          =   240
            Left            =   1290
            Picture         =   "FrmVerInforme.frx":053A
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   60
            Width           =   270
         End
         Begin VB.CommandButton cmd_periodo2 
            Height          =   240
            Left            =   2820
            Picture         =   "FrmVerInforme.frx":08BC
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   60
            Width           =   270
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "A"
            Height          =   195
            Left            =   1620
            TabIndex        =   13
            Top             =   120
            Width           =   105
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   0
            TabIndex        =   12
            Top             =   120
            Width           =   210
         End
         Begin VB.Label LblPerFin 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblPerFin"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1770
            TabIndex        =   11
            Top             =   30
            Width           =   1365
         End
         Begin VB.Label LblPerIni 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblPerIni"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   10
            Top             =   30
            Width           =   1365
         End
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   360
         TabIndex        =   0
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
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
         Valor           =   "25/04/2008"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   1980
         TabIndex        =   2
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
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
         Valor           =   "25/04/2008"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Left            =   60
         TabIndex        =   8
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   1785
         TabIndex        =   7
         Top             =   345
         Width           =   135
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4770
      Top             =   -180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":0C3E
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":1182
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":1514
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":1698
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":1AEC
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":1C04
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":2148
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":268C
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":27A0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":28B4
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":2D08
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":2E74
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":33BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":36D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":3A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerInforme.frx":3DFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar Asiento"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Configurar Formatos"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmVerInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--modificado
'Fecha      |Por            |Procedimiento      |Descripcion
'=========  ==============  =================== ===============
'11/11/09   |JCastro        |Cargar             |Mejorar informe de Balance
'                                               rubro Sobrejiro Bancario
'
'



Option Explicit
Dim RstTmp As New ADODB.Recordset
Dim SeEjecuto As Boolean

Dim mMesIni As Integer
Dim mMesFin As Integer
Dim BAND_INTERRUMPIR As Boolean '--interrumpir el procesos de la consulta

Dim mPosRegistro As Integer '--indica la posicion del numero de registro
 
Dim RstCpto As New ADODB.Recordset
Dim RstCtas As New ADODB.Recordset '--almacenara lista de cuentas relacionadas del informe

Dim Formula As New CProcessor

Private Sub cmd_periodo1_Click()
    mMesIni = SeleccionaMes(xCon)
    LblPerIni.Caption = Busca_Codigo(mMesIni, "id", "descripcion", "con_meses", "N", xCon)
End Sub

Private Sub cmd_periodo2_Click()
    mMesFin = SeleccionaMes(xCon)
    LblPerFin.Caption = Busca_Codigo(mMesFin, "id", "descripcion", "con_meses", "N", xCon)
End Sub



Private Sub Fg1_DblClick()
'    If fg1(Index).Row < fg1(Index).FixedRows Or Index = 3 Then Exit Sub
'    If fg1(Index).Row >= fg1(Index).Rows - 3 Then Exit Sub
'    '--mostrando la ventana del detalle
'    pHabilitarBotonEditor True, Index

End Sub

Private Sub fg2_DblClick()
'''    '--mostrar el asiento
'''    If fg2.Rows <= fg2.FixedRows Then Exit Sub
'''    Dim xfrm As New SGI2_funciones.formularios
'''    Me.MousePointer = vbHourglass
'''    xfrm.AsientoVer xCon, fg2.TextMatrix(fg2.Row, mPosRegistro)
'''    Set xfrm = Nothing
'''    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()

    If SeEjecuto = False Then
    
        SeEjecuto = True
        TabOne1.CurrTab = 0
        
        
        TxtIdMon.Text = 1
        TxtIdMon_Validate False
        
        opt_fecha(0).Value = True
        
        TxtInforme.Text = ""
        
        LblIdInforme.Caption = 0
        
        Setea
        
        PreparaRST
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    ElseIf KeyCode = vbKeyF8 Then
        pConsultar
    End If
End Sub

Private Sub Form_Load()
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    
    Frame7.BackColor = &H8000000F
    
        
    SeEjecuto = False
    
    LblPerIni.Caption = ""
    LblPerFin.Caption = ""
    
    TabOne1.CurrTab = 0
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Height > 3000 Then
        TabOne1.Top = 1110
        TabOne1.Width = Me.Width - 150
        TabOne1.Height = Me.Height - 1550
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Formula = Nothing
End Sub

Private Sub opt_fecha_Click(Index As Integer)
    If Index = 0 Then
        Frame7.Visible = False
    Else
        Frame7.Top = 240
        Frame7.Visible = True
        
    End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    BAND_INTERRUMPIR = False
    
    If Button.Index = 1 Then pConsultar
    
'    If Button.Index = 3 Then pBuscarAsiento
    
    'ExportarComprasExcel TabOne1.CurrTab
    If Button.Index = 5 Then pExportar TabOne1.CurrTab
        
    If Button.Index = 6 And TabOne1.CurrTab <> 3 Then pImprimir
    
'    If Button.Index = 7 And TabOne1.CurrTab <> 3 Then Configurar

    If Button.Index = 9 Then Unload Me

End Sub

'***********************************************************************************************

Private Sub TxtIdMon_Change()
    If Trim(TxtIdMon.Text) = "" Then LblMoneda.Caption = ""
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If NulosC(TxtIdMon.Text) <> "" Then
        LblMoneda.Caption = Busca_Codigo(TxtIdMon.Text, "id", "descripcion", "mae_moneda", "N", xCon)
        If NulosC(LblMoneda.Caption) = "" Then
            TxtIdMon.Text = ""
        End If
    End If
End Sub

Private Sub CmdBusMon_Click()
    
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":      xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":   xCampos(1, 1) = "id":              xCampos(1, 2) = "500":      xCampos(1, 3) = "N"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_moneda ORDER BY descripcion ;", xCampos(), "Buscando Moneda", "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.RecordCount = 0 Then GoTo SALIR
    TxtIdMon.Text = xRs("id") & ""
    LblMoneda.Caption = xRs("descripcion") & ""
    
    TxtInforme.SetFocus
SALIR:
    Set xRs = Nothing
End Sub

'***********************************************************************************************

Private Sub pExportar(Indice As Integer)
    Dim xFun As New SGI2_funciones.formularios
    Dim rst As New ADODB.Recordset
    If TabOne1.CurrTab = 0 Then
        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, TxtInforme.Text, "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en : " & LblMoneda.Caption, TxtInforme.Text
    ElseIf TabOne1.CurrTab = 1 Then
        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg2, TxtInforme.Text, "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en : " & LblMoneda.Caption, TxtInforme.Text
    End If
    
    
    Set xFun = Nothing
    
End Sub

'***********************************************************************************************

Private Function fValidarConsulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Function
    End If
    
    If opt_fecha(0).Value = True Then '--por fecha
        If NulosC(TxtFchIni.Valor) = "" Then
            MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        End If
        
        If NulosC(TxtFchFin.Valor) = "" Then
            MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchFin.SetFocus
            Exit Function
        End If
        
        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        End If
        
        If (Year(TxtFchIni.Valor) <> Year(TxtFchFin.Valor)) Then
            MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        ElseIf Year(TxtFchIni.Valor) <> CStr(AnoTra) Then
            MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        End If
        
    Else '--por periodo
    
        If mMesIni > mMesFin Then
            MsgBox "El periodo de inicio debe ser inferior o igual al periodo final", vbExclamation, xTitulo
            cmd_periodo1.SetFocus
            Exit Function
        End If
        
    End If
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Seleccione la Moneda", vbExclamation, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If
    
    
    If NulosN(LblIdInforme.Caption) = 0 Then
        MsgBox "Falta seleccionar el Informe", vbExclamation, xTitulo
        TxtInforme.SetFocus
        Exit Function
    End If
    
    
    fValidarConsulta = True
End Function

Private Sub pImprimir()
    Dim xMoneda As String
    Dim nPeriodo As String
    
    If opt_fecha(0).Value = True Then  '--por fecha
        If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
            nPeriodo = "Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
        Else
            nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
        End If
    Else '--por periodo
        If mMesIni = mMesFin Then
            nPeriodo = "Periodo : " & LblPerIni.Caption
        Else
            nPeriodo = "Periodo : De " + LblPerIni.Caption & " A " & LblPerFin.Caption
        End If
    End If
    
    xMoneda = LblMoneda.Caption
    
    
    On Error GoTo error
    Dim xPrint As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    If TabOne1.CurrTab = 0 Then
        xPrint.Imprimir_x_VSFlexGrid Fg1, "", "", "", False, False
    Else
        xPrint.Imprimir_x_VSFlexGrid Fg2, TxtInforme.Text, "(Expresado en " & LblMoneda.Caption & ")", nPeriodo, False, True
    End If
    
    Set xPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    MsgBox Err.Description & vbCr & Err.Source, vbCritical, xTitulo
    Err.Clear
   
End Sub

Private Sub pConsultar()
''    If fValidarConsulta() = False Then Exit Sub
    
    Cargar
      
End Sub


Private Sub Cargar()
    '===================================================================================================
    'Creado : 29/09/08 Por: Johan Castro
    'Propósito: 1.-CArgar en un temporal el resumen de todas las cuentas que se usan en le informe,
    '           este resume es idem a la hoja de trabajo.
    '           2.- CArgar los rubros con su respectivo detalle(cuenta contable), colocando el resumen por cada rubro
    '
    'Entradas:  Ninguno
    '
    'Resultados: Recordset Temporal Abierto, listo para agregar registro
    '
    'Otros: Este procedimiento servira de base para:
    '       asignar el importe a aquellos conceptos que llaman cuenta,
    '
    'Modificado :
    '11/11/09
    '        Considerar Informe de Balance; El rubro Sobregiro Bancario; Se agrega campo desctabal en con_planctas y en con_concepto
    '        para las cuentas que este como activo y pasivo se tomara en cuenta el saldo opuesto a su naturaleza
    '        Ej. cta 1040101 Cta Bco Continental 100018974 saldo= -118,117.38;
    '        su naturaleza es Deudor para este caso sera Acreedor
    '
    '===================================================================================================
    
    
    
    'LEYENDA:
    'SI: Saldos Iniciales
    'SP: Saldos del Movimiento del Periodo
    'SF: Saldos Finales
    'CB: Cuentas de Balance
    'CT: Cuentas de Transferencia
    'GN: Ganancias por Naturaleza
    'GF: Ganancias por Funcion
    'MP: Movimientos del Periodo
    'SM: Sumas del Mayor
    
    
    Dim nSQl As String
    Dim RstRubroCtas As New ADODB.Recordset '--cargar lista de cuenta segun rubro
    Dim nSQLAjuste  As String '--sentencia sql para considera los registros del diario se ajuste por diferencia de cambio
    Dim nSQLIdCpto  As String '--sentencia sql que almacenara la relacion de conceptos utilizados en el informe
                              '--esto servira para filtrar las cuentas utilizadas en estos conceptos
    Dim nSQLCierre As String '--sentencia sql para no mostrar el cierre
    
    Dim mRowIncio As Long '--Almacenara la fila inicial cuando se agregara detalle de un rubro, esto sera util
                          '--para hacer las sumatorias por rubro
    Dim A As Integer
    Dim FchIni, FchFin As Date '--fecha que seleccionara el usuario para mostrar el reporte
    
    Set RstCtas = Nothing
    
    '--cargar lista de conceptos que seran utilizados en el informe incluye ademas los conceptos utilizados en formula
    CargarListaCptos
    
    '--cargar los conceptos
    nSQLIdCpto = RstRegistroGenerarId(RstCpto, "id", "origen", "IN", True)
    If nSQLIdCpto <> "" Then nSQLIdCpto = " AND con_concepto." & Trim(nSQLIdCpto)
    
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " AND (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) "
    '-----------------------------------------------
    nSQLCierre = " AND (con_diario.idmes<>13) "
    '-----------------------------------------------

    Fg2.Rows = Fg2.FixedRows
    DoEvents
    
    '--consulta general

    nSQl = "SELECT con_concepto.id as idcpto,con_planctas.id AS IdCta ,con_planctas.cuenta, con_planctas.descripcion,con_planctas.tipsal, " _
        + vbCr + " IIf(((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol)))>0,((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))),0) AS SIDeb, " _
        + vbCr + " IIf(((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol)))>0,((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))),0) AS SIHab, " _
        + vbCr + " IIf(MovPeriodo.DebSol Is Null,0,MovPeriodo.DebSol) AS MPDeb, " _
        + vbCr + " IIf(MovPeriodo.HabSol Is Null,0,MovPeriodo.HabSol) AS MPHab, " _
        + vbCr + " [SIDeb]+[MPDeb] AS SMDeb, " _
        + vbCr + " [SIHab]+[MPHab] AS SMHab, " _
        + vbCr + " IIf((SMDeb-SMHab)>0,(SMDeb-SMHab),0) AS SADeb, " _
        + vbCr + " IIf((SMHab-SMDeb)>0,(SMHab-SMDeb),0) AS SAHab, " _
        + vbCr + " IIF(con_planctas.tipsal = 'D', IIF(SaldosIni.DebSol IS NULL,0,SaldosIni.DebSol) - IIF(SaldosIni.HabSol IS NULL,0,SaldosIni.HabSol),IIF(SaldosIni.HabSol IS NULL,0,SaldosIni.HabSol)- IIF(SaldosIni.DebSol IS NULL,0,SaldosIni.DebSol) ) as SI , " _
        + vbCr + " IIF(con_planctas.tipsal = 'D', IIF(MovPeriodo.DebSol IS NULL,0,MovPeriodo.DebSol) - IIF(MovPeriodo.HabSol IS NULL,0,MovPeriodo.HabSol),IIF(MovPeriodo.HabSol IS NULL,0,MovPeriodo.HabSol)- IIF(MovPeriodo.DebSol IS NULL,0,MovPeriodo.DebSol)) as SP , " _
        + vbCr + " SI + SP as SF, " _
        + vbCr + " IIf(con_planctas.iddes=1 Or con_planctas.iddes2=1,-1,0) AS CB, " _
        + vbCr + " IIf(con_planctas.iddes=4 Or con_planctas.iddes2=4,-1,0) AS CT, " _
        + vbCr + " IIf(con_planctas.iddes=2 Or con_planctas.iddes2=2,-1,0) AS GN, " _
        + vbCr + " IIf(con_planctas.iddes=3 Or con_planctas.iddes2=3,-1,0) As GF "



    If NulosN(TxtIdMon.Text) = 2 Then
        nSQl = Replace(nSQl, "DebSol", "DebDol")
        nSQl = Replace(nSQl, "HabSol", "HabDol")
    End If
    '--movimientos del periodo
    nSQl = nSQl _
        + vbCr + " FROM (  ((con_planctas INNER JOIN con_conceptodet ON con_planctas.id = con_conceptodet.idref) INNER JOIN con_concepto ON con_conceptodet.idcpto = con_concepto.id)  " _
        + vbCr + " LEFT JOIN " _
        + vbCr + " ( " _
        + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
        + vbCr + " FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id=con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc=con_tc.fecha " _
        + vbCr + " WHERE (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " & nSQLAjuste _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
        + vbCr + " ORDER BY con_planctas.cuenta " _
        + vbCr + " ) AS MovPeriodo ON con_planctas.id = MovPeriodo.IdCta) " _
        + vbCr + " Left Join "
    
    '--saldos iniciales
    nSQl = nSQl _
        + vbCr + " ( " _
        + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
        + vbCr + " FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
        + vbCr + " WHERE  ((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')) " & nSQLAjuste & nSQLCierre _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
        + vbCr + " ORDER BY con_planctas.cuenta " _
        + vbCr + " ) AS SaldosIni "
    

    nSQl = nSQl _
        + vbCr + " ON con_planctas.id = SaldosIni.IdCta " _
        + vbCr + " WHERE con_concepto.origen=0 " & nSQLIdCpto _
        + vbCr + " ORDER BY con_planctas.cuenta; "

    '--si seleccionar por periodo
    If opt_fecha(1).Value = True Then
        '--movimiento del periodo
        nSQl = Replace(nSQl, "(((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))", "( con_diario.idmes>=" & mMesIni & " And con_diario.idmes <= " & mMesFin & " )")
        '--saldos iniciales
        nSQl = Replace(nSQl, "((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "'))", "con_diario.idmes < " & mMesIni)
        
    End If
    
    '--limpiar el rst de cuentas
    Set RstCtas = Nothing
    
    Me.MousePointer = vbHourglass
    DoEvents
    RST_Busq RstCtas, nSQl, xCon
    
    '*************************************************************************************************************
    
    '--cargar detalle de rubros
    Fg2.Rows = Fg2.FixedRows
    Fg2.Rows = Fg2.Rows + 1
    DoEvents
    
    '--filtrar conceptos que invocan cuentas contables; los que estan relacionados con formulas se mostraran al final
    RstCpto.Filter = "origen=0"
    '-------
    
    If RstCpto.RecordCount = 0 Then
        MsgBox "El informe " & TxtInforme.Text & vbCr & "No tiene conceptos relacionados, vuelva a configurar el informe", vbExclamation, xTitulo
        Me.MousePointer = vbDefault
        TxtInforme.SetFocus
        Exit Sub
    Else
        RstCpto.MoveFirst
    End If
    
    '--Recorremos todos los conceptos que estan asociados en el informe y por cada concepto mostrara la lista
    '--de cuentas contables utilizadas(Todo en Pestaña: Detalle por rubros del informe)
    Do While Not RstCpto.EOF
        '--origen: 0= cuenta; 1= formula
        
        FORMATO_CELDA Fg2, Fg2.Rows - 1, 1, , True, , NulosC(RstCpto("variable")) '"Rubro =>>"
        FORMATO_CELDA Fg2, Fg2.Rows - 1, 2, , True, , NulosC(RstCpto("descripcion"))
        
        If RstCpto("origen") = 0 Then '--Si es cuenta
            '--Limpiando rst temporal
            Set RstRubroCtas = Nothing
                      
            '--Cargar lista de cuentas contables usados en el concepto
            nSQl = "SELECT con_planctas.id, con_planctas.cuenta, con_planctas.descripcion ,con_planctas.tipsal,con_planctas.desctabal,con_concepto.desctabal as desctabalcpto,con_concepto.idcat " _
                + vbCr + " FROM (con_conceptodet INNER JOIN con_planctas ON con_conceptodet.idref = con_planctas.id) INNER JOIN con_concepto ON con_conceptodet.idcpto = con_concepto.id " _
                + vbCr + " WHERE (((con_conceptodet.IdCpto) = " & NulosN(RstCpto("id")) & ")) " _
                + vbCr + " ORDER BY con_planctas.cuenta;"
            
            '--cargando rst temporal
            RST_Busq RstRubroCtas, nSQl, xCon
            
            If RstRubroCtas.RecordCount <> 0 Then
                '--almacenar el inicio de la fila para luego considerar en la suma de los totales por rubro
                mRowIncio = Fg2.Rows
                '-------
                
                RstRubroCtas.MoveFirst
                Do While Not RstRubroCtas.EOF
                                
                    '--agregar cuenta
                    Fg2.AddItem ""
                    Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(RstRubroCtas("cuenta"))
                    Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(RstRubroCtas("descripcion"))
                    Fg2.TextMatrix(Fg2.Rows - 1, 3) = NulosC(RstRubroCtas("tipsal"))
                        
                    '--buscar si cuenta en lista resumen de cuenta
                    RstCtas.Filter = "idcta=" & RstRubroCtas("id")
                    
                    If RstCtas.RecordCount <> 0 Then
                        
                        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(NulosN(RstCtas("sideb")), FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(NulosN(RstCtas("sihab")), FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(NulosN(RstCtas("mpdeb")), FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 7) = Format(NulosN(RstCtas("mphab")), FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 8) = Format(NulosN(RstCtas("smdeb")), FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 9) = Format(NulosN(RstCtas("smhab")), FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 10) = Format(NulosN(RstCtas("sadeb")), FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 11) = Format(NulosN(RstCtas("sahab")), FORMAT_MONTO)
                                            
                        Fg2.TextMatrix(Fg2.Rows - 1, 12) = Format(NulosN(RstCtas("si")), FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 13) = Format(NulosN(RstCtas("sp")), FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 14) = Format(NulosN(RstCtas("sf")), FORMAT_MONTO)
                        
                        Fg2.TextMatrix(Fg2.Rows - 1, 15) = Format(NulosN(RstCtas("idcta")), FORMAT_MONTO)
                        
                    Else
                        '--colocar por defecto valor a cero
                        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(0, FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(0, FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(0, FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 7) = Format(0, FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 8) = Format(0, FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 9) = Format(0, FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 10) = Format(0, FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 11) = Format(0, FORMAT_MONTO)
                        
                        Fg2.TextMatrix(Fg2.Rows - 1, 12) = Format(0, FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 13) = Format(0, FORMAT_MONTO)
                        Fg2.TextMatrix(Fg2.Rows - 1, 14) = Format(0, FORMAT_MONTO)
                        
                    End If
                    '-------------------------------------------------------------------------------------
                    '--verificar si se coloca en activo o pasivo segun saldo
                    '--considerar solo cuando el rubro es para balance
                    If NulosN(RstRubroCtas("desctabal")) = 3 And NulosN(RstRubroCtas("idcat")) = 1 Then
                        '--activo
                        If NulosN(RstRubroCtas("desctabalcpto")) = 1 Then
                            If NulosC(RstRubroCtas("tipsal")) = "D" And NulosN(RstCtas("sf")) >= 0 Then
                            '--no hacer nada, esta ok
                            
                            ElseIf NulosC(RstRubroCtas("tipsal")) = "D" And NulosN(RstCtas("sf")) < 0 Then
                            '--eliminar registro
                                Fg2.RemoveItem Fg2.Rows - 1
                            
                            ElseIf NulosC(RstRubroCtas("tipsal")) = "H" And NulosN(RstCtas("sf")) < 0 Then
                            '--cambiar signo
                                Fg2.TextMatrix(Fg2.Rows - 1, 12) = Format(NulosN(Abs(Fg2.TextMatrix(Fg2.Rows - 1, 12))), FORMAT_MONTO)
                                Fg2.TextMatrix(Fg2.Rows - 1, 13) = Format(NulosN(Abs(Fg2.TextMatrix(Fg2.Rows - 1, 13))), FORMAT_MONTO)
                                Fg2.TextMatrix(Fg2.Rows - 1, 14) = Format(NulosN(Abs(Fg2.TextMatrix(Fg2.Rows - 1, 14))), FORMAT_MONTO)
                            
                            ElseIf NulosC(RstRubroCtas("tipsal")) = "H" And NulosN(RstCtas("sf")) >= 0 Then
                            '--eliminar registro
                                Fg2.RemoveItem Fg2.Rows - 1
                            
                            End If
                            
                        '--pasivo y patrimonio
                        ElseIf NulosN(RstRubroCtas("desctabalcpto")) = 2 Then
                            If NulosC(RstRubroCtas("tipsal")) = "D" And NulosN(RstCtas("sf")) < 0 Then
                            '--cambiar signo
                                Fg2.TextMatrix(Fg2.Rows - 1, 12) = Format(NulosN(Abs(Fg2.TextMatrix(Fg2.Rows - 1, 12))), FORMAT_MONTO)
                                Fg2.TextMatrix(Fg2.Rows - 1, 13) = Format(NulosN(Abs(Fg2.TextMatrix(Fg2.Rows - 1, 13))), FORMAT_MONTO)
                                Fg2.TextMatrix(Fg2.Rows - 1, 14) = Format(NulosN(Abs(Fg2.TextMatrix(Fg2.Rows - 1, 14))), FORMAT_MONTO)
                                
                            ElseIf NulosC(RstRubroCtas("tipsal")) = "D" And NulosN(RstCtas("sf")) >= 0 Then
                            '--eliminar registro
                                Fg2.RemoveItem Fg2.Rows - 1
                                
                            ElseIf NulosC(RstRubroCtas("tipsal")) = "H" And NulosN(RstCtas("sf")) >= 0 Then
                            '--no hacer nada, esta ok
                            
                            ElseIf NulosC(RstRubroCtas("tipsal")) = "H" And NulosN(RstCtas("sf")) < 0 Then
                            '--eliminar registro
                                Fg2.RemoveItem Fg2.Rows - 1
                                
                            End If
                            
                        Else
                            
                        End If
                    End If
                    '-------------------------------------------------------------------------------------
                    
                    RstRubroCtas.MoveNext
                Loop
                
            Else
                mRowIncio = Fg2.Rows - 1
                
            End If
            
            '--colocar la sumatoria del detalle de los rubros
            Fg2.AddItem ""
            
            FORMATO_CELDA Fg2, Fg2.Rows - 1, 2, , True, , "Totales =>>"
            
            For A = 4 To 14
                FORMATO_CELDA Fg2, Fg2.Rows - 1, A, , True, , Format(GRID_SUMAR_COL(Fg2, A, mRowIncio, Fg2.Rows - 2), FORMAT_MONTO)
            Next A
            
            
            '-----------------------------
            mRowIncio = 0
            '-----------------------------
                        
            '--si termina de mostrar el detalle de las cuentas
            '--actualizar los datos el importe del rubro en campo formula para que tome como constante
            RstCpto("valor") = NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 14))
            'RstCpto("valor") = sSaldoF
            
           
        Else
            '--cuando es formula
            Fg2.AddItem ""
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(RstCpto("formula"))
            
        End If

        RstCpto.MoveNext
        If RstCpto.EOF = False Then Fg2.Rows = Fg2.Rows + 2
        
    Loop
    
    '--ajustar las columnas segun ancho de numeros
    For A = 4 To 14
        Fg2.AutoSize A
    Next A
    
    '*************************************************************************************************************
    '--proceder a hacer los calculos de las formulas
    Aplicar
    
    '--mostrar los valores en la presentacion preliminar
    '--luego de hacer los calculos segun las formulas de los conceptos
    PonerDatosEnPresentacion
    
SALIR:
    Me.MousePointer = vbDefault
    
End Sub

Sub PreparaRST()
    '===================================================================================================
    'Creado : 20/09/08 Por: Johan Castro
    'Propósito: Crear Recordset Temporal que almacenara los conceptos utilizados en el informe
    '
    'Entradas:  Ninguno
    '
    'Resultados: Recordset Temporal Abierto, listo para agregar registro
    
    'Modificado :
    
    '===================================================================================================
       
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    ReDim xCampos(8, 3) As String

    xCampos(0, 0) = "id":           xCampos(0, 1) = "N":      xCampos(0, 2) = "3"
    xCampos(1, 0) = "descripcion":  xCampos(1, 1) = "C":      xCampos(1, 2) = "240"
    xCampos(2, 0) = "variable":     xCampos(2, 1) = "C":      xCampos(2, 2) = "240"
    xCampos(3, 0) = "formula":      xCampos(3, 1) = "C":      xCampos(3, 2) = "240"
    xCampos(4, 0) = "origen":       xCampos(4, 1) = "N":      xCampos(4, 2) = "2"
    xCampos(5, 0) = "enlista":      xCampos(5, 1) = "N":      xCampos(5, 2) = "2"
    xCampos(6, 0) = "valor":        xCampos(6, 1) = "D":      xCampos(6, 2) = "2"
    xCampos(7, 0) = "desctabal":    xCampos(7, 1) = "N":      xCampos(7, 2) = "2"

    Set RstCpto = xFun.CrearRstTMP(xCampos)
    RstCpto.Open
    
    
    
End Sub

'***********************************************************************************************


Private Sub CmdBusInforme_Click()
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
   
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "5500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM con_informe  ORDER BY descripcion ", xCampos(), "Buscando Libro Contable", "descripcion", "descripcion", Principio
    If xRs.State = 1 Then
        TxtInforme.Text = NulosC(xRs("descripcion"))
        LblIdInforme.Caption = NulosC(xRs("id"))
        '--mostrar la presentacion del informe, adicionar el titulo del reporte
        LblNumCol.Caption = NulosN(xRs("cancol"))
        PresentacionPreliminar
        
        '--poner el cursor para seleccionar moneda
        TxtIdMon.SetFocus
    End If
    Set xRs = Nothing
End Sub


Private Sub TxtInforme_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtInforme_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusInforme_Click
    End If
End Sub


'***********************************************************************************************

Private Sub PresentacionPreliminar()
    '===================================================================================================
    'Creado : 22/09/08 Por: Johan Castro
    'Propósito: Mostrar el Informe con su formato listo para colocar los importes
    '
    'Entradas:  Ninguno
    '
    'Resultados: Presentacion del Informe sin Importe
    
    'Otros: Este Procedimiento requiere primero lo Sgt.
    '           Seleccionar el tipo de informe
    '           Indicar el periodo de consulta
    
    'Modificado : fecha Por: xxxxx
    '           *****
    '===================================================================================================

    '----------------------------------
    Dim RstTmp As New ADODB.Recordset
    Dim nSQl As String
    Dim A, B, C As Long
    
    DoEvents
    
    Fg1.Rows = 1
    Fg1.Cols = 1
    
    Fg2.Rows = Fg2.FixedRows
    DoEvents
    
    '--cantidad de columnas
    '--se considera valor 5 como constante para colocar las columnas
    Fg1.Cols = 5 * NulosN(LblNumCol.Caption) + 1
    
    '--agregar fila para titulo
    
    Fg1.Rows = Fg1.Rows + 1
    GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, TxtInforme.Text, flexAlignCenterCenter, , , , , True
    '--agregar fila para periodo
    Fg1.Rows = Fg1.Rows + 1
    
    '--agregar fila para moneda
    Fg1.Rows = Fg1.Rows + 1
    GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, "(Expresado en " & LblMoneda.Caption & ")", flexAlignCenterCenter, , , , , True
    '--agregar fila en blanco
    Fg1.Rows = Fg1.Rows + 2
        
    Fg1.FrozenRows = 4
    Fg1.Col = 1
    Fg1.Row = 1
    Fg1.CellFontSize = 12
    Fg1.RowHeight(0) = 50
    Fg1.RowHeight(1) = 320
    Fg1.ColWidth(0) = 50
    
    '****************************************************
    
    '--obtener la consulta del detalle del informe
    nSQl = "SELECT con_informedet.idcpto, con_informedet.corr, con_informedet.posicion, con_informedet.idcpto, con_informedet.descripcion, con_informedet.negrita, con_concepto.formula, con_concepto.variable,con_informedet.ancho " _
        + vbCr + " FROM con_informedet LEFT JOIN con_concepto ON con_informedet.idcpto = con_concepto.id " _
        + vbCr + " WHERE (((con_informedet.idinf) = " & NulosN(LblIdInforme.Caption) & ")) " _
        + vbCr + " ORDER BY con_informedet.corr; "
    
    RST_Busq RstTmp, nSQl, xCon
    
    
    '--recorrer la cantidad de filas del grid
    For A = 0 To NulosN(LblNumCol.Caption) - 1
        '--hacer el filtro para agregar el contenido del rst al grid
        RstTmp.Filter = "posicion=" & A + 1
        If RstTmp.RecordCount <> 0 Then
            If Fg1.Rows - 1 < RstTmp.RecordCount + 5 Then
                Fg1.Rows = Fg1.Rows + RstTmp.RecordCount + 1
            End If
            RstTmp.MoveFirst
        End If
        '--posicionar en la primera fila
        C = 5
        '--agregando al grid
        Do While Not RstTmp.EOF
            Fg1.TextMatrix(C, 2 + (5 * A)) = RstTmp("idcpto")
            Fg1.TextMatrix(C, 3 + (5 * A)) = NulosC(RstTmp("variable"))
            Fg1.TextMatrix(C, 4 + (5 * A)) = NulosC(RstTmp("descripcion"))
            '--incrementar la fila
            If NulosN(RstTmp("negrita")) = 0 Then
'                FORMATO_CELDA Fg1, C, 3 + (5 * A), vbBlack, False
                FORMATO_CELDA Fg1, C, 4 + (5 * A), vbBlack, False
                FORMATO_CELDA Fg1, C, 5 + (5 * A), vbBlack, False
            Else
'                FORMATO_CELDA Fg1, C, 3 + (5 * A), vbBlack, True
                FORMATO_CELDA Fg1, C, 4 + (5 * A), vbBlack, True
                FORMATO_CELDA Fg1, C, 5 + (5 * A), vbBlack, True
            End If
            
            If NulosN(RstTmp("ancho")) <> 0 Then
                Fg1.ColWidth(4 + (5 * A)) = NulosN(RstTmp("ancho"))
            End If
            '--alineacion
            Fg1.ColAlignment(4 + (5 * A)) = flexAlignLeftCenter
            Fg1.ColAlignment(5 + (5 * A)) = flexAlignRightCenter
            
            '--ocultando columnas
            Fg1.ColWidth(1 + (5 * A)) = 300 '--1ra columna
            Fg1.ColWidth(2 + (5 * A)) = 0 '--id concepto
            Fg1.ColWidth(3 + (5 * A)) = 700 '--variable
            
            
            '--incrementar la fila
            C = C + 1
         RstTmp.MoveNext
        Loop
    Next A
    
    
    Set RstTmp = Nothing
    
    '----------------------------------

End Sub

Private Sub Aplicar()
    '===================================================================================================
    'Creado : 22/09/08 Por: Johan Castro
    'Propósito: Obtener el importe por cada concepto de la lista que esta el recordset RstCpto
    '           Luego utilizaremos esta informacion para actualizar la presentacion preliminar y mostrar
    '           los importes
    '
    'Entradas:  Ninguno
    '
    'Resultados: lista de conceptos con su respectivo valor
    
    'Otros:
    
    'Modificado : fecha Por: xxxxx
    '           *****
    '===================================================================================================
    
    Dim xBook As Variant
    
    RstCpto.Filter = ""
    RstCpto.MoveFirst
    Dim xx As String
    Do While Not RstCpto.EOF
        
        '--obtener posicion
        xBook = RstCpto.Bookmark

        If RstCpto("origen") = -1 Then RstCpto("valor") = 0
        
        RstCpto("enlista") = 0
        AplicarFormula RstCpto("id")
        RstCpto.Filter = ""
        '--restablecer posicion
        RstCpto.Bookmark = xBook
        
        RstCpto.MoveNext
    Loop

End Sub


Private Sub AplicarFormula(IdCpto As Long)
    '===================================================================================================
    'Creado : 22/09/08 Por: Johan Castro
    'Propósito: Asignar un valor al concepto; Esta se obtendra mediante una formula o una constante
    '
    'Entradas:  IdCpto: codigo del concepto a indentificar el valor
    '
    'Resultados: Concepto con un valor
    
    'Otros: Es un procedimiento recursivo; Si el concepto es un formula primero se identificara los conceptos
    '       relacionados, se obtendra su valor por cada uno (se puede aplicar recursiva si uno de los conceptos es una formula)
    '       luego se aplicara la formula
    'Modificado : fecha Por: xxxxx
    '           *****
    '===================================================================================================

    Dim RstTmp As New ADODB.Recordset
    Dim nSQl As String

    Dim Xbookmark As Variant
    
    '--proceder a hacer los calculos segun corresponda
    '--se tomara la lista de rubros en rstcpto
    '*********************************************************************************************
  RstCpto.Filter = ""
  RstCpto.Filter = "id=" & IdCpto
  
  Xbookmark = RstCpto.Bookmark
  
  If RstCpto.RecordCount = 0 Then Exit Sub
  '--si ya esta asignada como variable entonces salir
  'If RstCpto("enlista") = 1 Then Exit Sub
  
  If RstCpto("origen") = 0 Or RstCpto("enlista") = 1 Then
  
   ' If RstCpto("origen") = 0 Then
    
        '--buscar en lista el valor para actualizar en rstcpto
        '--agregando el rubro como variable para el calculo de la formula
        
        Formula.DeclareConstant(RstCpto("variable")) = NulosC(RstCpto("valor"))
        
    Else
        '--cargar lista de conceptos para evaluar
        
        nSQl = "SELECT con_concepto.id, con_conceptodet.idref, con_concepto_1.descripcion, con_concepto_1.variable, con_concepto_1.formula, con_concepto_1.origen " _
            + vbCr + " FROM (con_concepto INNER JOIN con_conceptodet ON con_concepto.id = con_conceptodet.idcpto) INNER JOIN con_concepto AS con_concepto_1 ON con_conceptodet.idref = con_concepto_1.id " _
            + vbCr + " WHERE (((con_concepto.id)=" & IdCpto & ") AND ((con_concepto.origen)=-1)); "
    
        RST_Busq RstTmp, nSQl, xCon
        
        If RstTmp.RecordCount <> 0 Then
            RstTmp.MoveFirst
            
            Do While Not RstTmp.EOF
                '--volver a entrar al procedimiento para buscar los conceptos sin formula
                AplicarFormula RstTmp("idref")
                
                RstTmp.MoveNext
            Loop
                        
            '--quitando filtro al rst temporal
            RstCpto.Filter = ""
            '--posicionando en el registro correcto
            RstCpto.Bookmark = Xbookmark
            
            '--aplicar formula
            Formula.BaseCalculation = 1
            RstCpto("valor") = Format(Formula.Calculate(NulosC(RstCpto("formula"))), "#####0.00")
            
            '--asignando el valor de la formula para que sea utilizado en otra formula
            Formula.DeclareConstant(RstCpto("variable")) = NulosN(RstCpto("valor"))
            
            
            '----
        Else
            
            RstCpto("valor") = "0"
        End If
        
        Set RstTmp = Nothing
    
    End If
  
    '--actualizar el concepto para no volver a cargar la variable!!!!
    RstCpto("enlista") = 1
  
 
        
End Sub



Private Sub PonerDatosEnPresentacion()
    '===================================================================================================
    'Creado : 22/09/08 Por: Johan Castro
    'Propósito: Mostrar los importes en la presentacion preliminar
    '
    'Entradas:  Ninguno
    '
    'Resultados: Presentacion del Informe Completo con importes
    
    'Otros: Mostrara la sgt informacion adicional.
    '       Periodo como subtitulo
    '       Moneda como subtitulo
    '       Ajustara la columna donde se muestra los importes para mostrar completo los valores
    '
    'Modificado : fecha Por: xxxxx
    '           *****
    '===================================================================================================

    Dim A, B, C As Long
    Dim nPeriodo As String
    
    If opt_fecha(0).Value = True Then  '--por fecha
        If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
            nPeriodo = "Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
        Else
            nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
        End If
    Else '--por periodo
        If mMesIni = mMesFin Then
            nPeriodo = "Periodo : " & LblPerIni.Caption
        Else
            nPeriodo = "Periodo : De " + LblPerIni.Caption & " A " & LblPerFin.Caption
        End If
    End If

    GRID_COMBINAR Fg1, 2, 1, 2, Fg1.Cols - 1, nPeriodo, flexAlignCenterCenter, , , , , True
    
    '--actualizar la moneda en el informe, si ha cambiado
    GRID_COMBINAR Fg1, 3, 1, 3, Fg1.Cols - 1, "(Expresado en " & LblMoneda.Caption & ")", flexAlignCenterCenter, , , , , True
    
    DoEvents
    
    '*********************************************************************

    '--recorrer la cantidad de filas del grid
    For A = 0 To NulosN(LblNumCol.Caption) - 1

        '--agregando al grid
        For C = 5 To Fg1.Rows - 1
            '--verificar si la celda contiene el id de un concepto
            If NulosN(Fg1.TextMatrix(C, 2 + (5 * A))) <> 0 Then
                
                Fg1.TextMatrix(C, 5 + (5 * A)) = 0
                
                '--filtrar el concepto para obtener el calculo luego de aplicar la formula
                RstCpto.Filter = "id=" & NulosN(Fg1.TextMatrix(C, 2 + (5 * A)))
                If RstCpto.RecordCount <> 0 Then
                    '--verificar si es nagativo para dar un formato especial ej. (50) en vez de -50
                    '-- otro formato "###,##0"
                    If NulosN(RstCpto("valor")) >= 0 Then
                        Fg1.TextMatrix(C, 5 + (5 * A)) = Format(NulosN(RstCpto("valor")), FORMAT_MONTO)
                    Else
                        Fg1.TextMatrix(C, 5 + (5 * A)) = "(" & Format(Abs(NulosN(RstCpto("valor"))), FORMAT_MONTO) & ")"
                    End If
                    
                End If
                
            End If

        Next C
    Next A
    
    '--ajustando las columnas
    For A = 0 To NulosN(LblNumCol.Caption) - 1
        Fg1.AutoSize 5 + (5 * A)
    Next A
    
    
    RstCpto.Filter = ""

End Sub

Private Sub CargarListaCptos()
    '===================================================================================================
    'Creado : 22/09/08 Por: Johan Castro
    'Propósito: Cargara la relacion de conceptos utilizados en el informe
    '           ya sea concepto relacionado a cuentas contable o formulas; si es lo segundo
    '           cargara todos los conceptos que dependan de una formula utilizando una recursiva
    '
    'Entradas:  Ninguno
    '
    'Resultados: Lista de conceptos en el recordset temporal listos para asignar sus valores
    
    'Otros: Es un procediento
    '
    'Modificado : fecha Por: xxxxx
    '           *****
    '===================================================================================================
    
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQl As String
    
    
    '--reiniciar el rst del concepto
    Set RstCpto = Nothing
    PreparaRST
    
'    Dim Formula As New CProcessor
    
    Formula.ReleaseMemory
    
    '--obtener todos los rubros que contiene el informe
    nSQl = "SELECT con_informedet.idcpto, con_informedet.corr, con_informedet.posicion, con_informedet.idcpto, con_informedet.descripcion, con_informedet.negrita, con_concepto.formula, con_concepto.variable,con_informedet.ancho " _
        + vbCr + " FROM con_informedet LEFT JOIN con_concepto ON con_informedet.idcpto = con_concepto.id " _
        + vbCr + " WHERE (((con_informedet.idinf) = " & NulosN(LblIdInforme.Caption) & ")) AND con_concepto.origen IS NOT NULL " _
        + vbCr + " ORDER BY con_informedet.corr; "
    
    RST_Busq RstTmp, nSQl, xCon
    
    If RstTmp.RecordCount <> 0 Then
        RstTmp.MoveFirst
    
        '--agregando al grid
         Do While Not RstTmp.EOF
             
             CargaCptos RstTmp("idcpto")
             RstTmp.MoveNext
         Loop
         
    End If
     
     Set RstTmp = Nothing
     
   '--quitar filtro
    RstCpto.Filter = ""
    Set Formula = Nothing
End Sub


Private Sub CargaCptos(IdCpto As Long)
    '===================================================================================================
    'Creado : 20/09/08 Por: Johan Castro
    'Propósito: Cargar los conceptos utilizados en el informe, si hay conceptos que son utilizados en formulas
    '           tambien seran considerados para la carga
    '
    'Entradas:  IdCpto: codigo del concepto
    '
    'Resultados: Conceptos utilizados en el informe listo para mostrar en el reporte
    
    'Otros: Este procedimiento se usara como recursiva cuando el concepto es una formula, esta invocara de nuevo
    '       a este procedimiento para buscar los conceptos involucrados
    
    'Modificado :
    
    '===================================================================================================

    Dim RstTmp As New ADODB.Recordset
    Dim nSQl As String
    
    '*********************************************************************************************
    '--buscar el concepto en mension; si esta en la lista no hacer nada, caso contrario agregar en lista
    If RstCpto.RecordCount <> 0 Then
        RstCpto.MoveFirst
        RstCpto.Find "id=" & IdCpto
    End If
    If RstCpto.EOF = True Or RstCpto.BOF = True Then
    
        Set RstTmp = Nothing
        
        nSQl = "SELECT con_concepto.id, con_concepto.descripcion, con_concepto.variable, con_concepto.formula, con_concepto.origen,con_concepto.desctabal " _
            + vbCr + " FROM con_concepto " _
            + vbCr + " WHERE (((con_concepto.id)=" & IdCpto & ")); "
            
        '--ejecutar consulta
        RST_Busq RstTmp, nSQl, xCon
        
        '--agregar al rst temporal
        RstCpto.AddNew
        RstCpto("id") = RstTmp("id")
        RstCpto("descripcion") = NulosC(RstTmp("descripcion"))
        RstCpto("variable") = NulosC(RstTmp("variable"))
        RstCpto("formula") = NulosC(RstTmp("formula"))
        RstCpto("origen") = NulosN(RstTmp("origen"))
        If IsNumeric(RstTmp("formula")) = True Then
            RstCpto("valor") = NulosN(RstTmp("formula"))
        Else
            RstCpto("valor") = 0
        End If
        RstCpto("enlista") = 0 'Nota este RstCpto es un temporal
        RstCpto("desctabal") = NulosN(RstTmp("desctabal"))
        
        RstCpto.Update
        
        Set RstTmp = Nothing

    End If

    '*********************************************************************************************
    
    
    'Cargar registros que contiene la formula (con_concepto.origen = -1)
    nSQl = "SELECT con_concepto.id, con_conceptodet.idref, con_concepto_1.descripcion, con_concepto_1.variable, con_concepto_1.formula, con_concepto_1.origen " _
        + vbCr + " FROM (con_concepto INNER JOIN con_conceptodet ON con_concepto.id = con_conceptodet.idcpto) INNER JOIN con_concepto AS con_concepto_1 ON con_conceptodet.idref = con_concepto_1.id " _
        + vbCr + " WHERE (((con_concepto.id)=" & IdCpto & ") AND ((con_concepto.origen)=-1)); "

    RST_Busq RstTmp, nSQl, xCon
    
    If RstTmp.RecordCount <> 0 Then
        RstTmp.MoveFirst
        
        Do While Not RstTmp.EOF
            '--evaluar si el concepto es utilizado para cuenta o para formula
            
            '--para cuenta
            If RstTmp("origen") = 0 Then
                '--buscar si esta en la lista de conceptos, si no esta agregar en lista
                If RstCpto.RecordCount <> 0 Then
                    RstCpto.MoveFirst
                    RstCpto.Find "id=" & RstTmp("idref")
                End If
                If RstCpto.EOF = True Or RstCpto.BOF = True Then
                    RstCpto.AddNew
                    RstCpto("id") = RstTmp("idref")
                    RstCpto("descripcion") = NulosC(RstTmp("descripcion"))
                    RstCpto("variable") = NulosC(RstTmp("variable"))
                    RstCpto("formula") = NulosC(RstTmp("formula"))
                    RstCpto("origen") = NulosN(RstTmp("origen"))
                    RstCpto.Update
                    
                End If
                
            Else
                '--para formula
                '--volver a entrar al procedimiento para buscar los conceptos sin formula
                CargaCptos RstTmp("idref")
                
            End If
            RstTmp.MoveNext
        Loop
    
    End If
    
    RstCpto.Filter = ""
    
    Set RstTmp = Nothing
        
End Sub


Sub Setea()
    '===================================================================================================
    'Creado : 02/10/08 Por: Johan Castro
    'Propósito: Configurar la presentacion del detalle de los rubros, este formato contendra lo sgte.
    '           Nro y Descripcion de Cuenta Contable
    '           Saldos Iniciales(D y H)
    '           Movimiento del Periodo(D y H)
    '           Sumas del Mayor(D y H)
    '           Saldos Finales(Solo importes)
    '
    'Entradas:  Ninguno
    '
    'Resultados: Presentacion de los rubros con formato idem a la hoja de trabajo
    
    'Otros:
    
    'Modificado :
    
    '===================================================================================================
    
    Dim A As Integer
    Dim B As Integer
    

     Fg2.GridLines = flexGridNone
     Fg2.Rows = 2
     Fg2.FixedRows = 2
     Fg2.Cols = 16
     
     Fg2.TextMatrix(0, 1) = "          1"
     Fg2.TextMatrix(1, 1) = "          1"
     Fg2.TextMatrix(0, 1) = "Nº Cuenta"
     Fg2.TextMatrix(1, 1) = "Nº Cuenta"
     Fg2.TextMatrix(0, 2) = "Descripción"
     Fg2.TextMatrix(1, 2) = "Descripción"
     
     Fg2.TextMatrix(0, 3) = "Nat."
     Fg2.TextMatrix(1, 3) = "Saldo" '--naturaleza del saldo
              
     Fg2.Redraw = False
     
     Fg2.MergeCol(0) = True
     Fg2.MergeCol(1) = True
     Fg2.MergeCol(2) = True
     Fg2.MergeCol(3) = False
     
     Fg2.MergeCells = 2
     Fg2.Redraw = True
     
     With Fg2
         .MergeCells = flexMergeFree
         .MergeRow(-1) = True
         .Cell(flexcpText, 0, 4, 0, 5) = "Saldos Iniciales"
         .Cell(flexcpText, 0, 6, 0, 7) = "Movimiento del Periodo"
         .Cell(flexcpText, 0, 8, 0, 9) = "Sumas del Mayor"
         .Cell(flexcpText, 0, 10, 0, 11) = "Saldos Finales"
         
         .Cell(flexcpText, 0, 12, 0, 14) = "Resumen de Saldos"

         .Cell(flexcpBackColor, 0, 0, Fg2.Rows - 1, Fg2.Cols - 1) = &H8000000F
         
         '--alinear las celdas
         .Row = 0
         .Col = 4
         .CellAlignment = flexAlignCenterCenter
         .Col = 6
         .CellAlignment = flexAlignCenterCenter
         .Col = 8
         .CellAlignment = flexAlignCenterCenter
         .Col = 10
         .CellAlignment = flexAlignCenterCenter
         .Col = 12
         .CellAlignment = flexAlignCenterCenter
     End With
     
    
     Fg2.ColWidth(3) = 450
     
     Fg2.ColWidth(4) = 1100
     Fg2.ColWidth(5) = 1100
     Fg2.ColWidth(6) = 1100
     Fg2.ColWidth(7) = 1100
     Fg2.ColWidth(8) = 1100
     Fg2.ColWidth(9) = 1100
     Fg2.ColWidth(10) = 1100
     Fg2.ColWidth(11) = 1100
     
     Fg2.ColWidth(12) = 1100
     Fg2.ColWidth(13) = 1100
     Fg2.ColWidth(14) = 1100
         
     Fg2.TextMatrix(1, 4) = "Debe"
     Fg2.TextMatrix(1, 5) = "Haber"
     Fg2.TextMatrix(1, 6) = "Debe"
     Fg2.TextMatrix(1, 7) = "Haber"
     Fg2.TextMatrix(1, 8) = "Debe"
     Fg2.TextMatrix(1, 9) = "Haber"
     Fg2.TextMatrix(1, 10) = "Debe"
     Fg2.TextMatrix(1, 11) = "Haber"
     
     Fg2.TextMatrix(1, 12) = "Inicial"
     Fg2.TextMatrix(1, 13) = "Mov. Perido"
     Fg2.TextMatrix(1, 14) = "Final"
     
     Fg2.TextMatrix(1, 15) = "IdCta"
     
     
     For B = 4 To 14
         Fg2.ColAlignment(B) = flexAlignRightCenter
     Next B
        
     Fg2.ColWidth(15) = 0
         

End Sub

