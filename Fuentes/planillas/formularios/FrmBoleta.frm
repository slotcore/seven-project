VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmBoleta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Boleta de Pago"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11670
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
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":277E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":2A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":2E2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":327C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":3596
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":38B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoleta.frx":3BCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Automatico"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar Excel"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "Imprimir"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6630
      Left            =   15
      TabIndex        =   0
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
      Caption         =   "  &Consulta  |   &Automático    |   &Detalle   "
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
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   6210
         Left            =   12315
         TabIndex        =   45
         Top             =   375
         Width           =   11580
         Begin VB.CommandButton cb 
            Height          =   225
            Index           =   3
            Left            =   1815
            Picture         =   "FrmBoleta.frx":3F5C
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   1035
            Width           =   195
         End
         Begin VB.TextBox txt 
            BackColor       =   &H80000009&
            Height          =   315
            Index           =   2
            Left            =   2025
            MaxLength       =   10
            TabIndex        =   67
            Tag             =   "null"
            Text            =   "txt(2)"
            Top             =   1635
            Width           =   1305
         End
         Begin VB.TextBox txt 
            BackColor       =   &H80000009&
            Height          =   315
            Index           =   1
            Left            =   1275
            MaxLength       =   4
            TabIndex        =   66
            Tag             =   "null"
            Text            =   "txt(1)"
            Top             =   1635
            Width           =   645
         End
         Begin VB.TextBox txt_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   0
            Left            =   6165
            Locked          =   -1  'True
            TabIndex        =   56
            Text            =   "txt_total(0)"
            Top             =   5880
            Width           =   1305
         End
         Begin VB.TextBox txt_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   1
            Left            =   7500
            Locked          =   -1  'True
            TabIndex        =   55
            Text            =   "txt_total(1)"
            Top             =   5880
            Width           =   1305
         End
         Begin VB.TextBox txt_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00008000&
            Height          =   300
            Index           =   2
            Left            =   8835
            Locked          =   -1  'True
            TabIndex        =   54
            Text            =   "txt_total(2)"
            Top             =   5880
            Width           =   1305
         End
         Begin VB.TextBox txt_total 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H000080FF&
            Height          =   300
            Index           =   3
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   53
            Text            =   "txt_total(3)"
            Top             =   5880
            Width           =   1305
         End
         Begin VB.CommandButton cmdManual 
            Caption         =   "&Procesar Calculo"
            Height          =   600
            Index           =   0
            Left            =   10185
            TabIndex        =   69
            Top             =   1320
            Width           =   1290
         End
         Begin VB.CommandButton cb 
            Height          =   225
            Index           =   2
            Left            =   6705
            Picture         =   "FrmBoleta.frx":408E
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   720
            Width           =   195
         End
         Begin VB.CommandButton cb 
            Height          =   225
            Index           =   0
            Left            =   1815
            Picture         =   "FrmBoleta.frx":41C0
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   705
            Width           =   195
         End
         Begin VB.CommandButton cb 
            Height          =   225
            Index           =   1
            Left            =   1815
            Picture         =   "FrmBoleta.frx":42F2
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   1350
            Width           =   195
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0000FFFF&
            Height          =   315
            Index           =   0
            Left            =   2505
            TabIndex        =   49
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   -15
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Frame Frame6 
            Caption         =   "( Periodo )"
            Height          =   720
            Left            =   9525
            TabIndex        =   47
            Top             =   135
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
               TabIndex        =   48
               Top             =   330
               Width           =   1740
            End
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   0
            Left            =   1275
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   60
            Text            =   "txt_cb(0)"
            Top             =   675
            Width           =   765
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   1
            Left            =   1275
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   65
            Text            =   "txt_cb(1)"
            Top             =   1320
            Width           =   765
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   2
            Left            =   6165
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   62
            Text            =   "txt_cb(2)"
            Top             =   690
            Width           =   765
         End
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   3570
            Index           =   0
            Left            =   15
            TabIndex        =   71
            Top             =   2040
            Width           =   11475
            _cx             =   20241
            _cy             =   6297
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
            FormatString    =   $"FrmBoleta.frx":4424
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
         Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
            Height          =   300
            Index           =   0
            Left            =   1275
            TabIndex        =   57
            Top             =   345
            Width           =   1350
            _ExtentX        =   2381
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
            Valor           =   "14/11/2008"
         End
         Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
            Height          =   300
            Index           =   1
            Left            =   3600
            TabIndex        =   58
            Top             =   345
            Width           =   1350
            _ExtentX        =   2381
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
            Valor           =   "14/11/2008"
         End
         Begin VB.Frame fra 
            Height          =   645
            Index           =   0
            Left            =   30
            TabIndex        =   46
            Top             =   5550
            Width           =   6060
            Begin VB.CommandButton cmdManual 
               Caption         =   "Eliminar"
               Enabled         =   0   'False
               Height          =   345
               Index           =   2
               Left            =   1395
               TabIndex        =   73
               ToolTipText     =   "Eliminar Personal"
               Top             =   210
               Width           =   1200
            End
            Begin VB.CommandButton cmdManual 
               Caption         =   "Agregar"
               Enabled         =   0   'False
               Height          =   345
               Index           =   1
               Left            =   60
               TabIndex        =   70
               ToolTipText     =   "Agregar Personal"
               Top             =   210
               Width           =   1200
            End
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   3
            Left            =   1275
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   63
            Text            =   "txt_cb(3)"
            Top             =   1005
            Width           =   765
         End
         Begin VB.Label LblTipCam2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cambio"
            Height          =   195
            Index           =   1
            Left            =   5535
            TabIndex        =   103
            Top             =   1050
            Width           =   1110
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
            Height          =   285
            Index           =   1
            Left            =   6930
            TabIndex        =   102
            Top             =   1005
            Width           =   1080
         End
         Begin VB.Label lblIdCtaDoc 
            AutoSize        =   -1  'True
            Caption         =   "lblIdCtaDoc"
            Height          =   195
            Left            =   4395
            TabIndex        =   99
            Top             =   1050
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label lbl_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cod(3)"
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
            Index           =   3
            Left            =   2895
            TabIndex        =   92
            Top             =   1005
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl_capt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Documento"
            Height          =   195
            Index           =   3
            Left            =   15
            TabIndex        =   91
            Top             =   1079
            Width           =   1185
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Index           =   2
            Left            =   555
            TabIndex        =   89
            Top             =   75
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Serie"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   88
            Top             =   75
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000006&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   1942
            Top             =   1755
            Width           =   45
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            Caption         =   "N° Documento"
            Height          =   195
            Index           =   1
            Left            =   15
            TabIndex        =   87
            Top             =   1725
            Width           =   1050
         End
         Begin VB.Label lblMoneda 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblMoneda"
            Height          =   195
            Left            =   1155
            TabIndex        =   86
            Top             =   75
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbl_total 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Ingreso"
            Height          =   210
            Index           =   0
            Left            =   6165
            TabIndex        =   85
            Top             =   5670
            Width           =   930
         End
         Begin VB.Label lbl_total 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Descuento"
            Height          =   210
            Index           =   1
            Left            =   7500
            TabIndex        =   84
            Top             =   5670
            Width           =   1185
         End
         Begin VB.Label lbl_total 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Aporte"
            Height          =   210
            Index           =   2
            Left            =   8835
            TabIndex        =   83
            Top             =   5670
            Width           =   870
         End
         Begin VB.Label lbl_total 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total a Pagar"
            Height          =   210
            Index           =   3
            Left            =   10170
            TabIndex        =   82
            Top             =   5670
            Width           =   960
         End
         Begin VB.Label lbl_capt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            Height          =   195
            Index           =   2
            Left            =   5535
            TabIndex        =   81
            Top             =   780
            Width           =   585
         End
         Begin VB.Label lbl_capt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proceso"
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   80
            Top             =   780
            Width           =   585
         End
         Begin VB.Label lbl_capt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Personal"
            Height          =   195
            Index           =   1
            Left            =   15
            TabIndex        =   79
            Top             =   1401
            Width           =   615
         End
         Begin VB.Label lbl_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cod(2)"
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
            Index           =   2
            Left            =   7800
            TabIndex        =   78
            Top             =   690
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl_fecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fch. Pago"
            Height          =   195
            Index           =   1
            Left            =   2805
            TabIndex        =   77
            Top             =   435
            Width           =   735
         End
         Begin VB.Label lbl_fecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fch. Emisión"
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   76
            Top             =   435
            Width           =   900
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
            Left            =   2910
            TabIndex        =   75
            Top             =   675
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cod(1)"
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
            Index           =   1
            Left            =   2895
            TabIndex        =   74
            Top             =   1320
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "codigo"
            Height          =   195
            Index           =   0
            Left            =   1980
            TabIndex        =   72
            Top             =   90
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Boleta de Pago"
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
            Height          =   285
            Left            =   15
            TabIndex        =   68
            Top             =   30
            Width           =   11550
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
            Height          =   285
            Index           =   2
            Left            =   6930
            TabIndex        =   64
            Top             =   690
            Width           =   1725
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
            Left            =   2025
            TabIndex        =   61
            Top             =   675
            Width           =   2925
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
            Height          =   285
            Index           =   1
            Left            =   2025
            TabIndex        =   59
            Top             =   1320
            Width           =   6795
         End
         Begin VB.Label lbl_cb 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb(3)"
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
            Index           =   3
            Left            =   2025
            TabIndex        =   93
            Top             =   1005
            Width           =   2925
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6210
         Left            =   -12225
         TabIndex        =   1
         Top             =   375
         Width           =   11580
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   5550
            Left            =   15
            TabIndex        =   5
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
            Columns(1).Caption=   "N°.Registro"
            Columns(1).DataField=   "registro"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Apellidos y Nombres"
            Columns(2).DataField=   "nombres"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "TDoc"
            Columns(3).DataField=   "abrev"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "N° Documento"
            Columns(4).DataField=   "numerodoc"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fch. Emi"
            Columns(5).DataField=   "fchdoc"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Fch. Pago"
            Columns(6).DataField=   "fchpago"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "M"
            Columns(7).DataField=   "simbolo"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Importe"
            Columns(8).DataField=   "imptot"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Saldo"
            Columns(9).DataField=   "impsal"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1005"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=926"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1693"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1614"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=6006"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=5927"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1111"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1032"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2646"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2566"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1984"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1905"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1826"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1746"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=847"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=767"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=1508"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1429"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=770"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=1588"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1508"
            Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=770"
            Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=74,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=32,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=29,.parent=14,.alignment=1"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=30,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=31,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=59,.parent=14,.alignment=1"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=60,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=61,.parent=17"
            _StyleDefs(76)  =   "Named:id=33:Normal"
            _StyleDefs(77)  =   ":id=33,.parent=0"
            _StyleDefs(78)  =   "Named:id=34:Heading"
            _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(80)  =   ":id=34,.wraptext=-1"
            _StyleDefs(81)  =   "Named:id=35:Footing"
            _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(83)  =   "Named:id=36:Selected"
            _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=37:Caption"
            _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(87)  =   "Named:id=38:HighlightRow"
            _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=39:EvenRow"
            _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(91)  =   "Named:id=40:OddRow"
            _StyleDefs(92)  =   ":id=40,.parent=33"
            _StyleDefs(93)  =   "Named:id=41:RecordSelector"
            _StyleDefs(94)  =   ":id=41,.parent=34"
            _StyleDefs(95)  =   "Named:id=42:FilterBar"
            _StyleDefs(96)  =   ":id=42,.parent=33"
         End
         Begin VB.Label lblperiodo 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   0
            Left            =   9450
            TabIndex        =   6
            Top             =   75
            Width           =   1770
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Boleta de Pago"
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
            TabIndex        =   2
            Top             =   30
            Width           =   11550
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6210
         Left            =   45
         TabIndex        =   3
         Top             =   375
         Width           =   11580
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   6210
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   11580
            Begin VB.CommandButton cbA 
               Height          =   225
               Index           =   4
               Left            =   6255
               Picture         =   "FrmBoleta.frx":445F
               Style           =   1  'Graphical
               TabIndex        =   104
               Top             =   660
               Width           =   195
            End
            Begin VB.CommandButton cbA 
               Height          =   225
               Index           =   3
               Left            =   1635
               Picture         =   "FrmBoleta.frx":4591
               Style           =   1  'Graphical
               TabIndex        =   94
               Top             =   1290
               Width           =   195
            End
            Begin VB.CommandButton CmdAuto 
               Caption         =   "&Procesar Cálculo"
               Enabled         =   0   'False
               Height          =   420
               Index           =   0
               Left            =   30
               TabIndex        =   15
               Top             =   5760
               Width           =   1410
            End
            Begin VB.CommandButton CmdAuto 
               Caption         =   "&Ver Detalle"
               Enabled         =   0   'False
               Height          =   420
               Index           =   1
               Left            =   1485
               TabIndex        =   14
               Top             =   5760
               Width           =   1410
            End
            Begin VB.CommandButton cbA 
               Height          =   225
               Index           =   0
               Left            =   1635
               Picture         =   "FrmBoleta.frx":46C3
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   660
               Width           =   195
            End
            Begin VB.CommandButton CmdAuto 
               Caption         =   "&Cargar Personal"
               Height          =   585
               Index           =   2
               Left            =   10200
               TabIndex        =   26
               Top             =   990
               Width           =   1320
            End
            Begin VB.Frame Frame4 
               Caption         =   "( Periodo )"
               Height          =   720
               Left            =   9510
               TabIndex        =   12
               Top             =   210
               Width           =   2010
               Begin VB.Label lblperiodo 
                  Alignment       =   2  'Center
                  Caption         =   "lblperiodo(2)"
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
                  Index           =   2
                  Left            =   120
                  TabIndex        =   13
                  Top             =   330
                  Width           =   1740
               End
            End
            Begin VB.CommandButton cbA 
               Height          =   225
               Index           =   1
               Left            =   6255
               Picture         =   "FrmBoleta.frx":47F5
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   330
               Width           =   195
            End
            Begin VB.TextBox txt_totalA 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   0
               Left            =   6165
               Locked          =   -1  'True
               TabIndex        =   11
               Text            =   "txt_totalA(0)"
               Top             =   5880
               Width           =   1305
            End
            Begin VB.TextBox txt_totalA 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H000000FF&
               Height          =   300
               Index           =   1
               Left            =   7500
               Locked          =   -1  'True
               TabIndex        =   10
               Text            =   "txt_totalA(1)"
               Top             =   5880
               Width           =   1305
            End
            Begin VB.TextBox txt_totalA 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H00008000&
               Height          =   300
               Index           =   2
               Left            =   8835
               Locked          =   -1  'True
               TabIndex        =   9
               Text            =   "txt_totalA(2)"
               Top             =   5880
               Width           =   1305
            End
            Begin VB.TextBox txt_totalA 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H000080FF&
               Height          =   300
               Index           =   3
               Left            =   10170
               Locked          =   -1  'True
               TabIndex        =   8
               Text            =   "txt_totalA(3)"
               Top             =   5880
               Width           =   1305
            End
            Begin VB.CommandButton cbA 
               Height          =   225
               Index           =   2
               Left            =   1635
               Picture         =   "FrmBoleta.frx":4927
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   975
               Width           =   195
            End
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   3990
               Index           =   1
               Left            =   45
               TabIndex        =   16
               Top             =   1620
               Width           =   11475
               _cx             =   20241
               _cy             =   7038
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
               Cols            =   19
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmBoleta.frx":4A59
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
            Begin AspaTextBoxFecha.TextBoxFecha txtfechaA 
               Height          =   300
               Index           =   0
               Left            =   1125
               TabIndex        =   18
               Top             =   300
               Width           =   1350
               _ExtentX        =   2381
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
               Valor           =   "14/11/2008"
            End
            Begin AspaTextBoxFecha.TextBoxFecha txtfechaA 
               Height          =   300
               Index           =   1
               Left            =   3495
               TabIndex        =   19
               Top             =   300
               Width           =   1350
               _ExtentX        =   2381
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
               Valor           =   "14/11/2008"
            End
            Begin VB.TextBox txt_cbA 
               Height          =   300
               Index           =   1
               Left            =   5715
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   20
               Text            =   "txt_cbA(1)"
               Top             =   300
               Width           =   765
            End
            Begin VB.TextBox txt_cbA 
               Height          =   300
               Index           =   0
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   21
               Text            =   "txt_cbA(0)"
               Top             =   630
               Width           =   735
            End
            Begin VB.TextBox txt_cbA 
               Height          =   300
               Index           =   2
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   23
               Text            =   "txt_cbA(2)"
               Top             =   945
               Width           =   735
            End
            Begin VB.TextBox txt_cbA 
               Height          =   300
               Index           =   3
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   24
               Text            =   "txt_cbA(3)"
               Top             =   1260
               Width           =   735
            End
            Begin VB.TextBox txt_cbA 
               Height          =   300
               Index           =   4
               Left            =   5715
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   22
               Text            =   "txt_cbA(4)"
               Top             =   630
               Width           =   765
            End
            Begin VB.Label lbl_captA 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Número"
               Height          =   195
               Index           =   4
               Left            =   4980
               TabIndex        =   106
               Top             =   720
               Width           =   555
            End
            Begin VB.Label lbl_codA 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_codA(4)"
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
               Index           =   4
               Left            =   7440
               TabIndex        =   105
               Top             =   630
               Visible         =   0   'False
               Width           =   1110
            End
            Begin VB.Label LblTipCam2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Cambio"
               Height          =   195
               Index           =   0
               Left            =   7065
               TabIndex        =   101
               Top             =   1065
               Width           =   1110
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
               Height          =   285
               Index           =   0
               Left            =   8325
               TabIndex        =   100
               Top             =   960
               Width           =   1080
            End
            Begin VB.Label lblIdCtaDocA 
               AutoSize        =   -1  'True
               Caption         =   "lblIdCtaDocA"
               Height          =   195
               Left            =   4905
               TabIndex        =   98
               Top             =   1335
               Visible         =   0   'False
               Width           =   930
            End
            Begin VB.Label lbl_captA 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Documento"
               Height          =   195
               Index           =   3
               Left            =   45
               TabIndex        =   96
               Top             =   1365
               Width           =   825
            End
            Begin VB.Label lbl_codA 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_codA(3)"
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
               Index           =   3
               Left            =   2850
               TabIndex        =   95
               Top             =   1275
               Visible         =   0   'False
               Width           =   1110
            End
            Begin VB.Label lbl_fechaA 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fch. Pago"
               Height          =   195
               Index           =   1
               Left            =   2685
               TabIndex        =   42
               Top             =   405
               Width           =   735
            End
            Begin VB.Label lbl_fechaA 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fch. Emisión"
               Height          =   195
               Index           =   0
               Left            =   45
               TabIndex        =   41
               Top             =   405
               Width           =   900
            End
            Begin VB.Label lbl_codA 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_codA(0)"
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
               Left            =   2850
               TabIndex        =   40
               Top             =   630
               Visible         =   0   'False
               Width           =   1110
            End
            Begin VB.Label lbl_codA 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_codA(1)"
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
               Index           =   1
               Left            =   7440
               TabIndex        =   38
               Top             =   300
               Visible         =   0   'False
               Width           =   1110
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Procesar en Automático"
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
               Height          =   285
               Left            =   0
               TabIndex        =   37
               Top             =   0
               Width           =   11550
            End
            Begin VB.Label lbl_totalA 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total Ingreso"
               Height          =   210
               Index           =   0
               Left            =   6165
               TabIndex        =   36
               Top             =   5670
               Width           =   930
            End
            Begin VB.Label lbl_totalA 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total Descuento"
               Height          =   210
               Index           =   1
               Left            =   7500
               TabIndex        =   35
               Top             =   5670
               Width           =   1185
            End
            Begin VB.Label lbl_totalA 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total Aporte"
               Height          =   210
               Index           =   2
               Left            =   8835
               TabIndex        =   34
               Top             =   5670
               Width           =   870
            End
            Begin VB.Label lbl_totalA 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total a Pagar"
               Height          =   210
               Index           =   3
               Left            =   10170
               TabIndex        =   33
               Top             =   5670
               Width           =   960
            End
            Begin VB.Label lbl_codA 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_codA(2)"
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
               Index           =   2
               Left            =   2850
               TabIndex        =   32
               Top             =   960
               Visible         =   0   'False
               Width           =   1110
            End
            Begin VB.Label lbl_captA 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Categoría"
               Height          =   195
               Index           =   1
               Left            =   4980
               TabIndex        =   31
               Top             =   420
               Width           =   705
            End
            Begin VB.Label lbl_captA 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Proceso"
               Height          =   195
               Index           =   0
               Left            =   45
               TabIndex        =   30
               Top             =   720
               Width           =   585
            End
            Begin VB.Label lbl_captA 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda"
               Height          =   195
               Index           =   2
               Left            =   45
               TabIndex        =   29
               Top             =   1065
               Width           =   585
            End
            Begin VB.Label lblMonedaA 
               AutoSize        =   -1  'True
               Caption         =   "lblMonedaA"
               Height          =   195
               Left            =   3915
               TabIndex        =   28
               Top             =   1065
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label lbl_cbA 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cbA(2)"
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
               Index           =   2
               Left            =   1875
               TabIndex        =   44
               Top             =   960
               Width           =   1725
            End
            Begin VB.Label lbl_cbA 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cbA(0)"
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
               Left            =   1875
               TabIndex        =   43
               Top             =   630
               Width           =   2970
            End
            Begin VB.Label lbl_cbA 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cbA(1)"
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
               Index           =   1
               Left            =   6465
               TabIndex        =   39
               Top             =   300
               Width           =   2970
            End
            Begin VB.Label lbl_cbA 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cbA(3)"
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
               Index           =   3
               Left            =   1875
               TabIndex        =   97
               Top             =   1275
               Width           =   2970
            End
            Begin VB.Label lbl_cbA 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cbA(4)"
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
               Index           =   4
               Left            =   6465
               TabIndex        =   107
               Top             =   630
               Width           =   2970
            End
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
Attribute VB_Name = "FrmBoleta"
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

Dim RstIngreso As New ADODB.Recordset
Dim RstDescuento As New ADODB.Recordset
Dim RstAportacion As New ADODB.Recordset
Dim QueHaceTmp As Integer '--almacenar el estado de QueHace cuando se elija la opcion Automatico
'

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 2
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

Private Sub fg_EnterCell(Index As Integer)
    If QueHace = 3 Then
        Fg(Index).Editable = flexEDNone
        Exit Sub
    End If
'    If fg(Index).Col < 3 Then
'        fg(Index).Editable = flexEDNone
'    Else
        Fg(Index).Editable = flexEDKbdMouse
'    End If
End Sub
Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row < Fg(Index).FixedRows Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case Index
        Case 0 '--detalle
            Select Case Col
                Case 4, 9, 14
                    If validar_numero(KeyAscii) = False Then KeyAscii = 0
                Case Else
                    KeyAscii = 0
            End Select
        Case 1 '--auto
            Select Case Col
                Case 13, 14
                    If validar_numero(KeyAscii) = False Then KeyAscii = 0
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub fg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 1 Then Exit Sub
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then 'F3 = Agregar Item
        pRegistroAdd
    End If
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then
        pRegistroDel    'F4 = Eliminar Item
    End If
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If QueHace = 3 Then Exit Sub
    If Agregando = True Then Exit Sub
    Select Case Index
        Case 0 '--detalle
            Select Case Col
                Case 4, 9, 14
                    If IsNumeric(Fg(Index).TextMatrix(Row, Col)) = False Then
                        Fg(Index).TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                    Fg(Index).TextMatrix(Row, Col) = Format(Fg(Index).TextMatrix(Row, Col), FORMAT_MONTO)
                Case Else
'                    KeyAscii = 0
            End Select
        Case 1 '--auto
            Select Case Col
                Case 8, 10, 11
                    If IsDate(Fg(Index).TextMatrix(Row, Col)) = False Then
                        Fg(Index).TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                    Fg(Index).TextMatrix(Row, Col) = Format(Fg(Index).TextMatrix(Row, Col), FORMAT_DATE)
                Case 13, 14
                    If NulosN(Fg(Index).TextMatrix(Row, Col)) = 0 Then
                        Fg(Index).TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                    If Col = 13 Then
                        Fg(Index).TextMatrix(Row, Col) = Format(Fg(Index).TextMatrix(Row, Col), "0000")
                    Else
                        Fg(Index).TextMatrix(Row, Col) = Format(Fg(Index).TextMatrix(Row, Col), "0000000000")
                    End If
                    
                Case Else
                    
            End Select
    End Select

    Exit Sub
error:
    SHOW_ERROR Me.Name, "fg_CellChanged (" + CStr(Index) + ")"
End Sub

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String '--almacenará los codigos de conceptos ya seleccionados
    Dim nSQL As String
    Dim nTitulo As String
    Dim xCampos(3, 5) As String
    Dim mIdConceptoCat&
    If QueHace = 3 Then Exit Sub
    Select Case Index
        Case 0
            
            xCampos(0, 0) = "CodSun":       xCampos(0, 1) = "codsun":      xCampos(0, 2) = "900":  xCampos(0, 3) = "C": xCampos(0, 4) = "S"
            xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion": xCampos(1, 2) = "5000": xCampos(1, 3) = "C": xCampos(1, 4) = "N"
            xCampos(2, 0) = "Tipo":         xCampos(2, 1) = "tipnombre":   xCampos(2, 2) = "2200": xCampos(2, 3) = "C": xCampos(2, 4) = "N"

            Select Case Col
                Case 3 '--ingresos
                    nSQLId = GRID_GENERAR_SQL_ID(Fg(0), 1, "pla_concepto.id", "NOT IN", True)
                    If nSQLId <> "" Then nSQLId = " and " & nSQLId
                    mIdConceptoCat = 1
                    nTitulo = "Buscando Ingresos"

                Case 8 '--descuento
                    nSQLId = GRID_GENERAR_SQL_ID(Fg(0), 6, "pla_concepto.id", "NOT IN", True)
                    If nSQLId <> "" Then nSQLId = " and " & nSQLId
                    mIdConceptoCat = 3
                    nTitulo = "Buscando Descuento"
                Case 13 '--aportes
                    nSQLId = GRID_GENERAR_SQL_ID(Fg(0), 11, "pla_concepto.id", "NOT IN", True)
                    If nSQLId <> "" Then nSQLId = " and " & nSQLId
'                    nSQLId = nSQLId & " and pla_conceptotipo."
                    mIdConceptoCat = 2
                    nTitulo = "Buscando Aportes"
            End Select
            nSQL = "SELECT pla_concepto.id,pla_concepto.codsun, pla_concepto.descripcion, pla_conceptotipo.descripcion AS tipnombre " _
                + vbCr + " FROM pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
                + vbCr + " WHERE pla_concepto.variable is not null AND (((pla_conceptotipo.idcat)=" & mIdConceptoCat & ")) " & nSQLId & " and pla_concepto.activo=-1;"
            '---------------------------
            CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            If xRs.State = 0 Then GoTo salir
            If xRs.RecordCount = 0 Then GoTo salir
            
            Agregando = True
            
            Fg(0).TextMatrix(Fg(0).Row, Col) = NulosC(xRs.Fields("descripcion"))
            Fg(0).TextMatrix(Fg(0).Row, Col - 2) = NulosN(xRs.Fields("id"))
            Agregando = False
            '---------------------------
            
        Case 1
            If Col = 10 Or Col = 11 Then
                '--invocar al formulario de horas
                Dim obj As New SGI2_funciones.formularios
                obj.HoraSeleccionar Fg(1), Row, Col, Fg(1).TextMatrix(Row, Col)
                Set obj = Nothing
            End If
    End Select
    Exit Sub
salir:
    Agregando = False
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
            Select Case Index
            Case 0: PopupMenu Menu3
            Case 1: PopupMenu Menu1
            End Select
        End If
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    TabOne1.TabVisible(1) = False
    Dim Rpta As Integer
    SeEjecuto = False
    pConfigurarGrilla
    pCargarGrid
    Blanquea False
    Blanquea True
    
    SeEjecuto = True
'    If RstFrm.RecordCount = 0 Then
'        If MsgBox("No se ha registrado ninguna Boleta, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
'            nuevo
'        End If
'    End If
End Sub

Private Sub pCargarGrid()
    Dim nSQL  As String

    lblperiodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
    lblperiodo(1).Caption = lblperiodo(0).Caption
    lblperiodo(2).Caption = lblperiodo(0).Caption
    
    nSQL = "SELECT pla_boleta.*, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom AS nombres, mae_categoria.nomcor AS categoria, pla_boleta.numser & ' ' & pla_boleta.numdoc AS numerodoc, mae_documento.abrev, mae_moneda.simbolo, " _
        + vbCr + " IIf([pla_boleta].[numreg] Is Null Or [pla_boleta].[numreg]='','',Format([pla_boleta].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([pla_boleta].[numreg],3)) AS registro " _
        + vbCr + " FROM pla_proceso RIGHT JOIN (pla_empleados RIGHT JOIN ((((mae_moneda RIGHT JOIN pla_boleta ON mae_moneda.id = pla_boleta.idmon) LEFT JOIN mae_categoria ON pla_boleta.idcat = mae_categoria.id) LEFT JOIN mae_documento ON pla_boleta.iddoc = mae_documento.id) LEFT JOIN mae_libros ON pla_boleta.idlib = mae_libros.id) ON pla_empleados.id = pla_boleta.idemp) ON pla_proceso.id = pla_boleta.idproc " _
        + vbCr + " WHERE (((pla_boleta.ano)=" & AnoTra & ") AND ((pla_boleta.idmes)=" & xMes & "));"

    '--cargando datos
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon

    Set Dg3.DataSource = RstFrm
    Me.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    QueHace = 3
    QueHaceTmp = QueHace
    
    Dg3.Columns("fchdoc").NumberFormat = FORMAT_DATE
    Dg3.Columns("fchpago").NumberFormat = FORMAT_DATE
    Dg3.Columns("imptot").NumberFormat = FORMAT_MONTO
    Dg3.Columns("impsal").NumberFormat = FORMAT_MONTO
        
    txtfecha(0).Valor = Date
    txtfecha(1).Valor = Date

    CentrarFrm Me
    SeEjecuto = False
    Agregando = False

    Dg3.BatchUpdates = False

    TabOne1.CurrTab = 0
    '--
    Habilitar_Obj False
    '----

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set Dg3.DataSource = Nothing
End Sub


Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then MuestraSegundoTab
    End If
    If OldTab = 1 And NewTab = 2 Then
        Toolbar1.Buttons(7).ToolTipText = "Grabar Automático - Detalle"
        If Fg(1).Row < Fg(1).FixedRows Then Exit Sub
        pAutoPonerDatosEnDetalle Fg(1).Row
    ElseIf OldTab = 2 And NewTab = 1 Then
        Toolbar1.Buttons(7).ToolTipText = "Grabar Automático"
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pAutomatico
    If Button.Index = 3 Then nuevo
    If Button.Index = 4 Then Modificar
    If Button.Index = 5 Then Eliminar
    If Button.Index = 7 Then
        If QueHace = 4 Then '--automatico
            
            If Grabar(True) = True Then
                If TabOne1.CurrTab = 1 Then '--automatico
                    pAutoCargarDatosPersonal
                Else '--detalle del automatico
                    
                End If
                
            End If
            
        Else
            If Grabar(False) = True Then
                Cancelar
                RstFrm.Requery
                Dg3.Refresh
            End If
        End If
    End If
    If Button.Index = 8 Then Cancelar
    If Button.Index = 10 Then Filtrar
    If Button.Index = 11 Then RstFrm.Filter = ""
    If Button.Index = 12 Then Buscar
    If Button.Index = 14 Then CambiarMes
'    If Button.Index = 16 Then pExportarExcel
    If Button.Index = 17 Then pImprimir
    
    If Button.Index = 19 Then
        Set RstFrm = Nothing
        Unload Me
    End If
    
End Sub

Private Sub CambiarMes()
    TabOne1.CurrTab = 0
    xMes = SeleccionaMes(xCon)
    pCargarGrid
End Sub

Sub Eliminar()

    On Error GoTo error
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If

    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim xCod&
    
    xCod = RstFrm("id")



    If MsgBox("¿Esta seguro de eliminar el registro?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
        
            '--Eliminando los conceptos
        xCon.Execute "DELETE * FROM pla_boletadet WHERE idbol = " & xCod & ""
        '--Eliminando los centros de costo asociado a la boleta
        xCon.Execute "DELETE * FROM pla_boletacosto WHERE idbol = " & xCod & ""
        '--eliminado el diario
        xCon.Execute "DELETE * FROM con_diario WHERE idlib = 9 AND idmes = " & xMes & " and idmov = " & xCod & " ;"
        '********************************************************************************
        xCon.Execute "DELETE * FROM pla_boleta WHERE id= " & xCod & ""
        
        
        MsgBox "El registro fue eliminado con éxito", vbInformation, xTitulo
        RstFrm.Requery
        Dg3.Refresh
        TabOne1.CurrTab = 0
        If RstFrm.RecordCount = 0 Then
            If MsgBox("No hay registrado ningúna Boleta de Pago, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
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
    
    TabOne1.TabVisible(0) = True
    TabOne1.TabVisible(1) = False
    
    ActivaTool
    Habilitar_Obj False
    
    Toolbar1.Buttons(7).ToolTipText = "Grabar"
    Toolbar1.Buttons(8).ToolTipText = "Cancelar"
    
    Label1.Caption = "Detalle de la Boleta"
    TabOne1.CurrTab = 0
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
    ActivaTool
    Habilitar_Obj True
    
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 2
    End If

    Fg(1).Editable = flexEDKbdMouse
    Fg(1).SelectionMode = flexSelectionFree

    Label1.Caption = "Modificando Horario"

    Fg(1).ColFormat(3) = FORMAT_HORA_LARGO
    Fg(1).ColFormat(4) = FORMAT_HORA_LARGO

    txt(1).SetFocus

End Sub

Sub MuestraSegundoTab()
'    On Error GoTo error
    Dim QueHaceTmp1  As Integer
    With RstFrm
        Blanquea
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Then Exit Sub
        QueHaceTmp1 = QueHace
        QueHace = -1
        If IsDate(RstFrm("fchdoc")) = True Then
            txtfecha(0).Valor = CDate(RstFrm("fchdoc"))
        Else
            txtfecha(0).Valor = ""
        End If
        If IsDate(RstFrm("fchpago")) = True Then
            txtfecha(1).Valor = CDate(RstFrm("fchpago"))
        Else
            txtfecha(1).Valor = ""
        End If
        
        '--proceso
        If NulosN(RstFrm("idproc")) <> 0 Then
            txt_cb(0).Text = NulosN(RstFrm("idproc"))
            txt_cb_Validate 0, False
        End If
        '--personal
        If NulosN(RstFrm("idemp")) <> 0 Then
            txt_cb(1).Text = NulosN(RstFrm("idemp"))
            txt_cb_Validate 1, False
        End If
        
        txt(1).Text = NulosC(RstFrm("numser"))
        txt(2).Text = NulosC(RstFrm("numdoc"))
        
        '--moneda
        If NulosN(RstFrm("idmon")) <> 0 Then
            txt_cb(2).Text = NulosN(RstFrm("idmon"))
            txt_cb_Validate 2, False
            If NulosN(RstFrm("iddoc")) <> 0 Then
                txt_cb(3).Text = NulosN(RstFrm("iddoc"))
                txt_cb_Validate 3, False
            End If
        End If
        QueHace = QueHaceTmp1
        
'        txt(0).Text = NulosN(.Fields("id")) '--CODIGO
'        txt(1).Text = NulosC(.Fields("descripcion"))
'        If IsDate(.Fields("tolerancia")) = True Then
'            dtpk(0).Value = CDate(.Fields("tolerancia"))
'        End If
        '---

        '-----------------------------
        Set RstIngreso = Nothing
        Set RstDescuento = Nothing
        Set RstAportacion = Nothing
        '-----------------------------
    
        pConceptoDocumentoEmp RstIngreso, NulosN(.Fields("id")), e_Remuneracion
        pConceptoDocumentoEmp RstDescuento, NulosN(.Fields("id")), e_Descuento
        pConceptoDocumentoEmp RstAportacion, NulosN(.Fields("id")), e_Aportacion
        
        pCargarConceptosDetalle NulosN(.Fields("idemp"))
        

    End With

    Exit Sub
error:

    SHOW_ERROR
End Sub


Private Sub Habilitar_Obj(band As Boolean, Optional EsAuto As Boolean = False)
    habilitar_Locked txt, Not band
    habilitar_Locked txt_cb, Not band
    habilitar_Locked txtfecha, Not band
    
    habilitar_Locked txtfechaA, Not band
    habilitar_Locked txt_cbA, Not band
        
    habilitar cmdManual, band

    TabOne1.CurrTab = IIf(band = False, 0, 2)
    TabOne1.TabEnabled(2) = Not band
    
    TabOne1.TabEnabled(0) = Not band

    If band = False Then
        Fg(0).SelectionMode = flexSelectionByRow
        Fg(1).SelectionMode = flexSelectionByRow
    End If
End Sub

Private Sub Blanquea(Optional EsAuto As Boolean = False)
    
    If EsAuto = False Then
        LimpiaText txt
        LimpiaText txt_cb
        LimpiaText txtfecha
        LimpiaText txt_total
        Fg(0).Rows = Fg(0).FixedRows
        LblTipoCambio(1).Caption = ""
    Else
        LimpiaText txt_cbA
        habilitar_Locked txt_cbA, Not txt_cbA(0).Locked
        LimpiaText txtfechaA
        LimpiaText txt_totalA
        Fg(1).Rows = Fg(1).FixedRows
        LblTipoCambio(0).Caption = ""
    End If

End Sub

Private Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
    Toolbar1.Buttons(16).Enabled = True
    Toolbar1.Buttons(17).Enabled = True
End Sub

Private Sub nuevo()
    QueHace = 1
    ActivaTool
    Blanquea
    Habilitar_Obj True, False
    Label1.Caption = "Agregando Boleta"
    '------------

    Fg(0).Editable = flexEDKbdMouse
    Fg(0).SelectionMode = flexSelectionFree
    '------------
    TabOne1.CurrTab = 2
    txtfecha(0).SetFocus

End Sub


Private Function Grabar(Optional fEsAutomatico As Boolean = False) As Boolean
    If fValidarDatos(fEsAutomatico) = False Then Exit Function
    If fEsAutomatico = False Then
        If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el Registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo salir
    Else
        If TabOne1.CurrTab = 1 Then
            If MsgBox("Seguro desea grabar el Proceso Automático", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo salir
        Else
            If MsgBox("Seguro desea grabar el Proceso Automático - Detalle" & vbCr & "Personal: " & lbl_cb(1).Caption, vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo salir
        End If
    End If
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDiario As New ADODB.Recordset
    
    Dim RstTmp As New ADODB.Recordset
    Dim xCod&, xCol&, xFil&, mIdEmp&, mRow&
    Dim nSQL As String

    On Error GoTo LaCague
    Me.MousePointer = vbHourglass
    
    xCon.BeginTrans
    
    RST_Busq RstCab, "SELECT top 1 * FROM pla_boleta ", xCon
    RST_Busq RstDet, "SELECT top 1 * FROM pla_boletadet", xCon
    RST_Busq RstDiario, "SELECT top 1 * FROM con_diario", xCon

    If fEsAutomatico = False Then
        GrabarDet IIf(QueHace = 1, -1, RstFrm("id")), RstCab, RstDet, RstDiario, QueHace
    Else
        Dim QueHaceTmp As Integer
        pBloquearEnAuto False
        If TabOne1.CurrTab = 1 Then
            For mRow = Fg(1).FixedRows To Fg(1).Rows - 1
                DoEvents
                Fg(1).Row = mRow
                If NulosN(Fg(1).TextMatrix(mRow, 4)) = -1 Then
                    '--colocar los datos en el tab del detalle
                    pAutoPonerDatosEnDetalle mRow
                    If NulosN(Fg(1).TextMatrix(mRow, 1)) = 0 Then
                        QueHaceTmp = 1
                    Else
                        QueHaceTmp = 2
                        xCod = NulosN(Fg(1).TextMatrix(mRow, 1))
                    End If
                    
                    If NulosN(txt_total(0).Text) <> 0 Or NulosN(txt_total(1).Text) <> 0 Or NulosN(txt_total(2).Text) <> 0 Or NulosN(txt_total(3).Text) <> 0 Then
                        If GrabarDet(xCod, RstCab, RstDet, RstDiario, QueHaceTmp) = False Then GoTo LaCague
                    End If
                    
                End If
            Next
        Else '--grabar el detalle
            mRow = Fg(1).Row
            If NulosN(Fg(1).TextMatrix(mRow, 1)) = 0 Then
                QueHaceTmp = 1
            Else
                QueHaceTmp = 2
                xCod = NulosN(Fg(1).TextMatrix(mRow, 1))
            End If
            
            If NulosN(txt_total(0).Text) <> 0 Or NulosN(txt_total(1).Text) <> 0 Or NulosN(txt_total(2).Text) <> 0 Or NulosN(txt_total(3).Text) <> 0 Then
                If GrabarDet(xCod, RstCab, RstDet, RstDiario, QueHaceTmp) = False Then GoTo LaCague
            End If
            '******************************************************************************************
            '--actualizar los valores a la grilla
            Agregando = True
            Fg(1).TextMatrix(mRow, 1) = xCod
            Fg(1).TextMatrix(mRow, 4) = 0       '--check
            Fg(1).TextMatrix(mRow, 10) = Format(txtfecha(0).Valor, FORMAT_DATE) '--fecha doc
            Fg(1).TextMatrix(mRow, 11) = Format(txtfecha(1).Valor, FORMAT_DATE) '--fecha pago
            Fg(1).TextMatrix(mRow, 3) = NulosN(txt_cb(2).Text)      '--moneda codigo
            Fg(1).TextMatrix(mRow, 12) = NulosC(lblMoneda.Caption)  '--moneda simbolo
            
            Fg(1).TextMatrix(mRow, 13) = txt(1).Text '--serie
            Fg(1).TextMatrix(mRow, 14) = txt(2).Text '--numero doc
            
            Fg(1).TextMatrix(mRow, 15) = Format(NulosN(txt_total(0).Text), FORMAT_MONTO) '--ingreso
            Fg(1).TextMatrix(mRow, 16) = Format(NulosN(txt_total(1).Text), FORMAT_MONTO) '--descuento
            Fg(1).TextMatrix(mRow, 17) = Format(NulosN(txt_total(2).Text), FORMAT_MONTO) '--aportes
            Fg(1).TextMatrix(mRow, 18) = Format(NulosN(txt_total(3).Text), FORMAT_MONTO) '--total pago
            Agregando = False
            '******************************************************************************************
        End If
        
    End If
    '--------------------------------------------------------------------------------------------------
    xCon.CommitTrans
    If fEsAutomatico = False Then
        MsgBox "El registro se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    Else
        If TabOne1.CurrTab = 1 Then
            MsgBox "El Proceso se Actualizó correctamente", vbInformation, xTitulo
        Else
            MsgBox "El Registro del Proceso se Actualizó correctamente" & vbCr & "Personal: " & lbl_cb(1).Caption, vbInformation, xTitulo
        End If
    End If
    pBloquearEnAuto True
    Grabar = True
salir:
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstTmp = Nothing
    Me.MousePointer = vbDefault
    Exit Function
LaCague:
    pBloquearEnAuto True
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:   Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function



Private Function GrabarDet(xCod&, RstCab As ADODB.Recordset, _
                           RstDet As ADODB.Recordset, RstDiario As ADODB.Recordset, _
                           QueHaceTmp As Integer) As Boolean
   
    Dim RstTmp As New ADODB.Recordset
    Dim xCol&, xFil&, mIdEmp&
    Dim nSQL As String
    Dim xNumAsiento As String
    
    On Error GoTo LaCague
    
    mIdEmp = NulosN(lbl_cod(1).Caption)
    
    If QueHaceTmp = 1 Then
            xNumAsiento = NuevoNumAsiento(9, xMes, xCon)
            xCod = HallaCodigoTabla("pla_boleta", xCon, "id")
            RstCab.AddNew
            RstCab("id") = xCod
        Else
            
            RST_Busq RstCab, "SELECT * FROM pla_boleta WHERE id =" & xCod & "", xCon
            '--Eliminando los conceptos del detalle
            xCon.Execute "DELETE * FROM pla_boletadet WHERE idbol = " & xCod & ""
            '--Eliminando los centros de costo asociado a la boleta
            xCon.Execute "DELETE * FROM pla_boletacosto WHERE idbol = " & xCod & ""
            
            '********************************************************************************
            xNumAsiento = DevuelveNumAsiento(9, NulosN(xCod), xMes, xCon)
            If xNumAsiento = "" Then xNumAsiento = NuevoNumAsiento(9, xMes, xCon)
            '--eliminado el diario
            xCon.Execute "DELETE * FROM con_diario WHERE idlib = 9 AND idmes = " & xMes & " and idmov = " & xCod & " ;"
            '********************************************************************************
        End If

    '--------------------------------------------------------------------------------------------------
    RstCab("idlib") = 9
    RstCab("ano") = AnoTra
    RstCab("idmes") = xMes
    RstCab("idemp") = mIdEmp
    RstCab("idproc") = NulosN(lbl_cod(0).Caption)
    RstCab("idmon") = NulosN(lbl_cod(2).Caption)
    RstCab("iddoc") = NulosN(lbl_cod(3).Caption)
    RstCab("numser") = NulosC(txt(1).Text)
    RstCab("numdoc") = NulosC(txt(2).Text)
    RstCab("fchdoc") = CDate(txtfecha(0).Valor)
    RstCab("fchpago") = CDate(txtfecha(1).Valor)
    RstCab("impingr") = NulosN(txt_total(0).Text)
    RstCab("impapor") = NulosN(txt_total(1).Text)
    RstCab("impdesc") = NulosN(txt_total(2).Text)
    RstCab("imptot") = NulosN(txt_total(3).Text)
    '--obtener el ultimo saldo del documento
    RstCab("impsal") = NulosN(txt_total(3).Text)
    '-----
    RstCab("totseg") = 0
    
    RstCab("numreg") = Format(xMes, "00") + xNumAsiento
    If xMes <> 0 And xMes <> 13 Then
        RstCab("fchreg") = CDate("01/" + Format(xMes, "00") + "/" + AnoTra)
    End If
    
    '**********************************************************************************************************
    '--grabando la categoria,fecha de ingreso(de trabajo a la empresa)
    nSQL = "SELECT pla_empleados.*, pla_periodolaboral.idcat, mae_categoria.descripcion AS categoria, pla_periodolaboral.fchini AS fchingreso, pla_periodolaboral.fchfin AS fchsalida " _
        + vbCr + " FROM mae_categoria RIGHT JOIN (pla_empleados LEFT JOIN pla_periodolaboral ON pla_empleados.id = pla_periodolaboral.idemp) ON mae_categoria.id = pla_periodolaboral.idcat " _
        + vbCr + " WHERE (((pla_empleados.id)=" & mIdEmp & ") AND ((pla_periodolaboral.fchfin) Is Null));"
   
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount <> 0 Then
        RstCab("idcat") = NulosN(RstTmp("idcat"))
        
        RstCab("fchingreso") = CDate(RstTmp("fchingreso"))
        
        '--------------------------------------------------------------------------------------------------
        '--grabando el regimen pensionario del trabajador
        nSQL = "SELECT DISTINCT  pla_categoria1.idemp, 1 AS idcat, pla_categoria1.idregpen " _
            + vbCr + " FROM pla_categoria1 LEFT JOIN pla_conceptoregpen ON pla_categoria1.idregpen = pla_conceptoregpen.idregpen " _
            + vbCr + " WHERE  (((pla_categoria1.idemp)=" & mIdEmp & ") AND ((1)=" & NulosN(RstTmp("idcat")) & ")); " _
            + vbCr + " UNION " _
            + vbCr + " SELECT DISTINCT pla_categoria2.idemp, 1 AS idcat, pla_categoria2.idregpen " _
            + vbCr + " FROM pla_categoria2 LEFT JOIN pla_conceptoregpen ON pla_categoria2.idregpen = pla_conceptoregpen.idregpen " _
            + vbCr + " WHERE  (((pla_categoria2.idemp)=" & mIdEmp & ") AND ((2)=" & NulosN(RstTmp("idcat")) & "));"
        
        Set RstTmp = Nothing
        RST_Busq RstTmp, nSQL, xCon
        
        If RstTmp.RecordCount <> 0 Then RstCab("idregpen") = NulosN(RstTmp("idregpen"))
    End If
    Set RstTmp = Nothing
    '**********************************************************************************************************
    
    RstCab.Update

    '**********************************************************************************************************
    RstIngreso.Filter = "mIdEmp=" & mIdEmp
    RstDescuento.Filter = "mIdEmp=" & mIdEmp
    RstAportacion.Filter = "mIdEmp=" & mIdEmp
    
    If RstIngreso.RecordCount <> 0 Then RstIngreso.MoveFirst
    If RstDescuento.RecordCount <> 0 Then RstDescuento.MoveFirst
    If RstAportacion.RecordCount <> 0 Then RstAportacion.MoveFirst
    '--grabando los conceptos de ingresos
    DoEvents
    Do While Not RstIngreso.EOF
        If NulosN(RstIngreso("imptot")) <> 0 Then
            RstDet.AddNew
            RstDet("idbol") = xCod
            RstDet("idcpto") = NulosN(RstIngreso("idcpto"))
            RstDet("imptot") = NulosC(RstIngreso("imptot"))
            RstDet("aplanilla") = NulosN(RstIngreso("aplanilla"))
            RstDet.Update
        End If
        RstIngreso.MoveNext
    Loop
    '--grabando los conceptos de descuentos
    DoEvents
    Do While Not RstDescuento.EOF
        If NulosN(RstDescuento("imptot")) <> 0 Then
            RstDet.AddNew
            RstDet("idbol") = xCod
            RstDet("idcpto") = NulosN(RstDescuento("idcpto"))
            RstDet("imptot") = NulosC(RstDescuento("imptot"))
            RstDet("aplanilla") = NulosN(RstDescuento("aplanilla"))
            RstDet.Update
        End If
        RstDescuento.MoveNext
    Loop
    '--grabando los conceptos de aportes
    DoEvents
    Do While Not RstAportacion.EOF
        If NulosN(RstAportacion("imptot")) <> 0 Then
            RstDet.AddNew
            RstDet("idbol") = xCod
            RstDet("idcpto") = NulosN(RstAportacion("idcpto"))
            RstDet("imptot") = NulosC(RstAportacion("imptot"))
            RstDet("aplanilla") = NulosN(RstAportacion("aplanilla"))
            RstDet.Update
        End If
        RstAportacion.MoveNext
    Loop
    '**********************************************************************************************************
    '--insertando los centros de costo
    xCon.Execute "INSERT INTO pla_boletacosto (idbol,idcencos,impcos) " _
            + vbCr + "SELECT pla_boleta.id,pla_empleadoscos.idcencos, Sum(([pla_boleta].[imptot]*[pla_empleadoscos].[imppor])/100) AS costo " _
            + vbCr + "FROM pla_boleta INNER JOIN pla_empleadoscos ON pla_boleta.idemp = pla_empleadoscos.idemp " _
            + vbCr + "WHERE (((pla_empleadoscos.imppor) <> 0)) " _
            + vbCr + "GROUP BY pla_empleadoscos.idcencos,pla_boleta.id " _
            + vbCr + "HAVING (((pla_boleta.id)=" & xCod & ")); "
    
    '**********************************************************************************************************
    '**************** creando el asiento del detalle
    '**********************************************************************************************************
    '--se procedera a hacer la consulta para obtener las cuentas relacionados a los conceptos de la planilla
    '--
    
    '--cargar el documento
    nSQL = "SELECT con_planctas.id AS idcta, con_planctas.cuenta, con_planctas.descripcion AS nomcta, 'H' AS Origen,'Doc' as Tipo, pla_boleta.imptot AS total " _
        + vbCr + " FROM pla_boleta INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (pla_boleta.idmon = mae_documentocta.idmon) AND (pla_boleta.iddoc = mae_documentocta.iddoc) " _
        + vbCr + " WHERE (((pla_boleta.id)=" & xCod & ")); "
        
    '--cargar las cuentas que estan relacionadas a los conceptos
    nSQL = nSQL _
        + vbCr + " UNION " _
        + vbCr + " SELECT con_planctas.id AS idcta, con_planctas.cuenta, con_planctas.descripcion AS nomcta, 'D' AS Origen,'Detalle' as Tipo, Sum(pla_boletadet.imptot) AS total " _
        + vbCr + " FROM (con_planctas RIGHT JOIN pla_concepto ON con_planctas.id = pla_concepto.idctadeb) INNER JOIN pla_boletadet ON pla_concepto.id = pla_boletadet.idcpto " _
        + vbCr + " WHERE (((pla_boletadet.idbol)=" & xCod & ") AND ((pla_concepto.idctadeb) Is Not Null And (pla_concepto.idctadeb)<>0) AND ((pla_concepto.aplanilla)=-1)) " _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion; " _
        + vbCr + " UNION " _
        + vbCr + " SELECT con_planctas.id AS idcta, con_planctas.cuenta, con_planctas.descripcion AS nomcta, 'H' AS Origen,'Detalle' as Tipo, Sum(pla_boletadet.imptot) AS total " _
        + vbCr + " FROM (pla_concepto LEFT JOIN con_planctas ON pla_concepto.idctahab = con_planctas.id) INNER JOIN pla_boletadet ON pla_concepto.id = pla_boletadet.idcpto " _
        + vbCr + " WHERE (((pla_boletadet.idbol)=" & xCod & ") AND ((pla_concepto.idctahab) Is Not Null And (pla_concepto.idctahab)<>0) AND ((pla_concepto.aplanilla)=-1)) " _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion; "
        
    '--cargar las cuentas automaticas(debe y haber) relaciondas al debe del concepto
    nSQL = nSQL _
        + vbCr + " UNION " _
        + vbCr + " SELECT con_planctas_1.id AS idcta, con_planctas_1.cuenta, con_planctas_1.descripcion AS nomcta, 'D' AS Origen,'Automatico1' as tipo, Sum(pla_boletadet.imptot) AS total " _
        + vbCr + " FROM con_planctas AS con_planctas_1 RIGHT JOIN ((con_planctas RIGHT JOIN pla_concepto ON con_planctas.id = pla_concepto.idctadeb) INNER JOIN pla_boletadet ON pla_concepto.id = pla_boletadet.idcpto) ON con_planctas_1.id = con_planctas.ctadesdeb " _
        + vbCr + " WHERE (((pla_boletadet.idbol)=" & xCod & ") AND ((pla_concepto.idctadeb) Is Not Null And (pla_concepto.idctadeb)<>0) AND ((pla_concepto.aplanilla)=-1) AND ((con_planctas.ctadesdeb) Is Not Null And (con_planctas.ctadesdeb)<>0)) " _
        + vbCr + " GROUP BY con_planctas_1.id, con_planctas_1.cuenta, con_planctas_1.descripcion; " _
        + vbCr + " UNION " _
        + vbCr + " SELECT con_planctas_1.id AS idcta, con_planctas_1.cuenta, con_planctas_1.descripcion AS nomcta, 'H' AS Origen, 'Automatico1' AS tipo, Sum(pla_boletadet.imptot) AS total " _
        + vbCr + " FROM ((con_planctas RIGHT JOIN pla_concepto ON con_planctas.id = pla_concepto.idctadeb) INNER JOIN pla_boletadet ON pla_concepto.id = pla_boletadet.idcpto) LEFT JOIN con_planctas AS con_planctas_1 ON con_planctas.ctadeshab = con_planctas_1.id " _
        + vbCr + " WHERE (((pla_boletadet.idbol)=" & xCod & ") AND ((pla_concepto.idctadeb) Is Not Null And (pla_concepto.idctadeb)<>0) AND ((pla_concepto.aplanilla)=-1) AND ((con_planctas.ctadeshab) Is Not Null And (con_planctas.ctadeshab)<>0)) " _
        + vbCr + " GROUP BY con_planctas_1.id, con_planctas_1.cuenta, con_planctas_1.descripcion; "
    
    '--cargar las cuentas automaticas(debe y haber) relaciondas al haber del concepto
    nSQL = nSQL _
        + vbCr + " UNION " _
        + vbCr + " SELECT con_planctas_1.id AS idcta, con_planctas_1.cuenta, con_planctas_1.descripcion AS nomcta, 'D' AS Origen, 'Automatico2' AS tipo, Sum(pla_boletadet.imptot) AS total " _
        + vbCr + " FROM ((pla_concepto LEFT JOIN con_planctas ON pla_concepto.idctahab = con_planctas.id) INNER JOIN pla_boletadet ON pla_concepto.id = pla_boletadet.idcpto) LEFT JOIN con_planctas AS con_planctas_1 ON con_planctas.ctadesdeb = con_planctas_1.id " _
        + vbCr + " WHERE (((pla_boletadet.idbol)=" & xCod & ") AND ((pla_concepto.idctahab) Is Not Null And (pla_concepto.idctahab)<>0) AND ((pla_concepto.aplanilla)=-1) AND ((con_planctas.ctadesdeb) Is Not Null And (con_planctas.ctadesdeb)<>0)) " _
        + vbCr + " GROUP BY con_planctas_1.id, con_planctas_1.cuenta, con_planctas_1.descripcion; " _
        + vbCr + " UNION " _
        + vbCr + " SELECT con_planctas_1.id AS idcta, con_planctas_1.cuenta, con_planctas_1.descripcion AS nomcta, 'H' AS Origen, 'Automatico2' AS tipo, Sum(pla_boletadet.imptot) AS total " _
        + vbCr + " FROM ((pla_concepto LEFT JOIN con_planctas ON pla_concepto.idctahab = con_planctas.id) LEFT JOIN con_planctas AS con_planctas_1 ON con_planctas.ctadeshab = con_planctas_1.id) INNER JOIN pla_boletadet ON pla_concepto.id = pla_boletadet.idcpto " _
        + vbCr + " WHERE (((pla_boletadet.idbol)=" & xCod & ") AND ((pla_concepto.idctahab) Is Not Null And (pla_concepto.idctahab)<>0) AND ((pla_concepto.aplanilla)=-1) AND ((con_planctas.ctadeshab) Is Not Null And (con_planctas.ctadeshab)<>0)) " _
        + vbCr + " GROUP BY con_planctas_1.id, con_planctas_1.cuenta, con_planctas_1.descripcion;"
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        DoEvents
        pGenerarAsiento RstDiario, AnoTra, xMes, 9, xCod, 0, 0, xNumAsiento, NulosN(LblTipoCambio(1).Caption), CDate(txtfecha(0).Valor), NulosN(RstTmp.Fields("idcta")), NulosN(lbl_cod(2).Caption), NulosN(RstTmp.Fields("total")), IIf(RstTmp.Fields("origen") = "D", True, False)
        RstTmp.MoveNext
    Loop
    Set RstTmp = Nothing
    '**********************************************************************************************************
    '**************** fin creando el asiento del detalle
    '**********************************************************************************************************
    
    GrabarDet = True
salir:
     Exit Function
LaCague:
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    GrabarDet = False
End Function


Private Function fValidarDatos(Optional fEsAutomatico As Boolean = False) As Boolean
'    Dim mRow&, QGrid&
'    Dim band As Integer
'
'    band = Validar(txt_cb)
'    If band <> -1 Then
'       MsgBox "Llene el Campo de " & lbl_cb_capt(band).Caption, vbInformation, xTitulo
'       txt_cb(band).SetFocus
'       Exit Function
'    End If
'
'    band = Validar(txt)
'    If band > 0 Then
'       MsgBox "Llene el Campo de " & lbl(band).Caption, vbInformation, xTitulo
'       txt(band).SetFocus
'       Exit Function
'    End If
'
'    '--------------------------------
'    '--VALIDAR QUE NO ESTE REGISTRADO
'    Dim RstTmp As New ADODB.Recordset
'    If QueHace = 1 Then
'        RST_Busq RstTmp, "SELECT descripcion FROM pla_boleta WHERE ucase(descripcion)='" + UCase(Trim(txt(1).Text)) + "';", xCon
'    Else
'        RST_Busq RstTmp, "SELECT descripcion FROM pla_boleta WHERE ucase(descripcion)='" + UCase(Trim(txt(1).Text)) + "' AND id <> " + CStr(RstFrm.Fields("id")) + ";", xCon
'    End If
'    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then
'        MsgBox "El registro " + IIf(QueHace = 1, " ya fue ingresado", "ya existe"), vbExclamation, xTitulo
'        Set RstTmp = Nothing
'        Exit Function
'    End If
'    Set RstTmp = Nothing
'    '--------------------------------
''    If fEsAutomatico = False Then
'        If fg(0).Rows = fg(0).FixedRows Then
'            MsgBox "No hay conceptos Asignados", vbExclamation, xTitulo
'            cmdManual(0).SetFocus
'            Exit Function
'        End If
''    Else
''        If fg(1).Rows = 1 Then
''            MsgBox "No hay Personal", vbExclamation, xTitulo
''            cmd(3).SetFocus
''            Exit Function
''        End If
''    End If
'    '--------------------------------
'    With fg(0)
'        For mRow = .FixedRows To .Rows - 1
'            If IsDate(.TextMatrix(mRow, 3)) = False Then
'                MsgBox "Ingrese la Hora de Inicio" + vbCr + "Tipo de Hora: " + .TextMatrix(mRow, 2), vbExclamation, xTitulo
'                Agregando = True:  .Row = mRow:     .Col = 3: Agregando = False
'                fg(1).SetFocus
'                Exit Function
'            ElseIf IsDate(.TextMatrix(mRow, 4)) = False Then
'                MsgBox "Ingrese la Hora Final" + vbCr + "Tipo de Hora: " + .TextMatrix(mRow, 2), vbExclamation, xTitulo
'                Agregando = True:  .Row = mRow:     .Col = 4: Agregando = False
'                fg(1).SetFocus
'                Exit Function
'            End If
'        Next mRow
'    End With
'    '--------------------------------
    fValidarDatos = True
End Function

Sub Buscar()
'    On Error GoTo error
'    TabOne1.CurrTab = 0
'
'    Dim xRs As New ADODB.Recordset
'    Dim nSQL As String
'
'    Dim xCampos(2, 4) As String
'
'    xCampos(0, 0) = "Descripción":        xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":     xCampos(0, 3) = "C"
'    xCampos(1, 0) = "Tolerancia":         xCampos(1, 1) = "tolerancia":       xCampos(1, 2) = "1500":     xCampos(1, 3) = "F"
'
'
'    nSQL = "SELECT mae_horario.*, IIf([mae_horario].[vigencia]=-1,'Vigente','De Baja') AS estado " _
'        + vbCr + " FROM mae_horario " _
'        + vbCr + " ORDER BY mae_horario.descripcion;"
'
'    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Horario", "descripcion", "descripcion", Principio
'    If xRs.State = 0 Then GoTo salir
'    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir
'
'    RstFrm.MoveFirst
'    RstFrm.Find "id = " + CStr(xRs("id"))
'salir:
'    Set xRs = Nothing
'    Exit Sub
'error:
'    Set xRs = Nothing
'    SHOW_ERROR Me.Name, "Buscar"
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
    Agregando = True
    With Fg(0) '--del detalle de la planilla

        .Rows = 2
        .Cols = 15
        .FixedRows = 2
        .RowHeight(0) = 300
        .RowHeight(1) = 250

        GRID_COMBINAR Fg(0), 0, 1, 0, 4, "Remuneraciones", flexAlignLeftCenter, True, , &H800000, &HD8E9EC, True
        GRID_COMBINAR Fg(0), 0, 5, 1, 5, " ", flexAlignLeftCenter, False, , , &HDDDDFF
        GRID_COMBINAR Fg(0), 0, 6, 0, 9, "Descuentos y Aportes del Trabajador", flexAlignLeftCenter, True, , &H800000, &HD8E9EC, True
        GRID_COMBINAR Fg(0), 0, 10, 1, 10, " ", flexAlignLeftCenter, False, , , &HDDDDFF
        GRID_COMBINAR Fg(0), 0, 11, 0, 14, "Aportes del Empleador", flexAlignLeftCenter, True, , &H800000, &HD8E9EC, True

        .TextMatrix(1, 1) = "IdCpto":      .ColWidth(1) = 0:
        .TextMatrix(1, 2) = "Aplanilla":   .ColWidth(2) = 0:
        .TextMatrix(1, 3) = "Descripión":  .ColWidth(3) = 2750:   .ColAlignment(3) = flexAlignLeftCenter:   .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 4) = "Importe":     .ColWidth(4) = 800:    .ColAlignment(4) = flexAlignRightCenter:  .Row = 1: .Col = 4: .CellAlignment = flexAlignRightCenter
        .ColWidth(5) = 100:
        .TextMatrix(1, 6) = "IdCpto":      .ColWidth(6) = 0:
        .TextMatrix(1, 7) = "Aplanilla":   .ColWidth(7) = 0:
        .TextMatrix(1, 8) = "Descripión":  .ColWidth(8) = 2750:   .ColAlignment(8) = flexAlignLeftCenter:   .Row = 1: .Col = 8: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 9) = "Importe":     .ColWidth(9) = 800:    .ColAlignment(9) = flexAlignRightCenter:  .Row = 1: .Col = 9: .CellAlignment = flexAlignRightCenter
        .ColWidth(10) = 100:
        .TextMatrix(1, 11) = "IdCpto":     .ColWidth(11) = 0:
        .TextMatrix(1, 12) = "Aplanilla":  .ColWidth(12) = 0:
        .TextMatrix(1, 13) = "Descripión": .ColWidth(13) = 2750:  .ColAlignment(13) = flexAlignLeftCenter:  .Row = 1: .Col = 13: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 14) = "Importe":    .ColWidth(14) = 800:   .ColAlignment(14) = flexAlignRightCenter: .Row = 1: .Col = 14: .CellAlignment = flexAlignRightCenter

        .SelectionMode = flexSelectionByRow
        
        GRID_COMBOLIST Fg(0), 3
        GRID_COMBOLIST Fg(0), 8
'        GRID_COMBOLIST Fg(0), 13
        
    End With

    With Fg(1) '--proceso autimatico
        .Clear
        .Rows = 2
        .FixedRows = 2
        .Cols = 19
        .RowHeight(0) = 250
        .RowHeight(1) = 500
        .FrozenCols = 5

        GRID_COMBINAR Fg(1), 0, 1, 0, 3, "id's", flexAlignCenterCenter, True, flexMergeFree, vbBlack, &HD8E9EC, True
        GRID_COMBINAR Fg(1), 0, 4, 0, 11, " ", flexAlignCenterCenter, True, flexMergeFree, vbBlack, &HD8E9EC, True
        GRID_COMBINAR Fg(1), 0, 13, 0, 14, "Documento", flexAlignCenterCenter, True, flexMergeFree, vbBlack, &HD8E9EC, True
        GRID_COMBINAR Fg(1), 0, 15, 0, 17, "Totales", flexAlignCenterCenter, True, flexMergeFree, vbBlack, &HD8E9EC, True
        GRID_COMBINAR Fg(1), 0, 18, 0, 18, " ", flexAlignCenterCenter, True, flexMergeFree, vbBlack, &HD8E9EC, True

        .TextMatrix(1, 1) = "IdBol":        .ColWidth(1) = 0:
        .TextMatrix(1, 2) = "IdEmp":        .ColWidth(2) = 0:
        .TextMatrix(1, 3) = "IdMon":        .ColWidth(3) = 0:

        .TextMatrix(1, 4) = "Sel":          .ColWidth(4) = 350:    .ColAlignment(4) = flexAlignCenterCenter: .Row = 1: .Col = 4: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 5) = "Personal":     .ColWidth(5) = 2400:   .ColAlignment(5) = flexAlignLeftCenter:   .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 6) = "Cat.":         .ColWidth(6) = 450:    .ColAlignment(6) = flexAlignLeftCenter:   .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 7) = "Cargo":        .ColWidth(7) = 790:    .ColAlignment(7) = flexAlignLeftCenter:   .Row = 1: .Col = 7: .CellAlignment = flexAlignLeftCenter

        .TextMatrix(1, 8) = "Fecha" & vbCr & "Ingreso":   .ColWidth(8) = 800:    .ColAlignment(8) = flexAlignCenterCenter: .Row = 1: .Col = 8: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 9) = "Horas" & vbCr & "Trabajo":   .ColWidth(9) = 780:    .ColAlignment(9) = flexAlignRightCenter:  .Row = 1: .Col = 9: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 10) = "Fecha" & vbCr & "Emisión":  .ColWidth(10) = 800:   .ColAlignment(10) = flexAlignCenterCenter: .Row = 1: .Col = 10: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 11) = "Fecha" & vbCr & "Pago":     .ColWidth(11) = 800:   .ColAlignment(11) = flexAlignCenterCenter: .Row = 1: .Col = 11: .CellAlignment = flexAlignCenterCenter

        .TextMatrix(1, 12) = "M":          .ColWidth(12) = 450:  .ColAlignment(12) = flexAlignCenterCenter: .Row = 1: .Col = 12: .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(1, 13) = "Serie":      .ColWidth(13) = 450:  .ColAlignment(13) = flexAlignCenterCenter: .Row = 1: .Col = 13: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 14) = "Número":     .ColWidth(14) = 1000:  .ColAlignment(14) = flexAlignCenterCenter: .Row = 1: .Col = 14: .CellAlignment = flexAlignCenterCenter

        .TextMatrix(1, 15) = "Ingresos":    .ColWidth(15) = 700:  .ColAlignment(15) = flexAlignRightCenter:  .Row = 1: .Col = 15: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 16) = "Descuento":   .ColWidth(16) = 850:  .ColAlignment(16) = flexAlignRightCenter:  .Row = 1: .Col = 16: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 17) = "Aportes":     .ColWidth(17) = 700:  .ColAlignment(17) = flexAlignRightCenter:  .Row = 1: .Col = 17: .CellAlignment = flexAlignRightCenter

        .TextMatrix(1, 18) = "Neto a" & vbCr & "Pagar":  .ColWidth(18) = 850:  .ColAlignment(18) = flexAlignRightCenter:  .Row = 1: .Col = 18: .CellAlignment = flexAlignRightCenter


        .ColDataType(4) = flexDTBoolean
        .SelectionMode = flexSelectionByRow

    End With
    GRID_COMBOLIST Fg(1), 12
    Agregando = False
    DoEvents
End Sub

'****************************************************************************************
Private Sub cb_Click(Index As Integer)
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error

    If QueHace = 3 Then Exit Sub

    Select Case Index
        Case 0 '--proceso =>> manual
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"

            nTitulo = "Proceso"
            nSQL = "SELECT pla_proceso.id, pla_proceso.descripcion AS nombre, pla_proceso.id AS cod " _
                + vbCr + " FROM pla_proceso " _
                + vbCr + " WHERE (((pla_proceso.enproceso)=-1)); "

        Case 1 '--personal
            pCargarPersonal Me, Index
            txt(1).SetFocus
            Exit Sub
            
        Case 2 '--moneda =>>>  ::manual
            ReDim xCampos(3, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Símbolo":  xCampos(1, 1) = "simbolo":   xCampos(1, 2) = "800":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Id":       xCampos(2, 1) = "id":        xCampos(2, 2) = "500":    xCampos(2, 3) = "N"
            nTitulo = "Buscando Moneda"
            nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion AS nombre, mae_moneda.id AS cod, mae_moneda.simbolo " _
                + vbCr + " FROM mae_moneda;"
        Case 3 '--tipo documento
            If NulosN(lbl_cod(2).Caption) = 0 Then
                MsgBox "Seleccione una moneda", vbExclamation, xTitulo
                txt_cb(2).SetFocus
                Exit Sub
            End If
        
            ReDim xCampos(3, 3) As String
            xCampos(0, 0) = "Nombre":  xCampos(0, 1) = "nombre":   xCampos(0, 2) = "6500":  xCampos(0, 3) = "C"
            xCampos(1, 0) = "Abrev":   xCampos(1, 1) = "abrev":    xCampos(1, 2) = "700":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Id":      xCampos(2, 1) = "id":       xCampos(2, 2) = "500":   xCampos(2, 3) = "N"

            nTitulo = "Seleccionar el Documento"
            
            nSQL = "SELECT mae_documento.id, mae_documento.descripcion AS nombre, mae_documento.id AS cod, mae_documento.abrev, mae_documento.codsun, mae_documentocta.idcuen " _
                + vbCr + " FROM mae_documento INNER JOIN mae_documentocta ON mae_documento.id = mae_documentocta.iddoc " _
                + vbCr + " WHERE (((mae_documentocta.idmon)= " & NulosN(lbl_cod(2).Caption) & ") AND ((mae_documentocta.tipope)=0)); "

    End Select


    Dim xRs As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    lbl_cb(Index).ToolTipText = xRs.Fields(1) & "" '--NOMBRE
    
    If Index = 2 Then lblMoneda.Caption = NulosC(xRs.Fields(3))
    If Index = 3 Then lblIdCtaDoc.Caption = NulosN(xRs.Fields("idcuen"))
    
    If Trim(lbl_cod(Index).Tag) <> Trim(lbl_cod(Index).Caption) Then
        Select Case Index
            Case 0 '--manual tipo doc
                Fg(0).Rows = Fg(0).FixedRows
                LimpiaText txt_total
        End Select
    End If
    Select Case Index
        Case 0 '--tipo doc
            txt_cb(2).SetFocus '--moneda
        Case 1 '--personal
            txt(1).SetFocus
        Case 2 '--moneda
            txt_cb(3).SetFocus '--tipo doc
        Case 3 '--tipo doc
            txt_cb(1).SetFocus
    End Select

salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
        Me.lbl_cb(Index).Tag = ""
        If Index = 2 Then lblMoneda.Caption = ""
        If Index = 3 Then lblIdCtaDoc.Caption = ""
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
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then Exit Sub

    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--proceso =>> ::manual
            nSQL = "SELECT pla_proceso.id, pla_proceso.descripcion AS nombre, pla_proceso.id AS cod " _
                + vbCr + " FROM pla_proceso " _
                + vbCr + " WHERE (((pla_proceso.enproceso)=-1)) and pla_proceso.id  = " & NulosN(txt_cb(Index).Text) & ""

        Case 1 '--personal
            nSQL = "SELECT pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.id as cod,pla_empleados.numdoc, mae_dociden.abrev AS tipodoc, mae_sexo.abrev AS sexo, Format([pla_empleados].[fchnac],'dd/mm/yyyy') AS fchnac, pla_empleados.numtel, pla_empleados.email " _
                + vbCr + " FROM (mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_sexo.id = pla_empleados.idsex) INNER JOIN pla_periodolaboral ON pla_empleados.id = pla_periodolaboral.idemp " _
                + vbCr + " WHERE (((pla_periodolaboral.fchfin) Is Null)) and pla_empleados.id = " & NulosN(txt_cb(Index).Text) & ";"

        Case 2 '--moneda =>>> ::manual
            nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion AS nombre, mae_moneda.id AS cod, mae_moneda.simbolo " _
                + vbCr + " FROM mae_moneda " _
                + vbCr + " WHERE mae_moneda.id  = " & NulosN(txt_cb(Index).Text) & ""
        Case 3 '--tipo de documento
            If NulosN(lbl_cod(2).Caption) = 0 Then
                MsgBox "Seleccione una moneda", vbExclamation, xTitulo
                txt_cb(2).SetFocus
                Exit Sub
            End If
        
            nSQL = "SELECT mae_documento.id, mae_documento.descripcion AS nombre, mae_documento.id AS cod, mae_documento.abrev, mae_documento.codsun, mae_documentocta.idcuen " _
                + vbCr + " FROM mae_documento INNER JOIN mae_documentocta ON mae_documento.id = mae_documentocta.iddoc " _
                + vbCr + " WHERE (((mae_documentocta.idmon)=" & NulosN(lbl_cod(2).Caption) & ") AND ((mae_documentocta.tipope)=0)) and mae_documento.id  = " & NulosN(txt_cb(Index).Text) & ""
    
    End Select

    If xCon.State = 0 Then GoTo salir

    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cod(Index).Caption = RstTmp.Fields(2) & "" '--CODIGO
        lbl_cb(Index).ToolTipText = RstTmp.Fields(1) & "" '--NOMBRE
        '--------------
        If Index = 2 Then lblMoneda.Caption = NulosC(RstTmp.Fields(3))
        If Index = 3 Then lblIdCtaDoc.Caption = NulosN(RstTmp.Fields("idcuen"))
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
'***** OPERACION AUTOMATICO
'****************************************************************************************

Private Sub pAutomatico()

    QueHace = 4
    ActivaTool
    Blanquea True
    Habilitar_Obj True, True
    
    Toolbar1.Buttons(7).ToolTipText = "Grabar Automático"
    Toolbar1.Buttons(8).ToolTipText = "Cancelar Automático"
    
    pBloquearEnAuto True
    '------------
    Fg(1).Editable = flexEDNone ' flexEDKbdMouse
    Fg(1).SelectionMode = flexSelectionByRow
    '------------
    TabOne1.TabVisible(0) = False
    TabOne1.TabVisible(1) = True
    
    txtfechaA(0).Valor = Date
    txtfechaA(1).Valor = Date
    
    LimpiaText txtfechaA
    
    TabOne1.CurrTab = 1
    
    txtfechaA(0).SetFocus
    
End Sub

Private Sub cbA_Click(Index As Integer)
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error

    If QueHace <> 4 Then Exit Sub '--4:automatico
    
    Select Case Index
        Case 0 '--Proceso =>> 0::auto
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"

            nTitulo = "Seleccionar el Proceso"
            nSQL = "SELECT pla_proceso.id, pla_proceso.descripcion AS nombre, pla_proceso.id AS cod " _
                + vbCr + " FROM pla_proceso " _
                + vbCr + " WHERE (((pla_proceso.enproceso)=-1)); "

        Case 1 '--auto categoria
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Categoría"
            nSQL = "SELECT mae_categoria.id, mae_categoria.descripcion AS nombre, mae_categoria.id AS cod " _
                + vbCr + " FROM mae_categoria;"

        Case 2 '--moneda =>>> 2::auto
            ReDim xCampos(3, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Símbolo":  xCampos(1, 1) = "simbolo":   xCampos(1, 2) = "800":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Id":       xCampos(2, 1) = "id":        xCampos(2, 2) = "500":    xCampos(2, 3) = "N"
            nTitulo = "Buscando Moneda"
            nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion AS nombre, mae_moneda.id AS cod, mae_moneda.simbolo " _
                + vbCr + " FROM mae_moneda;"
        Case 3 '--tipo documento
            If NulosN(lbl_codA(2).Caption) = 0 Then
                MsgBox "Seleccione una moneda", vbExclamation, xTitulo
                txt_cbA(2).SetFocus
                Exit Sub
            End If
        
            ReDim xCampos(3, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Abrev":   xCampos(1, 1) = "abrev":    xCampos(1, 2) = "700":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Id":       xCampos(2, 1) = "id":        xCampos(2, 2) = "500":    xCampos(2, 3) = "N"

            nTitulo = "Seleccionar el Documento"
            nSQL = "SELECT mae_documento.id, mae_documento.descripcion AS nombre, mae_documento.id AS cod, mae_documento.abrev, mae_documento.codsun, mae_documentocta.idcuen " _
                + vbCr + " FROM mae_documento INNER JOIN mae_documentocta ON mae_documento.id = mae_documentocta.iddoc " _
                + vbCr + " WHERE (((mae_documentocta.idmon)=" & NulosN(lbl_codA(2).Caption) & ") AND ((mae_documentocta.tipope)=0));"

        Case 4 '--por implementar
            MsgBox "Pendiente", vbCritical
            Exit Sub
    End Select


    Dim xRs As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir

    lbl_codA(Index).Tag = lbl_codA(Index).Caption

    txt_cbA(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cbA(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_codA(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    lbl_cbA(Index).ToolTipText = xRs.Fields(1) & "" '--NOMBRE
    
    If Index = 2 Then lblMonedaA.Caption = NulosC(xRs.Fields(3))
    If Index = 3 Then lblIdCtaDocA.Caption = NulosN(xRs.Fields("idcuen"))
        
    If Trim(lbl_codA(Index).Tag) <> Trim(lbl_codA(Index).Caption) Then
        Select Case Index
            Case 0 '--auto tipo doc
                Fg(1).Rows = Fg(1).FixedRows
                LimpiaText txt_totalA
                Set RstIngreso = Nothing
                Set RstDescuento = Nothing
                Set RstAportacion = Nothing
        End Select
    End If
    Select Case Index
        Case 0 '--auto tipo doc
            txt_cbA(1).SetFocus '--categoria
        Case 1 '--auto categoria
            txt_cbA(2).SetFocus '--moneda
        Case 2 '--moneda
            txt_cbA(3).SetFocus
    End Select

salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cbA_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cbA_Change(Index As Integer)
    If txt_cbA(Index).Text = "" Then
        Me.lbl_cbA(Index).Caption = ""
        Me.lbl_codA(Index).Caption = ""
        Me.lbl_cbA(Index).Tag = ""
        If Index = 2 Then lblMonedaA.Caption = ""
        If Index = 3 Then lblIdCtaDocA.Caption = ""
    End If
End Sub

Private Sub txt_cbA_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If txt_cbA(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cbA_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cbA_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txt_cbA_Validate(Index As Integer, Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If txt_cbA(Index).Text = "" Then Exit Sub

    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--proceso =>> ::auto
            nSQL = "SELECT pla_proceso.id, pla_proceso.descripcion AS nombre, pla_proceso.id AS cod " _
                + vbCr + " FROM pla_proceso " _
                + vbCr + " WHERE (((pla_proceso.enproceso)=-1)) and pla_proceso.id  = " & NulosN(txt_cbA(Index).Text) & ""

        Case 1 '--personal
            nSQL = "SELECT mae_categoria.id, mae_categoria.descripcion AS nombre, mae_categoria.id AS cod " _
                + vbCr + " FROM mae_categoria " _
                + vbCr + " WHERE mae_categoria.id  = " & NulosN(txt_cbA(Index).Text) & ""

        Case 2 '--moneda =>>> ::auto
            nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion AS nombre, mae_moneda.id AS cod, mae_moneda.simbolo " _
                + vbCr + " FROM mae_moneda " _
                + vbCr + " WHERE mae_moneda.id  = " & NulosN(txt_cbA(Index).Text) & ""
    
        Case 3 '--tipo doc
            If NulosN(lbl_codA(2).Caption) = 0 Then
                MsgBox "Seleccione una moneda", vbExclamation, xTitulo
                txt_cbA(2).SetFocus
                Exit Sub
            End If
        
            nSQL = "SELECT mae_documento.id, mae_documento.descripcion AS nombre, mae_documento.id AS cod, mae_documento.abrev, mae_documento.codsun, mae_documentocta.idcuen " _
                + vbCr + " FROM mae_documento INNER JOIN mae_documentocta ON mae_documento.id = mae_documentocta.iddoc " _
                + vbCr + " WHERE (((mae_documentocta.idmon)=" & NulosN(lbl_codA(2).Caption) & ") AND ((mae_documentocta.tipope)=0)) and mae_documento.id  = " & NulosN(txt_cbA(Index).Text) & ""

        
    End Select

    If xCon.State = 0 Then GoTo salir

    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo salir

    lbl_codA(Index).Tag = lbl_codA(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cbA(Index).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cbA(Index).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_codA(Index).Caption = RstTmp.Fields(2) & "" '--CODIGO
        lbl_cbA(Index).ToolTipText = RstTmp.Fields(1) & "" '--NOMBRE
        
        If Index = 2 Then lblMonedaA.Caption = NulosC(RstTmp.Fields(3))
        If Index = 3 Then lblIdCtaDocA.Caption = NulosN(RstTmp.Fields("idcuen"))
    Else
        txt_cbA(Index).Text = "":    lbl_cbA(Index).Caption = "":    lbl_codA(Index).Caption = ""
    End If
    
    
    '--------------
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cbA_Validate (" + CStr(Index) + ")"
    Exit Sub
salir:
    Set RstTmp = Nothing
    txt_cbA(Index).Text = ""
End Sub

Private Sub CmdAuto_Click(Index As Integer)
    Select Case Index
        Case 0 '--procesar
            pAutoCalculos
        Case 1 '--ver detalle
            If Fg(1).FixedRows = Fg(1).Rows - 1 Then
                MsgBox "No hay Registros para mostrar el detalle", vbInformation, xTitulo
                Exit Sub
            End If
        
            If Fg(1).Row < Fg(1).FixedRows Then
                MsgBox "Seleccione de nuevo el Registro que desea ver el detalle", vbInformation, xTitulo
                Exit Sub
            End If

            TabOne1.CurrTab = 2
            Toolbar1.Buttons(7).ToolTipText = "Grabar Automático - Detalle"
            
            pAutoPonerDatosEnDetalle Fg(1).Row

        Case 2 '--cargar lista de personal
            '--cargar lista de personal en la grilla para proceder a efectuar el calculo
            pAutoCargarDatosPersonal
            '-----
            If Fg(1).Rows > Fg(1).FixedRows Then
                habilitar CmdAuto, True
            Else
                habilitar CmdAuto, False
                CmdAuto(2).Enabled = True
            End If
            TabOne1.TabEnabled(2) = True
    End Select
End Sub

Private Sub pAutoPonerDatosEnDetalle(mRow&)
    '--proceso
    If NulosN(txt_cbA(0).Text) <> 0 Then
        txt_cb(0).Text = txt_cbA(0).Text
        txt_cb_Validate 0, False
    End If
    '--personal
    If NulosN(Fg(1).TextMatrix(mRow, 2)) <> 0 Then
        txt_cb(1).Text = NulosN(Fg(1).TextMatrix(mRow, 2))
        lbl_cb(1).Caption = NulosC(Fg(1).TextMatrix(mRow, 5))
        lbl_cod(1).Caption = NulosN(Fg(1).TextMatrix(mRow, 2))
    End If
    '--moneda
    If NulosN(Fg(1).TextMatrix(mRow, 3)) <> 0 Then
        txt_cb(2).Text = NulosN(Fg(1).TextMatrix(mRow, 3))
        txt_cb_Validate 2, False
        '---colocando el tipo documento
        
        If NulosN(lbl_codA(3).Caption) <> 0 Then
            txt_cb(3).Text = txt_cbA(3).Text
            lbl_cb(3).Caption = lbl_cbA(3).Caption
            lbl_cod(3).Caption = lbl_codA(3).Caption
            lblIdCtaDoc.Caption = lblIdCtaDocA.Caption
        End If
        
    End If
    '--fch.emision
    If IsDate(Fg(1).TextMatrix(mRow, 10)) = True Then
        txtfecha(0).Valor = CDate(Fg(1).TextMatrix(mRow, 10))
    Else
        txtfecha(0).Valor = ""
    End If
    '--fch.pago
    If IsDate(Fg(1).TextMatrix(mRow, 11)) = True Then
        txtfecha(1).Valor = CDate(Fg(1).TextMatrix(mRow, 11))
    Else
        txtfecha(1).Valor = ""
    End If
    
    txt(1).Text = Fg(1).TextMatrix(mRow, 13)
    txt(2).Text = Fg(1).TextMatrix(mRow, 14)
    
    '--poner datos del detalla de los conceptos
    pCargarConceptosDetalle NulosN(Fg(1).TextMatrix(mRow, 2))

End Sub

Private Sub pAutoCargarDatosPersonal()
    If NulosN(txt_cbA(1).Text) = 0 Then
        MsgBox "Seleccione la Categoría", vbExclamation, xTitulo
        txt_cbA(1).SetFocus
        Exit Sub
    End If
    If NulosN(txt_cbA(0).Text) = 0 Then
        MsgBox "Seleccione el Tipo de Proceso", vbExclamation, xTitulo
        txt_cbA(0).SetFocus
        Exit Sub
    End If
    If NulosN(txt_cbA(2).Text) = 0 Then
        MsgBox "Seleccione la Moneda", vbExclamation, xTitulo
        txt_cbA(2).SetFocus
        Exit Sub
    End If
    
    If xMes = 0 Or xMes = 13 Then
        MsgBox "Seleccione el periodo ", vbInformation, xTitulo
        Exit Sub
    End If
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String

    pConfigurarGrilla
    '-----------------------------
    Set RstIngreso = Nothing
    Set RstDescuento = Nothing
    Set RstAportacion = Nothing
    '-----------------------------
        
    pCagarListaPersonal RstTmp, NulosN(txt_cbA(0).Text), NulosN(txt_cbA(1).Text), AnoTra, xMes
    
    Fg(1).Rows = Fg(1).FixedRows
    If RstTmp.State = 0 Then GoTo salir

    Agregando = True

    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        Fg(1).Rows = Fg(1).Rows + 1
        Fg(1).TextMatrix(Fg(1).Rows - 1, 1) = NulosN(RstTmp("idbol"))
        If NulosN(RstTmp("idbol")) <> 0 Then
            '--cargar los detalles de los documentos, porque pueda darse el caso de generar la lista de personal
            '--y ver el detalle de un registro que tiene movimiento
            pConceptoDocumentoEmp RstIngreso, NulosN(RstTmp("idbol")), e_Remuneracion
            pConceptoDocumentoEmp RstDescuento, NulosN(RstTmp("idbol")), e_Descuento
            pConceptoDocumentoEmp RstAportacion, NulosN(RstTmp("idbol")), e_Aportacion
            Fg(1).TextMatrix(Fg(1).Rows - 1, 4) = 0
        Else
            Fg(1).TextMatrix(Fg(1).Rows - 1, 4) = -1
        End If
        
        Fg(1).TextMatrix(Fg(1).Rows - 1, 2) = NulosC(RstTmp("idemp"))
        Fg(1).TextMatrix(Fg(1).Rows - 1, 3) = NulosC(RstTmp("idmon"))
        
        Fg(1).TextMatrix(Fg(1).Rows - 1, 5) = NulosC(RstTmp("nombres"))
        Fg(1).TextMatrix(Fg(1).Rows - 1, 6) = NulosC(RstTmp("catabrev"))
        Fg(1).TextMatrix(Fg(1).Rows - 1, 7) = NulosC(RstTmp("cargo"))
        Fg(1).TextMatrix(Fg(1).Rows - 1, 8) = Format(NulosC(RstTmp("ingreso")), FORMAT_DATE)
        Fg(1).TextMatrix(Fg(1).Rows - 1, 9) = ConvertHora(NulosN(RstTmp("totseg")))
        Fg(1).TextMatrix(Fg(1).Rows - 1, 10) = Format(NulosC(RstTmp("fchdoc")), FORMAT_DATE)
        Fg(1).TextMatrix(Fg(1).Rows - 1, 11) = Format(NulosC(RstTmp("fchpago")), FORMAT_DATE)
        
        Fg(1).TextMatrix(Fg(1).Rows - 1, 12) = NulosC(RstTmp("simbolo"))
        
        Fg(1).TextMatrix(Fg(1).Rows - 1, 13) = NulosC(RstTmp("numser"))
        Fg(1).TextMatrix(Fg(1).Rows - 1, 14) = NulosC(RstTmp("numdoc"))
                
        Fg(1).TextMatrix(Fg(1).Rows - 1, 15) = Format(RstTmp("impingr"), FORMAT_MONTO)
        Fg(1).TextMatrix(Fg(1).Rows - 1, 16) = Format(RstTmp("impdesc"), FORMAT_MONTO)
        Fg(1).TextMatrix(Fg(1).Rows - 1, 17) = Format(RstTmp("impapor"), FORMAT_MONTO)
        Fg(1).TextMatrix(Fg(1).Rows - 1, 18) = Format(RstTmp("imptot"), FORMAT_MONTO)
        
        RstTmp.MoveNext
    Loop
    '--aplicando el fonfo
    GRID_COLOR_FONDO Fg(1), Fg(1).FixedRows, 15, Fg(1).Rows - 1, 15, &HE7FEFC
    GRID_COLOR_FONDO Fg(1), Fg(1).FixedRows, 16, Fg(1).Rows - 1, 16, &HC0E0FF
    GRID_COLOR_FONDO Fg(1), Fg(1).FixedRows, 17, Fg(1).Rows - 1, 17, &HC0C0FF
    GRID_COLOR_FONDO Fg(1), Fg(1).FixedRows, 18, Fg(1).Rows - 1, 18, &HFFD3A8

salir:
    Set RstTmp = Nothing
    Agregando = False
End Sub


Private Sub pAutoCalculos()
    If Fg(1).Rows = Fg(1).FixedRows Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If
    If NulosN(txt_cbA(0).Text) = 0 Then
        MsgBox "Seleccione el Tipo de Documento", vbExclamation, xTitulo
        txt_cbA(0).SetFocus
        Exit Sub
    End If
    If NulosN(txt_cbA(2).Text) = 0 Then
        MsgBox "Seleccione la Moneda", vbExclamation, xTitulo
        txt_cbA(2).SetFocus
        Exit Sub
    End If

    '--ver si han seleccionado registros para procesar el calculo
    Dim fExisteRegistro As Boolean
    Dim K&
    With Fg(1)
        For K = .FixedRows To .Rows - 1
            If NulosN(.TextMatrix(K, 4)) = -1 Then
                fExisteRegistro = True
                Exit For
            End If
        Next K
    End With
    If fExisteRegistro = False Then
        MsgBox "Seleccione los registros que desea para efectuar el Cálculo", vbInformation, xTitulo
        Exit Sub
    End If
    '-----------------------------------------------------
    pBloquearEnAuto False
    '-----------------------------------------------------
    
    Set RstIngreso = Nothing
    Set RstDescuento = Nothing
    Set RstAportacion = Nothing
    
    '-----------------------------------------------------
       
    Dim TotIngreso As Double
    Dim TotDescuento As Double
    Dim TotAporte As Double

    Dim RstIngresoTmp As New ADODB.Recordset
    Dim RstDescuentoTmp As New ADODB.Recordset
    Dim RstAporteTmp As New ADODB.Recordset
    
    '-----------------------------------------------------
    
    Dim nSerie As String
    Dim mNumeroDoc
    
    nSerie = "0001"
    mNumeroDoc = fNumeroDoc(nSerie, NulosN(lbl_codA(3).Caption))   '--obtener el numero de serie
    
    For K = Fg(1).FixedRows To Fg(1).Rows - 1
        Fg(1).Row = K
        If NulosN(Fg(1).TextMatrix(K, 4)) = -1 Then
            DoEvents
            '--------------------------------------------------
            '--eliminar registros del empleado seleccionado, es útil cuando el usuario procesa mas de una vez
            '--finalidad:: eliminar duplicados
            RstRegistroEliminar RstIngreso, "mIdEmp", NulosN(Fg(1).TextMatrix(K, 2)), True
            RstRegistroEliminar RstDescuento, "mIdEmp", NulosN(Fg(1).TextMatrix(K, 2)), True
            RstRegistroEliminar RstAportacion, "mIdEmp", NulosN(Fg(1).TextMatrix(K, 2)), True
            '---------------------------------------------------
            TotIngreso = 0: TotDescuento = 0: TotAporte = 0
            Fg(1).TextMatrix(K, 13) = ""
            Fg(1).TextMatrix(K, 14) = ""
            Fg(1).TextMatrix(K, 15) = ""
            Fg(1).TextMatrix(K, 16) = ""
            Fg(1).TextMatrix(K, 17) = ""
            Fg(1).TextMatrix(K, 18) = ""
            '---------------------------------------------------
            DoEvents
            '********************************************************************************************
            '--finalidad::efectuar el calculo
            pObtenerValoresConcepto RstIngresoTmp, RstDescuentoTmp, RstAporteTmp, _
                                    NulosN(Fg(1).TextMatrix(K, 2)), NulosN(txt_cbA(0).Text)

            '--obteniendo los totales
            TotIngreso = RstRegistroSumar(RstIngresoTmp, "imptot")
            TotDescuento = RstRegistroSumar(RstDescuentoTmp, "imptot")
            TotAporte = RstRegistroSumar(RstAporteTmp, "imptot")
            '********************************************************************************************
            '--copiando datos   finalidad::actualizar la grilla en funcion a los valores seleccionados
            If TotIngreso <> 0 Or TotDescuento <> 0 Or TotAporte <> 0 Then
                Fg(1).TextMatrix(K, 10) = Format(txtfechaA(0).Valor, FORMAT_DATE)
                Fg(1).TextMatrix(K, 11) = Format(txtfechaA(1).Valor, FORMAT_DATE)
                '--moneda
                Fg(1).TextMatrix(K, 3) = NulosN(txt_cbA(2).Text) '--codigo
                Fg(1).TextMatrix(K, 12) = NulosC(lblMonedaA.Caption)  '--simbolo
                
                
                Fg(1).TextMatrix(K, 13) = nSerie   '--serie
                Fg(1).TextMatrix(K, 14) = Format(mNumeroDoc, "0000000000") '--numero doc
                mNumeroDoc = mNumeroDoc + 1 '--para el siguiente registro
                '--
                Fg(1).TextMatrix(K, 15) = Format(TotIngreso, FORMAT_MONTO)
                Fg(1).TextMatrix(K, 16) = Format(TotDescuento, FORMAT_MONTO)
                Fg(1).TextMatrix(K, 17) = Format(TotAporte, FORMAT_MONTO)
    
                Fg(1).TextMatrix(K, 18) = Format(TotIngreso - TotDescuento, FORMAT_MONTO)
            End If
            '********************************************************************************************
            '--limpiando los temporales
            Set RstIngresoTmp = Nothing
            Set RstDescuentoTmp = Nothing
            Set RstAporteTmp = Nothing
            '----
        End If
    Next

    pTotalizar True
    
    '---------
    pBloquearEnAuto True
    '---------
    
End Sub


Private Sub pTotalizar(Optional fEsAutomatico As Boolean = True)
    '--efectua el calculo de los totales segun sea el caso
    '--automatico o manual
    If fEsAutomatico = True Then
        txt_totalA(0).Text = Format(GRID_SUMAR_COL(Fg(1), 15), FORMAT_MONTO) '--tot ingresos
        txt_totalA(1).Text = Format(GRID_SUMAR_COL(Fg(1), 16), FORMAT_MONTO) '--tot descuento
        txt_totalA(2).Text = Format(GRID_SUMAR_COL(Fg(1), 17), FORMAT_MONTO) '--tot aporte
        txt_totalA(3).Text = Format(GRID_SUMAR_COL(Fg(1), 18), FORMAT_MONTO) '--tot a pagar
    Else
        txt_total(0).Text = Format(GRID_SUMAR_COL(Fg(0), 4), FORMAT_MONTO) '--tot ingresos
        txt_total(1).Text = Format(GRID_SUMAR_COL(Fg(0), 9), FORMAT_MONTO) '--tot descuento
        txt_total(2).Text = Format(GRID_SUMAR_COL(Fg(0), 14), FORMAT_MONTO) '--tot aporte
        txt_total(3).Text = Format(NulosN(txt_total(0).Text) - NulosN(txt_total(1).Text), FORMAT_MONTO) '--tot a pagar

    End If

End Sub


Private Sub pObtenerValoresConcepto(RstIngresoTmp As ADODB.Recordset, _
                                    RstDescuentoTmp As ADODB.Recordset, _
                                    RstAporteTmp As ADODB.Recordset, mIdEmp&, mIdDocumento&)
                                    
    '--
    '--

    Dim TotIngreso As Double
    Dim TotDescuento As Double
    Dim TotAporte As Double

    Dim RstCptoEmp As New ADODB.Recordset
    Dim RstCptoValores As New ADODB.Recordset
    Dim RstCptoFormulas As New ADODB.Recordset

    DoEvents
    '---------------------------------------------------
    '--cargar los ingresos(remuneraciones)
    pConceptoSueldoAsignadoEmp RstIngresoTmp, mIdEmp, mIdDocumento, AnoTra, xMes, e_Remuneracion
    '--cargar los descuentos
    pConceptoSueldoAsignadoEmp RstDescuentoTmp, mIdEmp, mIdDocumento, AnoTra, xMes, e_Descuento
    '--cargar las aportaciones
    pConceptoSueldoAsignadoEmp RstAporteTmp, mIdEmp, mIdDocumento, AnoTra, xMes, e_Aportacion
    
    '--si el personal esta dado de baja => no hacer nada
    If RstIngresoTmp.State = 0 And RstDescuentoTmp.State = 0 And RstAporteTmp.State = 0 Then Exit Sub
    
    '--unir ingresos, descuentos,aportes en un solo recordset
    DEFINIR_RST_TMP RstCptoEmp, RstIngresoTmp
    CARGAR_RST_TMP RstCptoEmp, RstIngresoTmp
    CARGAR_RST_TMP RstCptoEmp, RstDescuentoTmp
    CARGAR_RST_TMP RstCptoEmp, RstAporteTmp

    '--cargar los conceptos que tienen formulas
    pConceptosCagarEnFormula RstCptoFormulas

    '--cargar los valores que se usaran en los conceptos ejm. SNP,AFP's, Sueldo Basico,Bonificacion(si no esta en planilla)
    pConceptoCagarDetalleFormula RstCptoValores, mIdEmp, mIdDocumento, AnoTra, xMes

    '--efectuar el calculo de ingreso, descuento y aporte
    pConceptoEfectuarCalculo RstIngresoTmp, RstDescuentoTmp, RstAporteTmp, _
        TotIngreso, TotDescuento, TotAporte, _
        RstCptoFormulas, RstCptoValores, RstCptoEmp
        
    '--cargando a los recordset's generales ACUMULADOS
    '--esto servira cuando se desee ver el detalle de algun registro antes de grabar el automatico,
    '--asi como en ingreso manual
    If RstIngreso.State = 0 Then DEFINIR_RST_TMP RstIngreso, RstIngresoTmp
    If RstDescuento.State = 0 Then DEFINIR_RST_TMP RstDescuento, RstDescuentoTmp
    If RstAportacion.State = 0 Then DEFINIR_RST_TMP RstAportacion, RstAporteTmp
    
    '--cargar los temporales a recordset local, de uso exclusivo en el formulario
    CARGAR_RST_TMP RstIngreso, RstIngresoTmp
    CARGAR_RST_TMP RstDescuento, RstDescuentoTmp
    CARGAR_RST_TMP RstAportacion, RstAporteTmp

    '----

    Set RstCptoFormulas = Nothing
    Set RstCptoEmp = Nothing
    Set RstCptoValores = Nothing

End Sub


'****************************************************************************************
'******* OPERACION MANUAL
Private Sub cmdManual_Click(Index As Integer)
    Select Case Index
        Case 0 '--procesar manual auto
            pManualCalculos
        Case 1 '--agregar
            pRegistroAdd
        Case 2 '--eliminar
            pRegistroDel
    End Select
    
End Sub

Private Sub pManualCalculos()
    On Error GoTo error
    
    If NulosN(txt_cb(0).Text) = 0 Then
        MsgBox "Seleccione el Tipo de Documento", vbExclamation, xTitulo
        txt_cb(0).SetFocus
        Exit Sub
    End If
    If NulosN(txt_cb(1).Text) = 0 Then
        MsgBox "Seleccione el Personal", vbExclamation, xTitulo
        txt_cb(1).SetFocus
        Exit Sub
    End If
    If NulosN(txt_cb(2).Text) = 0 Then
        MsgBox "Seleccione la Moneda", vbExclamation, xTitulo
        txt_cb(2).SetFocus
        Exit Sub
    End If
    If NulosC(txt(1).Text) = "" Then
        MsgBox "Ingrese el Número de Serie", vbExclamation, xTitulo
        txt(1).SetFocus
        Exit Sub
    End If
    If NulosC(txt(2).Text) = "" Then
        MsgBox "Ingrese el Número de Documento", vbExclamation, xTitulo
        txt(2).SetFocus
        Exit Sub
    End If
    Dim TotIngreso As Double
    Dim TotDescuento As Double
    Dim TotAporte As Double

    Dim RstIngresoTmp As New ADODB.Recordset
    Dim RstDescuentoTmp As New ADODB.Recordset
    Dim RstAporteTmp As New ADODB.Recordset
    '--------------------------------------------------
    '--limpiar datos
    Fg(0).Rows = Fg(0).FixedRows
    LimpiaText txt_total
    DoEvents
    Me.MousePointer = vbHourglass
    '--------------------------------------------------
    '--eliminar registros del empleado seleccionado
    RstRegistroEliminar RstIngreso, "mIdEmp", NulosN(txt_cb(1).Text), True
    RstRegistroEliminar RstDescuento, "mIdEmp", NulosN(txt_cb(1).Text), True
    RstRegistroEliminar RstAportacion, "mIdEmp", NulosN(txt_cb(1).Text), True
    '---------------------------------------------------
    pObtenerValoresConcepto RstIngresoTmp, RstDescuentoTmp, RstAporteTmp, _
                            NulosN(txt_cb(1).Text), NulosN(txt_cb(0).Text)
    '---------------------------------------------------
    If RstIngresoTmp.State = 0 And RstDescuentoTmp.State = 0 And RstAporteTmp.State = 0 Then GoTo salir
    
    '--obteniendo los totales
    TotIngreso = RstRegistroSumar(RstIngresoTmp, "imptot")
    TotDescuento = RstRegistroSumar(RstDescuentoTmp, "imptot")
    TotAporte = RstRegistroSumar(RstAporteTmp, "imptot")
    
    If QueHace = 4 Then '--automatico
        '--copiando datos al proceso automatico
        Fg(1).TextMatrix(Fg(1).Row, 4) = -1 '--check activo
        
        Fg(1).TextMatrix(Fg(1).Row, 10) = Format(txtfecha(0).Valor, FORMAT_DATE)
        Fg(1).TextMatrix(Fg(1).Row, 11) = Format(txtfecha(1).Valor, FORMAT_DATE)
        '--moneda
        Fg(1).TextMatrix(Fg(1).Row, 3) = NulosN(txt_cb(2).Text) '--codigo
        Fg(1).TextMatrix(Fg(1).Row, 12) = NulosC(lblMoneda.Caption)  '--simbolo
        '--
        Fg(1).TextMatrix(Fg(1).Row, 13) = txt(1).Text   '--serie
        Fg(1).TextMatrix(Fg(1).Row, 14) = txt(2).Text   '--numero
        
        
        Fg(1).TextMatrix(Fg(1).Row, 15) = Format(TotIngreso, FORMAT_MONTO)
        Fg(1).TextMatrix(Fg(1).Row, 16) = Format(TotDescuento, FORMAT_MONTO)
        Fg(1).TextMatrix(Fg(1).Row, 17) = Format(TotAporte, FORMAT_MONTO)
    
        Fg(1).TextMatrix(Fg(1).Row, 18) = Format(TotIngreso - TotDescuento, FORMAT_MONTO)
    End If

    Set RstIngresoTmp = Nothing
    Set RstDescuentoTmp = Nothing
    Set RstAporteTmp = Nothing
    '--cargando los datos de nuevo al registro
    pCargarConceptosDetalle NulosN(txt_cb(1).Text)
salir:
    '--totalizar las categorias de los conceptos
    pTotalizar False
    If QueHace = 4 Then pTotalizar True
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    pTotalizar False
    If QueHace = 4 Then pTotalizar True
    SHOW_ERROR Me.Name, "pManualCalculos"
End Sub


Private Sub pCargarConceptosDetalle(mIdEmp&)
    Dim mCantidadFilas&
    Dim mRow&
    
    Fg(0).Rows = Fg(0).FixedRows
    If RstIngreso.State = 0 Or RstDescuento.State = 0 Or RstAportacion.State = 0 Then Exit Sub
    '************************************************************************************************
    RstIngreso.Filter = "mIdEmp=" & mIdEmp
    RstDescuento.Filter = "mIdEmp=" & mIdEmp
    RstAportacion.Filter = "mIdEmp=" & mIdEmp

    Agregando = True

    If RstIngreso.RecordCount <> 0 Then RstIngreso.MoveFirst
    If RstDescuento.RecordCount <> 0 Then RstDescuento.MoveFirst
    If RstAportacion.RecordCount <> 0 Then RstAportacion.MoveFirst
    '--obtener la cantidad de filas a considerar maximo
    mCantidadFilas = RstIngreso.RecordCount
    If mCantidadFilas < RstDescuento.RecordCount Then mCantidadFilas = RstDescuento.RecordCount
    If mCantidadFilas < RstAportacion.RecordCount Then mCantidadFilas = RstAportacion.RecordCount
    '--colocar la cantidad de filas
    Fg(0).Rows = Fg(0).FixedRows
    Fg(0).Rows = Fg(0).FixedRows + mCantidadFilas
    '--colocando los conceptos de ingresos
    mRow = Fg(0).FixedRows
    Do While Not RstIngreso.EOF
        If NulosN(RstIngreso("imptot")) <> 0 Then
            Fg(0).TextMatrix(mRow, 1) = NulosN(RstIngreso("idcpto"))
            Fg(0).TextMatrix(mRow, 2) = NulosN(RstIngreso("aplanilla"))
            Fg(0).TextMatrix(mRow, 3) = NulosC(RstIngreso("concepto"))
            Fg(0).TextMatrix(mRow, 4) = Format(NulosN(RstIngreso("imptot")), FORMAT_MONTO)
            mRow = mRow + 1
        End If
        RstIngreso.MoveNext
        
    Loop
    '--colocando los conceptos de descuentos
    mRow = Fg(0).FixedRows
    Do While Not RstDescuento.EOF
        If NulosN(RstDescuento("imptot")) <> 0 Then
            Fg(0).TextMatrix(mRow, 6) = NulosN(RstDescuento("idcpto"))
            Fg(0).TextMatrix(mRow, 7) = NulosN(RstDescuento("aplanilla"))
            Fg(0).TextMatrix(mRow, 8) = NulosC(RstDescuento("concepto"))
            Fg(0).TextMatrix(mRow, 9) = Format(NulosN(RstDescuento("imptot")), FORMAT_MONTO)
            mRow = mRow + 1
        End If
        RstDescuento.MoveNext
    Loop
    '--colocando los conceptos de aportes
    mRow = Fg(0).FixedRows
    Do While Not RstAportacion.EOF
        If NulosN(RstAportacion("imptot")) <> 0 Then
            Fg(0).TextMatrix(mRow, 11) = NulosN(RstAportacion("idcpto"))
            Fg(0).TextMatrix(mRow, 12) = NulosN(RstAportacion("aplanilla"))
            Fg(0).TextMatrix(mRow, 13) = NulosC(RstAportacion("concepto"))
            Fg(0).TextMatrix(mRow, 14) = Format(NulosN(RstAportacion("imptot")), FORMAT_MONTO)
            mRow = mRow + 1
        End If
        RstAportacion.MoveNext
    Loop
    
    If Fg(0).Rows > Fg(0).FixedRows Then
        '--eliminar filas que no se usan
        Do While NulosN(Fg(0).TextMatrix(Fg(0).Rows - 1, 1)) = 0 And NulosN(Fg(0).TextMatrix(Fg(0).Rows - 1, 6)) = 0 And NulosN(Fg(0).TextMatrix(Fg(0).Rows - 1, 11)) = 0 And Fg(0).Rows > Fg(0).FixedRows
            Fg(0).RemoveItem Fg(0).Rows - 1
        Loop
        '--colocando los colores
        GRID_COLOR_FONDO Fg(0), Fg(0).FixedRows, 1, Fg(0).Rows - 1, 4, &HE7FEFC
        GRID_COLOR_FONDO Fg(0), Fg(0).FixedRows, 6, Fg(0).Rows - 1, 9, &HC0E0FF
        GRID_COLOR_FONDO Fg(0), Fg(0).FixedRows, 11, Fg(0).Rows - 1, 14, &HC0C0FF
        '--separar entre categoria de conceptos
        GRID_COLOR_FONDO Fg(0), Fg(0).FixedRows, 5, Fg(0).Rows - 1, 5, &HDDDDFF
        GRID_COLOR_FONDO Fg(0), Fg(0).FixedRows, 10, Fg(0).Rows - 1, 10, &HDDDDFF
        
'        GRID_COMBINAR fg(0), 0, 5, fg(0).Rows - 1, 5, " - ", , False, , , &HDDDDFF
'        GRID_COMBINAR fg(0), 0, 10, fg(0).Rows - 1, 10, " - ", , False, , , &HDDDDFF
   
    End If

    pTotalizar False
    
    Agregando = False

End Sub

Private Sub pBloquearEnAuto(band As Boolean)
    
    If TabOne1.CurrTab = 1 Then
    
        habilitar CmdAuto, band
        habilitar cbA, band
        habilitar txt_cbA, band
        habilitar txtfechaA, band
        
        Fg(1).Enabled = band
        
        TabOne1.TabEnabled(2) = band
        
        Fg(1).SelectionMode = flexSelectionByRow
        If Fg(1).Rows = Fg(1).FixedRows Then
            Fg(1).Row = Fg(1).FixedRows - 1
        Else
            Fg(1).Row = Fg(1).FixedRows
        End If
   
   
        TabOne1.TabEnabled(2) = band
    ElseIf TabOne1.CurrTab = 2 Then
        TabOne1.TabEnabled(1) = band
        
        'habilitar CmdAuto, band
        habilitar cb, band
        habilitar txt_cb, band
        habilitar txtfecha, band
        habilitar txt, band
        habilitar cmdManual, band
    End If
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    Select Case Index
        Case 1, 2: If validar_numero(KeyAscii) = False Then KeyAscii = 0
    
    End Select
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 1
            If txt(Index).Text = "" Then Exit Sub
            If IsNumeric(txt(Index).Text) = False Then
                txt(Index).Text = ""
                Exit Sub
            End If
            txt(Index).Text = Format(txt(Index).Text, "0000")
            
            If QueHace = 1 Then
                If NulosN(lbl_cod(3).Caption) = 0 Then
                    MsgBox "Seleccione el Tipo de documento", vbExclamation, xTitulo
                    txt(Index).Text = ""
                    txt_cb(3).SetFocus
                    Exit Sub
                End If
                txt(2).Text = Format(fNumeroDoc(txt(1).Text, NulosN(lbl_cod(3).Caption)), "0000000000")
                
            End If
            
        Case 2
            If txt(Index).Text = "" Then Exit Sub
            If IsNumeric(txt(Index).Text) = False Then
                txt(Index).Text = ""
                Exit Sub
            End If
            txt(Index).Text = Format(txt(Index).Text, "0000000000")
    End Select
End Sub
'--

'*************Imprimir

Private Sub pImprimir()
    Dim RstEmp As New ADODB.Recordset
    Dim nNumBoleta As String
    Dim nPeriodo As String
    Dim mTotalDias As Integer
    Dim mIdMon As Integer
    Dim nTotalHN As String
    Dim nTotalHE1 As String
    Dim nTotalHE2 As String
    Dim nSQL As String
    Dim mIdEmp&
    
    Dim mTipoImpresion&
    Dim mCuenta&
    Dim K&
    
    Dim mTotalSegundosMes As Long
        
    mCuenta = -1
    
    mIdEmp = NulosN(txt_cb(1).Text) '--codigo del empleado
    
    nPeriodo = "'" & UCase(Format("01/" & xMes & "/01", "mmm")) & " " & AnoTra & "'"
    nNumBoleta = "'" & NulosC(txt(1).Text) & " " & NulosC(txt(2).Text) & "'"
    mIdMon = NulosN(txt_cb(2).Text)
    
    '--de las horas
    mTotalDias = HallaDiasMes(CDate("01/" & xMes & "/" & AnoTra))
    mTotalSegundosMes = mTotalDias * 8
    mTotalSegundosMes = mTotalSegundosMes * 60 * 60
    nTotalHN = "'" & ConvertHora(mTotalSegundosMes) & "'"
    nTotalHE1 = "'00:00:00'"
    nTotalHE2 = "'00:00:00'"
    
    If TabOne1.CurrTab = 0 Then
        mTipoImpresion = 1 '--desde la consulta
    ElseIf TabOne1.CurrTab = 1 Then
        mTipoImpresion = 2 'automatico
        For K = Fg(1).FixedRows To Fg(1).Rows - 1
            If NulosN(Fg(1).TextMatrix(K, 4)) = -1 Then mCuenta = mCuenta + 1
        Next K
        If mCuenta = -1 Then
            MsgBox "Seleccione los registos que desea imprimir", vbExclamation, xTitulo
            Exit Sub
        End If
        Dim RstEmpTmp As New ADODB.Recordset
        For K = Fg(1).FixedRows To Fg(1).Rows - 1
            If NulosN(Fg(1).TextMatrix(K, 4)) = -1 Then
                mIdEmp = NulosN(Fg(1).TextMatrix(K, 2)) '--codigo del empleado
                
                nNumBoleta = "'" & NulosC(Fg(1).TextMatrix(K, 13)) & " " & NulosC(Fg(1).TextMatrix(K, 14)) & "'"
                mIdMon = NulosN(Fg(1).TextMatrix(K, 3))
                
                nSQL = "SELECT " & nPeriodo & " AS periodo, mae_area.descripcion AS area, mae_cargo.descripcion AS cargo,pla_empleados.id as idemp, pla_empleados!apepat+' '+pla_empleados!apemat+', '+pla_empleados!nom AS apenom, pla_empleados.numdoc, pla_categoria1.cuspp, pla_empleados.numessalud, pla_empleados.basico, " & nNumBoleta & " AS numboleta, " & mTotalDias & " AS DiaTrabajo, " & nTotalHN & " AS TotalHN," & nTotalHE1 & " AS TotalHE1, " & nTotalHE2 & " AS TotalHE2, Last(pla_periodolaboral.fchini) AS fchingreso, Last(pla_periodolaboral.fchfin) AS fchcese, " & mIdMon & " AS idmon " _
                    + vbCr + " FROM ((mae_area RIGHT JOIN (pla_empleados LEFT JOIN mae_cargo ON pla_empleados.idcargo = mae_cargo.id) ON mae_area.id = pla_empleados.idarea) LEFT JOIN pla_categoria1 ON pla_empleados.id = pla_categoria1.idemp) LEFT JOIN pla_periodolaboral ON pla_empleados.id = pla_periodolaboral.idemp " _
                    + vbCr + " Group By 'MAY 2008', mae_area.descripcion, mae_cargo.descripcion, pla_empleados!apepat+' '+pla_empleados!apemat+', '+pla_empleados!nom, pla_empleados.numdoc, pla_categoria1.cuspp, pla_empleados.numessalud, pla_empleados.basico, pla_empleados.id " _
                    + vbCr + " Having (((pla_empleados.id) = " & mIdEmp & " )) " _
                    + vbCr + " ORDER BY Last(pla_periodolaboral.fchini), Last(pla_periodolaboral.fchfin);"
                Set RstEmpTmp = Nothing
                RST_Busq RstEmpTmp, nSQL, xCon
                If RstEmp.State = 0 Then DEFINIR_RST_TMP RstEmp, RstEmpTmp
                CARGAR_RST_TMP RstEmp, RstEmpTmp
                
            End If
        Next K
        Set RstEmpTmp = Nothing
    Else
        If QueHace <> 4 Then
            mTipoImpresion = 3 'detalle-normal
        Else
            mTipoImpresion = 4 '--detalle-automatico
        End If
    End If
    
    If mTipoImpresion <> 2 Then
        nSQL = "SELECT " & nPeriodo & " AS periodo, mae_area.descripcion AS area, mae_cargo.descripcion AS cargo,pla_empleados.id as idemp, pla_empleados!apepat+' '+pla_empleados!apemat+', '+pla_empleados!nom AS apenom, pla_empleados.numdoc, pla_categoria1.cuspp, pla_empleados.numessalud, pla_empleados.basico, " & nNumBoleta & " AS numboleta, " & mTotalDias & " AS DiaTrabajo, " & nTotalHN & " AS TotalHN," & nTotalHE1 & " AS TotalHE1, " & nTotalHE2 & " AS TotalHE2, Last(pla_periodolaboral.fchini) AS fchingreso, Last(pla_periodolaboral.fchfin) AS fchcese, " & mIdMon & " AS idmon " _
            + vbCr + " FROM ((mae_area RIGHT JOIN (pla_empleados LEFT JOIN mae_cargo ON pla_empleados.idcargo = mae_cargo.id) ON mae_area.id = pla_empleados.idarea) LEFT JOIN pla_categoria1 ON pla_empleados.id = pla_categoria1.idemp) LEFT JOIN pla_periodolaboral ON pla_empleados.id = pla_periodolaboral.idemp " _
            + vbCr + " Group By mae_area.descripcion, mae_cargo.descripcion, pla_empleados!apepat+' '+pla_empleados!apemat+', '+pla_empleados!nom, pla_empleados.numdoc, pla_categoria1.cuspp, pla_empleados.numessalud, pla_empleados.basico, pla_empleados.id " _
            + vbCr + " Having (((pla_empleados.id) = " & mIdEmp & " )) " _
            + vbCr + " ORDER BY Last(pla_periodolaboral.fchini), Last(pla_periodolaboral.fchfin);"
    
        RST_Busq RstEmp, nSQL, xCon
        
    End If
    
    FrmPrintBoleta.pRecibeRsts RstIngreso, RstDescuento, RstAportacion, RstEmp
    FrmPrintBoleta.Show
    
    

End Sub


'------

Private Sub pRegistroAdd()
    Dim mCol%
    Dim fInsertar As Boolean
    If QueHace = 3 Then Exit Sub
    Agregando = True
    If Fg(0).Rows > Fg(0).FixedRows Then
        If NulosN(Fg(0).TextMatrix(Fg(0).Rows - 1, 1)) = 0 And NulosN(Fg(0).TextMatrix(Fg(0).Rows - 1, 6)) = 0 And NulosN(Fg(0).TextMatrix(Fg(0).Rows - 1, 11)) = 0 Then
            MsgBox "Seleccione un Concepto de Remuneraciones, Descuento o Aportes", vbExclamation, xTitulo
        ElseIf NulosN(Fg(0).TextMatrix(Fg(0).Rows - 1, 1)) <> 0 And NulosN(Fg(0).TextMatrix(Fg(0).Rows - 1, 4)) = 0 Then
            MsgBox "Ingrese el Importe de Remuneraciones " + vbCr + "Concepto: " & NulosC(Fg(0).TextMatrix(Fg(0).Rows - 1, 3)), vbExclamation, xTitulo
        ElseIf NulosN(Fg(0).TextMatrix(Fg(0).Rows - 1, 6)) <> 0 And NulosN(Fg(0).TextMatrix(Fg(0).Rows - 1, 9)) = 0 Then
            MsgBox "Ingrese el Importe de Descuentos " + vbCr + "Concepto: " & NulosC(Fg(0).TextMatrix(Fg(0).Rows - 1, 8)), vbExclamation, xTitulo
        ElseIf NulosN(Fg(0).TextMatrix(Fg(0).Rows - 1, 11)) <> 0 And NulosN(Fg(0).TextMatrix(Fg(0).Rows - 1, 14)) = 0 Then
            MsgBox "Ingrese el Importe de Aportaciones " + vbCr + "Concepto: " & NulosC(Fg(0).TextMatrix(Fg(0).Rows - 1, 13)), vbExclamation, xTitulo
        Else
            fInsertar = True
        End If
    Else
        fInsertar = True
    End If
    
    If fInsertar = True Then
        Fg(0).AddItem ""
        '--colocando los colores
        GRID_COLOR_FONDO Fg(0), Fg(0).Rows - 1, 1, Fg(0).Rows - 1, 4, &HE7FEFC
        GRID_COLOR_FONDO Fg(0), Fg(0).Rows - 1, 6, Fg(0).Rows - 1, 9, &HC0E0FF
        GRID_COLOR_FONDO Fg(0), Fg(0).Rows - 1, 11, Fg(0).Rows - 1, 14, &HC0C0FF
        '--separar entre categoria de conceptos
        GRID_COLOR_FONDO Fg(0), Fg(0).Rows - 1, 5, Fg(0).Rows - 1, 5, &HDDDDFF
        GRID_COLOR_FONDO Fg(0), Fg(0).Rows - 1, 10, Fg(0).Rows - 1, 10, &HDDDDFF
        
    End If
    
    Fg(0).Row = Fg(0).Rows - 1
    Fg(0).Col = 4
    
    Fg(0).SetFocus
    Agregando = False
End Sub


Private Sub pRegistroDel()
    Dim mRowDel&, mRow&
    If Fg(0).Rows = 1 Then Exit Sub
    If Fg(0).Row < 1 Then Exit Sub



    If MsgBox("Seguro desea eliminar el Registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    mRow = Fg(0).Row
    Fg(0).RemoveItem Fg(0).Row
    '--agrupar -----------------------
    Me.MousePointer = vbHourglass
    Agregando = True
    If Fg(0).Rows > Fg(0).FixedRows Then
    
    Else
        Me.MousePointer = vbDefault
        cmdManual(0).SetFocus
        Exit Sub
    End If
    Agregando = False
    
    If Fg(0).Rows > Fg(0).FixedRows Then
        Fg(0).Row = Fg(0).Rows - 1
    ElseIf Fg(0).Rows = Fg(0).FixedRows Then
        Fg(0).Row = Fg(0).FixedRows - 1
    End If
    '------------
    pTotalizar
    '------------
    Fg(0).Col = 4
    Me.MousePointer = vbDefault
    '-------------------------------
End Sub


Private Sub pGenerarAsiento(RstDiario As ADODB.Recordset, nAnoTrabajo, mMesActivo, IDLibro, IDMov, mIdDocPro, mCorr, nAsiento, mTipoCambio, FchDoc, IDcuenta, IDMoneda, mImporte, Optional EsDEBE As Boolean = True)
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
    If EsDEBE = True Then
        If IDMoneda = 1 Then '--DE LA MONEDA
            RstDiario("impdebsol") = mImporte
            RstDiario("impdebdol") = 0
        Else
            RstDiario("impdebsol") = mImporte * mTipoCambio
            RstDiario("impdebdol") = mImporte
        End If
    Else
         If IDMoneda = 1 Then '--DE LA MONEDA
            RstDiario("imphabsol") = mImporte
            RstDiario("imphabdol") = 0
        Else
            RstDiario("imphabsol") = mImporte * mTipoCambio
            RstDiario("imphabdol") = mImporte
        End If
   End If

    RstDiario.Update
End Sub

Private Sub txtfecha_Validate(Index As Integer, Cancel As Boolean)
    If Index <> 0 Then Exit Sub
    If IsDate(txtfecha(0).Valor) = True Then
        LblTipoCambio(1).Caption = HallaTipoCambio(txtfecha(0).Valor, 2, xCon)
    Else
        LblTipoCambio(1).Caption = ""
    End If
End Sub

Private Sub txtfechaA_Validate(Index As Integer, Cancel As Boolean)
    If Index <> 0 Then Exit Sub
    If IsDate(txtfechaA(0).Valor) = True Then
        LblTipoCambio(0).Caption = HallaTipoCambio(txtfechaA(0).Valor, 2, xCon)
    Else
        LblTipoCambio(0).Caption = ""
    End If
End Sub

Private Function fNumeroDoc(mSerie As String, mIdDoc As Long) As Double
    Dim nSQL As String
    Dim RstTmp As New ADODB.Recordset
    Dim mNumeroDoc&
    
    nSQL = "SELECT TOP 1 pla_boleta.numser, pla_boleta.numdoc " _
        + vbCr + " FROM pla_boleta  WHERE (((pla_boleta.numser)='" & mSerie & "')) and  pla_boleta.iddoc = " & mIdDoc & "  ORDER BY pla_boleta.numdoc DESC; "
    
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount <> 0 Then
        mNumeroDoc = NulosN(RstTmp("numdoc")) + 1
    Else
        mNumeroDoc = 1
    End If
    
    fNumeroDoc = mNumeroDoc
    
    Set RstTmp = Nothing
    
End Function
