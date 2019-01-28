VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLicencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Licencia"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   11685
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   540
      Index           =   1
      Left            =   30
      TabIndex        =   31
      Top             =   6435
      Width           =   11625
      Begin VB.CommandButton cmd 
         Caption         =   "Eliminar"
         Height          =   345
         Index           =   2
         Left            =   2880
         TabIndex        =   34
         Top             =   105
         Width           =   1200
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Agregar"
         Height          =   345
         Index           =   0
         Left            =   75
         TabIndex        =   33
         Top             =   105
         Width           =   1200
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Modificar"
         Height          =   345
         Index           =   1
         Left            =   1320
         TabIndex        =   32
         Top             =   105
         Width           =   1200
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   2
         X1              =   -15
         X2              =   12000
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line lin 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   3
         X1              =   -30
         X2              =   12000
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   3
         X1              =   11610
         X2              =   11610
         Y1              =   -30
         Y2              =   5455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   800
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   420
      Index           =   12
      Left            =   30
      TabIndex        =   29
      Top             =   375
      Width           =   11625
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   4
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   11610
         X2              =   11610
         Y1              =   -30
         Y2              =   5455
      End
      Begin VB.Line lin 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   12000
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   -15
         X2              =   12000
         Y1              =   15
         Y2              =   15
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
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   1770
      End
   End
   Begin VB.Frame FraEditor 
      BorderStyle     =   0  'None
      Height          =   3885
      Left            =   3285
      TabIndex        =   1
      Top             =   1545
      Visible         =   0   'False
      Width           =   5430
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   5160
         Picture         =   "FrmLicencia.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   36
         ToolTipText     =   "Cerrar"
         Top             =   75
         Width           =   195
      End
      Begin VB.CommandButton CmdEditor 
         Caption         =   "&Grabar"
         Height          =   420
         Index           =   0
         Left            =   1530
         TabIndex        =   18
         Top             =   3375
         Width           =   1020
      End
      Begin VB.CommandButton CmdEditor 
         Caption         =   "&Cancelar"
         Height          =   420
         Index           =   1
         Left            =   2850
         TabIndex        =   20
         Top             =   3375
         Width           =   1020
      End
      Begin VB.TextBox txt 
         Height          =   870
         Index           =   1
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "FrmLicencia.frx":02EC
         Top             =   2340
         Width           =   5235
      End
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   3
         Left            =   4695
         Picture         =   "FrmLicencia.frx":02F5
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1755
         Width           =   255
      End
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   2
         Left            =   4695
         Picture         =   "FrmLicencia.frx":0427
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1410
         Width           =   255
      End
      Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
         Height          =   315
         Index           =   0
         Left            =   810
         TabIndex        =   10
         Top             =   375
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "  /  /    "
      End
      Begin VB.CommandButton cb 
         Height          =   225
         Index           =   1
         Left            =   1335
         Picture         =   "FrmLicencia.frx":0559
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Seleccione la Licencia"
         Top             =   1080
         Width           =   210
      End
      Begin VB.CommandButton cb 
         Height          =   225
         Index           =   0
         Left            =   1335
         Picture         =   "FrmLicencia.frx":068B
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Seleccione el ersonal"
         Top             =   750
         Width           =   210
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   0
         Left            =   810
         MaxLength       =   20
         TabIndex        =   11
         Text            =   "txt_cb(0)"
         Top             =   720
         Width           =   765
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   1
         Left            =   810
         MaxLength       =   20
         TabIndex        =   12
         Text            =   "txt_cb(1)"
         ToolTipText     =   "Ingrese el Sexo (1:Masculino, 2:Femenino)"
         Top             =   1050
         Width           =   765
      End
      Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
         Height          =   315
         Index           =   1
         Left            =   810
         TabIndex        =   13
         Top             =   1380
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "  /  /    "
      End
      Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
         Height          =   315
         Index           =   2
         Left            =   810
         TabIndex        =   15
         Top             =   1725
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "  /  /    "
      End
      Begin MSComCtl2.DTPicker dtpk 
         Height          =   300
         Index           =   0
         Left            =   3540
         TabIndex        =   14
         Top             =   1380
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         Format          =   49545218
         CurrentDate     =   39534
      End
      Begin MSComCtl2.DTPicker dtpk 
         Height          =   300
         Index           =   1
         Left            =   3540
         TabIndex        =   16
         Top             =   1725
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         Format          =   49545218
         CurrentDate     =   39534
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Información Adicional"
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   28
         Top             =   2115
         Width           =   1515
      End
      Begin VB.Label lbltxtfch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.Fin"
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   27
         Top             =   1845
         Width           =   345
      End
      Begin VB.Label lbldtpk 
         AutoSize        =   -1  'True
         Caption         =   "H.Fin"
         Height          =   195
         Index           =   1
         Left            =   2670
         TabIndex        =   26
         Top             =   1845
         Width           =   375
      End
      Begin VB.Label lbldtpk 
         AutoSize        =   -1  'True
         Caption         =   "H.Inicio"
         Height          =   195
         Index           =   0
         Left            =   2670
         TabIndex        =   25
         Top             =   1500
         Width           =   540
      End
      Begin VB.Label lbltxtfch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.Inicio"
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   22
         Top             =   1500
         Width           =   510
      End
      Begin VB.Label lbl_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   19
         Top             =   1185
         Width           =   480
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
         Left            =   4005
         TabIndex        =   9
         Top             =   1050
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personal"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   6
         Top             =   840
         Width           =   615
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
         Left            =   4005
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbltxtfch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emisión"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   3
         Top             =   495
         Width           =   540
      End
      Begin VB.Label LblTituloFrame 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Editor de Licencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   1560
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5415
         X2              =   5415
         Y1              =   -75
         Y2              =   4815
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -330
         X2              =   5715
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   5685
         Y1              =   3855
         Y2              =   3870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         Index           =   2
         X1              =   90
         X2              =   5280
         Y1              =   3285
         Y2              =   3285
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   300
         Index           =   1
         Left            =   30
         Top             =   45
         Width           =   5355
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
         Left            =   1575
         TabIndex        =   21
         Top             =   1050
         Width           =   3765
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
         Left            =   1575
         TabIndex        =   7
         Top             =   720
         Width           =   3765
      End
   End
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLicencia.frx":07BD
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLicencia.frx":0D01
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLicencia.frx":1093
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLicencia.frx":1217
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLicencia.frx":166B
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLicencia.frx":1783
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLicencia.frx":1CC7
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLicencia.frx":220B
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLicencia.frx":231F
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLicencia.frx":2433
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLicencia.frx":2887
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLicencia.frx":29F3
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLicencia.frx":2F3B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
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
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Periodo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Actualizar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
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
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   5565
      Left            =   30
      TabIndex        =   35
      Top             =   810
      Width           =   11625
      _cx             =   20505
      _cy             =   9816
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
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmLicencia.frx":32CD
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
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "&Agregar"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_4 
         Caption         =   "&Eliminar"
      End
   End
End
Attribute VB_Name = "FrmLicencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--Add Licencia
            nuevo
        Case 1 '--Modificar Licencia
            Modificar
        Case 2 '--Eliminar Licencia
            Eliminar
    End Select
End Sub

Private Sub CmdEditor_Click(Index As Integer)
    Select Case Index
        Case 0 'grabar
            If Grabar() = True Then
                pCargarGrid
                If QueHace = 1 Then
                    If MsgBox("Desea Agregar otro Registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbYes Then
                        nuevo
                    Else
                        CmdEditor_Click 1
                    End If
                Else
                    CmdEditor_Click 1
                End If
            End If
        Case 1 'cancelar
            Cancelar
    End Select
End Sub

Private Sub dtpk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If dtpk(Index).Enabled = False Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index + 2
    ElseIf KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub pCargarGrid()
    Dim nSQL  As String
    Dim RstTmp As New ADODB.Recordset
    On Error GoTo error
    lblperiodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)

    nSQL = "SELECT pla_Licencia.*, pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_tipoLicencia.descripcion AS Licencia " _
        + vbCr + " FROM pla_empleados RIGHT JOIN (mae_tipoLicencia RIGHT JOIN pla_Licencia ON mae_tipoLicencia.id = pla_Licencia.idlic) ON pla_empleados.id = pla_Licencia.idemp " _
        + vbCr + " WHERE (((Month([fchemi])) = " & xMes & ")) " _
        + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"

    Me.MousePointer = vbHourglass
    RST_Busq RstTmp, nSQL, xCon
    '---------------
    pConfigurarGrilla
    '---------------
    If RstTmp.RecordCount <> 0 Then
        Agregando = True
        If RstTmp.RecordCount <> 0 Then
            RstTmp.MoveFirst
            Do While Not RstTmp.EOF
                With Fg1
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = NulosN(RstTmp.Fields("id"))
                    .TextMatrix(.Rows - 1, 2) = NulosN(RstTmp.Fields("idemp"))
                    .TextMatrix(.Rows - 1, 3) = NulosN(RstTmp.Fields("idlic"))
                    
                    .TextMatrix(.Rows - 1, 4) = NulosC(RstTmp.Fields("fchemi"))
                    .TextMatrix(.Rows - 1, 5) = NulosC(RstTmp.Fields("nombres"))
                    .TextMatrix(.Rows - 1, 6) = NulosC(RstTmp.Fields("fchini"))
                    .TextMatrix(.Rows - 1, 7) = NulosC(RstTmp.Fields("horini"))
                    .TextMatrix(.Rows - 1, 8) = NulosC(RstTmp.Fields("fchfin"))
                    .TextMatrix(.Rows - 1, 9) = NulosC(RstTmp.Fields("horfin"))
                    .TextMatrix(.Rows - 1, 10) = NulosC(RstTmp.Fields("Licencia"))
                    .TextMatrix(.Rows - 1, 11) = NulosC(RstTmp.Fields("observacion"))
                    RstTmp.MoveNext
                End With
            Loop
        End If
    End If
    If Fg1.Rows > 1 Then
        Fg1.Row = Fg1.Rows - 1
        If Fg1.Enabled = True Then Fg1.SetFocus
    End If
    '---------------
    Me.MousePointer = vbDefault
    Exit Sub
error:
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

Private Sub Fg1_DblClick()
    cmd_Click 1
End Sub

Private Sub Fg1_EnterCell()
    Fg1.Editable = flexEDNone
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Fg1.Enabled = False Then Exit Sub
    If KeyCode = 45 Then
        nuevo
    End If
    If KeyCode = 46 Then
        Eliminar
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Fg1.Enabled = False Then Exit Sub
    If Button = 2 Then
        PopupMenu Menu1
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    SeEjecuto = False
    pCargarGrid
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If FraEditor.Visible = True Then CmdEditor_Click 1
    End If
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    Agregando = False
    QueHace = 3
   '--
    txtfecha(0).Valor = Date
    txtfecha(1).Valor = Date
    txtfecha(2).Valor = Date
    '--
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation, xTitulo
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub Menu1_1_Click()
    cmd_Click 0
End Sub

Private Sub menu1_2_Click()
cmd_Click 1
End Sub

Private Sub Menu1_3_Click()
    cmd_Click 4
End Sub

Private Sub menu1_4_Click()
    cmd_Click 2
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then CambiarMes
    If Button.Index = 3 Then Buscar
    If Button.Index = 4 Then pCargarGrid
    If Button.Index = 6 Then pExportarExcel
    If Button.Index = 7 Then pImprimir
    If Button.Index = 9 Then
        Unload Me
    End If
End Sub

Sub Eliminar()
    On Error GoTo error
    If Fg1.Rows <= 1 Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If
    If Fg1.Row < 1 Then
        MsgBox "Seleccione correctamente el registro", vbExclamation, xTitulo
        Exit Sub
    End If

    If MsgBox("¿Esta seguro de eliminar el registro?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
        '--eliminar asistencia
        pEliminarAsistencia NulosN(Fg1.TextMatrix(Fg1.Row, 1))
        '---------------------------------------------------------
        xCon.Execute "DELETe * FROM pla_Licencia WHERE id = " & NulosN(Fg1.TextMatrix(Fg1.Row, 1)) & "; "
        
        pCargarGrid
        MsgBox "El registro fue eliminado con éxito", vbInformation, xTitulo
    End If
Exit Sub
error:
    SHOW_ERROR Me.Name, "Eliminar"
End Sub

Private Sub Cancelar()
    QueHace = 3
    pHabilitarBotonEditor False
    If Fg1.Rows = 1 Then
        cmd(0).SetFocus
    Else
        Fg1.Row = 1
        Fg1.SetFocus
    End If
    
End Sub

Private Sub CambiarMes()
    xMes = SeleccionaMes(xCon)
    pCargarGrid
End Sub
Private Sub Modificar()
   '------
    If Fg1.Rows = 1 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If
    If Fg1.Row < 1 Then
        MsgBox "Seleccione correctamente el registro", vbExclamation, xTitulo
        Exit Sub
    End If
    
    QueHace = 2
    pHabilitarBotonEditor True
    pPonerDatos
    LblTituloFrame.Caption = "Modificar Licencia"
    txtfecha(0).SetFocus
    
End Sub

Private Sub Blanquea()
    LimpiaText txt
    LimpiaText txt_cb
    LimpiaText txtfecha
End Sub

Private Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub nuevo()
    QueHace = 1
    Blanquea
    pHabilitarBotonEditor True
    txtfecha(0).Valor = Date
    dtpk(0).Value = CDate("00:00:00")
    dtpk(1).Value = CDate("00:00:00")
    LblTituloFrame.Caption = "Agregar Licencia"
    txtfecha(0).SetFocus
End Sub


Function Grabar() As Boolean
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el Registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo salir
    
    Dim RstCab As New ADODB.Recordset
    Dim RstHora As New ADODB.Recordset
    Dim xCod&, xCol&, xFil&
    Dim nSQL As String
    
    On Error GoTo LaCague
    Me.MousePointer = vbHourglass
    '--iniciando transaccion
    xCon.BeginTrans
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM pla_Licencia ", xCon
        xCod = HallaCodigoTabla("pla_Licencia", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xCod
    Else
        xCod = NulosN(Fg1.TextMatrix(Fg1.Row, 1))
        
        RST_Busq RstCab, "SELECT * FROM pla_Licencia WHERE id =" & xCod & "", xCon
        '******************************************************************************************************
        '--Eliminando los registros de asistencia
        pEliminarAsistencia xCod
        '******************************************************************************************************

    End If
    
    RstCab("idemp") = NulosN(lbl_cod(0).Caption)
    RstCab("idlic") = NulosN(lbl_cod(1).Caption)
    RstCab("fchemi") = CDate(txtfecha(0).Valor)
    RstCab("fchini") = CDate(txtfecha(1).Valor)
    RstCab("horini") = CDate(dtpk(0).Value)
    RstCab("fchfin") = CDate(txtfecha(2).Valor)
    RstCab("horfin") = CDate(dtpk(1).Value)
    RstCab("observacion") = Trim(txt(1).Text)
    RstCab.Update
    '----
    '******************************************************************************************************
    '--generar los registros de asistencia en automatico
    Dim dFecha As Date
    '----
    For dFecha = CDate(txtfecha(1).Valor) To CDate(txtfecha(2).Valor)
        pMacacionDia dFecha, e_Asist_Licencia, NulosN(lbl_cod(0).Caption)
    Next dFecha
    '******************************************************************************************************
    
    xCon.CommitTrans
    '----
    MsgBox "El registro se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    
    Grabar = True
salir:
    Set RstCab = Nothing:    Set RstHora = Nothing
    Me.MousePointer = vbDefault
    Exit Function
LaCague:
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstHora = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function


Private Function fValidarDatos() As Boolean
    Dim mRow&, QGrid&, mCodigo&
    Dim band&
    '--------------------------------
    'ver si tiene horario no este registrado el horario
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    nSQL = "SELECT mae_horariohora.idhora, mae_tipohora.descripcion, mae_horariohora.hingreso, mae_horariohora.hsalida " _
        + vbCr + " FROM mae_tipohora INNER JOIN (mae_horariohora INNER JOIN mae_horarioemp ON mae_horariohora.idhor = mae_horarioemp.idhor) ON mae_tipohora.id = mae_horariohora.idhora " _
        + vbCr + " Where (((mae_horariohora.idhora) = 1) And ((mae_horarioemp.IdEmp) = " & NulosN(txt_cb(0).Text) & ") And ((mae_horarioemp.vigencia) = -1)) " _
        + vbCr + " ORDER BY mae_tipohora.prioridad;"
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount = 0 Then
        MsgBox "El personal no tiene Horario" + vbCr + "Configure el horario a " & lbl_cb(0).Caption, vbExclamation, xTitulo
        Set RstTmp = Nothing
        Exit Function
    End If
    Set RstTmp = Nothing
    '--------------------------------
    
    band = Validar(txt_cb)
    If band <> -1 Then
        MsgBox "Falta ingresar el Campo " & lbl_capt(band).Caption, vbExclamation, xTitulo
        txt_cb(band).SetFocus
        Exit Function
    End If
    band = validar_Fecha(txtfecha)
    If band <> -1 Then
        MsgBox "Falta ingresar el Campo " & lbltxtfch(band).Caption, vbExclamation, xTitulo
        txtfecha(band).SetFocus
        Exit Function
    End If
    
    band = validar_Fecha(dtpk)
    If band <> -1 Then
        MsgBox "Falta ingresar el Campo " & lbldtpk(band).Caption, vbExclamation, xTitulo
        dtpk(band).SetFocus
        Exit Function
    End If
    '--------------------------------
    If CDate(txtfecha(1).Valor) > CDate(txtfecha(2).Valor) Then
        MsgBox "La fecha de Inicio es superior a la fecha Final" + vbCr + "Modifique los valores para continuar", vbExclamation, xTitulo
        txtfecha(2).SetFocus
        Exit Function
    End If
    
    If (CDate(txtfecha(1).Valor) = CDate(txtfecha(2).Valor)) And (CDate(dtpk(0).Value) > CDate(dtpk(1).Value)) Then
        MsgBox "La Hora de Inicio es superior a la Hora Final" + vbCr + "Modifique los valores para continuar", vbExclamation, xTitulo
        dtpk(1).SetFocus
        Exit Function
    End If
    '--restringir si se programo permisos
    nSQL = "SELECT pla_permiso.id,pla_permiso.fchini,pla_permiso.fchfin,mae_tipopermiso.descripcion FROM pla_permiso LEFT JOIN mae_tipopermiso ON pla_permiso.idper=mae_tipopermiso.id " _
            + vbCr + " WHERE idemp = " & NulosN(txt_cb(0).Text) & " AND " _
            + vbCr + " ((fchini BETWEEN cdate('" & txtfecha(1).Valor & "') AND cdate('" & txtfecha(2).Valor & "')) OR " _
            + vbCr + " (fchfin BETWEEN cdate('" & txtfecha(1).Valor & "')  AND cdate('" & txtfecha(2).Valor & "'))) "
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then
        MsgBox "No se puede " + IIf(QueHace = 1, "agregar", "modificar") + vbCr + _
               "Porque se ha programado un permiso del " & Format(RstTmp.Fields("fchini"), "dd/mm/yy") & " al " & Format(RstTmp.Fields("fchfin"), "dd/mm/yy") + vbCr + _
               "Motivo: " & NulosC(RstTmp.Fields("descripcion")), vbExclamation, xTitulo
        Set RstTmp = Nothing
        Exit Function
    End If
    '--restringir si se programo permisos en el intervalo de fechas
    If QueHace = 1 Then
        nSQL = "SELECT id,fchini,fchfin FROM pla_Licencia " _
                + vbCr + " WHERE idemp = " & NulosN(txt_cb(0).Text) & " AND " _
                + vbCr + " ((fchini BETWEEN cdate('" & txtfecha(1).Valor & "') AND cdate('" & txtfecha(2).Valor & "')) OR " _
                + vbCr + "  (fchfin BETWEEN cdate('" & txtfecha(1).Valor & "') AND cdate('" & txtfecha(2).Valor & "'))) "
    Else
        nSQL = "SELECT id,fchini,fchfin FROM pla_Licencia " _
                + vbCr + " WHERE id <> " & NulosN(Fg1.TextMatrix(Fg1.Row, 1)) & " AND idemp = " & NulosN(txt_cb(0).Text) & " AND " _
                + vbCr + " ((fchini BETWEEN cdate('" & txtfecha(1).Valor & "') AND cdate('" & txtfecha(2).Valor & "')) OR " _
                + vbCr + "  (fchfin BETWEEN cdate('" & txtfecha(1).Valor & "') AND cdate('" & txtfecha(2).Valor & "'))) "
    End If
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then
        MsgBox "No se puede " + IIf(QueHace = 1, "agregar", "modificar") + vbCr + _
               "Porque tiene Licencia del " & Format(RstTmp.Fields("fchini"), "dd/mm/yy") & " al " & Format(RstTmp.Fields("fchfin"), "dd/mm/yy"), vbExclamation, xTitulo
        Set RstTmp = Nothing
        Exit Function
    End If
    '--------------------------------
    fValidarDatos = True
End Function
 
Sub Buscar()
    On Error GoTo error
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
   
    Dim xCampos(7, 4) As String
    
    xCampos(0, 0) = "Emisión":  xCampos(0, 1) = "fchemi":       xCampos(0, 2) = "1000":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "Personal": xCampos(1, 1) = "nombres":      xCampos(1, 2) = "3500":     xCampos(1, 3) = "C"
    xCampos(2, 0) = "Licencia":   xCampos(2, 1) = "Licencia":      xCampos(2, 2) = "1500":     xCampos(2, 3) = "C"
    xCampos(3, 0) = "F.Inicio": xCampos(3, 1) = "fchini":       xCampos(3, 2) = "1000":     xCampos(3, 3) = "F"
    xCampos(4, 0) = "H.Inicio": xCampos(4, 1) = "horini":       xCampos(4, 2) = "1000":     xCampos(4, 3) = "F"
    xCampos(5, 0) = "F.Fin":    xCampos(5, 1) = "fchfin":       xCampos(5, 2) = "1000":     xCampos(5, 3) = "F"
    xCampos(6, 0) = "H.Fin":    xCampos(6, 1) = "horfin":       xCampos(6, 2) = "1200":     xCampos(6, 3) = "F"
        
        
    nSQL = "SELECT pla_Licencia.*, pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_tipoLicencia.descripcion AS Licencia " _
        + vbCr + " FROM pla_empleados RIGHT JOIN (mae_tipoLicencia RIGHT JOIN pla_Licencia ON mae_tipoLicencia.id = pla_Licencia.idlic) ON pla_empleados.id = pla_Licencia.idemp " _
        + vbCr + " WHERE (((Month([fchemi])) = " & xMes & ")) " _
        + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"
        
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Licencias", "nombres", "nombres", Principio
    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir
    '--
    Dim A&
    Fg1.Row = 1
    For A = 1 To Fg1.Rows - 1
        DoEvents
        Fg1.Row = A
        If NulosN(Fg1.TextMatrix(A, 1)) = xRs("id") Then
            Exit For
        End If
    Next A
    '--
salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

Private Sub pImprimir()
    On Error GoTo error
    Me.MousePointer = vbHourglass
    Dim oPrint As New SGI2_funciones.formularios
    oPrint.Imprimir_x_VSFlexGrid Fg1, "Consulta de Licencias", , "Periodo:  " + lblperiodo(0).Caption, False, True
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

Private Sub pExportarExcel()
    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Consulta de Licencias", "Periodo:  " + lblperiodo(0).Caption, "", "Licencias"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
End Sub


Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
End Sub

Private Sub pHabilitarBotonEditor(band As Boolean)
    '--TRUE= MUESTRA LA OPCION PARA SELECCIONAR LA RUTA
    Dim K&
    If band = True Then
        Fg1.Enabled = False
        FraEditor.Top = 1545
        FraEditor.Left = 3285
    Else
        Fg1.Enabled = True
    End If
    FraEditor.Visible = band
    habilitar cmd, Not band
    For K = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(K).Enabled = Not band
    Next K
    
End Sub

'****************

Private Sub cb_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nOrden As String
    Dim nCampoBusca As String
    Dim nSQL As String
    On Error GoTo error
    Select Case Index

        Case 0 '--PERSONAL
            pCargarPersonal
            Exit Sub
            
        Case 1 '--MOTIVO
            ReDim xCampos(1, 3) As String
            xCampos(0, 0) = "Desripción":   xCampos(0, 1) = "nombre":         xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            
            nTitulo = "Buscando Licencias"
            nSQL = "SELECT mae_tipolicencia.id, mae_tipolicencia.descripcion AS nombre, mae_tipolicencia.id AS cod " _
                + vbCr + " FROM mae_tipolicencia;"

        Case 2, 3 '--hora de inicio, hora de fin
            If Index = 2 Or Index = 3 Then
            Dim obj As New SGI2_funciones.formularios
            obj.HoraSeleccionar dtpk(Index - 2), -1, -1, dtpk(Index - 2).Value
            Set obj = Nothing
            Select Case Index
                Case 2 '--HORA INICIO
                    txtfecha(2).SetFocus
                Case 3 '--HORA FIN
                    txt(1).SetFocus
            End Select
            Exit Sub
    End If
    
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
    Select Case Index
        Case 0 '--PERSONAL
            txt_cb(1).SetFocus
        Case 1 '--MOTIVO
            txtfecha(1).SetFocus
    End Select
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
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
   
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then Exit Sub
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index

        Case 0 '--PERSONAL
            nSQL = "SELECT pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.id as cod,pla_empleados.numdoc, mae_dociden.abrev AS tipodoc, mae_sexo.abrev AS sexo, Format([pla_empleados].[fchnac],'dd/mm/yyyy') AS fchnac, pla_empleados.numtel, pla_empleados.email " _
                + vbCr + " FROM mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_sexo.id = pla_empleados.idsex " _
                + vbCr + " WHERE pla_empleados.id = " & NulosN(txt_cb(Index).Text) & ";"
            
        Case 1 '--MOTIVO
            nSQL = "SELECT mae_tipolicencia.id, mae_tipolicencia.descripcion AS nombre, mae_tipolicencia.id AS cod " _
                + vbCr + " FROM mae_tipolicencia " _
                + vbCr + " WHERE mae_tipoLicencia.id = " & NulosN(txt_cb(Index).Text) & ";"
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
    Else
        txt_cb(Index).Text = ""
    End If
    '--------------
    Select Case Index
        Case 0 '--PERSONAL
            If Agregando = False Then txt_cb(1).SetFocus
        Case 1 '--MOTIVO
            If Agregando = False Then txtfecha(1).SetFocus
    End Select
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


'****************

Private Sub pPonerDatos()
    On Error GoTo error
    Dim mRow&
    
    QueHace = 2
    Agregando = True
    With Fg1
        mRow = .Row
        txtfecha(0).Valor = CDate(.TextMatrix(mRow, 4)) '--fecha emision
        If NulosN(.TextMatrix(mRow, 2)) <> 0 Then '--personal
            txt_cb(0).Text = NulosN(.TextMatrix(mRow, 2))
            txt_cb_Validate 0, False
        End If
        If NulosN(.TextMatrix(mRow, 3)) <> 0 Then '--motivo
            txt_cb(1).Text = NulosN(.TextMatrix(mRow, 3))
            txt_cb_Validate 1, False
        End If
        If IsDate(.TextMatrix(mRow, 6)) = True Then '--fecha inicio
            txtfecha(1).Valor = CDate(.TextMatrix(mRow, 6))
        Else
            txtfecha(1).Valor = ""
        End If
        If IsDate(.TextMatrix(mRow, 8)) = True Then '--fecha fin
            txtfecha(2).Valor = CDate(.TextMatrix(mRow, 8))
        Else
            txtfecha(2).Valor = ""
        End If
        If IsDate(.TextMatrix(mRow, 7)) = True Then '--hora inicio
            dtpk(0).Value = CDate(.TextMatrix(mRow, 7))
        Else
            dtpk(0).Value = ""
        End If
        If IsDate(.TextMatrix(mRow, 9)) = True Then '--hora fin
            dtpk(1).Value = CDate(.TextMatrix(mRow, 9))
        Else
            dtpk(1).Value = ""
        End If
        txt(1).Text = .TextMatrix(mRow, 11) '--observacion
    End With
    Agregando = False
    Exit Sub
error:
    Agregando = False
    CmdEditor(0).Enabled = False
    SHOW_ERROR Me.Name, "pPonerDatos"
End Sub

Private Sub pConfigurarGrilla()
    With Fg1
        .Rows = 1
        .Cols = 12
        .FixedRows = 1
        .RowHeight(0) = 250
        
        .TextMatrix(0, 1) = "id":       .ColWidth(1) = 0:     .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(0, 2) = "idemp":    .ColWidth(2) = 0:     .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(0, 3) = "idlic":    .ColWidth(3) = 0:     .ColAlignment(3) = flexAlignLeftCenter
        
        .TextMatrix(0, 4) = "Emisión":  .ColWidth(4) = 800:     .ColAlignment(4) = flexAlignCenterCenter:   .Row = 0: .Col = 4: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 5) = "Personal": .ColWidth(5) = 2200:    .ColAlignment(5) = flexAlignLeftCenter:     .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 6) = "F.Inicio": .ColWidth(6) = 800:     .ColAlignment(6) = flexAlignCenterCenter:   .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 7) = "H.Inicio": .ColWidth(7) = 1050:    .ColAlignment(7) = flexAlignCenterCenter:   .Row = 0: .Col = 7: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 8) = "F.Fin":    .ColWidth(8) = 800:     .ColAlignment(8) = flexAlignCenterCenter:   .Row = 0: .Col = 8: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 9) = "H.Fin":    .ColWidth(9) = 1050:    .ColAlignment(9) = flexAlignCenterCenter:   .Row = 0: .Col = 9: .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(0, 10) = "Motivo":   .ColWidth(10) = 2000:    .ColAlignment(10) = flexAlignLeftCenter:   .Row = 0: .Col = 10: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 11) = "Información Adicional":    .ColWidth(11) = 2700:    .ColAlignment(11) = flexAlignLeftCenter:   .Row = 0: .Col = 11: .CellAlignment = flexAlignLeftCenter
        .ColFormat(4) = FORMAT_DATE
        .ColFormat(6) = FORMAT_DATE
        .ColFormat(7) = FORMAT_HORA_AL_SEGUNDO
        .ColFormat(8) = FORMAT_DATE
        .ColFormat(9) = FORMAT_HORA_AL_SEGUNDO
        .SelectionMode = flexSelectionByRow
    End With
    '*****************************************
    DoEvents
End Sub


Private Sub pEliminarAsistencia(mIdCodigo&)
    '--Eliminando los registros de asistencia
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim mIdEmp&
    '--buscando se hay dias que se registraron las asistencias
    '--idori=3:licencia segun tabla pla_origenes
    nSQL = "SELECT pla_licencia.idemp, pla_marcacion.dia, pla_marcaciondet.idori, pla_marcaciondet.idmarca " _
        + vbCr + " FROM pla_licencia, pla_marcacion INNER JOIN pla_marcaciondet ON pla_marcacion.id = pla_marcaciondet.idmarca " _
        + vbCr + " WHERE (((pla_marcacion.dia) Between [pla_licencia].[fchini] And [pla_licencia].[fchfin]) " & _
                  "AND ((pla_marcaciondet.idori)=3) AND ((pla_licencia.id)=" & mIdCodigo & "));"
                  
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount <> 0 Then
        RstTmp.MoveFirst
        nSQL = ""
        mIdEmp = RstTmp.Fields("idemp")
        Do While Not RstTmp.EOF
            nSQL = nSQL & RstTmp.Fields("idmarca") & ","
            RstTmp.MoveNext
        Loop
        If nSQL <> "" Then nSQL = " (" + Left(nSQL, Len(nSQL) - 1) + ") "
        '--marcacion
        xCon.Execute "DELETE FROM pla_marcaciondet " & _
                     "WHERE idemp = " & mIdEmp & " AND idori=3 AND idmarca In " & nSQL & " ;"
        '--tipos de horas
        xCon.Execute "DELETE FROM pla_marcacionhora " & _
                     "WHERE idemp = " & mIdEmp & " AND idhora IN (7,8) AND idmarca In " & nSQL & " ;"
        '--7:hora licencia con goce haber; 8: hora licencia sin goce haber
    End If
    Set RstTmp = Nothing

End Sub

Private Sub pic_Click()
    CmdEditor_Click 1
End Sub

'--------------
Private Sub pCargarPersonal()
    Dim xRs As New ADODB.Recordset
    pBuscarPersonal xRs, True
    If xRs.State = 1 Then
        txt_cb(0) = xRs.Fields("id") & "" '--TEXTO A MOSTRAR
        lbl_cb(0).Caption = xRs.Fields("nombres") & "" '--NOMBRE
        lbl_cod(0).Caption = xRs.Fields("id") & "" '--CODIGO
        lbl_cb(0).ToolTipText = xRs.Fields("nombres") & "" '--NOMBRE
        txt_cb(1).SetFocus
    End If
    Set xRs = Nothing
End Sub
'--------------

