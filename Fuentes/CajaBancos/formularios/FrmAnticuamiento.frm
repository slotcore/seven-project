VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmAnticuamiento 
   Caption         =   "Caja y Bancos - Anticuamiento"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
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
      Height          =   885
      Left            =   7740
      TabIndex        =   21
      Top             =   360
      Width           =   4035
      Begin VB.CommandButton CmdBusMon 
         Height          =   240
         Left            =   1185
         Picture         =   "FrmAnticuamiento.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   420
         Width           =   210
      End
      Begin VB.TextBox TxtIdMon 
         Height          =   300
         Left            =   720
         MaxLength       =   1
         TabIndex        =   23
         Text            =   "TxtIdMon"
         Top             =   390
         Width           =   705
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
         Height          =   300
         Left            =   1425
         TabIndex        =   25
         Top             =   390
         Width           =   2490
      End
      Begin VB.Label LblTipCam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   24
         Top             =   480
         Width           =   585
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[  Seleccionar  ]"
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
      Height          =   885
      Left            =   2407
      TabIndex        =   14
      Top             =   360
      Width           =   5280
      Begin VB.CommandButton CmdBusCliPro 
         Enabled         =   0   'False
         Height          =   240
         Left            =   4920
         Picture         =   "FrmAnticuamiento.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   480
         Width           =   210
      End
      Begin VB.OptionButton OptSel2 
         Caption         =   "Seleccionar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1110
         TabIndex        =   16
         Top             =   240
         Width           =   1140
      End
      Begin VB.OptionButton OptSel1 
         Caption         =   "Todos"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.TextBox TxtCliPro 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "TxtCliPro"
         Top             =   450
         Width           =   4995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   2580
         TabIndex        =   20
         Top             =   210
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label LblIdCliPro 
         Caption         =   "LblIdCliPro"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3090
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   885
      Left            =   60
      TabIndex        =   10
      Top             =   360
      Width           =   2295
      Begin VB.OptionButton opt4ta 
         Caption         =   "Prestador de Servicio"
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
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   630
         Width           =   2160
      End
      Begin VB.OptionButton OptProvee 
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   405
         Width           =   1230
      End
      Begin VB.OptionButton OptCliente 
         Caption         =   "Cliente"
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
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   180
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.Frame FraDetalle 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6270
      Left            =   11970
      TabIndex        =   5
      Top             =   1260
      Visible         =   0   'False
      Width           =   11775
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   11490
         Picture         =   "FrmAnticuamiento.frx":0264
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   6
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg1 
         Height          =   5490
         Left            =   60
         TabIndex        =   8
         Top             =   720
         Width           =   11655
         _cx             =   20558
         _cy             =   9684
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
         BackColor       =   14745342
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   128
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14745342
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
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmAnticuamiento.frx":0550
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   345
         Left            =   60
         TabIndex        =   9
         Top             =   360
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   609
         ButtonWidth     =   609
         ButtonHeight    =   556
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Consultar"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Exportar a MSExcel"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Salir"
               ImageIndex      =   13
            EndProperty
         EndProperty
         BorderStyle     =   1
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   3285
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483637
            ImageWidth      =   16
            ImageHeight     =   15
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   14
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":06D7
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":0C1B
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":0FAD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":1107
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":1499
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":161D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":1A71
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":1B89
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":20CD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":2611
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":2725
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":2839
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":2C8D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAnticuamiento.frx":2DF9
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   0
         Y1              =   0
         Y2              =   6420
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000E&
         BorderWidth     =   2
         Index           =   3
         X1              =   15
         X2              =   11790
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   2
         X1              =   15
         X2              =   11790
         Y1              =   6255
         Y2              =   6255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   11760
         X2              =   11760
         Y1              =   15
         Y2              =   6390
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
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
         Left            =   135
         TabIndex        =   7
         Top             =   75
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   45
         Top             =   45
         Width           =   11685
      End
   End
   Begin VB.Frame fraBarra 
      BorderStyle     =   0  'None
      Caption         =   "FrmConsultaDiario"
      Height          =   780
      Left            =   12120
      TabIndex        =   1
      Top             =   6750
      Visible         =   0   'False
      Width           =   6285
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   150
         TabIndex        =   2
         Top             =   315
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   -15
         Y2              =   900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   6270
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   1
         X1              =   -75
         X2              =   6500
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   6270
         X2              =   6270
         Y1              =   -30
         Y2              =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Documentos"
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
         Height          =   195
         Left            =   195
         TabIndex        =   4
         Top             =   75
         Width           =   2130
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Interrumpir = ESC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   4605
         TabIndex        =   3
         Top             =   75
         Width           =   1530
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3285
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":3341
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":3885
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":3C17
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":3D71
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":4103
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":4287
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":46DB
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":47F3
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":4D37
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":527B
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":538F
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":54A3
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":58F7
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnticuamiento.frx":5A63
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg2 
      Height          =   6240
      Left            =   60
      TabIndex        =   26
      Top             =   1290
      Width           =   11745
      _cx             =   20717
      _cy             =   11007
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
      BackColor       =   14745342
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   128
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14745342
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
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmAnticuamiento.frx":5FAB
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
Attribute VB_Name = "FrmAnticuamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--frmAnticuamiento
'--Creado:      18/05/10 Johan Castro
'--Proposito:   Muestra en reporte el anticuamiento de los saldos por cobrar o pagar agrupados en periodos
'--             Muestra el resumen, seleccionando un registro se prodra ver el detalle
'--
'--Modificado: 15/06/11 Johan Castro
'--Proposito:   Dar flexibilidad al formulario para definir el tamaño del mismo


Option Explicit
Dim SeEjecuto As Boolean
Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
Dim OrigFX As Long '--para mover el frame posicion horizontal
Dim OrigFY As Long '--para mover el frame posicion vertical


Private Sub CmdBusCliPro_Click()
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    If OptCliente.Value = True Then
        xForm.Titulo = "Buscando Clientes"
        xForm.SQLCad = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id From mae_cliente ORDER BY mae_cliente.nombre"
        xCampos(0, 0) = "Cliente":       xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    ElseIf OptProvee.Value = True Then
        xForm.Titulo = "Buscando Proveedores"
        xForm.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov ORDER BY mae_prov.nombre"
        xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    ElseIf opt4ta.Value = True Then
        xForm.Titulo = "Buscando Prestador de Servicio"
        xForm.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov WHERE mae_prov.tipper = 1 ORDER BY mae_prov.nombre"
        xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
        
    End If

    xCampos(1, 0) = "Nº R.U.C.":     xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"

    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "nombre"
    xForm.CampoBusca = "nombre"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCliPro.Text = xRs("nombre")
        LblIdCliPro.Caption = xRs("id")
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Fg2_DblClick()
    If Fg2.Row >= Fg2.FixedRows And Fg2.Row < Fg2.Rows - 1 Then
        BAND_INTERRUMPIR = False
        Bloquea True
        CargarDetalle Fg2.TextMatrix(Fg2.Row, 1)
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        pConfigurarGrilla
        TxtIdMon.Text = 1
        TxtIdMon_Validate False
        SeEjecuto = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    ElseIf KeyCode = vbKeyF3 And Shift = 0 Then
        
    ElseIf KeyCode = vbKeyF8 Then
        pConsultar
    End If

End Sub

Private Sub Form_Load()
    SeEjecuto = False
    TxtCliPro.Text = ""
    LblMoneda.Caption = ""
    TxtIdMon.Text = ""
    Fg1.AutoSearch = flexSearchFromTop
    Fg2.AutoSearch = flexSearchFromTop
    
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
  
    If Me.Height > 3000 Then
        Fg2.Top = 1290
        Fg2.Width = Me.Width - 150
        Fg2.Height = Me.Height - 1700
    End If
End Sub

Private Sub opt4ta_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    Label1(0).Caption = "Prestador de Servicio"
End Sub

Private Sub OptSel1_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    TxtCliPro.Enabled = False
    CmdBusCliPro.Enabled = False
End Sub

Private Sub OptSel2_Click()
    TxtCliPro.Enabled = True
    CmdBusCliPro.Enabled = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 3 Then pExportar
    If Button.Index = 4 Then pImprimir
    If Button.Index = 6 Then
        Unload Me
    End If
End Sub

Sub pExportar()
    '===================================================================================================
    'Creado :  20/05/10 Por: Johan Castro
    'Propósito: Exportar el listado al MSExcel
    '
    'Entradas:  Ninguna
    '
    'Resultados:Nuevo archivo de MSExcel
    
    'Nota :Exportar datos del resumen y del detalle
    '
    '===================================================================================================

    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios
    Dim nPeriodo As String
    Dim nTitulo As String
    Dim nTitulo1 As String
    
    If AnoTra = Year(Date) Then
        nPeriodo = "Al  " + Format(Date, "dd/mm/yyyy")
    Else
        nPeriodo = "Al  31/12/" + AnoTra
    End If
    
    If OptCliente.Value = True Then
        nTitulo = "Anticuamiento de Clientes (Expresado en " & LblMoneda.Caption & ")"
    ElseIf OptProvee.Value = True Then
        nTitulo = "Anticuamiento de Proveedores (Expresado en " & LblMoneda.Caption & ")"
    ElseIf opt4ta.Value = True Then
        nTitulo = "Anticuamiento de Prestadores de Servicio (Expresado en " & LblMoneda.Caption & ")"
    End If
    
    '--datos del cliente, proveedor o prestador de servicio
    nTitulo1 = ""
        
    If FraDetalle.Visible = True Then '--detalle
        '--datos del cliente, proveedor o prestador de servicio
        nTitulo1 = LblTitulo.Caption
        
        oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, nTitulo, nPeriodo, nTitulo1, "Anticuamiento"
    Else
        oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg2, nTitulo, nPeriodo, nTitulo1, "Anticuamiento"
    End If
    
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportar"
End Sub



Private Sub TxtCliPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCliPro_Click
    End If
End Sub

'***********************************************************************************************
'------------CAMBIOS AL 020108

Private Sub CmdBusMon_Click()
    
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":      xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":   xCampos(1, 1) = "id":              xCampos(1, 2) = "500":      xCampos(1, 3) = "N"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_moneda ORDER BY descripcion ;", xCampos(), "Buscando Moneda", "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    TxtIdMon.Text = xRs("id") & ""
    LblMoneda.Caption = xRs("descripcion") & ""
    
Salir:
    Set xRs = Nothing
End Sub


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

Private Sub pConfigurarGrilla()
    Dim xrst As New ADODB.Recordset
    Dim nSQL As String '--Sentencia SQL para la consulta
    Dim A As Integer
    
    nSQL = "SELECT mae_rangos.id, mae_rangos.descripcion  FROM mae_rangos WHERE (((mae_rangos.id)<>0)) ORDER BY mae_rangos.id;"
    RST_Busq xrst, nSQL, xCon
    
    
    '--configurar el detalle
    With Fg1
        '-----
        .Rows = 2
        .FixedRows = 2
        .FrozenCols = 0
        .ColWidth(0) = 200
        .RowHeight(0) = 250
        .RowHeight(1) = 250
        .Cols = 11
        .WordWrap = True
        
        GRID_COMBINAR Fg1, 0, 1, 0, 10, "DATOS DEL DOCUMENTO", flexAlignCenterCenter, True, flexMergeFree, , &HC0C0C0, True

        .TextMatrix(1, 1) = "Nº Reg.":  .ColWidth(1) = 900:    .ColAlignment(1) = flexAlignCenterCenter:     .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 2) = "T.D.":         .ColWidth(2) = 450:    .ColAlignment(2) = flexAlignLeftCenter:     .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 3) = "N°.Documento": .ColWidth(3) = 1400:   .ColAlignment(3) = flexAlignLeftCenter:     .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 4) = "Fch.Emi.":     .ColWidth(4) = 840:    .ColAlignment(4) = flexAlignLeftCenter:     .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 5) = "Fch.Ven.":     .ColWidth(5) = 840:    .ColAlignment(5) = flexAlignLeftCenter:     .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 6) = "Dif Dias":     .ColWidth(6) = 450:    .ColAlignment(6) = flexAlignRightCenter:     .Row = 1: .Col = 6: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 7) = "Cond. Pago":   .ColWidth(7) = 800:   .ColAlignment(7) = flexAlignLeftCenter:     .Row = 1: .Col = 7: .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(1, 8) = "M":            .ColWidth(8) = 450:    .ColAlignment(8) = flexAlignLeftCenter:     .Row = 1: .Col = 8: .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(1, 9) = "T.C.":         .ColWidth(9) = 500:    .ColAlignment(9) = flexAlignRightCenter:    .Row = 1: .Col = 9: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 10) = "Imp":         .ColWidth(10) = 900:   .ColAlignment(10) = flexAlignRightCenter:   .Row = 1: .Col = 10: .CellAlignment = flexAlignRightCenter
        
        .FrozenCols = 10
        
        '--formato a la cabecera
        .Cell(flexcpForeColor, 1, 1, 1, 10) = &H800000
        .Cell(flexcpFontBold, 1, 1, 1, 10) = True
                
        '----------------------
        '--colocar las cabeceras
        xrst.MoveFirst
        Do While Not xrst.EOF
            .Cols = .Cols + 1

            GRID_COMBINAR Fg1, 0, .Cols - 1, 1, .Cols - 1, NulosC(xrst("descripcion")), flexAlignRightCenter, False, flexMergeFree, , &HC0C0C0, True
            .ColWidth(.Cols - 1) = 900
            xrst.MoveNext
        Loop
        
        .Cols = .Cols + 1
        GRID_COMBINAR Fg1, 0, .Cols - 1, 1, .Cols - 1, "Total", flexAlignRightCenter, False, flexMergeFree, , &HC0C0C0, True
        
        '----------------------
        '--hacer que no se formen grupos de celdas con datos iguales
        .MergeCells = flexMergeFixedOnly
        
        .SelectionMode = flexSelectionByRow
    End With
    
    With Fg2
        '-----
        .Clear
        .Rows = 1
        .Cols = 5
        .FixedRows = 1
        .FrozenCols = 0
        
        .ColWidth(0) = 200:
        .RowHeight(0) = 500
        
        .WordWrap = True
        
        .TextMatrix(0, 1) = "Id":       .ColWidth(1) = 0
        .TextMatrix(0, 2) = "R.U.C.":   .ColWidth(2) = 1200:  .ColAlignment(2) = flexAlignCenterCenter: .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Nombres":  .ColWidth(3) = 3500:  .ColAlignment(3) = flexAlignLeftCenter:   .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 4) = "Cant. Doc":  .ColWidth(4) = 600:  .ColAlignment(4) = flexAlignRightCenter:   .Row = 0: .Col = 4: .CellAlignment = flexAlignCenterCenter
        '--formato a la cabecera
        .Cell(flexcpForeColor, 0, 1, 0, 4) = &H800000
        .Cell(flexcpFontBold, 0, 1, 0, 4) = True
        
        .FrozenCols = 4
        '----------------------
        '--colocar las cabeceras
        xrst.MoveFirst
        Do While Not xrst.EOF
            .Cols = .Cols + 1

            GRID_COMBINAR Fg2, 0, .Cols - 1, 0, .Cols - 1, NulosC(xrst("descripcion")), flexAlignRightCenter, False, flexMergeFree, , &HC0C0C0, True
            .ColWidth(.Cols - 1) = 1000
            
            xrst.MoveNext
        Loop
        
        .Cols = .Cols + 1
        GRID_COMBINAR Fg2, 0, .Cols - 1, 0, .Cols - 1, "Total", flexAlignRightCenter, False, flexMergeFree, , &HC0C0C0, True
        
        '----------------------
        '--hacer que no se formen grupos de celdas con datos iguales
        .MergeCells = flexMergeFixedOnly
        '----------------------
        .SelectionMode = flexSelectionByRow
    End With
    
    DoEvents
End Sub


Private Sub pImprimir()
    '===================================================================================================
    'Creado :  20/05/10 Por: Johan Castro
    'Propósito: Imprimir listado
    '
    'Entradas:  Ninguna
    '
    'Resultados:Vista previa de la impresion
    
    'Nota :Imprime datos del resumen y del detalle
    '
    '===================================================================================================

    On Error GoTo error
    
    Dim oPrint As New SGI2_funciones.formularios
    Dim nPeriodo As String
    Dim nTitulo As String
    Dim nTitulo1 As String
    
    If AnoTra = Year(Date) Then
        nPeriodo = "Al  " + Format(Date, "dd/mm/yyyy")
    Else
        nPeriodo = "Al  31/12/" + AnoTra
    End If
    
    If OptCliente.Value = True Then
        nTitulo = "Anticuamiento de Clientes (Expresado en " & LblMoneda.Caption & ")"
    ElseIf OptProvee.Value = True Then
        nTitulo = "Anticuamiento de Proveedores (Expresado en " & LblMoneda.Caption & ")"
    ElseIf opt4ta.Value = True Then
        nTitulo = "Anticuamiento de Prestadores de Servicio (Expresado en " & LblMoneda.Caption & ")"
    End If
    
    '--datos del cliente, proveedor o prestador de servicio
    nTitulo1 = LblTitulo.Caption


    If FraDetalle.Visible = True Then
        oPrint.Imprimir_x_VSFlexGrid Fg1, nTitulo, nTitulo1, nPeriodo, True, True
    Else
        oPrint.Imprimir_x_VSFlexGrid Fg2, nTitulo, nTitulo1, nPeriodo, True, True
    End If
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

Private Sub pConsultar()
    BAND_INTERRUMPIR = False

    CargarResumen
End Sub



Sub CargarDetalle(IdPer As Long)
    '===================================================================================================
    'Creado :  17/05/10 Por: Johan Castro
    'Propósito: Muestra el listado del detalle de documentos por cliente,proveedor o Prestador de Servicio
    '           toma como base de calculo el saldo por pagar o cobrar
    '
    'Entradas:  IdPer = Codigo de cliente, proveedor o prestador de servicio
    '
    'Resultados:Reporte segun parametro ingresado
    '
    'Modificado: 11/05/11 Johan Castro
    '            Invocar a evento GenerarConsulta(IdPer) para generar la sentencia SQL

    '===================================================================================================
    
    Dim rst As New ADODB.Recordset
    Dim A, xFila As Long
    Dim nSQL As String
    Dim nCampoMuestra As String '--indica el campo que se mostrara esta en funcion de la moneda seleccionada
    Dim nSQLSub As String '--Sentencia SQL para identificar una subconsulta; está a nivel de detalle

    Dim xCol As Integer '--posicion de la columna
    
    
    On Error GoTo error

    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Seleccione una Moneda", vbExclamation, xTitulo
        TxtIdMon.SetFocus
        Bloquea False
        Exit Sub
    End If
    
    
    '--------------------------
    '--muestra el detalle
    FraDetalle.Top = 1320
    FraDetalle.Left = 30
    
    '--barra de progreso
    fraBarra.Visible = True
    fraBarra.Left = 2798
    fraBarra.Top = 2925
    ProgressBar1.Max = 1
    ProgressBar1.Value = 0
    fraBarra.Refresh
    
    '--limpiar el grid
    Fg1.Rows = Fg1.FixedRows
    
    '--mostrar todas las columnas, luego se ocultaran si no tiene importes
    For xCol = 11 To Fg1.Cols - 1
        Fg1.ColWidth(xCol) = 900
    Next xCol
    
    '--filtro para la presentacion expresada en la moneda segun seleccion
    If NulosN(TxtIdMon.Text) = 1 Then
        nCampoMuestra = "Sum(VistaDet.impsalexpmn)"
    Else
        nCampoMuestra = "Sum(VistaDet.impsalexpme)"
    End If
        
    '--colocando datos del cliente
    LblTitulo.Caption = "Nº R.U.C. : " & Fg2.TextMatrix(Fg2.Row, 2) & "     " & Fg2.TextMatrix(Fg2.Row, 3)
    
    
    nSQLSub = GenerarConsulta(IdPer)

    
    DoEvents

    nSQL = "TRANSFORM " & nCampoMuestra & " AS SumaDeimpsalexpmn " _
        + vbCr + " SELECT VistaDet.registro, VistaDet.numruc, VistaDet.nombre, VistaDet.abrev, VistaDet.numerodoc, VistaDet.fchdoc, VistaDet.fchven, VistaDet.numdias, VistaDet.condpago, VistaDet.simbolo, " _
        + vbCr + " VistaDet.tipcam, VistaDet.impreal, " & nCampoMuestra & " AS Total " _
        + vbCr + " FROM mae_rangos, " _
        + vbCr + " ( " & nSQLSub & " ) AS VistaDet  " _
        + vbCr + " WHERE (((VistaDet.numdias) Between [mae_rangos].[rangoini] And [mae_rangos].[rangofin])) " _
        + vbCr + " GROUP BY VistaDet.registro, VistaDet.numruc, VistaDet.nombre, VistaDet.abrev, VistaDet.numerodoc, VistaDet.fchdoc, VistaDet.fchven,VistaDet.numdias, VistaDet.condpago, VistaDet.simbolo, VistaDet.tipcam, VistaDet.impreal " _
        + vbCr + " ORDER BY VistaDet.nombre,VistaDet.abrev, VistaDet.numerodoc " _
        + vbCr + " PIVOT mae_rangos.descripcion; "

    '--ejecutar la conulta
    RST_Busq rst, nSQL, xCon
    
    If rst.RecordCount = 0 Then
        MsgBox "No hay registros para mostrar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fraBarra.Visible = False
        Set rst = Nothing
        Exit Sub
    End If
    
    ProgressBar1.Max = rst.RecordCount
    
    Me.MousePointer = vbHourglass
    
    If rst.RecordCount <> 0 Then
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo Salir:

        rst.MoveFirst
            
        Dim mRowIni As Integer
        
        '--------------
        rst.MoveFirst
        For A = 1 To rst.RecordCount    '--GRUPO DE CLIENTE/PROVEEDOR
            DoEvents
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo Salir:
            ProgressBar1.Value = A
            
            Fg1.Rows = Fg1.Rows + 1
            xFila = Fg1.Rows - 1
            
            Fg1.TextMatrix(xFila, 1) = NulosC(rst("registro"))
            Fg1.TextMatrix(xFila, 2) = NulosC(rst("abrev"))
            Fg1.TextMatrix(xFila, 3) = NulosC(rst("numerodoc"))
            Fg1.TextMatrix(xFila, 4) = Format(rst("fchdoc"), FORMAT_DATE)
            Fg1.TextMatrix(xFila, 5) = Format(rst("fchven"), FORMAT_DATE)
            Fg1.TextMatrix(xFila, 6) = NulosN(rst("numdias"))
            Fg1.TextMatrix(xFila, 7) = NulosC(rst("condpago"))
            Fg1.TextMatrix(xFila, 8) = NulosC(rst("simbolo"))
            Fg1.TextMatrix(xFila, 9) = Format(NulosN(rst("tipcam")), "###0.##0") & ""
            Fg1.TextMatrix(xFila, 10) = Format(NulosN(rst("impreal")), FORMAT_MONTO)
            
            '--colocar los importes de la tabla cruzada
            For xCol = 11 To Fg1.Cols - 1
                
                If RstRegistroBuscaCampo(rst, NulosC(Fg1.TextMatrix(1, xCol))) = True Then
                    Fg1.TextMatrix(xFila, xCol) = Format(NulosN(rst(Fg1.TextMatrix(1, xCol))), FORMAT_MONTO)
                    
                End If
            Next xCol
            '-------------------------------------------------------------
            
            rst.MoveNext
            
            If rst.EOF = True Then Exit For
            

        Next A
        '-------------------------------------------------------------
        
        Fg1.Rows = Fg1.Rows + 1
        xFila = xFila + 1
        
        FORMATO_CELDA Fg1, xFila, 10, , True, , "TOTAL -->"
        '--colocar los importes de la tabla cruzada
        For xCol = 11 To Fg1.Cols - 1
            
            FORMATO_CELDA Fg1, xFila, xCol, , True, , Format(GRID_SUMAR_COL(Fg1, xCol, Fg1.FixedRows, Fg1.Rows - 2), FORMAT_MONTO)
            
            If NulosN(Fg1.TextMatrix(xFila, xCol)) <> 0 Then
                Fg1.ColWidth(xCol) = 900
            Else
                Fg1.ColWidth(xCol) = 0
            End If
            
        Next xCol
        
        
        '-------------------------------------------------------------
    End If
    
    Set rst = Nothing
    fraBarra.Visible = False
    Me.MousePointer = vbDefault
'    MsgBox "La Consulta fue se realizó Correctamente", vbInformation, xTitulo
    Exit Sub
Salir:
    Me.MousePointer = vbDefault
    fraBarra.Visible = False
    Set rst = Nothing
    MsgBox "La Consulta fue Interrumpida", vbInformation, xTitulo
    Exit Sub
error:
    
    Me.MousePointer = vbDefault
'    Resume
    fraBarra.Visible = False
    Set rst = Nothing
    SHOW_ERROR Me.Name, "CargarCli2"
    
End Sub

Sub CargarResumen()
    '===================================================================================================
    'Creado :  19/05/10 Por: Johan Castro
    'Propósito: Muestra listado el resumen del anticuamiento por cliente,proveedor o Prestador de Servicio
    '           toma como base de calculo el saldo por pagar o cobrar
    '
    'Entradas:  Ninguna
    '
    'Resultados:Reporte segun parametros seleccionados
    
    'Nota :Permite ver el detalle haciendo doble clic en el registro
    'Modificado: 11/05/11 Johan Castro
    '            Invocar a evento GenerarConsulta() para generar la sentencia SQL
    '===================================================================================================
    
    Dim rst As New ADODB.Recordset
    Dim A, xFila As Long
    Dim nSQL As String
    Dim nCampoMuestra As String '--indica el campo que se mostrara esta en funcion de la moneda seleccionada
                                   '--esten relacionados a sus documentos de referencia
    Dim nSQLSub As String '--Sentencia SQL para identificar una subconsulta; está a nivel de detalle
    
    
    On Error GoTo error
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Seleccione una Moneda", vbExclamation, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    
    '--------------------------
    fraBarra.Visible = True
    fraBarra.Left = 2798
    fraBarra.Top = 2925
        
    ProgressBar1.Max = 1
    ProgressBar1.Value = 0
    fraBarra.Visible = True
    fraBarra.Refresh
    
    Fg2.Rows = Fg2.FixedRows
    
    DoEvents

    '--filtro para la presentacion expresada en la moneda segun seleccion
    If NulosN(TxtIdMon.Text) = 1 Then
        nCampoMuestra = "Sum(VistaDet.impsalexpmn)"
    Else
        nCampoMuestra = "Sum(VistaDet.impsalexpme)"
    End If
    
    nSQLSub = GenerarConsulta()

    nSQL = "TRANSFORM " & nCampoMuestra & " AS SumaDeimpsalexpmn " _
        + vbCr + " SELECT VistaDet.idper,VistaDet.numruc, VistaDet.nombre,count(VistaDet.registro) as candoc, " & nCampoMuestra & " AS Total " _
        + vbCr + " FROM mae_rangos, " _
        + vbCr + " (" & nSQLSub & ") AS VistaDet  " _
        + vbCr + " WHERE (((VistaDet.numdias) Between [mae_rangos].[rangoini] And [mae_rangos].[rangofin])) " _
        + vbCr + " GROUP BY VistaDet.idper,VistaDet.numruc, VistaDet.nombre " _
        + vbCr + " ORDER BY VistaDet.nombre " _
        + vbCr + " PIVOT mae_rangos.descripcion; "
    
 
    '--ejecutar la conulta
    RST_Busq rst, nSQL, xCon
    
    If rst.State = 0 Then Exit Sub
    
    If rst.RecordCount = 0 Then
        MsgBox "No hay registros para mostrar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fraBarra.Visible = False
        Set rst = Nothing
        Exit Sub
    End If
    
    ProgressBar1.Max = rst.RecordCount
    
    Dim xCol As Integer '--posicion de la columna
    
    Me.MousePointer = vbHourglass
     
    xFila = Fg2.FixedRows
    
    If rst.RecordCount <> 0 Then
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo Salir:
        
        '--------------
        rst.MoveFirst
        For A = 1 To rst.RecordCount    '--GRUPO DE CLIENTE/PROVEEDOR
            DoEvents
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo Salir:
            ProgressBar1.Value = A

            Fg2.Rows = Fg2.Rows + 1
            xFila = Fg2.Rows - 1
            
            Fg2.TextMatrix(xFila, 1) = NulosC(rst("idper"))
            Fg2.TextMatrix(xFila, 2) = NulosC(rst("numruc"))
            Fg2.TextMatrix(xFila, 3) = NulosC(rst("nombre"))
            Fg2.TextMatrix(xFila, 4) = NulosN(rst("candoc"))
            
            '--colocar los importes de la tabla cruzada
            For xCol = 5 To Fg2.Cols - 1
                '--verificar que la columna exista en el rst
                If RstRegistroBuscaCampo(rst, NulosC(Fg2.TextMatrix(0, xCol))) = True Then
                    Fg2.TextMatrix(xFila, xCol) = Format(NulosN(rst(Fg2.TextMatrix(0, xCol))), FORMAT_MONTO)
                    
                End If
            Next xCol
            '-------------------------------------------------------------
            rst.MoveNext
            
            If rst.EOF = True Then Exit For
        Next A
        '-------------------------------------------------------------
        
        Fg2.Rows = Fg2.Rows + 1
        xFila = xFila + 1
        FORMATO_CELDA Fg2, xFila, 3, , True, , "TOTAL -->"
        
        '--colocar los importes de la tabla cruzada
        For xCol = 5 To Fg2.Cols - 1
            
            FORMATO_CELDA Fg2, xFila, xCol, , True, , Format(GRID_SUMAR_COL(Fg2, xCol, Fg2.FixedRows, Fg2.Rows - 2), FORMAT_MONTO)
                                                
        Next xCol
        '-------------------------------------------------------------
    End If
    
    Set rst = Nothing
    fraBarra.Visible = False
    Me.MousePointer = vbDefault
    MsgBox "La Consulta fue se realizó Correctamente", vbInformation, xTitulo
    Exit Sub
Salir:
    Me.MousePointer = vbDefault
    fraBarra.Visible = False
    Set rst = Nothing
    MsgBox "La Consulta fue Interrumpida", vbInformation, xTitulo
    Exit Sub
error:
   
    Me.MousePointer = vbDefault
'    Resume
    fraBarra.Visible = False
    Set rst = Nothing
    SHOW_ERROR Me.Name, "CargarCli2"
    
End Sub



Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 3 Then pExportar
    If Button.Index = 4 Then pImprimir
    If Button.Index = 6 Then pic_Click
End Sub

Private Sub pic_Click()
    Bloquea False
End Sub

Private Sub Bloquea(band As Boolean)
    '--permitira bloquear  o desbloquear los controles
    Toolbar1.Enabled = Not band
    
    OptCliente.Enabled = Not band
    OptProvee.Enabled = Not band
    opt4ta.Enabled = Not band
    
    OptSel1.Enabled = Not band
    OptSel2.Enabled = Not band
    
    Fg2.Enabled = Not band
    
    TxtCliPro.Enabled = Not band
    CmdBusCliPro.Enabled = Not band
    
    TxtIdMon.Enabled = Not band
    CmdBusMon.Enabled = Not band
    
    FraDetalle.Visible = band
    
End Sub



Function GenerarConsulta(Optional IdPer As Long = 0) As String
    '===================================================================================================
    'creado: 11/05/11 Por Johan Castro
    'Propósito: Generar la consulta a nivel de detalle de documentos por cliente,proveedor o Prestador de Servicio
    '           toma como base de calculo el saldo por pagar o cobrar
    '
    'Entradas:  IdPer = Codigo de cliente, proveedor, prestador de servicio
    '                   Por defecto toma valor a cero
    '
    'Resultados: Consulta segun parametros indicados
    '
    '===================================================================================================
    
    Dim nSQL As String
    Dim nSQLFiltro As String
    Dim nFechaRaiz As String '--fecha para comparar el numero de dias que hay de diferencia con la fecha de vencimiento
    Dim nSQLWhere As String '--almacenara la condicion de la consulta
    Dim nSQLWhere1 As String '--almacenara el filtro para no mostrar las notas de credito de compra y venta que
    Dim nSQLFiltroDocRef As String '--almacenara el filtro para no mostrar las notas de credito de compra y venta que
                                   '--esten relacionados a sus documentos de referencia
    
    
    '--verificar el año de trabajo, si es distinto a la fecha actual tomar como base ultimo dia del año
    If AnoTra <> Year(Date) Then
        nFechaRaiz = "31/12/" & AnoTra
    Else
        nFechaRaiz = Date
    End If
    '---
    
    '--filtrando por cliente, proveedor o prestador de servicio
    If IdPer <> 0 Then
        nSQLWhere = " and vta_ventas.idcli= " & IdPer
        
        '--filtrando para percepcion
        If OptProvee.Value = True Then nSQLWhere1 = " and con_percepcion.idcli= " & IdPer
        
    End If
   
    '--aplicar filtro para no considerar en el filtro cuando se trate de renta de cuarta categoria
    If opt4ta.Value = False Then nSQLFiltroDocRef = " and vta_ventas.iddocref=0 "

    '--detalle
    nSQL = "SELECT Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4) AS registro, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, vta_ventas.numser & '-' & vta_ventas.numdoc AS numerodoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.abrev AS condpago, mae_condpago.numdia, mae_moneda.simbolo, " _
                + vbCr + " IIf(vta_ventas.tc <> 0, vta_ventas.tc, IIf(con_tc.impven Is Null, 0, con_tc.impven)) As tipcam, " _
                + vbCr + " IIf(vta_ventas.tipdoc=7,(-1)*vta_ventas.imptotdoc,vta_ventas.imptotdoc) AS impreal, " _
                + vbCr + " IIf(vta_ventas.tipdoc=7,(-1)*vta_ventas.impsal,vta_ventas.impsal) AS impsalreal, " _
                + vbCr + " IIf(vta_ventas.idmon=1,[impsalreal],[impsalreal]*[tipcam]) AS impsalexpmn, " _
                + vbCr + " IIf(vta_ventas.idmon=2,[impsalreal],IIf([tipcam]=0,0,[impsalreal]/[tipcam])) AS impsalexpme, " _
                + vbCr + " CDate('" & nFechaRaiz & "') AS fchraiz, DateDiff('d',[fchraiz],vta_ventas.fchven) AS numdias, " _
                + vbCr + " vta_ventas.idcli as idper, vta_ventas.id as iddoc " _
                + vbCr + " FROM (((((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha " _
                + vbCr + " Where (((vta_ventas.impsal) <> 0)) " & nSQLWhere & nSQLFiltroDocRef & " "
    
    '--modificar la consulta cuando seleccionen anticuamiento de proveedores o prestador de servicio
    If OptProvee.Value = True Then
        nSQL = Replace(nSQL, "vta_ventas", "com_compras")
        nSQL = Replace(nSQL, "mae_cliente", "mae_prov")
        nSQL = Replace(nSQL, ".idcli", ".idpro")
        nSQL = Replace(nSQL, ".imptotdoc", ".imptot")
        
    ElseIf opt4ta.Value = True Then
        nSQL = Replace(nSQL, "vta_ventas", "com_honorarios")
        nSQL = Replace(nSQL, "mae_cliente", "mae_prov")
        nSQL = Replace(nSQL, ".idcli", ".idpro")
        nSQL = Replace(nSQL, ".imptotdoc", ".imptot")
    End If
    
    
    If OptProvee.Value = True Then
    
        nSQL = nSQL _
                + vbCr + " UNION ALL " _
                + vbCr + " SELECT Left(con_percepcion.numreg,2) & mae_libros.codsun & Right(con_percepcion.numreg,4) AS registro, mae_prov.numruc, mae_prov.nombre, mae_documento.abrev, " _
                + vbCr + " con_percepcion.numser & '-' & con_percepcion.numdoc AS numerodoc, con_percepcion.fchdoc,con_percepcion.fchven, mae_condpago.abrev AS condpago, mae_condpago.numdia, mae_moneda.simbolo, " _
                + vbCr + " IIf(con_percepcion.tc <> 0, con_percepcion.tc, IIf(con_tc.impven Is Null, 0, con_tc.impven)) As tipcam, " _
                + vbCr + " IIf(con_percepcion.tipdoc=7,(-1)*con_percepcion.imptotper,con_percepcion.imptotper) AS impreal, " _
                + vbCr + " IIf(con_percepcion.tipdoc=7,(-1)*con_percepcion.impsal,con_percepcion.impsal) AS impsalreal, " _
                + vbCr + " IIf(con_percepcion.idmon=1,[impsalreal],[impsalreal]*[tipcam]) AS impsalexpmn, " _
                + vbCr + " IIf(con_percepcion.idmon=2,[impsalreal],IIf([tipcam]=0,0,[impsalreal]/[tipcam])) AS impsalexpme, " _
                + vbCr + " CDate('" & nFechaRaiz & "') AS fchraiz, DateDiff('d',[fchraiz],con_percepcion.fchven) AS numdias, " _
                + vbCr + " con_percepcion.idcli as idper, con_percepcion.id as iddoc " _
                + vbCr + " FROM (((((con_percepcion LEFT JOIN mae_libros ON con_percepcion.idlib = mae_libros.id) LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id) LEFT JOIN mae_documento ON con_percepcion.tipdoc = mae_documento.id) LEFT JOIN mae_condpago ON con_percepcion.idconpag = mae_condpago.id) LEFT JOIN mae_moneda ON con_percepcion.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_percepcion.fchdoc = con_tc.fecha " _
                + vbCr + " Where (((con_percepcion.impsal) <> 0)) " & nSQLWhere1
    End If
    
        
    GenerarConsulta = nSQL
    
End Function



Private Sub FraDetalle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    OrigFX = x
    OrigFY = y
    FraDetalle.ZOrder 0
End Sub

Private Sub FraDetalle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then
        With FraDetalle
            .Move .Left + x - OrigFX, .Top + y - OrigFY
        End With
    End If
End Sub


