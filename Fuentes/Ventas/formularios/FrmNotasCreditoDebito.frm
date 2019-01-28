VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmNotasCreditoDebito 
   Caption         =   "Ventas - Notas de Credito y Debito"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12075
   LinkTopic       =   "Form2"
   ScaleHeight     =   7740
   ScaleWidth      =   12075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fraseldoc 
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   645
      Left            =   1920
      TabIndex        =   67
      Top             =   720
      Visible         =   0   'False
      Width           =   825
      Begin VB.CommandButton cmdsalirseldoc 
         Height          =   555
         Left            =   2400
         Picture         =   "FrmNotasCreditoDebito.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   615
         Width           =   750
      End
      Begin VB.CommandButton cmdokseldoc 
         Height          =   555
         Left            =   1455
         Picture         =   "FrmNotasCreditoDebito.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   615
         Width           =   750
      End
      Begin VB.ComboBox cbonotascredeb 
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   240
         Width           =   4245
      End
   End
   Begin VB.Frame Fradocsproc 
      Caption         =   "Documentos Facturados en pantalla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3105
      Left            =   6885
      TabIndex        =   62
      Top             =   3930
      Visible         =   0   'False
      Width           =   4665
      Begin VB.CommandButton cmdSalirdocsproc 
         Height          =   630
         Left            =   3765
         Picture         =   "FrmNotasCreditoDebito.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   2415
         Width           =   750
      End
      Begin VB.CommandButton cmdOKdocsproc 
         Height          =   630
         Left            =   210
         Picture         =   "FrmNotasCreditoDebito.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   2415
         Width           =   750
      End
      Begin VB.CommandButton cmdEliminarOKdocsproc 
         Height          =   630
         Left            =   1140
         Picture         =   "FrmNotasCreditoDebito.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2415
         Width           =   765
      End
      Begin VSFlex7Ctl.VSFlexGrid fgdocsproc 
         Height          =   2070
         Left            =   120
         TabIndex        =   66
         Top             =   285
         Width           =   4395
         _cx             =   7752
         _cy             =   3651
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmNotasCreditoDebito.frx":0D2A
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
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   9195
      TabIndex        =   41
      Top             =   210
      Visible         =   0   'False
      Width           =   2970
      Begin VB.CommandButton CmdOk 
         Height          =   660
         Left            =   840
         Picture         =   "FrmNotasCreditoDebito.frx":0DC1
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2865
         Width           =   720
      End
      Begin VB.CommandButton CmdSalir 
         Height          =   660
         Left            =   1605
         Picture         =   "FrmNotasCreditoDebito.frx":10CB
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2865
         Width           =   720
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   210
         TabIndex        =   44
         Top             =   420
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   50593794
         CurrentDate     =   38919
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   0
         X2              =   2955
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   2955
         X2              =   2955
         Y1              =   15
         Y2              =   3630
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   3600
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   0
         X2              =   2955
         Y1              =   3615
         Y2              =   3615
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         Height          =   315
         Left            =   30
         Top             =   30
         Width           =   2910
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccionar Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   150
         TabIndex        =   45
         Top             =   75
         Width           =   1425
      End
   End
   Begin VB.Frame fraconsdocref 
      Caption         =   "Documentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1605
      Left            =   5955
      TabIndex        =   37
      Top             =   645
      Visible         =   0   'False
      Width           =   2880
      Begin VB.CommandButton CmdSalirRef 
         Height          =   630
         Left            =   3525
         Picture         =   "FrmNotasCreditoDebito.frx":13D5
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4095
         Width           =   750
      End
      Begin VB.CommandButton CmdOkRef 
         Height          =   630
         Left            =   2520
         Picture         =   "FrmNotasCreditoDebito.frx":16DF
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   4095
         Width           =   750
      End
      Begin VSFlex7Ctl.VSFlexGrid Fgdocref 
         Height          =   3555
         Left            =   105
         TabIndex        =   40
         Top             =   465
         Width           =   7485
         _cx             =   13203
         _cy             =   6271
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmNotasCreditoDebito.frx":19E9
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
            Picture         =   "FrmNotasCreditoDebito.frx":1AB7
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNotasCreditoDebito.frx":1FFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNotasCreditoDebito.frx":238D
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNotasCreditoDebito.frx":2511
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNotasCreditoDebito.frx":2965
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNotasCreditoDebito.frx":2A7D
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNotasCreditoDebito.frx":2FC1
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNotasCreditoDebito.frx":3505
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNotasCreditoDebito.frx":3619
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNotasCreditoDebito.frx":372D
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNotasCreditoDebito.frx":3B81
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNotasCreditoDebito.frx":3CED
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7800
      Left            =   -15
      TabIndex        =   0
      Top             =   345
      Width           =   11985
      _cx             =   21140
      _cy             =   13758
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
      Appearance      =   1
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7380
         Left            =   45
         TabIndex        =   48
         Top             =   375
         Width           =   11895
         Begin VB.Frame fraitems 
            Height          =   720
            Left            =   6465
            TabIndex        =   77
            Top             =   2400
            Width           =   5040
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "Agregar Items Doc. Ref."
               Height          =   360
               Left            =   240
               TabIndex        =   79
               Top             =   240
               Width           =   2220
            End
            Begin VB.CommandButton cmddocsprocesados 
               Caption         =   "Eliminar Items de  Doc Ref."
               Height          =   360
               Left            =   2640
               TabIndex        =   78
               Top             =   240
               Width           =   2220
            End
         End
         Begin VB.Frame fradocref 
            Caption         =   "Motivo de Emision de Nota de Credito / Debito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   840
            Left            =   270
            TabIndex        =   35
            Top             =   2370
            Width           =   5835
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "Eliminar Item"
               Height          =   360
               Left            =   4620
               TabIndex        =   75
               Top             =   420
               Width           =   1140
            End
            Begin VB.TextBox txtMotivo 
               Height          =   285
               Left            =   1320
               TabIndex        =   74
               Text            =   "TxtMotivo"
               Top             =   480
               Width           =   3255
            End
            Begin VB.OptionButton optdevolucion 
               Caption         =   "Devolucion"
               Height          =   195
               Left            =   2400
               TabIndex        =   73
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optdescuento 
               Caption         =   "Descuento"
               Height          =   195
               Left            =   1200
               TabIndex        =   72
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton optanulacion 
               Caption         =   "Anulacion"
               Height          =   195
               Left            =   120
               TabIndex        =   71
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label lbltextomotivo 
               AutoSize        =   -1  'True
               Caption         =   "Comentario Ref."
               Height          =   195
               Left            =   120
               TabIndex        =   76
               Top             =   510
               Width           =   1140
            End
         End
         Begin VB.CommandButton CmdBusNumSer 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmNotasCreditoDebito.frx":4235
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1425
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipItem 
            Height          =   240
            Left            =   2460
            Picture         =   "FrmNotasCreditoDebito.frx":4367
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   510
            Width           =   240
         End
         Begin VB.CommandButton CmdBusCli 
            Height          =   240
            Left            =   3300
            Picture         =   "FrmNotasCreditoDebito.frx":4499
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1140
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmNotasCreditoDebito.frx":45CB
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   810
            Width           =   240
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2895
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   23
            Text            =   "TxtNumDoc"
            Top             =   1425
            Width           =   1440
         End
         Begin VB.CommandButton CmdBusCondicion 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmNotasCreditoDebito.frx":46FD
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1770
            Width           =   240
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   6885
            Picture         =   "FrmNotasCreditoDebito.frx":482F
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   21
            Text            =   "TxtNumSer"
            Top             =   1425
            Width           =   645
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   13
            Text            =   "TxtTipDoc"
            Top             =   810
            Width           =   915
         End
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   17
            Text            =   "TxtNumRuc"
            Top             =   1110
            Width           =   1770
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   6225
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   9
            Text            =   "TxtIdMon"
            Top             =   375
            Width           =   915
         End
         Begin VB.TextBox TxtConPag 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   26
            Text            =   "TxtConPag"
            Top             =   1740
            Width           =   915
         End
         Begin VB.TextBox TxtTipItem 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   5
            Text            =   "TxtTipItem"
            Top             =   480
            Width           =   915
         End
         Begin VB.Frame Frame4 
            Height          =   600
            Left            =   240
            TabIndex        =   49
            Top             =   6210
            Width           =   11295
            Begin VB.TextBox txtisc 
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
               Left            =   8070
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   54
               TabStop         =   0   'False
               Text            =   "TxtIsc"
               Top             =   180
               Width           =   1100
            End
            Begin VB.TextBox txtinafecto 
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
               Left            =   3810
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   53
               TabStop         =   0   'False
               Text            =   "TxtInafecto"
               Top             =   180
               Width           =   1100
            End
            Begin VB.TextBox TxtBruto 
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
               Left            =   1110
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   52
               TabStop         =   0   'False
               Text            =   "TxtBruto"
               Top             =   180
               Width           =   1100
            End
            Begin VB.TextBox TxtIGV 
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
               Left            =   6225
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   51
               TabStop         =   0   'False
               Text            =   "TxtIGV"
               Top             =   180
               Width           =   1140
            End
            Begin VB.TextBox TxtTotal 
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
               Left            =   10005
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   50
               TabStop         =   0   'False
               Text            =   "TxtTotal"
               Top             =   180
               Width           =   1100
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "I.S.C."
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
               Index           =   3
               Left            =   7500
               TabIndex        =   60
               Top             =   270
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Inafecto"
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
               Left            =   2535
               TabIndex        =   59
               Top             =   270
               Width           =   1140
            End
            Begin VB.Label LblIgvTasa 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "LblIgvTasa"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   5400
               TabIndex        =   58
               Top             =   270
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Bruto"
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
               Index           =   0
               Left            =   135
               TabIndex        =   57
               Top             =   270
               Width           =   885
            End
            Begin VB.Label LblRotulo 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V."
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
               Left            =   5010
               TabIndex        =   56
               Top             =   270
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total"
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
               Index           =   2
               Left            =   9390
               TabIndex        =   55
               Top             =   270
               Width           =   450
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   3015
            Left            =   270
            TabIndex        =   61
            Top             =   3240
            Width           =   11250
            _cx             =   19844
            _cy             =   5318
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmNotasCreditoDebito.frx":4961
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchDoc 
            Height          =   300
            Left            =   1800
            TabIndex        =   32
            Top             =   2055
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
            Locked          =   -1  'True
            Valor           =   "03/01/2004"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchVen 
            Height          =   300
            Left            =   4965
            TabIndex        =   34
            Top             =   2055
            Width           =   1290
            _ExtentX        =   2275
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
            Valor           =   "03/01/2004"
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Notas de Credito / Debito"
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
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   11715
         End
         Begin VB.Label LblTipoItem 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipoItem"
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
            Left            =   2760
            TabIndex        =   7
            Top             =   480
            Width           =   2715
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Item"
            Height          =   195
            Index           =   6
            Left            =   270
            TabIndex        =   4
            Top             =   465
            Width           =   660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Vencimiento"
            Height          =   195
            Index           =   3
            Left            =   3675
            TabIndex        =   33
            Top             =   2100
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emision"
            Height          =   195
            Index           =   2
            Left            =   255
            TabIndex        =   31
            Top             =   2100
            Width           =   1260
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio"
            Height          =   195
            Left            =   7455
            TabIndex        =   29
            Top             =   1785
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label LblTipoCambio 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   8700
            TabIndex        =   30
            Top             =   1740
            Visible         =   0   'False
            Width           =   2790
         End
         Begin VB.Label LblCondPag 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblCondPag"
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
            Left            =   2895
            TabIndex        =   28
            Top             =   1755
            Width           =   3360
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
            Left            =   7200
            TabIndex        =   11
            Top             =   375
            Width           =   2790
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Index           =   7
            Left            =   255
            TabIndex        =   16
            Top             =   1140
            Width           =   480
         End
         Begin VB.Label LblNomCli 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomCli"
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
            Left            =   3585
            TabIndex        =   19
            Top             =   1110
            Width           =   4935
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Documento"
            Height          =   195
            Index           =   1
            Left            =   255
            TabIndex        =   12
            Top             =   825
            Width           =   1410
         End
         Begin VB.Label LblNomDoc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomDoc"
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
            Left            =   2760
            TabIndex        =   15
            Top             =   795
            Width           =   5760
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Documento"
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   20
            Top             =   1455
            Width           =   1275
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2760
            Top             =   1530
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Condicion de Pago"
            Height          =   195
            Index           =   4
            Left            =   255
            TabIndex        =   25
            Top             =   1785
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   5580
            TabIndex        =   8
            Top             =   465
            Width           =   585
         End
         Begin VB.Label LblIdCliente 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCliente"
            Height          =   195
            Left            =   7440
            TabIndex        =   24
            Top             =   1410
            Visible         =   0   'False
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7380
         Left            =   -12540
         TabIndex        =   46
         Top             =   375
         Width           =   11895
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6975
            Left            =   90
            TabIndex        =   2
            Top             =   330
            Width           =   11820
            _ExtentX        =   20849
            _ExtentY        =   12303
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "T.D."
            Columns(0).DataField=   "abrev"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Moneda"
            Columns(1).DataField=   "simbolo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numerodoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Emi"
            Columns(3).DataField=   "fchdoc"
            Columns(3).NumberFormat=   "Short Date"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Cliente"
            Columns(4).DataField=   "nombre"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Forma Pago"
            Columns(5).DataField=   "desccond"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Importe"
            Columns(6).DataField=   "imptotdoc"
            Columns(6).NumberFormat=   "0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Fch. Ven."
            Columns(7).DataField=   "fchven"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Saldo"
            Columns(8).DataField=   "impsal"
            Columns(8).NumberFormat=   "0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Estado"
            Columns(9).DataField=   "EstadoVenta"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=900"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=820"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1402"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1323"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2566"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2487"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1773"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1693"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=4313"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=4233"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2090"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2011"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=1640"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1561"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=1773"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1693"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=1667"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1588"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(55)=   "Column(9).Width=1561"
            Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1482"
            Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=516"
            Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta Notas de Credito / Debito"
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
            Index           =   0
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Width           =   11610
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label LblMes 
            AutoSize        =   -1  'True
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   8235
            TabIndex        =   47
            Top             =   30
            Width           =   765
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   12630
         X2              =   24525
         Y1              =   375
         Y2              =   7755
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   12075
      _ExtentX        =   21299
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Nota de Credito / Debito"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Restaurar Nota de Credito / Debito"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Anular Nota de Credito / Debito"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar Nota de Credito / Debito"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Emitir Nota de Credito / Debito Anulada"
               EndProperty
            EndProperty
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
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
            Object.ToolTipText     =   "Imprimir Guia"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   9195
      X2              =   9195
      Y1              =   195
      Y2              =   3795
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar Item            "
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar Item                "
      End
   End
End
Attribute VB_Name = "FrmNotasCreditoDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstVent As New ADODB.Recordset
Dim QueHace As Integer
Dim TasaImpuesto As Double
Dim CaracteresNumericos As String
Dim SeEjecuto As Boolean
Dim ValTipCam As Double
Dim xDescImp As String
Dim xIdCuenTasa As Integer  'codigo de la cuenta contable del impuesto
Dim xCuentaDoc As Integer   'codigo de la cuenta contable del documento
Dim xMes As Integer         'numero de mes en el que se realiza la operacion
Dim Mostrando As Boolean

'Dim swguiafact '0 No se facturaron, 1 Se facturaron




Sub CambiarMes()
    Toolbar1.Enabled = False
    TabOne1.Enabled = False
    Frame5.Left = 4455
    Frame5.Top = 2100
    Frame5.Visible = True
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    Dim rs As New ADODB.Recordset
    If RstVent.RecordCount = 0 Then
                MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
                Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de eliminar la " & RstVent("nomdoc") & " seleccionada?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM con_diario WHERE idmov = " & RstVent("id") & " AND idlib = 2 AND iddoc = " & RstVent("tipdoc") & ""
        
        
        RST_Busq rs, "SELECT * FROM vta_ventas WHERE  vta_ventas.idnumref =" & RstVent("id") & "", xCon
        
        Do While Not rs.EOF
            'Actualizamos el campo  idnumref = 0 de la tabla ventas
            xCon.Execute " UPDATE vta_ventas SET vta_ventas.idnumref = 0 WHERE vta_ventas.id = " & rs("id") & ""
            rs.MoveNext
        Loop
        
        
        xCon.Execute "DELETE * FROM vta_notascreabo WHERE id = " & RstVent("id") & ""
        
        
        MsgBox "La " & RstVent("nomdoc") & " se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstVent.Requery
        Dg1.Refresh
    End If
    
    Set rs = Nothing
      
End Sub

Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    Toolbar1.Buttons(11).Enabled = Not Toolbar1.Buttons(11).Enabled
    
    Toolbar1.Buttons(13).Enabled = Not Toolbar1.Buttons(13).Enabled
    Toolbar1.Buttons(15).Enabled = Not Toolbar1.Buttons(15).Enabled
End Sub
Sub RestaurarFactura()
    
    'Se restaura una Nota de Credito y/o Debito  Anulada
    Dim Rpta As Integer
    
    Rpta = MsgBox("Esta seguro de restaurar la " & RstVent("nomdoc") & " Nº " + RstVent("numser") & "-" & RstVent("numdoc"), vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption)
    If Rpta = vbYes Then
        xCon.Execute "UPDATE vta_notascreabo SET vta_notascreabo.Anulado = 0, " _
            & " WHERE vta_notascreabo.id =" & RstVent("id") & ""
        
        xCon.Execute "DELETE * FROM vta_notascreabodet WHERE vta_notascreabodet.idvta =" & RstVent("id") & ""
        RstVent.Requery
        Dg1.Refresh
        MsgBox "La " & RstVent("nomdoc") & " se restauro con exito", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
    End If
End Sub

Sub Anular()

    If RstVent.RecordCount = 0 Then
       MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
       Exit Sub
    End If
   
   
   'Validamos si la factura esta anulada
    If RstVent("Anulado") = -1 Then
        MsgBox "La " & RstVent("nomdoc") & " ya fue anulada, seleccione otra", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
        Exit Sub
    End If
    
    Dim rs As New ADODB.Recordset
    Dim Rpta As Integer
    Dim A As Integer
    Rpta = MsgBox("¿Esta seguro de anular la " & RstVent("nomdoc") & " " & RstVent("numser") & "-" & RstVent("numdoc") + "?", vbYesNo + vbDefaultButton1 + vbQuestion, xTitulo)
    
    If Rpta = vbYes Then
        xCon.Execute "UPDATE vta_notascreabo SET vta_notascreabo.Anulado = -1,   " _
            & " vta_notascreabo.impbru =0, vta_notascreabo.impinaf = 0, vta_notascreabo.impigv=0,  vta_notascreabo.impisc = 0,  " _
            & " vta_notascreabo.impotr =0, vta_notascreabo.imptotdoc = 0,  vta_notascreabo.impsal = 0  " _
            & " WHERE vta_notascreabo.id = " & RstVent("id") & " "
        
        
        'Buscamos todos los documentos relacionados con la NOTA DE CREDITO / DEBITO
        
        RST_Busq rs, "SELECT * FROM vta_ventas WHERE  vta_ventas.idnumref =" & RstVent("id") & "", xCon
        
        Do While Not rs.EOF
            'Actualizamos el campo  idnumref = 0 de la tabla ventas
            xCon.Execute " UPDATE vta_ventas SET vta_ventas.idnumref = 0 WHERE vta_ventas.id = " & rs("id") & ""
            rs.MoveNext
        Loop
        
        xCon.Execute "DELETE * FROM vta_notascreabodet WHERE vta_notascreabodet.idvta = " & RstVent("id") & ""
        
                
        MsgBox "La " & RstVent("nomdoc") & "se anulo con exito ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstVent.Requery
        Dg1.Refresh
    End If
    Set rs = Nothing
End Sub

Sub Cancelar()
Dim X As Integer
    Bloquea
    
    Fg1.ColComboList(1) = ""
    
    
    Label5.Caption = "Detalle de Notas de Credito / Debito "
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
     
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
       
       
       
    'Colocamos en el campo estado 0  de la tabla guia que indica no  esta facturado
    
    'If fgdocsproc.Rows - 1 > 0 Then
    'If swguiafact = 0 Then
    '    For X = 1 To fgdocsproc.Rows - 1
    '            xCon.Execute " UPDATE vta_guia SET Vta_guia.Estado = 0 WHERE vta_guia.id = " & Val(fgdocsproc.TextMatrix(X, 1)) & ""
    '    Next
    '    fgdocsproc.Rows = 1
    'End If
    'End If
    'swguiafact = 0
    
End Sub

Sub Nuevo()
    QueHace = 1
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Label5.Caption = "Agregando Notas de Credito / Debito "
    Fg1.ColComboList(1) = "|..."

    Fg1.SelectionMode = flexSelectionFree

    Fg1.Rows = 1
    'Fg1.Rows = Fg1.Rows + 1
    TxtTipItem.SetFocus
End Sub

Sub Modificar()
    If RstVent.RecordCount = 0 Then
            MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
            Exit Sub
    End If
    
    QueHace = 2
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    MuestraSegundoTab
    Label5.Caption = "Modificando Notas de Credito / Debito"
    Fg1.ColComboList(1) = "|..."
    
    Fg1.SelectionMode = flexSelectionFree
    
    TxtTipItem.SetFocus
End Sub

Sub MuestraSegundoTab()
    Dim rs As New ADODB.Recordset
    TxtTipItem.Text = IIf(RstVent("idtipo") = 0, "", NulosN(RstVent("idtipo")))
    
    TxtTipDoc.Text = RstVent("tipdoc")
    TxtNumRuc.Text = RstVent("numruc")
    TxtNumSer.Text = RstVent("numser")
    TxtNumDoc.Text = RstVent("numdoc")
    TxtFchDoc.Valor = RstVent("fchdoc")
    TxtFchVen.Valor = RstVent("fchven")
    TxtConPag.Text = RstVent("idconpag")
    TxtIdMon.Text = RstVent("idmon")
    
    TxtBruto.Text = Format(RstVent("impbru"), "0.00")
    TxtIGV.Text = Format(RstVent("impigv"), "0.00")
    TxtTotal.Text = Format(RstVent("imptotdoc"), "0.00")
    txtinafecto.Text = Format(RstVent("impinaf"), "0.00")
    txtisc.Text = Format(RstVent("impisc"), "0.00")
    
    LblTipoItem.Caption = NulosC(RstVent("desctipcom"))
    LblNomDoc.Caption = RstVent("nomdoc")
    LblNomCli.Caption = RstVent("nombre")
    LblCondPag.Caption = NulosC(RstVent("desccond"))
    TxtNumRuc.Text = RstVent("numruc")
    LblMoneda.Caption = RstVent("descmon")
    LblIdCliente.Caption = RstVent("idcli")
    
    'RST_Busq RS, " SELECT vta_ventas.id, vta_ventas.fchdoc, mae_documento.descripcion, vta_ventas.numser, vta_ventas.numdoc " _
    '            & " FROM mae_documento INNER JOIN vta_ventas ON mae_documento.id = vta_ventas.tipdoc " _
    '            & " WHERE vta_ventas.id = " & RstVent("idnumref") & "", xCon

    'If RS.RecordCount > 0 Then
    
    
    
    Set rs = Nothing
    
    'xIdCuenTasa = NulosN(RstVent("idcuen"))

    If RstVent("idmon") = 1 Then
        LblTipoCambio.Visible = False
    Else
        LblTipoCambio.Visible = True
        LblTipoCambio.Caption = RstVent("impven")
    End If
    
    Dim RstDet As New ADODB.Recordset
    
    Mostrando = True
    Fg1.Rows = 1
    
        
    RST_Busq RstDet, " SELECT vta_notascreabodet.*, alm_inventario.descripcion, mae_unidades.abrev,alm_inventario.idcuentaven,alm_inventario.idtipven  " _
                     & " FROM (vta_notascreabodet LEFT JOIN alm_inventario ON vta_notascreabodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON vta_notascreabodet.idunimed = mae_unidades.id " _
                     & " WHERE (((vta_notascreabodet.idvta)=" & RstVent("id") & "))", xCon
    
    If RstDet.RecordCount <> 0 Then
        Do While Not RstDet.EOF
            Fg1.Rows = Fg1.Rows + 1
                                    
            If NulosC(RstDet("descripcion")) = "" Then
                Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstDet("descripusu"))
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstDet("descripcion"))
            End If
            
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDet("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = IIf(RstDet("preuni") = 0, "", Format(RstDet("preuni"), "0.00"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = IIf(RstDet("canpro") = 0, "", Format(RstDet("canpro"), "0.00"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = IIf(RstDet("imptot") = 0, "", Format(RstDet("imptot"), "0.00"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = RstDet("iditem")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = RstDet("idunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(RstDet("idcuentaven"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(RstDet("idtipven"))
            RstDet.MoveNext
        Loop
    End If
    
    Set RstDet = Nothing
    Mostrando = False
    
    'cargamos el codigo de la cuenta contable del documento
    'Set RstDet = BuscaConCriterio("SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, " _
    '    & " con_planctas.cuenta, mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuenvta as cuentaimp " _
    '    & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuenvta = con_planctas.id) " _
    '    & " ON mae_documento.idimp = mae_impuestos.id WHERE mae_documento.id  = " & Val(TxtTipDoc.Text) & "", xCon)
    
    
    Set RstDet = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & Val(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & Val(TxtIdMon) & " and tipope = -1", xCon)

    If RstDet.RecordCount = 1 Then
        xCuentaDoc = RstDet("idcuen")
    End If
    Set RstDet = Nothing
    
End Sub

Sub Bloquea()
    TxtTipItem.Locked = Not TxtTipItem.Locked
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtFchDoc.Locked = Not TxtFchDoc.Locked
    TxtFchVen.Locked = Not TxtFchVen.Locked
    TxtConPag.Locked = Not TxtConPag.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
    
    'CmdAddItem.Enabled = False
    'CmdDelItem.Enabled = False
    
End Sub

Sub Blanquea()
    TxtTipItem.Text = ""
    TxtTipDoc.Text = ""
    TxtNumRuc.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtFchDoc.Valor = ""
    
    TxtFchVen.Valor = ""
    TxtConPag.Text = ""
    TxtIdMon.Text = ""
    
    LblNomDoc.Caption = ""
    LblNomCli.Caption = ""
    LblCondPag.Caption = ""
    LblMoneda.Caption = ""
    LblIdCliente.Caption = ""
    LblTipoItem.Caption = ""
    

    
    txtinafecto = ""
    txtisc = ""
    TxtBruto.Text = ""
    TxtIGV.Text = ""
    TxtTotal.Text = ""
    TxtMotivo = ""

    Fg1.Rows = 1
End Sub

Private Sub CmdAddItem_Click()
    
    Dim xRs1 As New ADODB.Recordset
    
    
    

    
         If Val(LblIdCliente) = 0 Then
           MsgBox "Seleccione Cliente", vbInformation, Me.Caption
           Exit Sub
         End If
    

            
        fraconsdocref.Height = 4800
        fraconsdocref.Width = 7620
        fraconsdocref.Visible = True
    
    
    
    RST_Busq xRs1, "SELECT vta_ventas.id, vta_ventas.fchdoc, Format([vta_ventas]![numser],'0000') & '-' & Format([vta_ventas]![numdoc],'0000000000') AS [NroDoc], mae_documento.descripcion as [Nomdoc], vta_ventas.imptotdoc, mae_moneda.descripcion " & _
                   "FROM mae_moneda INNER JOIN (mae_documento INNER JOIN vta_ventas ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon " & _
                   "WHERE vta_ventas.idcli = " & Val(Me.LblIdCliente) & " AND vta_ventas.idnumref = 0 ", xCon

         With Me.Fgdocref
         .Rows = 1
         
         
         .ColWidth(1) = 300  'Id
         .ColWidth(2) = 1200  'Fecha
         .ColWidth(3) = 1500 'Nro Doc
         .ColWidth(4) = 1200  'Nom Doc
         .ColWidth(5) = 1200  'Importe
         .ColWidth(6) = 1200  'Descripcion
         
         
         
         
     Do While Not xRs1.EOF
     
         .AddItem ""
         .TextMatrix(.Rows - 1, 1) = xRs1("ID")
         .TextMatrix(.Rows - 1, 2) = xRs1("fchdoc")
         .TextMatrix(.Rows - 1, 3) = xRs1("Nrodoc")
         .TextMatrix(.Rows - 1, 4) = xRs1("Nomdoc")
         .TextMatrix(.Rows - 1, 5) = Format(xRs1("Imptotdoc"), "0.00")
         .TextMatrix(.Rows - 1, 6) = xRs1("Descripcion")
         xRs1.MoveNext
     Loop
      
      
     End With
     Fgdocref.SetFocus

    Set xRs1 = Nothing
    

End Sub


Private Sub CmdBusCondicion_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_condpago ORDER BY descripcion"
    
    xform.Titulo = "Buscando Condicion de Pago"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtConPag.Text = xRs("id")
        LblCondPag.Caption = xRs("descripcion")
        TxtFchDoc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusNumSer_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "iddoc":       xCampos(0, 2) = "1000":    xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion": xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Serie":          xCampos(2, 1) = "numser":      xCampos(2, 2) = "1000":    xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nro Documento":  xCampos(3, 1) = "numdoc":      xCampos(3, 2) = "1500":    xCampos(3, 3) = "C"
    
    
    xform.SQLCad = "SELECT mae_documento.descripcion, mae_series.iddoc, mae_series.numser, mae_series.numdoc " & _
                   " FROM mae_documento INNER JOIN mae_series ON mae_documento.id = mae_series.iddoc where mae_series.iddoc = " & Val(TxtTipDoc) & ""

    xform.Titulo = "Buscando Series"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numser"
    xform.CampoBusca = "numser"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtNumSer.Text = Format(xRs("numser"), "0000")
        TxtNumDoc = HallaNumdocVenta(Val(TxtTipDoc.Text), NulosC(TxtNumSer.Text), xCon)
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdBusCli_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Cliente":    xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_cliente.nombre, mae_cliente.numruc, mae_cliente.id From mae_cliente"
    
    xform.Titulo = "Buscando Cliente"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtNumRuc.Text = xRs("numruc")
        LblNomCli.Caption = xRs("nombre")
        LblIdCliente.Caption = xRs("id")
        TxtNumSer.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_moneda ORDER BY descripcion"
    
    xform.Titulo = "Buscando Moneda"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdMon.Text = xRs("id")
        LblMoneda.Caption = xRs("descripcion")
        Fg1.SetFocus
        
        If Trim(TxtIdMon.Text) = "1" Then
            LblTipCam.Visible = False
            LblTipoCambio.Visible = False
        Else
            If TxtFchDoc.Valor = "" Then
                MsgBox "No ha especificado la fecha del documento, no se puede determinar " & Chr(13) _
                    & "la fecha del tipo de cambio para este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                
                TxtIdMon.Text = ""
                TxtFchDoc.SetFocus
                Exit Sub
            End If
            LblTipCam.Visible = True
            LblTipoCambio.Visible = True
            Set xRs = Nothing
            Set xRs = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = CDATE('" & TxtFchDoc.Valor & "')", xCon)
            If xRs.RecordCount = 1 Then
                LblTipoCambio.Caption = Format(xRs("impven"), "0.000")
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDoc_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
        & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuenvta  as cuentaimp" _
        & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuenvta = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id   WHERE mae_documento.id = 7 OR mae_documento.id = 8 "
    
    Dim xImpuesto As Double
    
    xform.Titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtTipDoc.Text = xRs("id")
        LblNomDoc.Caption = xRs("descripcion")
        TasaImpuesto = NulosN(xRs("tasa"))
        xDescImp = xRs("descripcion")
        xIdCuenTasa = NulosN(xRs("cuentaimp"))
        LblRotulo = Trim(NulosC(xRs("abreimp"))) + " (       )"
        LblIgvTasa.Caption = Trim(Str(TasaImpuesto)) + "%"

        
        Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & Val(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & Val(TxtIdMon) & " and tipope = -1", xCon)
            If xRs2.RecordCount > 0 Then
                xCuentaDoc = NulosN(xRs2("idcuen"))
            End If
            Set xRs2 = Nothing
        
        
        TxtNumRuc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipItem_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_tipoproducto.* FROM mae_tipoproducto"
    
    xform.Titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtTipItem.Text = xRs("id")
        LblTipoItem = xRs("descripcion")
        TxtTipDoc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDelItem_Click()
 
    
    If Fg1.Rows - 1 > 0 Then
    
        If Fg1.Rows - 1 = 1 Then
            If Fg1.TextMatrix(Fg1.Row, 8) = "" Then
                Fg1.Rows = 1
            Else
                MsgBox "Para eliminar items de Doc. de Referencia Click en Eliminar Items de Doc Ref.", vbInformation, xTitulo
            End If
            
        Else
            If Fg1.TextMatrix(Fg1.Row, 8) = "" Then
                Fg1.RemoveItem Fg1.Row
            Else
                
                MsgBox "Para eliminar items de Doc. de Referencia Click en Eliminar Items de Doc Ref.", vbInformation, xTitulo
            End If
        End If
        
    End If
    HallarTotal
End Sub



Private Sub cmddocsprocesados_Click()
    

Fradocsproc.Top = 3750
Fradocsproc.Left = 6885
Fradocsproc.Width = 4665
Fradocsproc.Height = 3105

    
    Toolbar1.Enabled = False
    TabOne1.Enabled = False
    Fradocsproc.Visible = True
    fgdocsproc.SetFocus

End Sub

Private Sub cmdEliminarOKdocsproc_Click()
Dim Rstventa As New ADODB.Recordset
Dim X As Integer
Dim Y As Integer



If fgdocsproc.Rows - 1 > 0 Then
        
        If fgdocsproc.Rows - 1 = 1 Then
            'Colocamos en el campo idnumref  = 0  de la tabla vta_ventas  que indica que tiene nota de credito / debito
            xCon.Execute " UPDATE vta_ventas SET Vta_ventas.idnumref = 0 WHERE Vta_ventas.id  = " & Val(fgdocsproc.TextMatrix(fgdocsproc.Row, 1)) & ""
             fgdocsproc.Rows = 1
             Fg1.Rows = 1
             HallarTotal
             Call cmdSalirdocsproc_Click
             Exit Sub
        Else
             
            
            
            
                With Fg1
                    If Val(TxtTipDoc) = 7 Then 'SI ES NOTA DE CREDITO ELIMINAMOS EL ITEM POR CODIGO DE ITEM
                        For X = 1 To Fg1.Rows - 1
                          
                            RST_Busq Rstventa, "Select vta_ventasdet.* From vta_ventasdet where vta_ventasdet.Idvta = " & Val(fgdocsproc.TextMatrix(fgdocsproc.Row, 1)) & " and vta_ventasdet.IdItem = " & Val(Fg1.TextMatrix(X, 6)) & "", xCon
                            
                            If Rstventa.RecordCount > 0 Then
                                    .TextMatrix(X, 4) = Val(.TextMatrix(X, 4)) - Rstventa("canpro")
                                
                            End If
                            
                        Next
                                                
                            Y = 1
                        
                                For X = 1 To Fg1.Rows - 1
                                    If Val(.TextMatrix(Y, 4)) = 0 Then
                                        Fg1.RemoveItem (Y)
                                    Else
                                        Y = Y + 1
                                    End If
                                Next
                        
                    Else 'SI ES NOTA DE DEBITO ELIMINAMOS EL ITEM POR ID DEL DOCUMENTO DE REFERENCIA
                        For X = 1 To Fg1.Rows - 1
                            If Val(.TextMatrix(X, 10)) = Val(fgdocsproc.TextMatrix(fgdocsproc.Row, 1)) Then
                                Fg1.RemoveItem (X)
                            End If
                        Next
                    End If
                    
                End With
            
            'Colocamos en el campo idnumref  = 0  de la tabla vta_ventas  que indica que tiene nota de credito / debito
                xCon.Execute " UPDATE vta_ventas SET Vta_ventas.idnumref = 0 WHERE Vta_ventas.id  = " & Val(fgdocsproc.TextMatrix(fgdocsproc.Row, 1)) & ""
                fgdocsproc.RemoveItem fgdocsproc.Row
                HallarTotal
                Call cmdSalirdocsproc_Click
        End If
End If
 


Set Rstventa = Nothing
End Sub


Private Sub CmdOk_Click()
        
    Dim xFecha As String
    xFecha = Format(MonthView1.Value, "dd/mm/yy")
        
    'xFchIni = "01/" + Mid(Format(CDate(xFecha), "dd/mm/yy"), 4, 5)
    'xfchfin = Trim(Format(HallaDiasMes(CDate(xFecha)), "00")) + "/" + Mid(Format(CDate(xFecha), "dd/mm/yy"), 4, 5)
    

    RST_Busq RstVent, "SELECT vta_notascreabo.*, mae_cliente.nombre, [vta_notascreabo]![numser]+'-'+[vta_notascreabo]![numdoc] AS numerodoc, IIf(vta_notascreabo.Anulado = 0, 'Facturado', 'Anulado') AS [EstadoVenta], " _
            & " mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, " _
            & " mae_moneda.descripcion AS descmon, mae_moneda.simbolo, mae_impuestos.idcuenvta, con_tc.impcom, mae_tipoproducto.descripcion AS desctipcom" _
            & " FROM (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN ((mae_documento LEFT JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) " _
            & " RIGHT JOIN (mae_condpago RIGHT JOIN (vta_notascreabo LEFT JOIN con_tc ON vta_notascreabo.fchdoc = con_tc.fecha) ON " _
            & " mae_condpago.id = vta_notascreabo.idconpag) ON mae_documento.id = vta_notascreabo.tipdoc) ON mae_moneda.id = vta_notascreabo.idmon) " _
            & " ON mae_cliente.id = vta_notascreabo.idcli) LEFT JOIN mae_tipoproducto ON vta_notascreabo.idtipo = mae_tipoproducto.id " _
            & " WHERE (((vta_notascreabo.numreg) Like '" & Format(xMes, "00") & "%'))", xCon
    
    
        
        Set Dg1.DataSource = RstVent
    
    
    CmdSalir_Click
End Sub

Private Sub cmdOKdocsproc_Click()
    Dim xRs As New ADODB.Recordset
        
    
         If fgdocsproc.Rows - 1 <= 0 Then
            cmdSalirdocsproc_Click
            Exit Sub
         End If
         
         fraconsdocref.Height = 4800
        fraconsdocref.Width = 7620
        fraconsdocref.Visible = True
         
         RST_Busq xRs, "SELECT vta_guia.numser, vta_guia.numdoc, alm_inventario.descripcion, vta_guiadet.canpro, mae_unidades.abrev " & _
                       "FROM vta_guia INNER JOIN (mae_unidades INNER JOIN (alm_inventario INNER JOIN vta_guiadet ON alm_inventario.id = vta_guiadet.iditem) ON mae_unidades.id = alm_inventario.idunimed) ON vta_guia.id = vta_guiadet.idgui " & _
                       " WHERE vta_guia.id =  " & Val(fgdocsproc.TextMatrix(Me.fgdocsproc.Row, 1)), xCon

         With Me.Fgdocref
             .Rows = 1
             .Cols = 5
                      
            .ColWidth(1) = 1500  'Nro Doc
            .ColWidth(2) = 3500 'Item
            .ColWidth(3) = 800 'Cantidad
            .ColWidth(4) = 800 'Uni Med
         
            .TextMatrix(0, 1) = "Nro Doc"
            .TextMatrix(0, 2) = "Item"
            .TextMatrix(0, 3) = "Cantidad"
            .TextMatrix(0, 4) = "Unid Med"
           
           Do While Not xRs.EOF
                .AddItem ""
                .TextMatrix(.Rows - 1, 1) = Format(xRs("numser"), "0000") & " - " & Format(xRs("numdoc"), "000000000")
                .TextMatrix(.Rows - 1, 2) = xRs("descripcion")
                .TextMatrix(.Rows - 1, 3) = xRs("canpro")
                .TextMatrix(.Rows - 1, 4) = xRs("abrev")
                 xRs.MoveNext
           Loop
             
        End With
        Fradocsproc.Visible = False
        Set xRs = Nothing
    
End Sub

Private Sub CmdOkRef_Click()
Dim Rstventa As New ADODB.Recordset
Dim Pos As Integer
Dim X As Integer
Dim swaviso As Integer
Dim xIdDocRef  As Integer 'Almacena el Id de la Nota de Credito / Debito

If Fgdocref.Rows - 1 = 0 Then
    CmdSalirRef_Click
    Exit Sub
End If





If Val(TxtTipDoc) = 7 Then   'si es nota de de credito

        RST_Busq Rstventa, "SELECT vta_ventasdet.preuni, vta_ventasdet.iditem, alm_inventario.idcuentaven, alm_inventario.idtipven, alm_inventario.descripcion, vta_ventasdet.canpro, vta_ventasdet.idunimed, mae_unidades.abrev, vta_ventas.idcli,vta_ventas.id, mae_cliente.nombre, mae_cliente.numruc  " & _
                     "FROM mae_cliente INNER JOIN (vta_ventas INNER JOIN (mae_unidades INNER JOIN (alm_inventario INNER JOIN vta_ventasdet ON alm_inventario.id = vta_ventasdet.iditem) ON mae_unidades.id = alm_inventario.idunimed) ON vta_ventas.id = vta_ventasdet.idvta) ON mae_cliente.id = vta_ventas.idcli " & _
                     "WHERE vta_ventas.id =" & Val(Fgdocref.TextMatrix(Me.Fgdocref.Row, 1)) & " ", xCon
    
        If Rstventa.RecordCount > 0 Then
                
                LblNomCli = Rstventa("nombre")
                LblIdCliente = Rstventa("idcli")
                TxtNumRuc = Rstventa("numruc")
                
               'Añadimos a la lista de documentos a facturarse
                With fgdocsproc
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 1) = Fgdocref.TextMatrix(Fgdocref.Row, 1) 'ID del Documento de Referencia
                    .TextMatrix(.Rows - 1, 2) = Fgdocref.TextMatrix(Fgdocref.Row, 2) 'Fecha
                    .TextMatrix(.Rows - 1, 3) = Fgdocref.TextMatrix(Fgdocref.Row, 4) 'Documento
                    .TextMatrix(.Rows - 1, 4) = Fgdocref.TextMatrix(Fgdocref.Row, 3) 'Nro de documento
                End With
                        
        With Me.Fg1
         Do While Not Rstventa.EOF
                                                          
                swaviso = 0
                Pos = 0
                For X = 1 To .Rows - 1
                   If Rstventa("iditem") = Val(.TextMatrix(X, 6)) Then
                      swaviso = 1
                      Pos = X
                      Exit For
                   End If
                Next
                
                If swaviso = 0 Then
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 1) = Rstventa("descripcion")
                    .TextMatrix(.Rows - 1, 2) = Rstventa("abrev")
                    .TextMatrix(.Rows - 1, 3) = Format(NulosN(Rstventa!preuni), "0.00")
                    .TextMatrix(.Rows - 1, 4) = Rstventa("canpro")
                    .TextMatrix(.Rows - 1, 6) = Rstventa("iditem")
                    .TextMatrix(.Rows - 1, 7) = Rstventa("idunimed")
                    .TextMatrix(.Rows - 1, 8) = Rstventa("idcuentaven")
                    .TextMatrix(.Rows - 1, 9) = Rstventa("idtipven")
                    .TextMatrix(.Rows - 1, 5) = Val(.TextMatrix(.Rows - 1, 3)) * Val(.TextMatrix(.Rows - 1, 4))
                    .TextMatrix(.Rows - 1, 5) = Format(.TextMatrix(.Rows - 1, 5), "0.00")
                Else
                    .TextMatrix(Pos, 4) = Val(.TextMatrix(Pos, 4)) + Rstventa("canpro")
                    .TextMatrix(Pos, 5) = Val(.TextMatrix(Pos, 3)) * Val(.TextMatrix(Pos, 4))
                    .TextMatrix(Pos, 5) = Format(.TextMatrix(Pos, 5), "0.00")
                End If
                     Rstventa.MoveNext
            Loop
                                                                                           
                    xIdDocRef = HallaCodigoTabla("vta_notascreabo", xCon, "id")
                        
                   'Colocamos en el campo idnumref  el valor del campo id de la tabla vta_notascreabo
                    xCon.Execute " UPDATE vta_ventas SET Vta_ventas.idnumref = " & xIdDocRef & " WHERE Vta_ventas.id  = " & Val(Fgdocref.TextMatrix(Fgdocref.Row, 1)) & ""
                       
                       'Eliminanamos la lista de Documentos de Referencia
                    If Fgdocref.Rows - 1 > 0 Then
                        If Fgdocref.Rows - 1 = 1 Then
                            Fgdocref.Rows = 1
                        Else
                            Fgdocref.RemoveItem Fgdocref.Row
                        End If
                    End If
            
            
        End With
    End If
Else 'SI ES NOTA DEBITO

        
        
        RST_Busq Rstventa, "SELECT vta_ventas.idcli,vta_ventas.id, vta_ventas.imptotdoc, mae_cliente.nombre, mae_cliente.numruc  " & _
                     "FROM mae_cliente INNER JOIN vta_ventas  ON mae_cliente.id = vta_ventas.idcli " & _
                     "WHERE vta_ventas.id =" & Val(Fgdocref.TextMatrix(Me.Fgdocref.Row, 1)) & " ", xCon
    
        If Rstventa.RecordCount > 0 Then
                
                LblNomCli = Rstventa("nombre")
                LblIdCliente = Rstventa("idcli")
                TxtNumRuc = Rstventa("numruc")
                
               'Añadimos a la lista de documentos a facturarse
                With fgdocsproc
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 1) = Fgdocref.TextMatrix(Fgdocref.Row, 1) 'ID del Documento de Referencia
                    .TextMatrix(.Rows - 1, 2) = Fgdocref.TextMatrix(Fgdocref.Row, 2) 'Fecha
                    .TextMatrix(.Rows - 1, 3) = Fgdocref.TextMatrix(Fgdocref.Row, 4) 'Documento
                    .TextMatrix(.Rows - 1, 4) = Fgdocref.TextMatrix(Fgdocref.Row, 3) 'Nro de documento
                End With
                        
                With Me.Fg1
                      Do While Not Rstventa.EOF
                        .AddItem ""
                        .TextMatrix(.Rows - 1, 3) = Format(NulosN(Rstventa!imptotdoc), "0.00")
                        .TextMatrix(.Rows - 1, 4) = 1
                        .TextMatrix(.Rows - 1, 5) = Val(.TextMatrix(.Rows - 1, 3)) * Val(.TextMatrix(.Rows - 1, 4))
                        .TextMatrix(.Rows - 1, 5) = Format(.TextMatrix(.Rows - 1, 5), "0.00")
                        .TextMatrix(.Rows - 1, 10) = Rstventa("id")
                        Rstventa.MoveNext
                      Loop
                    
                    xIdDocRef = HallaCodigoTabla("vta_notascreabo", xCon, "id")
                        
                   'Colocamos en el campo idnumref  el valor del campo id de la tabla vta_notascreabo
                    xCon.Execute " UPDATE vta_ventas SET Vta_ventas.idnumref = " & xIdDocRef & " WHERE Vta_ventas.id  = " & Val(Fgdocref.TextMatrix(Fgdocref.Row, 1)) & ""
                       
                       'Eliminanamos la lista de Documentos de Referencia
                    If Fgdocref.Rows - 1 > 0 Then
                        If Fgdocref.Rows - 1 = 1 Then
                            Fgdocref.Rows = 1
                        Else
                            Fgdocref.RemoveItem Fgdocref.Row
                        End If
                    End If
                End With
                
       End If
End If

Toolbar1.Enabled = True
TabOne1.Enabled = True
fraconsdocref.Visible = False
HallarTotal
Set Rstventa = Nothing

End Sub

Private Sub cmdokseldoc_Click()
Dim Rpta As Integer
    Dim RstCab As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim xId As Integer
    Dim xnumdoc As String
    Dim xnumser As String
    Dim xNumAsiento As String
    
    
   If Trim(cbonotascredeb) = "" Then
       MsgBox "Seleccione el documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
       Exit Sub
   End If
    
    RST_Busq xRs, "SELECT * FROM MAE_Series WHERE mae_series.iddoc = " & Val(Right(cbonotascredeb, 5)) & "", xCon
        
    If xRs.RecordCount > 0 Then
       xRs.MoveFirst
       xnumser = xRs("numser")
    Else
       MsgBox "Registre la serie en Mantenimiento de Series", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
       Set xRs = Nothing
       Exit Sub
    End If
    
       Rpta = MsgBox("Esta seguro de emitir una " & Trim(Mid(cbonotascredeb, 1, 100)) & " como anulada ", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
    If Rpta = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo LaCague
    xCon.BeginTrans
    
    'Validar si el nro de documento existe solo en modo adicionar documento
    RST_Busq RstCab, "SELECT * FROM vta_notascreabo", xCon
    

    xId = HallaCodigoTabla("vta_notascreabo", xCon, "id")
    RstCab.AddNew
    RstCab("id") = xId
    RstCab("idtipo") = 0
    RstCab("tipdoc") = Val(Right(cbonotascredeb, 5))
    RstCab("idcli") = 1
    
    xnumdoc = HallaNumdocVenta(Val(TxtTipDoc.Text), NulosC(TxtNumSer.Text), xCon)
    
    RstCab("numser") = Format(xnumser, "0000")
    RstCab("numdoc") = xnumdoc
    
    If NulosC(TxtFchDoc.Valor) <> "" Then RstCab("Fchdoc") = TxtFchDoc.Valor
    If NulosC(TxtFchVen.Valor) <> "" Then RstCab("Fchven") = TxtFchVen.Valor

    
    RstCab("idconpag") = 0
    RstCab("idmon") = 1
    RstCab("impbru") = 0
    RstCab("impinaf") = 0
    RstCab("impigv") = 0
    RstCab("impisc") = 0
    RstCab("impotr") = 0
    RstCab("imptotdoc") = 0
    RstCab("impsal") = 0
    RstCab("numreg") = Trim(Str(xMes)) + xNumAsiento
    RstCab("anulado") = -1
    
    'Determinamos si es una exportacion
    RstCab("idtipven") = 1 'en el cual puede ser venta afecta o inafecta para el registro de
                               'de ventas se valida por programa ver tabla mae_tipoventa
    RstCab.Update
                            
    'Actualizamos el numero de documento en la tabla Mae_series
    Call ActualizaNroDocumento(Val(xnumdoc), Val(Right(cbonotascredeb, 5)), Val(xnumser))
    
    MsgBox "La " & Trim(Mid(cbonotascredeb, 1, 100)) & " anulada se genero con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Toolbar1.Enabled = True
    TabOne1.Enabled = True
    Fraseldoc.Visible = False

    xCon.CommitTrans
    RstVent.Requery
    Dg1.Refresh
    Exit Sub

LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set xRs = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Private Sub CmdSalir_Click()
    Toolbar1.Enabled = True
    TabOne1.Enabled = True
    Frame5.Visible = False

End Sub

Private Sub cmdSalirdocsproc_Click()
    Fradocsproc.Visible = False
    Toolbar1.Enabled = True
    TabOne1.Enabled = True

End Sub

Private Sub CmdSalirRef_Click()
    Toolbar1.Enabled = True
    TabOne1.Enabled = True

fraconsdocref.Visible = False
End Sub


Private Sub cmdsalirseldoc_Click()
    Toolbar1.Enabled = True
    TabOne1.Enabled = True
    Fraseldoc.Visible = False
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
  If Val(TxtTipItem) > 0 Then
  If Val(Me.TxtTipDoc) = 7 Then ' SI ES NOTA DE CREDITO
  
  Else 'SI ES NOTA DE DEBITO
    
        Dim xCampos(3, 4) As String
    
        xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Unidad":       xCampos(1, 1) = "abrev":          xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
        xCampos(2, 0) = "Codigo":       xCampos(2, 1) = "codpro":         xCampos(2, 2) = "1500":    xCampos(2, 3) = "C"
    
        xform.SQLCad = "SELECT alm_inventario.*, mae_unidades.descripcion AS descuni,  mae_unidades.abrev " _
            & " FROM mae_unidades INNER JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
            & " WHERE alm_inventario.tippro = " & Val(TxtTipItem) & "  ORDER BY alm_inventario.descripcion "
    
        xform.Titulo = "Buscando Productos"
        xform.FormaBusca = CualquierParte
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"

        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        
        If xRs.State = 1 Then
            Fg1.TextMatrix(Fg1.Row, 1) = xRs("descripcion") & "de " & fgdocsproc.TextMatrix(fgdocsproc.Rows - 1, 3) + " " + fgdocsproc.TextMatrix(fgdocsproc.Rows - 1, 4)
            Fg1.TextMatrix(Fg1.Row, 2) = xRs("abrev")
            
            Fg1.TextMatrix(Fg1.Row, 6) = xRs("id")
            Fg1.TextMatrix(Fg1.Row, 7) = xRs("idunimed")
            Fg1.TextMatrix(Fg1.Row, 8) = xRs("idcuentaven")
            Fg1.TextMatrix(Fg1.Row, 9) = xRs("idtipven")
        End If
     End If
   End If
   

  



  
  Set xform = Nothing
  Set xRs = Nothing
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    
    If Mostrando = True Then Exit Sub
    If Col = 3 Or Col = 4 Then
        Fg1.TextMatrix(Fg1.Row, 5) = Val(Fg1.TextMatrix(Fg1.Row, 3)) * Val(Fg1.TextMatrix(Fg1.Row, 4))
        Fg1.TextMatrix(Fg1.Row, 5) = Format(Fg1.TextMatrix(Fg1.Row, 5), "0.00")
        HallarTotal
    End If
    
    
End Sub

Sub HallarTotal()
    Dim A As Integer
    Dim totalafec As Double
    Dim totalinaf As Double
    
    
    
    txtinafecto.Text = "0.00"
    TxtIGV.Text = "0.00"
    txtisc = "0.00"
    TxtTotal.Text = "0.00"
    
    For A = 1 To Fg1.Rows - 1
        
            If Val(TxtTipDoc) = 2 Then 'SI ES RECIBO POR HONORARIOS
                totalafec = totalafec + Val(Fg1.TextMatrix(A, 5)) 'venta  gravada
            Else
                If Fg1.TextMatrix(A, 9) = "1" Then 'si es venta gravada
                    totalafec = totalafec + Val(Fg1.TextMatrix(A, 5)) 'venta  gravada
                Else
                    totalinaf = totalinaf + Val(Fg1.TextMatrix(A, 5)) 'venta no gravada
                End If
            End If
    Next A
    
    
    
        
            TxtTotal.Text = (totalafec * ((TasaImpuesto / 100) + 1)) + totalinaf
            If totalafec > 0 Then
                TxtIGV.Text = (totalafec * ((TasaImpuesto / 100) + 1)) - totalafec
            
            End If
            txtinafecto = totalinaf
        
        
        TxtBruto.Text = Format(totalafec, "0.00")
        txtinafecto.Text = Format(txtinafecto.Text, "0.00")
        TxtIGV.Text = Format(TxtIGV.Text, "0.00")
        TxtTotal.Text = Format(TxtTotal.Text, "0.00")
End Sub

Private Sub Fg1_Click()
'Form1.Show
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then Exit Sub
    If Fg1.Col = 2 Or Fg1.Col = 5 Then
        Fg1.Editable = flexEDNone
    Else
        Fg1.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
 Call CmdDelItem_Click
End If

If KeyCode = 45 Then
 
 Call CmdAddItem_Click
End If

End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    
    If Button = 2 Then PopupMenu menu1

End Sub

Private Sub Fgdocref_DblClick()
 CmdOkRef_Click
End Sub

Private Sub Fgdocref_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmdOkRef_Click
End If
End Sub

Private Sub fgdocsproc_DblClick()
Call cmdOKdocsproc_Click
End Sub

Private Sub fgdocsproc_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
 cmdEliminarOKdocsproc_Click
End If

End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Dim Rpta As Integer
        Dim Rst As New ADODB.Recordset
        
        
        Set Rst = Nothing
        
        RST_Busq RstVent, "SELECT vta_notascreabo.*, mae_cliente.nombre, [vta_notascreabo]![numser]+'-'+[vta_notascreabo]![numdoc] AS numerodoc, IIf(vta_notascreabo.Anulado = 0, 'Generado', 'Anulado') AS [EstadoVenta], " _
            & " mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, " _
            & " mae_moneda.descripcion AS descmon, mae_moneda.simbolo, mae_impuestos.idcuenvta, con_tc.impcom, mae_tipoproducto.descripcion AS desctipcom" _
            & " FROM (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN ((mae_documento LEFT JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) " _
            & " RIGHT JOIN (mae_condpago RIGHT JOIN (vta_notascreabo LEFT JOIN con_tc ON vta_notascreabo.fchdoc = con_tc.fecha) ON " _
            & " mae_condpago.id = vta_notascreabo.idconpag) ON mae_documento.id = vta_notascreabo.tipdoc) ON mae_moneda.id = vta_notascreabo.idmon) " _
            & " ON mae_cliente.id = vta_notascreabo.idcli) LEFT JOIN mae_tipoproducto ON vta_notascreabo.idtipo = mae_tipoproducto.id " _
            & " WHERE  (vta_notascreabo.tipdoc = 7 or vta_notascreabo.tipdoc = 8) and   (((vta_notascreabo.numreg) Like '" & Format(xMes, "00") & "%'))", xCon
        
        Set Dg1.DataSource = RstVent
        If RstVent.RecordCount = 0 Then
            Rpta = MsgBox("No se ha registrado ninguna Nota de Credito / Debito, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Unload Me
            End If
        Else
            Dg1.SetFocus
        End If
        
    End If
End Sub

Private Sub Form_Load()
    Fg1.ColWidth(6) = 0
    QueHace = 3
    TabOne1.CurrTab = 0
    SeEjecuto = False
    
    CaracteresNumericos = "0123456789." & Chr(8)
    Fg1.ColWidth(6) = 0
    Fg1.ColWidth(7) = 0
    Fg1.ColWidth(8) = 0
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    
    fgdocsproc.Rows = 1
    
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    
    
    
    TasaImpuesto = 19
    LblIgvTasa.Caption = "(" & Trim(Str(TasaImpuesto)) + " % " & ")"
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ColComboList(1) = ""
    
    'swguiafact = 0
    xAño = 2006
    xMes = 12
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando una compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub menu1_1_Click()
    CmdAddItem_Click
End Sub

Private Sub menu1_3_Click()
    CmdDelItem_Click
End Sub






Private Sub optanulacion_Click()
TxtMotivo = "POR ANULACION DE DOCS."
End Sub

Private Sub optdescuento_Click()
TxtMotivo = "POR DESCUENTO DE DOCS. "
End Sub

Private Sub Option2_Click()

End Sub

Private Sub optdevolucion_Click()
TxtMotivo = "POR DEVOLUCION DE DOCS. "
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        
        'Validamos si la cuadricula tiene datos
        
            If QueHace = 3 Then
                If RstVent.RecordCount = 0 Then
                    MsgBox "No existe información para visualizar", vbInformation, Me.Caption
                    Blanquea
                    Exit Sub
                Else
                    MuestraSegundoTab
                End If
            End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.index = 1 Then Nuevo
    
    If Button.index = 2 Then Modificar
    
    If Button.index = 3 Then Eliminar
    
    If Button.index = 5 Then
        If Grabar = True Then
            Cancelar
            RstVent.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.index = 6 Then Cancelar
    
    
    If Button.index = 11 Then CambiarMes
    
    If Button.index = 15 Then
        Set RstVent = Nothing
        Unload Me
    End If
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  If ButtonMenu.Parent.index = 2 Then
        
       'MODIFICACION DE NOTAS DE CREDITO Y DEBITO
        If ButtonMenu.index = 1 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "La cuadricula de datos esta vacia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            If RstVent("anulado") = -1 Then
                        MsgBox "No puede modificar la " & RstVent("nomdoc") & " anulada proceda a restaurarlo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                        Exit Sub
            Else
                        Modificar
            End If
        End If
        
        'RESTAURAR documentoS
        If ButtonMenu.index = 2 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "La cuadricula de datos esta vacia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
                If RstVent("anulado") = -1 Then
                    RestaurarFactura
                End If
        End If
    
    End If
  
  If ButtonMenu.Parent.index = 3 Then
        If ButtonMenu.index = 1 Then Anular
        If ButtonMenu.index = 2 Then Eliminar
        If ButtonMenu.index = 3 Then EmitirAnulada
    End If
End Sub

Private Sub TxtConPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        
        If NulosC(TxtConPag.Text) = "" Then Exit Sub
        Dim xRs1 As New ADODB.Recordset
        
        RST_Busq xRs1, "SELECT * FROM mae_condpago WHERE id = " & Val(TxtConPag.Text) & "", xCon
        
        If xRs1.RecordCount = 0 Then
            TxtConPag.Text = ""
            LblCondPag.Caption = ""
        Else
            LblCondPag.Caption = Trim(xRs1("descripcion"))
        End If
        Set xRs1 = Nothing
        TxtFchDoc.SetFocus
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtConPag_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCondicion_Click
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        
        If NulosC(TxtIdMon.Text) = "" Then Exit Sub
        Dim xRs1 As New ADODB.Recordset
        
        'buscamos el codigo de la moneda         digitada
        RST_Busq xRs1, "SELECT * FROM mae_moneda WHERE id = " & Val(TxtIdMon.Text) & "", xCon
        
        If xRs1.RecordCount = 0 Then
            TxtIdMon.Text = ""
            LblMoneda.Caption = ""
        Else
            LblMoneda.Caption = Trim(xRs1("descripcion"))
            
            If Trim(TxtIdMon.Text) = "1" Then
                LblTipCam.Visible = False
                LblTipoCambio.Visible = False
            Else
                If TxtFchDoc.Valor = "" Then
                    MsgBox "No ha especificado la fecha del documento, no se puede determinar " & Chr(13) _
                        & "la fecha del tipo de cambio para este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    
                    TxtIdMon.Text = ""
                    LblMoneda.Caption = ""
                    TxtFchDoc.SetFocus
                    Exit Sub
                End If
                LblTipCam.Visible = True
                LblTipoCambio.Visible = True
                Set xRs1 = Nothing
                Set xRs1 = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = CDATE('" & TxtFchDoc.Valor & "')", xCon)
                If xRs1.RecordCount = 1 Then
                    LblTipoCambio.Caption = Format(xRs1("impven"), "0.000")
                    ValTipCam = xRs1("impven")
                Else
                    LblTipoCambio.Caption = "0.00"
                    ValTipCam = 0
                    
                End If
            End If
        End If
        Set xRs1 = Nothing
        TxtTipDoc.SetFocus
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtMotivo_KeyPress(KeyAscii As Integer)
Dim X As Integer
Dim numfac As String
Dim Rpta As Integer
If KeyAscii = 13 Then

   
   With Fg1
    If .Rows - 1 = 0 Then
        MsgBox "Proceda Agregar Items Doc. Ref. y luego comentarios de referencia para  la Nota de Credito ", vbInformation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Desea Añadir al comentario los Nro de documentos respectivo?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        
            .AddItem ""
            For X = 1 To fgdocsproc.Rows - 1
              numfac = numfac + fgdocsproc.TextMatrix(X, 4) + ", "
            Next
           .TextMatrix(.Rows - 1, 1) = Trim(TxtMotivo) + numfac
        
    Else
         .AddItem ""
        .TextMatrix(.Rows - 1, 1) = Trim(TxtMotivo)
    End If
    End With
End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      
        If NulosC(TxtNumDoc.Text) <> "" Then
            TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
        End If
      TxtConPag.SetFocus
    End If
End Sub


Private Sub TxtNumRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        
        Dim xRs1 As New ADODB.Recordset
        RST_Busq xRs1, "SELECT * FROM mae_cliente WHERE numruc like '" & TxtNumRuc.Text & "%' ORDER BY numruc", xCon
        If xRs1.RecordCount <> 0 Then
            TxtNumRuc.Text = xRs1("numruc")
            LblNomCli.Caption = xRs1("nombre")
            LblIdCliente.Caption = xRs1("id")
        Else
            TxtNumRuc.Text = ""
            LblNomCli.Caption = ""
            LblIdCliente.Caption = ""
        End If
        Set xRs1 = Nothing
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumRuc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCli_Click
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    Dim Rstdoc As New ADODB.Recordset
    If KeyAscii = 13 Then
        
        If NulosC(TxtNumSer.Text) <> "" Then
            TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
        End If
   
        RST_Busq Rstdoc, " SELECT mae_documento.descripcion, mae_series.iddoc, mae_series.numser, mae_series.numdoc " & _
                         " FROM mae_documento INNER JOIN mae_series ON mae_documento.id = mae_series.iddoc where mae_series.iddoc = " & Val(TxtTipDoc) & " and mae_series.numser =" & Val(TxtNumSer) & "  ", xCon
        
        
        If Rstdoc.RecordCount = 0 Then
            MsgBox "Registre la serie en Mantenimiento de Series", vbInformation, xTitulo
            TxtNumSer = ""
            TxtNumDoc = ""
        Else
            TxtNumSer.Text = Format(Rstdoc("numser"), "0000")
            TxtNumDoc = HallaNumdocVenta(Val(TxtTipDoc.Text), NulosC(TxtNumSer.Text), xCon)
        End If
            
            TxtNumDoc.SetFocus
    End If
    
End Sub

Private Sub TxtNumSer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
CmdBusNumSer_Click

End If

End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        
        If NulosC(TxtTipDoc.Text) = "" Then Exit Sub
        
        If Val(TxtTipDoc) = 7 Or Val(TxtTipDoc) = 8 Then
        
            Dim xRs As New ADODB.Recordset
            Dim xRs2 As New ADODB.Recordset
            
            RST_Busq xRs, "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
            & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuenvta as cuentaimp " _
            & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuenvta = con_planctas.id) " _
            & " ON mae_documento.idimp = mae_impuestos.id WHERE mae_documento.id  = " & Val(TxtTipDoc.Text) & "", xCon
            
            If xRs.RecordCount = 0 Then
                TxtTipDoc.Text = ""
                LblNomDoc.Caption = ""
            Else
                TxtTipDoc.Text = xRs("id")
                LblNomDoc.Caption = xRs("descripcion")
                TasaImpuesto = NulosN(xRs("tasa"))
                xDescImp = xRs("descripcion")
                xIdCuenTasa = NulosN(xRs("cuentaimp"))
                                
                Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & Val(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & Val(TxtIdMon) & " and tipope = -1", xCon)
                If xRs2.RecordCount > 0 Then
                    xCuentaDoc = NulosN(xRs2("idcuen"))
                End If
                Set xRs2 = Nothing
                LblRotulo = Trim(NulosC(xRs("abreimp"))) + " (       )"
                LblIgvTasa.Caption = Trim(Str(TasaImpuesto)) + "%"
                TxtNumRuc.SetFocus
            End If
                
                If Val(TxtTipDoc) = 7 Then
                    fradocref.Enabled = True
                    
                Else
                    fradocref.Enabled = False
                    
                End If
                
                Set xRs = Nothing
        Else
              MsgBox "Ingrese el Codigo de la Nota de Credito o Debito", vbInformation, xTitulo
              TxtTipDoc = ""
              Exit Sub
        End If
        
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    
End Sub

Private Sub TxtTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDoc_Click
    End If
End Sub

Sub EmitirAnulada()
Dim rs As New ADODB.Recordset
    Toolbar1.Enabled = False
    TabOne1.Enabled = False
    
    Fraseldoc.Left = 4185
    Fraseldoc.Top = 1635
    Fraseldoc.Width = 4425
    Fraseldoc.Height = 1335
    
    RST_Busq rs, "SELECT * FROM mae_documento WHERE Id = 7 OR Id = 8 ORDER BY descripcion ", xCon
    
    cbonotascredeb.Clear
    Do While Not rs.EOF
        cbonotascredeb.AddItem rs!Descripcion & Space(100) & rs!id
        rs.MoveNext
    Loop
    
    If rs.RecordCount > 0 Then cbonotascredeb.ListIndex = 0
    Fraseldoc.Visible = True
    Set rs = Nothing
End Sub

Function Grabar() As Boolean

            
    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento de venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Function
    End If
    
    If TxtNumRuc.Text = "" Then
        MsgBox "No ha especificado cliente de la venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Function
    End If
    
    If TxtNumSer.Text = "" Or TxtNumDoc.Text = "" Then
        MsgBox "No ha especificado el numero del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        Exit Function
    End If
    
    If TxtFchDoc.Valor = "" Then
        MsgBox "No ha especificado la fecha de emision del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchDoc.SetFocus
        Exit Function
    End If
    
    If TxtFchVen.Valor = "" Then
        MsgBox "No ha especificado la fecha de vencimiento del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchVen.SetFocus
        Exit Function
    End If
    
    If TxtConPag.Text = "" Then
        MsgBox "No ha especificado la condicion de pago del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtConPag.SetFocus
        Exit Function
    End If
    
    If TxtIdMon.Text = "" Then
        MsgBox "No ha especificado la moneda del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado items para la venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    If Fg1.Rows - 1 >= 1 Then
    If Fg1.TextMatrix(1, 8) = "" Then
        MsgBox "No ha especificado codigo para item añadido ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    End If
    
    If QueHace = 1 Then 'Validamos si existe el numero del documento en modo adicion
    
    Dim RstCab As New ADODB.Recordset
    
    RST_Busq RstCab, " Select * from vta_notascreabo where Clng(tipdoc) =" & CLng(TxtTipDoc) & " and Clng(Numser) =" & CLng(TxtNumSer) & " and Clng(numdoc) = " & CLng(Me.TxtNumDoc) & " ", xCon
    
    If RstCab.RecordCount > 0 Then
        MsgBox "El Nro de documento ha sido registrado por otro usuario se grabara con otro numero", vbInformation, Me.Caption
        TxtNumDoc = HallaNumdocVenta(Val(TxtTipDoc), NulosC(TxtNumSer.Text), xCon)
    End If
    Set RstCab = Nothing
    End If
    
    Dim RstDeta2 As New ADODB.Recordset
    Dim RstActPro As New ADODB.Recordset
    
    Dim RstDet As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim xIdCuen As Integer
    Dim xTotal As Double
    Dim xidtipven As String 'Determina si la venta es de tipo exportacion
    Dim xNumAsiento As String
    
    Dim xId As Integer
    Dim A As Integer
    Dim X As Integer
    Dim P As Integer
    On Error GoTo LaCague
    
  '  swguiafact = 1
    xCon.BeginTrans
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("vta_notascreabo", xCon, "id")
        xNumAsiento = HallaNumAsiento(xMes)
        
        RST_Busq RstCab, "SELECT * FROM vta_notascreabo", xCon
        RST_Busq RstDet, "SELECT * FROM vta_notascreabodet", xCon
        RST_Busq RstDia, "SELECT * FROM con_diario", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstVent("id")
        RST_Busq RstCab, "SELECT * FROM vta_notascreabo WHERE id = " & xId & "", xCon
        
        'Eliminamos el detalle de la venta
            xCon.Execute "DELETE * FROM vta_notascreabodet WHERE idvta = " & xId & ""
        
        RST_Busq RstDet, "SELECT * FROM vta_notascreabodet", xCon
        
        RST_Busq RstDia, "SELECT * FROM con_diario WHERE idmes = " & Format(CDate(TxtFchDoc.Valor), "mm") & " AND " _
                         & " idlib = 2 AND idmov = " & xId & " And iddoc = " & Val(TxtTipDoc) & "", xCon
            
        If RstDia.RecordCount <> 0 Then
            xNumAsiento = RstDia("numasi")
        End If
        
        Set RstDia = Nothing
        
       'Eliminamos el asiento contable
        xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & Format(CDate(TxtFchDoc.Valor), "mm") & " AND " _
                     & " idlib = 2 AND idmov = " & xId & " AND Iddoc = " & Val(TxtTipDoc) & ""
            
        RST_Busq RstDia, "SELECT * FROM con_diario", xCon
    End If
    
    RstCab("idtipo") = Val(TxtTipItem.Text)
    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("idcli") = NulosN(LblIdCliente.Caption)
    RstCab("numser") = TxtNumSer.Text
    RstCab("numdoc") = TxtNumDoc.Text
    RstCab("fchdoc") = TxtFchDoc.Valor
    RstCab("fchven") = TxtFchVen.Valor
    RstCab("idconpag") = NulosN(TxtConPag.Text)
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    RstCab("impbru") = NulosN(TxtBruto.Text)
    RstCab("impinaf") = NulosN(txtinafecto.Text)
    RstCab("impigv") = NulosN(TxtIGV.Text)
    RstCab("impisc") = NulosN(txtisc.Text)
    RstCab("impotr") = 0  'NulosN(Txtotr..Text)
    RstCab("imptotdoc") = NulosN(TxtTotal.Text)
    RstCab("impsal") = NulosN(TxtTotal.Text)
    RstCab("numreg") = Trim(Str(xMes)) + xNumAsiento
    
    
    'Determinamos si es una exportacion
    For A = 1 To Fg1.Rows - 1
        xidtipven = Val(Fg1.TextMatrix(A, 9))
    Next A
    
    If xidtipven = 2 And Val(TxtIGV.Text) = 0 Then 'si esta venta exportacion
        RstCab("idtipven") = 2
    Else
        RstCab("idtipven") = 0 'en el cual puede ser venta afecta o inafecta para el registro de
                               'de ventas se valida por programa ver tabla mae_tipoventa
    End If
            
    RstCab.Update
    
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idvta") = xId
        RstDet("iditem") = Val(Fg1.TextMatrix(A, 6))
        
        'Si no tiene codigo este item grabamos su descripcion segun registre el usuario
        If Val(Fg1.TextMatrix(A, 6)) = 0 Then
            RstDet("Descripusu") = Trim(Fg1.TextMatrix(A, 1))
        End If
        RstDet("idunimed") = Val(Fg1.TextMatrix(A, 7))
        RstDet("preuni") = Val(Fg1.TextMatrix(A, 3))
        RstDet("canpro") = Val(Fg1.TextMatrix(A, 4))
        RstDet("imptot") = Val(Fg1.TextMatrix(A, 5))
        RstDet.Update
    Next A
   

'    If P = 200 Then
    
       'Grabamos el libro diario del movimiento
        RstDia.AddNew
       'Grabamos el documento de venta en la tabla diario
        RstDia("año") = xAño
        RstDia("idmes") = xMes
        RstDia("idlib") = 2
        RstDia("idmov") = xId
        RstDia("iddoc") = Val(TxtTipDoc)
        RstDia("numasi") = xNumAsiento
        RstDia("tc") = ValTipCam
        RstDia("idcue") = xCuentaDoc
        
        If TxtIdMon.Text = "1" Then
            RstDia("imphabsol") = Val(TxtTotal.Text)
            RstDia("imphabdol") = 0
        Else
            RstDia("imphabsol") = Val(TxtTotal.Text) * Val(LblTipoCambio.Caption)
            RstDia("imphabdol") = Val(TxtTotal.Text)
        End If
        
        RstDia.Update
        
    'Grabamos el impuesto si la operacion esta afecta a el
    If Val(Me.TxtIGV) > 0 Then
        RstDia.AddNew
        RstDia("idmes") = xMes
        RstDia("año") = xAño
        RstDia("idlib") = 2
        RstDia("idmov") = xId
        RstDia("iddoc") = Val(TxtTipDoc)
        RstDia("numasi") = xNumAsiento
        RstDia("tc") = ValTipCam
        RstDia("idcue") = xIdCuenTasa
        If TxtIdMon.Text = "1" Then
            RstDia("impdebsol") = Val(TxtIGV.Text)
            RstDia("impdebdol") = 0
        Else
            RstDia("impdebsol") = Val(TxtIGV.Text) * Val(LblTipoCambio.Caption)
            RstDia("impdebdol") = Val(TxtIGV.Text)
        End If
        RstDia.Update
    End If
    
    
   
   '********Rutina para que extraer la base imponible sea afecta o inafecta
    
    
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(2, 3) As String
    Dim rstdocus As New ADODB.Recordset

            xCampos(0, 0) = "cuenta":     xCampos(0, 1) = "C":      xCampos(0, 2) = "12"
            xCampos(1, 0) = "Importe":    xCampos(1, 1) = "D":      xCampos(1, 2) = "2"
    
            Set rstdocus = xFun.CrearRstTMP(xCampos)
            rstdocus.Open


          For X = 1 To Fg1.Rows - 1
            
            If Trim(Fg1.TextMatrix(X, 8)) <> "" Then
                    xIdCuen = Trim(Fg1.TextMatrix(X, 8))
                    xTotal = Val(Fg1.TextMatrix(X, 5))
                    rstdocus.Find ("cuenta ='" & xIdCuen & "'")
                If rstdocus.EOF = True Then
                    rstdocus.AddNew
                    rstdocus("cuenta") = Trim(Fg1.TextMatrix(X, 8))
                    rstdocus("importe") = xTotal
                    rstdocus.Update
                Else
                    rstdocus("importe") = rstdocus("importe") + xTotal
                    rstdocus.Update
                End If
            End If
           
           Next X
           
           'Grabamos el diario
            If rstdocus.RecordCount > 0 Then
                rstdocus.MoveFirst
                Do While Not rstdocus.EOF
                    RstDia.AddNew
                    RstDia("año") = xAño
                    RstDia("idmes") = xMes               'LLAVE - CODIGO DEL MES
                    RstDia("idlib") = 2                  'LLAVE - CODIGO DEL LIBRO
                    RstDia("idmov") = xId                'LLAVE - CODIGO DEL MOVIMIENTO
                    RstDia("iddoc") = Val(TxtTipDoc)     'LLAVE - CODIGO DEL DOCUMENTO
                    RstDia("numasi") = xNumAsiento       'LLAVE - NUMERO DE ASIENTO
                    
                    RstDia("tc") = ValTipCam
                    RstDia("idcue") = rstdocus("cuenta")
                    If TxtIdMon.Text = "1" Then
                        RstDia("impdebsol") = rstdocus("importe")
                        RstDia("impdebdol") = 0
                    Else
                        RstDia("impdebsol") = rstdocus("importe") * Val(LblTipoCambio.Caption)
                        RstDia("impdebdol") = rstdocus("importe")
                    End If
                    RstDia.Update
                    rstdocus.MoveNext
                Loop
            End If
    
    
   'Actualizamos en el campo Idnumref de la tabla Ventas con el valor del campo id de la tabla vta_notascreabo
    For X = 1 To Me.fgdocsproc.Rows - 1
            xCon.Execute " UPDATE vta_ventas SET Vta_Ventas.idnumref = " & xId & " WHERE vta_ventas.id = " & Val(fgdocsproc.TextMatrix(X, 1)) & ""
    Next
 '   End If
    fgdocsproc.Rows = 1
    
    If QueHace = 1 Then
    'Grabamos ó actualizamos el ultimo numero del documento
    Call ActualizaNroDocumento(Val(TxtNumDoc), Val(TxtTipDoc), Val(TxtNumSer))
    End If
    
    xCon.CommitTrans
    MsgBox "La nota de credito y/o debito se registro con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDia = Nothing
    Grabar = True
    Exit Function
    
LaCague:
    'Resume
    xCon.RollbackTrans
    Set rstdocus = Nothing
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDia = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Function HallaNumAsiento(Mes As Integer) As String
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT con_diario.idmes, con_diario.idlib, con_diario.numasi From con_diario " _
        & " WHERE (((con_diario.idmes)=" & Mes & ") AND ((con_diario.idlib)=2)) ORDER BY numasi", xCon
    
    If Rst.RecordCount = 0 Then
        HallaNumAsiento = "0001"
    Else
        Rst.MoveLast
        HallaNumAsiento = Format(Val(Rst("numasi")) + 1, "0000")
    End If
    Exit Function
End Function

Private Sub TxtTipItem_KeyPress(KeyAscii As Integer)
    Dim RstTmp As New ADODB.Recordset
    If KeyAscii = 13 Then
        If NulosC(TxtTipItem.Text) <> "" Then
            Set RstTmp = BuscaConCriterio("SELECT * FROM mae_tipoproducto WHERE id = " & Val(TxtTipItem.Text) & "", xCon)
            If RstTmp.RecordCount > 0 Then
               LblTipoItem.Caption = RstTmp("descripcion")
            Else
                TxtTipItem.Text = ""
                LblTipoItem.Caption = ""
            End If
        End If
        TxtIdMon.SetFocus
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    Set RstTmp = Nothing

End Sub

Private Sub TxtTipItem_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipItem_Click
    End If

End Sub

