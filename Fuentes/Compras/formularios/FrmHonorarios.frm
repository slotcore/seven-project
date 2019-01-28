VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmHonorarios 
   Caption         =   "Compras - Ingreso de Honorarios"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdApertura 
      Caption         =   "&Apertura"
      Height          =   315
      Left            =   10575
      TabIndex        =   107
      Top             =   375
      Visible         =   0   'False
      Width           =   1320
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8010
      Top             =   15
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
            Picture         =   "FrmHonorarios.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame11 
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   2700
      Left            =   11040
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   7320
      Begin VB.CommandButton CmdAcep 
         Caption         =   "&Aceptar"
         Height          =   405
         Left            =   2985
         TabIndex        =   26
         Top             =   2220
         Width           =   1305
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg4 
         Height          =   1710
         Left            =   195
         TabIndex        =   27
         Top             =   465
         Width           =   6900
         _cx             =   12171
         _cy             =   3016
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
         Rows            =   50
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmHonorarios.frx":2B10
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
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos Adjuntos"
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
         Left            =   180
         TabIndex        =   28
         Top             =   135
         Width           =   1860
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   7305
         X2              =   7305
         Y1              =   15
         Y2              =   2685
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   15
         Y2              =   2670
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   7290
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   7305
         Y1              =   2685
         Y2              =   2685
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00400000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   330
         Left            =   45
         Top             =   60
         Width           =   7230
      End
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   3645
      Left            =   11145
      TabIndex        =   16
      Top             =   1830
      Visible         =   0   'False
      Width           =   8610
      Begin VB.TextBox TxtTotImp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7305
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "TxtTotImp"
         Top             =   2670
         Width           =   960
      End
      Begin VB.TextBox TxtTotPor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6330
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "TxtTotPor"
         Top             =   2670
         Width           =   975
      End
      Begin VB.CommandButton CmdDelCenCos 
         Caption         =   "&Eliminar C.C."
         Height          =   390
         Left            =   2865
         TabIndex        =   20
         Top             =   3120
         Width           =   1320
      End
      Begin VB.CommandButton CmdAddCenCos 
         Caption         =   "&Agregar C.C."
         Height          =   390
         Left            =   1500
         TabIndex        =   19
         Top             =   3120
         Width           =   1320
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   5790
         TabIndex        =   18
         Top             =   3120
         Width           =   1320
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   4425
         TabIndex        =   17
         Top             =   3120
         Width           =   1320
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg5 
         Height          =   2190
         Left            =   75
         TabIndex        =   23
         Top             =   465
         Width           =   8460
         _cx             =   14922
         _cy             =   3863
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmHonorarios.frx":2BED
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detallar Centro de Costos"
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
         TabIndex        =   24
         Top             =   90
         Width           =   2190
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   8595
         Y1              =   3630
         Y2              =   3630
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   3615
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   8595
         X2              =   8595
         Y1              =   15
         Y2              =   3645
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   8595
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   300
         Left            =   45
         Top             =   45
         Width           =   8520
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   15
      TabIndex        =   29
      Top             =   360
      Width           =   11880
      _cx             =   20955
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
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   12525
         TabIndex        =   35
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdBusTipDocRef 
            Height          =   240
            Left            =   2145
            Picture         =   "FrmHonorarios.frx":2CA8
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   2370
            Width           =   240
         End
         Begin VB.CommandButton CmdBusDocRef2 
            Height          =   240
            Left            =   8025
            Picture         =   "FrmHonorarios.frx":2DDA
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   2370
            Width           =   240
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   5
            Text            =   "TxtNumSer"
            Top             =   1710
            Width           =   915
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2760
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "TxtNumDoc"
            Top             =   1710
            Width           =   1440
         End
         Begin VB.Frame Frame10 
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
            Height          =   540
            Left            =   7215
            TabIndex        =   57
            Top             =   2640
            Width           =   1905
            Begin VB.CheckBox Check1 
               Caption         =   "Ingresar Neto"
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
               Left            =   210
               TabIndex        =   58
               Top             =   270
               Width           =   1500
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "[ Rta 4ta Cat. ]"
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
            Height          =   540
            Left            =   5385
            TabIndex        =   55
            Top             =   2640
            Width           =   1815
            Begin VB.CheckBox ChkImpRen4 
               Caption         =   "Aplicar Impuesto"
               Height          =   195
               Left            =   195
               TabIndex        =   56
               Top             =   270
               Width           =   1470
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "[ Opciones de Descuento]"
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
            Height          =   540
            Left            =   9135
            TabIndex        =   52
            Top             =   2640
            Width           =   2580
            Begin VB.OptionButton OptDes2 
               Caption         =   "Valor"
               Height          =   195
               Left            =   1590
               TabIndex        =   54
               Top             =   270
               Width           =   870
            End
            Begin VB.OptionButton OptDes1 
               Caption         =   "Porcentaje"
               Height          =   195
               Left            =   165
               TabIndex        =   53
               Top             =   270
               Width           =   1215
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "( Afecta :)"
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
            Height          =   600
            Left            =   1020
            TabIndex        =   49
            Top             =   4350
            Visible         =   0   'False
            Width           =   2805
            Begin VB.OptionButton OptSi 
               Caption         =   "Afecto"
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
               Height          =   195
               Left            =   105
               TabIndex        =   51
               Top             =   285
               Width           =   1125
            End
            Begin VB.OptionButton OptNo 
               Caption         =   "No Afecto"
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
               Height          =   195
               Left            =   1320
               TabIndex        =   50
               Top             =   270
               Width           =   1440
            End
         End
         Begin VB.CommandButton CmdBusTipoCompra 
            Height          =   240
            Left            =   2145
            Picture         =   "FrmHonorarios.frx":2F0C
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   1110
            Width           =   240
         End
         Begin VB.Frame Frame5 
            Height          =   495
            Left            =   9630
            TabIndex        =   46
            Top             =   210
            Width           =   2115
            Begin VB.Label LblPeriodo2 
               Alignment       =   2  'Center
               Caption         =   "LblPeriodo2"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   120
               TabIndex        =   47
               Top             =   150
               Width           =   1860
            End
         End
         Begin VB.CommandButton CmdBusCondicion 
            Height          =   240
            Left            =   6735
            Picture         =   "FrmHonorarios.frx":303E
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   795
            Width           =   240
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   2145
            Picture         =   "FrmHonorarios.frx":3170
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   1440
            Width           =   240
         End
         Begin VB.Frame Frame9 
            Caption         =   "[ Opciones de Compra ]"
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
            ForeColor       =   &H00400000&
            Height          =   540
            Left            =   120
            TabIndex        =   39
            Top             =   2640
            Width           =   5250
            Begin VB.OptionButton OptOpera3 
               Caption         =   "Doc. Entrada"
               Height          =   195
               Left            =   1110
               TabIndex        =   43
               ToolTipText     =   "Documentos de Entrada"
               Top             =   270
               Width           =   1260
            End
            Begin VB.OptionButton OptOpera1 
               Caption         =   "Normal"
               Height          =   195
               Left            =   105
               TabIndex        =   42
               ToolTipText     =   "Operacion Normal"
               Top             =   270
               Width           =   825
            End
            Begin VB.CommandButton CmdCargaDoc 
               Caption         =   "Adicionar"
               Height          =   300
               Left            =   4080
               TabIndex        =   41
               Top             =   165
               Width           =   1095
            End
            Begin VB.OptionButton OptOpera2 
               Caption         =   "Ord. de Compra"
               Height          =   195
               Left            =   2535
               TabIndex        =   40
               ToolTipText     =   "Orden de Compra"
               Top             =   270
               Width           =   1410
            End
         End
         Begin VB.TextBox TxtGlosa 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "TxtGlosa"
            Top             =   2025
            Width           =   10050
         End
         Begin VB.CommandButton CmdBusProv 
            Height          =   240
            Left            =   3150
            Picture         =   "FrmHonorarios.frx":32A2
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   480
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   6735
            Picture         =   "FrmHonorarios.frx":33D4
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1425
            Width           =   240
         End
         Begin VB.TextBox TxtDocRef 
            Height          =   300
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   7
            Text            =   "TxtDocRef"
            Top             =   1710
            Width           =   2025
         End
         Begin VB.CommandButton CmdBusDocRef 
            Height          =   240
            Left            =   8010
            Picture         =   "FrmHonorarios.frx":3506
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1740
            Width           =   240
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2430
            Left            =   105
            TabIndex        =   12
            Top             =   3195
            Width           =   11610
            _cx             =   20479
            _cy             =   4286
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
            Rows            =   20
            Cols            =   17
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmHonorarios.frx":3638
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
            Left            =   1665
            TabIndex        =   1
            Top             =   765
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
            Valor           =   "22/05/2008"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchVen 
            Height          =   300
            Left            =   10485
            TabIndex        =   3
            Top             =   750
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
            Valor           =   "22/05/2008"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchPago 
            Height          =   300
            Left            =   10485
            TabIndex        =   8
            Top             =   1710
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
            Valor           =   "22/05/2008"
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   4
            Text            =   "TxtIdMon"
            Top             =   1395
            Width           =   750
         End
         Begin VB.TextBox TxtConPag 
            Height          =   300
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   2
            Text            =   "TxtConPag"
            Top             =   765
            Width           =   750
         End
         Begin VB.TextBox TxtTipCom 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   103
            Text            =   "TxtTipCom"
            Top             =   1080
            Width           =   750
         End
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   0
            Text            =   "TxtNumRuc"
            Top             =   450
            Width           =   1770
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   104
            Text            =   "TxtTipDoc"
            Top             =   1395
            Width           =   750
         End
         Begin VB.Frame Frame4 
            Height          =   1215
            Left            =   105
            TabIndex        =   59
            Top             =   5580
            Width           =   11610
            Begin VB.CommandButton CmdVerAsiento 
               Caption         =   "&Ver Asiento Contable"
               Height          =   300
               Left            =   4980
               TabIndex        =   108
               Top             =   855
               Width           =   4020
            End
            Begin VB.CommandButton CmdSeleccionar 
               Caption         =   "Seleccionar Item"
               Height          =   345
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   66
               ToolTipText     =   "Seleccionar Items "
               Top             =   660
               Width           =   1395
            End
            Begin VB.CommandButton CmdDetCenCos 
               Caption         =   "Centro de Costo"
               Height          =   345
               Left            =   1650
               Style           =   1  'Graphical
               TabIndex        =   65
               ToolTipText     =   "Centro de Costos"
               Top             =   660
               Width           =   1395
            End
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "Agregar Item"
               Enabled         =   0   'False
               Height          =   345
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   64
               ToolTipText     =   "Agregar Item"
               Top             =   300
               Width           =   1395
            End
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "Eliminar Item"
               Enabled         =   0   'False
               Height          =   345
               Left            =   1650
               Style           =   1  'Graphical
               TabIndex        =   63
               ToolTipText     =   "Eliminar Item"
               Top             =   300
               Width           =   1395
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
               Left            =   10215
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   13
               TabStop         =   0   'False
               Text            =   "TxtBruto"
               Top             =   180
               Width           =   1230
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
               Left            =   10215
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   14
               TabStop         =   0   'False
               Text            =   "TxtIGV"
               Top             =   495
               Width           =   1230
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
               Left            =   10215
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   15
               TabStop         =   0   'False
               Text            =   "TxtTotal"
               Top             =   810
               Width           =   1230
            End
            Begin VB.CheckBox ChkAjusta 
               Caption         =   "Ajustar Totales"
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
               Left            =   6990
               TabIndex        =   62
               Top             =   570
               Visible         =   0   'False
               Width           =   1605
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Ver His. Precios"
               Height          =   345
               Left            =   3075
               Style           =   1  'Graphical
               TabIndex        =   61
               ToolTipText     =   "Historico de Precios"
               Top             =   300
               Width           =   1395
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Presupuesto"
               Height          =   345
               Left            =   3075
               Style           =   1  'Graphical
               TabIndex        =   60
               ToolTipText     =   "Presupuesto"
               Top             =   660
               Width           =   1395
            End
            Begin VB.Label lblrotulo 
               Caption         =   "lblrotulo"
               Height          =   240
               Left            =   4980
               TabIndex        =   96
               Top             =   510
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000005&
               Index           =   1
               X1              =   4785
               X2              =   4785
               Y1              =   90
               Y2              =   1200
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000003&
               Index           =   0
               X1              =   4770
               X2              =   4770
               Y1              =   120
               Y2              =   1185
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Imponible"
               Height          =   195
               Index           =   0
               Left            =   9465
               TabIndex        =   72
               Top             =   225
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total"
               Height          =   195
               Index           =   2
               Left            =   9780
               TabIndex        =   71
               Top             =   855
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Impuesto"
               Height          =   195
               Index           =   6
               Left            =   9495
               TabIndex        =   70
               Top             =   555
               Width           =   645
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   6435
               TabIndex        =   69
               Top             =   510
               Width           =   2565
            End
            Begin VB.Label LblIgvTasa 
               Alignment       =   2  'Center
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
               Height          =   225
               Left            =   7875
               TabIndex        =   68
               Top             =   255
               Width           =   1110
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Tasa Renta 4ta"
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
               Left            =   6405
               TabIndex        =   67
               Top             =   255
               Width           =   1335
            End
         End
         Begin VB.TextBox TxtDocRef2 
            Height          =   300
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   11
            Text            =   "TxtDocRef2"
            Top             =   2340
            Width           =   2025
         End
         Begin VB.TextBox TxtTipDocRef 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   10
            Text            =   "Txt"
            Top             =   2340
            Width           =   750
         End
         Begin VB.Label lblReg 
            Caption         =   "lblReg"
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
            Height          =   270
            Left            =   9555
            TabIndex        =   105
            Top             =   1095
            Width           =   2190
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tip de Doc. Ref."
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   102
            Top             =   2385
            Width           =   1185
         End
         Begin VB.Label LblTipDocref 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipDocref"
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
            Left            =   2430
            TabIndex        =   101
            Top             =   2340
            Width           =   2325
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Doc. Referencia"
            Height          =   195
            Index           =   13
            Left            =   4815
            TabIndex        =   100
            Top             =   2385
            Width           =   1395
         End
         Begin VB.Label LblIdDocRef2 
            AutoSize        =   -1  'True
            Caption         =   "LblIdDocRef2"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8385
            TabIndex        =   99
            Top             =   2385
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label LblIdAlmacen 
            AutoSize        =   -1  'True
            Caption         =   "LblIdAlmacen"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   2655
            TabIndex        =   95
            Top             =   270
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label LblIdCenCos 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCenCos"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   4050
            TabIndex        =   94
            Top             =   285
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Pago"
            Height          =   195
            Index           =   8
            Left            =   9285
            TabIndex        =   93
            Top             =   1755
            Width           =   1095
         End
         Begin VB.Label LblTipoCompra 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipoCompra"
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
            Left            =   2430
            TabIndex        =   92
            Top             =   1080
            Width           =   2325
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Item"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   91
            Top             =   1155
            Width           =   660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Venc."
            Height          =   195
            Index           =   3
            Left            =   9600
            TabIndex        =   90
            Top             =   810
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emisión"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   89
            Top             =   810
            Width           =   1260
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "T.C."
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
            Left            =   7635
            TabIndex        =   88
            Top             =   1125
            Visible         =   0   'False
            Width           =   375
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
            Height          =   270
            Left            =   8070
            TabIndex        =   87
            Top             =   1095
            Visible         =   0   'False
            Width           =   1260
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
            Left            =   7005
            TabIndex        =   86
            Top             =   765
            Width           =   2325
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
            Left            =   2430
            TabIndex        =   85
            Top             =   1395
            Width           =   2325
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Honorarios"
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
            Left            =   90
            TabIndex        =   84
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Documento"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   83
            Top             =   1755
            Width           =   1275
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2610
            Top             =   1830
            Width           =   105
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1110
            TabIndex        =   82
            Top             =   270
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Condición de Pago"
            Height          =   195
            Index           =   4
            Left            =   4860
            TabIndex        =   81
            Top             =   810
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   80
            Top             =   1455
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   79
            Top             =   2055
            Width           =   405
         End
         Begin VB.Label LblNomPro 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomPro"
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
            Left            =   3420
            TabIndex        =   78
            Top             =   450
            Width           =   5910
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Prestador de Servicio"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   77
            Top             =   480
            Width           =   1515
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
            Left            =   7020
            TabIndex        =   76
            Top             =   1395
            Width           =   2325
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Documento"
            Height          =   195
            Index           =   1
            Left            =   4800
            TabIndex        =   75
            Top             =   1455
            Width           =   1410
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Referente al Documento"
            Height          =   195
            Index           =   9
            Left            =   4470
            TabIndex        =   74
            Top             =   1755
            Width           =   1740
         End
         Begin VB.Label LblIdDocRef 
            AutoSize        =   -1  'True
            Caption         =   "LblIdDocRef"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8310
            TabIndex        =   73
            Top             =   1755
            Visible         =   0   'False
            Width           =   900
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   45
         TabIndex        =   30
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6480
            Left            =   30
            TabIndex        =   31
            Top             =   300
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11430
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
            Columns(1).Caption=   "Nº Reg."
            Columns(1).DataField=   "numreg1"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "T.D."
            Columns(2).DataField=   "abrev"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Nº Documento"
            Columns(3).DataField=   "numerodoc"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Emi"
            Columns(4).DataField=   "fchdoc1"
            Columns(4).NumberFormat=   "Short Date"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fch. Ven."
            Columns(5).DataField=   "fchven1"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Proveedor"
            Columns(6).DataField=   "nombre"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "M"
            Columns(7).DataField=   "simbolo"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "T.C."
            Columns(8).DataField=   "impven1"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Importe"
            Columns(9).DataField=   "impbru1"
            Columns(9).NumberFormat=   "0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Reten"
            Columns(10).DataField=   "impigv1"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Total"
            Columns(11).DataField=   "imptot1"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "Saldo"
            Columns(12).DataField=   "impsal1"
            Columns(12).NumberFormat=   "0.00"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   13
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   397
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=13"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1535"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1455"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=900"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=820"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2514"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2434"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1561"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1482"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1535"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1455"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=4419"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=4339"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=767"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=688"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=513"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=953"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=873"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=1455"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1376"
            Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(62)=   "Column(10).Width=1164"
            Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=1085"
            Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=514"
            Splits(0)._ColumnProps(67)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(68)=   "Column(11).Width=1429"
            Splits(0)._ColumnProps(69)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(70)=   "Column(11)._WidthInPix=1349"
            Splits(0)._ColumnProps(71)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(72)=   "Column(11)._ColStyle=514"
            Splits(0)._ColumnProps(73)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(74)=   "Column(12).Width=1561"
            Splits(0)._ColumnProps(75)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(76)=   "Column(12)._WidthInPix=1482"
            Splits(0)._ColumnProps(77)=   "Column(12)._EditAlways=0"
            Splits(0)._ColumnProps(78)=   "Column(12)._ColStyle=514"
            Splits(0)._ColumnProps(79)=   "Column(12).Order=13"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=74,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=78,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=86,.parent=13,.alignment=1"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=83,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=84,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=85,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13,.alignment=1"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=63,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=64,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=65,.parent=17"
            _StyleDefs(88)  =   "Named:id=33:Normal"
            _StyleDefs(89)  =   ":id=33,.parent=0"
            _StyleDefs(90)  =   "Named:id=34:Heading"
            _StyleDefs(91)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(92)  =   ":id=34,.wraptext=-1"
            _StyleDefs(93)  =   "Named:id=35:Footing"
            _StyleDefs(94)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(95)  =   "Named:id=36:Selected"
            _StyleDefs(96)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(97)  =   "Named:id=37:Caption"
            _StyleDefs(98)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(99)  =   "Named:id=38:HighlightRow"
            _StyleDefs(100) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(101) =   "Named:id=39:EvenRow"
            _StyleDefs(102) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(103) =   "Named:id=40:OddRow"
            _StyleDefs(104) =   ":id=40,.parent=33"
            _StyleDefs(105) =   "Named:id=41:RecordSelector"
            _StyleDefs(106) =   ":id=41,.parent=34"
            _StyleDefs(107) =   "Named:id=42:FilterBar"
            _StyleDefs(108) =   ":id=42,.parent=33"
         End
         Begin VB.Label LblPeriodo 
            Alignment       =   2  'Center
            Caption         =   "LblPeriodo"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   9810
            TabIndex        =   34
            Top             =   0
            Width           =   1860
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Honorarios"
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
            Left            =   105
            TabIndex        =   33
            Top             =   30
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
            TabIndex        =   32
            Top             =   30
            Width           =   765
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   106
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   609
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
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
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
   Begin VB.Menu menu1 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu1_1_1 
         Caption         =   "Agregar Item"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar Item"
      End
      Begin VB.Menu menu1_4 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_5 
         Caption         =   "Ver Historico de Precios"
      End
   End
   Begin VB.Menu opciones 
      Caption         =   "Opciones"
      Visible         =   0   'False
      Begin VB.Menu opciones_1 
         Caption         =   "Agregar Documentos de Entrada"
      End
      Begin VB.Menu opciones_2 
         Caption         =   "Agregar Documentos de Entrada Registrados"
      End
   End
End
Attribute VB_Name = "FrmHonorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMHONORARIOS.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO DONDE SE REGISTRAN LOS RECIBOS POR HONORARIOS, PERMITIENDO DETALLAR
'*                    LOS ITEMS DE LA COMPRA Y SU RESPECTIVO CENTRO DE COSTOS, ASI MISMO SE GENERA EL
'                     PROCESO CONTABLE PARA LA COMPRA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 17/09/09
'* VERSION          : 1.0
'*****************************************************************************************************

Option Explicit
Dim RstComp As New ADODB.Recordset     ' RECORDSET PRINCIPAL QUE ALMACENARA TODOS LOS REGISTROS
Dim QueHace As Integer                 ' VARIABLE QUE INDICA EN QUE ESTADO SE ENCUENTA EL FORMULARIO 1 = NUVEO; 2  MODIFIVA; 3 = SOLO LECTURA
Dim TasaImpuesto As Double             ' ALAMCENA LA TASA DEL IMPUESTO
Dim CaracteresNumericos As String      ' ALMACENA LOS CARACTERES NUMERICOS QUE SE PERMITIRAN INGRESAR EN LOS CONTROLES TextBox
Dim CaracteresNumericos2 As String     ' ALMACENA LOS CARACTERES NUMERICOS QUE SE PERMITIRAN INGRESAR EN ALGUNOS CONTROLES TextBox
Dim SeEjecuto As Boolean               ' VARIABLE QUE VALIDA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim ValTipCam As Double                ' ALMACENA EL VALOR DEL TIPO DE CAMBIO
Dim xDescImp As String                 ' ALMACENA LA DESCRIPCION DEL IMPUESTO
Dim xIdCuenTasa As Integer             ' codigo de la cuenta contable del impuesto
Dim xCuentaDoc As Integer              ' codigo de la cuenta contable del documento
Dim Mostrando As Boolean               ' INFORMA SI SE ESTA LLENANDO INFORMACION EN UN CONTROL FlexGrid
Dim RstTmp As New ADODB.Recordset      ' RECORDSET TEMPORAL UTILIZADO PARA CARGAR ALGUNAS TABLAS
Dim xFchFin, xFchIni, xFechaMes As String
Dim RstTempISC As New ADODB.Recordset  ' RECORSET TEMPORAL QUE ALMACENARA INFORMACION DEL IMPUESTO SELECTIVO AL CONSUMO
Dim AgePer As Boolean                  ' INDICA QUE ES UN AGENTE DE PERCEPCION
Dim AgeRet As Boolean                  ' INDICA QUE ES UN AGENTE DE RETENCION
Dim DetCenCos As Boolean               ' especifica si se va a detallar el centro de costos
Dim CodSunatDoc As String              ' especifica el codigo de la sunat del documento
Dim xPorIgv  As Double                 ' ESPECIFICA EL PORCENTAJE DEL IGV
Dim xHorIni As Date                    ' ALMACENA LA HORA DE INICIO

Dim fOrdenLista As Boolean             ' especfica el orden de la lista de la consulta
Dim mMesActivo As Integer              ' indica el mes activo
Dim fCierrePeriodo As Boolean          ' indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim mIdRegistro&                           ' --identificador del registro
'Para mover el frame
Dim OrigFX As Long
Dim OrigFY As Long


'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'* OBSERVACIONES    : NO VALIDA QUE EL REGISTRO ESTE VINCULADOS CON OTRAS OPERACIONES, DEBERIA DE
'*                    VALIDAR SI EL REGISTRO NO ESTA VINCULADO CON CAJA Y BANCOS, REVISAR ES URGENTE
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    TabOne1.CurrTab = 0
    If RstComp.State = 0 Then Exit Sub
    If RstComp.RecordCount = 0 Then
        MsgBox "No hay Registro de Honorario para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    '**********************************************************************************************************************
    '--evaluar si el registro de honorario si esta vinculado con modulo tesoreria
    Dim nSQL As String
    Dim Rst As New ADODB.Recordset
    Dim xId&
    
    xId = RstComp("id")
    
    nSQL = "SELECT Left(tes_caja.[numreg], 2) & '01' & Right(tes_caja.[numreg],4) AS registro   " _
        + vbCr + " FROM tes_caja INNER JOIN tes_cajadestinodet ON tes_caja.id = tes_cajadestinodet.idtes " _
        + vbCr + " WHERE (((tes_cajadestinodet.iddoc)=" & xId & ") AND ((tes_cajadestinodet.idmod)=9) AND ((tes_caja.tipmov)=2));"
    RST_Busq Rst, nSQL, xCon
    If Rst.RecordCount <> 0 Then
        MsgBox "El registro de Honorario está vinculado con: " + vbCr + "Módulo: Tesoreria - Egresos" & vbCr & "Nº. Registro: " & NulosC(Rst("registro")) & vbCr & "Si desea continuar, Elimine primero el Registro " & NulosC(Rst("registro")) & " del módulo de Tesoreria - Egresos", vbInformation, xTitulo
        Set Rst = Nothing
        Exit Sub
    End If
    Set Rst = Nothing
    '**********************************************************************************************************************
    
    Rpta = MsgBox("¿Esta seguro de eliminar el honorario seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        If RstComp("tipcom") = 3 Then
            ' actualizamos orden de compra
            xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idfac = 0 WHERE (((com_ordencompra.idfac)=" & xId & "))"
        End If
        '--Eliminar referencia en transferencia de documento si lo hubiera, asimismo cambiar estado preparado para transferir
        xCon.Execute "UPDATE tra_documento SET tra_documento.idmod = 0, tra_documento.iddoc = 0, tra_documento.estado = 0 WHERE (((tra_documento.idmod)=9) AND ((tra_documento.iddoc)=" & xId & ")) "
       
        ' ELIMINAMOS LOS REGISTROS DEL DIARIO
        xCon.Execute "DELETE * FROM con_diario WHERE idmov = " & xId & " AND idlib = 40"
        ' ELIMINAMOS EL REGISTRO
        xCon.Execute "DELETE * FROM com_honorarios WHERE id = " & xId & ""
        xCon.Execute "DELETE * FROM com_honorariosdet WHERE idhon = " & xId & ""
       
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & IdMenuActivo
        
        MsgBox "El Honorario se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstComp.Requery
        Dg1.Refresh
        
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LA BARRA DE HERRAMIENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
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

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA LOS PROCESOS DE AGREGAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    Bloquea
    Fg1.ColComboList(1) = ""
    Label5.Caption = "Detalle de Compra"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.Editable = flexEDNone
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Label5.Caption = "Agregando Honorarios"
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    OptSi.Value = True
    Fg1.Rows = 1
    Fg5.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
    OptDes1.Value = True
    OptDes1_Click
    OptOpera1.Value = True
    If xOrigen = 1 Then
        CargarValoresDefecto
    End If
    TxtTipCom.Text = "5"
    TxtTipCom_Validate True
    
    TxtIdMon.Text = "1"
    TxtIdMon_Validate True
    
    TxtTipDoc.Text = "2"
    TxtTipDoc_Validate True
    xHorIni = Time
    TxtNumRuc.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : CargarValoresDefecto
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA VALORES POR DEFECTO PARA EL FORMULARIO CUANDO SE ESTE AGREGANDO UN NUEVO
'*                    REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargarValoresDefecto()
    TxtFchDoc.Valor = Date
    TxtTipCom.Text = "1"
    TxtTipCom_Validate True
    TxtIdMon.Text = 1
    TxtIdMon_Validate True
    TxtTipDoc.Text = "1"
    TxtTipDoc_Validate True
    TxtConPag.Text = "1"
    TxtConPag_Validate True
    TxtFchVen.Valor = Date
    OptOpera1.Value = True
    OptOpera1_Click
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    MuestraSegundoTab
    Label5.Caption = "Modificando Honorario"
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    xHorIni = Time
    TxtFchDoc.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LOS DATOS DEL REGISTRO EN LA PESTAÑA CONSULTA DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Blanquea
    Dim xRs As New ADODB.Recordset
    If RstComp.EOF = True Or RstComp.BOF = True Or RstComp.RecordCount = 0 Then Exit Sub
    lblReg.Caption = "Nº Reg. " & NulosC(RstComp("numreg1"))
    
    TxtTipCom.Text = NulosN(RstComp("idtipo"))
    TxtTipDoc.Text = NulosN(RstComp("tipdoc"))
    TxtNumRuc.Text = NulosC(RstComp("numruc"))
    TxtNumSer.Text = NulosC(RstComp("numser"))
    TxtNumDoc.Text = NulosC(RstComp("numdoc"))
    If IsDate(RstComp("fchdoc")) = True Then TxtFchDoc.Valor = RstComp("fchdoc")
    If IsDate(RstComp("fchven")) = True Then TxtFchVen.Valor = RstComp("fchven")
    If IsDate(RstComp("fchpag")) = True Then TxtFchPago.Valor = RstComp("fchpag")
    
    TxtConPag.Text = RstComp("idconpag")
    TxtIdMon.Text = RstComp("idmon")
    TxtGlosa.Text = NulosC(RstComp("glosa"))
    
    LblTipoCompra.Caption = RstComp("desctipcom")
    LblNomDoc.Caption = RstComp("nomdoc")
    LblNomPro.Caption = RstComp("nombre")
    LblCondPag.Caption = NulosC(RstComp("desccond"))
    TxtNumRuc.Text = NulosC(RstComp("numruc"))
    LblMoneda.Caption = NulosC(RstComp("descmon"))
    LblIdProveedor.Caption = RstComp("idpro")
    LblIdAlmacen.Caption = NulosN(RstComp("idalm"))
    
    If NulosN(TxtTipDoc.Text) = 7 Then
        Label3(9).Visible = True
        TxtDocRef.Visible = True
        CmdBusDocRef.Visible = True
    Else
        Label3(9).Visible = False
        TxtDocRef.Visible = False
        CmdBusDocRef.Visible = False
    End If
        
    ' mostramos el documento de referencia de la compra
    Dim Rst As New ADODB.Recordset
    
    If NulosN(RstComp("idtipdocref")) <> 0 Then
        TxtTipDocRef.Text = NulosN(RstComp("idtipdocref"))
    Else
        TxtTipDocRef.Text = ""
    End If
    TxtTipDocRef_Validate False
    LblIdDocRef2.Caption = NulosN(RstComp("iddocref2"))
    
    ' SI EL DOCUMENTO DE REFERENCIA ES UNA ORDEN DE COMPRA.
    If NulosN(TxtTipDocRef.Text) = 1 Then
        RST_Busq Rst, "SELECT com_ordencompra.id, [com_ordencompra]![numser] & '-' & [com_ordencompra]![numdoc] AS numdoc From com_ordencompra " _
            & " WHERE (((com_ordencompra.id)=" & NulosN(LblIdDocRef2.Caption) & "))", xCon
    End If
    If NulosN(TxtTipDocRef.Text) = 2 Then
    End If
    If NulosN(TxtTipDocRef.Text) = 3 Then
    End If
    
    ' SI EL DOCUMENTO DE REFERENCIA ES UNA ORDEN DE DESPACHO
    If NulosN(TxtTipDocRef.Text) = 4 Then
        RST_Busq Rst, "SELECT var_ordendespacho.id, [var_ordendespacho]![año] & [var_ordendespacho]![idaduana] & [var_ordendespacho]![idregimen] & [var_ordendespacho]![numdoc] AS numdoc" _
            & " From var_ordendespacho WHERE (((var_ordendespacho.id)=" & NulosN(LblIdDocRef2.Caption) & "))", xCon
    End If
    
    If Rst.State <> 0 Then
        If Rst.RecordCount <> 0 Then
            TxtDocRef2.Text = Rst("numdoc")
            LblIdDocRef2.Caption = Rst("id")
        Else
            TxtDocRef2.Text = ""
            LblIdDocRef2.Caption = ""
        End If
    End If
    Set Rst = Nothing
    
    If RstComp("idmon") = 1 Then
        LblTipoCambio.Visible = False
    Else
        LblTipoCambio.Visible = True
        If mMesActivo = 0 Then
            LblTipoCambio.Caption = HallaTipoCambio("01/01/" + Trim(AnoTra), 2, Venta, xCon)
        Else
            LblTipoCambio.Caption = HallaTipoCambio(RstComp("fchdoc"), 2, Venta, xCon)
        End If
    End If
    
    ' mostramos el tipo de descuento que se le aplica a la compra
    Mostrando = True
    If RstComp("tipdes") = 1 Or NulosN(RstComp("tipdes")) = 0 Then
        OptDes1.Value = True
    End If
    
    If RstComp("tipdes") = 2 Then
        OptDes2.Value = True
    End If
    Mostrando = False
    
    ' Preguntamos en que contexto se realizo la compra
    If RstComp("tipcom") = 1 Then
        ' Se registro una compra normal
        OptOpera1.Value = True
        OptOpera1_Click
    End If
    
    If RstComp("tipcom") = 2 Then
        ' Se registro una compra con documento de ingreso
        OptOpera3.Value = True
        OptOpera3_Click
        CargarIngresoAlmacen RstComp("id")
    End If
    
    If RstComp("tipcom") = 3 Then
        ' Se registro una compra con orden de compra
        OptOpera2.Value = True
        OptOpera2_Click
    End If
    
    '--------------------------------------
    'revisar si este pedaso de codigo sirve
    If RstComp("afecto") = -1 Then
        OptSi.Value = True
        'OptSi_Click
    Else
        OptNo.Value = True
    End If
    '--------------------------------------
    
    TxtTipDoc_Validate True
    
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    
    ' cargamos la cuenta del igv
    RST_Busq RstDet, "SELECT mae_impuestos.idcuen, mae_impuestos.tasa, mae_documento.id FROM mae_documento LEFT JOIN mae_impuestos " _
        & " ON mae_documento.idimp = mae_impuestos.id WHERE (((mae_documento.id)=" & NulosN(TxtTipDoc.Text) & "))", xCon

    If RstDet.RecordCount <> 0 Then
        xIdCuenTasa = NulosN(RstDet("idcuen"))
        TasaImpuesto = NulosN(RstDet("tasa"))
    End If
    Set RstDet = Nothing
    
    Set RstDet = Nothing
    Mostrando = True
    Fg1.Rows = 1
    
    ' CARGAMOS EL DETALLE DEL DOCUMENTO
    RST_Busq RstDet, "SELECT com_honorariosdet.*, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idcuenta, " _
        & " alm_inventario.idtipcom, con_planctas.ctadesdeb, con_planctas.ctadeshab " _
        & " FROM con_planctas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN com_honorariosdet ON " _
        & " alm_inventario.id = com_honorariosdet.iditem) ON mae_unidades.id = alm_inventario.idunimed) ON " _
        & " con_planctas.id = alm_inventario.idcuenta WHERE (((com_honorariosdet.idhon)=" & RstComp("id") & "))", xCon
                       
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstDet("descripcion"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDet("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(NulosN(RstDet("canpro")), "0.0000")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosN(RstDet("preunibru")), "0.000000")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(RstDet("preunibruina")), "0.000000")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(RstDet("valdes")), "0.0000")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(RstDet("preuni")), "0.000000")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(RstDet("imptot")), "0.0000")
            
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosN(RstDet("iditem"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosN(RstDet("idunimed"))
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(RstDet("idcuenta"))
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosN(RstDet("idtipcom"))
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosN(RstDet("ctadesdeb"))
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosN(RstDet("ctadeshab"))
            
            RstDet.MoveNext
            If RstDet.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    ' BUSCAMOS SI ESTA AFECTO A ALGUN IMPUESTO
    BuscarImpuestos
    AgregarCentroCosto2 True, RstComp("id")
    
    If NulosN(TxtTipDoc.Text) = 2 Then
        'recibo por honorarios
        TxtBruto.Text = Format(RstComp("impbru"), FORMAT_MONTO)
        TxtIGV.Text = Format(RstComp("impigv"), FORMAT_MONTO)
        TxtTotal.Text = Format(RstComp("impbru") - RstComp("impigv"), FORMAT_MONTO)
    Else
        TxtBruto.Text = Format(RstComp("impbru"), FORMAT_MONTO)
        TxtIGV.Text = Format(RstComp("impigv"), FORMAT_MONTO)
        TxtTotal.Text = Format(RstComp("imptot"), FORMAT_MONTO)
    End If
    
    Set RstDet = Nothing
    Mostrando = False
    xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
    Set RstDet = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : CargarIngresoAlmacen
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DOCUMENTOS DE INGRESO QUE ESTEN RELACIONADOS CON LA COMPRA
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    IdCompra  |  INTEGER    |  ESPECIFICA EL ID DE LA COMPRA
'* Devuelve         :
'*****************************************************************************************************
Sub CargarIngresoAlmacen(IdCompra As Integer)
    Dim Rst As New ADODB.Recordset
    Dim A As Integer

    RST_Busq Rst, "SELECT alm_ingresodoc.iddoc, alm_ingreso.tipmov, alm_ingreso.fching, mae_documento.abrev, alm_ingreso.nombre, alm_ingreso.id, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS numdoc " _
        & " FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) RIGHT JOIN alm_ingresodoc ON alm_ingreso.id = alm_ingresodoc.id " _
        & " WHERE (((alm_ingresodoc.iddoc)=" & IdCompra & ") AND ((alm_ingreso.tipmov)=-1))", xCon

    Fg4.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(A, 1) = Rst("fching")
            Fg4.TextMatrix(A, 2) = Rst("abrev")
            Fg4.TextMatrix(A, 3) = Rst("numdoc")
            Fg4.TextMatrix(A, 4) = Rst("nombre")
            Fg4.TextMatrix(A, 5) = Rst("id")

            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea()
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtFchDoc.Locked = Not TxtFchDoc.Locked
    TxtFchVen.Locked = Not TxtFchVen.Locked
    TxtFchPago.Locked = Not TxtFchPago.Locked
    TxtConPag.Locked = Not TxtConPag.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtGlosa.Locked = Not TxtGlosa.Locked
    TxtBruto.Locked = Not TxtBruto.Locked
    TxtTipDocRef.Locked = Not TxtTipDocRef.Locked
    TxtDocRef2.Locked = Not TxtDocRef2.Locked

    Frame9.Enabled = Not Frame9.Enabled
    
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LIMPIA LOS CONTROLES DEL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    lblReg.Caption = ""
    TxtTipCom.Text = ""
    TxtTipDoc.Text = ""
    TxtNumRuc.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtFchDoc.Valor = ""
    TxtFchVen.Valor = ""
    TxtFchPago.Valor = ""
    TxtConPag.Text = ""
    TxtIdMon.Text = ""
    TxtGlosa.Text = ""
    TxtDocRef.Text = ""
    
    LblIdCenCos.Caption = ""
    LblNomDoc.Caption = ""
    LblNomPro.Caption = ""
    LblCondPag.Caption = ""
    LblMoneda.Caption = ""
    LblIdProveedor.Caption = ""
    LblTipoCompra.Caption = ""
    
    TxtBruto.Text = "0.00"
    TxtIGV.Text = "0.00"
    TxtTotal.Text = "0.00"
    
    Label3(9).Visible = False
    TxtDocRef.Visible = False
    CmdBusDocRef.Visible = False
    
    TxtTipDocRef.Text = ""
    TxtDocRef2.Text = ""
    LblTipDocref.Caption = ""
    LblIdDocRef2.Caption = ""
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Fg1.ColWidth(1) = 4500 - 2000
        Fg1.ColWidth(15) = 1000
        Fg1.ColWidth(16) = 1000
    Else
        Fg1.ColWidth(1) = 4500
        Fg1.ColWidth(15) = 0
        Fg1.ColWidth(16) = 0
    End If
End Sub

Private Sub ChkAjusta_Click()
    If ChkAjusta.Value = 1 Then
        TxtBruto.Locked = False
        TxtIGV.Locked = False
        TxtTotal.Locked = False
    Else
        TxtBruto.Locked = True
        TxtIGV.Locked = True
        TxtTotal.Locked = True
    End If
End Sub

Private Sub ChkImpRen4_Click()
    ' BUSCA LOS IMPUESTO RELACIONADOS CON EL DOCUMENTO EN FUNCION A LOS ITEMS REGISTRADOS
    BuscarImpuestos
End Sub

Private Sub CmdAcep_Click()
    TabOne1.Enabled = True
    Toolbar1.Enabled = True
    
    Frame11.Visible = False
End Sub

Private Sub CmdAceptar_Click()
    ' VALIDA LA DISTRIBUCION DE LOS CENTROS DE COSTO PARA LA COMPRA
    If QueHace = 3 Then
        ActivarEntorno
        Frame6.Visible = False
        Exit Sub
    End If
    
    Dim xTot As Double
    xTot = NulosN(TxtBruto.Text)
    
    If NulosN(Format(xTot, "0.00")) <> NulosN(Format(TxtTotImp.Text, "0.00")) Then
        MsgBox "la distribucion del centro de costo no coincide con el importe del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    LblIdCenCos.Caption = ""
    
    DetCenCos = True
    Frame6.Visible = False
    ActivarEntorno
End Sub

Private Sub CmdAddCenCos_Click()
    ' AGREGA CENTRO DE COSTOS
    If QueHace = 3 Then Exit Sub
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Dim xfrm As New SGI2_funciones.formularios
    Set Rst = xfrm.SeleCentroCosto(xCon)
    
    If Rst.State = 1 Then
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                Fg5.Rows = Fg5.Rows + 1
                Fg5.TextMatrix(Fg5.Rows - 1, 1) = Rst("codigo")
                Fg5.TextMatrix(Fg5.Rows - 1, 2) = Rst("descripcion")
                Fg5.TextMatrix(Fg5.Rows - 1, 5) = Rst("idcencos")
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If
    End If
    Set xfrm = Nothing
End Sub

Private Sub CmdAddItem_Click()
    ' AGREGA UN ITEM A LA COMPRA
    If QueHace = 3 Then Exit Sub
    If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 1)) = "" Then Exit Sub
    Fg1.Rows = Fg1.Rows + 1
    
    With Fg1
        .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, 1
    End With
    
    fg1_CellButtonClick Fg1.Rows - 1, 1
    If Fg1.Row >= 1 Then Fg1.Col = 4
    Fg1.SetFocus
End Sub

Private Sub CmdBusAlm_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT alm_almacenes.* FROM alm_almacenes"
    
    xform.Titulo = "Buscando Almacenes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtNumRuc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdApertura_Click()
    AperturaDocumento xCon, xIdUsuario, 40, IdMenuActivo
    ' refrescar la consulta
    RstComp.Filter = ""
    TDB_FiltroLimpiar Dg1
    RstComp.Requery
End Sub

Private Sub CmdBusCondicion_Click()
    ' BUSCA LA CONDICION DE PAGO EN QUE SE EFECTUA LA COMPRA
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
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
        If xRs.RecordCount <> 0 Then
            TxtConPag.Text = xRs("id")
            LblCondPag.Caption = xRs("descripcion")
            TxtFchVen.Valor = CDate(TxtFchDoc.Valor) + xRs("numdia")
            TxtFchVen.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDocRef_Click()
    ' BUSCA EL DOCUMENTO DE REFERENCIA PARA LA COMPRA
    If QueHace = 3 Then Exit Sub

    If NulosN(LblIdProveedor.Caption) = 0 Then
        MsgBox "No ha especificado el proveedor para referenciar este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(6, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Tipo. Doc.":       xCampos(0, 1) = "abrev":                xCampos(0, 2) = "1000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fch. Doc.":        xCampos(1, 1) = "fchdoc":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Documento":     xCampos(2, 1) = "numdoc":               xCampos(2, 2) = "1500":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Fch. Ven.":        xCampos(3, 1) = "fchven":               xCampos(3, 2) = "1000":         xCampos(3, 3) = "C"
    xCampos(4, 0) = "Total":            xCampos(4, 1) = "imptot":               xCampos(4, 2) = "1000":         xCampos(4, 3) = "N"
    xCampos(5, 0) = "Condicion":        xCampos(5, 1) = "descripcion":          xCampos(5, 2) = "1000":         xCampos(5, 3) = "C"
    
    xform.SQLCad = "SELECT mae_documento.abrev, com_honorarios.fchdoc, [com_honorarios]![numser]+'-'+[com_honorarios]![numdoc] AS numdoc, com_honorarios.fchven," _
        & " mae_prov.nombre, mae_condpago.descripcion, com_honorarios.id, com_honorarios.imptot FROM mae_condpago LEFT JOIN (mae_documento RIGHT JOIN " _
        & " (mae_prov RIGHT JOIN com_honorarios ON mae_prov.id = com_honorarios.idpro) ON mae_documento.id = com_honorarios.tipdoc) ON mae_condpago.id = com_honorarios.idconpag " _
        & " WHERE (((com_honorarios.idpro)=" & NulosN(LblIdProveedor.Caption) & ") AND ((com_honorarios.tipdoc)<>7))"
    
    xform.Titulo = "Buscando Documentos del Proveedor"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtDocRef.Text = xRs("numdoc")
            LblIdDocRef.Caption = xRs("id")
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDocRef2_Click()
    ' BUSCA EL DOCUMENTO DE REFERENCIA PARA LA COMPRA
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(4, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nº Documento":      xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fch. Emi.":         xCampos(1, 1) = "fchemi":      xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Fch. Ven.":         xCampos(2, 1) = "fchven":      xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Proveedor":         xCampos(3, 1) = "nombre":      xCampos(3, 2) = "4000":         xCampos(3, 3) = "C"
    
    
    xform.FormaBusca = Principio
    
    If NulosN(TxtTipDocRef.Text) = 1 Then
        'Orden de Compra
        xform.SQLCad = "SELECT com_ordencompra.id, [com_ordencompra]![numser] & '-' & [com_ordencompra]![numdoc] AS numdoc, com_ordencompra.fchemi, " _
            & " com_ordencompra.fchven, mae_prov.nombre FROM mae_prov RIGHT JOIN com_ordencompra ON mae_prov.id = com_ordencompra.idpro " _
            & " Where (((com_ordencompra.idest) = 2)) ORDER BY com_ordencompra.fchemi"
        xform.Titulo = "Orden de Compra"
    ElseIf NulosN(TxtTipDocRef.Text) = 2 Then
        'Orden de Produccion
        MsgBox "Opcion no disponible"
        xform.Titulo = "Orden de Produccion"
        Exit Sub
    ElseIf NulosN(TxtTipDocRef.Text) = 3 Then
        'Orden de Mantenimiento
        xform.Titulo = "Orden de Matenimiento"
        MsgBox "Opcion no disponible"
        Exit Sub
    ElseIf NulosN(TxtTipDocRef) = 4 Then
        'Orden de Despacho
        xCampos(3, 0) = "Cliente"
        xform.SQLCad = "SELECT var_ordendespacho.id, var_ordendespacho.numerodoc AS numdoc,mae_cliente.nombre, var_ordendespacho.idcli, var_ordendespacho.fchemi, var_ordendespacho.fchven  " _
            & " FROM var_ordendespacho LEFT JOIN mae_cliente ON var_ordendespacho.idcli = mae_cliente.id "
        
        xform.Titulo = "Buscando Orden de Despacho"
        xform.FormaBusca = CualquierParte
        
        
    Else
        Exit Sub
    End If
    
    xform.Criterio = ""
    xform.Ordenado = "numdoc"
    xform.CampoBusca = "numdoc"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtDocRef2.Text = NulosC(xRs("numdoc"))
            LblIdDocRef2.Caption = xRs("id")
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusProv_Click()
    ' BUSCA UN PROVEEDOR
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Proveedor":    xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_prov.nombre, mae_prov.numruc, mae_prov.id, mae_prov.idcondpag From mae_prov WHERE (((mae_prov.activo)=-1) AND ((mae_prov.tipper)=1) AND ((mae_prov.idtipdoc)=5))"
    
    xform.Titulo = "Buscando Prestador de Servicio"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumRuc.Text = xRs("numruc")
            LblNomPro.Caption = xRs("nombre")
            LblIdProveedor.Caption = xRs("id")
            
            If xRs("idcondpag") <> 0 Then
                TxtConPag.Text = xRs("idcondpag")
                TxtConPag_Validate True
            End If
            
            TxtFchDoc.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    ' BUSCA LA MONEDA DE LA COMPRA
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
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
        If xRs.RecordCount <> 0 Then
            TxtIdMon.Text = xRs("id")
            LblMoneda.Caption = xRs("descripcion")
            TxtTipDoc.SetFocus
            
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
                LblTipoCambio.Caption = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
            End If
            xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDocRef_Click()
    ' BUSCA EL TIPO DE DOCUMENTO DE REFERENCIA PARA LA COMPRA
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_docreferencia ORDER BY descripcion"
    
    xform.Titulo = "Buscando Tipo de Documento de Referencia"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipDocRef.Text = xRs("id")
            LblTipDocref.Caption = xRs("descripcion")
            TxtDocRef2.Text = ""
            LblIdDocRef2.Caption = ""
            TxtDocRef2.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipoCompra_Click()
    ' BUSCA TIPO DE PRODUCTO
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_tipoproducto ORDER BY descripcion"
    
    xform.Titulo = "Buscando Tipo"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipCom.Text = xRs("id")
            LblTipoCompra.Caption = xRs("descripcion")
            TxtIdMon.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdCancelar_Click()
    ActivarEntorno
    DetCenCos = False
    Frame6.Visible = False
End Sub

Private Sub CmdCargaDoc_Click()
    If OptOpera3.Value = True Then
        PopupMenu opciones
    End If
    If OptOpera2.Value = True Then
        'AdjuntarOrdenCompra
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : AdjuntarEntradas
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : VINCULA LOS INGRESO DE ALMACEN CON LA COMPRA ACTUAL
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Tipo      |  INTEGER          |  ESPECIFICA SI SE CARGARA UN INGRESO NUEVO O UNO
'*                                                     YA ASIGNADO A OTRA COMPRA :
'*                                                     Tipo = 1  muestra las entradas no procesadas
'*                                                     Tipo = 2  muestra las entradas procesadas
'* Devuelve         :
'*****************************************************************************************************
Sub AdjuntarEntradas(Tipo As Integer)
    'Tipo = 1  muestra las entradas no procesadas
    'Tipo = 2  muestra las entradas procesadas
    Dim xfrm As New eps_librerias.FormSeleccion
    Dim xCampos(4, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim xRs1 As New ADODB.Recordset
    
    xCampos(0, 0) = "Documento":       xCampos(0, 1) = "abrev":         xCampos(0, 2) = "1200":   xCampos(0, 3) = "C":     xCampos(0, 4) = "N"
    xCampos(1, 0) = "Nº Documento":    xCampos(1, 1) = "numdoc":        xCampos(1, 2) = "1500":   xCampos(1, 3) = "C":     xCampos(1, 4) = "S"
    xCampos(2, 0) = "Fch. Giro":       xCampos(2, 1) = "fchdoc":        xCampos(2, 2) = "1000":   xCampos(2, 3) = "C":     xCampos(2, 4) = "N"
    xCampos(3, 0) = "Proveedor":       xCampos(3, 1) = "nombre":        xCampos(3, 2) = "3000":   xCampos(3, 3) = "C":     xCampos(3, 4) = "N"

    If Tipo = 1 Then
        ' entradas no procesadas
        xfrm.SQLCad = "SELECT alm_ingreso.fchdoc, mae_documento.abrev, alm_ingreso.nombre, alm_ingreso!numser+'-'+alm_ingreso!numdoc AS numdoc, " _
            & " alm_ingreso.id, (SELECT Count(1) AS numdocs From alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocs " _
            & " FROM alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id WHERE ((((SELECT Count(1) AS numdocs From alm_ingresodoc " _
            & " WHERE (((alm_ingresodoc.id)=alm_ingreso.id))))=0)) ORDER BY alm_ingreso.fchdoc"
    Else
        ' ENTRADAS PROCESADAS O ASIGNADAS A UNA COMPRA
        xfrm.SQLCad = "SELECT alm_ingreso.fchdoc, mae_documento.abrev, alm_ingreso.nombre, alm_ingreso!numser+'-'+alm_ingreso!numdoc AS numdoc, " _
            & " alm_ingreso.id, (SELECT Count(1) AS numdocs From alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocs " _
            & " FROM alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id WHERE ((((SELECT Count(1) AS numdocs From alm_ingresodoc " _
            & " WHERE (((alm_ingresodoc.id)=alm_ingreso.id))))<>0)) ORDER BY alm_ingreso.fchdoc"
    End If
        
    xfrm.Titulo = "Buscando Entradas a Almacen"
    Set xfrm.Coneccion = xCon
    Set xRs = xfrm.Seleccionar(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount = 0 Then
            Set xRs = Nothing
            Exit Sub
        End If
        Dim xCadWHERE As String
        Dim A As Integer
        Dim Rst As New ADODB.Recordset
        
        Fg4.Rows = 1
        xRs.MoveFirst
        
        'CARGAMOS LOS DOCUMENTOS ADJUNTOS Y LO MOSTRAMOS EN LA LISTA DE "DOCUMENTOS ADJUNTOS"
        For A = 1 To xRs.RecordCount
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(A, 1) = xRs("fchdoc")
            Fg4.TextMatrix(A, 2) = xRs("abrev")
            Fg4.TextMatrix(A, 3) = xRs("numdoc")
            Fg4.TextMatrix(A, 4) = xRs("nombre")
            Fg4.TextMatrix(A, 5) = xRs("id")
            xRs.MoveNext
            If xRs.EOF = True Then Exit For
        Next A
        
        CargarItems
    End If
    Set xfrm = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : CargarItems
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS ITEMS DE LOS DOCUMENTOS DE INGRESO CARGADOS EN EL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargarItems()
    Dim A As Integer
    Dim xCadWHERE As String
    Dim Rst As New ADODB.Recordset
    
    ' PREPARAMOS LA CADENA WHERE DE LA COSULTA A EJECUTARSE, ESTA CADENA WHERE SOLO CARGARA LOS DOCUMENTOS DE INGRESO CARGADOS
    For A = 1 To Fg4.Rows - 1
        xCadWHERE = xCadWHERE + "(alm_ingresodet.id = " & NulosN(Fg4.TextMatrix(A, 5)) & ")"
        If A = Fg4.Rows - 1 Then
            Exit For
        End If
        xCadWHERE = xCadWHERE + " OR "
    Next A
    
    xCadWHERE = "(" + xCadWHERE + ")"
    
    ' EJECUTAMOS LA CONSULTA
    RST_Busq Rst, "SELECT alm_inventario.codpro, mae_unidades.abrev, alm_inventario.descripcion, Sum(alm_ingresodet.cantidad) AS cantidad, " _
        & " con_planctas.ctadesdeb, con_planctas.ctadeshab, alm_inventario.idcuenta, alm_inventario.iddet, alm_inventario.idtipcom, alm_inventario.id, " _
        & " alm_inventario.idunimed " _
        & " FROM con_planctas RIGHT JOIN ((alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades " _
        & " ON alm_inventario.idunimed = mae_unidades.id) ON con_planctas.id = alm_inventario.idcuenta " _
        & " Where " + xCadWHERE _
        & " GROUP BY alm_inventario.codpro, mae_unidades.abrev, alm_inventario.descripcion, con_planctas.ctadesdeb, con_planctas.ctadeshab, " _
        & " alm_inventario.idcuenta, alm_inventario.iddet, alm_inventario.idtipcom, alm_inventario.id, alm_inventario.idunimed", xCon
    
    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Mostrando = True
        
        ' MOSTRAMOS LOS ITEMS CARGADOS
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = Rst("descripcion")
            Fg1.TextMatrix(A, 2) = Rst("abrev")
            Fg1.TextMatrix(A, 3) = Format(Rst("cantidad"), "0.0000")
            Fg1.TextMatrix(A, 4) = 0
            
            Fg1.TextMatrix(A, 9) = Rst("id")
            Fg1.TextMatrix(A, 10) = Rst("idunimed")
            Fg1.TextMatrix(A, 11) = NulosN(Rst("idcuenta"))
            Fg1.TextMatrix(A, 12) = NulosN(Rst("idtipcom"))
            Fg1.TextMatrix(A, 13) = NulosN(Rst("ctadesdeb"))
            Fg1.TextMatrix(A, 14) = NulosN(Rst("ctadeshab"))
        
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
        Mostrando = False
    End If
End Sub

Private Sub CmdDelCenCos_Click()
    ' BORRA UN CENTRO DE COSTO
    If Fg5.Row < Fg5.FixedRows Then Exit Sub
    Fg5.RemoveItem Fg5.Row
End Sub

Private Sub CmdDelItem_Click()
    ' BORRA UNA FILA DEL CONTROL FlexGrid Fg1
    If QueHace = 3 Then Exit Sub
    If Fg1.Rows = 1 Then
        MsgBox "No hay items para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If Fg1.Row < 1 Then
        MsgBox "Seleccione una fila correcta para eliminar", vbExclamation, xTitulo
        Exit Sub
    End If
    
    Fg1.RemoveItem Fg1.Row
    HallarTotal
End Sub

'*****************************************************************************************************
'* Nombre           : ActivarEntorno
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TabOne1 y Toolbar1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivarEntorno()
    TabOne1.Enabled = Not TabOne1.Enabled
    Toolbar1.Enabled = Not Toolbar1.Enabled
End Sub

Private Sub CmdDetCenCos_Click()
    If QueHace = 3 Then
        CmdAddCenCos.Enabled = False
        CmdDelCenCos.Enabled = False
        CmdCancelar.Enabled = False
        Fg5.Editable = flexEDNone
        Fg5.SelectionMode = flexSelectionByRow
    Else
        CmdAddCenCos.Enabled = True
        CmdDelCenCos.Enabled = True
        CmdCancelar.Enabled = True
        Fg5.Editable = flexEDKbdMouse
        Fg5.SelectionMode = flexSelectionFree
    End If
    ActivarEntorno
    TxtTotPor.Text = ""
    TxtTotImp.Text = ""
    Frame6.Left = 1545
    Frame6.Top = 2190
    Frame6.Visible = True
    HallarTotCenCos
End Sub

Private Sub CmdSeleccionar_Click()
    If Trim(CmdSeleccionar.Caption) = "Ver Documentos" Then
        TabOne1.Enabled = False
        Toolbar1.Enabled = False
        
        Frame11.Left = 2280
        Frame11.Top = 2550
        Frame11.Visible = True
        Exit Sub
    End If

    If QueHace = 3 Then Exit Sub
    
    If xOrigen = 0 Then
        If NulosC(TxtTipCom.Text) = "" Then
            MsgBox "No ha especificado el tipo de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtTipCom.SetFocus
            Exit Sub
        End If
    End If
    
    Dim xCampos(3, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLId As String
    Dim A&
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "4000":         xCampos(0, 3) = "C":    xCampos(0, 4) = "N"
    xCampos(1, 0) = "Uni. Med":       xCampos(1, 1) = "abrev":         xCampos(1, 2) = "1000":         xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "Codigo":         xCampos(2, 1) = "codpro":        xCampos(2, 2) = "1800":         xCampos(2, 3) = "C":    xCampos(2, 4) = "S"

    '*******************************************************************************************
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 9, "alm_inventario.id", " NOT IN ", True)
    '*******************************************************************************************
    If xOrigen = 0 Then
        If nSQLId <> "" Then nSQLId = " AND " & nSQLId
        nSQL = "SELECT CONSULTA1.*, CONSULTA2.precio FROM " _
            & " [SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.descripcion AS descuni, mae_unidades.abrev, " _
            & " con_planctas.ctadesdeb, con_planctas.ctadeshab,  alm_inventario.idunimed,  alm_inventario.idcuenta, alm_inventario.idtipcom " _
            & " FROM mae_unidades INNER JOIN (con_planctas RIGHT JOIN alm_inventario ON con_planctas.id = alm_inventario.idcuenta) ON mae_unidades.id = alm_inventario.idunimed " _
            & " Where (((alm_inventario.tippro) = " & NulosN(TxtTipCom.Text) & ")) ORDER BY alm_inventario.descripcion]. AS CONSULTA1 LEFT JOIN " _
            & " [SELECT com_honorariosdet.iditem, Min(com_honorariosdet.preuni) AS precio From com_honorariosdet GROUP BY com_honorariosdet.iditem]. AS CONSULTA2 ON CONSULTA1.id = CONSULTA2.iditem ORDER BY CONSULTA1.descripcion"
    Else
        If nSQLId <> "" Then nSQLId = " WHERE " & nSQLId
        nSQL = "SELECT alm_inventario.*, mae_unidades.descripcion AS descuni, mae_unidades.abrev, " _
                & " con_planctas.ctadesdeb, con_planctas.ctadeshab FROM con_planctas RIGHT JOIN (mae_unidades INNER JOIN " _
                & " alm_inventario ON mae_unidades.id = alm_inventario.idunimed) ON con_planctas.id = alm_inventario.idcuenta " _
                & " " & nSQLId & " ORDER BY alm_inventario.descripcion"
    End If
    
   '*******************************************************************************************
    CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Productos"
    '*******************************************************************************************
    If xRs.State = 1 Then
        Mostrando = True
        If xRs.RecordCount <> 0 Then xRs.MoveFirst
        Do While Not xRs.EOF
            Fg1.Rows = Fg1.Rows + 1
            
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("descripcion"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = xRs("abrev")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosN(xRs("precio")), "0.0000")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = xRs("id")
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosN(xRs("idunimed"))
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(xRs("idcuenta"))
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosN(xRs("idtipcom"))
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosN(xRs("ctadesdeb"))
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosN(xRs("ctadeshab"))
           
            xRs.MoveNext
        Loop
    End If
    Mostrando = False
    Set xRs = Nothing
End Sub

Private Sub CmdVerAsiento_Click()
    VerAsiento
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstComp
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LAS COLUMNAS DEL CONTROL DataGrid Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstComp.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstComp("id")), xCon
    End If
End Sub

Private Sub fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    ' PERMITE BUSCAR UN ITEM
    If xOrigen = 0 Then
        If NulosN(TxtTipCom.Text) = 0 Then
            MsgBox "No ha especificado el tipo de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtTipCom.SetFocus
            Exit Sub
        End If
    End If
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5400":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Unid.":        xCampos(1, 1) = "abrev":          xCampos(1, 2) = "600":     xCampos(1, 3) = "C"
    xCampos(2, 0) = "Codigo":       xCampos(2, 1) = "codpro":         xCampos(2, 2) = "2000":    xCampos(2, 3) = "C"
    
    '*******************************************************************************************
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 9, "alm_inventario.id", " NOT IN ", True)
    '*******************************************************************************************
    If xOrigen = 0 Then
        If nSQLId <> "" Then nSQLId = " and " & nSQLId
        xform.SQLCad = "SELECT CONSULTA1.*, CONSULTA2.precio FROM " _
            & " [SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.descripcion AS descuni, mae_unidades.abrev, " _
            & " con_planctas.ctadesdeb, con_planctas.ctadeshab,  alm_inventario.idunimed,  alm_inventario.idcuenta, alm_inventario.idtipcom " _
            & " FROM mae_unidades INNER JOIN (con_planctas RIGHT JOIN alm_inventario ON con_planctas.id = alm_inventario.idcuenta) ON mae_unidades.id = alm_inventario.idunimed " _
            & " Where (alm_inventario.tipo) In (1,3) and (((alm_inventario.tippro) = " & NulosN(TxtTipCom.Text) & ")) ORDER BY alm_inventario.descripcion]. AS CONSULTA1 LEFT JOIN " _
            & " [SELECT com_honorariosdet.iditem, Min(com_honorariosdet.preuni) AS precio From com_honorariosdet GROUP BY com_honorariosdet.iditem]. AS CONSULTA2 ON CONSULTA1.id = CONSULTA2.iditem ORDER BY CONSULTA1.descripcion"
    Else
        If nSQLId <> "" Then nSQLId = " and " & nSQLId
        xform.SQLCad = "SELECT alm_inventario.*, mae_unidades.descripcion AS descuni, mae_unidades.abrev, " _
            & " con_planctas.ctadesdeb, con_planctas.ctadeshab FROM con_planctas RIGHT JOIN (mae_unidades INNER JOIN " _
            & " alm_inventario ON mae_unidades.id = alm_inventario.idunimed) ON con_planctas.id = alm_inventario.idcuenta " _
            & " WHERE (((alm_inventario.tipo) In (1,3)))  " & nSQLId & " AND alm_inventario.idcuenta <> 0 ORDER BY alm_inventario.descripcion "
    End If
    
    xform.Titulo = "Buscando Productos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "codpro"
    xform.CampoBusca = "codpro"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    Mostrando = True
    Dim A As Integer
    
    If xRs.State = 1 Then
        If Fg1.Rows <> 1 Then
            'VALIDAMOS QUE EL ITEM SELECCIONADO NO ESTE AGREGADO
            For A = 1 To Fg1.Rows - 1
                If Fg1.TextMatrix(A, 9) = xRs("id") Then
                    MsgBox "El item seleccionado ya fue agregado, seleccione otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    A = Fg1.Rows - 1
                    Set xRs = Nothing
                    Exit Sub
                End If
            Next A
        End If
        
        ' AGREGAMOS EL ITEM SELECCIONADO AL FlexGrid Fg1
        If xRs.RecordCount <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 1) = xRs("descripcion")
            Fg1.TextMatrix(Fg1.Row, 2) = xRs("abrev")
            Fg1.TextMatrix(Fg1.Row, 3) = 1
            Fg1.TextMatrix(Fg1.Row, 4) = Format(NulosN(xRs("precio")), "0.0000")
            Fg1.TextMatrix(Fg1.Row, 9) = xRs("id")
            Fg1.TextMatrix(Fg1.Row, 10) = xRs("idunimed")
            Fg1.TextMatrix(Fg1.Row, 11) = NulosN(xRs("idcuenta"))
            Fg1.TextMatrix(Fg1.Row, 12) = NulosN(xRs("idtipcom"))
            Fg1.TextMatrix(Fg1.Row, 13) = NulosN(xRs("ctadesdeb"))
            Fg1.TextMatrix(Fg1.Row, 14) = NulosN(xRs("ctadeshab"))
        End If
    End If
    Mostrando = False
    Set xform = Nothing
    If Fg1.Row >= 1 Then Fg1.Col = 4
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : AgregarCentroCosto2
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AGREGA UN CENTRO DE COSTO
'* Paranetros       : NOMBRE        |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    CargarGrabado |  BOOLEAN    |  especifica que se levantara un centro de costos
'*                                                   que haya sido grabado
'*                    IdCompra      |  INTEGER    |  ESPECIFICA EL ID DE LA COMPRA
'* Devuelve         :
'*****************************************************************************************************
Sub AgregarCentroCosto2(CargarGrabado As Boolean, Optional IdCompra As Integer)
    'CargarGrabado = especifica que se levantara un centro de costos que haya sido grabado
    Dim Rst As New ADODB.Recordset
    Dim A, B, C, xFila As Integer
    Dim SeEncontro As Boolean
    
    Fg5.Rows = 1
        
    If CargarGrabado = True Then
        RST_Busq Rst, "SELECT com_honorarioscosto.idcom, com_honorarioscosto.idcencos, con_centrocosto.codigo, con_centrocosto.descripcion, com_honorarioscosto.imppor, com_honorarioscosto.impcos, " _
            & " con_centrocosto.tipo FROM con_centrocosto INNER JOIN com_honorarioscosto ON con_centrocosto.id = com_honorarioscosto.idcencos " _
            & " WHERE (((com_honorarioscosto.idcom)=" & IdCompra & "))", xCon
            
        If Rst.RecordCount <> 0 Then
            Fg5.Rows = 1
            Rst.MoveFirst
            Mostrando = True
            For A = 1 To Rst.RecordCount
                Fg5.Rows = Fg5.Rows + 1
                Fg5.TextMatrix(Fg5.Rows - 1, 1) = NulosC(Rst("codigo"))
                Fg5.TextMatrix(Fg5.Rows - 1, 2) = NulosC(Rst("descripcion"))
                Fg5.TextMatrix(Fg5.Rows - 1, 3) = Format(NulosN(Rst("imppor")), "0.00")
                Fg5.TextMatrix(Fg5.Rows - 1, 4) = Format(NulosN(Rst("impcos")), "0.00")
                Fg5.TextMatrix(Fg5.Rows - 1, 5) = NulosN(Rst("idcencos"))
                
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
            Mostrando = False
        End If
    Else
        For A = 1 To Fg1.Rows - 1
            ' buscamos si el item actual tiene centros de costo definido
            RST_Busq Rst, "SELECT alm_invencencos.idpro, alm_invencencos.idcencos, con_centrocosto.codigo, con_centrocosto.descripcion, " _
            & " alm_invencencos.imppor FROM alm_invencencos LEFT JOIN con_centrocosto ON alm_invencencos.idcencos = con_centrocosto.id " _
            & " WHERE (((alm_invencencos.idpro)=" & NulosN(Fg1.TextMatrix(A, 9)) & "))", xCon
            
            If Rst.RecordCount <> 0 Then
                ' si tiene centro de costos agregamos a la cuadricula centro de costos
                Rst.MoveFirst
                For B = 1 To Rst.RecordCount
                    ' buscamos si el cetro de costo ya fue agregado a la cuadricula
                    SeEncontro = False
                    xFila = 0
                    For C = 1 To Fg5.Rows - 1
                        If Fg5.TextMatrix(C, 5) = Rst("idcencos") Then
                            SeEncontro = True
                            xFila = C
                        End If
                    Next C
                    
                    If SeEncontro = True Then
                        ' nos pocisionamos en la fila que contiene el centro de costos y sumamos el valor
                        If Rst("imppor") < 100 Then
                            Fg5.TextMatrix(Fg5.Rows - 1, 4) = NulosN(Fg5.TextMatrix(Fg5.Rows - 1, 4)) + (NulosN(Fg1.TextMatrix(A, 8)) * ((Rst("imppor") / 100) + 1))
                        Else
                            Fg5.TextMatrix(Fg5.Rows - 1, 4) = NulosN(Fg5.TextMatrix(Fg5.Rows - 1, 4)) + NulosN(Fg1.TextMatrix(A, 8))
                        End If
                    Else
                        ' agregamos una nueva fila a la cuadricula centro de costos
                        Fg5.Rows = Fg5.Rows + 1
                        Fg5.TextMatrix(Fg5.Rows - 1, 1) = NulosC(Rst("codigo"))
                        Fg5.TextMatrix(Fg5.Rows - 1, 2) = NulosC(Rst("descripcion"))
                        Fg5.TextMatrix(Fg5.Rows - 1, 3) = Format(NulosN(Rst("imppor")), "0.00")
                        Fg5.TextMatrix(Fg5.Rows - 1, 5) = NulosN(Rst("idcencos"))
                        If NulosN(Rst("imppor")) < 100 Then
                            Fg5.TextMatrix(Fg5.Rows - 1, 4) = NulosN(Fg1.TextMatrix(A, 8)) * ((Rst("imppor") / 100) + 1)
                        Else
                            Fg5.TextMatrix(Fg5.Rows - 1, 4) = NulosN(Fg1.TextMatrix(A, 8))
                        End If
                        Fg5.TextMatrix(Fg5.Rows - 1, 4) = Format(Fg5.TextMatrix(Fg5.Rows - 1, 4), "0.00")
                    End If
                    
                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                Next B
            Else
                If NulosN(NulosN(Fg1.TextMatrix(A, 9))) <> 0 Then
                    'MsgBox "El item " & NulosC(Fg1.TextMatrix(A, 1)) & ", no tiene especificado un centro de costos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                End If
            End If
        Next A
            
        If NulosN(TxtBruto.Text) <> 0 Then
            For A = 1 To Fg5.Rows - 1
                Fg5.TextMatrix(A, 3) = (NulosN(Fg5.TextMatrix(A, 4)) / (NulosN(TxtBruto.Text))) * 100
                Fg5.TextMatrix(A, 3) = Format(Fg5.TextMatrix(A, 3), "0.00")
            Next A
        End If
    End If
    Set Rst = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : AgregarCentroCosto
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AGREGA UN CENTRO DE COSTO
'* Paranetros       : NOMBRE        |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    xIdProducto   |  INTEGER    |  ESPECIFICA EL ID DEL ITEM
'*                    IdCompra      |  INTEGER    |  ESPECIFICA EL IMPORTE DEL ITEM
'* Devuelve         :
'*****************************************************************************************************
Sub AgregarCentroCosto(xIdProducto As Integer, xImporte As Double)
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    Dim SeEncontro As Boolean
    
    ' buscamos si el producto tiene centro de costo asignado
    RST_Busq Rst, "SELECT alm_invencencos.idpro, alm_invencencos.idcencos, con_centrocosto.codigo, con_centrocosto.descripcion, " _
        & " alm_invencencos.imppor FROM alm_invencencos LEFT JOIN con_centrocosto ON alm_invencencos.idcencos = con_centrocosto.id " _
        & " WHERE (((alm_invencencos.idpro)=" & xIdProducto & "))", xCon
    
    If Rst.RecordCount <> 0 Then
        For A = 1 To Rst.RecordCount
            For B = 1 To Fg5.Rows - 1
                SeEncontro = False
                If Fg5.TextMatrix(B, 5) = Rst("idcencos") Then
                    SeEncontro = True
                    Exit For
                End If
            Next B
            If SeEncontro = False Then
                ' si no lo encuentra lo debe de agregar a la lista de centro de costos
                Fg5.Rows = Fg5.Rows + 1
                Fg5.TextMatrix(Fg5.Rows - 1, 1) = Rst("codigo")
                Fg5.TextMatrix(Fg5.Rows - 1, 2) = Rst("descripcion")
                Fg5.TextMatrix(Fg5.Rows - 1, 3) = Format(Rst("imppor"), "0.00")
                If Rst("imppor") = 100 Then
                    Fg5.TextMatrix(Fg5.Rows - 1, 4) = xImporte * 1
                Else
                    Fg5.TextMatrix(Fg5.Rows - 1, 4) = xImporte * ((Rst("imppor") / 100) + 1)
                End If
                Fg5.TextMatrix(Fg5.Rows - 1, 4) = Format(Fg5.TextMatrix(Fg5.Rows - 1, 4), "0.00")
                Fg5.TextMatrix(Fg5.Rows - 1, 5) = Rst("idcencos")
            Else
                ' si el centro de costo ya existe, agregarlo al centro de costo ya existente
                MsgBox "Falta hacer esta opcion"
            End If
        Next A
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Mostrando = True Then Exit Sub
    If Row = 0 Then Exit Sub
    If Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Then
        Fg1.TextMatrix(Row, 3) = Format(Fg1.TextMatrix(Row, 3), "0.0000")
        Fg1.TextMatrix(Row, 4) = Format(Fg1.TextMatrix(Row, 4), "0.000000")
        Fg1.TextMatrix(Row, 5) = Format(Fg1.TextMatrix(Row, 5), "0.000000")
        Fg1.TextMatrix(Row, 6) = Format(Fg1.TextMatrix(Row, 6), "0.0000")
        
        ' verificamos si hay descuento
        ' chequeamo si es por porcentaje
        If OptDes1.Value = True Then
            ' Se esta aplicando descuento por porcentaje
            Dim xPorcen As Double
            If NulosN(Fg1.TextMatrix(Row, 6)) <> 0 Then
                xPorcen = (NulosN(Fg1.TextMatrix(Row, 6)) / 100)
                Fg1.TextMatrix(Row, 7) = (NulosN(Fg1.TextMatrix(Row, 4)) + NulosN(Fg1.TextMatrix(Row, 5))) * xPorcen
                Fg1.TextMatrix(Row, 7) = (NulosN(Fg1.TextMatrix(Row, 4)) + NulosN(Fg1.TextMatrix(Row, 5))) - NulosN(Fg1.TextMatrix(Row, 7))
            Else
                Fg1.TextMatrix(Row, 7) = (NulosN(Fg1.TextMatrix(Row, 4)) + NulosN(Fg1.TextMatrix(Row, 5)))
            End If
        End If
        If OptDes2.Value = True Then
            ' Se esta aplicando descuento por importe
            If NulosN(Fg1.TextMatrix(Row, 6)) <> 0 Then
                Fg1.TextMatrix(Row, 7) = (NulosN(Fg1.TextMatrix(Row, 4)) + NulosN(Fg1.TextMatrix(Row, 5))) - NulosN(Fg1.TextMatrix(Row, 6))
                Fg1.TextMatrix(Row, 7) = Format(Fg1.TextMatrix(Row, 7), "0.0000")
            Else
                Fg1.TextMatrix(Row, 7) = (NulosN(Fg1.TextMatrix(Row, 4)) + NulosN(Fg1.TextMatrix(Row, 5)))
            End If
        End If
        
        Fg1.TextMatrix(Row, 7) = Format(Fg1.TextMatrix(Row, 7), "0.000000")
        Fg1.TextMatrix(Row, 8) = NulosN(Fg1.TextMatrix(Row, 3)) * NulosN(Fg1.TextMatrix(Row, 7))
        Fg1.TextMatrix(Row, 8) = Format(Fg1.TextMatrix(Row, 8), "0.0000")
        
        HallarTotal
        BuscarImpuestos
    End If
    
    If Col = 15 Or Col = 16 Then
        Fg1.TextMatrix(Fg1.Row, 3) = Format(Fg1.TextMatrix(Fg1.Row, 16), "0.0000")
        
        Fg1.TextMatrix(Fg1.Row, 7) = "0.0000"
        Fg1.TextMatrix(Fg1.Row, 15) = Format(Fg1.TextMatrix(Fg1.Row, 15), "0.00")
        If NulosN(Fg1.TextMatrix(Fg1.Row, 16)) = 0 Then
            Fg1.TextMatrix(Fg1.Row, 4) = NulosN(Fg1.TextMatrix(Fg1.Row, 15)) / ((NulosN(LblIgvTasa.Caption) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 4) = (NulosN(Fg1.TextMatrix(Fg1.Row, 15)) / ((NulosN(LblIgvTasa.Caption) / 100) + 1)) / NulosN(Fg1.TextMatrix(Fg1.Row, 16))
        End If
        Fg1.TextMatrix(Fg1.Row, 4) = Format(Fg1.TextMatrix(Fg1.Row, 4), "0.0000")
        Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 4), "0.0000")
        
        Fg1.TextMatrix(Fg1.Row, 8) = NulosN(Fg1.TextMatrix(Fg1.Row, 7)) * NulosN(Fg1.TextMatrix(Fg1.Row, 3))
        Fg1.TextMatrix(Fg1.Row, 8) = Format(Fg1.TextMatrix(Fg1.Row, 8), "0.0000")
        BuscarImpuestos
        HallarTotal
    End If
End Sub


'*****************************************************************************************************
'* Nombre           : BuscarImpuestos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BUSCA LOS IMPUESTOS EN FUNCION A LOS ITEMS CARGADOS EN LA COMPRA, BUSCA EL I.G.V
'*                    Y EL I.S.C.
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub BuscarImpuestos()
    If Fg1.Rows = 1 Then Exit Sub
    Dim A As Integer
    Dim xImpSEL, xImpIGV As Double
    
    Dim Rst As New ADODB.Recordset
    
    Set RstTempISC = Nothing
    PreparaRST_ISC
    xImpSEL = 0
    
    'buscando selectivo
    For A = 1 To Fg1.Rows - 1
        If NulosC(Fg1.TextMatrix(A, 1)) <> "" Then
            RST_Busq Rst, "SELECT mae_impuestos.tasa, mae_impuestos.idcuen, con_planctas.cuenta " _
                & " FROM (alm_inventario LEFT JOIN mae_impuestos ON alm_inventario.idimpsel = mae_impuestos.id) " _
                & " LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id WHERE " _
                & " ((alm_inventario.id = " & NulosN(Fg1.TextMatrix(A, 9)) & " ))", xCon
            
            If Rst.RecordCount <> 0 Then
                If NulosN(Rst("idcuen")) <> 0 Then
                    xImpSEL = xImpSEL + NulosN(Fg1.TextMatrix(A, 5)) * (Rst("tasa") / 100)
                    
                    If RstTempISC.RecordCount = 0 Then
                        RstTempISC.AddNew
                        RstTempISC("idcuen") = Rst("idcuen")
                        RstTempISC("total") = RstTempISC("total") + NulosN(Fg1.TextMatrix(A, 5)) * (Rst("tasa") / 100)
                    Else
                        RstTempISC.MoveFirst
                        RstTempISC.Find "idcuen = " & Rst("idcuen") & ""
                        
                        If RstTempISC.EOF = False Then
                            RstTempISC("idcuen") = Rst("idcuen")
                            RstTempISC("total") = RstTempISC("total") + NulosN(Fg1.TextMatrix(A, 5)) * (Rst("tasa") / 100)
                        End If
                    End If
                End If
            End If
        End If
    Next A
    
    ' buscando el impuesto a las ventas
    xImpIGV = 0
    
    If CodSunatDoc = "01" Or CodSunatDoc = "07" Or CodSunatDoc = "04" Or CodSunatDoc = "08" Or CodSunatDoc = "12" Or CodSunatDoc = "14" Or CodSunatDoc = "29" Then
        xImpIGV = NulosN(TxtBruto.Text) * (NulosN(TasaImpuesto) / 100)
    End If
    
    If CodSunatDoc = "02" Then
        If ChkImpRen4.Value = 1 Then
            xImpIGV = NulosN(TxtBruto.Text) * (NulosN(TasaImpuesto) / 100)
        Else
            xImpIGV = 0
        End If
    End If
    
    If NulosN(TxtTipDoc.Text) <> 2 Then
        TxtIGV.Text = Format(xImpIGV, FORMAT_MONTO)
        TxtTotal.Text = NulosN(TxtBruto.Text) + NulosN(TxtIGV.Text)
        TxtTotal.Text = Format(TxtTotal.Text, FORMAT_MONTO)
    Else
        TxtIGV.Text = Format(xImpIGV, FORMAT_MONTO)
        TxtTotal.Text = NulosN(TxtBruto.Text) - NulosN(TxtIGV.Text)
        TxtTotal.Text = Format(TxtTotal.Text, FORMAT_MONTO)
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : PreparaRST_ISC
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL RECORDSET TEMPORAL RstTempISC
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub PreparaRST_ISC()
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(9, 3) As String

    xCampos(0, 0) = "idcuen":        xCampos(0, 1) = "N":      xCampos(0, 2) = "2"
    xCampos(1, 0) = "Total":         xCampos(1, 1) = "D":      xCampos(1, 2) = "2"
    Set RstTempISC = xFun.CrearRstTMP(xCampos)

    RstTempISC.Open
End Sub

'*****************************************************************************************************
'* Nombre           : HallarTotal
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : HALLA LOS TOTALES DE LOS ITEMS AGREGADOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub HallarTotal()
    Dim A As Integer
    Dim Total, TotalIna As Double
    Dim xPorcen As Double
    Dim PreDes As Double
    Fg5.Rows = 1
    Dim Valor As Double
    Total = 0
    TotalIna = 0
    For A = 1 To Fg1.Rows - 1
        If OptDes1.Value = True Then
            ' Se esta aplicando descuento por porcentaje
            If NulosN(Fg1.TextMatrix(A, 6)) <> 0 Then
                xPorcen = ((NulosN(Fg1.TextMatrix(A, 6)) / 100))
                PreDes = NulosN(Fg1.TextMatrix(A, 4)) - (NulosN(Fg1.TextMatrix(A, 4)) * xPorcen)
                Valor = PreDes * NulosN(Fg1.TextMatrix(A, 3))
                Total = Total + Valor
                
                Valor = (NulosN(Fg1.TextMatrix(A, 5)) / xPorcen) * NulosN(Fg1.TextMatrix(A, 3))
                TotalIna = TotalIna + Valor
            Else
                Valor = NulosN(Fg1.TextMatrix(A, 4)) * NulosN(Fg1.TextMatrix(A, 3))
                Total = Total + Valor
                
                Valor = NulosN(Fg1.TextMatrix(A, 5)) * NulosN(Fg1.TextMatrix(A, 3))
                TotalIna = TotalIna + Valor
            End If
        End If
        If OptDes2.Value = True Then
            ' Se esta aplicando descuento por importe
            If NulosN(Fg1.TextMatrix(A, 6)) <> 0 Then
                If NulosN(Fg1.TextMatrix(A, 4)) <> 0 Then
                    Valor = (NulosN(Fg1.TextMatrix(A, 4)) - NulosN(Fg1.TextMatrix(A, 6))) * NulosN(Fg1.TextMatrix(A, 3))
                    Total = Total + Valor
                End If
                If NulosN(Fg1.TextMatrix(A, 5)) <> 0 Then
                    Valor = (NulosN(Fg1.TextMatrix(A, 5)) - NulosN(Fg1.TextMatrix(A, 6))) * NulosN(Fg1.TextMatrix(A, 3))
                    TotalIna = TotalIna + Valor
                End If
            Else
                Total = Total + (NulosN(Fg1.TextMatrix(A, 4)) * NulosN(Fg1.TextMatrix(A, 3)))
                TotalIna = TotalIna + (NulosN(Fg1.TextMatrix(A, 5)) * NulosN(Fg1.TextMatrix(A, 3)))
            End If
        End If
    Next A

    TxtBruto.Text = Format(NulosN(Total), FORMAT_MONTO)
    AgregarCentroCosto2 False
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then Exit Sub
    If Fg1.Col = 2 Or Fg1.Col = 7 Or Fg1.Col = 8 Then
        Fg1.Editable = flexEDNone
    Else
        Fg1.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Or Col = 15 Or Col = 16 Then
        If InStr(CaracteresNumericos2, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = 45 Then
        CmdAddItem_Click
    End If
    If KeyCode = 46 Then
        CmdDelItem_Click
    End If
    If KeyCode = 93 Then
        PopupMenu menu1
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    If Button = 2 Then PopupMenu menu1
End Sub

'*****************************************************************************************************
'* Nombre           : CargarRSTCom
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA EL RECORSET PRINCIPAL DEL FORMULARIO, ESTE RECORSET SE VISUALIZARA EN LA
'*                    PANTALLA CONSULTA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargarRSTCom()
    Dim cSQL As String
    
    cSQL = "SELECT com_honorarios.*, mae_prov.nombre, IIf(IsNull([com_honorarios]![numser])=-1,[com_honorarios]![numdoc],[com_honorarios]![numser] & '-' & [com_honorarios]![numdoc]) AS numerodoc, mae_documento.descripcion AS nomdoc, " _
                & "mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_prov.numruc, mae_moneda.descripcion AS descmon, mae_moneda.simbolo, " _
                & "mae_tipoproducto.descripcion AS desctipcom, con_tc.impcom, Mid([com_honorarios].[numreg],1,2)+[mae_libros].[codsun]+Mid([com_honorarios].[numreg],3,4) AS numreg1, " _
                & "com_honorarios.fchdoc & '' as fchdoc1, com_honorarios.fchven & '' as fchven1, com_honorarios.imptot & '' as imptot1, com_honorarios.impsal & ''  as impsal1,  " _
                & "IIf([com_honorarios].[tc]=0,[con_tc].[impven],[com_honorarios].[tc]) & '' AS impven1,com_honorarios.impbru & '' as impbru1,com_honorarios.impigv & '' as impigv1  " _
            + vbCr + "FROM (mae_condpago RIGHT JOIN (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((com_honorarios LEFT JOIN mae_tipoproducto " _
                & "ON com_honorarios.idtipo = mae_tipoproducto.id) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) ON mae_documento.id = com_honorarios.tipdoc) " _
                & "ON mae_moneda.id = com_honorarios.idmon) ON mae_prov.id = com_honorarios.idpro) ON mae_condpago.id = com_honorarios.idconpag) LEFT JOIN mae_libros " _
                & "ON com_honorarios.idlib = mae_libros.id " _
            + vbCr + "WHERE (((com_honorarios.numreg) Like '" & Format(mMesActivo, "00") & "%')) " _
            + vbCr + "ORDER BY com_honorarios.numreg DESC"
    
    RST_Busq RstComp, cSQL, xCon
End Sub

Private Sub Fg5_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Mostrando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Dim xTot As Double
    xTot = NulosN(TxtBruto.Text)
    
    If Col = 3 Then
        If NulosN(Fg5.TextMatrix(Fg5.Row, 3)) > 100 Then
            Fg5.TextMatrix(Fg5.Row, 3) = ""
            Fg5.TextMatrix(Fg5.Row, 4) = ""
            Exit Sub
        End If
        If NulosN(Fg5.TextMatrix(Fg5.Row, 3)) <> 0 Then
            Fg5.TextMatrix(Fg5.Row, 4) = xTot * NulosN(Fg5.TextMatrix(Fg5.Row, 3) / 100)
        End If
    End If
    
    If Col = 4 Then
        If Fg5.TextMatrix(Fg5.Row, 4) > xTot Then Exit Sub
        
        If NulosN(Fg5.TextMatrix(Fg5.Row, 4)) <> 0 And xTot <> 0 Then
            Fg5.TextMatrix(Fg5.Row, 3) = ((NulosN(Fg5.TextMatrix(Fg5.Row, 4)) / xTot) * 100)
            Fg5.TextMatrix(Fg5.Row, 3) = Format(Fg5.TextMatrix(Fg5.Row, 3), "0.00")
        End If
    End If
    
    HallarTotCenCos
End Sub

'*****************************************************************************************************
'* Nombre           : HallarTotCenCos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : HALLA LA SUMA TOTAL DE LOS CENTROS DE COSTO CARGADOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub HallarTotCenCos()
    Dim A As Integer
    Dim TotPor, TotImp As Double
    
    For A = 1 To Fg5.Rows - 1
        TotPor = TotPor + NulosN(Fg5.TextMatrix(A, 3))
        TotImp = TotImp + NulosN(Fg5.TextMatrix(A, 4))
    Next A
    
    TxtTotPor.Text = Format(TotPor, "0.00")
    TxtTotImp.Text = Format(TotImp, "0.00")
End Sub

Private Sub Fg5_EnterCell()
    If QueHace = 3 Then Exit Sub
    If Fg5.Col = 3 Or Fg5.Col = 4 Then
        Fg5.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE DESPUES DE CARGAR EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        Dim Rpta As Integer
        Dim Rst As New ADODB.Recordset
        
        mMesActivo = xMes
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        xFechaMes = "01/" + Trim(Format(mMesActivo, "00")) + "/" + Trim(Format(AnoTra, "0000"))
        xFchIni = xFechaMes
        xFchFin = Format(HallaDiasMes(CDate(xFechaMes)), "00") + "/" + Mid(xFechaMes, 4, 7)
        
        If xOrigen = 1 Then
        
            LblPeriodo2.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
            Nuevo
        Else
            If CONTABILIZAR = True Then
                OpcionesPeriodo
            Else
                RST_Busq RstComp, "SELECT com_honorarios.*, mae_prov.nombre,  IIf(IsNull([com_honorarios]![numser])=-1,[com_honorarios]![numdoc],[com_honorarios]![numser] & '-' & [com_honorarios]![numdoc]) AS numerodoc, " _
                    & " mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_prov.numruc, " _
                    & " mae_moneda.descripcion AS descmon, mae_moneda.simbolo, mae_tipoproducto.descripcion AS desctipcom, " _
                    & " con_tc.impcom, IIf([com_honorarios].[tc]=0,[con_tc].[impven],[com_honorarios].[tc]) & '' AS impven1 ,com_honorarios.impbru & '' as impbru1,com_honorarios.impigv & '' as impigv1  " _
                    & " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_condpago RIGHT JOIN ((com_honorarios LEFT " _
                    & " JOIN mae_tipoproducto ON com_honorarios.idtipo = mae_tipoproducto.id) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) " _
                    & " ON mae_condpago.id = com_honorarios.idconpag) ON mae_documento.id = com_honorarios.tipdoc) ON mae_moneda.id = com_honorarios.idmon) " _
                    & " ON mae_prov.id = com_honorarios.idpro", xCon
            End If
            Set Rst = Nothing
            
            Set Dg1.DataSource = RstComp
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then ' F3 Nuevo
        If fCierrePeriodo = True Then Exit Sub
        
        If QueHace <> 3 Then Exit Sub
        
        Nuevo
    End If
    
    If KeyCode = 115 Then ' F4 Modificar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace <> 3 Then Exit Sub
        If RstComp.RecordCount = 0 Then Exit Sub
        Modificar
    End If
    
    If KeyCode = 113 Then ' F2 Grabar
        If fCierrePeriodo = False Then Exit Sub
        
        If QueHace = 3 Then Exit Sub
        
        If Grabar = True Then
            If xOrigen = 0 Then
                Cancelar
                RstComp.Requery
                Dg1.Refresh
            Else
                QueHace = 3
                Set RstComp = Nothing
                Unload Me
                Exit Sub
            End If
        End If
    End If
    
    'If KeyCode = 116 Then ' F5 actualizar
    
    If KeyCode = 117 Then ' F6 cancelar
        If fCierrePeriodo = False Then Exit Sub
        
        If QueHace = 3 Then Exit Sub
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE AL CARGAR EL FORMULARIO
    QueHace = 3
    TabOne1.CurrTab = 0
    SeEjecuto = False
    
    Dg1.Columns("fchdoc1").NumberFormat = FORMAT_DATE
    Dg1.Columns("fchven1").NumberFormat = FORMAT_DATE
    
    Dg1.Columns("impbru1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("impigv1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("imptot1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("impsal1").NumberFormat = FORMAT_MONTO
    
    CaracteresNumericos = "0123456789." & Chr(8)
    CaracteresNumericos2 = "0123456789.-" & Chr(8) & Chr(13)
    
    Fg4.ColWidth(5) = 0
    Fg5.ColWidth(5) = 0
    
    Fg1.ColWidth(1) = 4500 + (1005 * 2)
    Fg1.ColWidth(2) = 0
    Fg1.ColWidth(3) = 0
    Fg1.ColWidth(5) = 0
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
    Fg1.ColWidth(12) = 0
    Fg1.ColWidth(13) = 0
    Fg1.ColWidth(14) = 0
    Fg1.ColWidth(15) = 0
    Fg1.ColWidth(16) = 0
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    
    ChkImpRen4.Value = 1
    
    LblIgvTasa.Caption = ""
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ColComboList(1) = ""
    If CONTABILIZAR = True Then
        Toolbar1.Buttons(11).Visible = True
        LblPeriodo.Visible = True
        Frame5.Visible = True
    Else
        Toolbar1.Buttons(11).Visible = False
        LblPeriodo.Visible = False
        Frame5.Visible = False
    End If
    
    Fg4.SelectionMode = flexSelectionByRow
    Fg4.Editable = flexEDNone
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando una compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Menu1_1_Click()
    CmdAddItem_Click
End Sub

Private Sub menu1_1_1_Click()
    CmdAddItem_Click
End Sub

Private Sub Menu1_3_Click()
    CmdDelItem_Click
End Sub

'Private Sub Menu1_5_Click()
'    Command1_Click
'End Sub

Private Sub opciones_1_Click()
    AdjuntarEntradas 1
End Sub

Private Sub opciones_2_Click()
    AdjuntarEntradas 2
End Sub

Private Sub OptDes1_Click()
    ' ESPECIFICA QUE EL DESCUENTO SERA EN PORCENTAJE
    If OptDes1.Value = True Then
        Fg1.TextMatrix(0, 6) = " Dsct en %"
        
        Dim A As Integer
        For A = 1 To Fg1.Rows - 1
            Fg1_CellChanged A, 3
        Next A
    End If
End Sub

Private Sub OptDes2_Click()
    ' ESPECIFICA QUE EL DESCUENTO SERA EN IMPORTE
    If OptDes2.Value = True Then
        Fg1.TextMatrix(0, 6) = "Dsct en Imp."
        
        Dim A As Integer
        For A = 1 To Fg1.Rows - 1
            Fg1_CellChanged A, 3
        Next A
    End If
End Sub

Private Sub OptNo_Click()
    If OptNo.Value = True Then HallarTotal
End Sub

Private Sub OptOpera1_Click()
    ' ESPECIFICA QUE SE ESTA REGISTRANDO UNA COMPRA NORMAL
    If OptOpera1.Value = True Then
        Fg1.Rows = 1
        Fg4.Rows = 1
        CmdSeleccionar.Caption = "Seleccionar Item"
        CmdAddItem.Enabled = True
        CmdDelItem.Enabled = True
    End If
End Sub

Private Sub OptOpera2_Click()
    ' ESPECIFICA QUE SE ESTA REGISTRANDO UNA COMPRA AMARRADA A UNO O VARIOS DOCUMENTOS DE INGRESO
    If OptOpera2.Value = True Then
        CmdSeleccionar.Caption = "Ver Documentos"
        CmdAddItem.Enabled = False
        CmdDelItem.Enabled = False
    End If
End Sub

Private Sub OptOpera3_Click()
    ' ESPECIFICA QUE SE ESTA REGISTRANDO UNA COMPRA AMARRADA A UNA ORDEN DE COMPRA
    If OptOpera3.Value = True Then
        CmdSeleccionar.Caption = "Ver Documentos"
        CmdAddItem.Enabled = False
        CmdDelItem.Enabled = False
    End If
End Sub

Private Sub OptSi_Click()
    If OptSi.Value = True Then HallarTotal
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If xOrigen = 0 Then
            If RstComp.RecordCount = 0 Then Exit Sub
            If QueHace = 3 Then MuestraSegundoTab
        End If
    End If
End Sub

Sub Filtrar()
    ' PERMITE AFECTUAR UN FILTRO EN EL RECORDSET RstComp
    Dim xform As New eps_librerias.FormFiltrar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(7, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Tipo Documento":     xCampos(0, 1) = "abrev":         xCampos(0, 2) = "C":         xCampos(0, 3) = "1500"
    xCampos(1, 0) = "Moneda":             xCampos(1, 1) = "simbolo":       xCampos(1, 2) = "C":         xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Fch. Emi.":          xCampos(2, 1) = "fchdoc":        xCampos(2, 2) = "F":         xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Proveedor":            xCampos(3, 1) = "nombre":      xCampos(3, 2) = "C":         xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Forma Pago":         xCampos(4, 1) = "desccond":      xCampos(4, 2) = "C":         xCampos(4, 3) = "1500"
    xCampos(5, 0) = "Fch. Vencimiento":   xCampos(5, 1) = "fchven":        xCampos(5, 2) = "C":         xCampos(5, 3) = "1500"
    xCampos(6, 0) = "Importe":            xCampos(6, 1) = "imptot":        xCampos(6, 2) = "C":         xCampos(6, 3) = "1500"
    xCampos(7, 0) = "Saldo":              xCampos(7, 1) = "impsal":        xCampos(7, 2) = "C":         xCampos(7, 3) = "1500"
    
    Set xform.Coneccion = xCon
    Set xform.Rst = RstComp   'recorset que llena el grid
    Set RstComp = xform.FiltrarReg(xCampos)
    Set Dg1.DataSource = RstComp
    Dg1.Refresh
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then
        ' preguntamos si la compra esta vinculada a una orden de compra
        If RstComp("idordcom") <> 0 Then
            ' no se puede modificar una compra que tenga un orden de compra asignada
            MsgBox "La compra no se puede modificar por tener una Orden de Compra asignada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            If xOrigen = 0 Then
                Cancelar
                RstComp.Requery
                Dg1.Refresh
                
                If RstComp.RecordCount <> 0 Then
                    RstComp.MoveFirst
                    RstComp.Find "id=" & mIdRegistro
                    If RstComp.EOF = True Then RstComp.MoveFirst
                End If
                
            Else
                QueHace = 3
                Set RstComp = Nothing
                Unload Me
                Exit Sub
            End If
        End If
    End If
    
    If Button.Index = 6 Then
        If xOrigen = 1 Then
            QueHace = 3
            IdCompraReg = 0
            Unload Me
            Exit Sub
        End If
        
        Cancelar
    End If
    
    If Button.Index = 8 Then
        Filtrar
    End If
    
    If Button.Index = 9 Then
'        TabOne1.CurrTab = 0
'        TDB_FiltroLimpiar Dg1
'        RstComp.Filter = ""
        TDB_Actualizar Me, TabOne1, Dg1, RstComp
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 11 Then
        Dim xMesProv As Integer
        xMesProv = mMesActivo
        mMesActivo = SeleccionaMes(xCon)
        OpcionesPeriodo
    End If
    
    If Button.Index = 13 Then pExportar
    
    
    If Button.Index = 14 Then
        If RstComp("tipdoc") = 4 Then
            Imprimir
        Else
            MsgBox "No puede imprimir este documento, seleccione una liquidación de compras para efectuar esta operación", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    
    If Button.Index = 16 Then
        Set RstComp = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : OpcionesPeriodo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LAS OPCIONES ACTIVAS DEL PERIODO ESPECIFICADO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub OpcionesPeriodo()
'Modificado 12/01/11 Johan Castro
'           Agregar envío de parametro xIdUsuario a procidimiento CierrePeriodo

    Dim NomMes As String
    Dim Cerrado As Boolean
    Dim Rpta  As Integer
    Dim xFechaMes As String
    Dim xFchIni, xFchFin As Date
    
    If mMesActivo = 0 Then CmdApertura.Visible = True Else CmdApertura.Visible = False
        
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    TDB_FiltroLimpiar Dg1
    Set RstComp = Nothing
    '------------------------------------------------------------------------------------------
    
    LblPeriodo.Caption = LblMes.Caption
    LblPeriodo2.Caption = LblMes.Caption
    
    xFechaMes = "01/" + Trim(Format(mMesActivo, "00")) + "/" + Trim(Format(AnoTra, "0000"))
    
    Set RstComp = Nothing
    
    CargarRSTCom
    
    Set Dg1.DataSource = RstComp
    
    TabOne1.CurrTab = 0
    Dg1.SetFocus
    
End Sub

Private Sub TxtBruto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtBruto_Validate(Cancel As Boolean)
    If NulosN(TxtBruto.Text) <> 0 Then
        TxtBruto.Text = Format(TxtBruto.Text, FORMAT_MONTO)
        TxtIGV.Text = NulosN(TxtBruto.Text) * xPorIgv
        TxtIGV.Text = Format(TxtIGV.Text, FORMAT_MONTO)
    End If
End Sub

Private Sub TxtConPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtConPag_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCondicion_Click
    End If
End Sub

Private Sub TxtConPag_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtConPag.Text) = "" Then Exit Sub
    Dim xRs1 As New ADODB.Recordset
    
    RST_Busq xRs1, "SELECT * FROM mae_condpago WHERE id = " & NulosN(TxtConPag.Text) & "", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtConPag.Text = ""
        LblCondPag.Caption = ""
    Else
        LblCondPag.Caption = Trim(xRs1("descripcion"))
        If NulosC(TxtFchDoc.Valor) <> "" Then
            TxtFchVen.Valor = CDate(TxtFchDoc.Valor) + xRs1("numdia")
        End If
    End If
    Set xRs1 = Nothing
End Sub

Private Sub TxtDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDocRef_Click
    End If
End Sub

Private Sub TxtDocRef2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Fg1.Rows = 1 Then
            CmdAddItem.SetFocus
        Else
            Fg1.Row = 1
            Fg1.Col = 1
            Fg1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDocRef2_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 116 Then
        CmdBusDocRef2_Click
    End If
    If KeyCode = 46 Then
        TxtDocRef2.Text = ""
        LblIdDocRef2.Caption = ""
    End If
End Sub

Private Sub TxtFchDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtFchDoc.Valor) <> "" Then
        Dim xRs1 As New ADODB.Recordset
        
        RST_Busq xRs1, "SELECT * FROM mae_condpago WHERE id = " & NulosN(TxtConPag.Text) & "", xCon
        
        If xRs1.RecordCount = 0 Then
            TxtConPag.Text = ""
            LblCondPag.Caption = ""
        Else
            If NulosC(TxtFchDoc.Valor) <> "" Then
                TxtFchVen.Valor = CDate(TxtFchDoc.Valor) + xRs1("numdia")
            End If
        End If
        Set xRs1 = Nothing
    End If
End Sub

Private Sub TxtIdAlmacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdAlmacen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusAlm_Click
    End If
End Sub

Private Sub TxtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtIdMon.Text) = "" Then Exit Sub
    Dim xRs1 As New ADODB.Recordset
    
    'buscamos el codigo de la moneda digitada
    RST_Busq xRs1, "SELECT * FROM mae_moneda WHERE id = " & NulosN(TxtIdMon.Text) & "", xCon
    
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
            LblTipoCambio.Caption = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
        End If
    End If
    Set xRs1 = Nothing
    xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
End Sub

Private Sub TxtIGV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIGV2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIGV3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtInafecto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtISC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtNumDoc.Text) <> "" Then
        TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
        
        If NulosC(TxtNumDoc.Text) <> "" And NulosC(TxtNumSer.Text) <> "" Then
            If ExisteNumDocCompra = True Then
                Exit Sub
            End If
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ExisteNumDocCompra
'* Tipo             : FUNCION
'* Descripcion      : VALIDA QUE EL NUEVO NUMERO DE DOCUMENTO INGRESADO NO EXISTA COMO COMPRA
'*                    REGISTRADA, ESTA FUNCION DEVUELVER VERDADERO SI ENCUENTRA EL REGISTRO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function ExisteNumDocCompra() As Boolean
    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
    If QueHace <> 1 Then nSQL = " and com_honorarios.id <> " & NulosN(RstComp("id"))
    
    RST_Busq Rst, "SELECT com_honorarios.fchdoc, Left([com_honorarios].[numreg],2) & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & Right([com_honorarios].[numreg],4) AS registro FROM com_honorarios LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id WHERE numser = '" & NulosC(TxtNumSer.Text) & "' and numdoc = '" & NulosC(TxtNumDoc.Text) & "' AND idpro = " & NulosN(LblIdProveedor.Caption) & nSQL, xCon
    If Rst.RecordCount = 0 Then
        ExisteNumDocCompra = False
    Else
        MsgBox "El número de documento ingresado ya fue registrado" & vbCr & "Nº Registro: " & NulosC(Rst("registro")) & vbCr & "Fecha Doc.   " & NulosC(Rst("fchdoc")) & vbCr & "Ingrese Otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.Text = ""

        ExisteNumDocCompra = True
    End If
    Set Rst = Nothing
End Function

Private Sub TxtNumRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumRuc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub

Private Sub TxtNumRuc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If TxtNumRuc.Text = "" Then Exit Sub
    Dim xRs1 As New ADODB.Recordset
    RST_Busq xRs1, "SELECT * FROM mae_prov WHERE numruc like '" & TxtNumRuc.Text & "%' ORDER BY numruc", xCon
    If xRs1.RecordCount <> 0 Then
        TxtNumRuc.Text = xRs1("numruc")
        LblNomPro.Caption = xRs1("nombre")
        LblIdProveedor.Caption = xRs1("id")
        If xRs1("idcondpag") <> 0 Then
            TxtConPag.Text = xRs1("idcondpag")
            TxtConPag_Validate True
        End If
    Else
        TxtNumRuc.Text = ""
        LblNomPro.Caption = ""
        LblIdProveedor.Caption = ""
    End If
    Set xRs1 = Nothing
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtNumSer.Text) <> "" Then
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
        
        If NulosC(TxtNumDoc.Text) <> "" And NulosC(TxtNumSer.Text) <> "" Then
            If ExisteNumDocCompra = True Then
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub TxtTipCom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipCom_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipoCompra_Click
    End If
End Sub

Private Sub TxtTipCom_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtTipCom.Text) <> "" Then
        Set RstTmp = BuscaConCriterio("SELECT * FROM mae_tipoproducto WHERE id = " & NulosN(TxtTipCom.Text) & "", xCon)
        If RstTmp.RecordCount <> 0 Then
            LblTipoCompra.Caption = RstTmp("descripcion")
        Else
            TxtTipCom.Text = ""
            LblTipoCompra.Caption = ""
        End If
    End If
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : PERMITE GRABAR UN REGISTRO EN LA TABAL com_honorarios, DEVUELVE VERDADERO SI
'*                    TIENE EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim A, B, Rpta As Integer
    
    If NulosN(TxtTipDoc.Text) <> 0 Then
        ' VERIFICAMOS QUE EL DOCUMENTO DE COMPRA, TENGA ASIGNADO UNA CUENTA CONTABLE
        If xCuentaDoc = 0 Then
            MsgBox "No se ha asignado una cuenta contable al documento " + LblNomDoc.Caption & Chr(13) _
                & "Asignele una cuenta en el menu Contabilidad opcion Asignar Ctas. Contables a documentos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    
        ' VERIFICAMOS QUE EL IMPUESTO ASIGNADO AL DOCUMENTO TENGA UNA CUENTA CONTABLE
        If xIdCuenTasa = 0 Then
            MsgBox "El impuesto asignado al documento " + LblNomDoc.Caption & Chr(13) & " no tiene cuenta contable" & Chr(13) _
                & "Asignele una cuenta en el menu Contabilidad opcion Maestro de Impuestos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    End If
    
    Dim Rst As New ADODB.Recordset
    
    For A = 1 To Fg1.Rows - 1
        ' validamos que el precio ingresado este en un rango de precios especificado
        RST_Busq Rst, "SELECT * FROM com_precios WHERE idpro = " & NulosN(Fg1.TextMatrix(A, 8)) & "", xCon
        If Rst.RecordCount <> 0 Then
            If NulosN(Fg1.TextMatrix(A, 4)) > NulosN(Rst("pretop")) Then
                Set Rst = Nothing
                ' buscamos una autorizacion de ingreso para el precio del proveedor
                RST_Busq Rst, "SELECT com_preciosdet.idpro, com_preciosdet.fecreg, com_preciosdet.idprov, com_preciosdet.precio" _
                    & " From com_preciosdet  " _
                    & " WHERE (((com_preciosdet.idpro)=" & NulosN(Fg1.TextMatrix(A, 8)) & ") AND ((com_preciosdet.fecreg)=CDate('" & Format(TxtFchDoc.Valor, "dd/mm/yyyy") & " ')) " _
                    & " AND ((com_preciosdet.idprov)=" & NulosN(LblIdProveedor.Caption) & "))", xCon
                
                If Rst.RecordCount = 0 Then
                    ' si no encontramos una autorizacion de precio para el proveedor en el dia de la operacion se rechaza
                    MsgBox "El precio ingresado para el item " + NulosC(Fg1.TextMatrix(A, 1)) & Chr(13) _
                        & "excede el precio fijado por el administrador de precios, verifique el precio fijado" & Chr(13) _
                        & "en el modulo de Compras opcion  Fijar Precios de Compra a Item", vbCritical + vbOKOnly + vbDefaultButton1, xTitulo
                    Set Rst = Nothing
                    Exit Function
                Else
                    If NulosN(Fg1.TextMatrix(A, 4)) > NulosN(Rst("precio")) Then
                        ' si el precio ingresado es aun mayor que el precio autorizado se rechaza la compra
                        MsgBox "El precio ingresado para el item " + NulosC(Fg1.TextMatrix(A, 1)) & Chr(13) _
                            & "excede el precio fijado por el administrador de precios, verifique el precio fijado" & Chr(13) _
                            & "en el modulo de Compras opcion  Fijar Precios de Compra a Item", vbCritical + vbOKOnly + vbDefaultButton1, xTitulo
                        Set Rst = Nothing
                        Exit Function
                    End If
                End If
            End If
        End If
        
        Set Rst = Nothing
        
        'validamos que el ingreso de items no exceda el stock maximo
        'validamos la cuenta contable del item
        If NulosN(Fg1.TextMatrix(A, 10)) = 0 Then
            MsgBox "No se le ha asignado una Cuenta Contable para Venta al item : " & Chr(13) _
                & Fg1.TextMatrix(A, 1) & Chr(13) _
                & "Asígnele una cuenta en el menu Almacén opción Mantenimiento Items de Compra y Venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    Next A
        
    If TxtTipCom.Text = "" Then
        MsgBox "No ha especificado el Tipo de Compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipCom.SetFocus
        Exit Function
    End If
    
    If TxtNumRuc.Text = "" Then
        MsgBox "No ha especificado proveedor de la compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
        MsgBox "No ha especificado items para la compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.Col = 1
        Fg1.SetFocus
        Exit Function
    End If
    
    ' verificamos que la fecha de vencimiento no sea menor a la fecha de vencimiento
    If CDate(TxtFchDoc.Valor) > CDate(TxtFchVen.Valor) Then
        MsgBox "La fecha de vencimiento del documento no puede ser menor a la fecha de emision", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchVen.SetFocus
        Exit Function
    End If
    
    ' verificamos que la fecha de vencimiento no sea mayor al periodo contable
    If CDate(TxtFchVen.Valor) > (CDate(xFchFin)) Then
        If NulosC(TxtFchPago.Valor) = "" Then
            MsgBox "No puede registrar este documento en el mes de " + Trim(LblPeriodo.Caption) + ", la fecha de " & Chr(13) _
                & "vencimiento es mayor a la fecha del periodo, para registrar este documento en el periodo" & Chr(13) _
                & "actual ingrese la fecha de pago menor o igual a la fecha de cierre del periodo ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    End If
        
    ' VERIFICAMOS QUE LOS ITEMS IGRESADOS SON LOS CORRECTOS
    ' VERIFICAMOS QUE NO EXISTAS FILAS SIN ITEMAS
    For A = 1 To Fg1.Rows - 1
        If NulosC(Fg1.TextMatrix(A, 1)) = "" Then
            Fg1.RemoveItem A
        End If
    Next A
    
    If Fg1.Rows <> 1 Then
        For A = 1 To Fg1.Rows - 1
            If NulosN(Fg1.TextMatrix(A, 3)) = 0 Then
                MsgBox "No ha especificado la cantidad para el item : " + Trim(Fg1.TextMatrix(A, 1)), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Fg1.Col = 3: Fg1.Row = A
                Fg1.SetFocus
                Exit Function
            End If
        Next A
    Else
        MsgBox "No se ha especificado ningun item para esta compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    
    ' verificamos que el total de items sea igual al total de los totales
    A = NulosN(Format(GRID_SUMAR_COL(Fg1, 8), "0.0000"))
    B = NulosN(TxtBruto.Text) '+ NulosN(TxtBruto2.Text) + NulosN(TxtBruto3.Text) + NulosN(TxtInafecto.Text)
    If Round(A, 2) <> Round(B, 2) Then
        MsgBox "El monto del detalle del documento no coincide con la sumatoria de los totales", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtBruto.SetFocus
        Exit Function
    End If
    
    Dim RstDeta2 As New ADODB.Recordset
    Dim RstActPro As New ADODB.Recordset
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
'    Dim RstDia As New ADODB.Recordset
    Dim RstCosto As New ADODB.Recordset
    
    Dim xIdCuen As Integer
    Dim xTotal As Double
    Dim xNumAsiento As String
    Dim xId As Double
    Dim nSQL As String
    
    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        ' SI SE ESTA AGREGANDO UN NUEVO REGISTRO
        xId = HallaCodigoTabla("com_honorarios", xCon, "id")
        
        xNumAsiento = NuevoNumAsiento(40, mMesActivo, xCon)
        
        RST_Busq RstCab, "SELECT TOP 1 * FROM com_honorarios", xCon
        RST_Busq RstDet, "SELECT TOP 1 * FROM com_honorariosdet", xCon
'        RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
        RST_Busq RstCosto, "SELECT TOP 1 * FROM com_honorarioscosto", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
        IdCompraReg = xId
    Else
        ' SI SE ESTA MODIFICANDO UN REGISTRO
        xId = RstComp("id")
        RST_Busq RstCab, "SELECT * FROM com_honorarios WHERE id = " & xId & "", xCon
        
        ' eliminamos el sotck agregado con la compra
        RST_Busq RstDeta2, "SELECT com_honorariosdet.* From com_honorariosdet WHERE (((com_honorariosdet.idhon)=" & xId & "))", xCon

        If RstDeta2.RecordCount <> 0 Then
            RstDeta2.MoveFirst
            For A = 1 To RstDeta2.RecordCount
                RST_Busq RstActPro, "SELECT alm_inventario.id, alm_inventario.stckact  From alm_inventario WHERE ((alm_inventario.id=" & RstDeta2("iditem") & "))", xCon
                If RstActPro.RecordCount = 1 Then
                    RstActPro("stckact") = RstActPro("stckact") - RstDeta2("canpro")
                    RstActPro.Update
                End If
                Set RstActPro = Nothing
            Next A
        End If
        Set RstDeta2 = Nothing
        
        ' eliminamos el detalle de la compra
        xCon.Execute "DELETE * FROM com_honorariosdet WHERE idhon = " & xId & ""
        RST_Busq RstDet, "SELECT TOP 1 * FROM com_honorariosdet", xCon
        
'''        If mMesActivo = 0 Then
'''            RST_Busq RstDia, "SELECT * FROM con_diario WHERE idmes = " & mMesActivo & " AND idlib = 36 AND idmov = " & xId & "", xCon
'''        Else
'''            RST_Busq RstDia, "SELECT * FROM con_diario WHERE idmes = " & mMesActivo & " AND idlib = 1 AND idmov = " & xId & "", xCon
'''        End If
'''
'''        If RstDia.RecordCount <> 0 Then
'''            xNumAsiento = RstDia("numasi")
'''        End If
        
'''        Set RstDia = Nothing
'''
'''        ' eliminamos el asiento contable
'''        If mMesActivo = 0 Then
'''            xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & mMesActivo & " AND idlib = 36 AND idmov = " & xId & ""
'''        Else
'''            xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & mMesActivo & " AND idlib = 40 AND idmov = " & xId & ""
'''        End If
''
'''        RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon

        ' Eliminamos el centro de costos
        xCon.Execute "DELETE * FROM com_honorarioscosto WHERE idcom = " & xId & ""
        
        RST_Busq RstCosto, "SELECT TOP 1 * FROM com_honorarioscosto", xCon
        
        xNumAsiento = Mid(RstComp("numreg"), 3, 4)
        
        'Borramos los flag de las tablas alm_ingreso y com_ordencompra
        If OptOpera3 = True Then
            'actualizamos campo idfac en la tabla alm_igreso a 0 para que se vuelva a procesar
            xCon.Execute "DELETE * FROM alm_ingresodoc WHERE iddoc = " & RstComp("id") & " "
        End If
    
        If OptOpera2 = True Then
            'actualizamos campo idfac en la tabla com_ordencompra a 0 para que se vuelva a procesar
            xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idfac = 0 WHERE (((com_ordencompra.idfac)=" & RstComp("id") & "))"
        End If
    End If
    
    mIdRegistro = xId
    
    'GRABAMOS LOS DATOS DEL NUEVO REGISTRO
    RstCab("idlib") = 40
    RstCab("idtipo") = NulosN(TxtTipCom.Text)
    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("idpro") = NulosN(LblIdProveedor.Caption)
    RstCab("numser") = TxtNumSer.Text
    RstCab("numdoc") = TxtNumDoc.Text
    RstCab("fchdoc") = TxtFchDoc.Valor
    RstCab("fchven") = TxtFchVen.Valor
    If NulosC(TxtFchPago.Valor) <> "" Then RstCab("fchpag") = TxtFchPago.Valor
    RstCab("idconpag") = NulosN(TxtConPag.Text)
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    RstCab("impbru") = NulosN(TxtBruto.Text)
    RstCab("impigv") = NulosN(TxtIGV.Text)
    RstCab("imptot") = NulosN(TxtTotal.Text)
    RstCab("glosa") = NulosC(TxtGlosa.Text)
    RstCab("impsal") = NulosN(TxtTotal.Text)
    
    RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    
    If NulosN(TxtTipDocRef.Text) <> 0 Then
        RstCab("idtipdocref") = NulosN(TxtTipDocRef.Text)
        RstCab("iddocref2") = NulosN(LblIdDocRef2.Caption)
    Else
        RstCab("idtipdocref") = 0
        RstCab("iddocref2") = 0
    End If
    
    If CONTABILIZAR = True Then
        RstCab("numreg") = Format(Trim(Str(mMesActivo)), "00") + xNumAsiento
    End If
    
    ' grabamos el tipo de descuento
    If OptDes1.Value = True Then
        RstCab("tipdes") = 1
    End If
    If OptDes2.Value = True Then
        RstCab("tipdes") = 2
    End If
    
    If OptSi.Value = True Then
        RstCab("afecto") = -1
    Else
        RstCab("afecto") = 0
    End If
    
    ' especificamos como en que contexto se esta haciendo la compra
    If OptOpera1.Value = True Then RstCab("tipcom") = 1  'Compra normal
    If OptOpera3.Value = True Then RstCab("tipcom") = 2  'Compra vinculada con documentos de entrada
    If OptOpera2.Value = True Then RstCab("tipcom") = 3  'Compra vinculada con Orden de Compra
    
    RstCab.Update
    
    ' Grabamos los items de la compra
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idhon") = xId
        RstDet("iditem") = NulosN(Fg1.TextMatrix(A, 9))
        RstDet("idunimed") = NulosN(Fg1.TextMatrix(A, 10))
        RstDet("canpro") = NulosN(Fg1.TextMatrix(A, 3))
        RstDet("preunibru") = NulosN(Fg1.TextMatrix(A, 4))    ' precio bruto afecto
        RstDet("preunibruina") = NulosN(Fg1.TextMatrix(A, 5)) ' precio bruto inafecto
        RstDet("valdes") = NulosN(Fg1.TextMatrix(A, 6))
        RstDet("preuni") = NulosN(Fg1.TextMatrix(A, 7))
        RstDet("imptot") = NulosN(Fg1.TextMatrix(A, 8))
        RstDet.Update
        
        If NulosN(TxtTipCom.Text) = 1 Or NulosN(TxtTipCom.Text) = 4 Or NulosN(TxtTipCom.Text) = 2 Then
            RST_Busq RstActPro, "SELECT alm_inventario.id, alm_inventario.stckact  From alm_inventario WHERE (((alm_inventario.id)=" & NulosN(NulosN(Fg1.TextMatrix(A, 8))) & "))", xCon

            If RstActPro.RecordCount = 1 Then
                RstActPro("stckact") = NulosN(RstActPro("stckact")) + NulosN(Fg1.TextMatrix(A, 3))
                RstActPro.Update
            End If
            Set RstActPro = Nothing
        End If
    Next A
    
    ' Actualizamos los documentos relacionados con la factura
    If OptOpera3.Value = True Then
        If Fg4.Rows <> 1 Then
            For A = 1 To Fg4.Rows - 1
                ' actualizamos el flag de los partes de entrada para saber con que documento de compra se valorizaran
                xCon.Execute "INSERT INTO alm_ingresodoc (id, iddoc) values (" & NulosN(Fg4.TextMatrix(A, 5)) & "," & xId & ")"
            Next A
        End If
    End If
    
    If OptOpera2.Value = True Then
        If Fg4.Rows <> 3 Then
            For A = 1 To Fg4.Rows - 1
                ' actualizamos el flag de las ordenes de compra para saber con que documento ingresaron las ordenes de compra
                xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idfac = " & xId & " WHERE (((com_ordencompra.id)=" & NulosN(Fg4.TextMatrix(A, 5)) & "))"
            Next A
        End If
    End If
    
    ' grabamos el centro de costos
    If Fg5.Rows > 1 Then
        For A = 1 To Fg5.Rows - 1
            RstCosto.AddNew
            RstCosto("idcom") = xId
            RstCosto("idcencos") = NulosN(Fg5.TextMatrix(A, 5))
            RstCosto("imppor") = NulosN(Fg5.TextMatrix(A, 3))
            RstCosto("impcos") = NulosN(Fg5.TextMatrix(A, 4))
            RstCosto.Update
        Next A
    End If
    
    ' En caso de estar vinculada a una orden de compra actualizamos la orden de compra "3 = PROCESADA"
'''    If CONTABILIZAR = True Then
'''        ' Grabamos el libro diario del movimiento
'''
'''        ' grabamos a facturas por pagar Plan de cuentas 42.1 o dependiendo del caso
'''        RstDia.AddNew
'''        RstDia("año") = AnoTra
'''        RstDia("idmes") = mMesActivo
'''        RstDia("idlib") = 40
'''        RstDia("idmov") = xId
'''        RstDia("numasi") = xNumAsiento
'''        RstDia("tc") = ValTipCam
'''        RstDia("idcue") = xCuentaDoc
'''        RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''        RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''        If NulosN(TxtTipDoc.Text) <> 0 Then
'''            If NulosN(TxtTipDoc.Text) <> 7 Then
'''                ' cuando se factura u otro comprabante excepto nota de credito hace su asiento norma
'''                If TxtIdMon.Text = "1" Then
'''                    RstDia("imphabsol") = Format(NulosN(TxtTotal.Text), "0.000000")
'''                    RstDia("imphabdol") = 0
'''                Else
'''                    RstDia("imphabsol") = Format(NulosN(TxtTotal.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'''                    RstDia("imphabdol") = Format(NulosN(TxtTotal.Text), "0.000000")
'''                End If
'''            Else
'''                ' cuando sea nota de credito hace el asiento inverso al de una venta
'''                If TxtIdMon.Text = "1" Then
'''                    RstDia("impdebsol") = Format(NulosN(TxtTotal.Text), "0.000000")
'''                    RstDia("impdebdol") = 0
'''                Else
'''                    RstDia("impdebsol") = Format(NulosN(TxtTotal.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'''                    RstDia("impdebdol") = Format(NulosN(TxtTotal.Text), "0.000000")
'''                End If
'''            End If
'''        End If
'''        RstDia.Update
'''
'''        ' grabamos el impuesto si la operacion esta afecta a el
'''        If NulosN(TxtIGV.Text) <> 0 Then
'''            RstDia.AddNew
'''            RstDia("año") = AnoTra
'''            RstDia("idmes") = mMesActivo
'''            If mMesActivo = 0 Then
'''                RstDia("idlib") = 36
'''            Else
'''                RstDia("idlib") = 40
'''            End If
'''            RstDia("idmov") = xId
'''            RstDia("numasi") = xNumAsiento
'''            RstDia("tc") = ValTipCam
'''            RstDia("idcue") = xIdCuenTasa
'''            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''            RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''            If NulosN(TxtTipDoc.Text) <> 0 Then
'''                If NulosN(TxtTipDoc.Text) <> 7 And NulosN(TxtTipDoc.Text) <> 2 Then
'''                    If TxtIdMon.Text = "1" Then
'''                        RstDia("impdebsol") = Format(NulosN(TxtIGV.Text), "0.000000")
'''                        RstDia("impdebdol") = 0
'''                    Else
'''                        RstDia("impdebsol") = Format(NulosN(TxtIGV.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'''                        RstDia("impdebdol") = Format(NulosN(TxtIGV.Text), "0.000000")
'''                    End If
'''                Else
'''                    If TxtIdMon.Text = "1" Then
'''                        RstDia("imphabsol") = Format(NulosN(TxtIGV.Text), "0.000000")
'''                        RstDia("imphabdol") = 0
'''                    Else
'''                        RstDia("imphabsol") = Format(NulosN(TxtIGV.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'''                        RstDia("imphabdol") = Format(NulosN(TxtIGV.Text), "0.000000")
'''                    End If
'''                End If
'''            End If
'''            RstDia.Update
'''        End If
'''
'''        ' grabamos el imponible en function a los items de la factura
'''        Set Rst = Nothing
'''        RST_Busq Rst, "SELECT com_honorariosdet.idhon, alm_inventario.idcuenta, Sum(com_honorariosdet.imptot) AS SumaDeimptot FROM alm_inventario INNER JOIN com_honorariosdet " _
'''            & " ON alm_inventario.id = com_honorariosdet.iditem GROUP BY com_honorariosdet.idhon, alm_inventario.idcuenta HAVING (((com_honorariosdet.idhon)=" & xId & "))", xCon
'''
'''        If Rst.RecordCount <> 0 Then
'''            Rst.MoveFirst
'''            For A = 1 To Rst.RecordCount
'''                RstDia.AddNew
'''                RstDia("año") = AnoTra
'''                RstDia("idmes") = mMesActivo          ' LLAVE - CODIGO DEL MES
'''                If mMesActivo = 0 Then
'''                    RstDia("idlib") = 36              ' LLAVE - CODIGO DEL LIBRO
'''                Else
'''                    RstDia("idlib") = 40
'''                End If
'''                RstDia("idmov") = xId                 ' LLAVE - CODIGO DEL MOVIMIENTO
'''                RstDia("numasi") = xNumAsiento        ' LLAVE - NUMERO DE ASIENTO
'''                RstDia("tc") = ValTipCam
'''                RstDia("idcue") = Rst("idcuenta")
'''                RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''                RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''                If NulosN(TxtTipDoc.Text) <> 0 Then
'''                    If NulosN(TxtTipDoc.Text) <> 7 Then
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("impdebsol") = Format(Rst("SumaDeimptot"), "0.000000")
'''                            RstDia("impdebdol") = 0
'''                        Else
'''                            RstDia("impdebsol") = Format(Rst("SumaDeimptot"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.0000000")
'''                            RstDia("impdebdol") = Format(Rst("SumaDeimptot"), "0.000000")
'''                        End If
'''                    Else
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("imphabsol") = Format(Rst("SumaDeimptot"), "0.000000")
'''                            RstDia("imphabdol") = 0
'''                        Else
'''                            RstDia("imphabsol") = Format(Rst("SumaDeimptot"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.0000000")
'''                            RstDia("imphabdol") = Format(Rst("SumaDeimptot"), "0.000000")
'''                        End If
'''                    End If
'''                End If
'''                RstDia.Update
'''
'''                Rst.MoveNext
'''                If Rst.EOF = True Then Exit For
'''            Next A
'''        End If
'''
'''        'grabamos los asientos automaticos
'''
'''        'grabamos la cuenta de destino debe
'''        Set Rst = Nothing
'''
'''        RST_Busq Rst, "SELECT com_honorariosdet.idhon, con_planctas.ctadesdeb, Sum(com_honorariosdet.imptot) AS SumaDeimptot FROM con_planctas RIGHT JOIN (alm_inventario " _
'''            & " INNER JOIN com_honorariosdet ON alm_inventario.id = com_honorariosdet.iditem) ON con_planctas.id = alm_inventario.idcuenta GROUP BY com_honorariosdet.idhon, " _
'''            & " con_planctas.ctadesdeb HAVING (((com_honorariosdet.idhon)=" & xId & "))", xCon
'''
'''        If Rst.RecordCount <> 0 Then
'''            Rst.MoveFirst
'''            For A = 1 To Rst.RecordCount
'''                If Rst("ctadesdeb") <> 0 Then
'''                    RstDia.AddNew
'''                    RstDia("año") = AnoTra
'''                    RstDia("idmes") = mMesActivo               ' LLAVE - CODIGO DEL MES
'''                    RstDia("idlib") = 40                       ' LLAVE - CODIGO DEL LIBRO
'''                    RstDia("idmov") = xId                      ' LLAVE - CODIGO DEL MOVIMIENTO
'''                    RstDia("numasi") = xNumAsiento             ' LLAVE - NUMERO DE ASIENTO
'''                    RstDia("tc") = ValTipCam
'''                    RstDia("idcue") = Rst("ctadesdeb") 'xIdCuen
'''                    RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''                    RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''                    If NulosN(TxtTipDoc.Text) <> 0 Then
'''                        If NulosN(TxtTipDoc.Text) <> 7 Then
'''                            If TxtIdMon.Text = "1" Then
'''                                RstDia("impdebsol") = Format(Rst("SumaDeimptot"), "0.000000")
'''                                RstDia("impdebdol") = 0
'''                            Else
'''                                RstDia("impdebsol") = Format(Rst("SumaDeimptot"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.0000000")
'''                                RstDia("impdebdol") = Format(Rst("SumaDeimptot"), "0.000000")
'''                            End If
'''                        Else
'''                            If TxtIdMon.Text = "1" Then
'''                                RstDia("imphabsol") = Format(Rst("SumaDeimptot"), "0.000000")
'''                                RstDia("imphabdol") = 0
'''                            Else
'''                                RstDia("imphabsol") = Format(Rst("SumaDeimptot"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.0000000")
'''                                RstDia("imphabdol") = Format(Rst("SumaDeimptot"), "0.000000")
'''                            End If
'''                        End If
'''                    End If
'''                    RstDia.Update
'''                End If
'''
'''                Rst.MoveNext
'''                If Rst.EOF = True Then Exit For
'''            Next A
'''        End If
'''
'''
'''        ' grabamos la cuenta de destino haber
'''        Set Rst = Nothing
'''
'''        RST_Busq Rst, "SELECT com_honorariosdet.idhon, con_planctas.ctadeshab, Sum(com_honorariosdet.imptot) AS SumaDeimptot FROM con_planctas RIGHT JOIN (alm_inventario " _
'''            & " INNER JOIN com_honorariosdet ON alm_inventario.id = com_honorariosdet.iditem) ON con_planctas.id = alm_inventario.idcuenta GROUP BY com_honorariosdet.idhon, " _
'''            & " con_planctas.ctadeshab HAVING (((com_honorariosdet.idhon)=" & xId & "))", xCon
'''
'''        If Rst.RecordCount <> 0 Then
'''            Rst.MoveFirst
'''            For A = 1 To Rst.RecordCount
'''                If Rst("ctadeshab") <> 0 Then
'''                    RstDia.AddNew
'''                    RstDia("año") = AnoTra
'''                    RstDia("idmes") = mMesActivo               ' LLAVE - CODIGO DEL MES
'''                    RstDia("idlib") = 40                       ' LLAVE - CODIGO DEL LIBRO
'''                    RstDia("idmov") = xId                      ' LLAVE - CODIGO DEL MOVIMIENTO
'''                    RstDia("numasi") = xNumAsiento             ' LLAVE - NUMERO DE ASIENTO
'''                    RstDia("tc") = ValTipCam
'''                    RstDia("idcue") = Rst("ctadeshab") 'xIdCuen
'''                    RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''                    RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''                    If NulosN(TxtTipDoc.Text) <> 0 Then
'''                        If NulosN(TxtTipDoc.Text) <> 7 Then
'''                            If TxtIdMon.Text = "1" Then
'''                                RstDia("imphabsol") = Format(Rst("SumaDeimptot"), "0.000000")
'''                                RstDia("imphabdol") = 0
'''                            Else
'''                                RstDia("imphabsol") = Format(Rst("SumaDeimptot"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.0000000")
'''                                RstDia("imphabdol") = Format(Rst("SumaDeimptot"), "0.000000")
'''                            End If
'''                        Else
'''                            If TxtIdMon.Text = "1" Then
'''                                RstDia("impdebsol") = Format(Rst("SumaDeimptot"), "0.000000")
'''                                RstDia("impdebdol") = 0
'''                            Else
'''                                RstDia("impdebsol") = Format(Rst("SumaDeimptot"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.0000000")
'''                                RstDia("impdebdol") = Format(Rst("SumaDeimptot"), "0.000000")
'''                            End If
'''                        End If
'''                    End If
'''                    RstDia.Update
'''                End If
'''
'''                Rst.MoveNext
'''                If Rst.EOF = True Then Exit For
'''            Next A
'''        End If
'''    End If
'''
'''    '----------------------------------------------------------
'''    ' grabamos el selectivo en funcion a los items de la factura
'''    If RstTempISC.RecordCount <> 0 Then
'''        RstTempISC.MoveFirst
'''
'''        For A = 1 To RstTempISC.RecordCount
'''            RstDia.AddNew
'''            RstDia("año") = AnoTra
'''            RstDia("idmes") = mMesActivo               ' LLAVE - CODIGO DEL MES
'''            RstDia("idlib") = 40                       ' LLAVE - CODIGO DEL LIBRO
'''            RstDia("idmov") = xId                      ' LLAVE - CODIGO DEL MOVIMIENTO
'''            RstDia("numasi") = xNumAsiento             ' LLAVE - NUMERO DE ASIENTO
'''            RstDia("tc") = ValTipCam
'''            RstDia("idcue") = RstTempISC("idcuen")
'''            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''            RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''            If NulosN(TxtTipDoc.Text) <> 0 Then
'''                If NulosN(TxtTipDoc.Text) <> 7 Then
'''                    If TxtIdMon.Text = "1" Then
'''                        RstDia("impdebsol") = Format(RstTempISC("total"), "0.000000")
'''                        RstDia("impdebdol") = 0
'''                    Else
'''                        RstDia("impdebsol") = Format(RstTempISC("total"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'''                        RstDia("impdebdol") = Format(RstTempISC("total"), "0.000000")
'''                    End If
'''                Else
'''                    If TxtIdMon.Text = "1" Then
'''                        RstDia("imphabsol") = Format(RstTempISC("total"), "0.000000")
'''                        RstDia("imphabdol") = 0
'''                    Else
'''                        RstDia("imphabsol") = Format(RstTempISC("total"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'''                        RstDia("imphabdol") = Format(RstTempISC("total"), "0.000000")
'''                    End If
'''                End If
'''            End If
'''            RstDia.Update
'''
'''            RstTempISC.MoveNext
'''
'''            If RstTempISC.EOF = True Then
'''                Exit For
'''            End If
'''        Next A
'''    End If
    
    ' AVERIGUAMOS SI EL ITEM ESTA AFECTO A LA DETRACCION
'''    RST_Busq Rst, "SELECT mae_detraccion.id, mae_detraccion.descripcion, mae_detraccion.tasa, alm_inventario.iddet " _
'''        & " FROM alm_inventario LEFT JOIN mae_detraccion ON alm_inventario.iddet = mae_detraccion.id " _
'''        & " WHERE ((alm_inventario.id= " & NulosN(Fg1.TextMatrix(Fg1.Row, 9)) & "))", xCon

    RST_Busq Rst, "SELECT mae_detraccion.id, mae_detraccion.descripcion, mae_detraccion.tasa, alm_inventario.iddet " _
        & " FROM com_honorariosdet INNER JOIN (alm_inventario INNER JOIN mae_detraccion ON alm_inventario.iddet = mae_detraccion.id) ON com_honorariosdet.iditem = alm_inventario.id " _
        & " WHERE (((com_honorariosdet.idhon)=" & xId & "));", xCon


    If Rst.RecordCount <> 0 Then
        If Rst("iddet") <> 0 Then
            ' SI ESTA AFECTO CREAMOS LA DETRACCION EN LA TABLA con_detraccion
            MsgBox "Se ha detectado que la compra registrada esta afecta al regimen de la Detraccion " + Chr(13) _
                & "Decripcion : " + Rst("descripcion") + Chr(13) _
                & "tasa : " + Format(Rst("tasa"), "0.00") + "%", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                
            Dim RstDeta As New ADODB.Recordset
            Dim xId2 As Double
            
            If QueHace = 1 Then
                xId2 = HallaCodigoTabla("con_detraccion", xCon, "id")
                RST_Busq RstDeta, "SELECT * FROM con_detraccion", xCon
                RstDeta.AddNew
                RstDeta("id") = xId2
            Else
                'buscamos detraccion de compra con_detraccion.tipo=2
                RST_Busq RstDeta, "SELECT con_detraccion.* From con_detraccion " _
                    & " WHERE (((con_detraccion.iddoc)=" & xId & ")) AND con_detraccion.tipo = 1", xCon
            End If
            
            If RstDeta.RecordCount = 0 Then
                ' este procedimiento es solo para cuando se este modificando una compra afecta a la detraccion y no
                ' se le haya hecho la detraccion a la hora de ingresar la compra
                xId2 = HallaCodigoTabla("con_detraccion", xCon, "id")
                RstDeta.AddNew
                RstDeta("id") = xId2
            End If
            
            RstDeta("iddet") = Rst("iddet")
            RstDeta("por") = NulosN(Rst("tasa"))
            RstDeta("iddoc") = xId
            RstDeta("idmon") = NulosN(TxtIdMon.Text)
            RstDeta("tipo") = 1
            RstDeta("fchmov") = Date
            RstDeta("Glosa") = ""
            RstDeta("imp") = Format((NulosN(TxtTotal.Text) * NulosN(Rst("tasa") / 100)), "0.00")
            RstDeta("numdet") = "SIN NUMERO"
            RstDeta.Update
        End If
    End If
    
    '----------------------------------------------------------------------------------
    '---generar asiento
    xNumAsiento = GenerarAsiento(xCon, 40, xId, AnoTra, mMesActivo, 1)
    If xNumAsiento = "" Then GoTo LaCague
    '----------------------------------------------------------------------------------
    
    
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
''    '--grabar datos adicionales en el diario
''    nSQL = "UPDATE ((com_honorarios INNER JOIN con_diario ON com_honorarios.id = con_diario.idmov) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha  " _
''        + vbCr + " SET con_diario.tc = [con_tc].[impven] , con_diario.fchdoc=com_honorarios.fchdoc, con_diario.idmon=com_honorarios.idmon, con_diario.ridlib = 40, con_diario.ridtipper = 1, con_diario.ridper = [com_honorarios].[idpro], con_diario.rtipdoc = [com_honorarios].[tipdoc], con_diario.rfchope = [com_honorarios].[fchdoc], con_diario.rnumerodoc = IIf([com_honorarios].[numser] Is Null Or [com_honorarios].[numser]='','',[com_honorarios].[numser] & '-') & [com_honorarios].[numdoc], con_diario.rglosaope = [com_honorarios].[glosa] & '', con_diario.rregistro = Left([com_honorarios].[numreg],2) & [mae_libros].[codsun] & Right([com_honorarios].[numreg],4) " _
''        + vbCr + " WHERE (((con_diario.idlib)=40) AND ((con_diario.idmov)=" & xId & ")); "
''
''    xCon.Execute nSQL
    
    xCon.CommitTrans
            
    MsgBox "El Honorario se registró con éxito" & vbCr & "Registro Nº: " & xNumAsiento, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Set RstDeta = Nothing
    Set RstCab = Nothing
    Set RstDet = Nothing
'    Set RstDia = Nothing
    Set RstCosto = Nothing
    Grabar = True
    Exit Function

LaCague:
    xCon.RollbackTrans
    Set RstDeta = Nothing
    Set RstCab = Nothing
    Set RstDet = Nothing
'    Set RstDia = Nothing
    Set RstCosto = Nothing
    
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
  
End Function

'*****************************************************************************************************
'* Nombre           : HallaNumAsiento
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GENERA EL NUMERO DE ASIENTO DEL PERIODO ESPECIFICADO, DEVUELVE UNA CADENA QUE
'*                    ES EL NUMERO DE REGISTRO
'* Paranetros       : NOMBRE    |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Mes       |  INTEGER     |  ESPECIFICA EL ID DEL MES
'* Devuelve         : String
'*****************************************************************************************************
Function HallaNumAsiento(Mes As Integer) As String
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT con_diario.idmes, con_diario.idlib, con_diario.numasi From con_diario " _
        & " WHERE (((con_diario.idmes)=" & Mes & ") AND ((con_diario.idlib)=40)) ORDER BY numasi", xCon
    
    If Rst.RecordCount = 0 Then
        HallaNumAsiento = "0001"
    Else
        Rst.MoveLast
        HallaNumAsiento = Format(NulosN(Rst("numasi")) + 1, "0000")
    End If
    Exit Function
End Function

Private Sub TxtTipDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtTipDoc.Text) = "" Then Exit Sub
    Dim xRs As New ADODB.Recordset
    
    RST_Busq xRs, "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
        & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuen as cuentaimp " _
        & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id WHERE mae_documento.id  = " & NulosN(TxtTipDoc.Text) & "", xCon
    
    If NulosN(TxtTipDoc.Text) = 2 Then
        Frame8.Visible = True
    Else
        Frame8.Visible = False
    End If
    
    If TxtTipDoc.Text = 7 Then
        Label3(9).Visible = True
        TxtDocRef.Visible = True
        CmdBusDocRef.Visible = True
    Else
        Label3(9).Visible = False
        TxtDocRef.Visible = False
        CmdBusDocRef.Visible = False
    End If
    
    If xRs.RecordCount = 0 Then
        TxtTipDoc.Text = ""
        LblNomDoc.Caption = ""
    Else
        CodSunatDoc = NulosC(xRs("codsun"))
        TxtTipDoc.Text = xRs("id")
        LblNomDoc.Caption = xRs("descripcion")
        TasaImpuesto = NulosN(xRs("tasa"))
        xDescImp = xRs("descripcion")
        xIdCuenTasa = NulosN(xRs("cuentaimp"))
        
        LblIgvTasa.Caption = Trim(Str(TasaImpuesto)) + "%"
        xPorIgv = (TasaImpuesto / 100)
        Frame3.Caption = "( Afecta : " + NulosC(xRs("descimp")) + ")"
    End If
    
    Set xRs = Nothing
    xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
    If xCuentaDoc = 0 Then
        TxtTipDoc.Text = ""
        LblNomDoc.Caption = ""
    End If
End Sub


'*****************************************************************************************************
'* Nombre           : Imprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MANDA A IMPRIMIR UNA COPIA DEL COMPROBANTE DE DOCUMENTO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Imprimir()
    Dim RsPDoc As New ADODB.Recordset
    Dim RsPCab As New ADODB.Recordset
    Dim RsPDet As New ADODB.Recordset
    Dim xRsDoc As New ADODB.Recordset
    Dim xRsDet As New ADODB.Recordset
    Dim RstGui As New ADODB.Recordset
    Dim A As Integer
    Dim xCadGuias As String

    RST_Busq xRsDoc, "SELECT com_honorarios.fchdoc, mae_prov.nombre, mae_prov.numdoc, com_honorarios.imptot, com_honorarios.tipdoc, com_honorarios.idmon, " _
        & " mae_prov.dir FROM mae_prov RIGHT JOIN com_honorarios " _
        & " ON mae_prov.id = com_honorarios.idpro Where (((com_honorarios.id) = " & RstComp("id") & "))", xCon
    
    RST_Busq xRsDet, "SELECT com_honorariosdet.idcom, alm_inventario.descripcion, mae_unidades.abrev, com_honorariosdet.canpro, com_honorariosdet.preuni, " _
        & " com_honorariosdet.imptot FROM mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN com_honorariosdet ON alm_inventario.id = com_honorariosdet.iditem) " _
        & " ON mae_unidades.id = com_honorariosdet.idunimed WHERE (((com_honorariosdet.idcom)=" & RstComp("id") & "))", xCon

    RST_Busq RsPDoc, "SELECT * FROM var_plantilladoc WHERE tipdoc = " & xRsDoc("tipdoc") & " ", xCon

    If RsPDoc.RecordCount = 0 Then
        MsgBox "No se ha definido la plantilla de impresion para este tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set xRsDoc = Nothing
        Set xRsDet = Nothing
        Set RsPDoc = Nothing
        Exit Sub
    End If
    RST_Busq RsPCab, "SELECT * FROM var_plantilladoc WHERE tipdoc = " & RsPDoc("tipdoc") & " ", xCon
    If RsPCab.RecordCount <> 0 Then
        A = RsPCab("id")
        RST_Busq RsPCab, "SELECT * FROM var_plantillacab WHERE idplan = " & A & " ORDER BY item", xCon
        RST_Busq RsPDet, "SELECT * FROM var_plantilladet WHERE idplan = " & A & " ORDER BY item", xCon
    End If

    Printer.Font = "Super Draft 15cpi"
    Printer.FontBold = True
    Printer.FontSize = 11
    Printer.ScaleMode = 6

    Dim xCam, xFor As String

    'imprime cabezera
    Do While RsPCab.EOF = False
        xCam = RsPCab("campo")
        xFor = NulosC(RsPCab("formato"))

        Printer.CurrentX = RsPCab("posx")
        Printer.CurrentY = RsPCab("posy")

        If NulosC(UCase(xCam)) <> UCase("x-numeletra") And NulosC(UCase(xCam)) <> UCase("x-numguia") And NulosC(UCase(xCam)) <> UCase("x-docref") Then
            Printer.Print Format((NulosC(xRsDoc(xCam))), xFor)
        Else
            If NulosC(UCase(xCam)) = UCase("x-numeletra") Then
                Printer.Print "Son : "; NumeroLetra(xRsDoc("imptot"), xRsDoc("idmon"))
            End If
            If NulosC(UCase(xCam)) = UCase("x-numguia") Then
                Printer.Print xCadGuias
            End If
            If NulosC(UCase(xCam)) = UCase("x-docref") Then
                Printer.Print "Referente a Factura(s) : "; xRsDoc("docref")
            End If
        End If

        RsPCab.MoveNext
    Loop

    'imprime detalle
    Dim Fila As Integer

    Fila = RsPDet("posy")
    xRsDet.MoveFirst
    Do While xRsDet.EOF = False
        RsPDet.MoveFirst
        Do While RsPDet.EOF = False
            xCam = RsPDet("campo")
            xFor = NulosC(RsPDet("formato"))
            Printer.CurrentX = RsPDet("posx")
            Printer.CurrentY = Fila
            If xFor = "" Then
                Printer.Print NulosC(xRsDet(xCam))
            Else
                Printer.Print Format((NulosC(xRsDet(xCam))), xFor)
            End If
            RsPDet.MoveNext
        Loop
        Fila = Fila + 4

        xRsDet.MoveNext
    Loop

    Printer.EndDoc
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITE BUSCAR UN REGISTRO EN EL RECORDSET RstComp
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    TabOne1.CurrTab = 0
    Dim xRs As New ADODB.Recordset
    
    Dim nSQL As String
    Dim xCampos(8, 4) As String
    
    xCampos(0, 0) = "N°Reg":          xCampos(0, 1) = "numreg":     xCampos(0, 2) = "820":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "T.D.":           xCampos(1, 1) = "abrev":      xCampos(1, 2) = "400":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "N°. Documento":  xCampos(2, 1) = "numerodoc":  xCampos(2, 2) = "1400":  xCampos(2, 3) = "C"
    xCampos(3, 0) = "FchEmi":         xCampos(3, 1) = "fchdoc":     xCampos(3, 2) = "830":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "FchVenc":        xCampos(4, 1) = "fchven":     xCampos(4, 2) = "830":   xCampos(4, 3) = "C"
    xCampos(5, 0) = "Proveedor":      xCampos(5, 1) = "nombre":     xCampos(5, 2) = "2600":  xCampos(5, 3) = "C"
    xCampos(6, 0) = "M":              xCampos(6, 1) = "simbolo":    xCampos(6, 2) = "450":   xCampos(6, 3) = "C"
    xCampos(7, 0) = "Importe":        xCampos(7, 1) = "imptot":     xCampos(7, 2) = "850":   xCampos(7, 3) = "N"
    
    nSQL = "SELECT com_honorarios.id,Mid([com_honorarios].[numreg],1,2)+[mae_libros].[codsun]+Mid([com_honorarios].[numreg],3,4) AS numreg, mae_prov.nombre, [com_honorarios]![numser]+'-'+[com_honorarios]![numdoc] AS numerodoc, mae_documento.abrev, format(com_honorarios.fchdoc,'dd/mm/yy') as fchdoc, format(com_honorarios.fchven,'dd/mm/yy') as fchven, mae_prov.numruc, mae_moneda.simbolo, com_honorarios.imptot, com_honorarios.impsal " _
        + vbCr + " FROM (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN com_honorarios ON mae_documento.id = com_honorarios.tipdoc) ON mae_moneda.id = com_honorarios.idmon) ON mae_prov.id = com_honorarios.idpro) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id " _
        + vbCr + " WHERE (((com_honorarios.numreg) Like '" & Format(mMesActivo, "00") & "%')) " _
        + vbCr + " ORDER BY com_honorarios.numreg DESC;"
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Compras", "nombre", "nombre", Principio

    If xRs.State = 1 Then
        RstComp.MoveFirst
        RstComp.Find "id = " & xRs("id") & ""
    End If
    
    Set xRs = Nothing
End Sub

Private Sub TxtTipDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDocRef_Click
    End If
End Sub

Private Sub TxtTipDocRef_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosN(TxtTipDocRef.Text) = 0 Then Exit Sub
    
    Dim xRs1 As New ADODB.Recordset
    
    RST_Busq xRs1, "SELECT * FROM mae_docreferencia WHERE id = " & NulosN(TxtTipDocRef.Text) & "", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtTipDocRef.Text = ""
        LblTipDocref.Caption = ""
    Else
        LblTipDocref.Caption = Trim(xRs1("descripcion"))
        TxtDocRef2.Text = ""
        LblIdDocRef2.Caption = ""
    End If
    Set xRs1 = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : pExportar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A EXCEL LA PESTAÑA CONSULTA DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportar()
    TabOne1.CurrTab = 0
    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset
    Dim xCampos(13, 3) As String
    
    '0::Nombre a Mostrar;
    '1::nombre de Campo del Rst;
    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
    '3::ancho de columna
    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Id":                       xCampos(0, 1) = "id":           xCampos(0, 2) = 2:    xCampos(0, 3) = "500"
    xCampos(1, 0) = "Nº Reg":                   xCampos(1, 1) = "numreg1":      xCampos(1, 2) = 0:    xCampos(1, 3) = "900"
    xCampos(2, 0) = "R.U.C.":                   xCampos(2, 1) = "numruc":       xCampos(2, 2) = 0:    xCampos(2, 3) = "1200"
    xCampos(3, 0) = "Prestador de Servicio":    xCampos(3, 1) = "nombre":       xCampos(3, 2) = 0:    xCampos(3, 3) = "3290"
    xCampos(4, 0) = "T.D.":                     xCampos(4, 1) = "abrev":        xCampos(4, 2) = 0:    xCampos(4, 3) = "350"
    xCampos(5, 0) = "Num. Doc":                 xCampos(5, 1) = "numerodoc":    xCampos(5, 2) = 0:    xCampos(5, 3) = "1600"
    xCampos(6, 0) = "Fch.Emi":                  xCampos(6, 1) = "fchdoc1":      xCampos(6, 2) = 1:    xCampos(6, 3) = "900"
    xCampos(7, 0) = "Fch. Venc":                xCampos(7, 1) = "fchven1":      xCampos(7, 2) = 1:    xCampos(7, 3) = "900"
    xCampos(8, 0) = "Glosa":                    xCampos(8, 1) = "glosa":        xCampos(8, 2) = 1:    xCampos(8, 3) = "2000"
    xCampos(9, 0) = "M":                        xCampos(9, 1) = "simbolo":      xCampos(9, 2) = 1:    xCampos(9, 3) = "500"
    xCampos(10, 0) = "T.C.":                    xCampos(10, 1) = "impven1":     xCampos(10, 2) = 2:   xCampos(10, 3) = "700"
    xCampos(11, 0) = "Importe":                 xCampos(11, 1) = "impbru":      xCampos(11, 2) = 2:   xCampos(11, 3) = "900"
    xCampos(12, 0) = "Impuesto":                xCampos(12, 1) = "impigv":      xCampos(12, 2) = 2:   xCampos(12, 3) = "900"
    xCampos(13, 0) = "Total":                   xCampos(13, 1) = "imptot":      xCampos(13, 2) = 2:   xCampos(13, 3) = "900"
        
    Set RstTmp = RstComp.Clone
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "LISTADO DE HONORARIOS", "Periodo " & LblMes.Caption, "", "Listado de Honorarios - " & LblMes.Caption, RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
End Sub

Private Sub VerAsiento()
    '===================================================================================================
    'Creado : 20/11/09 Por: Johan Castro
    'Propósito: Mostrar el asiento
    '
    'Entradas:  Tomara como base la informacion degistrada para que a partir de allí genere el asiento
    '
    'Resultados:Asiento en pantalla
    '===================================================================================================
    
    
    Dim RstAsi As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim ValTipCam As Double
    Dim mFila As Long
    Dim mIdCta As Long '--codigo de la cuenta
    
    '--validar datos
    If IsDate(TxtFchDoc.Valor) = False Then
        MsgBox "Falta especificar la Fecha de Emision", vbInformation, xTitulo
        TxtFchDoc.SetFocus
        Exit Sub
    End If
    
    '---------------------------------------------------------------------------------------------------------------------------
    '--definir la estructura del rst
    RST_Busq RstTmp, "SELECT TOP 1 con_diario.idcue, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, con_diario.tc AS tipcam, con_diario.impdebsol AS impdebmn, con_diario.imphabsol AS imphabmn, con_diario.imphabdol AS impdebme, con_diario.imphabdol AS imphabme " _
                   & " FROM con_diario INNER JOIN con_planctas ON con_diario.idcue = con_planctas.id; ", xCon
    
    DEFINIR_RST_TMP RstAsi, RstTmp
   
    '---------------------------------------------------------------------------------------------------------------------------
    'ValTipCam = NulosN(LblTipoCambio.Caption)
    '--almacenar el tipo de cambio para hacer las conversiones mas adelante
    ValTipCam = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
    If ValTipCam = 0 Then
        MsgBox "No hay tipo de Cambio" & vbCr & "Indique el tipo de cambio para continuar", vbInformation, xTitulo
        Exit Sub
    End If
        
    '-------------------------------------------------------------------------
    'grabamos a facturas por pagar Plan de cuentas 42.1 o dependiendo del caso
    RstAsi.AddNew
    RstAsi("idcue") = xCuentaDoc
    RstAsi("ctanum") = NulosC(Busca_Codigo(xCuentaDoc, "id", "cuenta", "con_planctas", "N", xCon))
    RstAsi("ctadesc") = NulosC(Busca_Codigo(xCuentaDoc, "id", "descripcion", "con_planctas", "N", xCon))
    RstAsi("tipcam") = ValTipCam

    If NulosN(TxtTipDoc.Text) <> 0 Then
        If NulosN(TxtTipDoc.Text) <> 7 Then
            'cuando se factura u otro comprabante excepto nota de credito hace su asiento norma
            If TxtIdMon.Text = "1" Then
                RstAsi("imphabmn") = NulosN(TxtTotal.Text)
                RstAsi("imphabme") = NulosN(TxtTotal.Text) / ValTipCam
            Else
                RstAsi("imphabmn") = NulosN(TxtTotal.Text) * ValTipCam
                RstAsi("imphabme") = NulosN(TxtTotal.Text)
            End If
        Else
            'cuando sea nota de credito hace el asiento inverso al de una venta
            If TxtIdMon.Text = "1" Then
                RstAsi("impdebmn") = NulosN(TxtTotal.Text)
                RstAsi("impdebme") = NulosN(TxtTotal.Text) / ValTipCam
            Else
                RstAsi("impdebmn") = NulosN(TxtTotal.Text) * ValTipCam
                RstAsi("impdebme") = NulosN(TxtTotal.Text)
            End If
        End If
    End If
    RstAsi.Update
        
    '-----------------------------------------------------
    'grabamos el impuesto si la operacion esta afecta a el
    If NulosN(TxtIGV.Text) <> 0 Then
        RstAsi.AddNew
        RstAsi("idcue") = xIdCuenTasa
        RstAsi("ctanum") = NulosC(Busca_Codigo(xIdCuenTasa, "id", "cuenta", "con_planctas", "N", xCon))
        RstAsi("ctadesc") = NulosC(Busca_Codigo(xIdCuenTasa, "id", "descripcion", "con_planctas", "N", xCon))
        RstAsi("tipcam") = ValTipCam
            
        If NulosN(TxtTipDoc.Text) <> 0 Then
            If NulosN(TxtTipDoc.Text) <> 7 And NulosN(TxtTipDoc.Text) <> 2 Then
                If TxtIdMon.Text = "1" Then
                    RstAsi("impdebmn") = NulosN(TxtIGV.Text)
                    RstAsi("impdebme") = NulosN(TxtIGV.Text) / ValTipCam
                Else
                    RstAsi("impdebmn") = NulosN(TxtIGV.Text) * ValTipCam
                    RstAsi("impdebme") = NulosN(TxtIGV.Text)
                End If
            Else
                If TxtIdMon.Text = "1" Then
                    RstAsi("imphabmn") = NulosN(TxtIGV.Text)
                    RstAsi("imphabme") = NulosN(TxtIGV.Text) / ValTipCam
                Else
                    RstAsi("imphabmn") = NulosN(TxtIGV.Text) * ValTipCam
                    RstAsi("imphabme") = NulosN(TxtIGV.Text)
                End If
            End If
        End If

        RstAsi.Update
    End If
   
                
    '***********************************************************************
    '--grabamos el imponible en function a los items de la factura
    For mFila = 1 To Fg1.Rows - 1
        mIdCta = NulosN(Fg1.TextMatrix(mFila, 11))
        If mIdCta <> 0 Then
            RstAsi.AddNew
            RstAsi("idcue") = mIdCta
            RstAsi("ctanum") = NulosC(Busca_Codigo(mIdCta, "id", "cuenta", "con_planctas", "N", xCon))
            RstAsi("ctadesc") = NulosC(Busca_Codigo(mIdCta, "id", "descripcion", "con_planctas", "N", xCon))
            RstAsi("tipcam") = ValTipCam
                
            If NulosN(TxtTipDoc.Text) <> 0 Then
                If NulosN(TxtTipDoc.Text) <> 7 Then
                    If TxtIdMon.Text = "1" Then
                        RstAsi("impdebmn") = NulosN(Fg1.TextMatrix(mFila, 8))
                        RstAsi("impdebme") = NulosN(Fg1.TextMatrix(mFila, 8)) / ValTipCam
                    Else
                        RstAsi("impdebmn") = NulosN(Fg1.TextMatrix(mFila, 8)) * ValTipCam
                        RstAsi("impdebme") = NulosN(Fg1.TextMatrix(mFila, 8))
                    End If
                Else
                    If TxtIdMon.Text = "1" Then
                        RstAsi("imphabmn") = NulosN(Fg1.TextMatrix(mFila, 8))
                        RstAsi("imphabme") = NulosN(Fg1.TextMatrix(mFila, 8)) / ValTipCam
                    Else
                        RstAsi("imphabmn") = NulosN(Fg1.TextMatrix(mFila, 8)) * ValTipCam
                        RstAsi("imphabme") = NulosN(Fg1.TextMatrix(mFila, 8))
                    End If
                End If
            End If
            RstAsi.Update
        End If
        
    Next mFila
        
    '***********************************************************************
    
    '--mostramos los asientos automaticos
    For mFila = 1 To Fg1.Rows - 1
        '--cta debe
        mIdCta = NulosN(Busca_Codigo(NulosN(Fg1.TextMatrix(mFila, 11)), "id", "ctadesdeb", "con_planctas", "N", xCon))
        If mIdCta <> 0 Then
            RstAsi.AddNew
            RstAsi("idcue") = mIdCta
            RstAsi("ctanum") = NulosC(Busca_Codigo(mIdCta, "id", "cuenta", "con_planctas", "N", xCon))
            RstAsi("ctadesc") = NulosC(Busca_Codigo(mIdCta, "id", "descripcion", "con_planctas", "N", xCon))
            RstAsi("tipcam") = ValTipCam
            
            If TxtIdMon.Text = "1" Then
                RstAsi("impdebmn") = NulosN(Fg1.TextMatrix(mFila, 8))
                RstAsi("impdebme") = NulosN(Fg1.TextMatrix(mFila, 8)) / ValTipCam
            Else
                RstAsi("impdebmn") = NulosN(Fg1.TextMatrix(mFila, 8)) * ValTipCam
                RstAsi("impdebme") = NulosN(Fg1.TextMatrix(mFila, 8))
            End If
            
            RstAsi("imphabmn") = 0
            RstAsi("imphabme") = 0
                
            RstAsi.Update
        End If
        
        '-----------------
        '--cta haber
        mIdCta = NulosN(Busca_Codigo(NulosN(Fg1.TextMatrix(mFila, 11)), "id", "ctadeshab", "con_planctas", "N", xCon))
        If mIdCta <> 0 Then
            RstAsi.AddNew
            RstAsi("idcue") = mIdCta
            RstAsi("ctanum") = NulosC(Busca_Codigo(mIdCta, "id", "cuenta", "con_planctas", "N", xCon))
            RstAsi("ctadesc") = NulosC(Busca_Codigo(mIdCta, "id", "descripcion", "con_planctas", "N", xCon))
            RstAsi("tipcam") = ValTipCam
            
            RstAsi("impdebmn") = 0
            RstAsi("impdebme") = 0
            
            If TxtIdMon.Text = "1" Then
                RstAsi("imphabmn") = NulosN(Fg1.TextMatrix(mFila, 8))
                RstAsi("imphabme") = NulosN(Fg1.TextMatrix(mFila, 8)) / ValTipCam
            Else
                RstAsi("imphabmn") = NulosN(Fg1.TextMatrix(mFila, 8)) * ValTipCam
                RstAsi("imphabme") = NulosN(Fg1.TextMatrix(mFila, 8))
            End If
            RstAsi.Update
        End If
        
    Next mFila
    
    '***********************************************************************
    '--mostrar el asiento
    Dim xfrm As New SGI2_funciones.formularios
    Dim xId As Double
    
    '--verificar que accion se esta haciendo
    If QueHace = 1 Then
        xId = 0
    Else
        xId = RstComp("id")
    End If
    
    xfrm.AsientoVerTmp xCon, RstAsi, 40, xId
    
    Set xfrm = Nothing

End Sub



Private Sub Frame6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    Frame6.ZOrder 0
End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With Frame6
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub

