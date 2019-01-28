VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPercepcionesVenta 
   Caption         =   "Igv Percepciones - Ventas"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12690
   LinkTopic       =   "Form2"
   ScaleHeight     =   8880
   ScaleWidth      =   12690
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
      Height          =   1245
      Left            =   4080
      TabIndex        =   48
      Top             =   840
      Visible         =   0   'False
      Width           =   4425
      Begin VB.ComboBox cbodocumentos 
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   240
         Width           =   4245
      End
      Begin VB.CommandButton cmdokseldoc 
         Height          =   555
         Left            =   1455
         Picture         =   "FrmPercepcionesVenta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   615
         Width           =   750
      End
      Begin VB.CommandButton cmdsalirseldoc 
         Height          =   555
         Left            =   2400
         Picture         =   "FrmPercepcionesVenta.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   615
         Width           =   750
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   3630
      Left            =   6885
      TabIndex        =   6
      Top             =   1890
      Visible         =   0   'False
      Width           =   2970
      Begin VB.CommandButton CmdOk 
         Height          =   660
         Left            =   840
         Picture         =   "FrmPercepcionesVenta.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2865
         Width           =   720
      End
      Begin VB.CommandButton CmdSalir 
         Height          =   660
         Left            =   1605
         Picture         =   "FrmPercepcionesVenta.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2865
         Width           =   720
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   16449538
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
         X1              =   -90
         X2              =   -90
         Y1              =   60
         Y2              =   3660
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   0
         X2              =   2955
         Y1              =   4035
         Y2              =   4035
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
         TabIndex        =   10
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
      Height          =   4860
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   6030
      Begin VB.CommandButton CmdSalirRef 
         Height          =   630
         Left            =   3525
         Picture         =   "FrmPercepcionesVenta.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4095
         Width           =   750
      End
      Begin VB.CommandButton CmdOkRef 
         Height          =   630
         Left            =   2520
         Picture         =   "FrmPercepcionesVenta.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4095
         Width           =   750
      End
      Begin VSFlex7Ctl.VSFlexGrid Fgdocref 
         Height          =   3555
         Left            =   105
         TabIndex        =   5
         Top             =   465
         Width           =   5760
         _cx             =   10160
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
         FormatString    =   $"FrmPercepcionesVenta.frx":123C
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
            Picture         =   "FrmPercepcionesVenta.frx":130A
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPercepcionesVenta.frx":184E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPercepcionesVenta.frx":1BE0
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPercepcionesVenta.frx":1D64
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPercepcionesVenta.frx":21B8
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPercepcionesVenta.frx":22D0
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPercepcionesVenta.frx":2814
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPercepcionesVenta.frx":2D58
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPercepcionesVenta.frx":2E6C
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPercepcionesVenta.frx":2F80
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPercepcionesVenta.frx":33D4
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPercepcionesVenta.frx":3540
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7965
      Left            =   -15
      TabIndex        =   0
      Top             =   360
      Width           =   11910
      _cx             =   21008
      _cy             =   14049
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
         Height          =   7545
         Left            =   45
         TabIndex        =   15
         Top             =   375
         Width           =   11820
         Begin VB.CommandButton CmdAddItem 
            Caption         =   "Agregar Item"
            Height          =   360
            Left            =   645
            TabIndex        =   44
            Top             =   2310
            Width           =   1260
         End
         Begin VB.CommandButton CmdBusNumSer 
            Height          =   240
            Left            =   2430
            Picture         =   "FrmPercepcionesVenta.frx":3A88
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1455
            Width           =   240
         End
         Begin VB.CommandButton CmdDelItem 
            Caption         =   "Eliminar Item"
            Height          =   360
            Left            =   1965
            TabIndex        =   28
            Top             =   2310
            Width           =   1260
         End
         Begin VB.CommandButton CmdBusCli 
            Height          =   240
            Left            =   3300
            Picture         =   "FrmPercepcionesVenta.frx":3BBA
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1140
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmPercepcionesVenta.frx":3CEC
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   810
            Width           =   240
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2895
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   25
            Text            =   "TxtNumDoc"
            Top             =   1425
            Width           =   1440
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   2460
            Picture         =   "FrmPercepcionesVenta.frx":3E1E
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   465
            Width           =   240
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   23
            Text            =   "TxtNumSer"
            Top             =   1425
            Width           =   885
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   22
            Text            =   "TxtTipDoc"
            Top             =   780
            Width           =   915
         End
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   21
            Text            =   "TxtNumRuc"
            Top             =   1110
            Width           =   1770
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   20
            Text            =   "TxtIdMon"
            Top             =   420
            Width           =   915
         End
         Begin VB.Frame Frame4 
            Height          =   600
            Left            =   210
            TabIndex        =   16
            Top             =   6765
            Width           =   11295
            Begin VB.TextBox TxtImpPer 
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
               Left            =   7455
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   19
               TabStop         =   0   'False
               Text            =   "TxtImpPer"
               Top             =   180
               Width           =   1100
            End
            Begin VB.TextBox TxtImporte 
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
               Left            =   4500
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   18
               TabStop         =   0   'False
               Text            =   "TxtImporte"
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
               TabIndex        =   17
               TabStop         =   0   'False
               Text            =   "TxtTotal"
               Top             =   180
               Width           =   1100
            End
            Begin VB.Label lbltotal 
               AutoSize        =   -1  'True
               Caption         =   "Total Cobrado"
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
               Left            =   8670
               TabIndex        =   47
               Top             =   195
               Width           =   1215
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Percepción"
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
               Left            =   5940
               TabIndex        =   46
               Top             =   195
               Width           =   1395
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Importe"
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
               Left            =   3615
               TabIndex        =   45
               Top             =   195
               Width           =   645
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   3885
            Left            =   240
            TabIndex        =   30
            Top             =   2805
            Width           =   11280
            _cx             =   19897
            _cy             =   6853
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
            FormatString    =   $"FrmPercepcionesVenta.frx":3F50
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
            TabIndex        =   31
            Top             =   1770
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
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Percepciones"
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
            Left            =   105
            TabIndex        =   43
            Top             =   45
            Width           =   11625
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emision"
            Height          =   195
            Index           =   2
            Left            =   255
            TabIndex        =   42
            Top             =   1785
            Width           =   1260
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio"
            Height          =   195
            Left            =   4485
            TabIndex        =   41
            Top             =   1755
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
            Left            =   5655
            TabIndex        =   40
            Top             =   1725
            Visible         =   0   'False
            Width           =   2865
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00808080&
            Height          =   495
            Index           =   1
            Left            =   555
            Top             =   2250
            Width           =   2775
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FFFFFF&
            Height          =   495
            Index           =   0
            Left            =   555
            Top             =   2250
            Width           =   2790
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
            Left            =   2760
            TabIndex        =   39
            Top             =   450
            Width           =   2790
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Index           =   7
            Left            =   255
            TabIndex        =   38
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
            TabIndex        =   37
            Top             =   1110
            Width           =   4935
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Documento"
            Height          =   195
            Index           =   1
            Left            =   255
            TabIndex        =   36
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
            TabIndex        =   35
            Top             =   795
            Width           =   5760
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Documento"
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   34
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
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   255
            TabIndex        =   33
            Top             =   435
            Width           =   585
         End
         Begin VB.Label LblIdCliente 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCliente"
            Height          =   195
            Left            =   7665
            TabIndex        =   32
            Top             =   1425
            Visible         =   0   'False
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7545
         Left            =   -12465
         TabIndex        =   11
         Top             =   375
         Width           =   11820
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   7155
            Left            =   30
            TabIndex        =   12
            Top             =   375
            Width           =   11790
            _ExtentX        =   20796
            _ExtentY        =   12621
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "T.D."
            Columns(0).DataField=   "Abrev"
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
            Columns(5).Caption=   "Imp. Doc."
            Columns(5).DataField=   "imptotdoc"
            Columns(5).NumberFormat=   "0.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Imp. Percep."
            Columns(6).DataField=   "imptotper"
            Columns(6).NumberFormat=   "0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Total"
            Columns(7).DataField=   "imptotcob"
            Columns(7).NumberFormat=   "0.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Estado"
            Columns(8).DataField=   "Estado"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
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
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2540"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2461"
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
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=2170"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2090"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=2275"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2196"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=514"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=1852"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1773"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
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
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Named:id=33:Normal"
            _StyleDefs(73)  =   ":id=33,.parent=0"
            _StyleDefs(74)  =   "Named:id=34:Heading"
            _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(76)  =   ":id=34,.wraptext=-1"
            _StyleDefs(77)  =   "Named:id=35:Footing"
            _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(79)  =   "Named:id=36:Selected"
            _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=37:Caption"
            _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(83)  =   "Named:id=38:HighlightRow"
            _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=39:EvenRow"
            _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(87)  =   "Named:id=40:OddRow"
            _StyleDefs(88)  =   ":id=40,.parent=33"
            _StyleDefs(89)  =   "Named:id=41:RecordSelector"
            _StyleDefs(90)  =   ":id=41,.parent=34"
            _StyleDefs(91)  =   "Named:id=42:FilterBar"
            _StyleDefs(92)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Percepciones"
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
            TabIndex        =   14
            Top             =   45
            Width           =   11625
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
            TabIndex        =   13
            Top             =   30
            Width           =   765
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   12555
         X2              =   24375
         Y1              =   375
         Y2              =   7920
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12690
      _ExtentX        =   22384
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
                  Text            =   "Modificar Documento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Restaurar Documento"
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
                  Text            =   "Anular Documento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar Documento"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Emitir Documento Anulada"
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
      Visible         =   0   'False
      X1              =   9270
      X2              =   9270
      Y1              =   255
      Y2              =   3855
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
Attribute VB_Name = "FrmPercepcionesVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstvent As New ADODB.Recordset
Dim QueHace As Integer

Dim CaracteresNumericos As String
Dim seEjecuto As Boolean
Dim ValTipCam As Double
Dim xDescImp As String
Dim xCuentaDoc As Integer   'codigo de la cuenta contable del documento
Dim xMes As Integer         'numero de mes en el que se realiza la operacion
Dim Mostrando As Boolean
Dim swguiafact '0 No se facturaron, 1 Se facturaron
Dim visdatos As Byte  'Para identificar si se muestra datos de 0 Guias 1 Notas de Credito , Nota Debito

Sub CambiarMes()
    Toolbar1.Enabled = False
    TabOne1.Enabled = False
    Frame5.Left = 4455
    Frame5.Top = 2100
    Frame5.Visible = True
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    
    If rstvent.RecordCount = 0 Then
                MsgBox "La cuadricula de datos esta vacia", vbInformation, xTitulo
                Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de eliminar " + rstvent("nomdoc") + " seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
    Dim rs As New ADODB.Recordset
        
        xCon.Execute "DELETE * FROM con_diario WHERE idmov = " & rstvent("id") & " AND idlib = 4"
        
             'Buscamos todos los documentos relacionados con la percepcion y actualizamos el campo idnumper =0 para generar
        'nueva percepciones
        RST_Busq rs, "SELECT * FROM con_percepciondeta WHERE  con_percepciondeta.id =" & rstvent("id") & "", xCon
        
        Do While Not rs.EOF
            'Actualizamos el campo  idnumper con el valor del campo  id de la tabla con_percepcion
            xCon.Execute " UPDATE vta_ventas SET vta_ventas.idnumper = 0 WHERE vta_ventas.id = " & rs("iddoc") & ""
            rs.MoveNext
        Loop
   
        xCon.Execute "DELETE * FROM con_percepcion WHERE id = " & rstvent("id") & ""
        
        
        MsgBox "El documento se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        rstvent.Requery
        Dg1.Refresh
    End If
    
    
      'Actualizar Saldos
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
    
    'Se restaura una documento anulado
    
    Dim Rpta As Integer
    
    
    
    
    Rpta = MsgBox("Esta seguro de restaurar " + rstvent("nomdoc") + " Nº " + rstvent("numser") & "-" + rstvent("numdoc"), vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption)
    If Rpta = vbYes Then
        xCon.Execute "UPDATE con_percepcion SET con_percepcion.Anulado = 0, " _
            & " con_percepcion.idcli = 1  " _
            & " WHERE con_percepcion.id =" & rstvent("id") & ""
        
        xCon.Execute "DELETE * FROM con_percepciondeta WHERE con_percepciondeta.id =" & rstvent("id") & ""
        rstvent.Requery
        Dg1.Refresh
        MsgBox rstvent("nomdoc") + " se restauro con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Sub Anular()


    If rstvent.RecordCount = 0 Then
                MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
                Exit Sub
            End If
   
   
   'Validamos si la factura esta anulada
    If rstvent("Anulado") = -1 Then
        MsgBox "el documento ya fue anulado, seleccione otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
                
    Dim Rpta As Integer
    Dim A As Integer
    Rpta = MsgBox("¿Esta seguro de anular el documento Nº " + rstvent("numser") & "-" & rstvent("numdoc") + "?", vbYesNo + vbDefaultButton1 + vbQuestion, xTitulo)
    
    If Rpta = vbYes Then
        Dim rs As New ADODB.Recordset
        
        xCon.Execute "UPDATE con_percepcion SET con_percepcion.Anulado = -1,  " _
            & " con_percepcion.idcli = 1, con_percepcion.imptotdoc = 0 , con_percepcion.imptotper = 0 , con_percepcion.imptotcob = 0, con_percepcion.impsal = 0    " _
            & " WHERE con_percepcion.id = " & rstvent("id") & " "
        
        
        'Buscamos todos los documentos relacionados con la percepcion y actualizamos el campo idnumper =0 para generar
        'nueva percepciones
        RST_Busq rs, "SELECT * FROM con_percepciondeta WHERE  con_percepciondeta.id =" & rstvent("id") & "", xCon
        
        Do While Not rs.EOF
            'Actualizamos el campo  idnumper con el valor del campo  id de la tabla con_percepcion
            xCon.Execute " UPDATE vta_ventas SET vta_ventas.idnumper = 0 WHERE vta_ventas.id = " & rs("iddoc") & ""
            rs.MoveNext
        Loop
        
        xCon.Execute "DELETE * FROM con_percepciondeta WHERE con_percepciondeta.id = " & rstvent("id") & ""
        
        
        Set rs = Nothing
        rstvent.Requery
        Dg1.Refresh
        MsgBox "El documento se anulo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        'Actualizar Saldo de documentos
    End If
End Sub

Sub Cancelar()
Dim X As Integer
    Bloquea
    Fg1.ColComboList(1) = ""
    Label5.Caption = "Detalle de Venta"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
     
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
       
       
       
    
    
    
    
End Sub

Sub Nuevo()
    QueHace = 1
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Label5.Caption = "Agregando Percepción"
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    TxtIdMon.SetFocus
    Fg1.Rows = 1


End Sub

Sub Modificar()
    If rstvent.RecordCount = 0 Then
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
    Label5.Caption = "Modificando Venta"
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    

End Sub

Sub MuestraSegundoTab()

    TxtTipDoc.Text = rstvent("tipdoc")
    TxtNumRuc.Text = rstvent("numruc")
    TxtNumSer.Text = rstvent("numser")
    TxtNumDoc.Text = rstvent("numdoc")
    TxtFchDoc.Valor = rstvent("fchdoc")
    TxtIdMon.Text = rstvent("idmon")
    TxtImporte = Format(rstvent("imptotdoc"), "0.00")
    TxtTotal.Text = Format(rstvent("imptotcob"), "0.00")
    TxtImpPer.Text = Format(rstvent("imptotper"), "0.00")
    LblNomDoc.Caption = rstvent("nomdoc")
    LblNomCli.Caption = rstvent("nombre")
    
    LblMoneda.Caption = rstvent("descmon")
    LblIdCliente.Caption = rstvent("idcli")
    
    

    If rstvent("idmon") = 1 Then
        LblTipoCambio.Visible = False
    Else
        LblTipoCambio.Visible = True
        LblTipoCambio.Caption = rstvent("impven")
    End If
    
    
    Dim Rstdet As New ADODB.Recordset
    Dim xRs2   As New ADODB.Recordset
    
    Mostrando = True
    Fg1.Rows = 1
    
     
    RST_Busq Rstdet, " SELECT mae_documento.descripcion, vta_ventas.numser + '-' + vta_ventas.numdoc as [nrodoc] , vta_ventas.fchdoc, con_percepciondeta.porper, con_percepciondeta.impper, con_percepciondeta.impdoc, con_percepciondeta.impcob, con_percepciondeta.iddoc, con_percepciondeta.idper, vta_ventas.id " & _
                     " FROM mae_documento INNER JOIN (con_percepciondeta INNER JOIN vta_ventas ON con_percepciondeta.iddoc = vta_ventas.id) ON mae_documento.id = vta_ventas.tipdoc " & _
                     " WHERE con_percepciondeta.id = " & rstvent("id") & "", xCon

    
    If Rstdet.RecordCount <> 0 Then
        Do While Not Rstdet.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Row, 1) = Rstdet("descripcion")
        Fg1.TextMatrix(Fg1.Row, 2) = Rstdet("nrodoc")
        Fg1.TextMatrix(Fg1.Row, 3) = Rstdet("Fchdoc")
        Fg1.TextMatrix(Fg1.Row, 4) = Format(Rstdet("impdoc"), "0.00")
        Fg1.TextMatrix(Fg1.Row, 5) = Rstdet("porper") & "%"
        Fg1.TextMatrix(Fg1.Row, 10) = Rstdet("idper") 'id de item de percepcion
        Fg1.TextMatrix(Fg1.Row, 6) = Format(Rstdet("impper"), "0.00")
        Fg1.TextMatrix(Fg1.Row, 7) = Format(Rstdet("impcob"), "0.00")
        Fg1.TextMatrix(Fg1.Row, 9) = Rstdet("id") 'id del documento afecto a percepcion
        Rstdet.MoveNext
        Loop
        
        Set xRs2 = Nothing
        
        RST_Busq xRs2, "SELECT idcuenvta From mae_impuestos WHERE id = 4 ", xCon
        
        If xRs2.RecordCount > 0 Then
            Fg1.TextMatrix(Fg1.Row, 8) = xRs2("idcuenvta") 'id de cuenta para formar el haber cuenta 40x
        End If
                
    End If
    Set Rstdet = Nothing
    Mostrando = False
    
    'Cargamos el codigo de la cuenta contable del documento  para fomar el debe cuenta 12x
    
    Set Rstdet = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & Val(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & Val(TxtIdMon) & " and tipope = -1", xCon)

    If Rstdet.RecordCount = 1 Then
        xCuentaDoc = Rstdet("idcuen")
    End If
    
    Set Rstdet = Nothing
    Set xRs2 = Nothing
    
End Sub

Sub Bloquea()
    
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtFchDoc.Locked = Not TxtFchDoc.Locked
    
    
    TxtIdMon.Locked = Not TxtIdMon.Locked
    
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
    
    
    
End Sub

Sub Blanquea()

    TxtTipDoc.Text = ""
    TxtNumRuc.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtFchDoc.Valor = ""
    


    TxtIdMon.Text = ""
    
    LblNomDoc.Caption = ""
    LblNomCli.Caption = ""

    LblMoneda.Caption = ""
    LblIdCliente.Caption = ""

    

    
    TxtImpPer = ""


    TxtImporte = ""
    TxtTotal.Text = ""

    Fg1.Rows = 1
End Sub

Private Sub CmdAddItem_Click()

 If QueHace = 3 Then Exit Sub
    If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 1)) = "" Then Exit Sub
    Fg1.Rows = Fg1.Rows + 1
    
    With Fg1
        .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, 1
    End With
    
    Fg1_CellButtonClick Fg1.Rows - 1, 1
    Fg1.SetFocus
End Sub





Private Sub CmdBusNumSer_Click()
    Dim xform As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset

    
    If QueHace = 3 Then Exit Sub
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
              
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "iddoc":       xCampos(0, 2) = "800":    xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion": xCampos(1, 2) = "2200":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Serie":          xCampos(2, 1) = "numser":      xCampos(2, 2) = "1500":    xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nro Documento":  xCampos(3, 1) = "numdoc":      xCampos(3, 2) = "1500":    xCampos(3, 3) = "C"
    
    
    xform.SqlCad = "SELECT mae_documento.descripcion, mae_series.iddoc, mae_series.numser, mae_series.numdoc " & _
                   " FROM mae_documento INNER JOIN mae_series ON mae_documento.id = mae_series.iddoc WHERE mae_series.iddoc = " & Val(TxtTipDoc) & ""

    xform.Titulo = "Buscando Series"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numser"
    xform.CampoBusca = "numser"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtNumSer.Text = Format(xRs("numser"), "0000")
        TxtNumDoc = HallaNumdocVenta(Val(TxtTipDoc), NulosC(TxtNumSer.Text), xCon)
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdBusCli_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Cliente":    xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SqlCad = "SELECT mae_cliente.nombre, mae_cliente.numruc, mae_cliente.id From mae_cliente"
    
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

    Dim xform As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SqlCad = "SELECT * FROM mae_moneda ORDER BY descripcion"
    
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

    Dim xform As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SqlCad = "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
        & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuenvta  as cuentaimp" _
        & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuenvta = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id"
    
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
        
        xDescImp = xRs("descripcion")
        
        Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & Val(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & Val(TxtIdMon) & " and tipope = -1", xCon)
            If xRs2.RecordCount > 0 Then
                xCuentaDoc = NulosN(xRs2("idcuen"))
            End If
            TxtNumRuc.SetFocus
    End If
    Set xRs2 = Nothing
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub CmdDelItem_Click()
    If QueHace = 3 Then Exit Sub
    If Fg1.Rows - 1 > 0 Then
        
        If Fg1.Rows - 1 = 1 Then
             Fg1.Rows = 1
            Exit Sub
        Else
            Fg1.RemoveItem Fg1.Row
        End If
        
    End If
    
    
    
    HallarTotal
End Sub





Private Sub CmdOk_Click()
        Dim xFchIni As String
        Dim xFchFin As String
    Dim xFecha As String
    xFecha = Format(MonthView1.Value, "dd/mm/yy")
        
    xFchIni = "01/" + Mid(Format(CDate(xFecha), "dd/mm/yy"), 4, 5)
    xFchFin = Trim(Format(HallaDiasMes(CDate(xFecha)), "00")) + "/" + Mid(Format(CDate(xFecha), "dd/mm/yy"), 4, 5)
    

    RST_Busq rstvent, " SELECT con_percepcion.tipdoc, con_percepcion.numser, con_percepcion.numdoc, con_percepcion.fchdoc, con_percepcion.idcli, con_percepcion.imptotdoc, con_percepcion.imptotper, con_percepcion.imptotcob, con_percepcion.idmon, mae_cliente.nombre, mae_cliente.numruc, con_tc.impven,  mae_documento.abrev, mae_documento.descripcion as [nomdoc], mae_moneda.descripcion as [descmon],mae_moneda.Simbolo, [con_percepcion]![numser] + '-' + [con_percepcion]![numdoc] AS Numerodoc ,con_percepcion.id, iif(con_percepcion.anulado = 0,'Generado','Anulado') as Estado  " & _
                      " FROM con_tc INNER JOIN (mae_moneda INNER JOIN (mae_documento INNER JOIN (con_percepcion INNER JOIN mae_cliente ON con_percepcion.id = mae_cliente.id) ON mae_documento.id = con_percepcion.tipdoc) ON mae_moneda.id = con_percepcion.idmon) ON con_tc.fecha = con_percepcion.fchdoc " _
                    & " WHERE (((con_percepcion.numreg) Like '" & Format(xMes, "00") & "%'))", xCon
    
        Set Dg1.DataSource = rstvent
    
    
    CmdSalir_Click
End Sub


Private Sub cmdokseldoc_Click()


    
    Dim Rpta As Integer
    Dim RstCab As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim xId As Integer
    Dim xnumdoc As String
    Dim xnumser As String
    Dim xNumAsiento As String
    
    
    RST_Busq xRs, "SELECT * FROM mae_series WHERE mae_series.iddoc = " & Val(Right(cbodocumentos, 5)), xCon
        
    If xRs.RecordCount > 0 Then
        xRs.MoveFirst
        xnumser = xRs("numser")
    Else
        MsgBox "Registre la serie en Mantenimiento de Series", vbInformation, Me.Caption
        Set xRs = Nothing
        Exit Sub
    End If
    
    
    Rpta = MsgBox("Esta seguro de emitir " & Trim(Mid(cbodocumentos, 1, 100)) & " como anulado", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
    If Rpta = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo LaCague
    xCon.BeginTrans
    
    'Validar si el nro de documento existe solo en modo adicionar documento
    
    RST_Busq RstCab, "SELECT * FROM con_percepcion ", xCon
    

    xId = HallaCodigoTabla("con_percepcion", xCon, "id")
    RstCab.AddNew
    RstCab("id") = xId
    RstCab("tipdoc") = Val(Right(cbodocumentos, 5))
    RstCab("idcli") = 1
    RstCab("tipo") = 1
    
    
    xnumdoc = HallaNumdocVenta(Val(Right(cbodocumentos, 5)), Val(xnumser), xCon)
    
    RstCab("numser") = Format(xnumser, "0000")
    RstCab("numdoc") = xnumdoc
    
    If NulosC(TxtFchDoc.Valor) <> "" Then RstCab("Fchdoc") = TxtFchDoc.Valor
    
    RstCab("idmon") = 1
    RstCab("imptotdoc") = 0
    RstCab("imptotper") = 0
    RstCab("imptotcob") = 0
    RstCab("numreg") = Trim(Str(xMes)) + xNumAsiento
    RstCab("anulado") = -1
    RstCab.Update
                            
    'Actualizamos el numero de documento en la tabla Mae_series
    Call ActualizaNroDocumento(Val(xnumdoc), Val(Right(cbodocumentos, 5)), Val(xnumser))
    
    MsgBox Trim(Mid(cbodocumentos, 1, 100)) & " anulado se genero con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    xCon.CommitTrans
    rstvent.Requery
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
    Dim xform As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    Dim impper As Double

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(7, 4) As String
    
    xCampos(0, 0) = "Id":               xCampos(0, 1) = "id":           xCampos(0, 2) = "300":     xCampos(0, 3) = "N"
    xCampos(1, 0) = "Fecha Doc":        xCampos(1, 1) = "fchdoc":       xCampos(1, 2) = "1100":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Ruc":              xCampos(2, 1) = "numruc":       xCampos(2, 2) = "1200":    xCampos(2, 3) = "C"
    xCampos(3, 0) = "Razon Social":     xCampos(3, 1) = "nombre":       xCampos(3, 2) = "1600":    xCampos(3, 3) = "C"
    xCampos(4, 0) = "Documento":        xCampos(4, 1) = "descripcion":  xCampos(4, 2) = "1200":    xCampos(4, 3) = "C"
    xCampos(5, 0) = "Nro Documento":    xCampos(5, 1) = "nrodoc":       xCampos(5, 2) = "1500":    xCampos(5, 3) = "C"
    xCampos(6, 0) = "Total Doc":        xCampos(6, 1) = "imptotdoc":    xCampos(6, 2) = "1200":    xCampos(6, 3) = "N"
    
    
    

    
    xform.SqlCad = " SELECT vta_ventas.id, vta_ventas.fchdoc, mae_cliente.numruc, mae_cliente.nombre, mae_documento.descripcion, vta_ventas.numser+'-'+vta_ventas.numdoc AS NroDoc, vta_ventas.imptotdoc " _
                   & " FROM mae_documento INNER JOIN (mae_cliente INNER JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON mae_documento.id = vta_ventas.tipdoc " _
                   & " WHERE vta_ventas.idcli = " & Val(LblIdCliente) & " AND vta_ventas.percepcion = -1 AND vta_ventas.idnumper = 0 AND vta_ventas.anulado = 0 "
        
    
    
    xform.Titulo = "Buscando Documentos"
    xform.FormaBusca = CualquierParte
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        
        Fg1.TextMatrix(Fg1.Row, 1) = xRs("descripcion")
        Fg1.TextMatrix(Fg1.Row, 2) = xRs("nrodoc")
        Fg1.TextMatrix(Fg1.Row, 3) = xRs("Fchdoc")
        Fg1.TextMatrix(Fg1.Row, 4) = Format(xRs("imptotdoc"), "0.00")
                
        RST_Busq xRs2, " SELECT vta_ventasdet.tasaper, alm_inventario.idper " & _
        " FROM alm_inventario INNER JOIN vta_ventasdet ON alm_inventario.id = vta_ventasdet.iditem " & _
        " WHERE vta_ventasdet.idvta =" & xRs("id") & "", xCon
        
       'Extraemos el % de la percepción y los calculos respectivos
        
       If xRs2.RecordCount > 0 Then
            impper = Format(xRs("imptotdoc") * (xRs2("tasaper") / 100), "0.00")
            Fg1.TextMatrix(Fg1.Row, 5) = xRs2("tasaper") & "%"
            Fg1.TextMatrix(Fg1.Row, 10) = xRs2("idper") 'id de item de percepcion
        End If
        
        Fg1.TextMatrix(Fg1.Row, 6) = Format(impper, "0.00")
        Fg1.TextMatrix(Fg1.Row, 7) = Format((xRs("imptotdoc") + impper), "0.00")
                
        Set xRs2 = Nothing
        RST_Busq xRs2, "SELECT idcuenvta From mae_impuestos WHERE id = 4 ", xCon
        
        If xRs2.RecordCount > 0 Then
            Fg1.TextMatrix(Fg1.Row, 8) = xRs2("idcuenvta") 'id de cuenta para formar haber
        End If
            Fg1.TextMatrix(Fg1.Row, 9) = xRs("id") 'id del documento afecto a percepcion
            
    End If
        
        
    
    
    HallarTotal
    Set xform = Nothing
    Set xRs = Nothing
    Set xRs2 = Nothing
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    
    If Mostrando = True Then Exit Sub
    HallarTotal
    
    
End Sub

Sub HallarTotal()
    Dim A As Integer
    Dim totaldoc As Double
    Dim totalper  As Double
    Dim totalcob As Double
    
    TxtImpPer = "0.00"
    TxtImporte = "0.00"
    TxtTotal = "0.00"
    
    For A = 1 To Fg1.Rows - 1
            totaldoc = totaldoc + Val(Fg1.TextMatrix(A, 4)) 'total doc
            totalper = totalper + Val(Fg1.TextMatrix(A, 6)) 'total percibido
            totalcob = totalcob + Val(Fg1.TextMatrix(A, 7)) 'total cob
    Next A
    
        TxtImporte = Format(totaldoc, "0.00")
        TxtImpPer = Format(totalper, "0.00")
        TxtTotal = Format(totalcob, "0.00")
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then Exit Sub
    
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



Private Sub Form_Activate()
    If seEjecuto = False Then
        seEjecuto = True
        Dim Rpta As Integer
        Dim Rst As New ADODB.Recordset
        
        
        Set Rst = Nothing
                                   
        RST_Busq rstvent, " SELECT con_percepcion.tipdoc, con_percepcion.numser, con_percepcion.numdoc, con_percepcion.fchdoc, con_percepcion.idcli, con_percepcion.imptotdoc, con_percepcion.imptotper, con_percepcion.imptotcob, con_percepcion.idmon, mae_cliente.nombre, mae_cliente.numruc, con_tc.impven,  mae_documento.abrev, mae_documento.descripcion as [nomdoc], mae_moneda.descripcion as [descmon],mae_moneda.Simbolo, [con_percepcion]![numser] + '-' + [con_percepcion]![numdoc] AS Numerodoc ,con_percepcion.id, iif(con_percepcion.anulado = 0,'Generado','Anulado') as Estado ,con_percepcion.Anulado  " & _
                          " FROM con_tc INNER JOIN (mae_moneda INNER JOIN (mae_documento INNER JOIN (con_percepcion INNER JOIN mae_cliente ON con_percepcion.idcli = mae_cliente.id) ON mae_documento.id = con_percepcion.tipdoc) ON mae_moneda.id = con_percepcion.idmon) ON con_tc.fecha = con_percepcion.fchdoc " _
                          & " WHERE (((con_percepcion.numreg) Like '" & Format(xMes, "00") & "%'))", xCon

        
        Set Dg1.DataSource = rstvent
        If rstvent.RecordCount = 0 Then
            Rpta = MsgBox("No se ha registrado ninguna percepción , ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            If Rpta = vbYes Then
                Nuevo
            Else
                Unload Me
            End If
        Else
            Dg1.SetFocus
        End If
        
    End If
    TxtFchDoc.Valor = Date
End Sub

Private Sub Form_Load()
 
    QueHace = 3
    TabOne1.CurrTab = 0
    seEjecuto = False
    
    CaracteresNumericos = "0123456789." & Chr(8)
    Fg1.ColWidth(8) = 0
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    
    
      
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ColComboList(1) = ""
    swguiafact = 0
    xaño = 2006
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





Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        
        'Validamos si la cuadricula tiene datos
        
            If QueHace = 3 Then
                If rstvent.RecordCount = 0 Then
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
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Anular
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            rstvent.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    
    If Button.Index = 11 Then CambiarMes
    
    If Button.Index = 15 Then
        Set rstvent = Nothing
        Unload Me
    End If
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  If ButtonMenu.Parent.Index = 2 Then
        
       'MODIFICACION DE FACTURAS
        If ButtonMenu.Index = 1 Then
            If rstvent.RecordCount = 0 Then
                MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
                Exit Sub
            End If
            If rstvent("anulado") = -1 Then
                        MsgBox "No puede modificar un documento anulado proceda a restaurarlo", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
                        Exit Sub
            Else
                        Modificar
            End If
        End If
        
        'RESTAURAR documentoS
        If ButtonMenu.Index = 2 Then
            If rstvent.RecordCount = 0 Then
                MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
                Exit Sub
            End If
                If rstvent("anulado") = -1 Then
                    RestaurarFactura
                End If
        End If
    
    End If
  
  If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then Anular
        If ButtonMenu.Index = 2 Then Eliminar
        If ButtonMenu.Index = 3 Then EmitirAnulada
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

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        If NulosC(TxtNumDoc.Text) <> "" Then
            TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
        End If
        TxtFchDoc.SetFocus
    End If
End Sub


Private Sub TxtNumRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'SendKeys vbTab
        
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
        TxtNumSer.SetFocus
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
                         " FROM mae_documento INNER JOIN mae_series ON mae_documento.id = mae_series.iddoc where mae_series.iddoc = " & CLng(TxtTipDoc) & " and mae_series.numser =" & CLng(Me.TxtNumSer) & "  ", xCon
        
        
        If Rstdoc.RecordCount = 0 Then
            MsgBox "Registre la serie en Mantenimiento de Series", vbInformation, Me.Caption
            TxtNumSer = ""
            TxtNumDoc = ""
        Else
            TxtNumSer.Text = Format(Rstdoc("numser"), "0000")
            TxtNumDoc = HallaNumdocVenta(CLng(TxtTipDoc), CLng(Trim(TxtNumSer.Text)), xCon)
        End If
        TxtNumDoc.SetFocus
            'SendKeys vbTab
    End If
    
End Sub

Private Sub TxtNumSer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
CmdBusNumSer_Click

End If

End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'SendKeys vbTab
        
        If NulosC(TxtTipDoc.Text) = "" Then Exit Sub
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

            xDescImp = xRs("descripcion")

            
            
            Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & Val(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & Val(TxtIdMon) & " and tipope = -1", xCon)
            If xRs2.RecordCount > 0 Then
                xCuentaDoc = NulosN(xRs2("idcuen"))
            End If
            Set xRs2 = Nothing
            
            

            
            TxtNumRuc.SetFocus
        End If
        
                
        
                
        Set xRs = Nothing
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
    
    
    RST_Busq rs, "SELECT * FROM mae_documento WHERE Id  <> 7  AND Id <> 8 ORDER BY descripcion ", xCon
    
    cbodocumentos.Clear
    Do While Not rs.EOF
        cbodocumentos.AddItem rs!Descripcion & Space(100) & rs!id
        rs.MoveNext
    Loop
    
    If rs.RecordCount > 0 Then cbodocumentos.ListIndex = 0
    Fraseldoc.Visible = True
    Set rs = Nothing
End Sub

Function Grabar() As Boolean

        
    
    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Function
    End If
    
    If TxtNumRuc.Text = "" Then
        MsgBox "No ha especificado cliente ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
    
    If TxtIdMon.Text = "" Then
        MsgBox "No ha especificado la moneda del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado items para generar la percepción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    Dim RstCab As New ADODB.Recordset
    'Validar si el nro de documento existe
    If QueHace = 1 Then
    RST_Busq RstCab, " Select * from con_percepcion where Clng(tipdoc) =" & CLng(TxtTipDoc) & " and Clng(Numser) =" & CLng(TxtNumSer) & " and Clng(numdoc) = " & CLng(Me.TxtNumDoc) & " ", xCon
    
    If RstCab.RecordCount > 0 Then
        MsgBox "El Nro de documento ha sido registrado por otro usuario se grabara con otro numero", vbInformation, Me.Caption
        TxtNumDoc = HallaNumdocVenta(Val(TxtTipDoc), NulosC(TxtNumSer.Text), xCon)
    End If
    
    Set RstCab = Nothing
    End If
    
    Dim RstDeta2 As New ADODB.Recordset
    Dim Rstdet As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim xIdCuen As Integer
    Dim xTotal As Double
    Dim xNumAsiento As String
    
    Dim xId As Integer
    Dim A As Integer
    Dim X As Integer
 
    On Error GoTo LaCague
    
    swguiafact = 1
    xCon.BeginTrans
    
    
    If QueHace = 1 Then 'Nuevo Registro
        xId = HallaCodigoTabla("con_percepcion", xCon, "id")
        xNumAsiento = HallaNumAsiento(xMes)
        
        RST_Busq RstCab, "SELECT * FROM con_percepcion ", xCon
        RST_Busq Rstdet, "SELECT * FROM con_percepciondeta ", xCon
        RST_Busq RstDia, "SELECT * FROM con_diario ", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = rstvent("id")
        RST_Busq RstCab, "SELECT * FROM con_percepcion WHERE  id = " & xId & "", xCon
        
            
        'Eliminamos el detalle de la venta
         xCon.Execute "DELETE * FROM con_percepciondeta WHERE id = " & xId & ""
        
        RST_Busq Rstdet, "SELECT * FROM con_percepciondeta ", xCon
        
        RST_Busq RstDia, "SELECT * FROM con_diario WHERE idmes = " & Format(CDate(TxtFchDoc.Valor), "mm") & " AND " _
                         & " idlib = 4 AND idmov = " & xId & "", xCon
            
        If RstDia.RecordCount <> 0 Then
            xNumAsiento = RstDia("numasi")
        End If
        
        Set RstDia = Nothing
        
        'eliminamos el asiento contable
        xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & Format(CDate(TxtFchDoc.Valor), "mm") & " AND " _
            & " idlib = 4 AND idmov = " & xId & ""
            
        RST_Busq RstDia, "SELECT * FROM con_diario", xCon
    End If
    
    
    
    RstCab("idcli") = NulosN(LblIdCliente.Caption)
    RstCab("tipdoc") = Val(TxtTipDoc)
    RstCab("numser") = TxtNumSer.Text
    RstCab("numdoc") = TxtNumDoc.Text
    RstCab("tipo") = 1 'Ingreso
    RstCab("fchdoc") = TxtFchDoc.Valor
    RstCab("numreg") = Trim(Str(xMes)) + xNumAsiento
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    RstCab("imptotdoc") = NulosN(TxtImporte.Text)
    RstCab("imptotper") = NulosN(TxtImpPer.Text)
    RstCab("imptotcob") = NulosN(TxtTotal.Text)
    RstCab("impsal") = NulosN(TxtTotal.Text)
    RstCab.Update
                           
    
    For A = 1 To Fg1.Rows - 1
        Rstdet.AddNew
        Rstdet("id") = xId
        Rstdet("idper") = Val(Fg1.TextMatrix(A, 10))
        Rstdet("porper") = Val(Left(Fg1.TextMatrix(A, 5), 1))
        
        Rstdet("iddoc") = Val(Fg1.TextMatrix(A, 9))
         
         'Actualizamos el campo  idnumper con el valor del campo  id de la tabla con_percepcion
         xCon.Execute " UPDATE vta_ventas SET vta_ventas.idnumper = " & xId & " WHERE vta_ventas.id = " & Val(Fg1.TextMatrix(A, 9)) & ""
        
        
        Rstdet("impdoc") = Val(Fg1.TextMatrix(A, 4))
        Rstdet("impper") = Val(Fg1.TextMatrix(A, 6))
        Rstdet("impcob") = Val(Fg1.TextMatrix(A, 7))
        Rstdet.Update
    
    Next A
   

    'Grabamos el documento de percepcion en la tabla diario
    
    RstDia.AddNew
    RstDia("año") = xaño
    RstDia("idmes") = xMes
    RstDia("idlib") = 4
    RstDia("idmov") = xId
    RstDia("numasi") = xNumAsiento
    RstDia("tc") = ValTipCam
    RstDia("iddoc") = Val(TxtTipDoc)
    RstDia("idcue") = xCuentaDoc
    If TxtIdMon.Text = "1" Then
        RstDia("impdebsol") = Val(TxtTotal.Text)
        RstDia("impdebdol") = 0
    Else
        RstDia("impdebsol") = Val(TxtTotal.Text) * Val(LblTipoCambio.Caption)
        RstDia("impdebdol") = Val(TxtTotal.Text)
    End If
    
    RstDia.Update
    
    
    '********Formamos el Haber
    
    Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xCampos(2, 3) As String
    Dim rstdocus As New ADODB.Recordset

            xCampos(0, 0) = "cuenta":     xCampos(0, 1) = "C":      xCampos(0, 2) = "12"
            xCampos(1, 0) = "Importe":    xCampos(1, 1) = "D":      xCampos(1, 2) = "2"

            Set rstdocus = xFun.CrearRstTMP(xCampos)
            rstdocus.Open


          For X = 1 To Fg1.Rows - 1
            xIdCuen = Trim(Fg1.TextMatrix(X, 8))
            xTotal = Val(Fg1.TextMatrix(X, 7))
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
          Next X
           
           'Grabamos el diario
          If rstdocus.RecordCount > 0 Then
                rstdocus.MoveFirst
                Do While Not rstdocus.EOF
                    RstDia.AddNew
                    RstDia("año") = xaño
                    RstDia("idmes") = xMes               'LLAVE - CODIGO DEL MES
                    RstDia("idlib") = 4                  'LLAVE - CODIGO DEL LIBRO
                    RstDia("idmov") = xId                'LLAVE - CODIGO DEL MOVIMIENTO
                    RstDia("numasi") = xNumAsiento       'LLAVE - NUMERO DE ASIENTO
                    RstDia("iddoc") = Val(TxtTipDoc)
                    RstDia("tc") = ValTipCam
                    RstDia("idcue") = rstdocus("cuenta")
                    If TxtIdMon.Text = "1" Then
                        RstDia("imphabsol") = rstdocus("importe")
                        RstDia("imphabdol") = 0
                    Else
                        RstDia("imphabsol") = rstdocus("importe") * Val(LblTipoCambio.Caption)
                        RstDia("imphabdol") = rstdocus("importe")
                    End If
                    RstDia.Update
                    rstdocus.MoveNext
                Loop
            End If

    
   'Grabamos ó actualizamos el ultimo numero del documento
    If QueHace = 1 Then
       Call ActualizaNroDocumento(Val(TxtNumDoc), Val(TxtTipDoc), Val(TxtNumSer))
    End If
    
    
    xCon.CommitTrans
    MsgBox "El documento se registro con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Set RstCab = Nothing
    Set Rstdet = Nothing
    Set RstDia = Nothing
    Grabar = True
    Exit Function
    
LaCague:
    'Resume
    xCon.RollbackTrans
    
    Set RstCab = Nothing
    Set Rstdet = Nothing
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



