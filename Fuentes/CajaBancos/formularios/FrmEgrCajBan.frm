VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmEgrCajBan 
   Caption         =   "Caja y Bancos - Egresos"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11880
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
            Picture         =   "FrmEgrCajBan.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgrCajBan.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgrCajBan.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgrCajBan.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgrCajBan.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgrCajBan.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgrCajBan.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgrCajBan.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgrCajBan.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgrCajBan.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgrCajBan.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgrCajBan.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   0
      TabIndex        =   11
      Top             =   375
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
         Caption         =   "LblidDocumento"
         Height          =   6810
         Left            =   12525
         TabIndex        =   18
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame6 
            Height          =   1320
            Left            =   9120
            TabIndex        =   59
            Top             =   2505
            Width           =   2595
            Begin VB.OptionButton OptDe2 
               Caption         =   "x Cuenta"
               Enabled         =   0   'False
               Height          =   195
               Left            =   1455
               TabIndex        =   67
               Top             =   810
               Width           =   945
            End
            Begin VB.OptionButton OptDe1 
               Caption         =   "x Descipcion"
               Enabled         =   0   'False
               Height          =   195
               Left            =   135
               TabIndex        =   66
               Top             =   810
               Width           =   1230
            End
            Begin VB.CommandButton CmdDelCon 
               Caption         =   "Eliminar Destino"
               Enabled         =   0   'False
               Height          =   285
               Left            =   315
               TabIndex        =   61
               Top             =   450
               Width           =   1860
            End
            Begin VB.CommandButton CmdAddCon 
               Caption         =   "&Agregar Destino"
               Enabled         =   0   'False
               Height          =   285
               Left            =   315
               TabIndex        =   60
               Top             =   150
               Width           =   1860
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   360
            Left            =   9165
            TabIndex        =   29
            Top             =   1155
            Visible         =   0   'False
            Width           =   9030
            Begin VB.CommandButton CmdBusMedPag 
               Height          =   240
               Left            =   2190
               Picture         =   "FrmEgrCajBan.frx":277E
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   60
               Width           =   240
            End
            Begin VB.TextBox TxtIdMedioPago 
               Height          =   300
               Left            =   1545
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   3
               Text            =   "TxtIdMedioPago"
               Top             =   30
               Width           =   915
            End
            Begin VB.CommandButton CmdBusPro 
               Height          =   240
               Left            =   2190
               Picture         =   "FrmEgrCajBan.frx":28B0
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   60
               Width           =   240
            End
            Begin VB.Label LblDesMedPag 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDesMedPag"
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
               Left            =   2505
               TabIndex        =   33
               Top             =   30
               Width           =   6420
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Origen del Egreso"
               Height          =   195
               Index           =   6
               Left            =   0
               TabIndex        =   32
               Top             =   60
               Width           =   1260
            End
         End
         Begin VB.Frame Frame7 
            Height          =   2295
            Left            =   10410
            TabIndex        =   63
            Top             =   4185
            Width           =   1305
            Begin VB.CommandButton CmdEliminar 
               Caption         =   "&Eliminar"
               Height          =   630
               Left            =   90
               TabIndex        =   65
               Top             =   1185
               Width           =   1140
            End
            Begin VB.CommandButton CmdAgregar 
               Caption         =   "&Agregar Documento"
               Height          =   630
               Left            =   90
               TabIndex        =   64
               Top             =   510
               Width           =   1140
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg3 
            Height          =   1230
            Left            =   120
            TabIndex        =   7
            Top             =   2595
            Width           =   8940
            _cx             =   15769
            _cy             =   2170
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
            Rows            =   50
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEgrCajBan.frx":29E2
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
         Begin VB.CommandButton CmdMP 
            Height          =   240
            Left            =   2310
            Picture         =   "FrmEgrCajBan.frx":2AAF
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   1620
            Width           =   240
         End
         Begin VB.TextBox TxtTotal4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   9060
            Locked          =   -1  'True
            TabIndex        =   55
            Text            =   "TxtTotal4"
            Top             =   6495
            Width           =   1035
         End
         Begin VB.TextBox TxtTotal3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   8100
            Locked          =   -1  'True
            TabIndex        =   54
            Text            =   "TxtTotal3"
            Top             =   6495
            Width           =   975
         End
         Begin VB.TextBox TxtTotal2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7140
            Locked          =   -1  'True
            TabIndex        =   53
            Text            =   "TxtTotal2"
            Top             =   6495
            Width           =   975
         End
         Begin VB.TextBox TxtTotal1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6180
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   "TxtTotal1"
            Top             =   6495
            Width           =   975
         End
         Begin VB.OptionButton OptBanco 
            Caption         =   "Banco"
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
            Height          =   240
            Left            =   2910
            TabIndex        =   35
            Top             =   885
            Width           =   1170
         End
         Begin VB.OptionButton OptCaja 
            Caption         =   "Caja"
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
            Height          =   240
            Left            =   1680
            TabIndex        =   34
            Top             =   885
            Width           =   1170
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   360
            Left            =   120
            TabIndex        =   25
            Top             =   1155
            Visible         =   0   'False
            Width           =   9030
            Begin VB.CommandButton CmdBusCueBan 
               Height          =   240
               Left            =   3255
               Picture         =   "FrmEgrCajBan.frx":2BE1
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   60
               Width           =   240
            End
            Begin VB.TextBox TxtNumCue 
               Height          =   300
               Left            =   1545
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   2
               Text            =   "TxtNumCue"
               Top             =   30
               Width           =   1980
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Origen del Egreso"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   28
               Top             =   60
               Width           =   1260
            End
            Begin VB.Label LblBanco 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblBanco"
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
               Left            =   3570
               TabIndex        =   27
               Top             =   30
               Width           =   5355
            End
         End
         Begin VB.CommandButton CmdNumDoc 
            Height          =   240
            Left            =   11445
            Picture         =   "FrmEgrCajBan.frx":2D13
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1950
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.TextBox TxtImporte 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   7680
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   10
            Text            =   "TxtImporte"
            Top             =   3840
            Width           =   975
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   9120
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   6
            Text            =   "TxtNumDoc"
            Top             =   1920
            Width           =   2595
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   6705
            Picture         =   "FrmEgrCajBan.frx":2E45
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   870
            Width           =   240
         End
         Begin VB.Frame Frame5 
            Caption         =   "( Periodo )"
            Height          =   720
            Left            =   9120
            TabIndex        =   21
            Top             =   435
            Width           =   2595
            Begin VB.Label LblMes 
               Alignment       =   2  'Center
               Caption         =   "LblMes"
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
               Left            =   330
               TabIndex        =   22
               Top             =   255
               Width           =   1860
            End
         End
         Begin VB.CommandButton CmdBusDoc 
            Height          =   240
            Left            =   2310
            Picture         =   "FrmEgrCajBan.frx":2F77
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1950
            Width           =   240
         End
         Begin VB.CommandButton CmdBusCliente 
            Enabled         =   0   'False
            Height          =   240
            Left            =   5955
            Picture         =   "FrmEgrCajBan.frx":30A9
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3990
            Width           =   240
         End
         Begin VB.TextBox TxtProv 
            Height          =   300
            Left            =   1215
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   8
            Text            =   "TxtProv"
            Top             =   3960
            Width           =   5010
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   6060
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "TxtIdMon"
            Top             =   840
            Width           =   915
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchMov 
            Height          =   300
            Left            =   1665
            TabIndex        =   0
            Top             =   540
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
            Valor           =   "07/12/2007"
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   2190
            Left            =   120
            TabIndex        =   9
            Top             =   4275
            Width           =   10245
            _cx             =   18071
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEgrCajBan.frx":31DB
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
         Begin VB.TextBox TxtIdDoc 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   5
            Text            =   "TxtIdDoc"
            Top             =   1920
            Width           =   915
         End
         Begin VB.TextBox TxtMedPag 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   4
            Text            =   "TxtMedPag"
            Top             =   1590
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Destino del Egreso"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   62
            Top             =   2325
            Width           =   1335
         End
         Begin VB.Label LblMedPag 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblMedPag"
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
            Left            =   2625
            TabIndex        =   58
            Top             =   1590
            Width           =   6420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Medio de Pago"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   57
            Top             =   1620
            Width           =   1080
         End
         Begin VB.Label LblCtaHaber 
            Caption         =   "LblCtaHaber"
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   5460
            TabIndex        =   51
            Top             =   585
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label LblCtaDebe 
            Caption         =   "LblCtaDebe"
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   5460
            TabIndex        =   50
            Top             =   315
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Haber"
            Height          =   240
            Index           =   1
            Left            =   4365
            TabIndex        =   49
            Top             =   585
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Debe"
            Height          =   240
            Index           =   0
            Left            =   4365
            TabIndex        =   48
            Top             =   315
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   47
            Top             =   4005
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   5325
            TabIndex        =   46
            Top             =   885
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
            Height          =   195
            Index           =   4
            Left            =   6990
            TabIndex        =   45
            Top             =   3870
            Width           =   525
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
            Height          =   195
            Index           =   0
            Left            =   9135
            TabIndex        =   44
            Top             =   1665
            Width           =   1050
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Operacion"
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
            Left            =   60
            TabIndex        =   43
            Top             =   30
            Width           =   11610
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
            Left            =   7020
            TabIndex        =   42
            Top             =   840
            Width           =   1365
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emision"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   41
            Top             =   570
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Operacion"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   40
            Top             =   870
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   39
            Top             =   1950
            Width           =   825
         End
         Begin VB.Label LblIdCliente 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCliente"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   10320
            TabIndex        =   38
            Top             =   6585
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label LblDescDoc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescDoc"
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
            Left            =   2625
            TabIndex        =   37
            Top             =   1920
            Width           =   6420
         End
         Begin VB.Label LblIdCueBan 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCueBan "
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
            Left            =   7380
            TabIndex        =   36
            Top             =   525
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   45
         TabIndex        =   12
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   13
            Top             =   360
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11324
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Reg."
            Columns(0).DataField=   "numregi"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "T. M."
            Columns(1).DataField=   "motmov"
            Columns(1).NumberFormat=   "0.00"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Mov."
            Columns(2).DataField=   "fchope"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Importe"
            Columns(3).DataField=   "importe"
            Columns(3).NumberFormat=   "0.00"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "M"
            Columns(4).DataField=   "simbolo"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Origen"
            Columns(5).DataField=   "descori"
            Columns(5).NumberFormat=   "Short Date"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "T.D."
            Columns(6).DataField=   "abredoc"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Nº Documento"
            Columns(7).DataField=   "numdoc"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Nº Cuenta"
            Columns(8).DataField=   "numcue"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Banco"
            Columns(9).DataField=   "descban"
            Columns(9).NumberFormat=   "0.00"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1191"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1111"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1693"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1614"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1588"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1508"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=514"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=556"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=476"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=4207"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=4128"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=1005"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=926"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=2646"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2566"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=512"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=2408"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=2328"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(55)=   "Column(9).Width=2910"
            Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=2831"
            Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=514"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=74,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=62,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
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
         Begin VB.Label LblMes1 
            AutoSize        =   -1  'True
            Caption         =   "LblMes1"
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
            TabIndex        =   16
            Top             =   30
            Width           =   885
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Egresos"
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
            TabIndex        =   15
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblPeriodo 
            Alignment       =   2  'Center
            Caption         =   "LblMes"
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
            TabIndex        =   14
            Top             =   0
            Width           =   1860
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
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
End
Attribute VB_Name = "FrmEgrCajBan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim RstMov As New ADODB.Recordset
Dim xCuentaHaber As Integer        'para almacenar el codigo de la cuenta haber de la operacion
Dim Rst As New ADODB.Recordset     'para recorset temporales
Dim xSQL As String                 'sentencia SQL para los recorset temporales
Dim xFchPer As String
Dim RstTMPDoc As New ADODB.Recordset

Sub Eliminar()
    Dim Rpta, A As Integer
    Dim Rst As New ADODB.Recordset
    
    TabOne1.CurrTab = 0
    Rpta = MsgBox("Esta seguro de eliminar el movimiento seleccionado", vbInformation + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        'eliminamos el movimiento en el diario
        
        xCon.Execute "DELETE * From con_diario WHERE (((con_diario.idlib)=6) AND ((con_diario.idmov)= " & RstMov("id") & "))"
        
        'actualizamos el saldo del documento
        RST_Busq Rst, "SELECT * FROM con_cajabancodet WHERE id = " & RstMov("id") & "", xCon
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                xCon.Execute "UPDATE com_compras SET com_compras.impsal = [com_compras]![impsal]+" & Rst("impabo") & " WHERE (((com_compras.id)=" & Rst("iddoc") & "))"
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If
        
        'eliminamos el movimiento
        xCon.Execute "DELETE * FROM con_cajabanco WHERE id =" & RstMov("id") & ""
        
        MsgBox "El movimiento se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstMov.Requery
        Dg1.Refresh
    End If
End Sub

Sub MuestraSegundoTab()
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    
    Blanquea
    OptDe1.Value = True
    
    If RstMov.RecordCount = 0 Then Exit Sub
    
    TxtFchMov.Valor = RstMov("fchope")
    
    If RstMov("tipope") = 1 Then
        OptCaja.Value = True
    Else
        OptBanco.Value = True
        TxtMedPag.Text = NulosN(RstMov("idmedpag"))
        If NulosN(RstMov("idmedpag")) <> 0 Then
            LblMedPag.Caption = Busca_Codigo(RstMov("idmedpag"), "id", "descripcion", "con_mediopago", "N", xCon)
        End If
    End If
    
    TxtIdMon.Text = RstMov("idmon")
    LblMoneda.Caption = Busca_Codigo(RstMov("idmon"), "id", "descripcion", "mae_moneda", "N", xCon)
    
    If RstMov("tipope") = 1 Then
        TxtIdMedioPago.Text = RstMov("idori")
        LblDesMedPag.Caption = RstMov("descori")
        
        Set RstDet = BuscaConCriterio("SELECT con_origen.id, con_origen.idcue From con_origen WHERE (((con_origen.id)=" & RstMov("idori") & "))", xCon)
        LblCtaHaber.Caption = RstDet("idcue")
    Else
        TxtNumCue.Text = NulosC(RstMov("numcue"))
        LblBanco.Caption = NulosC(RstMov("descban"))
        LblIdCueBan.Caption = RstMov("idcueban")
        
        Set RstDet = BuscaConCriterio("SELECT con_bancocuenta.id, con_bancocuenta.idcuen FROM con_bancocuenta WHERE (((con_bancocuenta.id)=" & NulosN(LblIdCueBan.Caption) & "))", xCon)
        If RstDet.RecordCount <> 0 Then LblCtaHaber.Caption = RstDet("idcuen")         'Rst("cuenta")
    End If
    Set RstDet = Nothing
    
    TxtIdDoc.Text = RstMov("iddoc")
    TxtNumDoc.Text = RstMov("numdoc")
    TxtImporte.Text = Format(RstMov("importe"), "0.00")
    
    LblDescDoc.Caption = NulosC(RstMov("descdoc"))
    
    'Mostramos los destinos del egreso
    RST_Busq RstDet, "SELECT con_cajabancoorides.idorides, con_destino.descripcion AS descdest, con_destino.idcuen, con_planctas.cuenta, con_cajabancoorides.importe, " _
        & " con_planctas.descripcion AS desccta, con_destino.entgen FROM con_cajabanco LEFT JOIN (con_planctas RIGHT JOIN (con_cajabancoorides " _
        & " LEFT JOIN con_destino ON con_cajabancoorides.idorides = con_destino.id) ON con_planctas.id = con_destino.idcuen) ON con_cajabanco.id = con_cajabancoorides.id " _
        & " WHERE (((con_cajabanco.id)=" & RstMov("id") & ") AND ((con_cajabanco.tipmov)=2))", xCon
    
    'SELECT con_cajabancoorides.idorides, con_destino.descripcion AS descdest, con_destino.idcuen, con_planctas.cuenta, con_cajabancoorides.importe, " _
        & " con_planctas.descripcion AS desccta, con_destino.entgen, con_destino.iddoc FROM con_cajabanco LEFT JOIN (con_planctas RIGHT JOIN (con_cajabancoorides LEFT JOIN con_destino " _
        & " ON con_cajabancoorides.idorides = con_destino.id) ON con_planctas.id = con_destino.idcuen) ON con_cajabanco.id = con_cajabancoorides.id " _
        & " WHERE (((con_cajabanco.id)=" & RstMov("id") & ") AND ((con_cajabanco.tipmov)=2))", xCon

    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            Fg3.Rows = Fg3.Rows + 1
            Fg3.TextMatrix(A, 1) = NulosC(RstDet("descdest"))
            Fg3.TextMatrix(A, 2) = NulosN(RstDet("idorides"))
            Fg3.TextMatrix(A, 3) = NulosN(RstDet("idcuen"))
            Fg3.TextMatrix(A, 4) = NulosN(RstDet("entgen"))
            'Fg3.TextMatrix(A, 5) = NulosN(RstDet("iddoc"))
            Fg3.TextMatrix(A, 6) = Format(NulosN(RstDet("importe")), "0.00")
            RstDet.MoveNext
            If RstDet.EOF = True Then
                Exit For
            End If
        Next A
        SumarDestinos
    End If
    Set RstDet = Nothing
    
    RST_Busq RstDet, "SELECT con_cajabancodet.id, con_cajabancodet.iddoc, [con_percepcion]![numser]+'-'+[con_percepcion]![numdoc] AS numdoc, mae_documento.abrev AS nomdoc, " _
        & " con_cajabancodet.salant, con_cajabancodet.impabo, mae_moneda.simbolo, mae_prov.nombre, con_percepcion.fchdoc, con_percepcion.imptotper AS imptot, " _
        & " con_cajabancodet.idorigen, con_cajabancodet.idcue FROM con_cajabancodet LEFT JOIN (((con_percepcion LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id) LEFT JOIN mae_documento " _
        & " ON con_percepcion.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON con_percepcion.idmon = mae_moneda.id) ON con_cajabancodet.iddoc = con_percepcion.id " _
        & " Where (((con_cajabancodet.id) = " & RstMov("id") & ") And ((con_cajabancodet.idorigen) = 2)) " _
        & " Union " _
        & " SELECT con_cajabancodet.id, con_cajabancodet.iddoc, com_compras!numser+'-'+com_compras!numdoc AS numdoc, mae_documento.abrev AS nomdoc, " _
        & " con_cajabancodet.salant, con_cajabancodet.impabo, mae_moneda.simbolo, mae_prov.nombre, com_compras.fchdoc, com_compras.imptot, con_cajabancodet.idorigen, con_cajabancodet.idcue " _
        & " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (con_cajabancodet LEFT JOIN com_compras ON con_cajabancodet.iddoc = com_compras.id) " _
        & " ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
        & " WHERE (((con_cajabancodet.id) = " & RstMov("id") & ") AND ((con_cajabancodet.idorigen)=1))", xCon
    
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = NulosC(RstDet("nombre"))
            Fg2.TextMatrix(A, 2) = NulosC(RstDet("nomdoc"))
            Fg2.TextMatrix(A, 3) = Format(RstDet("fchdoc"), "dd/mm/yy")
            Fg2.TextMatrix(A, 4) = NulosC(RstDet("simbolo"))
            Fg2.TextMatrix(A, 5) = NulosC(RstDet("numdoc"))
            Fg2.TextMatrix(A, 6) = Format(RstDet("imptot"), "0.00")
            Fg2.TextMatrix(A, 7) = Format(RstDet("salant"), "0.00")
            Fg2.TextMatrix(A, 8) = Format(RstDet("impabo"), "0.00")
            Fg2.TextMatrix(A, 9) = Format((RstDet("salant") - RstDet("impabo")), "0.00")
            Fg2.TextMatrix(A, 10) = RstDet("iddoc")
            Fg2.TextMatrix(A, 11) = RstDet("idorigen")
            Fg2.TextMatrix(A, 15) = NulosN(Fg3.TextMatrix(Fg3.Rows - 1, 3)) 'NulosN(RstDet("idcue"))
            
            RstDet.MoveNext
            If RstDet.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    HallarTotales
End Sub

Private Sub CmdAddCon_Click()
    If Fg3.TextMatrix(Fg3.Rows - 1, 1) = "" Then Exit Sub
    
    Fg3.Rows = Fg3.Rows + 1
    Fg3_CellButtonClick Fg3.Rows - 1, 1
End Sub

Private Sub CmdAgregar_Click()
    If Fg3.Rows = 1 Then Exit Sub
    If Fg3.TextMatrix(Fg3.Row, 1) = "" Then
        MsgBox "Seleccione un destino para el egreso", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    CargarFacturasPorPagar Val(LblIdCliente.Caption)
End Sub

Function CadWhere(idDestino As Integer) As String
    'esta funcion permite filtrar a los proveedores cuyos documentos esten en la lista de documentos del destino del egreso
    Dim Rst2 As New ADODB.Recordset
    Dim A As Integer
    Dim xCadWhere As String
    'preparamos la linea WHERE de la consulta para ver los documentos que tenga asignado el destino del egreso
    RST_Busq Rst2, "SELECT * FROM con_destinodoc WHERE id = " & idDestino & "", xCon
    
    If Rst2.RecordCount <> 0 Then
        Rst2.MoveFirst
        For A = 1 To Rst2.RecordCount
            xCadWhere = xCadWhere + "(com_compras.tipdoc=" & Rst2("iddoc") & ")"
            Rst2.MoveNext
            If Rst2.EOF = True Then Exit For
            xCadWhere = xCadWhere + " OR "
        Next A
    End If
    Set Rst2 = Nothing
    CadWhere = xCadWhere
End Function

Private Sub CmdBusCliente_Click()
    If QueHace = 3 Then Exit Sub

    Dim xCadWhere As String
    
    xCadWhere = CadWhere(NulosN(Fg3.TextMatrix(Fg3.Row, 2)))
    
    If NulosC(xCadWhere) = "" Then
        MsgBox "El destino seleccionado no tiene documentos de compra asignado para su cancelacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":     xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"
    
    'buscamos los proveedores que tengan el documento especificado
    xForm.SQLCad = "SELECT DISTINCT mae_prov.id, mae_prov.numruc, mae_prov.nombre FROM mae_prov RIGHT JOIN com_compras ON mae_prov.id = com_compras.idpro " _
        & " WHERE ((com_compras.impsal<>0) " _
        & " AND " & xCadWhere & ")"

    xForm.Titulo = "Buscando Proveedores"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "nombre"
    xForm.CampoBusca = "nombre"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtProv.Text = xRs("nombre")
        LblIdCliente.Caption = xRs("id")
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Sub CargarFacturasPorPagar(IdProveedor As Integer)
    Dim xForm As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim xCadWhere As String
    
    xCadWhere = CadWhere(NulosN(Fg3.TextMatrix(Fg3.Row, 2)))
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(7, 5) As String
    
    xCampos(0, 0) = "Nº Documento":  xCampos(0, 1) = "numdoc":         xCampos(0, 2) = "1500":    xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "codsun":         xCampos(1, 2) = "600":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "Fch. Emi.":     xCampos(2, 1) = "fchdoc":         xCampos(2, 2) = "1000":    xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
    xCampos(3, 0) = "Proveedor":     xCampos(3, 1) = "nombre":         xCampos(3, 2) = "4000":    xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
    xCampos(4, 0) = "Moneda":        xCampos(4, 1) = "simbolo":        xCampos(4, 2) = "800":     xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
    xCampos(5, 0) = "Importe":       xCampos(5, 1) = "imptot":         xCampos(5, 2) = "1200":    xCampos(5, 3) = "N":    xCampos(5, 4) = "N"
    xCampos(6, 0) = "Saldo":         xCampos(6, 1) = "impsal":         xCampos(6, 2) = "1200":    xCampos(6, 3) = "N":    xCampos(6, 4) = "N"

    'xForm.SQLCad = "SELECT com_compras.id, mae_prov.nombre, mae_documento.codsun, com_compras.fchdoc, com_compras.fchven, " _
        & " [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, mae_moneda.simbolo, com_compras.imptot, com_compras.impsal, 'Compras' AS origen, 1 AS idori " _
        & " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) " _
        & " ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro WHERE (((com_compras.impsal)<>0) AND ((com_compras.idpro)=" & Val(LblIdCliente.Caption) & "))" _
        & " Union " _
        & " SELECT con_percepcion.id, mae_prov.nombre, mae_documento.codsun, con_percepcion.fchdoc, '' AS fchven, " _
        & " [con_percepcion]![numser]+'-'+[con_percepcion]![numdoc] AS numdoc, mae_moneda.simbolo, con_percepcion.imptotper AS imptot, con_percepcion.impsal, " _
        & " 'Percepcion' AS origen, 2 AS idori FROM ((con_percepcion LEFT JOIN mae_documento ON con_percepcion.tipdoc = mae_documento.id) " _
        & " LEFT JOIN mae_moneda ON con_percepcion.idmon = mae_moneda.id) LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id " _
        & " Where (((con_percepcion.impsal) <> 0) And ((con_percepcion.idcli) = " & Val(LblIdCliente.Caption) & "))"
    
    '((com_compras.tipdoc)=1 Or (com_compras.tipdoc)=3 Or (com_compras.tipdoc)=4 Or (com_compras.tipdoc)=8 Or (com_compras.tipdoc)=12 Or (com_compras.tipdoc)=14 Or (com_compras.tipdoc)=20 Or (com_compras.tipdoc)=40 Or (com_compras.tipdoc)=41 Or (com_compras.tipdoc)=95));
    
    If TxtProv.Text = "" Then
        xForm.SQLCad = "SELECT com_compras.id, mae_prov.nombre, mae_documento.codsun, com_compras.fchdoc, com_compras.fchven, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, " _
            & " mae_moneda.simbolo, com_compras.imptot, 'Compras' AS origen, 1 AS idori, com_compras.impsal FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento " _
            & " RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
            & " WHERE (((com_compras.impsal)<>0) AND " & " ( " & xCadWhere & "))" _
            & " Union " _
            & " SELECT con_percepcion.id, mae_prov.nombre, mae_documento.codsun, con_percepcion.fchdoc, '' AS fchven, [con_percepcion]![numser]+'-'+[con_percepcion]![numdoc] AS numdoc, " _
            & " mae_moneda.simbolo, con_percepcion.imptotper AS imptot, 'Percepcion' AS origen, 2 AS idori, con_percepcion.impsal FROM ((con_percepcion LEFT JOIN mae_documento " _
            & " ON con_percepcion.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON con_percepcion.idmon = mae_moneda.id) LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id " _
            & " WHERE (((con_percepcion.impsal)<>0))"

    Else
        xForm.SQLCad = "SELECT com_compras.id, mae_prov.nombre, mae_documento.codsun, com_compras.fchdoc, com_compras.fchven, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, " _
            & " mae_moneda.simbolo, com_compras.imptot, 'Compras' AS origen, 1 AS idori, com_compras.impsal FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento " _
            & " RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
            & " WHERE (((com_compras.impsal)<>0) AND ((com_compras.idpro)=" & NulosN(LblIdCliente.Caption) & ") AND " _
            & " ( " & xCadWhere & "))" _
            & " UNION " _
            & " SELECT con_percepcion.id, mae_prov.nombre, mae_documento.codsun, con_percepcion.fchdoc, '' AS fchven, [con_percepcion]![numser]+'-'+[con_percepcion]![numdoc] AS numdoc, " _
            & " mae_moneda.simbolo, con_percepcion.imptotper AS imptot, 'Percepcion' AS origen, 2 AS idori, con_percepcion.impsal FROM ((con_percepcion LEFT JOIN mae_documento " _
            & " ON con_percepcion.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON con_percepcion.idmon = mae_moneda.id) LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id " _
            & " WHERE (((con_percepcion.impsal)<>0) AND ((con_percepcion.idcli)=" & NulosN(LblIdCliente.Caption) & "))"
    End If

    '& " ( " (com_compras.tipdoc)=1 Or (com_compras.tipdoc)=4 Or (com_compras.tipdoc)=3 Or (com_compras.tipdoc)=8 Or (com_compras.tipdoc)=12 Or (com_compras.tipdoc)=20 Or (com_compras.tipdoc)=40 Or (com_compras.tipdoc)=41 Or (com_compras.tipdoc)=95));
    
    'SELECT com_compras.id, mae_prov.nombre, mae_documento.codsun, com_compras.fchdoc, com_compras.fchven, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, " _
        & " mae_moneda.simbolo, com_compras.imptot, 'Compras' AS origen, 1 AS idori, com_compras.impsal FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento " _
        & " RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
        & " WHERE ((com_compras.impsal<>0) AND (com_compras.idpro=971)) AND " _
        & xCadWhere

    
    
    xForm.Titulo = "Buscando Documentos de Proveedores"
    Set xForm.Coneccion = xCon
    Set xRs = Nothing
    Set xRs = xForm.Seleccionar(xCampos)
    If xRs.State = 1 Then
        Dim A As Integer
        Dim xFila As Integer
        xFila = Fg2.Rows - 1
        
        If xRs.RecordCount <> 0 Then
            xRs.MoveFirst
            For A = 1 To xRs.RecordCount
                Fg2.Rows = Fg2.Rows + 1
                xFila = xFila + 1
                Fg2.TextMatrix(xFila, 1) = xRs("nombre")
                Fg2.TextMatrix(xFila, 2) = xRs("codsun")
                Fg2.TextMatrix(xFila, 3) = xRs("fchdoc")
                Fg2.TextMatrix(xFila, 4) = xRs("simbolo")
                Fg2.TextMatrix(xFila, 5) = xRs("numdoc")
                Fg2.TextMatrix(xFila, 6) = Format(xRs("imptot"), "0.00")
                Fg2.TextMatrix(xFila, 7) = Format(xRs("impsal"), "0.00")
                Fg2.TextMatrix(xFila, 10) = xRs("id")     'id del documento
                Fg2.TextMatrix(xFila, 11) = xRs("idori")  'id de la entidad que origina el documento, puede ser proveedor, retencion, detraccion
                Fg2.TextMatrix(xFila, 12) = Fg3.TextMatrix(Fg3.Row, 4)       'id del destino, osea a donde mandamos el egreso
                Fg2.TextMatrix(xFila, 15) = Fg3.TextMatrix(Fg3.Row, 3)
                xRs.MoveNext
                If xRs.EOF = True Then Exit For
            Next A
        End If
    End If
End Sub

Private Sub CmdBusCueBan_Click()
    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtIdMon.Text) = "" Then
        MsgBox "No ha seleccionado la moneda para la operacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Banco":           xCampos(0, 1) = "desban":        xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº Cuenta":       xCampos(1, 1) = "numcue":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Moneda":          xCampos(2, 1) = "desmon":        xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nº Cta Contable": xCampos(3, 1) = "cuenta":        xCampos(3, 2) = "1500":         xCampos(3, 3) = "C"
    
    xForm.SQLCad = "SELECT mae_bancos.descripcion AS desban, con_bancocuenta.*, mae_moneda.descripcion AS desmon, " _
        & " con_planctas.cuenta FROM mae_bancos INNER JOIN (con_planctas RIGHT JOIN (con_bancocuenta LEFT JOIN mae_moneda " _
        & " ON con_bancocuenta.idmon = mae_moneda.id) ON con_planctas.id = con_bancocuenta.idcuen) ON " _
        & " mae_bancos.id = con_bancocuenta.idban Where (((con_bancocuenta.idmon) = " & Val(Val(TxtIdMon.Text)) & ")) " _
        & " ORDER BY mae_bancos.descripcion"
    
    xForm.Titulo = "Buscando Cuentas de Banco"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "desban"
    xForm.CampoBusca = "desban"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtNumCue.Text = xRs("numcue")
        LblIdCueBan = xRs("id")
        LblBanco.Caption = Trim(xRs("desban")) '"   Cuenta Nº " & xRs("numcue")
        xCuentaHaber = xRs("idcuen")
        LblCtaHaber.Caption = xRs("idcuen") 'xRs("cuenta")
        'TxtIdDoc.SetFocus
        'TxtIdMov.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDoc_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "id":           xCampos(1, 1) = "id":            xCampos(1, 2) = "1200":         xCampos(1, 3) = "N"
    
    If OptCaja.Value = True Then
        xForm.SQLCad = "SELECT mae_doccajaban.tipo, * From mae_doccajaban Where (((mae_doccajaban.tipo) = 1)) " _
            & " ORDER BY mae_doccajaban.descripcion"
    Else
        xForm.SQLCad = "SELECT mae_doccajaban.tipo, * From mae_doccajaban Where (((mae_doccajaban.tipo) = 2)) " _
            & " ORDER BY mae_doccajaban.descripcion"
    End If
    
    xForm.Titulo = "Buscando Documentos"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdDoc.Text = xRs("id")
        LblDescDoc.Caption = xRs("descripcion")
'        If xRs("selecciona") = -1 Then
'            TxtNumDoc.Locked = True
'        Else
'            TxtNumDoc.Locked = False
'        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMedPag_Click()
    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtIdMon.Text) = "" Then
        MsgBox "No ha especificado la moneda para la operacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Codigo":        xCampos(0, 1) = "id":            xCampos(0, 2) = "1000":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":   xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Cuenta":        xCampos(2, 1) = "desccuen":      xCampos(2, 2) = "3000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nº Cuenta":     xCampos(3, 1) = "cuenta":        xCampos(3, 2) = "1000":         xCampos(3, 3) = "C"
    
    'xForm.SQLCad = "SELECT con_destino.*, con_planctas.descripcion AS descta, con_planctas.cuenta " _
        & " FROM con_planctas RIGHT JOIN con_destino ON con_planctas.id = con_destino.idcuenta " _
        & " WHERE (((con_destino.idmon)=" & Val(TxtIdMon.Text) & "))"
    
    xForm.SQLCad = "SELECT con_origen.*, con_planctas.cuenta, con_planctas.descripcion AS desccuen, con_origen.idmon FROM con_planctas RIGHT JOIN con_origen " _
        & " ON con_planctas.id = con_origen.idcue WHERE (((con_origen.idmon)=" & Val(TxtIdMon.Text) & ") AND ((con_origen.tipmov)=2))"

    
    xForm.Titulo = "Buscando Origen del Egreso"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "id"
    xForm.CampoBusca = "id"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblDesMedPag.Caption = xRs("descripcion")
        TxtIdMedioPago.Text = xRs("id")
        xCuentaHaber = xRs("idcue")
        LblCtaHaber.Caption = xRs("idcue") 'xRs("cuenta")
        
        'If NulosN(xRs("iddoc")) <> 0 Then
        '    TxtIdDoc.Text = xRs("iddoc")
        '    LblDescDoc.Caption = Busca_Codigo(xRs("iddoc"), "id", "descripcion", "mae_doccajaban", "N", xCon)
        'Else
        '    TxtIdDoc.Text = ""
        '    LblDescDoc.Caption = ""
        'End If
        'TxtIdDoc.SetFocus
        'TxtIdMov.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "id":            xCampos(0, 2) = "1200":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
    
    'filtramos por tipo de movimiento  = 1 (Ingreso)
    xForm.SQLCad = "SELECT * FROM  mae_moneda ORDER BY descripcion"

    xForm.Titulo = "Buscando Moneda"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "id"
    xForm.CampoBusca = "id"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdMon.Text = xRs("id")
        LblMoneda.Caption = xRs("descripcion")
        If OptCaja.Value = True Then
            TxtIdMedioPago.SetFocus
        Else
            TxtNumCue.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

'Private Sub CmdBusMovi_Click()
'    If QueHace = 3 Then Exit Sub
'
'    Dim xform As New eps_librerias.FormBuscar
'
'    Dim xRs As New ADODB.Recordset
'    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
'
'    Dim xCampos(4, 4) As String
'
'    xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "id":            xCampos(0, 2) = "1200":         xCampos(0, 3) = "N"
'    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
'    xCampos(2, 0) = "Cuenta":       xCampos(2, 1) = "descuen":       xCampos(2, 2) = "3000":         xCampos(2, 3) = "C"
'    xCampos(3, 0) = "Nº Cuenta":    xCampos(3, 1) = "cuenta":        xCampos(3, 2) = "1200":         xCampos(3, 3) = "C"
'
'    'filtramos por tipo de movimiento  = 2 (Egresos)
'    xform.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, con_origen.id, con_origen.idmon, " _
'        & " con_origen.descripcion, con_origen.idcue, con_origen.tipmov FROM con_planctas INNER JOIN " _
'        & " con_origen ON con_planctas.id = con_origen.idcue WHERE (con_origen.tipmov = 2) " _
'        & " AND (con_origen.idmon = " & Val(TxtIdMon.Text) & ")"
'
'    xform.Titulo = "Buscando Origen del Egreso"
'    xform.FormaBusca = Principio
'    xform.Criterio = ""
'    xform.Ordenado = "id"
'    xform.CampoBusca = "id"
'    Set xform.Coneccion = xCon
'    Set xRs = xform.BuscarReg(xCampos)
'    If xRs.State = 1 Then
'        'TxtIdMov.Text = xRs("id")
'        'LblDescMov.Caption = xRs("descripcion")
'        LblCtaDebe.Caption = xRs("idcue") 'xRs("cuenta")
'        'If OptCaja.Value = True Then
'        '    TxtIdMedioPago.SetFocus
'        'Else
'        '    TxtNumCue.SetFocus
'        'End If
'        TxtIdDoc.SetFocus
'    End If
'    Set xform = Nothing
'    Set xRs = Nothing
'End Sub

Private Sub CmdBusPro_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Codigo":        xCampos(0, 1) = "id":            xCampos(0, 2) = "1000":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":   xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Cuenta":        xCampos(2, 1) = "descta":        xCampos(2, 2) = "3000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nº Cuenta":     xCampos(3, 1) = "cuenta":        xCampos(3, 2) = "1000":         xCampos(3, 3) = "C"
    
    If OptCaja.Value = True Then
        xForm.SQLCad = "SELECT con_destino.*, con_planctas.descripcion AS descta, con_planctas.cuenta " _
            & " FROM con_planctas RIGHT JOIN con_destino ON con_planctas.id = con_destino.idcuenta " _
            & " WHERE (((con_destino.idmon)=1))"
    Else
        xForm.SQLCad = "SELECT con_destino.*, con_planctas.descripcion AS descta, con_planctas.cuenta " _
            & " FROM con_planctas RIGHT JOIN con_destino ON con_planctas.id = con_destino.idcuenta " _
            & " WHERE (((con_destino.idmon)=2))"
    End If
    
    xForm.Titulo = "Buscando Destino del Ingreso"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "id"
    xForm.CampoBusca = "id"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblDesMedPag.Caption = xRs("descripcion")
        TxtIdMedioPago.Text = xRs("id")
        xCuentaHaber = xRs("idcuenta")
        LblCtaHaber.Caption = xRs("cuenta")
        
        If NulosN(xRs("iddoc")) <> 0 Then
            TxtIdDoc.Text = xRs("iddoc")
            LblDescDoc.Caption = Busca_Codigo(xRs("iddoc"), "id", "descripcion", "mae_doccajaban", "N", xCon)
            TxtNumDoc.Text = HallarNumeroDocumentoCaja(Val(TxtIdDoc.Text))
            TxtIdDoc.SetFocus
        Else
            
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Function HallarNumeroDocumentoCaja(CodigoDocumento As Integer) As String
    Dim Rst  As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT con_cajabanco.iddoc, con_cajabanco.numdoc From con_cajabanco " _
        & " WHERE (((con_cajabanco.iddoc)=" & CodigoDocumento & ")) ORDER BY numdoc", xCon

    If Rst.RecordCount = 0 Then
        HallarNumeroDocumentoCaja = "000001"
    Else
        Rst.MoveLast
        HallarNumeroDocumentoCaja = Format(Val(Rst("numdoc")) + 1, "000000")
    End If
End Function

Private Sub CmdDelCon_Click()
    If Fg3.Rows = 1 Then Exit Sub
    Fg3.RemoveItem Fg3.Row
End Sub

Private Sub CmdEliminar_Click()
    If Fg2.Rows = 1 Then
        MsgBox "No hay documento para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Fg2.RemoveItem Fg2.Row
End Sub

Private Sub CmdMP_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "id":            xCampos(0, 2) = "1000":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "7000":         xCampos(1, 3) = "C"
    
    'filtramos por tipo de movimiento  = 1 (Ingreso)
    xForm.SQLCad = "SELECT * FROM  con_mediopago ORDER BY descripcion"

    xForm.Titulo = "Buscando Medio de Pago"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "id"
    xForm.CampoBusca = "id"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtMedPag.Text = xRs("id")
        LblMedPag.Caption = xRs("descripcion")
        TxtIdDoc.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdNumDoc_Click()
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(5, 4) As String
    
    xCampos(0, 0) = "Documento":       xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "3500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº Documento":    xCampos(1, 1) = "numdoc":        xCampos(1, 2) = "1300":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Fch. Ope.":       xCampos(2, 1) = "fchope":        xCampos(2, 2) = "1100":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Importe":         xCampos(3, 1) = "importe":       xCampos(3, 2) = "1100":         xCampos(3, 3) = "N"
    xCampos(4, 0) = "Saldo":           xCampos(4, 1) = "saldo":         xCampos(4, 2) = "1100":         xCampos(4, 3) = "N"
    
    xForm.SQLCad = "SELECT con_cajabanco.id, con_cajabanco.iddoc, con_cajabanco.idmon, con_cajabanco.numdoc, " _
        & " con_cajabanco.fchope, con_cajabanco.importe, con_cajabanco.saldo, mae_doccajaban.descripcion " _
        & " FROM mae_doccajaban RIGHT JOIN (mae_moneda RIGHT JOIN con_cajabanco ON mae_moneda.id = con_cajabanco.idmon) " _
        & " ON mae_doccajaban.id = con_cajabanco.iddoc WHERE (((con_cajabanco.iddoc)=" & Val(TxtIdDoc.Text) & ") " _
        & " AND ((con_cajabanco.idmon)=" & Val(TxtIdMon.Text) & ") AND ((con_cajabanco.saldo)<>0))"
    
    xForm.Titulo = "Buscando " + Trim(LblDescDoc.Caption)
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "numdoc"
    xForm.CampoBusca = "numdoc"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtNumDoc.Text = xRs("numdoc")
        TxtImporte.Text = xRs("saldo")
        TxtImporte.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
        RstMov.Sort = "numregi DESC"
    End If
    If ColIndex = 1 Then
        RstMov.Sort = "motmov"
    End If
    If ColIndex = 2 Then
        RstMov.Sort = "fchope"
    End If
    If ColIndex = 3 Then
        RstMov.Sort = "importe"
    End If
    If ColIndex = 4 Then
        RstMov.Sort = "simbolo"
    End If
    If ColIndex = 5 Then
        RstMov.Sort = "descori"
    End If
    If ColIndex = 6 Then
        RstMov.Sort = "abredoc"
    End If
    If ColIndex = 7 Then
        RstMov.Sort = "numdoc"
    End If
    If ColIndex = 8 Then
        RstMov.Sort = "numcue"
    End If
    If ColIndex = 9 Then
        RstMov.Sort = "descban"
    End If
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Fg3.TextMatrix(Fg3.Row, 4) = "5" Then Exit Sub
    
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    If Col = 2 Then
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "abrev":            xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
            
        xForm.SQLCad = "SELECT * FROM mae_documento"
        
        xForm.Titulo = "Buscando Tipo de Documento"
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            Fg2.TextMatrix(Row, 2) = NulosC(xRs("codsun"))
            Fg2.TextMatrix(Row, 13) = xRs("id")
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
    
    If Col = 4 Then
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "id":            xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
            
        xForm.SQLCad = "SELECT * FROM mae_moneda "
        
        xForm.Titulo = "Buscando Tipo de Documento"
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            Fg2.TextMatrix(Row, 4) = NulosC(xRs("simbolo"))
            Fg2.TextMatrix(Row, 14) = xRs("id")
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
    
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Fg3.TextMatrix(Fg3.Row, 4) = "5" Then
        If Col <= 7 Then Exit Sub
    End If
    If Col = 6 Then
        Fg2.TextMatrix(Fg2.Row, 6) = Format(Fg2.TextMatrix(Fg2.Row, 6), "0.00")
        Fg2.TextMatrix(Fg2.Row, 7) = Fg2.TextMatrix(Fg2.Row, 6)
        HallarTotales
    End If
    
    If Col = 8 Then
        Fg2.TextMatrix(Fg2.Row, 8) = Format(Fg2.TextMatrix(Fg2.Row, 8), "0.00")
        Fg2.TextMatrix(Fg2.Row, 9) = NulosN(Fg2.TextMatrix(Fg2.Row, 7)) - NulosN(Fg2.TextMatrix(Fg2.Row, 8))
        Fg2.TextMatrix(Fg2.Row, 9) = Format(Fg2.TextMatrix(Fg2.Row, 9), "0.00")
        HallarTotales
    End If
    
    If Fg3.TextMatrix(Fg3.Row, 4) = "5" Then
        Fg3.TextMatrix(Fg3.Row, 5) = TxtTotal3.Text
    End If
End Sub

Private Sub Fg2_EnterCell()
    If Fg3.TextMatrix(Fg3.Row, 4) = "5" Then Exit Sub
    Fg2.Editable = flexEDKbdMouse
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        CmdAgregar_Click
    End If
    If KeyCode = 46 Then
        CmdEliminar_Click
        HallarTotales
    End If
End Sub

Private Sub Fg3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        Dim xForm As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        Dim xCampos(3, 4) As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        
        If OptDe1.Value = True Then
            xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "3000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Cuenta":       xCampos(1, 1) = "descuen":       xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Nº Cuenta":    xCampos(2, 1) = "cuenta":        xCampos(2, 2) = "1200":         xCampos(2, 3) = "C"
            xForm.Ordenado = "descripcion"
            xForm.CampoBusca = "descripcion"
        End If
        If OptDe2.Value = True Then
            xCampos(0, 0) = "Nº Cuenta":    xCampos(0, 1) = "cuenta":        xCampos(0, 2) = "1200":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Cuenta":       xCampos(1, 1) = "descuen":       xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Descripcion":  xCampos(2, 1) = "descripcion":   xCampos(2, 2) = "3000":         xCampos(2, 3) = "C"
            xForm.Ordenado = "cuenta"
            xForm.CampoBusca = "cuenta"
        End If
        
        xForm.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, con_destino.id, con_destino.idmon, con_destino.descripcion, con_destino.idcuen, " _
            & " con_destino.tipmov, con_destino.entgen, (SELECT Count([iddoc]) AS Expr1 From con_destinodoc WHERE (con_destinodoc.id=con_destino.id)) AS numdocasi " _
            & " FROM con_planctas RIGHT JOIN con_destino ON con_planctas.id = con_destino.idcuen WHERE (((con_destino.idmon)=" & Val(TxtIdMon.Text) & ") AND ((con_destino.tipmov)=2))"
        
        'SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, con_destino.id, con_destino.idmon, con_destino.descripcion, con_destino.idcuen, " _
            & " con_destino.tipmov, con_destino.entgen FROM con_planctas RIGHT JOIN con_destino ON con_planctas.id = con_destino.idcuen WHERE (((con_destino.idmon)=" & Val(TxtIdMon.Text) & ") " _
            & " AND ((con_destino.tipmov)=2))"

        xForm.Titulo = "Buscando Destino del Egreso"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            Fg3.TextMatrix(Row, 1) = xRs("descripcion")
            Fg3.TextMatrix(Row, 2) = xRs("id")
            Fg3.TextMatrix(Row, 3) = xRs("idcuen")
            Fg3.TextMatrix(Row, 4) = NulosN(xRs("entgen"))
            Fg3.TextMatrix(Row, 5) = NulosN(xRs("numdocasi"))   'especifica el numero de documentos asignado al destino
            
            If xRs("entgen") = 5 Then
                CmdBusCliente.Enabled = True
            Else
                CmdBusCliente.Enabled = False
                TxtProv.Text = ""
                LblIdCliente.Caption = ""
            End If
            
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg3_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 6 Then
        Fg3.TextMatrix(Fg3.Row, 6) = Format(Fg3.TextMatrix(Fg3.Row, 6), "0.00")
        SumarDestinos
    End If
End Sub

Sub SumarDestinos()
    Dim A As Integer
    Dim xTot As Double
    For A = 1 To Fg3.Rows - 1
        xTot = xTot + NulosN(Fg3.TextMatrix(A, 6))
    Next A
    
    TxtImporte.Text = Format(xTot, "0.00")
End Sub

Private Sub Fg3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        CmdAddCon_Click
    End If
    If KeyCode = 46 Then
        CmdDelCon_Click
    End If
End Sub

Private Sub Fg3_RowColChange()
    If NulosN(Fg3.TextMatrix(Fg3.Row, 4)) = 5 Then
        CmdBusCliente.Enabled = True
        CmdAgregar.Enabled = True
        CmdEliminar.Enabled = True
        Fg2.Editable = flexEDKbdMouse
    Else
        CmdBusCliente.Enabled = False
        CmdAgregar.Enabled = False
        CmdEliminar.Enabled = False
        Fg2.Editable = flexEDNone
        TxtProv.Text = ""
        LblIdCliente.Caption = ""
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim Rpta As Integer
        SeEjecuto = True
        
        CargarRSTCom
        
        Set Dg1.DataSource = RstMov
        OpcionesPeriodo
        If RstMov.RecordCount = 0 Then
            Rpta = MsgBox("No se ha registrado ninguna pago, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            If Rpta = vbYes Then
                Nuevo
            Else
                xMes = SeleccionaMes(xCon)
                OpcionesPeriodo
                CargarRSTCom
                
                If RstMov.State <> 0 Then
                    If RstMov.RecordCount = 0 Then
                        Rpta = MsgBox("No se ha registrado ningun movimiento, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
                        If Rpta = vbYes Then
                            Nuevo
                        Else
                            Set RstMov = Nothing
                            Unload Me
                        End If
                    End If
                Else
                    Set RstMov = Nothing
                    Unload Me
                End If
            End If
        Else
            OpcionesPeriodo
            Dg1.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Frame3.BackColor = &H8000000F
    Frame4.BackColor = &H8000000F
    
    Fg2.ColWidth(10) = 0
    Fg2.ColWidth(11) = 0
    Fg2.ColWidth(12) = 0
    Fg2.ColWidth(13) = 0
    Fg2.ColWidth(14) = 0
    Fg2.ColWidth(15) = 0
    
    Fg3.ColWidth(2) = 0
    Fg3.ColWidth(3) = 0
    Fg3.ColWidth(4) = 0
    Fg3.ColWidth(5) = 0
End Sub

Sub OpcionesPeriodo()
    Dim NomMes As String
    Dim Cerrado As Boolean
    Dim xFechaMes As String
    Dim xFchIni, xFchFin As Date
    Dim Rpta As Integer
    
    LblMes.Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
    Cerrado = Busca_Codigo(xMes, "id", "cerrado", "con_meses", "N", xCon)
    
    If Cerrado = True Then
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = False
        Toolbar1.Buttons(3).Visible = False
        Toolbar1.Buttons(4).Visible = False
    Else
        Toolbar1.Buttons(1).Visible = True
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(3).Visible = True
        Toolbar1.Buttons(4).Visible = True
    End If
    If xMes <> 0 Then
        xFechaMes = "01/" + Trim(Format(xMes, "00")) + "/" + Trim(Format(Year(Date), "0000"))
        xFchIni = xFechaMes
        xFchFin = Format(HallaDiasMes(CDate(xFechaMes)), "00") + "/" + Mid(xFechaMes, 4, 7)
        LblMes.Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
        LblMes1.Caption = LblMes.Caption
    End If
End Sub

Private Sub OptBanco_Click()
    Frame4.Visible = False
    Frame3.Visible = True
    Frame3.Left = 120
    Frame3.Top = 1155
    TxtNumCue.Locked = True
    TxtMedPag.Locked = False
    CmdMP.Enabled = True
    TxtIdDoc.Text = ""
    LblDescDoc.Caption = ""
End Sub

Private Sub OptCaja_Click()
    Frame3.Visible = False
    Frame4.Visible = True
    Frame4.Left = 120
    Frame4.Top = 1155
    TxtMedPag.Text = ""
    LblMedPag.Caption = ""
    TxtMedPag.Locked = True
    CmdMP.Enabled = False
    TxtIdDoc.Text = ""
    LblDescDoc.Caption = ""
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            MuestraSegundoTab
        End If
    End If
End Sub

Sub Modificar()
    QueHace = 2
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Operacion"
    Blanquea
    Bloquea
    Fg2.Rows = 1
    
    MuestraSegundoTab
    Fg2.Editable = flexEDKbdMouse
    Fg3.Editable = flexEDKbdMouse
    TxtFchMov.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then
        Modificar
    End If
    
    If Button.Index = 3 Then
        Eliminar
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstMov.Requery
            Dg1.Refresh
        End If
    End If
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 11 Then
        TabOne1.CurrTab = 0
        xMes = SeleccionaMes(xCon)
        OpcionesPeriodo
        If xMes = 0 Then
            Set RstMov = Nothing
            Unload Me
            Exit Sub
        End If
        
        CargarRSTCom
        'If RstMov.RecordCount = 0 Then
        '    Set RstMov = Nothing
        '    Unload Me
        'End If
    End If
    
    If Button.Index = 15 Then
        Set RstMov = Nothing
        Unload Me
    End If
End Sub

Sub CargarRSTCom()
    Dim Rpta As Integer
    
    If xMes = 0 Then
        Set RstMov = Nothing
        Exit Sub
    End If
    xFchPer = "01/" + Format(Trim(Str(xMes)), "00") + "/" + Trim(Str(AnoTra))
    
    RST_Busq RstMov, "SELECT con_cajabanco.*, mae_moneda.simbolo, mae_doccajaban.descripcion AS descdoc, con_bancocuenta.numcue, mae_bancos.descripcion AS descban, " _
        & " con_origen.descripcion AS descori, con_origen.idcue AS idcueori, mae_doccajaban.abrev AS abredoc, IIf(con_cajabanco!tipope=1,'Caja','Banco') AS motmov, " _
        & " con_cajabanco.tipmov, Mid([con_cajabanco]![numreg],1,2)+[mae_libros]![codsun]+Mid([con_cajabanco]![numreg],3,4) AS numregi FROM (mae_moneda RIGHT JOIN " _
        & " (mae_bancos RIGHT JOIN (((con_cajabanco LEFT JOIN con_bancocuenta ON con_cajabanco.idcueban = con_bancocuenta.id) LEFT JOIN mae_doccajaban " _
        & " ON con_cajabanco.iddoc = mae_doccajaban.id) LEFT JOIN con_origen ON con_cajabanco.idori = con_origen.id) ON mae_bancos.id = con_bancocuenta.idban) " _
        & " ON mae_moneda.id = con_cajabanco.idmon) LEFT JOIN mae_libros ON con_cajabanco.idlib = mae_libros.id WHERE (((con_cajabanco.tipmov)=2) " _
        & " AND ((con_cajabanco.fchreg)=CDate('" & xFchPer & "'))) ORDER BY con_cajabanco.id DESC", xCon
    Set Dg1.DataSource = RstMov
End Sub

Sub Cancelar()
    ActivaTool
    Bloquea
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    QueHace = 3
End Sub

Function Grabar() As Boolean
    If NulosC(TxtFchMov.Valor) = "" Then
        MsgBox "No ha especificado la fecha de la operacion", vbInformation + vbOKOnly + vbDefaultButton1
        TxtFchMov.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtIdMon.Text) = "" Then
        MsgBox "No ha especificado la moneda", vbInformation + vbOKOnly + vbDefaultButton1
        TxtIdMon.SetFocus
        Exit Function
    End If
    
    If OptCaja.Value = True Then
        If NulosC(TxtIdMedioPago.Text) = "" Then
            MsgBox "No ha especificado el destino del movimiento", vbInformation + vbOKOnly + vbDefaultButton1
            TxtIdMedioPago.SetFocus
            Exit Function
        End If
    Else
        If NulosC(TxtNumCue.Text) = "" Then
            MsgBox "No ha especificado el destino del movimiento", vbInformation + vbOKOnly + vbDefaultButton1
            TxtNumCue.SetFocus
            Exit Function
        End If
    End If

    If NulosC(TxtImporte.Text) = "" Then
        MsgBox "No ha especificado el importe del documento", vbInformation + vbOKOnly + vbDefaultButton1
        TxtImporte.SetFocus
        Exit Function
    End If

    If Fg3.Rows = 1 Then
        MsgBox "No ha especificado el destino del egreso", vbInformation + vbOKOnly + vbDefaultButton1
        Fg3.SetFocus
        Exit Function
    End If
    
    'If Fg2.Rows = 1 Then
    '    MsgBox "No ha especificado que documentos se estan cancelando con el movimiento", vbInformation + vbOKOnly + vbDefaultButton1
    '    TxtProv.SetFocus
    '    Exit Function
    'End If
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstOri As New ADODB.Recordset
    Dim xId, A As Integer
    Dim xNumAsiento As String
    Dim RstDia As New ADODB.Recordset
    Dim Rst As New ADODB.Recordset

On Error GoTo LaCague
    xCon.BeginTrans
    
    If QueHace = 1 Then
        xNumAsiento = NuevoNumAsiento(6, xMes, xCon)
        xId = HallaCodigoTabla("con_cajabanco", xCon, "id")
        
        RST_Busq RstCab, "SELECT * FROM con_cajabanco", xCon
        RST_Busq RstDet, "SELECT * FROM con_cajabancodet", xCon
        RST_Busq RstDia, "SELECT * FROM con_diario", xCon
        RST_Busq RstOri, "SELECT * FROM con_cajabancoorides", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstMov("id")
        xNumAsiento = DevuelveNumAsiento(6, RstMov("id"), xMes, xCon)
        xCon.Execute "DELETE * FROM con_cajabancodet WHERE id = " & RstMov("id") & ""
        xCon.Execute "DELETE * FROM con_diario WHERE idmov = " & RstMov("id") & " AND idlib = 6"
        xCon.Execute "DELETE * FROM con_cajabancoorides WHERE id = " & RstMov("id") & " "
        
        
        'actualizamos el saldo del documento cancelados
        RST_Busq Rst, "SELECT * FROM con_cajabancodet WHERE id = " & RstMov("id") & "", xCon
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                xCon.Execute "UPDATE com_compras SET com_compras.impsal = [com_compras]![impsal]+" & Rst("impabo") & " WHERE (((com_compras.id)=" & Rst("iddoc") & "))"
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If
        Set Rst = Nothing
        
        RST_Busq RstCab, "SELECT * FROM con_cajabanco WHERE id = " & RstMov("id") & "", xCon
        RST_Busq RstDet, "SELECT * FROM con_cajabancodet", xCon
        RST_Busq RstDia, "SELECT * FROM con_diario", xCon
        RST_Busq RstOri, "SELECT * FROM con_cajabancoorides", xCon
    End If
    
    If OptCaja.Value = True Then
        RstCab("tipope") = 1
        RstCab("idori") = Val(TxtIdMedioPago.Text)
    Else
        RstCab("tipope") = 2
        RstCab("idcueban") = Val(LblIdCueBan.Caption)
        RstCab("idmedpag") = NulosN(TxtMedPag.Text)
    End If
    RstCab("idlib") = 6
    RstCab("iddoc") = Val(TxtIdDoc.Text)
    RstCab("numdoc") = NulosC(TxtNumDoc.Text)
    RstCab("idmon") = Val(TxtIdMon.Text)
    RstCab("importe") = Val(TxtImporte.Text)
    RstCab("fchope") = TxtFchMov.Valor
    RstCab("fchreg") = CDate("01/" + Format(xMes, "00") + "/" + AnoTra)
    RstCab("tipmov") = 2
    
    RstCab.Update
    
    'grabamos los destinos del egreso
    For A = 1 To Fg3.Rows - 1
        RstOri.AddNew
        
        RstOri("id") = xId
        RstOri("idorides") = Fg3.TextMatrix(A, 2)
        RstOri("importe") = Fg3.TextMatrix(A, 6)
        RstOri.Update
    Next A
    
    'grabamos los documentos del egreso
    For A = 1 To Fg2.Rows - 1
        RstDet.AddNew
        
        RstDet("id") = xId
        RstDet("iddoc") = Fg2.TextMatrix(A, 10)
        RstDet("salant") = Fg2.TextMatrix(A, 7)
        RstDet("impabo") = Fg2.TextMatrix(A, 8)
        RstDet("idorigen") = Val(Fg2.TextMatrix(A, 11))
        RstDet("idcue") = Val(Fg2.TextMatrix(A, 15))
        
        RstDet.Update
        
        If Val(Fg2.TextMatrix(A, 11)) = 1 Then
            xCon.Execute "UPDATE com_compras SET com_compras.impsal = " & Val(Fg2.TextMatrix(A, 9)) & "" _
                & " WHERE (((com_compras.id)=" & Val(Fg2.TextMatrix(A, 10)) & "))"
        End If
        
        If Val(Fg2.TextMatrix(A, 11)) = 2 Then
            xCon.Execute "UPDATE con_percepcion SET con_percepcion.impsal = " & Val(Fg2.TextMatrix(A, 9)) & "" _
                & " WHERE (((con_percepcion.id)=" & Val(Fg2.TextMatrix(A, 10)) & "))"
        End If
    Next A
    
    Dim ValTipCam As Double
    ValTipCam = 0
    
    Dim B As Integer
    
    'grabamos el diario del movimiento
    '-----------------------------------------
    'ESCRIBIRMOS LA CUENTA DEBE DEL MOVIMIENTO
    For B = 1 To Fg3.Rows - 1
        If NulosN(Fg3.TextMatrix(B, 4)) = 5 Then
            'si es documentos de proveedor recorremos el grid para grabar el detalle
            For A = 1 To Fg2.Rows - 1
                RstDia.AddNew
                RstDia("año") = AnoTra
                RstDia("idmes") = xMes
                RstDia("idlib") = 6
                RstDia("idmov") = xId
                RstDia("numasi") = xNumAsiento
                RstDia("tc") = ValTipCam
                RstDia("iddocpro") = NulosN(Fg2.TextMatrix(A, 10))
                RstDia("idcue") = NulosN(Fg3.TextMatrix(B, 3))  'NulosN(Fg2.TextMatrix(A, 15))
                RstDia("fchasi") = "01/" + Format(xMes, "00") + "/" + Format(AnoTra, "0000")
                RstDia("fchdoc") = TxtFchMov.Valor
                If NulosC(TxtIdMon.Text) = "1" Then
                    RstDia("impdebsol") = NulosN(Fg2.TextMatrix(A, 8)) 'Val(TxtTotal3.Text)
                    RstDia("impdebdol") = 0
                Else
                    RstDia("impdebdol") = NulosN(Fg2.TextMatrix(A, 8))
                End If
                RstDia.Update
            Next A
        Else
            RstDia.AddNew
            RstDia("año") = AnoTra
            RstDia("idmes") = xMes
            RstDia("idlib") = 6
            RstDia("idmov") = xId
            RstDia("numasi") = xNumAsiento
            RstDia("tc") = ValTipCam
            RstDia("iddocpro") = 0
            RstDia("idcue") = NulosN(Fg3.TextMatrix(B, 3))
            RstDia("fchasi") = "01/" + Format(xMes, "00") + "/" + Format(AnoTra, "0000")
            RstDia("fchdoc") = TxtFchMov.Valor
            
            If NulosC(TxtIdMon.Text) = "1" Then
                RstDia("impdebsol") = NulosN(Fg3.TextMatrix(B, 6)) 'Val(TxtTotal3.Text)
                RstDia("impdebdol") = 0
            Else
                'RstDia("impdebsol") = Val(Fg2.TextMatrix(A, 8)) * Val(LblTipoCambio.Caption)
                RstDia("impdebdol") = NulosN(Fg3.TextMatrix(B, 6))
            End If
            RstDia.Update
        End If
    Next B
    
    '------------------------------------------
    'ESCRIBIRMOS LA CUENTA HABER DEL MOVIMIENTO
    RstDia.AddNew
    RstDia("año") = AnoTra
    RstDia("idmes") = xMes
    RstDia("idlib") = 6
    RstDia("idmov") = xId
    RstDia("numasi") = xNumAsiento
    RstDia("tc") = ValTipCam
    RstDia("idcue") = Val(LblCtaHaber.Caption)
    RstDia("fchasi") = "01/" + Format(xMes, "00") + "/" + Format(AnoTra, "0000")
    RstDia("fchdoc") = TxtFchMov.Valor
    
    If NulosC(TxtIdMon.Text) = "1" Then
        RstDia("imphabsol") = NulosN(TxtImporte.Text)
        RstDia("imphabdol") = 0
    Else
        'RstDia("imphabsol") = Val(TxtTotal3.Text) * Val(LblTipoCambio.Caption)
        RstDia("imphabdol") = NulosN(TxtImporte.Text)
    End If
    RstDia.Update
    
    'actualizamos el numero de registro contable de la operacion en la tabla caja y bancos
    xCon.Execute "UPDATE con_cajabanco SET con_cajabanco.numreg = '" & Trim(Format(xMes, "00")) + xNumAsiento & "' WHERE (((con_cajabanco.id)=" & xId & "))"

    xCon.CommitTrans
    
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDia = Nothing
    
    MsgBox "El movimiento se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
    
LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDia = Nothing
    
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Label5.Caption = "Agregando Operacion"
    Blanquea
    Bloquea
    OptCaja.Value = True
    OptCaja_Click
    Fg2.Rows = 1
    Fg3.Rows = 1
    Fg3.ColComboList(1) = "|..."
    
    Fg2.ColComboList(2) = "|..."
    Fg2.ColComboList(4) = "|..."
    
    Fg3.Rows = Fg3.Rows + 1
    Fg3.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    OptDe1.Value = True
    TxtFchMov.SetFocus
End Sub

Sub Bloquea()
    TxtFchMov.Locked = Not TxtFchMov.Locked
    TxtNumCue.Locked = Not TxtNumCue.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtImporte.Locked = Not TxtImporte.Locked
    TxtIdMedioPago.Locked = Not TxtIdMedioPago.Locked
    TxtIdDoc.Locked = Not TxtIdDoc.Locked
    
    CmdAddCon.Enabled = Not CmdAddCon.Enabled
    CmdDelCon.Enabled = Not CmdDelCon.Enabled
    OptDe1.Enabled = Not OptDe1.Enabled
    OptDe2.Enabled = Not OptDe2.Enabled
End Sub

Sub Blanquea()
    TxtFchMov.Valor = ""
    TxtIdMon.Text = ""
    TxtNumCue.Text = ""
    TxtIdMedioPago.Text = ""
    TxtProv.Text = ""
    TxtImporte.Text = ""

    TxtIdDoc.Text = ""
    TxtNumDoc.Text = ""
    TxtMedPag.Text = ""
    LblBanco.Caption = ""

    LblDesMedPag.Caption = ""
    LblMoneda.Caption = ""
    LblDescDoc.Caption = ""
    LblMedPag.Caption = ""
    
    TxtTotal1.Text = ""
    TxtTotal2.Text = ""
    TxtTotal3.Text = ""
    TxtTotal4.Text = ""
    
    Fg3.Rows = 1
    Fg2.Rows = 1
    
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

Private Sub TxtIdDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDoc_Click
    End If
End Sub

Private Sub TxtIdDoc_Validate(Cancel As Boolean)
    If TxtIdDoc.Text <> "" Then
        If OptCaja.Value = True Then
            xSQL = "SELECT mae_doccajaban.tipo, * From mae_doccajaban Where ((mae_doccajaban.tipo = 1) AND (id = " & Val(TxtIdDoc.Text) & ")) " _
                & " ORDER BY mae_doccajaban.descripcion"
        Else
            xSQL = "SELECT mae_doccajaban.tipo, * From mae_doccajaban Where ((mae_doccajaban.tipo = 2) AND (id = " & Val(TxtIdDoc.Text) & ")) " _
                & " ORDER BY mae_doccajaban.descripcion"
        End If
        
        Set Rst = BuscaConCriterio(xSQL, xCon)
        
        If Rst.RecordCount <> 0 Then
            LblDescDoc.Caption = Rst("descripcion")
            'If Rst("selecciona") = -1 Then
            '    TxtNumDoc.Locked = True
            'Else
            '    TxtNumDoc.Locked = False
            'End If
        End If
        Set Rst = Nothing
    End If
End Sub

Private Sub TxtIdMedioPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdMedioPago_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMedPag_Click
    End If
End Sub

Private Sub TxtIdMedioPago_Validate(Cancel As Boolean)
    If TxtIdMedioPago.Text <> "" Then
        xSQL = "SELECT con_origen.*, con_planctas.cuenta, con_planctas.descripcion AS desccuen, con_origen.idmon FROM con_planctas RIGHT JOIN con_origen " _
            & " ON con_planctas.id = con_origen.idcue WHERE (((con_origen.idmon)=1) AND ((con_origen.tipmov)=2) AND ((con_origen.id)=" & Val(TxtIdMedioPago.Text) & "))"

        'SELECT con_origen.*, con_planctas.cuenta, con_planctas.descripcion AS desccuen, con_origen.idmon FROM con_planctas RIGHT JOIN con_origen " _
            & " ON con_planctas.id = con_origen.idcue WHERE (((con_origen.idmon)=" & Val(TxtIdMon.Text) & ") AND ((con_origen.tipmov)=2))"

        Set Rst = BuscaConCriterio(xSQL, xCon)
        If Rst.RecordCount <> 0 Then
            LblDesMedPag.Caption = Rst("descripcion")
            LblCtaHaber.Caption = Rst("idcue") 'Rst("cuenta")
        
            'If NulosN(Rst("iddoc")) <> 0 Then
            '    TxtIdDoc.Text = Rst("iddoc")
            '    LblDescDoc.Caption = Busca_Codigo(Rst("iddoc"), "id", "descripcion", "mae_doccajaban", "N", xCon)
            'Else
            '    TxtIdDoc.Text = ""
            '    LblDescDoc.Caption = ""
            'End If
            'If NulosN(Rst("iddoc")) <> 0 Then
'                If Busca_Codigo(Rst("iddoc"), "id", "selecciona", "mae_doccajaban", "N", xCon) = -1 Then
'                    TxtNumDoc.Locked = True
'                Else
'                    TxtNumDoc.Locked = False
'                End If
            'Else
            '    TxtNumDoc.Locked = False
            'End If
        Else
            TxtIdMedioPago.Text = ""
            LblDesMedPag.Caption = ""
            LblCtaHaber.Caption = " "
        End If
        
        TxtNumDoc.Text = ""
        TxtImporte.Text = ""
        
        'TxtIdMov.SetFocus
        Set Rst = Nothing
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Sub HallarTotales()
    Dim A As Integer
    Dim Total1, Total2, Total3, Total4 As Double
    
    For A = 1 To Fg2.Rows - 1
        Total1 = Total1 + Val(Fg2.TextMatrix(A, 6))
        Total2 = Total2 + Val(Fg2.TextMatrix(A, 7))
        Total3 = Total3 + Val(Fg2.TextMatrix(A, 8))
        Total4 = Total4 + Val(Fg2.TextMatrix(A, 9))
    Next A
    
    TxtTotal1.Text = Format(Total1, "0.00")
    TxtTotal2.Text = Format(Total2, "0.00")
    TxtTotal3.Text = Format(Total3, "0.00")
    TxtTotal4.Text = Format(Total4, "0.00")
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If TxtIdMon.Text <> "" Then
        LblMoneda.Caption = Busca_Codigo(Val(TxtIdMon.Text), "id", "descripcion", "mae_moneda", "N", xCon)
        If LblMoneda.Caption = "" Then
            TxtIdMon.Text = ""
        End If
    End If
End Sub

'Private Sub TxtIdMov_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys vbTab
'    End If
'End Sub
'
'Private Sub TxtIdMov_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 116 Then
'        CmdBusMovi_Click
'    End If
'End Sub

'Private Sub TxtIdMov_Validate(Cancel As Boolean)
'   If TxtIdMov.Text <> "" And TxtIdMon.Text <> 0 Then
'        xSql = "SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, con_origen.id, con_origen.idmon, " _
'            & " con_origen.descripcion, con_origen.idcue, con_origen.tipmov, con_origen.id " _
'            & " FROM con_planctas INNER JOIN con_origen ON con_planctas.id = con_origen.idcue " _
'            & " WHERE (((con_origen.idmon)=" & Val(TxtIdMon.Text) & ") AND ((con_origen.tipmov)=2) AND ((con_origen.id)=" & Val(TxtIdMov.Text) & "))"
'
'        Set Rst = BuscaConCriterio(xSql, xCon)
'        If Rst.RecordCount <> 0 Then
'            LblDescMov.Caption = Rst("descripcion")
'            LblCtaDebe.Caption = Rst("idcue") 'Rst("cuenta")
'        Else
'            TxtIdMov.Text = ""
'            LblDescMov.Caption = ""
'            LblCtaDebe.Caption = " "
'        End If
'
'        If OptCaja.Value = True Then
'            TxtIdMedioPago.SetFocus
'        Else
'            TxtNumCue.SetFocus
'        End If
'        Set Rst = Nothing
'    End If
'End Sub

Private Sub TxtImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtImporte_Validate(Cancel As Boolean)
    If NulosN(TxtImporte.Text) <> 0 Then
        TxtImporte.Text = Format(TxtImporte.Text, "0.00")
    End If
End Sub

Private Sub TxtMedPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtMedPag_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdMP_Click
    End If
End Sub

Private Sub TxtMedPag_Validate(Cancel As Boolean)
    If TxtMedPag.Text <> "" Then
        LblMedPag.Caption = Busca_Codigo(TxtMedPag.Text, "id", "descripcion", "con_mediopago", "N", xCon)
        If LblMedPag.Caption = "" Then
            TxtMedPag.Text = ""
        End If
    End If
End Sub

Private Sub TxtNumCue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumCue_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCueBan_Click
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdNumDoc_Click
    End If
End Sub

Private Sub TxtProv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtProv_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        If Fg3.TextMatrix(Fg3.Row, 2) = "" Then Exit Sub
        CmdBusCliente_Click
    End If
End Sub


Sub CrearRstTmp()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(13, 3) As String

    xCampos(0, 0) = "nombre":       xCampos(0, 1) = "C":      xCampos(0, 2) = "5"
    xCampos(1, 0) = "tipdoc":       xCampos(1, 1) = "C":      xCampos(1, 2) = "15"
    xCampos(2, 0) = "fchemi":       xCampos(2, 1) = "F":      xCampos(2, 2) = "2"
    xCampos(3, 0) = "moneda":       xCampos(3, 1) = "F":      xCampos(3, 2) = "2"
    xCampos(4, 0) = "numdoc":       xCampos(4, 1) = "D":      xCampos(4, 2) = "2"
    xCampos(5, 0) = "impdoc":       xCampos(5, 1) = "D":      xCampos(5, 2) = "2"
    xCampos(6, 0) = "saldoc":       xCampos(6, 1) = "C":      xCampos(6, 2) = "10"
    xCampos(7, 0) = "impacu":       xCampos(7, 1) = "D":      xCampos(7, 2) = "2"
    xCampos(8, 0) = "newsal":       xCampos(8, 1) = "D":      xCampos(8, 2) = "2"
    xCampos(9, 0) = "iddoc":        xCampos(9, 1) = "N":      xCampos(9, 2) = "2"
    xCampos(10, 0) = "idpro":       xCampos(10, 1) = "N":     xCampos(10, 2) = "2"
    xCampos(11, 0) = "idtipdoc":    xCampos(11, 1) = "N":      xCampos(11, 2) = "2"
    xCampos(12, 0) = "idmon":       xCampos(12, 1) = "N":      xCampos(12, 2) = "2"
    
    Set RstTMPDoc = xFun.CrearRstTmp(xCampos)

    RstTMPDoc.Open
End Sub
