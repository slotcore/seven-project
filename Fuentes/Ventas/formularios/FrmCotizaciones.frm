VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCotizaciones 
   Caption         =   "Ventas - Cotizaciones"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCotizaciones.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCotizaciones.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCotizaciones.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCotizaciones.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCotizaciones.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCotizaciones.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCotizaciones.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCotizaciones.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCotizaciones.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCotizaciones.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCotizaciones.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCotizaciones.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCotizaciones.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   -15
      TabIndex        =   12
      Top             =   375
      Width           =   11895
      _cx             =   20981
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
         Height          =   6810
         Left            =   45
         TabIndex        =   40
         Top             =   375
         Width           =   11805
         Begin VB.Frame Frame5 
            Height          =   1080
            Left            =   7650
            TabIndex        =   63
            Top             =   2175
            Width           =   4035
            Begin VB.CommandButton CmdAprobada 
               Caption         =   "Aprobar"
               Height          =   315
               Left            =   135
               TabIndex        =   65
               Top             =   675
               Width           =   1860
            End
            Begin VB.CommandButton CmdRecha 
               Caption         =   "Rechazar"
               Height          =   315
               Left            =   2025
               TabIndex        =   64
               Top             =   675
               Width           =   1860
            End
            Begin VB.Label LblEstado 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Pendiente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   465
               Left            =   120
               TabIndex        =   66
               Top             =   195
               Width           =   3765
            End
         End
         Begin VB.TextBox TxtReferencias 
            Height          =   300
            Left            =   1755
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "TxtReferencias"
            Top             =   2955
            Width           =   4740
         End
         Begin VB.CommandButton CmdBusCondicion 
            Height          =   240
            Left            =   2400
            Picture         =   "FrmCotizaciones.frx":2A98
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   780
            Width           =   240
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2670
            Left            =   150
            TabIndex        =   11
            Top             =   3300
            Width           =   9705
            _cx             =   17119
            _cy             =   4710
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
            Rows            =   15
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCotizaciones.frx":2BCA
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
         Begin VB.CommandButton CmdBusAutoriza 
            Height          =   240
            Left            =   6225
            Picture         =   "FrmCotizaciones.frx":2D09
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   2670
            Width           =   240
         End
         Begin VB.TextBox TxtAutoriza 
            Height          =   300
            Left            =   1755
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   9
            Text            =   "TxtAutoriza"
            Top             =   2640
            Width           =   4740
         End
         Begin VB.Frame Frame3 
            Height          =   2790
            Left            =   9915
            TabIndex        =   33
            Top             =   3195
            Width           =   1770
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "Agregar Item"
               Enabled         =   0   'False
               Height          =   600
               Left            =   240
               TabIndex        =   34
               Top             =   780
               Width           =   1260
            End
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "Eliminar Item"
               Enabled         =   0   'False
               Height          =   600
               Left            =   240
               TabIndex        =   35
               Top             =   1455
               Width           =   1260
            End
         End
         Begin VB.CommandButton CmdBusContacto 
            Height          =   240
            Left            =   6225
            Picture         =   "FrmCotizaciones.frx":2E3B
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   2040
            Width           =   240
         End
         Begin VB.TextBox TxtContacto 
            Height          =   300
            Left            =   1755
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   7
            Text            =   "TxtContacto"
            Top             =   2010
            Width           =   4740
         End
         Begin VB.CommandButton CmdBusVen 
            Height          =   240
            Left            =   2400
            Picture         =   "FrmCotizaciones.frx":2F6D
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2355
            Width           =   240
         End
         Begin VB.TextBox TxtIdVen 
            Height          =   300
            Left            =   1755
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   8
            Text            =   "TxtIdVen"
            Top             =   2325
            Width           =   915
         End
         Begin VB.CommandButton CmdBusTipItem 
            Height          =   240
            Left            =   2400
            Picture         =   "FrmCotizaciones.frx":309F
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   465
            Width           =   240
         End
         Begin VB.CommandButton CmdBusCli 
            Height          =   240
            Left            =   3270
            Picture         =   "FrmCotizaciones.frx":31D1
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1725
            Width           =   240
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   8085
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   1
            Text            =   "TxtNumDoc"
            Top             =   435
            Width           =   1770
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   2400
            Picture         =   "FrmCotizaciones.frx":3303
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1410
            Width           =   240
         End
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1755
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   6
            Text            =   "TxtNumRuc"
            Top             =   1695
            Width           =   1770
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   1755
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   5
            Text            =   "TxtIdMon"
            Top             =   1380
            Width           =   915
         End
         Begin VB.TextBox TxtTipItem 
            Height          =   300
            Left            =   1755
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   0
            Text            =   "TxtTipItem"
            Top             =   435
            Width           =   915
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchDoc 
            Height          =   300
            Left            =   1755
            TabIndex        =   3
            Top             =   1065
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
            Left            =   4245
            TabIndex        =   4
            Top             =   1065
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
         Begin VB.Frame Frame4 
            Height          =   810
            Left            =   150
            TabIndex        =   41
            Top             =   5970
            Width           =   11535
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
               Left            =   7080
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   46
               TabStop         =   0   'False
               Text            =   "TxtIsc"
               Top             =   420
               Width           =   1200
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
               Left            =   4320
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   45
               TabStop         =   0   'False
               Text            =   "TxtInafecto"
               Top             =   420
               Width           =   1200
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
               Left            =   2940
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   44
               TabStop         =   0   'False
               Text            =   "TxtBruto"
               Top             =   420
               Width           =   1200
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
               Left            =   5700
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   43
               TabStop         =   0   'False
               Text            =   "TxtIGV"
               Top             =   420
               Width           =   1200
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
               Left            =   8475
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   42
               TabStop         =   0   'False
               Text            =   "TxtTotal"
               Top             =   420
               Width           =   1200
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
               Left            =   7065
               TabIndex        =   52
               Top             =   195
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
               Left            =   4320
               TabIndex        =   51
               Top             =   195
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
               Left            =   6360
               TabIndex        =   50
               Top             =   195
               Width           =   570
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
               Left            =   2940
               TabIndex        =   49
               Top             =   195
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
               Left            =   5700
               TabIndex        =   48
               Top             =   195
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
               Left            =   8445
               TabIndex        =   47
               Top             =   195
               Width           =   450
            End
         End
         Begin VB.TextBox TxtConPag 
            Height          =   300
            Left            =   1755
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   2
            Text            =   "TxtConPag"
            Top             =   750
            Width           =   915
         End
         Begin VB.Label LblIdEstado 
            AutoSize        =   -1  'True
            Caption         =   "LblIdEstado"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   6570
            TabIndex        =   67
            Top             =   2295
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Condicion de Pago"
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   62
            Top             =   810
            Width           =   1350
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
            Left            =   2715
            TabIndex        =   61
            Top             =   750
            Width           =   2820
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emision"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   59
            Top             =   1110
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Venc."
            Height          =   195
            Index           =   3
            Left            =   3375
            TabIndex        =   58
            Top             =   1110
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Autorizante"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   57
            Top             =   2685
            Width           =   795
         End
         Begin VB.Label LblIdAutoriza 
            AutoSize        =   -1  'True
            Caption         =   "LblIdAutoriza"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   6585
            TabIndex        =   56
            Top             =   2535
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Contacto"
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   25
            Top             =   2040
            Width           =   645
         End
         Begin VB.Label LblIdContacto 
            AutoSize        =   -1  'True
            Caption         =   "LblIdContacto"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   6540
            TabIndex        =   27
            Top             =   2055
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label lblreferencias 
            AutoSize        =   -1  'True
            Caption         =   "Referenc."
            Height          =   195
            Left            =   180
            TabIndex        =   54
            Top             =   3000
            Width           =   705
         End
         Begin VB.Label LblNomVen 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomVen"
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
            Left            =   2715
            TabIndex        =   32
            Top             =   2325
            Width           =   3765
         End
         Begin VB.Label Lblvendedor 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   180
            TabIndex        =   30
            Top             =   2370
            Width           =   690
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Cotizaciones"
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
            TabIndex        =   53
            Top             =   30
            Width           =   11610
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
            Left            =   2715
            TabIndex        =   16
            Top             =   435
            Width           =   2820
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Item"
            Height          =   195
            Index           =   6
            Left            =   180
            TabIndex        =   14
            Top             =   480
            Width           =   660
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "T. Cambio"
            Height          =   195
            Left            =   7125
            TabIndex        =   28
            Top             =   1455
            Visible         =   0   'False
            Width           =   720
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
            Left            =   8085
            TabIndex        =   29
            Top             =   1380
            Visible         =   0   'False
            Width           =   2160
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
            Left            =   2715
            TabIndex        =   20
            Top             =   1380
            Width           =   2820
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   21
            Top             =   1725
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
            TabIndex        =   23
            Top             =   1695
            Width           =   6660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Cotización"
            Height          =   195
            Index           =   0
            Left            =   6660
            TabIndex        =   17
            Top             =   480
            Width           =   1185
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   18
            Top             =   1410
            Width           =   585
         End
         Begin VB.Label LblIdCliente 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCliente"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   10275
            TabIndex        =   24
            Top             =   1740
            Visible         =   0   'False
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   -12450
         TabIndex        =   37
         Top             =   375
         Width           =   11805
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6450
            Left            =   30
            TabIndex        =   13
            Top             =   345
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11377
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
            Columns(1).Caption=   "Nº Documento"
            Columns(1).DataField=   "numerodoc"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Emi"
            Columns(2).DataField=   "fchdoc"
            Columns(2).NumberFormat=   "Short Date"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Ven."
            Columns(3).DataField=   "fchven"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Cliente"
            Columns(4).DataField=   "nombre"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Moneda"
            Columns(5).DataField=   "simbolo"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Importe"
            Columns(6).DataField=   "imptotdoc"
            Columns(6).NumberFormat=   "0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Condicion"
            Columns(7).DataField=   "Condicion"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2566"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2487"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1746"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1667"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1773"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1693"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=6403"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=6324"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1402"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1323"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1640"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1561"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=514"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=2170"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2090"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
            _StyleDefs(68)  =   "Named:id=33:Normal"
            _StyleDefs(69)  =   ":id=33,.parent=0"
            _StyleDefs(70)  =   "Named:id=34:Heading"
            _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(72)  =   ":id=34,.wraptext=-1"
            _StyleDefs(73)  =   "Named:id=35:Footing"
            _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(75)  =   "Named:id=36:Selected"
            _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=37:Caption"
            _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(79)  =   "Named:id=38:HighlightRow"
            _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=39:EvenRow"
            _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(83)  =   "Named:id=40:OddRow"
            _StyleDefs(84)  =   ":id=40,.parent=33"
            _StyleDefs(85)  =   "Named:id=41:RecordSelector"
            _StyleDefs(86)  =   ":id=41,.parent=34"
            _StyleDefs(87)  =   "Named:id=42:FilterBar"
            _StyleDefs(88)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta Cotizaciones"
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
            Left            =   90
            TabIndex        =   39
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
            TabIndex        =   38
            Top             =   30
            Width           =   765
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   12540
         X2              =   24345
         Y1              =   375
         Y2              =   7185
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
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
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Guia"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
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
Attribute VB_Name = "FrmCotizaciones"
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
Dim NumDias As Integer      'para saber a cuantos dias se esta emitiendo el documento
Dim xIdCuenTasa As Integer  'codigo de la cuenta contable del impuesto
Dim xMes As Integer         'numero de mes en el que se realiza la operacion
Dim Mostrando As Boolean
Dim swguiafact '0 No se facturaron, 1 Se facturaron
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO

Sub EstadoCotizacion(numindex As Integer)
    Dim Rpta As Integer
    '1 Pendiente
    '2 Aprobada
    '3 Procesada
    '4 Rechazada

    If numindex = 1 Then
        Rpta = MsgBox("¿Desea actualizar a [Pendiente] la cotizacion ?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
        If Rpta = vbYes Then
            xCon.Execute " UPDATE vta_cotizacion SET vta_cotizacion.idest = 1, " _
                & " vta_cotizacion.idaut = 0 WHERE vta_cotizacion.id = " & RstVent("id") & " "
        End If
    End If

    If numindex = 2 Then
        LblEstado.ForeColor = &H8000&
        LblEstado.Caption = "Aprobado"
        xCon.Execute " UPDATE vta_cotizacion SET vta_cotizacion.idest = 2 WHERE vta_cotizacion.id = " & RstVent("id") & " "
    End If

    If numindex = 4 Then
        LblEstado.Caption = "Rechazado"
        LblEstado.ForeColor = &HFF&
        xCon.Execute " UPDATE vta_cotizacion SET vta_cotizacion.idest = 4, vta_cotizacion.idaut = 0 WHERE vta_cotizacion.id = " & RstVent("id") & " "
    End If
    
    MsgBox "Cotización se actualizado con exito ", vbInformation, xTitulo
    RstVent.Requery
    Dg1.Refresh
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    
    If RstVent.RecordCount = 0 Then
        MsgBox "No hay cotizaciones para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If RstVent("idest") = 3 Then
        MsgBox "No se puede eliminar una cotizacion que ha sido procesada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    TabOne1.CurrTab = 0
    Rpta = MsgBox("¿ Esta seguro de eliminar la cotizacion seleccionada ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM  vta_cotizacion WHERE id =" & RstVent("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstVent("id") & " AND idform = " & IdMenuActivo

        
        MsgBox "La cotizacion se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstVent.Requery
        Dg1.Refresh
    
        If RstVent.RecordCount = 0 Then
            Rpta = MsgBox("No se ha registrado ninguna cotizacion, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstVent = Nothing
                Unload Me
            End If
        End If
    End If
End Sub

Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    Toolbar1.Buttons(4).Enabled = Not Toolbar1.Buttons(4).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
    Toolbar1.Buttons(14).Enabled = Not Toolbar1.Buttons(14).Enabled
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
    Dim rs As New ADODB.Recordset

    QueHace = 1
    xHorIni = Time
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Label5.Caption = "Agregando Cotización"
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    TxtFchDoc.Valor = Date
    Fg1.Rows = 1
    
    RST_Busq rs, "SELECT vta_cotizacion.* FROM vta_cotizacion ORDER BY vta_cotizacion.NUMDOC ", xCon
   
    If rs.RecordCount > 0 Then
        rs.MoveLast
        TxtNumDoc = Format(Val(rs!NumDoc) + 1, "0000000000")
    Else
        TxtNumDoc = "0000000001"
    End If
    LblIdEstado.Caption = "1"
    LblEstado.Caption = "Pendiente"
    LblEstado.ForeColor = &HC0FFFF
    Set rs = Nothing
    TxtTipItem.SetFocus
End Sub

Sub Modificar()
    QueHace = 2
    xHorIni = Time
    TabOne1.CurrTab = 1
    Blanquea
    Bloquea
    TabOne1.TabEnabled(0) = False
    ActivaTool
    MuestraSegundoTab
    Label5.Caption = "Modificando Cotizaciones"
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    TxtTipItem.SetFocus
End Sub

Sub MuestraSegundoTab()
    TxtTipItem.Text = NulosN(RstVent("idtipo"))
    TxtNumRuc.Text = RstVent("numruc")
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
    
    LblTipoItem.Caption = RstVent("desctipcom")

    TxtReferencias = NulosC(RstVent("referencia"))
    LblNomCli.Caption = RstVent("nombre")
    LblCondPag.Caption = NulosC(RstVent("desccond"))
    TxtNumRuc.Text = RstVent("numruc")
    LblMoneda.Caption = RstVent("descmon")
    LblIdCliente.Caption = RstVent("idcli")
    
    If RstVent("idmon") = 1 Then
        LblTipoCambio.Visible = False
    Else
        LblTipoCambio.Visible = True
        LblTipoCambio.Caption = RstVent("impven")
    End If
    
    LblIdContacto = RstVent("idcon")
    
    LblIdEstado.Caption = RstVent("idest")
    If LblIdEstado.Caption = "1" Then LblEstado.Caption = RstVent("condicion"): LblEstado.ForeColor = &HC0FFFF
    If LblIdEstado.Caption = "2" Then LblEstado.Caption = RstVent("condicion"): LblEstado.ForeColor = &H8000&
    If LblIdEstado.Caption = "3" Then LblEstado.Caption = RstVent("condicion"): LblEstado.ForeColor = &HFF0000
    If LblIdEstado.Caption = "4" Then LblEstado.Caption = RstVent("condicion"): LblEstado.ForeColor = &HFF&
    
    Dim RstTmp As New ADODB.Recordset
    Set RstTmp = BuscaConCriterio("SELECT * FROM mae_clicontacto WHERE idcli = " & Val(LblIdCliente.Caption) & "", xCon)
    
    If RstTmp.RecordCount <> 0 Then
        TxtContacto.Text = UCase(Trim(RstTmp("nomcon"))) + ", " + Trim(RstTmp("apecon"))
    Else
        TxtContacto.Text = ""
    End If
    Set RstTmp = Nothing

    LblIdAutoriza = NulosN(RstVent("idaut"))
    TxtAutoriza = NulosC(RstVent("apenomaut"))
    
    'Obtenemos el nombre del vendedor
    'Set RstTmp = BuscaConCriterio("SELECT   pla_empleados.apellnom " _
        & " FROM pla_empleados INNER JOIN vta_vendedores ON pla_empleados.id = vta_vendedores.idpers " _
        & " WHERE vta_vendedores.id = " & rstvent("idven") & "", xCon)
            
    'If RstTmp.RecordCount > 0 Then
    TxtIdVen = NulosN(RstVent("idven"))
    LblNomVen = NulosC(RstVent("apenomven"))
    'End If
    
    Set RstTmp = Nothing
 
    Dim RstDet As New ADODB.Recordset
    
    Mostrando = True
    Fg1.Rows = 1
    
    RST_Busq RstDet, "SELECT vta_cotizaciondet.*, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idcuentaven,alm_inventario.idtipven " _
        & " FROM mae_unidades RIGHT JOIN (alm_inventario INNER JOIN vta_cotizaciondet ON alm_inventario.id = vta_cotizaciondet.iditem) " _
        & " ON mae_unidades.id = alm_inventario.idunimed WHERE (((vta_cotizaciondet.idvta)=" & RstVent("id") & "))", xCon
    
    If RstDet.RecordCount <> 0 Then
        Do While Not RstDet.EOF
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = RstDet("descripcion")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = RstDet("abrev")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(RstDet("preuni"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(RstDet("canpro"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(RstDet("imptot"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = RstDet("iditem")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = RstDet("idunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = RstDet("idcuentaven")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = RstDet("idtipven")
            RstDet.MoveNext
        Loop
    End If
    
    Set RstDet = Nothing
    Set RstDet = BuscaConCriterio("SELECT mae_impuestos.tasa from mae_impuestos WHERE mae_impuestos.id = 1 ", xCon)
    
    If RstDet.RecordCount = 1 Then
        TasaImpuesto = NulosN(RstDet("tasa"))
        LblIgvTasa.Caption = Trim(Str(TasaImpuesto)) + "%"
    End If
    
    Mostrando = False
End Sub

Sub Bloquea()
    CmdAprobada.Enabled = Not CmdAprobada.Enabled
    CmdRecha.Enabled = Not CmdRecha.Enabled
    
    TxtTipItem.Locked = Not TxtTipItem.Locked
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtFchDoc.Locked = Not TxtFchDoc.Locked
    TxtFchVen.Locked = Not TxtFchVen.Locked
    TxtConPag.Locked = Not TxtConPag.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtContacto.Locked = Not TxtContacto.Locked
    TxtIdVen.Locked = Not TxtIdVen.Locked
    TxtReferencias.Locked = Not TxtReferencias.Locked
    
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
    
End Sub

Sub Blanquea()
    TxtTipItem.Text = ""
    TxtNumRuc.Text = ""
    TxtContacto.Text = ""
    TxtNumDoc.Text = ""
    TxtFchDoc.Valor = ""
    TxtFchVen.Valor = ""
    TxtConPag.Text = ""
    TxtIdMon.Text = ""
    
    LblNomCli.Caption = ""
    LblCondPag.Caption = ""
    LblMoneda.Caption = ""
    LblIdCliente.Caption = ""
    LblTipoItem.Caption = ""
    TxtIdVen = ""
    LblNomVen = ""
    
    txtinafecto = ""
    txtisc = ""
    TxtBruto.Text = ""
    TxtIGV.Text = ""
    TxtTotal.Text = ""
    TxtReferencias = ""
    TxtAutoriza = ""
    
    Fg1.Rows = 1
End Sub

Private Sub CmdAddItem_Click()
    If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 1)) = "" Then Exit Sub
    Fg1.Rows = Fg1.Rows + 1
    
    With Fg1
        .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, 1
    End With
    
    Fg1_CellButtonClick Fg1.Rows - 1, 1
    Fg1.SetFocus
End Sub

Private Sub CmdAprobada_Click()
    If Val(LblIdEstado.Caption) = 3 Then
        MsgBox "No se puede aprobar una cotizacion procesada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    EstadoCotizacion 2
End Sub

Private Sub CmdBusAutoriza_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Apellido Nombre":    xCampos(0, 1) = "apenom":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":             xCampos(1, 1) = "id":          xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT pla_empleados.id, [pla_empleados]![nombre] AS apenom" _
        & " FROM pla_empleados"
    
    xform.Titulo = "Buscando Personal"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "apenom"
    xform.CampoBusca = "apenom"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtAutoriza.Text = xRs("apenom")
        LblIdAutoriza.Caption = xRs("id")
        TxtReferencias.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
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
        NumDias = xRs("numdia")
        
        If NulosC(TxtFchDoc.Valor) = "" Then
            TxtFchDoc.Valor = Format(Date, "dd/mm/yyyy")
            TxtFchVen.Valor = Date + xRs("numdia")
        Else
            TxtFchVen.Valor = Date + xRs("numdia")
        End If
        TxtFchDoc.SetFocus
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
        TxtContacto.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusContacto_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Apellidos y Nombres":   xCampos(0, 1) = "apenom":    xCampos(0, 2) = "3500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Telefono":              xCampos(1, 1) = "numcel":    xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Email":                 xCampos(2, 1) = "email":     xCampos(2, 2) = "2000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Codigo":                xCampos(3, 1) = "id":        xCampos(3, 2) = "1000":         xCampos(3, 3) = "N"
    
    xform.SQLCad = "SELECT mae_clicontacto.id, UCase([mae_clicontacto]![apecon])+', '+[mae_clicontacto]![nomcon] AS apenom, " _
        & " mae_clicontacto.numcel, mae_clicontacto.email From mae_clicontacto WHERE (((mae_clicontacto.idcli)=" & Val(LblIdCliente.Caption) & "))"

    xform.Titulo = "Buscando Contactos del Cliente"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "apenom"
    xform.CampoBusca = "apenom"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        TxtContacto.Text = xRs("apenom")
        LblIdContacto.Caption = xRs("id")
        TxtIdVen.SetFocus
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
        TxtConPag.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusVen_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Codigo":    xCampos(0, 1) = "id":         xCampos(0, 2) = "800":          xCampos(0, 3) = "N"
    xCampos(1, 0) = "Vendedor":  xCampos(1, 1) = "apenom":     xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Basico":    xCampos(2, 1) = "basico":     xCampos(2, 2) = "1200":         xCampos(2, 3) = "N"
    xCampos(3, 0) = "Comision":  xCampos(3, 1) = "comision":   xCampos(3, 2) = "1200":         xCampos(3, 3) = "N"
    
    xform.SQLCad = "SELECT vta_vendedores.*, pla_empleados!nombre AS apenom " _
                & " FROM vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id"
    
    xform.Titulo = "Buscando Vendedores"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "apenom"
    xform.CampoBusca = "apenom"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        LblNomVen.Caption = xRs("apenom")
        TxtIdVen.Text = xRs("id")
        TxtAutoriza.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDelItem_Click()
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

Private Sub CmdRecha_Click()
    If Val(LblIdEstado.Caption) = 3 Then
        MsgBox "No se puede rechazar una cotizaicon procesada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    Else
        EstadoCotizacion 4
    End If
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstVent("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Unidad":       xCampos(1, 1) = "abrev":          xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Codigo":       xCampos(2, 1) = "codpro":         xCampos(2, 2) = "1500":    xCampos(2, 3) = "C"
    
    xform.SQLCad = " SELECT alm_inventario.*, mae_unidades.descripcion AS descuni, mae_unidades.abrev, mae_percepcion.tasa " _
        & "FROM mae_unidades INNER JOIN (mae_percepcion RIGHT JOIN alm_inventario ON mae_percepcion.id = alm_inventario.idper) " _
        & " ON mae_unidades.id = alm_inventario.idunimed WHERE alm_inventario.tippro = " & Val(TxtTipItem) & " " _
        & " ORDER BY alm_inventario.descripcion"

    xform.Titulo = "Buscando Productos"
    xform.FormaBusca = CualquierParte
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        Fg1.TextMatrix(Fg1.Row, 1) = xRs("descripcion")
        Fg1.TextMatrix(Fg1.Row, 2) = xRs("abrev")
        Fg1.TextMatrix(Fg1.Row, 3) = Format(NulosN(xRs("preuni")), "0.00")
        Fg1.TextMatrix(Fg1.Row, 6) = xRs("id")
        Fg1.TextMatrix(Fg1.Row, 7) = xRs("idunimed")
        Fg1.TextMatrix(Fg1.Row, 8) = xRs("idcuentaven")
        Fg1.TextMatrix(Fg1.Row, 9) = xRs("idtipven")
        Fg1.TextMatrix(Fg1.Row, 10) = NulosN(xRs("tasa"))
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
        If Fg1.TextMatrix(A, 9) = "1" Then 'si es venta gravada
            totalafec = totalafec + Val(Fg1.TextMatrix(A, 5)) 'venta  gravada
        Else
            totalinaf = totalinaf + Val(Fg1.TextMatrix(A, 5)) 'venta no gravada
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

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then Exit Sub
    
    If Fg1.Col = 2 Or Fg1.Col = 5 Then
        Fg1.Editable = flexEDNone
    Else
        Fg1.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = 46 Then CmdDelItem_Click
    If KeyCode = 45 Then CmdAddItem_Click
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    If Button = 2 Then PopupMenu menu1
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon

        RST_Busq RstVent, "SELECT vta_cotizacion.*, mae_cliente.nombre, [vta_cotizacion]![numdoc] AS numerodoc, mae_condpago.descripcion AS desccond, " _
            & " mae_cliente.numruc, mae_moneda.descripcion AS descmon, mae_moneda.simbolo, con_tc.impcom, mae_tipoproducto.descripcion AS desctipcom, " _
            & " mae_estadoordcom.descripcion AS Condicion,  [pla_empleados]![nombre] AS apenomven, [pla_empleados_1]![nombre] AS apenomaut " _
            & " FROM ((vta_vendedores RIGHT JOIN (mae_estadoordcom RIGHT JOIN ((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_condpago " _
            & " RIGHT JOIN (vta_cotizacion LEFT JOIN con_tc ON vta_cotizacion.fchdoc = con_tc.fecha) ON mae_condpago.id = vta_cotizacion.idconpag) " _
            & " ON mae_moneda.id = vta_cotizacion.idmon) ON mae_cliente.id = vta_cotizacion.idcli) LEFT JOIN mae_tipoproducto " _
            & " ON vta_cotizacion.idtipo = mae_tipoproducto.id) ON mae_estadoordcom.id = vta_cotizacion.idest) ON vta_vendedores.id = vta_cotizacion.idven) " _
            & " LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id) LEFT JOIN pla_empleados AS pla_empleados_1 " _
            & " ON vta_cotizacion.idaut = pla_empleados_1.id ORDER BY vta_cotizacion.fchdoc DESC", xCon

        Set Dg1.DataSource = RstVent
    
    End If
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    QueHace = 3
    TabOne1.CurrTab = 0
    SeEjecuto = False
    
    CaracteresNumericos = "0123456789." & Chr(8)
    Fg1.ColWidth(6) = 0
    Fg1.ColWidth(7) = 0
    Fg1.ColWidth(8) = 0
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    
    Set rs = BuscaConCriterio("SELECT mae_impuestos.tasa from mae_impuestos WHERE mae_impuestos.id = 1 ", xCon)
    If rs.RecordCount = 1 Then
        TasaImpuesto = NulosN(rs("tasa"))
        LblIgvTasa.Caption = Trim(Str(TasaImpuesto)) + "%"
    End If
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ColComboList(1) = ""
    
    TxtFchDoc.Valor = Date
    TxtFchVen.Valor = Date
    TxtFchVen.Valor = ""
    TxtFchVen.Valor = ""
    swguiafact = 0
    xAño = 2006
    xMes = 12
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando una cotizacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then
        If RstVent.RecordCount = 0 Then
            MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
            Exit Sub
        End If
        If RstVent("idest") = 3 Then
            MsgBox "No se puede modificar una cotizacion que ha sido procesada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.Index = 3 Then
        Eliminar
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstVent.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 14 Then
        Set RstVent = Nothing
        Unload Me
    End If
End Sub

Private Sub TxtAutoriza_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtReferencias.SetFocus
    End If
End Sub

Private Sub TxtAutoriza_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusAutoriza_Click
    End If
End Sub

Private Sub TxtConPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If NulosC(TxtConPag.Text) = "" Then
            SendKeys vbTab
            Exit Sub
        End If
        Dim xRs1 As New ADODB.Recordset
        
        RST_Busq xRs1, "SELECT * FROM mae_condpago WHERE id = " & Val(TxtConPag.Text) & "", xCon
        
        If xRs1.RecordCount = 0 Then
            TxtConPag.Text = ""
            LblCondPag.Caption = ""
        Else
            LblCondPag.Caption = Trim(xRs1("descripcion"))
            NumDias = xRs1("numdia")
            
            If NulosC(TxtFchDoc.Valor) = "" Then
                TxtFchDoc.Valor = Format(Date, "dd/mm/yyyy")
                TxtFchVen.Valor = Date + NumDias
            Else
                TxtFchVen.Valor = CDate(TxtFchDoc.Valor) + NumDias
            End If
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

Private Sub TxtContacto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtIdVen.SetFocus
    End If
End Sub

Private Sub TxtContacto_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusContacto_Click
    End If
End Sub

Private Sub TxtFchDoc_Validate(Cancel As Boolean)
    If TxtFchDoc.Valor <> "" Then
        TxtFchVen.Valor = CDate(TxtFchDoc.Valor) + NumDias
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If NulosC(TxtIdMon.Text) = "" Then
            SendKeys vbTab
            Exit Sub
        End If
        Dim xRs1 As New ADODB.Recordset
        
        'buscamos el codigo de la moneda  digitada
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
        TxtNumRuc.SetFocus
        Set xRs1 = Nothing
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdVen_KeyPress(KeyAscii As Integer)
    Dim RstTmp As New ADODB.Recordset
    
    If KeyAscii = 13 Then
        If NulosC(TxtIdVen.Text) <> "" Then
            Set RstTmp = BuscaConCriterio("SELECT vta_vendedores.*, pla_empleados!nombre AS apenom " _
                & " FROM vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id WHERE vta_vendedores.id = " & Val(TxtIdVen.Text) & "", xCon)
            
            If RstTmp.RecordCount <> 0 Then
                LblNomVen.Caption = NulosC(RstTmp("apenom"))
            Else
                TxtIdVen.Text = ""
                LblNomVen.Caption = ""
            End If
        End If
        TxtAutoriza.SetFocus
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    Set RstTmp = Nothing
End Sub

Private Sub TxtIdVen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusVen_Click
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
        
        TxtContacto.SetFocus
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

Function Grabar() As Boolean
    If TxtTipItem.Text = "" Then
        MsgBox "No ha especificado el tipo de item", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipItem.SetFocus
        Exit Function
    End If
    
    If TxtNumRuc.Text = "" Then
        MsgBox "No ha especificado cliente de la venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
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
    
    If CDate(TxtFchVen.Valor) < CVDate(TxtFchDoc.Valor) Then
        MsgBox " La fecha de vencimiento del documento no puede ser menor a la fecha de emision", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
    
    If QueHace = 1 Then 'Validamos si existe el numero del documento en modo adicion
        Dim RstCab As New ADODB.Recordset
        
        RST_Busq RstCab, " Select * from  vta_cotizacion where numdoc = '" & TxtNumDoc & "' ", xCon
        
        If RstCab.RecordCount > 0 Then
            MsgBox "El Nro de documento ha sido registrado por otro usuario se grabara con otro numero", vbInformation, Me.Caption
        End If
        
        Set RstCab = Nothing
    End If
    
    Dim RstDet As New ADODB.Recordset
    
    Dim xIdCuen As Integer
    Dim xTotal As Double
    
    Dim xidtipven As String 'Determina si la venta es de tipo exportacion
    Dim xNumAsiento As String
    
    Dim xId As Double
    Dim A As Integer
    Dim X As Integer
    Dim P As Integer
    
    'On Error GoTo LaCague
    swguiafact = 1
    xCon.BeginTrans
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("vta_cotizacion", xCon, "id")
        xNumAsiento = HallaNumAsiento(xMes)
        
        RST_Busq RstCab, "SELECT * FROM vta_cotizacion", xCon
        RST_Busq RstDet, "SELECT * FROM vta_cotizaciondet", xCon
                
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstVent("id")
        RST_Busq RstCab, "SELECT * FROM vta_cotizacion WHERE id = " & xId & "", xCon
        
        'Eliminamos el detalle de la cotizacion
         xCon.Execute "DELETE * FROM vta_cotizaciondet WHERE idvta = " & xId & ""
         RST_Busq RstDet, "SELECT * FROM vta_cotizaciondet", xCon
    End If
    
    RstCab("idtipo") = Val(TxtTipItem.Text)
    RstCab("idcli") = NulosN(LblIdCliente.Caption)
    RstCab("numdoc") = TxtNumDoc.Text
    RstCab("fchdoc") = TxtFchDoc.Valor
    RstCab("fchven") = TxtFchVen.Valor
    RstCab("idconpag") = NulosN(TxtConPag.Text)
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    RstCab("impbru") = NulosN(TxtBruto.Text)
    RstCab("impinaf") = NulosN(txtinafecto.Text)
    RstCab("impigv") = NulosN(TxtIGV.Text)
    RstCab("impisc") = NulosN(txtisc.Text)
    RstCab("impotr") = 0 'NulosN(me.txtir Txtotr...Text)
    RstCab("imptotdoc") = NulosN(TxtTotal.Text)
    RstCab("anulado") = 0
    RstCab("idest") = 1
    RstCab("idaut") = Val(LblIdAutoriza)
    RstCab("idcon") = Val(LblIdContacto.Caption)
    RstCab("idven") = NulosN(TxtIdVen)
    RstCab.Update
    
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idvta") = xId
        RstDet("iditem") = Val(Fg1.TextMatrix(A, 6))
        RstDet("idunimed") = Val(Fg1.TextMatrix(A, 7))
        RstDet("preuni") = Val(Fg1.TextMatrix(A, 3))
        RstDet("canpro") = Val(Fg1.TextMatrix(A, 4))
        RstDet("imptot") = Val(Fg1.TextMatrix(A, 5))
        RstDet("tasaper") = Val(Fg1.TextMatrix(A, 10))
        RstDet.Update
    Next A
   
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
   
   
    xCon.CommitTrans
    MsgBox "Cotización se registro con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Set RstCab = Nothing
    Set RstDet = Nothing
    
    Grabar = True
    Exit Function
    
'LaCague:
    'Resume
    xCon.RollbackTrans
    
    Set RstCab = Nothing
    Set RstDet = Nothing
    
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

Private Sub TxtReferencias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdAddItem.Enabled = False Then
            Fg1.SetFocus
        Else
            CmdAddItem.SetFocus
        End If
    End If
End Sub

Private Sub TxtTipItem_KeyPress(KeyAscii As Integer)
    Dim RstTmp As New ADODB.Recordset
    If KeyAscii = 13 Then
        If NulosC(TxtTipItem.Text) <> "" Then
            Set RstTmp = BuscaConCriterio("SELECT * FROM mae_tipoproducto WHERE id = " & Val(TxtTipItem.Text) & "", xCon)
            If RstTmp.RecordCount <> 0 Then
                LblTipoItem.Caption = RstTmp("descripcion")
            Else
                TxtTipItem.Text = ""
                LblTipoItem.Caption = ""
            End If
        End If
        TxtConPag.SetFocus
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

