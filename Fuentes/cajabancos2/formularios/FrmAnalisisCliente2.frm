VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "Codejock.CommandBars.v12.0.0.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.ocx"
Begin VB.Form FrmAnalisisCliente2 
   Caption         =   "Analisis de Cta Cte del Cliente  por Documento de Referencia"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   14370
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   885
      Left            =   840
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   6180
      Begin XtremeSuiteControls.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Top             =   465
         Width           =   5925
         _Version        =   786432
         _ExtentX        =   10451
         _ExtentY        =   503
         _StockProps     =   93
         Appearance      =   6
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   6165
         Y1              =   870
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   6150
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   6165
         X2              =   6165
         Y1              =   15
         Y2              =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   840
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   270
         Left            =   60
         Top             =   60
         Width           =   6075
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando Cta. Cte."
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
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   90
         Width           =   1845
      End
   End
   Begin SizerOneLibCtl.ElasticOne EO 
      Height          =   5775
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   555
      Width           =   14355
      _cx             =   25321
      _cy             =   10186
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   8
      BorderWidth     =   3
      ChildSpacing    =   3
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   2
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmAnalisisCliente2.frx":0000
      Begin SizerOneLibCtl.TabOne TabOne1 
         Height          =   4575
         Left            =   45
         TabIndex        =   4
         Top             =   1155
         Width           =   14265
         _cx             =   25162
         _cy             =   8070
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
         FrontTabForeColor=   -2147483630
         Caption         =   "   Resumen     |    Detalle    "
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
         Begin VB.Frame FrameDetalle 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   4155
            Left            =   -14820
            TabIndex        =   7
            Top             =   45
            Width           =   14175
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   4110
               Left            =   -105
               TabIndex        =   8
               Top             =   0
               Width           =   11550
               _cx             =   20373
               _cy             =   7250
               _ConvInfo       =   1
               Appearance      =   1
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
               BackColor       =   14876411
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   128
               ForeColorSel    =   16777215
               BackColorBkg    =   -2147483636
               BackColorAlternate=   14876411
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
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmAnalisisCliente2.frx":0042
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
         Begin VB.Frame FramResumen 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   4155
            Left            =   45
            TabIndex        =   5
            Top             =   45
            Width           =   14175
            Begin VSFlex7Ctl.VSFlexGrid Fg3 
               Height          =   4125
               Left            =   0
               TabIndex        =   6
               Top             =   0
               Width           =   11550
               _cx             =   20373
               _cy             =   7276
               _ConvInfo       =   1
               Appearance      =   1
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
               BackColor       =   14876411
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   128
               ForeColorSel    =   16777215
               BackColorBkg    =   -2147483636
               BackColorAlternate=   14876411
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
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmAnalisisCliente2.frx":0117
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
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne1 
         Height          =   1065
         Left            =   45
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   45
         Width           =   14265
         _cx             =   25162
         _cy             =   1879
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   0
         ChildSpacing    =   1
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   1
         GridCols        =   4
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmAnalisisCliente2.frx":01ED
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   1065
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   3360
            Begin VB.CommandButton CmdBusMon 
               Enabled         =   0   'False
               Height          =   240
               Left            =   1230
               Picture         =   "FrmAnalisisCliente2.frx":024B
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   720
               Width           =   240
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   870
               TabIndex        =   29
               Top             =   60
               Width           =   1305
               _ExtentX        =   2302
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
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
               Height          =   300
               Left            =   870
               TabIndex        =   30
               Top             =   375
               Width           =   1305
               _ExtentX        =   2302
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
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   870
               Locked          =   -1  'True
               TabIndex        =   28
               Text            =   "TxtIdMon"
               Top             =   690
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Inicio"
               Height          =   195
               Left            =   45
               TabIndex        =   34
               Top             =   135
               Width           =   735
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Venc."
               Height          =   195
               Left            =   45
               TabIndex        =   33
               Top             =   420
               Width           =   780
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Left            =   45
               TabIndex        =   32
               Top             =   720
               Width           =   585
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
               Left            =   1515
               TabIndex        =   31
               Top             =   690
               Width           =   1770
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   1065
            Left            =   3375
            TabIndex        =   23
            Top             =   0
            Width           =   4485
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   750
               Left            =   0
               TabIndex        =   24
               Top             =   255
               Width           =   4380
               _cx             =   7726
               _cy             =   1323
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
               BackColor       =   14876411
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   128
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   14876411
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
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmAnalisisCliente2.frx":037D
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
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
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
               Height          =   195
               Left            =   15
               TabIndex        =   25
               Top             =   30
               Width           =   600
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   1065
            Left            =   7875
            TabIndex        =   15
            Top             =   0
            Width           =   3390
            Begin VB.OptionButton Option4 
               Caption         =   "Pendientes"
               Height          =   195
               Left            =   210
               TabIndex        =   21
               Top             =   330
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Cancelados"
               Height          =   195
               Left            =   210
               TabIndex        =   20
               Top             =   555
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Todos"
               Height          =   195
               Left            =   210
               TabIndex        =   19
               Top             =   780
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.Frame Frame66 
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   720
               Left            =   1755
               TabIndex        =   16
               Top             =   285
               Width           =   1575
               Begin VB.CheckBox Check1 
                  Caption         =   "Apertura"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   18
                  Top             =   45
                  Width           =   1170
               End
               Begin VB.CheckBox Check2 
                  Caption         =   "Año Trabajo"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   17
                  Top             =   270
                  Width           =   1290
               End
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Opciones"
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
               Left            =   15
               TabIndex        =   22
               Top             =   30
               Visible         =   0   'False
               Width           =   810
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   1065
            Left            =   11280
            TabIndex        =   10
            Top             =   0
            Width           =   2985
            Begin VB.CommandButton Command1 
               Caption         =   "Agregar Documentos"
               Height          =   360
               Left            =   105
               TabIndex        =   13
               Top             =   180
               Width           =   2070
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Eliminar OD No Marcadas"
               Height          =   360
               Left            =   105
               TabIndex        =   12
               Top             =   540
               Visible         =   0   'False
               Width           =   2070
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Cargar Datos"
               Height          =   735
               Left            =   2205
               TabIndex        =   11
               Top             =   180
               Width           =   600
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Mostrar Solo"
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
               Left            =   90
               TabIndex        =   14
               Top             =   30
               Visible         =   0   'False
               Width           =   1080
            End
         End
      End
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   3765
      Top             =   0
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   4185
      Top             =   30
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmAnalisisCliente2.frx":03CD
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   750
      TabIndex        =   35
      Top             =   165
      Width           =   1290
   End
End
Attribute VB_Name = "FrmAnalisisCliente2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SeEjecuto As Boolean
Dim TopEO As Integer

Dim TAMAÑO_TOOL As TOOL_TAMAÑO_ICO

Const BTN_BUS = 1
Const BTN_EXP = 2
Const BTN_IMP = 3
Const BTN_CON = 4
Const BTN_SAL = 5

Dim RstDet As New ADODB.Recordset
Dim RstDetAux As New ADODB.Recordset

Private Sub CmdBusMon_Click()
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
        If xRs.RecordCount <> 0 Then
            TxtIdMon.Text = xRs("id")
            LblMoneda.Caption = xRs("descripcion")
            TxtIdMon_Validate True
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Command1_Click()
'    If NulosC(Fg1.TextMatrix(Fg1.Row, 8)) <> "Nº DE DOC. ==>" Then
'        MsgBox "No ha seleccionado el numero de orden de despacho que se desea eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Exit Sub
'    End If
'    EliminarOrden Fg1.TextMatrix(Fg1.Row, 9)
'    MsgBox "La Ordden de Depacho se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    FrmAgregaDoc.Show
End Sub

Sub EliminarOrden(OrdenDespacho As String)
    Dim xIdCli As Integer
    xIdCli = NulosN(Fg2.TextMatrix(0, 2))
    ' ELIMINAMOS LAS VENTAS
    xCon.Execute "DELETE * FROM vta_ventas WHERE numerodocref = '" & OrdenDespacho & "' and idcli = " & xIdCli & ""

    ' ELIMINAMOS LAS LGD
    xCon.Execute "DELETE * FROM vta_gastodebito WHERE numerodocref = '" & OrdenDespacho & "' and idcli = " & xIdCli & ""

    ' ELIMINAMOS LAS LETRAS
'    118102008001165
'    123456789012345
'    xCon.Execute "DELETE let_letra.* From let_letra " _
        & " WHERE ((let_letra.numorden='" & Mid(OrdenDespacho, 10, 6) & "') AND (let_letra.anoorden=" & Val(Mid(OrdenDespacho, 6, 4)) & ") " _
        & " AND (let_letra.idaduana=" & Val(Mid(OrdenDespacho, 1, 3)) & ") AND (let_letra.idregimen=" & Val(Mid(OrdenDespacho, 4, 2)) & ") AND (idclipro = " & xIdCli & "))"

    xCon.Execute "DELETE let_letra.* From let_letra " _
        & " WHERE ((numrefjunto = '" & OrdenDespacho & "') AND (idclipro = " & xIdCli & "))"
   
End Sub

Private Sub Command2_Click()
'    Dim A As Integer
'    For A = 2 To Fg1.Rows - 1
'        If NulosC(Fg1.TextMatrix(A, 8)) = "Nº DE DOC. ==>" Then
'            If NulosN(Fg1.TextMatrix(A, Fg1.Cols - 1)) = 0 Then
'                EliminarOrden Fg1.TextMatrix(A, 9)
'            End If
'        End If
'    Next A
'
'    MsgBox "Las Orddene de Depacho se eliminaron con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Sub

Private Sub Command3_Click()
    Dim xCadWhere As String
    Dim A As Integer
    Fg1.Rows = 2
    Fg3.Rows = 2
    
    If Check1.Value = 0 And Check2.Value = 0 Then
        MsgBox "No ha especificado que datos se van a mostrar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Check1.SetFocus
        Exit Sub
    End If
       
    If TxtFchIni.Valor = "" Then
        MsgBox "No ha especificado la fecha de inicio ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If

    If TxtFchFin.Valor = "" Then
        MsgBox "No ha especificado la fecha final ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If

    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Fg1.MergeCells = flexMergeFixedOnly
    
    TraerDatos
End Sub

Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Id = 1 Then
        If Check1.Value = 0 And Check2.Value = 0 Then
            MsgBox "No ha especificado que datos se van a mostrar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Check1.SetFocus
            Exit Sub
        End If
           
        If TxtFchIni.Valor = "" Then
            MsgBox "No ha especificado la fecha de inicio ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Sub
        End If
    
        If TxtFchFin.Valor = "" Then
            MsgBox "No ha especificado la fecha final ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchFin.SetFocus
            Exit Sub
        End If
    
        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        
        CargarSelect
    End If
    If Control.Id = 2 Then pExportar
    If Control.Id = 5 Then
        Unload Me
    End If
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Cliente":       xCampos(0, 1) = "nombre":           xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "abrev":            xCampos(1, 2) = "800":          xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Documento":  xCampos(2, 1) = "numruc":           xCampos(2, 2) = "1500":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Tipo Empresa":  xCampos(3, 1) = "tipemp":           xCampos(3, 2) = "1500":         xCampos(3, 3) = "C"
    
    xform.SQLCad = "SELECT mae_cliente.nombre, mae_dociden.abrev, mae_tipoempresa.descripcion AS tipemp, mae_cliente.numruc, " _
        & " mae_cliente.id FROM (mae_cliente LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id) LEFT JOIN mae_tipoempresa " _
        & " ON mae_cliente.tipper = mae_tipoempresa.id Where (((mae_cliente.activo) = -1)) ORDER BY mae_cliente.nombre"
    
    xform.Titulo = "Buscando Cliente"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = xRs("nombre")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = xRs("id")
            
            If Fg2.TextMatrix(Fg2.Rows - 1, 1) <> "" Then
                Fg2.Rows = Fg2.Rows + 1
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then ' INSERTA UNA FILA
        If NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 2)) <> 0 Then
            Fg2.Rows = Fg2.Rows + 1
        End If
    End If
    If KeyCode = 46 Then    ' ELIMINA UNA FILA
        If NulosN(Fg2.TextMatrix(Fg2.Row, 2)) <> 0 Then
            Fg2.RemoveItem Fg2.Row
            If Fg2.Rows = 0 Then Fg2.Rows = Fg2.Rows + 1
            
        End If
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
         SeEjecuto = True
    End If
End Sub

Private Sub Form_Load()
    Me.WindowState = 2
    SeEjecuto = False
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
    
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Frame5.BackColor = &H8000000F
    Frame4.BackColor = &H8000000F
    
    SetearCuadricula Fg1, 5, xCon, 5, 3, False
    SetearCuadricula Fg3, 5, xCon, 5, 1, False
    
    Blanquea
    CrearTool
    
    If TAMAÑO_TOOL = I16x16 Then EO.Top = 400: TopEO = 400
    If TAMAÑO_TOOL = I24x24 Then EO.Top = 520: TopEO = 520
    If TAMAÑO_TOOL = I32x32 Then EO.Top = 640: TopEO = 640
    If TAMAÑO_TOOL = I48x48 Then EO.Top = 890: TopEO = 890
    
    Fg2.Rows = 0
    Fg2.Rows = Fg2.Rows + 1
    Fg2.ColComboList(1) = "|..."
    Fg2.SelectionMode = flexSelectionFree
    Fg2.ColWidth(2) = 0
    Fg2.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionByRow
    
    TxtIdMon.Text = "1"
    TxtIdMon_Validate True
    
    Option4.Value = True
    Check2.Value = 1
    'TxtFchIni.SetFocus
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    EO.Width = Me.Width - 130
    If Me.Height <= (TopEO + 2375) Then
        Me.Height = (TopEO + 2375)
    Else
        EO.Height = (Me.Height - (TopEO + 400))
    End If
    
    Me.Refresh
    Fg1.Height = FrameDetalle.Height '- 10
    Fg1.Width = FrameDetalle.Width '- 10

    Fg3.Height = FramResumen.Height '- 10
    Fg3.Width = FramResumen.Width '- 10
    
    Frame3.Left = ((Me.Width - Frame3.Width) / 2)
    Frame3.Top = ((Me.Height - Frame3.Height) / 2)
    Frame3.Visible = False
End Sub


Sub Blanquea()
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    TxtIdMon.Text = ""
    LblMoneda.Caption = ""
    Fg2.Rows = 0
    'Fg3.Rows = 0
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If NulosN(TxtIdMon.Text) = 0 Then Exit Sub
    
    LblMoneda.Caption = Busca_Codigo(TxtIdMon.Text, "id", "descripcion", "mae_moneda", "N", xCon)
    If NulosC(LblMoneda.Caption) = "" Then
        TxtIdMon.Text = ""
    Else
'        If TxtIdMon = 2 Then
'            Fg1.ColWidth(9) = 1100
'            Fg1.ColWidth(10) = 1100
'            Fg1.ColWidth(11) = 1100
'
'            Fg1.ColWidth(12) = 0
'            Fg1.ColWidth(13) = 0
'            Fg1.ColWidth(14) = 0
'        Else
'            Fg1.ColWidth(9) = 0
'            Fg1.ColWidth(10) = 0
'            Fg1.ColWidth(11) = 0
'
'            Fg1.ColWidth(12) = 1100
'            Fg1.ColWidth(13) = 1100
'            Fg1.ColWidth(14) = 1100
'        End If
    End If
End Sub

Sub CrearTool()
    'CREAMOS EL TOOLBAR
    Dim Opciones(4, 3) As String
    
    Opciones(0, 0) = Str(BTN_BUS):    Opciones(0, 1) = "Buscar":                      Opciones(0, 2) = "0":      Opciones(0, 3) = "Ejecutar Busqueda"
    Opciones(1, 0) = Str(BTN_EXP):    Opciones(1, 1) = "Exportar Excel":              Opciones(1, 2) = "0":      Opciones(1, 3) = "Exportar Excel"
    Opciones(2, 0) = Str(BTN_IMP):    Opciones(2, 1) = "Imprimir":                    Opciones(2, 2) = "0":      Opciones(2, 3) = "Imprimir"
    Opciones(3, 0) = Str(BTN_CON):    Opciones(3, 1) = "Configurar":                  Opciones(3, 2) = "0":      Opciones(3, 3) = "Configurar"
    Opciones(4, 0) = Str(BTN_SAL):    Opciones(4, 1) = "Salir":                       Opciones(4, 2) = "1":      Opciones(4, 3) = "Salir"
        
    Dim xFun As New eps_librerias.Codejock
    'PocisionarContenedor
    xFun.BORRARMENU = True
    TAMAÑO_TOOL = I24x24
    xFun.CrearToolBar Opciones, CommandBars1, ImageManager1, TAMAÑO_TOOL
    Set xFun = Nothing
End Sub

Private Sub pExportar()
    Dim xFun As New SGI2_funciones.formularios
    Dim Rst As New ADODB.Recordset
    
    'If TabOne1.CurrTab = 0 Then
        If Fg1.Rows = 1 Then
            MsgBox "No hay registro para exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        If TabOne1.CurrTab = 0 Then
            'If Fg1.Rows <= 65000 Then
                GRID_EXPORTAR_MSEXCELTMP Fg1, xCon, flexFileCustomText, True, "Analisis de Cuenta x Cliente - DETALLADO", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en " & LblMoneda.Caption
            'Else
            '    ExportarMultipleHojasExcel Fg1, "Prueba.xls"
            'End If
        Else
            GRID_EXPORTAR_MSEXCELTMP Fg3, xCon, flexFileCustomText, True, "Analisis de Cuenta x Cliente - RESUMEM", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en " & LblMoneda.Caption
        End If
        Set xFun = Nothing
    'End If
End Sub

'Sub ExportarMultipleHojasExcel(Fg As Object, NombreArchivo As String)
'    Dim NumHojas As Integer
'    Dim A As Integer
'    NumHojas = (Fg.Rows / 65500)
'    Dim FgExp As Object
'    Dim xUltimaFila, xPrimeraFila As Double
'
'    xUltimaFila = 65500
'    xPrimeraFila = 1
'    For A = 1 To NumHojas
'        Set FgExp = Fg
'        FgExp.RemoveItem A
'    Next A
'End Sub

Sub TraerDatos()
    Dim xCad As String
    Dim xIdCliente As Integer
    Dim A As Integer
    
    '*******************************************************************************************
    ' CARGAMOS LOS DATOS DE CAJA Y BANCOS

On Error GoTo ExisteTabla

    xCon.Execute "DROP TABLE diario_libros"
    xCon.Execute "DROP TABLE con_diario_final"
    
    ' COPIAMOS LOS REGISTROS DEL LIBRO CAJA BANCOS A UNA TABLA TEMPORAL
    xCon.Execute "SELECT con_diario.* INTO diario_libros From con_diario WHERE (((con_diario.idlib)=6))"
    
    xCad = " SELECT DISTINCT diario_libros.idcue, con_planctas.cuenta, con_planctas.descripcion, diario_libros.fchdoc, tes_cajadestinodet.docctacte AS abrev, " _
        & " diario_libros.rnumerodoc, diario_libros.idmon, diario_libros.tc, [diario_libros]![imphabsol] AS impdebsol, [diario_libros]![impdebsol] AS imphabsol, " _
        & " [diario_libros]![imphabdol] AS impdebdol, [diario_libros]![impdebdol] AS imphabdol, diario_libros.rnumerodoc1, diario_libros.ridper AS idper, " _
        & " mae_cliente.nombre, mid([tes_caja]![numreg],1,2) & [mae_libros]![codsun] & mid([tes_caja]![numreg],3,4) AS numreg INTO con_diario_final FROM (tes_caja LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) " _
        & " LEFT JOIN (tes_cajadestino LEFT JOIN ((((diario_libros LEFT JOIN con_planctas ON diario_libros.idcue = con_planctas.id) RIGHT JOIN tes_cajadestinodet " _
        & " ON diario_libros.idmov = tes_cajadestinodet.idtes) LEFT JOIN mae_documento ON diario_libros.rtipdoc = mae_documento.id) LEFT JOIN mae_cliente " _
        & " ON tes_cajadestinodet.idper = mae_cliente.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) " _
        & " ON tes_caja.id = tes_cajadestino.idtes WHERE (((diario_libros.idlib)=6) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND ((tes_caja.tipmov)=1) AND ((diario_libros.imphabsol)<>0)) OR (((diario_libros.idlib)=6) " _
        & " AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND ((tes_caja.tipmov)=1) " _
        & " AND ((diario_libros.imphabdol)<>0)) ORDER BY diario_libros.rnumerodoc1"

    ' EJECUTAMOS LA CONSULTA DE CREACION DE TABLA PARA CREAR LA TABLA con_diario_final
    xCon.Execute xCad
    
    '*******************************************************************************************
    ' INSERTAMOS LOS REGISTROS DE EGRESOS DE CAJA Y BANCOS
    'xCad = "INSERT INTO con_diario_final ( idcue, cuenta, descripcion, fchdoc, abrev, rnumerodoc, idmon, tc, impdebsol, imphabsol, impdebdol, imphabdol, rnumerodoc1, " _
        & " idper, nombre, numreg ) SELECT DISTINCT diario_libros.idcue, con_planctas.cuenta, con_planctas.descripcion, diario_libros.fchdoc, tes_cajadestinodet.docctacte AS abrev, " _
        & " diario_libros.rnumerodoc, diario_libros.idmon, diario_libros.tc, [diario_libros]![imphabsol] AS impdebsol, [diario_libros]![impdebsol] AS imphabsol, " _
        & " [diario_libros]![imphabdol] AS impdebdol, [diario_libros]![impdebdol] AS imphabdol, diario_libros.rnumerodoc1, diario_libros.ridper AS idper, mae_cliente.nombre, " _
        & " [mae_libros]![codsun] & [tes_caja]![numreg] AS numreg FROM (tes_caja LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) LEFT JOIN (tes_cajadestino " _
        & " LEFT JOIN ((((diario_libros LEFT JOIN con_planctas ON diario_libros.idcue = con_planctas.id) RIGHT JOIN tes_cajadestinodet ON diario_libros.idmov = tes_cajadestinodet.idtes) " _
        & " LEFT JOIN mae_documento ON diario_libros.rtipdoc = mae_documento.id) LEFT JOIN mae_cliente ON tes_cajadestinodet.idper = mae_cliente.id) ON " _
        & " (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes " _
        & " WHERE (((diario_libros.idlib)=6) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((tes_caja.tipmov)=2) AND ((diario_libros.impdebsol)<>0)) OR (((diario_libros.idlib)=6) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND ((tes_caja.tipmov)=2) AND ((diario_libros.impdebdol)<>0)) " _
        & " ORDER BY diario_libros.rnumerodoc1 "
       
    xCad = "INSERT INTO con_diario_final ( idcue, cuenta, descripcion, fchdoc, abrev, rnumerodoc, idmon, tc, imphabsol, impdebsol, imphabdol, impdebdol, rnumerodoc1, " _
        & " idper, nombre, numreg ) SELECT DISTINCT diario_libros.idcue, con_planctas.cuenta, con_planctas.descripcion, diario_libros.fchdoc, tes_cajadestinodet.docctacte AS abrev, " _
        & " diario_libros.rnumerodoc, diario_libros.idmon, diario_libros.tc, [diario_libros]![imphabsol] AS impdebsol, [diario_libros]![impdebsol] AS imphabsol, " _
        & " [diario_libros]![imphabdol] AS impdebdol, [diario_libros]![impdebdol] AS imphabdol, diario_libros.rnumerodoc1, diario_libros.ridper AS idper, " _
        & " mae_cliente.nombre, mid([tes_caja]![numreg],1,2) & [mae_libros]![codsun] & mid([tes_caja]![numreg],3,4) AS numreg FROM (tes_caja LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) LEFT JOIN " _
        & " (tes_cajadestino LEFT JOIN ((((diario_libros LEFT JOIN con_planctas ON diario_libros.idcue = con_planctas.id) RIGHT JOIN tes_cajadestinodet ON diario_libros.idmov = tes_cajadestinodet.idtes) " _
        & " LEFT JOIN mae_documento ON diario_libros.rtipdoc = mae_documento.id) LEFT JOIN mae_cliente ON tes_cajadestinodet.idper = mae_cliente.id) " _
        & " ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes " _
        & " WHERE (((diario_libros.idlib)=6) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((tes_caja.tipmov)=2) AND ((diario_libros.impdebsol)<>0)) OR (((diario_libros.idlib)=6) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND ((tes_caja.tipmov)=2) AND ((diario_libros.impdebdol)<>0)) ORDER BY diario_libros.rnumerodoc1"
   
    ' EJECUTAMOS LA CONSULTA DE CREACION DE TABLA PARA CREAR LA TABLA con_diario_final
    xCon.Execute xCad
        
        
    '*******************************************************************************************
    ' INSERTAMOS LOS REGISTROS DE APERTURA
    xCad = "INSERT INTO con_diario_final ( idcue, cuenta, descripcion, fchdoc, abrev, rnumerodoc, idmon, tc, impdebsol, imphabsol, impdebdol, imphabdol, rnumerodoc1, " _
        & " idper, nombre, numreg ) SELECT con_provicionesdetdoc_AP.idcue, con_planctas.cuenta, con_planctas.descripcion, con_provicionesdetdoc_AP.fchemi, " _
        & " mae_documento.abrev, IIf([con_provicionesdetdoc_AP]![numser] Is Not Null,Trim([con_provicionesdetdoc_AP]![numser]) & '-' & Trim([con_provicionesdetdoc_AP]![numdoc])," _
        & " Trim([con_provicionesdetdoc_AP]![numdoc])) AS Expr2, con_provicionesdetdoc_AP.idmon, con_tc.impven, con_provicionesdetdoc_AP.cargosol, con_provicionesdetdoc_AP.abonosol, " _
        & " con_provicionesdetdoc_AP.cargodol, con_provicionesdetdoc_AP.abonodol, con_provicionesdetdoc_AP.numorden AS Expr1, con_provicionesdetdoc_AP.idclipro, " _
        & " mae_cliente.nombre, [mae_libros]![codsun] & [numreg] AS numreg1 FROM mae_libros RIGHT JOIN ((((((con_provicionesdetdoc_AP LEFT JOIN con_planctas " _
        & " ON con_provicionesdetdoc_AP.idcue = con_planctas.id) LEFT JOIN mae_documento ON con_provicionesdetdoc_AP.tipdoc = mae_documento.id) LEFT JOIN con_tc " _
        & " ON con_provicionesdetdoc_AP.fchemi = con_tc.fecha) LEFT JOIN con_provicionesdet ON (con_provicionesdetdoc_AP.idcue = con_provicionesdet.idcuen) " _
        & " AND (con_provicionesdetdoc_AP.idpro = con_provicionesdet.id)) LEFT JOIN mae_cliente ON con_provicionesdetdoc_AP.idclipro = mae_cliente.id) " _
        & " LEFT JOIN con_proviciones ON con_provicionesdetdoc_AP.idpro = con_proviciones.id) ON mae_libros.id = con_proviciones.idlib WHERE (((con_proviciones.idmes)=0))"


    ' EJECUTAMOS LA CONSULTA DE DATOS ANEXADOS
    xCon.Execute xCad
    
    
    '*******************************************************************************************
    ' CARGAMOS LAS LIQUIDACIONES GASTO DEBITO
    
    ' ELIMINAMOS LA TABLA diario libros
    xCon.Execute "DROP TABLE diario_libros"
    
    ' COPIAMOS LOS REGISTRO DEL LIBRO LIQUIDACION GASTO DEBITO  A LA TABLA TEMPORAL diario_libros
    xCon.Execute "SELECT con_diario.* INTO diario_libros From con_diario WHERE (((con_diario.idlib)=41))"

    ' CARGAMOS LA LIQUIDACION GASTO CREDITO
    xCad = "INSERT INTO con_diario_final ( idcue, cuenta, descripcion, fchdoc, abrev, rnumerodoc, idmon, tc, impdebsol, imphabsol, impdebdol, imphabdol, rnumerodoc1, " _
        & " idper, nombre, numreg ) SELECT DISTINCT diario_libros.idcue, con_planctas.cuenta, con_planctas.descripcion, diario_libros.fchdoc, mae_documento.abrev, " _
        & " vta_gastodebito!numser & '-' & vta_gastodebito!numdoc AS rnumerodoc, diario_libros.idmon, diario_libros.tc, diario_libros.impdebsol, diario_libros.imphabsol, " _
        & " diario_libros.impdebdol, diario_libros.imphabdol, diario_libros.rnumerodoc1, vta_gastodebito.idcli, mae_cliente.nombre, " _
        & " Format([diario_libros]![idmes],'00') & [mae_libros]![codsun] & [diario_libros]![numasi] AS numreg FROM ((((diario_libros LEFT JOIN con_planctas ON diario_libros.idcue = con_planctas.id) " _
        & " LEFT JOIN vta_gastodebito ON diario_libros.idmov = vta_gastodebito.id) LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) " _
        & " LEFT JOIN mae_cliente ON vta_gastodebito.idcli = mae_cliente.id) LEFT JOIN mae_libros ON diario_libros.idlib = mae_libros.id WHERE (((diario_libros.impdebsol)<>0) " _
        & " AND ((diario_libros.idlib)=41) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((vta_gastodebito.tipdoc)=126)) OR (((diario_libros.impdebdol)<>0) AND ((diario_libros.idlib)=41) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_gastodebito.tipdoc)=126)) ORDER BY diario_libros.rnumerodoc1"

    ' EJECUTAMOS LA CONSULTA DE DATOS ANEXADOS Y AGREGAMOS LOS REGISTRO DE LA TABLA TEMPORAL A LA TABLA con_diario_final
    xCon.Execute xCad
    
    ' CARGAMOS LA LIQUIDACION GASTO DEBITO
    xCad = "INSERT INTO con_diario_final ( idcue, cuenta, descripcion, fchdoc, abrev, rnumerodoc, idmon, tc, imphabsol, impdebsol, imphabdol, impdebdol, rnumerodoc1, " _
        & " idper, nombre, numreg ) SELECT DISTINCT diario_libros.idcue, con_planctas.cuenta, con_planctas.descripcion, diario_libros.fchdoc, mae_documento.abrev, " _
        & " vta_gastodebito!numser & '-' & vta_gastodebito!numdoc AS rnumerodoc, diario_libros.idmon, diario_libros.tc, diario_libros.impdebsol, diario_libros.imphabsol, " _
        & " diario_libros.impdebdol, diario_libros.imphabdol, diario_libros.rnumerodoc1, vta_gastodebito.idcli, mae_cliente.nombre, " _
        & " Format([diario_libros]![idmes],'00') & [mae_libros]![codsun] & [diario_libros]![numasi] AS numreg FROM ((((diario_libros LEFT JOIN con_planctas " _
        & " ON diario_libros.idcue = con_planctas.id) LEFT JOIN vta_gastodebito ON diario_libros.idmov = vta_gastodebito.id) LEFT JOIN mae_documento " _
        & " ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN mae_cliente ON vta_gastodebito.idcli = mae_cliente.id) LEFT JOIN mae_libros ON diario_libros.idlib = mae_libros.id " _
        & " WHERE (((diario_libros.impdebsol)<>0) AND ((diario_libros.idlib)=41) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') And " _
        & " (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_gastodebito.tipdoc)=120)) OR (((diario_libros.impdebdol)<>0) AND ((diario_libros.idlib)=41) " _
        & " AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_gastodebito.tipdoc)=120)) " _
        & " ORDER BY diario_libros.rnumerodoc1"

    ' EJECUTAMOS LA CONSULTA DE DATOS ANEXADOS Y AGREGAMOS LOS REGISTRO DE LA TABLA TEMPORAL A LA TABLA con_diario_final
    xCon.Execute xCad
    
    
    
    '*******************************************************************************************
    ' CARGAMOS LAS RETENCIONES
    
    ' ELIMINAMOS LA TABLA diario libros
    xCon.Execute "DROP TABLE diario_libros"
    
    'COPIAMOS LOS REGISTRO DE LA RETENCIONES A LA TABLA TEMPORAL diario_libros
    xCon.Execute "SELECT con_diario.* INTO diario_libros From con_diario WHERE (((con_diario.idlib)=5))"
    
    xCad = "INSERT INTO con_diario_final ( idcue, cuenta, descripcion, fchdoc, abrev, rnumerodoc, idmon, tc, imphabsol, impdebsol, imphabdol, impdebdol, rnumerodoc1, " _
        & " idper, nombre, numreg ) SELECT DISTINCT diario_libros.idcue, con_planctas.cuenta, con_planctas.descripcion, diario_libros.fchdoc, mae_documento.abrev, " _
        & " [con_retencion]![numser] & '-' & [con_retencion]![numdoc] AS rnumerodoc, diario_libros.idmon, diario_libros.tc, diario_libros.impdebsol, diario_libros.imphabsol, " _
        & " diario_libros.impdebdol, diario_libros.imphabdol, diario_libros.rnumerodoc1, con_retencion.idpro, mae_cliente.nombre, " _
        & " Format([diario_libros]![idmes],'00') & [mae_libros]![codsun] & [diario_libros]![numasi] AS numreg FROM ((((diario_libros LEFT JOIN con_planctas " _
        & " ON diario_libros.idcue = con_planctas.id) LEFT JOIN mae_documento ON diario_libros.rtipdoc = mae_documento.id) LEFT JOIN con_retencion ON diario_libros.idmov = con_retencion.id) " _
        & " LEFT JOIN mae_cliente ON con_retencion.idpro = mae_cliente.id) LEFT JOIN mae_libros ON diario_libros.idlib = mae_libros.id WHERE (((diario_libros.imphabsol)<>0) " _
        & " AND ((diario_libros.idlib)=5) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "'))) " _
        & " OR (((diario_libros.imphabdol)<>0) AND ((diario_libros.idlib)=5) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')))"

    ' EJECUTAMOS LA CONSULTA DE DATOS ANEXADOS Y AGREGAMOS LOS REGISTRO DE LA TALA TEMPORAL A LA TABLA con_diario_final
    xCon.Execute xCad
    
        
    '*******************************************************************************************
    ' CARGAMOS LAS VENTAS (facturas boletas de venta y notas de debito)
    
    ' ELIMINAMOS LA TABLA diario libros
    xCon.Execute "DROP TABLE diario_libros"
    
    'COPIAMOS LOS REGISTRO DE VENTA A LA TABLA TEMPORAL diario_libros
    xCon.Execute "SELECT con_diario.* INTO diario_libros From con_diario WHERE (((con_diario.idlib)=2))"

    xCad = "INSERT INTO con_diario_final ( idcue, cuenta, descripcion, fchdoc, abrev, rnumerodoc, idmon, tc, imphabsol, impdebsol, imphabdol, impdebdol, rnumerodoc1, " _
        & " idper, nombre, numreg ) SELECT DISTINCT diario_libros.idcue, con_planctas.cuenta, con_planctas.descripcion, diario_libros.fchdoc, mae_documento.abrev, " _
        & " vta_ventas!numser & '-' & vta_ventas!numdoc AS rnumerodoc, diario_libros.idmon, diario_libros.tc, IIf([diario_libros]![idmon]=1,[impdebsol],0) AS impdebsol2, " _
        & " diario_libros.imphabsol, diario_libros.impdebdol, diario_libros.imphabdol, diario_libros.rnumerodoc1, vta_ventas.idcli, mae_cliente.nombre, " _
        & " Format([diario_libros]![idmes],'00') & [mae_libros]![codsun] & [diario_libros]![numasi] AS numreg FROM (mae_cliente RIGHT JOIN (((diario_libros LEFT JOIN " _
        & " con_planctas ON diario_libros.idcue = con_planctas.id) LEFT JOIN vta_ventas ON diario_libros.idmov = vta_ventas.id) LEFT JOIN mae_documento " _
        & " ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_libros ON diario_libros.idlib = mae_libros.id " _
        & " WHERE (((IIf([diario_libros]![idmon]=1,[impdebsol],0))<>0) AND ((diario_libros.idlib)=2) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_ventas.tipdoc)<>7)) OR (((diario_libros.impdebdol)<>0) AND ((diario_libros.idlib)=2) " _
        & " AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_ventas.tipdoc)<>7)) " _
        & " ORDER BY diario_libros.rnumerodoc1"


    ' EJECUTAMOS LA CONSULTA DE DATOS ANEXADOS Y AGREGAMOS LOS REGISTRO DE LA TALA TEMPORAL A LA TABLA con_diario_final
    xCon.Execute xCad
      
      
    '*******************************************************************************************
    ' CARGAMOS LAS VENTAS (Notas de Credito)
    
    xCad = "INSERT INTO con_diario_final ( idcue, cuenta, descripcion, fchdoc, abrev, rnumerodoc, idmon, tc, impdebsol, imphabsol, impdebdol, imphabdol, rnumerodoc1, " _
        & " idper, nombre, numreg ) SELECT DISTINCT diario_libros.idcue, con_planctas.cuenta, con_planctas.descripcion, diario_libros.fchdoc, mae_documento.abrev, " _
        & " vta_ventas!numser & '-' & vta_ventas!numdoc AS rnumerodoc, diario_libros.idmon, diario_libros.tc, IIf([diario_libros]![idmon]=1,[impdebsol],0) AS impdebsol2, " _
        & " diario_libros.imphabsol, diario_libros.impdebdol, diario_libros.imphabdol, diario_libros.rnumerodoc1, vta_ventas.idcli, mae_cliente.nombre, " _
        & " Format([diario_libros]![idmes],'00') & [mae_libros]![codsun] & [diario_libros]![numasi] AS numreg FROM (mae_cliente RIGHT JOIN (((diario_libros LEFT JOIN " _
        & " con_planctas ON diario_libros.idcue = con_planctas.id) LEFT JOIN vta_ventas ON diario_libros.idmov = vta_ventas.id) LEFT JOIN mae_documento " _
        & " ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_libros ON diario_libros.idlib = mae_libros.id " _
        & " WHERE (((IIf([diario_libros]![idmon]=1,[impdebsol],0))<>0) AND ((diario_libros.idlib)=2) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') And " _
        & " (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_ventas.tipdoc)=7)) OR (((diario_libros.impdebdol)<>0) AND ((diario_libros.idlib)=2) AND " _
        & " ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_ventas.tipdoc)=7)) " _
        & " ORDER BY diario_libros.rnumerodoc1"

    ' EJECUTAMOS LA CONSULTA DE DATOS ANEXADOS Y AGREGAMOS LOS REGISTRO DE LA TALA TEMPORAL A LA TABLA con_diario_final
    xCon.Execute xCad
    
    
    
    '*******************************************************************************************
    ' CARGAMOS LAS LETRAS
    xCon.Execute "DROP TABLE diario_libros"
    
    'COPIAMOS LOS REGISTRO DE LETRAS A LA TABLA TEMPORAL diario_libros
    xCon.Execute "SELECT con_diario.* INTO diario_libros From con_diario WHERE (((con_diario.idlib)=37))"

    xCad = "INSERT INTO con_diario_final ( idcue, cuenta, descripcion, fchdoc, abrev, rnumerodoc, idmon, tc, imphabsol, impdebsol, imphabdol, impdebdol, rnumerodoc1, " _
        & " idper, nombre, numreg ) SELECT DISTINCT diario_libros.idcue, con_planctas.cuenta, con_planctas.descripcion, diario_libros.fchdoc, mae_documento.abrev, " _
        & " diario_libros.rnumerodoc, diario_libros.idmon, diario_libros.tc, diario_libros.imphabsol AS impdebsol, diario_libros.impdebsol AS imphabsol, " _
        & " diario_libros.imphabdol AS impdebdol, diario_libros.impdebdol AS imphabdol, diario_libros.rnumerodoc1, let_letra.idclipro AS idcli, mae_cliente.nombre, " _
        & " Format(diario_libros!idmes,'00') & mae_libros!codsun & diario_libros!numasi AS numreg FROM mae_documento RIGHT JOIN (((((diario_libros LEFT JOIN con_planctas " _
        & " ON diario_libros.idcue = con_planctas.id) LEFT JOIN let_letra ON diario_libros.idmov = let_letra.id) LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id) " _
        & " LEFT JOIN mae_libros ON diario_libros.idlib = mae_libros.id) LEFT JOIN let_letradet ON let_letra.id = let_letradet.idlet) ON mae_documento.id = let_letra.tipdoc " _
        & " WHERE (((diario_libros.impdebsol)<>0) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((diario_libros.idlib)=37)) OR (((diario_libros.impdebdol)<>0) AND ((diario_libros.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (diario_libros.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND ((diario_libros.idlib)=37)) ORDER BY diario_libros.rnumerodoc1"

    ' EJECUTAMOS LA CONSULTA DE DATOS ANEXADOS Y AGREGAMOS LOS REGISTRO DE LA TALA TEMPORAL A LA TABLA con_diario_final
    xCon.Execute xCad
    
    ' ELIMINAMOS LAS OPERACIONES DE CAJA Y BANCOS QUE NO DEBEN DE MOSTRARSE 'ABONO LETRAS COBRANZA    'ABONO LETRAS DESCUENTO SEGUN EL CHATIN
    xCon.Execute "DELETE con_diario_final.abrev From con_diario_final WHERE (((con_diario_final.abrev)='ABONO LETRAS COBRANZA')) OR (((con_diario_final.abrev)='ABONO LETRAS DESCUENTO'))"
    xCon.Execute "DELETE con_diario_final.abrev From con_diario_final WHERE (((con_diario_final.abrev)='CHEQUES GIRADOS')) OR (((con_diario_final.abrev)='ABONO LETRAS DESCUENTO'))"
    
    MsgBox "El proceso termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Exit Sub
    
ExisteTabla:
    If Err.Number = -2147217865 Then      ' EL ERROR DE QUE LA TABLA NO EXISTE
        'MsgBox "Se ha detectado el siguiente error : " & Err.Description
        Resume Next
    'Else
    '    MsgBox "Se ha detectado el siguiente error : " & Err.Description
    End If
End Sub

Sub CargarSelect()
    Dim xCad As String
    Dim A As Integer
    Dim xFchApe As String
    
    xFchApe = "01/01/09"
    Fg1.Rows = 2
    Fg3.Rows = 2
    
On Error GoTo LaCagada

    ' ELIMINAMOS LAS FILAS EN BLANCO
    If Fg2.Rows <> 0 Then
        For A = 0 To Fg2.Rows - 1
            If NulosN(Fg2.TextMatrix(A, 2)) = 0 Then
                Fg2.RemoveItem A
            End If
        Next A
    End If
    
    If Fg2.Rows = 0 Then
        ' MOSTRAMOS TODOS LOS CLIENTE
        xCad = "SELECT con_diario_final.* From con_diario_final Where (((con_diario_final.nombre) Is Not Null)) " _
            & " ORDER BY con_diario_final.nombre, con_diario_final.rnumerodoc1, con_diario_final.fchdoc"
    Else
        ' MOSTRAMOS SOLOS LOS CLIENTES ESPECIFICADOS
        Dim xCadWhere As String
        xCadWhere = "("
        For A = 0 To Fg2.Rows - 1
            xCadWhere = xCadWhere & "(con_diario_final.idper = " & Fg2.TextMatrix(A, 2) & ")"
            If A = Fg2.Rows - 1 Then
                Exit For
            End If
            xCadWhere = xCadWhere & " OR "
        Next A
        xCadWhere = xCadWhere & ")"
        
        xCad = "SELECT con_diario_final.* From con_diario_final Where (((con_diario_final.nombre) Is Not Null) " _
        & " And " & xCadWhere & ")" _
        & " ORDER BY con_diario_final.nombre, con_diario_final.rnumerodoc1, con_diario_final.fchdoc"
    End If

    RST_Busq RstDet, xCad, xCon
    
    ' MOSTRAMOS SOLO EL APERTURA
    If Check1.Value = 1 And Check2.Value = 0 Then
        RstDet.Filter = "fchdoc < '" & xFchApe & "'"
    End If
    
    ' MOSTRAMOS SOLO LOS MOVIMIENTOS DEL PERIODO ACTUAL
    If Check1.Value = 0 And Check2.Value = 1 Then
        RstDet.Filter = "fchdoc >= '" & TxtFchIni.Valor & "'"
    End If
    
    ' MOSTRAMOS EL DETALLE DEL MOVIMIENTO
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        ImprimirDetalleFG2
    End If
    
    ' MOSTRAMOS EL RESUMEN
    Set RstDet = Nothing
    If Fg2.Rows = 0 Then
        If Check1.Value = 1 And Check2.Value = 1 Then
            ' MOSTRAMOS TODOS LOS REGISTROS
            'xCad = "SELECT con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1, Sum(con_diario_final.impdebsol) AS SumaDeimpdebsol, " _
                & " Sum(con_diario_final.imphabsol) AS SumaDeimphabsol, Sum(con_diario_final.impdebdol) AS SumaDeimpdebdol, Sum(con_diario_final.imphabdol) AS SumaDeimphabdol " _
                & " From con_diario_final GROUP BY con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1 Having (((con_diario_final.nombre) Is Not Null)) " _
                & " ORDER BY con_diario_final.nombre"
                
            xCad = "SELECT con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1, Sum(IIf([idmon]=1,[impdebsol],0)) AS impdebsol2, " _
                & " Sum(IIf([idmon]=1,[imphabsol],0)) AS imphabsol2, Sum(con_diario_final.impdebdol) AS SumaDeimpdebdol, Sum(con_diario_final.imphabdol) AS SumaDeimphabdol" _
                & " From con_diario_final GROUP BY con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1 Having (((con_diario_final.nombre) Is Not Null)) " _
                & " ORDER BY con_diario_final.nombre"

        End If
        
        If Check1.Value = 1 And Check2.Value = 0 Then
            ' MOSTRAMOS SOLO EL APERTURA
            'xCad = "SELECT con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1, Sum(con_diario_final.impdebsol) AS SumaDeimpdebsol, " _
                & " Sum(con_diario_final.imphabsol) AS SumaDeimphabsol, Sum(con_diario_final.impdebdol) AS SumaDeimpdebdol, Sum(con_diario_final.imphabdol) AS SumaDeimphabdol " _
                & " From con_diario_final WHERE (((con_diario_final.fchdoc) < CDate('" & xFchApe & "'))) " _
                & " GROUP BY con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1 " _
                & " Having (((con_diario_final.nombre) Is Not Null)) ORDER BY con_diario_final.nombre"
                
             xCad = "SELECT con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1, Sum(IIf([idmon]=1,[impdebsol],0)) AS impdebsol2, " _
                & " Sum(IIf([idmon]=1,[imphabsol],0)) AS imphabsol2, Sum(con_diario_final.impdebdol) AS SumaDeimpdebdol, Sum(con_diario_final.imphabdol) AS SumaDeimphabdol " _
                & " From con_diario_final WHERE (((con_diario_final.fchdoc)<CDate('" & xFchApe & "'))) GROUP BY con_diario_final.idper, con_diario_final.nombre, " _
                & " con_diario_final.rnumerodoc1 Having (((con_diario_final.nombre) Is Not Null)) ORDER BY con_diario_final.nombre"
   
                
        End If
    
        If Check1.Value = 0 And Check2.Value = 1 Then
            ' MOSTRAMOS SOLO LOS MOVIMIENTOS DEL PERIODO
            'xCad = "SELECT con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1, Sum(con_diario_final.impdebsol) AS SumaDeimpdebsol, " _
                & " Sum(con_diario_final.imphabsol) AS SumaDeimphabsol, Sum(con_diario_final.impdebdol) AS SumaDeimpdebdol, Sum(con_diario_final.imphabdol) AS SumaDeimphabdol " _
                & " From con_diario_final WHERE (((con_diario_final.fchdoc) > CDate('01/01/09'))) " _
                & " GROUP BY con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1 " _
                & " Having (((con_diario_final.nombre) Is Not Null)) ORDER BY con_diario_final.nombre"
                
            xCad = "SELECT con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1, Sum(IIf([idmon]=1,[impdebsol],0)) AS impdebsol2, " _
                & " Sum(IIf([idmon]=1,[imphabsol],0)) AS imphabsol2, Sum(con_diario_final.impdebdol) AS SumaDeimpdebdol, Sum(con_diario_final.imphabdol) AS SumaDeimphabdol " _
                & " From con_diario_final WHERE (((con_diario_final.fchdoc)>CDate('" & xFchApe & "'))) GROUP BY con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1 " _
                & " Having (((con_diario_final.nombre) Is Not Null)) ORDER BY con_diario_final.nombre"

        End If
    Else
        If Check1.Value = 1 And Check2.Value = 1 Then
            ' MOSTRAMOS TODOS LOS REGISTROS
            'xCad = "SELECT con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1, Sum(con_diario_final.impdebsol) AS SumaDeimpdebsol, " _
                & " Sum(con_diario_final.imphabsol) AS SumaDeimphabsol, Sum(con_diario_final.impdebdol) AS SumaDeimpdebdol, Sum(con_diario_final.imphabdol) AS SumaDeimphabdol " _
                & " From con_diario_final GROUP BY con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1 " _
                & " Having (" & xCadWhere & " And ((con_diario_final.nombre) Is Not Null)) " _
                & " ORDER BY con_diario_final.nombre"
            xCad = "SELECT con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1, Sum(IIf([idmon]=1,[impdebsol],0)) AS impdebsol2, " _
                & " Sum(IIf([idmon]=1,[imphabsol],0)) AS imphabsol2, Sum(con_diario_final.impdebdol) AS SumaDeimpdebdol, Sum(con_diario_final.imphabdol) AS SumaDeimphabdol " _
                & " From con_diario_final GROUP BY con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1 " _
                & " Having (" & xCadWhere & " And ((con_diario_final.nombre) Is Not Null)) " _
                & " ORDER BY con_diario_final.nombre"
        End If
        
        If Check1.Value = 1 And Check2.Value = 0 Then
            ' MOSTRAMOS SOLO EL APERTURA
            'xCad = "SELECT con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1, Sum(con_diario_final.impdebsol) AS SumaDeimpdebsol, " _
                & " Sum(con_diario_final.imphabsol) AS SumaDeimphabsol, Sum(con_diario_final.impdebdol) AS SumaDeimpdebdol, Sum(con_diario_final.imphabdol) AS SumaDeimphabdol " _
                & " From con_diario_final WHERE (((con_diario_final.fchdoc) < CDate('" & xFchApe & "')))" _
               & " GROUP BY con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1 " _
                & " Having (" & xCadWhere & " And ((con_diario_final.nombre) Is Not Null)) " _
                & " ORDER BY con_diario_final.nombre"
                
            xCad = "SELECT con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1, Sum(IIf([idmon]=1,[impdebsol],0)) AS impdebsol2, " _
                & " Sum(IIf([idmon]=1,[imphabsol],0)) AS imphabsol2, Sum(con_diario_final.impdebdol) AS SumaDeimpdebdol, Sum(con_diario_final.imphabdol) AS SumaDeimphabdol " _
                & " From con_diario_final WHERE (((con_diario_final.fchdoc)<CDate('" & xFchApe & "'))) GROUP BY con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1 " _
                & " Having (" & xCadWhere & " And ((con_diario_final.nombre) Is Not Null)) " _
                & " ORDER BY con_diario_final.nombre"

        End If
    
        If Check1.Value = 0 And Check2.Value = 1 Then
            ' MOSTRAMOS SOLO LOS MOVIMIENTOS DEL PERIODO
            'xCad = "SELECT con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1, Sum(con_diario_final.impdebsol) AS SumaDeimpdebsol, " _
                & " Sum(con_diario_final.imphabsol) AS SumaDeimphabsol, Sum(con_diario_final.impdebdol) AS SumaDeimpdebdol, Sum(con_diario_final.imphabdol) AS SumaDeimphabdol " _
                & " From con_diario_final WHERE (((con_diario_final.fchdoc) >= CDate('" & TxtFchIni.Valor & "')))" _
                & " GROUP BY con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1 " _
                & " Having (" & xCadWhere & " And ((con_diario_final.nombre) Is Not Null)) " _
                & " ORDER BY con_diario_final.nombre"
                
            xCad = "SELECT con_diario_final.idper, con_diario_final.nombre, con_diario_final.rnumerodoc1, Sum(IIf([idmon]=1,[impdebsol],0)) AS impdebsol2, " _
                & " Sum(IIf([idmon]=1,[imphabsol],0)) AS imphabsol2, Sum(con_diario_final.impdebdol) AS SumaDeimpdebdol, Sum(con_diario_final.imphabdol) AS SumaDeimphabdol " _
                & " From con_diario_final WHERE (((con_diario_final.fchdoc)>=CDate('" & TxtFchIni.Valor & "'))) GROUP BY con_diario_final.idper, con_diario_final.nombre, " _
                & " con_diario_final.rnumerodoc1 " _
                & " Having (" & xCadWhere & " And ((con_diario_final.nombre) Is Not Null)) " _
                & " ORDER BY con_diario_final.nombre"

        End If
    End If
    
    RST_Busq RstDet, xCad, xCon
    
    If RstDet.RecordCount <> 0 Then
        MostrarDetalle
    End If
    
    Exit Sub

LaCagada:
    MsgBox "Error Inesperado sucedio lo siguiente : " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Resume
End Sub

Sub MostrarDetalle()
    Dim A As Double
    Dim xIdCli As Integer
    
    Fg1.Rows = 2
    xIdCli = RstDet("idper")
    
    For A = 1 To RstDet.RecordCount
        Fg1.Rows = Fg1.Rows + 1
        If RstDet("rnumerodoc1") = "118102008011655" Then
            MsgBox ""
        End If
        
        If A = 1 Then
            GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 2, "CLIENTE : " & RstDet("nombre"), flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
            Fg1.Rows = Fg1.Rows + 1
        End If
        
        If RstDet("idper") <> xIdCli Then
            Fg1.Rows = Fg1.Rows + 1
            GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 2, "CLIENTE : " & RstDet("nombre"), flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
            xIdCli = RstDet("idper")
            Fg1.Rows = Fg1.Rows + 1
        End If
        
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = "ORDEN DE DESPACHO"
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDet("rnumerodoc1"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format((NulosN(RstDet("SumaDeimphabdol")) - NulosN(RstDet("SumaDeimpdebdol"))), "0.00")
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format((NulosN(RstDet("imphabsol2")) - NulosN(RstDet("impdebsol2"))), "0.00")
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(RstDet("nombre"))
        RstDet.MoveNext
    Next A
End Sub

Sub ImprimirDetalleFG2()
    Dim B As Double
    Dim xNumOrd As String
    Dim xIdCliente As Integer
    Dim xForeColor As Long
    Dim xTotDebSol, xTotHabSol, xTotDebDol, xTotHabDol As Double
    Dim xxTotalSol, xxTotalDol As Double
    Dim xRucCli As String
    
On Error GoTo LaCagada2

    Fg3.Rows = 2
    If RstDet.RecordCount <> 0 Then
        ProgressBar1.Max = RstDet.RecordCount
        Frame3.Left = ((Me.Width - Frame3.Width) / 2)
        Frame3.Top = ((Me.Height - Frame3.Height) / 2)
        Frame3.Visible = True
        Frame3.Refresh
        
        xNumOrd = NulosC(RstDet("rnumerodoc1"))
        xIdCliente = NulosN(RstDet("idper"))
        xRucCli = Busca_Codigo(xIdCliente, "id", "numruc", "mae_cliente", "N", xCon)
        
        For B = 1 To RstDet.RecordCount
            ProgressBar1.Value = B
            
'            If Mid(RstDet("nombre"), 1, 10) = "CAJA RURAL" Then
'                MsgBox ""
'            End If
            
            If B = 1 Then
                Fg3.Rows = Fg3.Rows + 1
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, 2, "CLIENTE : " & RstDet("nombre"), flexAlignLeftCenter, True, , &HFF0000, &HE2FEFB, True
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 9, "Nº R.U.C. : " & xRucCli, flexAlignLeftCenter, True, , &HFF0000, &HE2FEFB, True
                
                Fg3.Rows = Fg3.Rows + 1
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, 2, "DOC. REF. :  ORDEN DE DESPACHO  ", flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 9, "Nº DOC. REF. : " & NulosC(RstDet("rnumerodoc1")), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                
                Fg3.Rows = Fg3.Rows + 1
            Else
                If (NulosC(RstDet("rnumerodoc1")) <> xNumOrd) Or (NulosC(RstDet("rnumerodoc1")) = xNumOrd And NulosC(RstDet("idper")) <> xIdCliente) Then
                    Fg3.Rows = Fg3.Rows + 1
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 8, "TOTAL ==>", flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 11, &H0&, True, &HE2FEFB, Format(NulosN(xTotDebDol), "0.00")
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 12, &H0&, True, &HE2FEFB, Format(NulosN(xTotHabDol), "0.00")
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 13, &H0&, True, &HE2FEFB, Format(NulosN(xTotDebSol), "0.00")
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 14, &H0&, True, &HE2FEFB, Format(NulosN(xTotHabSol), "0.00")
                    
                    Fg3.Rows = Fg3.Rows + 1
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 8, "SALDO ==>", flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
                    If (xTotDebDol - xTotHabDol) < 0 Then xForeColor = &HFF0000  'azul
                    If (xTotDebDol - xTotHabDol) > 0 Then xForeColor = &HFF&     'rojo
                    If (xTotDebDol - xTotHabDol) = 0 Then xForeColor = &H0&      'negro
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 12, xForeColor, True, &HE2FEFB, Format(xTotHabDol - xTotDebDol, "0.00")
                    
                    If (xTotDebSol - xTotHabSol) < 0 Then xForeColor = &HFF0000  'azul
                    If (xTotDebSol - xTotHabSol) > 0 Then xForeColor = &HFF&     'rojo
                    If (xTotDebSol - xTotHabSol) = 0 Then xForeColor = &H0&      'negro
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 14, xForeColor, True, &HE2FEFB, Format(xTotHabSol - xTotDebSol, "0.00")
                    
                    xxTotalDol = xxTotalDol + (xTotHabDol - xTotDebDol)
                    xxTotalSol = xxTotalSol + (xTotHabSol - xTotDebSol)
                    
                    xTotDebDol = 0:    xTotHabDol = 0:     xTotDebSol = 0:     xTotHabSol = 0
                    Fg3.Rows = Fg3.Rows + 2
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, 2, "DOC. REF. :  ORDEN DE DESPACHO  ", flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 10, "Nº DOC. REF. : " & NulosC(RstDet("rnumerodoc1")), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                    
                    xNumOrd = NulosC(RstDet("rnumerodoc1"))
                Else
                
                End If
                
                If NulosC(RstDet("idper")) <> xIdCliente Then
                    Fg3.Rows = Fg3.Rows - 1
                    ' IMPRIMIMOS EL TOTAL POR EMPRESA
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 8, "TOTAL EMPRESA ==>", flexAlignLeftTop, True, , &H40&, &HE2FEFB, True
                    
                    If (xxTotalDol) > 0 Then xForeColor = &HFF0000  'azul
                    If (xxTotalDol) < 0 Then xForeColor = &HFF&     'rojo
                    If (xxTotalDol) = 0 Then xForeColor = &H0&      'negro
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 12, xForeColor, True, &HE2FEFB, Format(xxTotalDol, "0.00")
                    
                    If (xxTotalSol) > 0 Then xForeColor = &HFF0000  'azul
                    If (xxTotalSol) < 0 Then xForeColor = &HFF&     'rojo
                    If (xxTotalSol) = 0 Then xForeColor = &H0&      'negro
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 14, xForeColor, True, &HE2FEFB, Format(xxTotalSol, "0.00")

                    xxTotalDol = 0
                    xxTotalSol = 0
                    Fg3.Rows = Fg3.Rows + 3
                    
                    xIdCliente = NulosN(RstDet("idper"))
                    xRucCli = Busca_Codigo(xIdCliente, "id", "numruc", "mae_cliente", "N", xCon)
                    
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, 2, "CLIENTE : " & RstDet("nombre"), flexAlignLeftCenter, True, , &HFF0000, &HE2FEFB, True
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 9, "Nº R.U.C. : " & xRucCli, flexAlignLeftCenter, True, , &HFF0000, &HE2FEFB, True
                    xIdCliente = RstDet("idper")
                    Fg3.Rows = Fg3.Rows + 1
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, 2, "DOC. REF. :  ORDEN DE DESPACHO  ", flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 10, "Nº DOC. REF. : " & NulosC(RstDet("rnumerodoc1")), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                    
                End If
                
                Fg3.Rows = Fg3.Rows + 1
            End If
            
            Fg3.TextMatrix(Fg3.Rows - 1, 1) = NulosC(RstDet("cuenta"))
            Fg3.TextMatrix(Fg3.Rows - 1, 2) = NulosC(RstDet("descripcion"))
            Fg3.TextMatrix(Fg3.Rows - 1, 3) = RstDet("numreg")
            Fg3.TextMatrix(Fg3.Rows - 1, 4) = NulosC(RstDet("rnumerodoc1"))
            
            Fg3.TextMatrix(Fg3.Rows - 1, 5) = NulosC(RstDet("fchdoc"))
            Fg3.TextMatrix(Fg3.Rows - 1, 6) = NulosC(RstDet("abrev"))
            Fg3.TextMatrix(Fg3.Rows - 1, 7) = NulosC(RstDet("rnumerodoc"))
            Fg3.TextMatrix(Fg3.Rows - 1, 15) = NulosC(RstDet("nombre"))
            
            If RstDet("idmon") = 1 Then
                Fg3.TextMatrix(Fg3.Rows - 1, 8) = "S/."
            Else
                Fg3.TextMatrix(Fg3.Rows - 1, 8) = "US $"
            End If
            Fg3.TextMatrix(Fg3.Rows - 1, 9) = Format(NulosN(RstDet("tc")), "0.000")
            Fg3.TextMatrix(Fg3.Rows - 1, 10) = "0.00"
            
            If RstDet("idmon") = 2 Then
                Fg3.TextMatrix(Fg3.Rows - 1, 11) = Format(NulosN(RstDet("impdebdol")), "0.00")
                Fg3.TextMatrix(Fg3.Rows - 1, 12) = Format(NulosN(RstDet("imphabdol")), "0.00")
                Fg3.TextMatrix(Fg3.Rows - 1, 13) = 0
                Fg3.TextMatrix(Fg3.Rows - 1, 14) = 0
                xTotDebDol = xTotDebDol + NulosN(RstDet("impdebdol"))
                xTotHabDol = xTotHabDol + NulosN(RstDet("imphabdol"))
                xTotDebSol = xTotDebSol + 0
                xTotHabSol = xTotHabSol + 0
            Else
                Fg3.TextMatrix(Fg3.Rows - 1, 11) = 0
                Fg3.TextMatrix(Fg3.Rows - 1, 12) = 0
                Fg3.TextMatrix(Fg3.Rows - 1, 13) = Format(NulosN(RstDet("impdebsol")), "0.00")
                Fg3.TextMatrix(Fg3.Rows - 1, 14) = Format(NulosN(RstDet("imphabsol")), "0.00")
                
                xTotDebDol = xTotDebDol + 0
                xTotHabDol = xTotHabDol + 0
                xTotDebSol = xTotDebSol + NulosN(RstDet("impdebsol"))
                xTotHabSol = xTotHabSol + NulosN(RstDet("imphabsol"))
            End If
            
            RstDet.MoveNext
            If RstDet.EOF = True Then
                Fg3.Rows = Fg3.Rows + 1
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 8, "TOTAL ==>", flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 11, &H0&, True, &HE2FEFB, Format(NulosN(xTotDebDol), "0.00")
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 12, &H0&, True, &HE2FEFB, Format(NulosN(xTotHabDol), "0.00")
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 13, &H0&, True, &HE2FEFB, Format(NulosN(xTotDebSol), "0.00")
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 14, &H0&, True, &HE2FEFB, Format(NulosN(xTotHabSol), "0.00")
                
                Fg3.Rows = Fg3.Rows + 1
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 8, "SALDO ==>", flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
                
                If (xTotDebDol - xTotHabDol) > 0 Then xForeColor = &HFF&    'rojo
                If (xTotDebDol - xTotHabDol) < 0 Then xForeColor = &HFF0000 'azul
                If (xTotDebDol - xTotHabDol) = 0 Then xForeColor = &H0&     'negro
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 12, xForeColor, True, &HE2FEFB, Format(xTotHabDol - xTotDebDol, "0.00")
                
                If (xTotDebSol - xTotHabSol) > 0 Then xForeColor = &HFF&    'rojo
                If (xTotDebSol - xTotHabSol) < 0 Then xForeColor = &HFF0000 'azul
                If (xTotDebSol - xTotHabSol) = 0 Then xForeColor = &H0&     'negro
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 14, xForeColor, True, &HE2FEFB, Format(xTotHabSol - xTotDebSol, "0.00")
                
                xxTotalDol = xxTotalDol + (xTotDebDol - xTotHabDol)
                xxTotalSol = xxTotalSol + (xTotDebSol - xTotHabSol)
                
                Fg3.Rows = Fg3.Rows + 1
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 8, "TOTAL EMPRESA ==>", flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
                If (xxTotalDol < 0) Then xForeColor = &HFF&    'rojo
                If (xxTotalDol > 0) Then xForeColor = &HFF0000 'azul
                If (xxTotalDol = 0) Then xForeColor = &H0&     'negro
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 12, xForeColor, True, &HE2FEFB, Format(xxTotalDol, "0.00")
                
                If (xxTotalSol < 0) Then xForeColor = &HFF&    'rojo
                If (xxTotalSol > 0) Then xForeColor = &HFF0000 'azul
                If (xxTotalSol = 0) Then xForeColor = &H0&     'negro
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 14, xForeColor, True, &HE2FEFB, Format(xxTotalSol, "0.00")
                                
                Fg3.Rows = Fg3.Rows + 1
                Exit For
            End If
        Next B
    End If
    Frame3.Visible = False
    Frame3.Refresh
    Exit Sub

LaCagada2:
    MsgBox "En el Procedimiento ImprimirDetalleFG2 se produjo el siguiente error : " & Err.Description
    Resume Next
End Sub

