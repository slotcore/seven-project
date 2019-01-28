VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "Codejock.CommandBars.v12.0.0.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.ocx"
Begin VB.Form FrmAnalisisCliente 
   Caption         =   "Contabilidad - Analisis de Cuenta x Cliente"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14385
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   14385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   885
      Left            =   840
      TabIndex        =   30
      Top             =   2685
      Visible         =   0   'False
      Width           =   6180
      Begin XtremeSuiteControls.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   135
         TabIndex        =   31
         Top             =   465
         Width           =   5925
         _Version        =   786432
         _ExtentX        =   10451
         _ExtentY        =   503
         _StockProps     =   93
         Appearance      =   6
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
         TabIndex        =   32
         Top             =   90
         Width           =   1845
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
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   840
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
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   6150
         Y1              =   15
         Y2              =   15
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
   End
   Begin SizerOneLibCtl.ElasticOne EO 
      Height          =   5775
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
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
      _GridInfo       =   $"FrmAnalisisCliente.frx":0000
      Begin SizerOneLibCtl.TabOne TabOne1 
         Height          =   4575
         Left            =   45
         TabIndex        =   14
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
         Begin VB.Frame FramResumen 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   4155
            Left            =   45
            TabIndex        =   17
            Top             =   45
            Width           =   14175
            Begin VSFlex7Ctl.VSFlexGrid Fg3 
               Height          =   4125
               Left            =   0
               TabIndex        =   18
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
               FormatString    =   $"FrmAnalisisCliente.frx":0042
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
         Begin VB.Frame FrameDetalle 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   4155
            Left            =   -14820
            TabIndex        =   15
            Top             =   45
            Width           =   14175
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   4110
               Left            =   -105
               TabIndex        =   16
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
               FormatString    =   $"FrmAnalisisCliente.frx":0118
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
         TabIndex        =   5
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
         _GridInfo       =   $"FrmAnalisisCliente.frx":01ED
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   1065
            Left            =   11280
            TabIndex        =   21
            Top             =   0
            Width           =   2985
            Begin VB.CommandButton Command3 
               Caption         =   "Command3"
               Height          =   720
               Left            =   2250
               TabIndex        =   25
               Top             =   255
               Visible         =   0   'False
               Width           =   585
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Eliminar OD No Marcadas"
               Height          =   360
               Left            =   105
               TabIndex        =   24
               Top             =   540
               Visible         =   0   'False
               Width           =   2070
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Eliminar OD Seleccionada"
               Height          =   360
               Left            =   105
               TabIndex        =   23
               Top             =   180
               Visible         =   0   'False
               Width           =   2070
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
               TabIndex        =   22
               Top             =   30
               Visible         =   0   'False
               Width           =   1080
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   1065
            Left            =   7875
            TabIndex        =   19
            Top             =   0
            Width           =   3390
            Begin VB.Frame Frame66 
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   720
               Left            =   1755
               TabIndex        =   33
               Top             =   285
               Width           =   1575
               Begin VB.CheckBox Check2 
                  Caption         =   "Año Trabajo"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   35
                  Top             =   270
                  Width           =   1290
               End
               Begin VB.CheckBox Check1 
                  Caption         =   "Apertura"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   34
                  Top             =   45
                  Width           =   1170
               End
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Todos"
               Height          =   195
               Left            =   210
               TabIndex        =   29
               Top             =   780
               Width           =   1275
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Cancelados"
               Height          =   195
               Left            =   210
               TabIndex        =   28
               Top             =   555
               Width           =   1275
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Pendientes"
               Height          =   195
               Left            =   210
               TabIndex        =   27
               Top             =   330
               Width           =   1275
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
               TabIndex        =   20
               Top             =   30
               Width           =   810
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   1065
            Left            =   3375
            TabIndex        =   10
            Top             =   0
            Width           =   4485
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   750
               Left            =   0
               TabIndex        =   3
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
               FormatString    =   $"FrmAnalisisCliente.frx":024B
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
               TabIndex        =   11
               Top             =   30
               Width           =   600
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   1065
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   3360
            Begin VB.CommandButton CmdBusMon 
               Enabled         =   0   'False
               Height          =   240
               Left            =   1230
               Picture         =   "FrmAnalisisCliente.frx":029B
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   720
               Width           =   240
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   870
               Locked          =   -1  'True
               TabIndex        =   2
               Text            =   "TxtIdMon"
               Top             =   690
               Width           =   615
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   870
               TabIndex        =   0
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
               TabIndex        =   1
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
               TabIndex        =   13
               Top             =   690
               Width           =   1770
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Left            =   45
               TabIndex        =   9
               Top             =   720
               Width           =   585
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Venc."
               Height          =   195
               Left            =   45
               TabIndex        =   8
               Top             =   420
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Inicio"
               Height          =   195
               Left            =   45
               TabIndex        =   7
               Top             =   135
               Width           =   735
            End
         End
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   750
      TabIndex        =   26
      Top             =   210
      Width           =   1290
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   4185
      Top             =   75
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmAnalisisCliente.frx":03CD
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   3765
      Top             =   45
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "FrmAnalisisCliente"
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

Dim RstRes1 As New ADODB.Recordset
Dim RstRes2 As New ADODB.Recordset
Dim RstDet1 As New ADODB.Recordset
Dim RstDet2 As New ADODB.Recordset

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
    If NulosC(Fg1.TextMatrix(Fg1.Row, 8)) <> "Nº DE DOC. ==>" Then
        MsgBox "No ha seleccionado el numero de orden de despacho que se desea eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    EliminarOrden Fg1.TextMatrix(Fg1.Row, 9)
    MsgBox "La Ordden de Depacho se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
    Dim A As Integer
    For A = 2 To Fg1.Rows - 1
        If NulosC(Fg1.TextMatrix(A, 8)) = "Nº DE DOC. ==>" Then
            If NulosN(Fg1.TextMatrix(A, Fg1.Cols - 1)) = 0 Then
                EliminarOrden Fg1.TextMatrix(A, 9)
            End If
        End If
    Next A

    MsgBox "Las Orddene de Depacho se eliminaron con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Sub

Private Sub Command3_Click()
    'Frame3.Visible = True
    CargarSelect
'    FrmImportaCtaCte.LblIdCliente.Caption = Fg2.TextMatrix(Fg2.Row, 2)
'    FrmImportaCtaCte.Show
End Sub

Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Id = 1 Then
        'CargarDatos
        CargarSelect
        'CargarResumen
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

'Sub CargaDatosFormatoSavar()
'
'    If TxtFchIni.Valor = "" Then
'        MsgBox "No ha especificado la fecha de inicio ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtFchIni.SetFocus
'        Exit Sub
'    End If
'
'    If TxtFchFin.Valor = "" Then
'        MsgBox "No ha especificado la fecha final ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtFchFin.SetFocus
'        Exit Sub
'    End If
'
'    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
'        MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Exit Sub
'    End If
'
'    SetearCuadricula Fg1, 5, xCon, 5, 1, False
'    Fg1.Rows = 2
'
'    Dim RstDat As New ADODB.Recordset
'    Dim RstCli As New ADODB.Recordset
'    Dim xCadWhere As String
'    Dim xSQL2 As String
'    Dim xSql As String
'    Dim A, B, C As Integer
'    Dim xSaldoD, xSaldoS, xTotalSaldoDol, xTotalSaldoSol, xTotEmpSol, xTotEmpDol As Double
'    Dim xColor As Long
'
'    Fg1.MergeCells = flexMergeFixedOnly
'
'    ' ELIMINAMOS LAS FILAS EN BLANCO
'    For A = 0 To Fg2.Rows - 1
'        If NulosN(Fg2.TextMatrix(A, 2)) = 0 Then
'            Fg2.RemoveItem A
'        End If
'    Next A
'
'    'ARMAMOS LAS SENTENCIA WHERE PARA LAS CONSULTAS
'    xCadWhere = ""
'    If Fg2.Rows <> 0 Then
'        xCadWhere = "WHERE ("
'        For A = 0 To Fg2.Rows - 1
'            xCadWhere = xCadWhere & "(consulta3.idcli = " & NulosC(Fg2.TextMatrix(A, 2)) & ")"
'            If A = Fg2.Rows - 1 Then
'                xCadWhere = xCadWhere & ")"
'                Exit For
'            End If
'            xCadWhere = xCadWhere & " OR "
'        Next A
'    End If
'
'    ' OBTEMOS TODOS LOS DATOS PARA MOSTRAR
'
'    xSql = xSql & "SELECT consulta3.* FROM " _
'        & " ( "
'
'    xSql = xSql & "SELECT con_planctas.cuenta, con_planctas.descripcion, vta_ventas.fchreg, mae_cliente.numruc, mae_cliente.nombre, " _
'        & " mae_documento.abrev, vta_ventas.fchdoc, vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc, mae_moneda.simbolo, " _
'        & " IIf(vta_ventas!tc=0,con_tc!impven,vta_ventas!tc) AS tc, vta_ventas.imptotdoc, vta_ventas.numerodocref, " _
'        & " IIf([vta_ventas]![idmon]=1,0,[con_diario]![impdebdol]) AS impdebdol, IIf([vta_ventas]![idmon]=1,0,[con_diario]![imphabdol]) AS imphabdol, " _
'        & " IIf([vta_ventas]![idmon]=2,0,[con_diario]![impdebsol]) AS impdebsol, IIf([vta_ventas]![idmon]=2,0,[con_diario]![imphabsol]) AS imphabsol, " _
'        & " vta_ventas.idcli FROM (((con_diario LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) RIGHT JOIN (mae_cliente " _
'        & " RIGHT JOIN (vta_ventas LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) ON mae_cliente.id = vta_ventas.idcli) " _
'        & " ON con_diario.idmov = vta_ventas.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc " _
'        & " ON vta_ventas.fchdoc = con_tc.fecha WHERE (((con_planctas.cuenta) Like '12%') AND ((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
'        & " And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=2) AND ((vta_ventas.anulado)=0)) " _
'        & " OR (((con_diario.idlib)=3) AND ((vta_ventas.idmes)=0))" _
'        & " Union  " _
'        & " SELECT con_planctas.cuenta, con_planctas.descripcion, vta_gastodebito.fchreg, mae_cliente.numruc, mae_cliente.nombre, " _
'        & " mae_documento.abrev, vta_gastodebito.fchemi AS fchdoc, vta_gastodebito!numser+'-'+vta_gastodebito!numdoc AS numdoc, " _
'        & " mae_moneda.simbolo, IIf(vta_gastodebito!tc=0,con_tc!impven,vta_gastodebito!tc) AS tc, vta_gastodebito.imptotdoc AS imptotdoc, " _
'        & " vta_gastodebito.numerodocref, IIf([vta_gastodebito]![idmon]=1,0,[vta_gastodebito]![imptotdoc]) AS impdebdol, " _
'        & " con_diario.imphabdol, IIf([vta_gastodebito]![idmon]=2,0,[con_diario]![impdebsol]) AS impdebsol, con_diario.imphabsol, " _
'        & " vta_gastodebito.idcli FROM ((((((con_diario LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) RIGHT JOIN vta_gastodebito " _
'        & " ON con_diario.idmov = vta_gastodebito.id) LEFT JOIN mae_cliente ON vta_gastodebito.idcli = mae_cliente.id) LEFT JOIN mae_documento " _
'        & " ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id) LEFT JOIN mae_libros " _
'        & " ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha " _
'        & " WHERE (((vta_gastodebito.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_gastodebito.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
'        & " AND ((con_diario.idlib)=41) AND ((con_diario.impdebsol)<>0)) OR (((con_diario.idlib)=3) AND ((vta_gastodebito.idmes)=0) " _
'        & " AND ((con_diario.impdebsol)<>0))"
'
'    xSql = xSql & " UNION " _
'        & " SELECT con_planctas.cuenta, con_planctas.descripcion, let_letra.fchreg, mae_cliente.numruc, mae_cliente.nombre, " _
'        & " mae_documento.abrev, let_letradet.fchemi, let_letradet!numser & '-' & let_letradet!numdoc AS numdoc, mae_moneda.simbolo, " _
'        & " IIf(let_letra!tc=0,con_tc!impven,let_letra!tc) AS tc, let_letradet.implet, let_letra!idaduana & let_letra!idregimen & " _
'        & " let_letra!anoorden & let_letra!numorden AS numerodocref, IIf([let_letra]![idmon]=1,0,[con_diario].[impdebdol]) AS impdebdol, " _
'        & " IIf([let_letra]![idmon]=1,0,[con_diario]![imphabdol]) AS imphabdol, IIf([let_letra]![idmon]=2,0,[con_diario].[impdebsol]) AS impdebsol, " _
'        & " IIf([let_letra]![idmon]=2,0,[con_diario]![imphabsol]) AS imphabsol, let_letra.idclipro FROM ((((let_letra LEFT JOIN mae_cliente " _
'        & " ON let_letra.idclipro = mae_cliente.id) LEFT JOIN mae_documento ON let_letra.idtipdoc = mae_documento.id) LEFT JOIN mae_moneda " _
'        & " ON let_letra.idmon = mae_moneda.id) LEFT JOIN con_tc ON let_letra.fchemi = con_tc.fecha) LEFT JOIN ((con_diario " _
'        & " RIGHT JOIN let_letradet ON (con_diario.correlativo = let_letradet.corr) AND (con_diario.idmov = let_letradet.idlet)) " _
'        & " LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) ON let_letra.id = let_letradet.idlet " _
'        & " WHERE (((let_letra.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (let_letra.fchreg)<=CDate('" & TxtFchFin.Valor & "'))) ORDER BY mae_cliente.nombre" _
'        & " ) AS consulta3 "
'
'    If xCadWhere <> "" Then
'        xSql = xSql & " " & xCadWhere
'    End If
'
'
'    ' OBTENEMOS LOS CLIENTES UNICOS
'    xSQL2 = xSQL2 & " SELECT DISTINCT Consulta3.numruc, Consulta3.nombre, Consulta3.idcli From " _
'        & " ( "
'
'    xSQL2 = xSQL2 & " SELECT con_planctas.cuenta, con_planctas.descripcion, vta_ventas.fchreg, mae_cliente.numruc, mae_cliente.nombre, " _
'        & " mae_documento.abrev, vta_ventas.fchdoc, vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc, mae_moneda.simbolo, " _
'        & " IIf(vta_ventas!tc=0,con_tc!impven,vta_ventas!tc) AS tc, vta_ventas.imptotdoc, vta_ventas.numerodocref, " _
'        & " IIf([vta_ventas]![idmon]=1,0,[con_diario]![impdebdol]) AS impdebdol, IIf([vta_ventas]![idmon]=1,0,[con_diario]![imphabdol]) AS imphabdol, " _
'        & " IIf([vta_ventas]![idmon]=2,0,[con_diario]![impdebsol]) AS impdebsol, IIf([vta_ventas]![idmon]=2,0,[con_diario]![imphabsol]) AS imphabsol, " _
'        & " vta_ventas.idcli FROM (((con_diario LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) RIGHT JOIN (mae_cliente " _
'        & " RIGHT JOIN (vta_ventas LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) ON mae_cliente.id = vta_ventas.idcli) " _
'        & " ON con_diario.idmov = vta_ventas.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc " _
'        & " ON vta_ventas.fchdoc = con_tc.fecha WHERE (((con_planctas.cuenta) Like '12*') AND ((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
'        & " And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=2) AND ((vta_ventas.anulado)=0)) " _
'        & " OR (((con_diario.idlib)=3) AND ((vta_ventas.idmes)=0))" _
'        & " Union " _
'        & " SELECT con_planctas.cuenta, con_planctas.descripcion, vta_gastodebito.fchreg, mae_cliente.numruc, mae_cliente.nombre, " _
'        & " mae_documento.abrev, vta_gastodebito.fchemi AS fchdoc, vta_gastodebito!numser+'-'+vta_gastodebito!numdoc AS numdoc, " _
'        & " mae_moneda.simbolo, IIf(vta_gastodebito!tc=0,con_tc!impven,vta_gastodebito!tc) AS tc, vta_gastodebito.imptotdoc AS imptotdoc, " _
'        & " vta_gastodebito.numerodocref, IIf([vta_gastodebito]![idmon]=1,0,[vta_gastodebito]![imptotdoc]) AS impdebdol, " _
'        & " con_diario.imphabdol, IIf([vta_gastodebito]![idmon]=2,0,[con_diario]![impdebsol]) AS impdebsol, con_diario.imphabsol, " _
'        & " vta_gastodebito.idcli FROM ((((((con_diario LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) RIGHT JOIN " _
'        & " vta_gastodebito ON con_diario.idmov = vta_gastodebito.id) LEFT JOIN mae_cliente ON vta_gastodebito.idcli = mae_cliente.id) " _
'        & " LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id) " _
'        & " LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha " _
'        & " WHERE (((vta_gastodebito.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_gastodebito.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
'        & " AND ((con_diario.idlib)=41) AND ((con_diario.impdebsol)<>0)) OR (((con_diario.idlib)=3) AND ((vta_gastodebito.idmes)=0))"
'
'    xSQL2 = xSQL2 & " Union " _
'        & " SELECT con_planctas.cuenta, con_planctas.descripcion, let_letra.fchreg, mae_cliente.numruc, mae_cliente.nombre, " _
'        & " mae_documento.abrev, let_letradet.fchemi, let_letradet!numser & '-' & let_letradet!numdoc AS numdoc, mae_moneda.simbolo, " _
'        & " IIf(let_letra!tc=0,con_tc!impven,let_letra!tc) AS tc, let_letradet.implet, let_letra!idaduana & let_letra!idregimen & " _
'        & " let_letra!anoorden & let_letra!numorden AS numerodocref, IIf([let_letra]![idmon]=1,0,[con_diario].[impdebdol]) AS impdebdol, " _
'        & " IIf([let_letra]![idmon]=1,0,[con_diario]![imphabdol]) AS imphabdol, IIf([let_letra]![idmon]=2,0,[con_diario].[impdebsol]) AS impdebsol, " _
'        & " IIf([let_letra]![idmon]=2,0,[con_diario]![imphabsol]) AS imphabsol, let_letra.idclipro FROM ((((let_letra LEFT JOIN mae_cliente " _
'        & " ON let_letra.idclipro = mae_cliente.id) LEFT JOIN mae_documento ON let_letra.idtipdoc = mae_documento.id) LEFT JOIN mae_moneda " _
'        & " ON let_letra.idmon = mae_moneda.id) LEFT JOIN con_tc ON let_letra.fchemi = con_tc.fecha) LEFT JOIN ((con_diario RIGHT JOIN let_letradet " _
'        & " ON (con_diario.correlativo = let_letradet.corr) AND (con_diario.idmov = let_letradet.idlet)) LEFT JOIN con_planctas " _
'        & " ON con_diario.idcue = con_planctas.id) ON let_letra.id = let_letradet.idlet WHERE (((let_letra.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
'        & " And (let_letra.fchreg)<=CDate('" & TxtFchFin.Valor & "'))) ) as Consulta3"
'
'    If xCadWhere <> "" Then
'        xSQL2 = xSQL2 & " " & xCadWhere
'    End If
'
'    RST_Busq RstDat, xSql, xCon
'    RST_Busq RstCli, xSQL2, xCon
'
'    TabOne1.CurrTab = 0
'    If RstCli.RecordCount <> 0 Then
'        RstCli.Sort = "nombre"
'        RstCli.MoveFirst
'        For A = 1 To RstCli.RecordCount
'            Fg1.Rows = Fg1.Rows + 1
'            Fg1.TextMatrix(Fg1.Rows - 1, 1) = "CLIENTE   :"
'
'            GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 6, "CLIENTE ==> " & RstCli("nombre"), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
'            GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "Nº DE RUC ==> ", flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
'            GRID_COMBINAR Fg1, Fg1.Rows - 1, 9, Fg1.Rows - 1, 10, RstCli("numruc"), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
'
'            ' FILTRAMOS LOS MOVIMIENTOS QUE TENGA EL CLIENTE ACTUAL
'            RstDat.Filter = adFilterNone
'            RstDat.Filter = "numruc = '" & RstCli("numruc") & "'"
'
'            If RstDat.RecordCount <> 0 Then
'                Dim xNumRef As String
'                Dim xTotDebDol, xTotHabDol, xTotDebSol, xTotHabSol As Double
'
'                ' ORDENAMOS POR NUMERO DE DOCUMENTO DE REFERENCIA Y POR FECHA DE EMISION DEL DOCUMENTO
'                RstDat.Sort = "numerodocref, fchdoc"
'                RstDat.MoveFirst
'
'                xNumRef = NulosC(RstDat("numerodocref"))
'                Fg1.Rows = Fg1.Rows + 1
'
'                GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 6, "DOC. REF. ==> " & "ORDEN DE DESPACHO", flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'                GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "Nº DE DOC. ==> ", flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'                GRID_COMBINAR Fg1, Fg1.Rows - 1, 9, Fg1.Rows - 1, 10, xNumRef, flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'
'                xTotalSaldoDol = 0
'                xTotalSaldoSol = 0
'                For B = 1 To RstDat.RecordCount
'                    If B > 1 Then
'                        If NulosC(xNumRef) = NulosC(RstDat("numerodocref")) Then
'                            Fg1.Rows = Fg1.Rows + 1
'                        Else
'                            Fg1.Rows = Fg1.Rows + 1
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "TOTAL ==>", flexAlignLeftCenter, True, , &H0&, &HE2FEFB, True
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H0&, True, &HE2FEFB, Format(xTotDebDol, "0.00")
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &H0&, True, &HE2FEFB, Format(NulosN(xTotHabDol), "0.00")
'
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H0&, True, &HE2FEFB, Format(xTotDebSol, "0.00")
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H0&, True, &HE2FEFB, Format(xTotHabSol, "0.00")
'
'                            ' ESCRIBIMOS EL SALDO DE CADA MONEDA
'                            Fg1.Rows = Fg1.Rows + 1
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "SALDO ==>", flexAlignLeftCenter, True, , &H0&, &HE2FEFB, True
'
'                            'COLOR PARA LOS DOLARES
'                            xTotEmpDol = xTotEmpDol + (NulosN(xTotDebDol) - NulosN(xTotHabDol))
'                            If NulosN(xTotDebDol) - NulosN(xTotHabDol) > 0 Then
'                                xColor = &HFF0000
'                            Else
'                                xColor = &HFF&
'                            End If
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, xColor, True, &HE2FEFB, Format(NulosN(xTotDebDol) - NulosN(xTotHabDol), "0.00")
'
'                            'COLOR PARA LOS SOLES
'                            xTotEmpSol = xTotEmpSol + (NulosN(xTotDebSol) - NulosN(xTotHabSol))
'                            If NulosN(xTotDebSol) - NulosN(xTotHabSol) > 0 Then
'                                xColor = &HFF0000
'                            Else
'                                xColor = &HFF&
'                            End If
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, xColor, True, &HE2FEFB, Format(xTotDebSol - xTotHabSol, "0.00")
'                            xTotalSaldoDol = 0
'                            xTotalSaldoSol = 0
'
'                            Fg1.Rows = Fg1.Rows + 1
'                            xTotDebDol = 0
'                            xTotHabDol = 0
'                            xTotDebSol = 0
'                            xTotHabSol = 0
'
'                            Fg1.Rows = Fg1.Rows + 1
'                            xNumRef = RstDat("numerodocref")
'
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 6, "DOC. REF. ==> " & "ORDEN DE DESPACHO", flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "Nº DE DOC. ==> ", flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 9, Fg1.Rows - 1, 10, xNumRef, flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'
'                            Fg1.Rows = Fg1.Rows + 1
'                        End If
'                    Else
'                        Fg1.Rows = Fg1.Rows + 1
'                    End If
'
'                    Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstDat("cuenta"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDat("descripcion"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 3) = RstDat("fchdoc")
'                    Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(RstDat("abrev"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(RstDat("numdoc"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = RstDat("simbolo")
'                    Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosN(RstDat("tc"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(RstDat("imptotdoc")), "0.00")
'
'                    If RstDat("abrev") = "LE" Or RstDat("abrev") = "DPA" Or RstDat("abrev") = "DPC" Or RstDat("abrev") = "CDT" Or RstDat("abrev") = "LGC" Or RstDat("abrev") = "DOB" Or RstDat("abrev") = "DE" Or RstDat("abrev") = "CR" Then
'                        'aqui hacemos la inversa del asiento para poder hallar el saldo esto segun carlos
'                        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(RstDat("imphabdol"), "0.00")
'                        Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(RstDat("impdebdol"), "0.00")
'                        'xTotalSaldoDol = (xTotalSaldoDol + RstDat("imphabdol")) - RstDat("impdebdol")
'                        'If xTotalSaldoDol > 0 Then
'                        '    xColor = &HFF0000
'                        'Else
'                        '    xColor = &HFF&
'                        'End If
'                        'FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, xColor, True, &HE2FEFB, Format(xTotalSaldoDol, "0.00")
'
'                        Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(RstDat("imphabsol"), "0.00")
'                        Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(RstDat("impdebsol"), "0.00")
'                        'xTotalSaldoSol = (xTotalSaldoSol + NulosN(RstDat("imphabsol"))) - NulosN(RstDat("impdebsol"))
'                        'FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, xColor, True, &HE2FEFB, Format(xTotalSaldoSol, "0.00")
'
'                        xTotDebDol = xTotDebDol + NulosN(RstDat("imphabdol"))
'                        xTotHabDol = xTotHabDol + NulosN(RstDat("impdebdol"))
'                        xTotDebSol = xTotDebSol + NulosN(RstDat("imphabsol"))
'                        xTotHabSol = xTotHabSol + NulosN(RstDat("impdebsol"))
'
'                    Else
'                        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(RstDat("impdebdol"), "0.00")
'                        Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(RstDat("imphabdol"), "0.00")
'                        xTotalSaldoDol = (xTotalSaldoDol + RstDat("impdebdol")) - RstDat("imphabdol")
''                        If xTotalSaldoDol > 0 Then
''                            xColor = &HFF0000
''                        Else
''                            xColor = &HFF&
''                        End If
''                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, xColor, True, &HE2FEFB, Format(xTotalSaldoDol, "0.00")
'
'
'                        Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(RstDat("impdebsol"), "0.00")
'                        Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(RstDat("imphabsol"), "0.00")
''                        xTotalSaldoSol = (xTotalSaldoSol + NulosN(RstDat("impdebsol"))) - NulosN(RstDat("imphabsol"))
''                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, xColor, True, &HE2FEFB, Format(xTotalSaldoSol, "0.00")
'
'                        xTotDebDol = xTotDebDol + NulosN(RstDat("impdebdol"))
'                        xTotHabDol = xTotHabDol + NulosN(RstDat("imphabdol"))
'                        xTotDebSol = xTotDebSol + NulosN(RstDat("impdebsol"))
'                        xTotHabSol = xTotHabSol + NulosN(RstDat("imphabsol"))
'                    End If
'
'                    RstDat.MoveNext
'                    If RstDat.EOF = True Then
'                        Fg1.Rows = Fg1.Rows + 1
'                        GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "TOTAL ==>", flexAlignLeftCenter, True, , &H0&, &HE2FEFB, True
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H0&, True, &HE2FEFB, Format(xTotDebDol, "0.00")
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &H0&, True, &HE2FEFB, Format(xTotHabDol, "0.00")
'
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H0&, True, &HE2FEFB, Format(xTotDebSol, "0.00")
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H0&, True, &HE2FEFB, Format(xTotHabSol, "0.00")
'
'                        ' ESCRIBIMOS EL SALDO DE CADA MONEDA
'                        Fg1.Rows = Fg1.Rows + 1
'                        GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "SALDO ==>", flexAlignLeftCenter, True, , &H0&, &HE2FEFB, True
'
'                        'COLOR PARA LOS DOLARES
'                        xTotEmpDol = xTotEmpDol + (NulosN(xTotDebDol) - NulosN(xTotHabDol))
'                        If NulosN(xTotDebDol) - NulosN(xTotHabDol) > 0 Then
'                            xColor = &HFF0000
'                        Else
'                            xColor = &HFF&
'                        End If
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, xColor, True, &HE2FEFB, Format(NulosN(xTotDebDol) - NulosN(xTotHabDol), "0.00")
'
'                        'COLOR PARA LOS SOLES
'                        xTotEmpSol = xTotEmpSol + (NulosN(xTotDebSol) - NulosN(xTotHabSol))
'                        If NulosN(xTotDebSol) - NulosN(xTotHabSol) > 0 Then
'                            xColor = &HFF0000
'                        Else
'                            xColor = &HFF&
'                        End If
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, xColor, True, &HE2FEFB, Format(xTotDebSol - xTotHabSol, "0.00")
'
'                        'MOSTRAMOS EL TOTAL DE LA EMPRESA
'                        Fg1.Rows = Fg1.Rows + 1
'
'                        GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "TOTAL EMPRESA ==>", flexAlignLeftCenter, True, , &H0&, &HE2FEFB, True
'                        If xTotEmpDol > 0 Then
'                            xColor = &HFF0000
'                        Else
'                            xColor = &HFF&
'                        End If
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, xColor, True, &HE2FEFB, Format(NulosN(xTotEmpDol), "0.00")
'
'                        If xTotEmpSol > 0 Then
'                            xColor = &HFF0000
'                        Else
'                            xColor = &HFF&
'                        End If
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, xColor, True, &HE2FEFB, Format(NulosN(xTotEmpSol), "0.00")
'
'                        Fg1.Rows = Fg1.Rows + 1
'                        xTotDebDol = 0
'                        xTotHabDol = 0
'                        xTotDebSol = 0
'                        xTotHabSol = 0
'                        Exit For
'                    End If
'                Next B
'            End If
'
'
'            RstCli.MoveNext
'            If RstCli.EOF = True Then Exit For
'            Fg1.Rows = Fg1.Rows + 1
'        Next A
'    Else
'        MsgBox "No se ha encontrado registros en el periodo especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'    End If
'    Fg1.Cols = Fg1.Cols + 1
'    Fg1.ColDataType(Fg1.Cols - 1) = flexDTBoolean
'    Fg1.Editable = flexEDKbdMouse
'    If Fg1.Rows <> 2 Then
'        With Fg1
'            'VERDE
'            .Select 2, 9, Fg1.Rows - 1, 10
'            .FillStyle = flexFillRepeat
'            .CellBackColor = &HE6F8E0
'        End With
'    End If
'    If Fg1.Rows >= 3 Then Fg1.Select 2, 1
'
'    If Fg2.Rows = 0 Then Fg2.Rows = Fg2.Rows + 1
'
'End Sub
'

'***************************************************************************************************************
'***************************************************************************************************************
'Sub CargarDatos()
'    If TxtFchIni.Valor = "" Then
'        MsgBox "No ha especificado la fecha de inicio ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtFchIni.SetFocus
'        Exit Sub
'    End If
'
'    If TxtFchFin.Valor = "" Then
'        MsgBox "No ha especificado la fecha final ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtFchFin.SetFocus
'        Exit Sub
'    End If
'
'    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
'        MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Exit Sub
'    End If
'
'    Fg1.Rows = 2
'
'    Dim RstDat As New ADODB.Recordset
'    Dim RstCli As New ADODB.Recordset
'    Dim xCadWhere As String
'    Dim xSQL2 As String
'    Dim xSql As String
'    Dim A, B, C As Integer
'    Dim xSaldoD, xSaldoS, xTotalSaldoDol, xTotalSaldoSol As Double
'    Dim xColor As Long
'
'    Fg1.MergeCells = flexMergeFixedOnly
'
'    ' ELIMINAMOS LAS FILAS EN BLANCO
'    For A = 0 To Fg2.Rows - 1
'        If NulosN(Fg2.TextMatrix(A, 2)) = 0 Then
'            Fg2.RemoveItem A
'        End If
'    Next A
'
'    'ARMAMOS LAS SENTENCIA WHERE PARA LAS CONSULTAS
'    xCadWhere = ""
'    If Fg2.Rows <> 0 Then
'        xCadWhere = "WHERE ("
'        For A = 0 To Fg2.Rows - 1
'            xCadWhere = xCadWhere & "(Consulta4.idcli = " & NulosC(Fg2.TextMatrix(A, 2)) & ")"
'            If A = Fg2.Rows - 1 Then
'                xCadWhere = xCadWhere & ")"
'                Exit For
'            End If
'            xCadWhere = xCadWhere & " OR "
'        Next A
'    End If
'
'    ' OBTEMOS TODOS LOS DATOS PARA MOSTRAR
'    xSql = xSql + "SELECT Consulta4.* From " _
'        & "(" _
'        & " SELECT con_planctas.cuenta, con_planctas.descripcion, vta_ventas.fchreg, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, " _
'        & " vta_ventas.fchdoc, vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc, mae_moneda.simbolo, IIf(vta_ventas!tc=0,con_tc!impven,vta_ventas!tc) AS tc, " _
'        & " vta_ventas.imptotdoc, vta_ventas.numerodocref, IIf(vta_ventas!idmon=1, IIf(vta_ventas!tc=0,con_diario!impdebsol/con_tc!impven,con_diario!impdebsol/vta_ventas!tc), " _
'        & " con_diario!impdebdol) AS impdebdol, IIf(vta_ventas!idmon=1,IIf(vta_ventas!tc=0,con_diario!imphabsol/con_tc!impven,con_diario!imphabsol/vta_ventas!tc),con_diario!imphabdol) AS imphabdol, " _
'        & " IIf(vta_ventas!idmon=2, IIf(vta_ventas!tc=0,con_diario!impdebdol*con_tc!impven,con_diario!impdebdol*vta_ventas!tc),con_diario!impdebsol) AS impdebsol, " _
'        & " IIf(vta_ventas!idmon=2, IIf(vta_ventas!tc=0,con_diario!imphabdol*con_tc!impven,con_diario!imphabdol*vta_ventas!tc),con_diario!imphabsol) AS imphabsol, " _
'        & " vta_ventas.idcli FROM (((con_diario LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) RIGHT JOIN (mae_cliente RIGHT JOIN (vta_ventas " _
'        & " LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) ON mae_cliente.id = vta_ventas.idcli) ON con_diario.idmov = vta_ventas.id) LEFT JOIN mae_documento  " _
'        & " ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha WHERE (((con_planctas.cuenta) Like '12%')  " _
'        & " AND ((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=2) AND ((vta_ventas.anulado)=0) " _
'        & " AND ((vta_ventas.idmes)<>0)) " _
'        & " UNION SELECT con_planctas.cuenta, con_planctas.descripcion, vta_gastodebito.fchreg, mae_cliente.numruc, mae_cliente.nombre, " _
'        & " mae_documento.abrev,  vta_gastodebito.fchemi AS fchdoc, vta_gastodebito!numser+'-'+vta_gastodebito!numdoc AS numdoc, mae_moneda.simbolo, " _
'        & " IIf(vta_gastodebito!tc=0,con_tc!impven,vta_gastodebito!tc) AS tc, vta_gastodebito.imptotdoc AS imptotdoc, vta_gastodebito.numerodocref, " _
'        & " IIf(vta_gastodebito!idmon=1, IIf(vta_gastodebito!tc=0,con_diario!impdebsol/con_tc!impven, con_diario!impdebsol/vta_gastodebito!tc), vta_gastodebito!imptotdoc) AS impdebdol, " _
'        & " con_diario.imphabdol, con_diario.impdebsol, con_diario.imphabsol, vta_gastodebito.idcli FROM ((((((con_diario LEFT JOIN con_planctas " _
'        & " ON con_diario.idcue = con_planctas.id) RIGHT JOIN vta_gastodebito ON con_diario.idmov = vta_gastodebito.id) LEFT JOIN mae_cliente " _
'        & " ON vta_gastodebito.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN mae_moneda " _
'        & " ON vta_gastodebito.idmon = mae_moneda.id) LEFT JOIN mae_libros  ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha " _
'        & " WHERE (((vta_gastodebito.fchreg)>=CDate('" & TxtFchIni.Valor & "')  And (vta_gastodebito.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.impdebsol)<>0) " _
'        & " AND ((con_diario.idlib)=41) AND ((vta_gastodebito.idmes)<>0))  " _
'
'    xSql = xSql + " UNION " _
'        & " SELECT con_planctas.cuenta, con_planctas.descripcion, let_letra.fchreg, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, let_letradet.fchemi, " _
'        & " let_letradet!numser & '-' & let_letradet!numdoc AS numdoc, mae_moneda.simbolo, IIf(let_letra!tc=0,con_tc!impven,let_letra!tc) AS tc, let_letradet.implet, " _
'        & " let_letra!idaduana & let_letra!idregimen & let_letra!anoorden & let_letra!numorden AS numerodocref, con_diario.impdebdol, con_diario.imphabdol, " _
'        & " con_diario.impdebsol, con_diario.imphabsol, let_letra.idclipro FROM ((((let_letra LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id)  " _
'        & " LEFT JOIN mae_documento ON let_letra.idtipdoc = mae_documento.id) LEFT JOIN mae_moneda ON let_letra.idmon = mae_moneda.id) LEFT JOIN con_tc ON  " _
'        & " let_letra.fchemi = con_tc.fecha) LEFT JOIN ((con_diario RIGHT JOIN let_letradet ON (con_diario.idmov = let_letradet.idlet) AND " _
'        & " (con_diario.correlativo = let_letradet.corr)) LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) ON let_letra.id = let_letradet.idlet  " _
'        & " WHERE (((let_letra.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (let_letra.fchreg)<=CDate('" & TxtFchFin.Valor & "'))) ORDER BY mae_cliente.nombre " _
'        & " ) as Consulta4"
'
'    If xCadWhere <> "" Then
'        xSql = xSql & " " & xCadWhere
'    End If
'
'
'    ' OBTENEMOS LOS CLIENTES UNICOS
'    xSQL2 = "SELECT DISTINCT CONSULTA4.numruc, CONSULTA4.nombre, CONSULTA4.idcli " _
'        & " From (SELECT con_planctas.cuenta, con_planctas.descripcion, vta_ventas.fchreg, " _
'        & " mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, vta_ventas.fchdoc,  vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc, mae_moneda.simbolo, " _
'        & " IIf(vta_ventas!tc=0,con_tc!impven,vta_ventas!tc) AS tc, vta_ventas.imptotdoc, vta_ventas.numerodocref, " _
'        & " IIf(vta_ventas!idmon=1,IIf(vta_ventas!tc=0,con_diario!impdebsol/con_tc!impven,con_diario!impdebsol/vta_ventas!tc),con_diario!impdebdol) AS impdebdol, " _
'        & " IIf(vta_ventas!idmon=1,IIf(vta_ventas!tc=0,con_diario!imphabsol/con_tc!impven,con_diario!imphabsol/vta_ventas!tc),con_diario!imphabdol) AS imphabdol, " _
'        & " IIf(vta_ventas!idmon=2,IIf(vta_ventas!tc=0,con_diario!impdebdol*con_tc!impven,con_diario!impdebdol*vta_ventas!tc),con_diario!impdebsol) AS impdebsol, " _
'        & " IIf(vta_ventas!idmon=2,IIf(vta_ventas!tc=0,con_diario!imphabdol*con_tc!impven,con_diario!imphabdol*vta_ventas!tc),con_diario!imphabsol) AS imphabsol,  " _
'        & " vta_ventas.idcli FROM (((con_diario LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) RIGHT JOIN (mae_cliente " _
'        & " RIGHT JOIN (vta_ventas LEFT JOIN mae_moneda  ON vta_ventas.idmon = mae_moneda.id) ON mae_cliente.id = vta_ventas.idcli) ON con_diario.idmov = vta_ventas.id) " _
'        & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha WHERE (((con_planctas.cuenta) Like '12%') " _
'        & " AND ((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=2) AND ((vta_ventas.anulado)=0) " _
'        & " AND ((vta_ventas.idmes)<>0)) " _
'        & " UNION " _
'        & " SELECT con_planctas.cuenta, con_planctas.descripcion, let_letra.fchreg, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev," _
'        & " let_letradet.fchemi, let_letradet!numser & '-' & let_letradet!numdoc AS numdoc, mae_moneda.simbolo, IIf(let_letra!tc=0, con_tc!impven,let_letra!tc) AS tc, " _
'        & " let_letradet.implet, let_letra!idaduana & let_letra!idregimen & let_letra!anoorden & let_letra!numorden AS numerodocref, con_diario.impdebdol, " _
'        & " con_diario.imphabdol, con_diario.impdebsol, con_diario.imphabsol, let_letra.idclipro FROM ((((let_letra LEFT JOIN mae_cliente  " _
'        & " ON let_letra.idclipro = mae_cliente.id) LEFT JOIN mae_documento ON let_letra.idtipdoc = mae_documento.id) LEFT JOIN mae_moneda " _
'        & " ON let_letra.idmon = mae_moneda.id)  LEFT JOIN con_tc ON let_letra.fchemi = con_tc.fecha) LEFT JOIN ((con_diario RIGHT JOIN let_letradet " _
'        & " ON (con_diario.idmov = let_letradet.idlet)  AND (con_diario.correlativo = let_letradet.corr)) LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) " _
'        & " ON let_letra.id = let_letradet.idlet WHERE (((let_letra.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (let_letra.fchreg)<=CDate('" & TxtFchFin.Valor & "'))) "
'
'    xSQL2 = xSQL2 & " Union  " _
'        & " SELECT con_planctas.cuenta, con_planctas.descripcion, vta_gastodebito.fchreg, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, " _
'        & " vta_gastodebito.fchemi AS fchdoc,  vta_gastodebito!numser+'-'+vta_gastodebito!numdoc AS numdoc, mae_moneda.simbolo, " _
'        & " IIf(vta_gastodebito!tc=0,con_tc!impven,vta_gastodebito!tc) AS tc, vta_gastodebito.imptotdoc AS imptotdoc, vta_gastodebito.numerodocref,  " _
'        & " IIf(vta_gastodebito!idmon=1,IIf(vta_gastodebito!tc=0,con_diario!impdebsol/con_tc!impven,con_diario!impdebsol/vta_gastodebito!tc),vta_gastodebito!imptotdoc) " _
'        & " AS impdebdol, con_diario.imphabdol, con_diario.impdebsol, con_diario.imphabsol, vta_gastodebito.idcli FROM ((((((con_diario LEFT JOIN con_planctas " _
'        & " ON con_diario.idcue = con_planctas.id)  RIGHT JOIN vta_gastodebito ON con_diario.idmov = vta_gastodebito.id) LEFT JOIN mae_cliente " _
'        & " ON vta_gastodebito.idcli = mae_cliente.id) LEFT JOIN mae_documento  ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN mae_moneda " _
'        & " ON vta_gastodebito.idmon = mae_moneda.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha " _
'        & " WHERE (((vta_gastodebito.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_gastodebito.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.impdebsol)<>0) " _
'        & " AND ((con_diario.idlib)=41) AND ((vta_gastodebito.idmes)<>0)) " _
'        & " ) AS CONSULTA4"
'
'    If xCadWhere <> "" Then
'        xSQL2 = xSQL2 & " " & xCadWhere
'    End If
'
'    RST_Busq RstDat, xSql, xCon
'    RST_Busq RstCli, xSQL2, xCon
'
'    If RstCli.RecordCount <> 0 Then
'        RstCli.Sort = "nombre"
'        RstCli.MoveFirst
'        For A = 1 To RstCli.RecordCount
'            Fg1.Rows = Fg1.Rows + 1
'            Fg1.TextMatrix(Fg1.Rows - 1, 1) = "CLIENTE   :"
'
'            GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 6, "CLIENTE ==> " & RstCli("nombre"), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
'            GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "Nº DE RUC ==> ", flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
'            GRID_COMBINAR Fg1, Fg1.Rows - 1, 9, Fg1.Rows - 1, 11, RstCli("numruc"), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
'            ' FILTRAMOS LOS MOVIMIENTOS QUE TENGA EL CLIENTE ACTUAL
'            RstDat.Filter = adFilterNone
'            RstDat.Filter = "numruc = '" & RstCli("numruc") & "'"
'
'            If RstDat.RecordCount <> 0 Then
'                Dim xNumRef As String
'                Dim xTotDebDol, xTotHabDol, xTotDebSol, xTotHabSol As Double
'
'                ' ORDENAMOS POR NUMERO DE DOCUMENTO DE REFERENCIA Y POR FECHA DE EMISION DEL DOCUMENTO
'                RstDat.Sort = "numerodocref, fchdoc"
'                RstDat.MoveFirst
'
'                xNumRef = NulosC(RstDat("numerodocref"))
'                Fg1.Rows = Fg1.Rows + 1
'
'                GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 6, "DOC. REF. ==> " & "ORDEN DE DESPACHO", flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'                GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "Nº DE DOC. ==> ", flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'                GRID_COMBINAR Fg1, Fg1.Rows - 1, 9, Fg1.Rows - 1, 11, xNumRef, flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'
'                xTotalSaldoDol = 0
'                xTotalSaldoSol = 0
'                For B = 1 To RstDat.RecordCount
'                    If B > 1 Then
'                        If NulosC(xNumRef) = NulosC(RstDat("numerodocref")) Then
'                            Fg1.Rows = Fg1.Rows + 1
'                        Else
'                            Fg1.Rows = Fg1.Rows + 1
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "TOTAL ==>", flexAlignLeftCenter, True, , &H0&, &HE2FEFB, True
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H0&, True, &HE2FEFB, Format(xTotDebDol, "0.00")
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &H0&, True, &HE2FEFB, Format(xTotHabDol, "0.00")
'
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H0&, True, &HE2FEFB, Format(xTotDebSol, "0.00")
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H0&, True, &HE2FEFB, Format(xTotHabSol, "0.00")
'
'                            xTotalSaldoDol = 0
'                            xTotalSaldoSol = 0
'
'                            Fg1.Rows = Fg1.Rows + 1
'                            xTotDebDol = 0
'                            xTotHabDol = 0
'                            xTotDebSol = 0
'                            xTotHabSol = 0
'
'                            Fg1.Rows = Fg1.Rows + 1
'                            xNumRef = RstDat("numerodocref")
'
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 6, "DOC. REF. ==> " & "ORDEN DE DESPACHO", flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "Nº DE DOC. ==> ", flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 9, Fg1.Rows - 1, 11, xNumRef, flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'
'                            Fg1.Rows = Fg1.Rows + 1
'                        End If
'                    Else
'                        Fg1.Rows = Fg1.Rows + 1
'                    End If
'
'                    Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstDat("cuenta"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDat("descripcion"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 3) = RstDat("fchdoc")
'                    Fg1.TextMatrix(Fg1.Rows - 1, 4) = RstDat("abrev")
'                    Fg1.TextMatrix(Fg1.Rows - 1, 5) = RstDat("numdoc")
'                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = RstDat("simbolo")
'                    Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosN(RstDat("tc"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(RstDat("imptotdoc")), "0.00")
'
'                    If RstDat("abrev") = "LE" Then
'                        'aqui hacemos la inversa del asiento para poder hallar el saldo esto segun carlos
'                        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(RstDat("imphabdol"), "0.00")
'                        Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(RstDat("impdebdol"), "0.00")
'                        xTotalSaldoDol = (xTotalSaldoDol + RstDat("imphabdol")) - RstDat("impdebdol")
'                        If xTotalSaldoDol > 0 Then
'                            xColor = &HFF0000
'                        Else
'                            xColor = &HFF&
'                        End If
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, xColor, True, &HE2FEFB, Format(xTotalSaldoDol, "0.00")
'
'                        Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(RstDat("imphabsol"), "0.00")
'                        Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(RstDat("impdebsol"), "0.00")
'                        xTotalSaldoSol = (xTotalSaldoSol + NulosN(RstDat("imphabsol"))) - NulosN(RstDat("impdebsol"))
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, xColor, True, &HE2FEFB, Format(xTotalSaldoSol, "0.00")
'
'                        xTotDebDol = xTotDebDol + NulosN(RstDat("imphabdol"))
'                        xTotHabDol = xTotHabDol + NulosN(RstDat("impdebdol"))
'                        xTotDebSol = xTotDebSol + NulosN(RstDat("imphabsol"))
'                        xTotHabSol = xTotHabSol + NulosN(RstDat("impdebsol"))
'
'                    Else
'                        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(RstDat("impdebdol"), "0.00")
'                        Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(RstDat("imphabdol"), "0.00")
'                        xTotalSaldoDol = (xTotalSaldoDol + RstDat("impdebdol")) - RstDat("imphabdol")
'                        If xTotalSaldoDol > 0 Then
'                            xColor = &HFF0000
'                        Else
'                            xColor = &HFF&
'                        End If
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, xColor, True, &HE2FEFB, Format(xTotalSaldoDol, "0.00")
'
'
'                        Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(RstDat("impdebsol"), "0.00")
'                        Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(RstDat("imphabsol"), "0.00")
'                        xTotalSaldoSol = (xTotalSaldoSol + NulosN(RstDat("impdebsol"))) - NulosN(RstDat("imphabsol"))
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, xColor, True, &HE2FEFB, Format(xTotalSaldoSol, "0.00")
'
'                        xTotDebDol = xTotDebDol + NulosN(RstDat("impdebdol"))
'                        xTotHabDol = xTotHabDol + NulosN(RstDat("imphabdol"))
'                        xTotDebSol = xTotDebSol + NulosN(RstDat("impdebsol"))
'                        xTotHabSol = xTotHabSol + NulosN(RstDat("imphabsol"))
'                    End If
'
'                    RstDat.MoveNext
'                    If RstDat.EOF = True Then
'                        Fg1.Rows = Fg1.Rows + 1
'                        GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "TOTAL ==>", flexAlignLeftCenter, True, , &H0&, &HE2FEFB, True
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H0&, True, &HE2FEFB, Format(xTotDebDol, "0.00")
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &H0&, True, &HE2FEFB, Format(xTotHabDol, "0.00")
'
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H0&, True, &HE2FEFB, Format(xTotDebSol, "0.00")
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H0&, True, &HE2FEFB, Format(xTotHabSol, "0.00")
'
'                        Fg1.Rows = Fg1.Rows + 1
'                        xTotDebDol = 0
'                        xTotHabDol = 0
'                        xTotDebSol = 0
'                        xTotHabSol = 0
'                        Exit For
'                    End If
'                Next B
'            End If
'
'
'            RstCli.MoveNext
'            If RstCli.EOF = True Then Exit For
'            Fg1.Rows = Fg1.Rows + 1
'        Next A
'    Else
'        MsgBox "No se ha encontrado registros en el periodo especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'    End If
'
'    With Fg1
'        'VERDE
'        .Select 2, 9, Fg1.Rows - 1, 11
'        .FillStyle = flexFillRepeat
'        .CellBackColor = &HE6F8E0
'    End With
'    If Fg1.Rows >= 3 Then Fg1.Select 2, 1
'    If Fg2.Rows = 0 Then Fg2.Rows = Fg2.Rows + 1
'End Sub

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
    Dim xFun As New SGI2_funciones.Formularios
    Dim Rst As New ADODB.Recordset
    
    'If TabOne1.CurrTab = 0 Then
        If Fg1.Rows = 1 Then
            MsgBox "No hay registro para exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        If TabOne1.CurrTab = 0 Then
            GRID_EXPORTAR_MSEXCELTMP Fg1, xCon, flexFileCustomText, True, "Analisis de Cuenta x Cliente - DETALLADO", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en " & LblMoneda.Caption
        Else
            GRID_EXPORTAR_MSEXCELTMP Fg3, xCon, flexFileCustomText, True, "Analisis de Cuenta x Cliente - RESUMEM", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en " & LblMoneda.Caption
        End If
        Set xFun = Nothing
    'End If
'    If TabOne1.CurrTab = 1 Then
'        If Fg2.Rows = 2 Then
'            MsgBox "No hay registro para exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            Exit Sub
'        End If
''        ExportarExcelResumen
'        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg2, "RESUMEN DEL DIARIO", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en " & LblMoneda.Caption, "Diario - Resumen"   ', Rst, ""
'        Set xFun = Nothing
'    End If
End Sub

'Private Sub pImprimirDet()
'    TabOne1.CurrTab = 0
'    If Me.TabOne1.CurrTab = 0 Then
'        If Fg1.Rows = 1 Then
'            MsgBox "No hay registros para imprimir", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            Exit Sub
'        End If
'    End If
'
'    Dim nPeriodo   As String
'    Dim xMoneda As String
'    If opt_fecha(0).Value = True Then
'        If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
'            nPeriodo = " Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
'        Else
'            nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
'        End If
'    Else
'        If mMesIni = mMesFin Then
'            nPeriodo = "Periodo: " + lbl_periodo(0).Caption
'        Else
'            nPeriodo = "Periodo: De " + lbl_periodo(0).Caption & " A " & lbl_periodo(1).Caption
'        End If
'    End If
'
'    xMoneda = LblMoneda.Caption
'
'    Dim RstTmp As New ADODB.Recordset
'    Dim A As Integer
'    Dim Rst As New ADODB.Recordset
'    RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet Where (((con_formatostipodet.idformato) = 1) And ((con_formatostipodet.idformatotipo) = 2) " _
'        & " And ((con_formatostipodet.mostrar) = -1)) ORDER BY con_formatostipodet.orden", xCon
'
'    Dim xCampos() As String
'    Dim xFil, xCol As Double
'
'    ReDim xCampos(Fg1.Rows - 2, Fg1.Cols - 1)
'
'    Dim xFila As Double
'    xFila = 0
'    For xFil = 1 To Fg1.Rows - 1
'        For xCol = 1 To Fg1.Cols - 1
'            xCampos(xFila, xCol) = Fg1.TextMatrix(xFil, xCol)
'        Next xCol
'        xFila = xFila + 1
'    Next xFil
'
'    Rst.MoveFirst
'    For A = 1 To Rst.RecordCount
'        If xCampos(0, A) = Rst("abrev") Then
'            If Rst("imprimir") = False Then
'                xCampos(0, A) = ""
'            End If
'        End If
'        Rst.MoveNext
'        If Rst.EOF = True Then Exit For
'    Next A
'
'    Dim xfrm As New eps_librerias.Imprimir
'
'    xfrm.Cabecera1 = NomEmp
'    xfrm.Cabecera2 = "RUC Nº: " & NumRuc
'    xfrm.Fecha = Format(Date, "dd/mm/yyyy")
'    xfrm.Titulo1 = "LIBRO DIARIO " & "(Expresado en " & xMoneda & ")"
'    xfrm.Titulo2 = nPeriodo
'    xfrm.TamañoFuente = 6
'    xfrm.TamañoCabecera = 8
'    xfrm.FuenteCabecera = "Courier New"
'    xfrm.Posicion_Hoja = Vertical
'    xfrm.Tamaño_Hoja = A_4
'    xfrm.TextoConsiderar = "LIBRO"
'    xfrm.TextoConsiderarAncho = 5
'    xfrm.ImprimirArray xCampos, Rst
'    Set xfrm = Nothing
'End Sub

'Sub CargarResumen()
'    Dim xSql As String
'    Dim xSqlDet As String
'
'    xSql = "SELECT DISTINCT Consulta5.cuenta, Consulta5.descripcion From " _
'        & " ( " _
'        & " SELECT con_planctas.cuenta, con_planctas.descripcion, vta_ventas.fchreg, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, " _
'        & " vta_ventas.fchdoc, vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc, mae_moneda.simbolo, IIf(vta_ventas!tc=0,con_tc!impven,vta_ventas!tc) AS tc, " _
'        & " vta_ventas.imptotdoc, vta_ventas.numerodocref, IIf(vta_ventas!idmon=1,IIf(vta_ventas!tc=0,con_diario!impdebsol/con_tc!impven,con_diario!impdebsol/vta_ventas!tc),con_diario!impdebdol) AS impdebdol, " _
'        & " IIf(vta_ventas!idmon=1,IIf(vta_ventas!tc=0,con_diario!imphabsol/con_tc!impven,con_diario!imphabsol/vta_ventas!tc),con_diario!imphabdol) AS imphabdol, " _
'        & " IIf(vta_ventas!idmon=2,IIf(vta_ventas!tc=0,con_diario!impdebdol*con_tc!impven,con_diario!impdebdol*vta_ventas!tc),con_diario!impdebsol) AS impdebsol, " _
'        & " IIf(vta_ventas!idmon=2,IIf(vta_ventas!tc=0,con_diario!imphabdol*con_tc!impven,con_diario!imphabdol*vta_ventas!tc),con_diario!imphabsol) AS imphabsol, vta_ventas.idcli " _
'        & " FROM (((con_diario LEFT JOIN con_planctas ON con_diario.idcue=con_planctas.id) RIGHT JOIN (mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_moneda ON vta_ventas.idmon=mae_moneda.id) " _
'        & " ON mae_cliente.id=vta_ventas.idcli) ON con_diario.idmov=vta_ventas.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc=mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc=con_tc.fecha " _
'        & " WHERE (((con_planctas.cuenta) Like '12%') AND ((vta_ventas.fchreg)>=CDate('01/01/08') And (vta_ventas.fchreg)<=CDate('31/01/08')) AND ((con_diario.idlib)=2) " _
'        & " AND ((vta_ventas.anulado)=0) AND ((vta_ventas.idmes)<>0)) " _
'        & " Union " _
'        & " SELECT con_planctas.cuenta, con_planctas.descripcion, vta_gastodebito.fchreg, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, vta_gastodebito.fchemi AS fchdoc, " _
'        & " vta_gastodebito!numser+'-'+vta_gastodebito!numdoc AS numdoc, mae_moneda.simbolo, IIf(vta_gastodebito!tc=0,con_tc!impven,vta_gastodebito!tc) AS tc, " _
'        & " vta_gastodebito.imptot AS imptotdoc, vta_gastodebito.numerodocref, IIf(vta_gastodebito!idmon=1,IIf(vta_gastodebito!tc=0,con_diario!impdebsol/con_tc!impven,con_diario!impdebsol/vta_gastodebito!tc),vta_gastodebito!imptot) AS impdebdol, " _
'        & " con_diario.imphabdol, con_diario.impdebsol, con_diario.imphabsol, vta_gastodebito.idcli FROM ((((((con_diario LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) " _
'        & " RIGHT JOIN vta_gastodebito ON con_diario.idmov = vta_gastodebito.id) LEFT JOIN mae_cliente ON vta_gastodebito.idcli = mae_cliente.id) " _
'        & " LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id) LEFT JOIN mae_libros " _
'        & " ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha WHERE (((vta_gastodebito.fchreg)>=CDate('01/01/08') " _
'        & " And (vta_gastodebito.fchreg)<=CDate('31/01/08')) AND ((con_diario.impdebsol)<>0) AND ((con_diario.idlib)=41) AND ((vta_gastodebito.idmes)<>0)) "
'
'    xSql = xSql + "UNION " _
'        & " SELECT con_planctas.cuenta, con_planctas.descripcion, let_letra.fchreg, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, let_letradet.fchemi, " _
'        & " let_letradet!numser & '-' & let_letradet!numdoc AS numdoc, mae_moneda.simbolo, IIf(let_letra!tc=0,con_tc!impven,let_letra!tc) AS tc, let_letradet.implet, " _
'        & " let_letra!idaduana & let_letra!idregimen & let_letra!anoorden & let_letra!numorden AS numerodocref, con_diario.impdebdol, con_diario.imphabdol, con_diario.impdebsol, " _
'        & " con_diario.imphabsol, let_letra.idclipro FROM ((((let_letra LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id) LEFT JOIN mae_documento " _
'        & " ON let_letra.idtipdoc = mae_documento.id) LEFT JOIN mae_moneda ON let_letra.idmon = mae_moneda.id) LEFT JOIN con_tc ON let_letra.fchemi = con_tc.fecha) " _
'        & " LEFT JOIN ((con_diario RIGHT JOIN let_letradet ON (con_diario.idmov = let_letradet.idlet) AND (con_diario.correlativo = let_letradet.corr)) LEFT JOIN " _
'        & " con_planctas ON con_diario.idcue = con_planctas.id) ON let_letra.id = let_letradet.idlet WHERE (((let_letra.fchreg)>=CDate('01/01/08') And (let_letra.fchreg)<=CDate('31/01/08'))) " _
'        & " ORDER BY mae_cliente.nombre " _
'        & " ) AS Consulta5 ORDER BY Consulta5.cuenta"
'
'
'    xSqlDet = "SELECT con_planctas.cuenta, con_planctas.descripcion, vta_ventas.fchreg, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, vta_ventas.fchdoc, " _
'        & " vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc, mae_moneda.simbolo, IIf(vta_ventas!tc=0,con_tc!impven,vta_ventas!tc) AS tc, vta_ventas.imptotdoc, vta_ventas.numerodocref, " _
'        & " IIf(vta_ventas!idmon=1,IIf(vta_ventas!tc=0,con_diario!impdebsol/con_tc!impven,con_diario!impdebsol/vta_ventas!tc),con_diario!impdebdol) AS impdebdol, " _
'        & " IIf(vta_ventas!idmon=1,IIf(vta_ventas!tc=0,con_diario!imphabsol/con_tc!impven,con_diario!imphabsol/vta_ventas!tc),con_diario!imphabdol) AS imphabdol, " _
'        & " IIf(vta_ventas!idmon=2,IIf(vta_ventas!tc=0,con_diario!impdebdol*con_tc!impven,con_diario!impdebdol*vta_ventas!tc),con_diario!impdebsol) AS impdebsol, " _
'        & " IIf(vta_ventas!idmon=2,IIf(vta_ventas!tc=0,con_diario!imphabdol*con_tc!impven,con_diario!imphabdol*vta_ventas!tc),con_diario!imphabsol) AS imphabsol, " _
'        & " vta_ventas.idcli FROM (((con_diario LEFT JOIN con_planctas ON con_diario.idcue=con_planctas.id) RIGHT JOIN (mae_cliente RIGHT JOIN (vta_ventas " _
'        & " LEFT JOIN mae_moneda ON vta_ventas.idmon=mae_moneda.id) ON mae_cliente.id=vta_ventas.idcli) ON con_diario.idmov=vta_ventas.id) LEFT JOIN mae_documento " _
'        & " ON vta_ventas.tipdoc=mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc=con_tc.fecha WHERE (((con_planctas.cuenta) Like '12%') AND " _
'        & " ((vta_ventas.fchreg)>=CDate('01/01/08') And (vta_ventas.fchreg)<=CDate('31/01/08')) AND ((con_diario.idlib)=2) AND ((vta_ventas.anulado)=0) " _
'        & " AND ((vta_ventas.idmes)<>0)) " _
'        & " Union " _
'        & " SELECT con_planctas.cuenta, con_planctas.descripcion, vta_gastodebito.fchreg, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, vta_gastodebito.fchemi AS fchdoc, " _
'        & " vta_gastodebito!numser+'-'+vta_gastodebito!numdoc AS numdoc, mae_moneda.simbolo, IIf(vta_gastodebito!tc=0,con_tc!impven,vta_gastodebito!tc) AS tc, " _
'        & " vta_gastodebito.imptot AS imptotdoc, vta_gastodebito.numerodocref, " _
'        & " IIf(vta_gastodebito!idmon=1,IIf(vta_gastodebito!tc=0,con_diario!impdebsol/con_tc!impven,con_diario!impdebsol/vta_gastodebito!tc),vta_gastodebito!imptot) AS impdebdol, " _
'        & " con_diario.imphabdol, con_diario.impdebsol, con_diario.imphabsol, vta_gastodebito.idcli FROM ((((((con_diario LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) " _
'        & " RIGHT JOIN vta_gastodebito ON con_diario.idmov = vta_gastodebito.id) LEFT JOIN mae_cliente ON vta_gastodebito.idcli = mae_cliente.id) " _
'        & " LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id) LEFT JOIN mae_libros " _
'        & " ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha WHERE (((vta_gastodebito.fchreg)>=CDate('01/01/08') " _
'        & " And (vta_gastodebito.fchreg)<=CDate('31/01/08')) AND ((con_diario.impdebsol)<>0) AND ((con_diario.idlib)=41) AND ((vta_gastodebito.idmes)<>0)) "
'
'    xSqlDet = xSqlDet & " UNION " _
'        & " SELECT con_planctas.cuenta, con_planctas.descripcion, let_letra.fchreg, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, let_letradet.fchemi, " _
'        & " let_letradet!numser & '-' & let_letradet!numdoc AS numdoc, mae_moneda.simbolo, IIf(let_letra!tc=0,con_tc!impven,let_letra!tc) AS tc, let_letradet.implet, " _
'        & " let_letra!idaduana & let_letra!idregimen & let_letra!anoorden & let_letra!numorden AS numerodocref, con_diario.impdebdol, con_diario.imphabdol, " _
'        & " con_diario.impdebsol, con_diario.imphabsol, let_letra.idclipro FROM ((((let_letra LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id) " _
'        & " LEFT JOIN mae_documento ON let_letra.idtipdoc = mae_documento.id) LEFT JOIN mae_moneda ON let_letra.idmon = mae_moneda.id) LEFT JOIN con_tc ON " _
'        & " let_letra.fchemi = con_tc.fecha) LEFT JOIN ((con_diario RIGHT JOIN let_letradet ON (con_diario.idmov = let_letradet.idlet) AND (con_diario.correlativo = let_letradet.corr)) " _
'        & "  LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) ON let_letra.id = let_letradet.idlet WHERE (((let_letra.fchreg)>=CDate('01/01/08') " _
'        & " And (let_letra.fchreg)<=CDate('31/01/08'))) ORDER BY mae_cliente.nombre"
'
'
'    Dim RstRes As New ADODB.Recordset
'    Dim RstDet As New ADODB.Recordset
'    Dim xTotalSol, xTotalDol As Double
'    RST_Busq RstRes, xSql, xCon
'    RST_Busq RstDet, xSqlDet, xCon
'
'    Dim A, B As Integer
'
'    If RstRes.RecordCount <> 0 Then
'        RstRes.MoveFirst
'
'        FgRes.Rows = 2
'        For A = 1 To RstRes.RecordCount
'            FgRes.Rows = FgRes.Rows + 1
'            FgRes.TextMatrix(FgRes.Rows - 1, 1) = "Nº CUENTA ==> " & NulosC(RstRes("cuenta"))
'            GRID_COMBINAR FgRes, FgRes.Rows - 1, 1, FgRes.Rows - 1, 2, "Nº CUENTA ==> " & NulosC(RstRes("cuenta")), flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'            GRID_COMBINAR FgRes, FgRes.Rows - 1, 4, FgRes.Rows - 1, 11, "CUENTA ==> " & RstRes("descripcion"), flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'            'RstRes.MoveNext
'            RstDet.MoveFirst
'            RstDet.Filter = "cuenta = '" & RstRes("cuenta") & "'"
'
'            If RstDet.RecordCount <> 0 Then
'                RstDet.MoveFirst
'                RstDet.Sort = "nombre, numerodocref, fchdoc"
'
'                For B = 1 To RstDet.RecordCount
'                    FgRes.Rows = FgRes.Rows + 1
'                    FgRes.TextMatrix(FgRes.Rows - 1, 1) = RstDet("numruc")
'                    FgRes.TextMatrix(FgRes.Rows - 1, 2) = RstDet("nombre")
'                    FgRes.TextMatrix(FgRes.Rows - 1, 3) = NulosC(RstDet("numerodocref"))
'                    FgRes.TextMatrix(FgRes.Rows - 1, 4) = NulosC(RstDet("abrev"))
'                    FgRes.TextMatrix(FgRes.Rows - 1, 5) = RstDet("fchdoc")
'                    FgRes.TextMatrix(FgRes.Rows - 1, 6) = RstDet("numdoc")
'                    FgRes.TextMatrix(FgRes.Rows - 1, 7) = RstDet("simbolo")
'                    FgRes.TextMatrix(FgRes.Rows - 1, 8) = Format(RstDet("tc"), "0.000")
'                    FgRes.TextMatrix(FgRes.Rows - 1, 9) = Format(RstDet("imptotdoc"), "0.00")
'
'                    If NulosC(RstDet("simbolo")) = "S/." Then
'                        FgRes.TextMatrix(FgRes.Rows - 1, 10) = Format(RstDet("imptotdoc") / RstDet("tc"), "0.00")
'                        FgRes.TextMatrix(FgRes.Rows - 1, 11) = Format(RstDet("imptotdoc"), "0.00")
'                    Else
'                        FgRes.TextMatrix(FgRes.Rows - 1, 10) = Format(RstDet("imptotdoc"), "0.00")
'                        FgRes.TextMatrix(FgRes.Rows - 1, 11) = Format(RstDet("imptotdoc") * RstDet("tc"), "0.00")
'                    End If
'
'                    xTotalDol = xTotalDol + NulosN(FgRes.TextMatrix(FgRes.Rows - 1, 10))
'                    xTotalSol = xTotalSol + NulosN(FgRes.TextMatrix(FgRes.Rows - 1, 11))
'
'                    RstDet.MoveNext
'                    If RstDet.EOF = True Then
'                        FgRes.Rows = FgRes.Rows + 1
'                        GRID_COMBINAR FgRes, FgRes.Rows - 1, 8, FgRes.Rows - 1, 9, "TOTAL ==> ", flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'                        FORMATO_CELDA FgRes, FgRes.Rows - 1, 10, &H80&, True, &HE2FEFB, Format(xTotalDol, "0.00")
'                        FORMATO_CELDA FgRes, FgRes.Rows - 1, 11, &H80&, True, &HE2FEFB, Format(xTotalSol, "0.00")
'
'                        FgRes.Rows = FgRes.Rows + 1
'                        Exit For
'                    End If
'                Next B
'
'            End If
'            RstRes.MoveNext
'
'            If RstRes.EOF = True Then Exit For
'            RstDet.Filter = adFilterNone
'        Next A
'    Else
'        MsgBox "No hay registros para mostrar en el periodo especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'    End If
'    Set RstRes = Nothing
'End Sub



'Sub CargaDatosFormatoSavar2()
'    If TxtFchIni.Valor = "" Then
'        MsgBox "No ha especificado la fecha de inicio ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtFchIni.SetFocus
'        Exit Sub
'    End If
'
'    If TxtFchFin.Valor = "" Then
'        MsgBox "No ha especificado la fecha final ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtFchFin.SetFocus
'        Exit Sub
'    End If
'
'    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
'        MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Exit Sub
'    End If
'
'    SetearCuadricula Fg1, 5, xCon, 5, 1, False
'    Fg1.Rows = 2
'
'    Dim RstDat As New ADODB.Recordset
'    Dim RstCli As New ADODB.Recordset
'    Dim xCadWhere As String
'    Dim xSQL2 As String
'    Dim xSql As String
'    Dim A, B, C As Integer
'    Dim xSaldoD, xSaldoS, xTotalSaldoDol, xTotalSaldoSol, xTotEmpSol, xTotEmpDol As Double
'    Dim xColor As Long
'
'    Fg1.MergeCells = flexMergeFixedOnly
'
'    ' ELIMINAMOS LAS FILAS EN BLANCO
'    For A = 0 To Fg2.Rows - 1
'        If NulosN(Fg2.TextMatrix(A, 2)) = 0 Then
'            Fg2.RemoveItem A
'        End If
'    Next A
'
'    'ARMAMOS LAS SENTENCIA WHERE PARA LAS CONSULTAS
'    xCadWhere = ""
'    If Fg2.Rows <> 0 Then
'        xCadWhere = " AND ("
'        For A = 0 To Fg2.Rows - 1
'            xCadWhere = xCadWhere & "(ctacte2.idclipro = " & NulosC(Fg2.TextMatrix(A, 2)) & ")"
'            If A = Fg2.Rows - 1 Then
'                xCadWhere = xCadWhere & ")"
'                Exit For
'            End If
'            xCadWhere = xCadWhere & " OR "
'        Next A
'    End If
'
'    ' OBTEMOS TODOS LOS DATOS PARA MOSTRAR
'    xSql = ""
'    xSQL2 = ""
'    xSql = "SELECT ctacte1.* From " _
'        & " ( " _
'        & " SELECT con_provicionesdetdoc.idclipro, mae_cliente.nombre, " _
'        & " Sum(IIf([con_provicionesdetdoc]![idmon]=2,IIf([con_provicionesdet]![tipo]=-1,[con_provicionesdetdoc]![impdoc],0),0)) AS debedol, " _
'        & " Sum(IIf([con_provicionesdetdoc]![idmon]=2,IIf([con_provicionesdet]![tipo]=0,[con_provicionesdetdoc]![impdoc],0),0)) AS haberdol, " _
'        & " Sum(IIf([con_provicionesdetdoc]![idmon]=1,IIf([con_provicionesdet]![tipo]=-1,[con_provicionesdetdoc]![impdoc],0),0)) AS debesol, " _
'        & " Sum(IIf([con_provicionesdetdoc]![idmon]=1,IIf([con_provicionesdet]![tipo]=0,[con_provicionesdetdoc]![impdoc],0),0)) AS habersol, " _
'        & " con_provicionesdetdoc.numdocref FROM con_proviciones RIGHT JOIN (con_provicionesdet RIGHT JOIN (((mae_cliente RIGHT JOIN (con_provicionesdetdoc " _
'        & " LEFT JOIN con_planctas ON con_provicionesdetdoc.idcue = con_planctas.id) ON mae_cliente.id = con_provicionesdetdoc.idclipro) LEFT JOIN mae_documento " _
'        & " ON con_provicionesdetdoc.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON con_provicionesdetdoc.idmon = mae_moneda.id) " _
'        & " ON (con_provicionesdet.id = con_provicionesdetdoc.idpro) AND (con_provicionesdet.idcuen = con_provicionesdetdoc.idcue)) ON con_proviciones.id = con_provicionesdet.id " _
'        & " GROUP BY con_provicionesdetdoc.idclipro, mae_cliente.nombre, con_provicionesdetdoc.numdocref, con_proviciones.idmes " _
'        & " Having (((con_proviciones.idmes) = 0)) ORDER BY mae_cliente.nombre, con_provicionesdetdoc.numdocref " _
'        & " ) AS ctacte2 " _
'        & " Left Join " _
'        & " ( " _
'        & " SELECT con_provicionesdetdoc.idclipro, mae_cliente.nombre, con_provicionesdetdoc.idcue, con_planctas.cuenta, con_planctas.descripcion, " _
'        & " mae_documento.abrev, con_provicionesdetdoc.fchemi, mae_moneda.simbolo, [con_provicionesdetdoc]![numser] & '-' & [con_provicionesdetdoc]![numdoc] AS numdoc, " _
'        & " con_provicionesdetdoc.impdoc, IIf([con_provicionesdetdoc]![idmon]=2,IIf([con_provicionesdet]![tipo]=-1,[con_provicionesdetdoc]![impdoc],0),0) AS debedol, " _
'        & " IIf([con_provicionesdetdoc]![idmon]=2,IIf([con_provicionesdet]![tipo]=0,[con_provicionesdetdoc]![impdoc],0),0) AS haberdol, " _
'        & " IIf([con_provicionesdetdoc]![idmon]=1,IIf([con_provicionesdet]![tipo]=-1,[con_provicionesdetdoc]![impdoc],0),0) AS debesol, " _
'        & " IIf([con_provicionesdetdoc]![idmon]=1,IIf([con_provicionesdet]![tipo]=0,[con_provicionesdetdoc]![impdoc],0),0) AS habersol, " _
'        & " con_provicionesdetdoc.numdocref FROM con_proviciones RIGHT JOIN (con_provicionesdet RIGHT JOIN (((mae_cliente RIGHT JOIN " _
'        & " (con_provicionesdetdoc LEFT JOIN con_planctas ON con_provicionesdetdoc.idcue = con_planctas.id) ON mae_cliente.id = con_provicionesdetdoc.idclipro) "
'
'    xSql = xSql + "LEFT JOIN mae_documento ON con_provicionesdetdoc.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON con_provicionesdetdoc.idmon = mae_moneda.id) " _
'        & " ON (con_provicionesdet.idcuen = con_provicionesdetdoc.idcue) AND (con_provicionesdet.id = con_provicionesdetdoc.idpro)) ON con_proviciones.id = con_provicionesdet.id " _
'        & " Where (((con_proviciones.idmes) = 0)) ORDER BY con_provicionesdetdoc.numdocref " _
'        & " ) AS ctacte1 " _
'        & " ON (ctacte2.idclipro = ctacte1.idclipro) AND (ctacte2.numdocref = ctacte1.numdocref) " _
'        & " Where ((([ctacte2]![debedol] - [ctacte2]![haberdol]) <> 0)) Or ((([ctacte2]![debesol] - [ctacte2]![habersol]) <> 0)) " _
'        & " ORDER BY ctacte2.nombre, ctacte2.numdocref"
'
'
'    '------------------------------
'    ' OBTENEMOS LOS CLIENTES UNICOS
'    xSQL2 = "SELECT DISTINCT ctacte2.idclipro, ctacte2.numruc, ctacte2.nombre" _
'        & " FROM " _
'        & " (SELECT con_provicionesdetdoc.idclipro, mae_cliente.nombre, " _
'        & " Sum(IIf([con_provicionesdetdoc]![idmon]=2,IIf([con_provicionesdet]![tipo]=-1,[con_provicionesdetdoc]![impdoc],0),0)) AS debedol, " _
'        & " Sum(IIf([con_provicionesdetdoc]![idmon]=2,IIf([con_provicionesdet]![tipo]=0,[con_provicionesdetdoc]![impdoc],0),0)) AS haberdol, " _
'        & " Sum(IIf([con_provicionesdetdoc]![idmon]=1,IIf([con_provicionesdet]![tipo]=-1,[con_provicionesdetdoc]![impdoc],0),0)) AS debesol, " _
'        & " Sum(IIf([con_provicionesdetdoc]![idmon]=1,IIf([con_provicionesdet]![tipo]=0,[con_provicionesdetdoc]![impdoc],0),0)) AS habersol, " _
'        & " con_provicionesdetdoc.numdocref, mae_docreferencia.descripcion AS docref, mae_cliente.numruc FROM con_proviciones RIGHT JOIN " _
'        & " ((con_provicionesdet RIGHT JOIN (((mae_cliente RIGHT JOIN (con_provicionesdetdoc LEFT JOIN con_planctas ON con_provicionesdetdoc.idcue = con_planctas.id) " _
'        & " ON mae_cliente.id = con_provicionesdetdoc.idclipro) LEFT JOIN mae_documento ON con_provicionesdetdoc.tipdoc = mae_documento.id) " _
'        & " LEFT JOIN mae_moneda ON con_provicionesdetdoc.idmon = mae_moneda.id) ON (con_provicionesdet.idcuen = con_provicionesdetdoc.idcue) " _
'        & " AND (con_provicionesdet.id = con_provicionesdetdoc.idpro)) LEFT JOIN mae_docreferencia ON con_provicionesdetdoc.idtipdocref = mae_docreferencia.id) " _
'        & " ON con_proviciones.id = con_provicionesdet.id GROUP BY con_provicionesdetdoc.idclipro, mae_cliente.nombre, con_provicionesdetdoc.numdocref, " _
'        & " mae_docreferencia.descripcion, mae_cliente.numruc, con_proviciones.idmes Having (((con_proviciones.idmes) = 0))  ORDER BY mae_cliente.nombre, " _
'        & " con_provicionesdetdoc.numdocref) AS ctacte2 " _
'        & " Where ((([ctacte2]![debedol] - [ctacte2]![haberdol]) <> 0 " & xCadWhere & ") " _
'        & " Or     (([ctacte2]![debesol] - [ctacte2]![habersol]) <> 0 " & xCadWhere & ")) " _
'        & " ORDER BY ctacte2.nombre"
'
'    RST_Busq RstDat, xSql, xCon
'    RST_Busq RstCli, xSQL2, xCon
'
'     TabOne1.CurrTab = 0
'    If RstCli.RecordCount <> 0 Then
'        RstCli.Sort = "nombre"
'        RstCli.MoveFirst
'        For A = 1 To RstCli.RecordCount
'            Fg1.Rows = Fg1.Rows + 1
'            Fg1.TextMatrix(Fg1.Rows - 1, 1) = "CLIENTE   :"
'
'            GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 6, "CLIENTE ==> " & RstCli("nombre"), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
'            GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "Nº DE RUC ==> ", flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
'            GRID_COMBINAR Fg1, Fg1.Rows - 1, 9, Fg1.Rows - 1, 10, RstCli("numruc"), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
'
'            ' FILTRAMOS LOS MOVIMIENTOS QUE TENGA EL CLIENTE ACTUAL
'            RstDat.Filter = adFilterNone
'            'RstDat.Filter = "numruc = '" & RstCli("numruc") & "'"
'            RstDat.Filter = "idclipro = '" & RstCli("idclipro") & "'"
'            '16681
'            If RstDat.RecordCount <> 0 Then
'                Dim xNumRef As String
'                Dim xTotDebDol, xTotHabDol, xTotDebSol, xTotHabSol As Double
'
'                ' ORDENAMOS POR NUMERO DE DOCUMENTO DE REFERENCIA Y POR FECHA DE EMISION DEL DOCUMENTO
'                RstDat.Sort = "numdocref, fchemi"
'                RstDat.MoveFirst
'
'                xNumRef = NulosC(RstDat("numdocref"))
'                Fg1.Rows = Fg1.Rows + 1
'
'                GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 6, "DOC. REF. ==> " & "ORDEN DE DESPACHO", flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'                GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "Nº DE DOC. ==> ", flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'                GRID_COMBINAR Fg1, Fg1.Rows - 1, 9, Fg1.Rows - 1, 10, xNumRef, flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'
'                xTotalSaldoDol = 0
'                xTotalSaldoSol = 0
'                For B = 1 To RstDat.RecordCount
'                    If B > 1 Then
'                        If NulosC(xNumRef) = NulosC(RstDat("numdocref")) Then
'                            Fg1.Rows = Fg1.Rows + 1
'                        Else
'                            Fg1.Rows = Fg1.Rows + 1
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "TOTAL ==>", flexAlignLeftCenter, True, , &H0&, &HE2FEFB, True
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H0&, True, &HE2FEFB, Format(xTotDebDol, "0.00")
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &H0&, True, &HE2FEFB, Format(NulosN(xTotHabDol), "0.00")
'
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H0&, True, &HE2FEFB, Format(xTotDebSol, "0.00")
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H0&, True, &HE2FEFB, Format(xTotHabSol, "0.00")
'
'                            ' ESCRIBIMOS EL SALDO DE CADA MONEDA
'                            Fg1.Rows = Fg1.Rows + 1
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "SALDO ==>", flexAlignLeftCenter, True, , &H0&, &HE2FEFB, True
'
'                            'COLOR PARA LOS DOLARES
'                            xTotEmpDol = xTotEmpDol + (NulosN(xTotDebDol) - NulosN(xTotHabDol))
'                            If NulosN(xTotDebDol) - NulosN(xTotHabDol) > 0 Then
'                                xColor = &HFF0000
'                            Else
'                                xColor = &HFF&
'                            End If
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, xColor, True, &HE2FEFB, Format(NulosN(xTotDebDol) - NulosN(xTotHabDol), "0.00")
'
'                            'COLOR PARA LOS SOLES
'                            xTotEmpSol = xTotEmpSol + (NulosN(xTotDebSol) - NulosN(xTotHabSol))
'                            If NulosN(xTotDebSol) - NulosN(xTotHabSol) > 0 Then
'                                xColor = &HFF0000
'                            Else
'                                xColor = &HFF&
'                            End If
'                            FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, xColor, True, &HE2FEFB, Format(xTotDebSol - xTotHabSol, "0.00")
'                            xTotalSaldoDol = 0
'                            xTotalSaldoSol = 0
'
'                            Fg1.Rows = Fg1.Rows + 1
'                            xTotDebDol = 0
'                            xTotHabDol = 0
'                            xTotDebSol = 0
'                            xTotHabSol = 0
'
'                            Fg1.Rows = Fg1.Rows + 1
'                            xNumRef = RstDat("numdocref")
'
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 6, "DOC. REF. ==> " & "ORDEN DE DESPACHO", flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "Nº DE DOC. ==> ", flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'                            GRID_COMBINAR Fg1, Fg1.Rows - 1, 9, Fg1.Rows - 1, 10, xNumRef, flexAlignLeftCenter, True, , &H80&, &HE2FEFB, True
'
'                            Fg1.Rows = Fg1.Rows + 1
'                        End If
'                    Else
'                        Fg1.Rows = Fg1.Rows + 1
'                    End If
'
'                    Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstDat("cuenta"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDat("descripcion"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 3) = RstDat("fchemi")
'                    Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(RstDat("abrev"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(RstDat("numdoc"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = RstDat("simbolo")
'                    'Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosN(RstDat("tc"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(RstDat("impdoc")), "0.00")
'
'                    Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(RstDat("debedol"), "0.00")
'                    Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(RstDat("haberdol"), "0.00")
'                    Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(RstDat("debesol"), "0.00")
'                    Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(RstDat("habersol"), "0.00")
'
'                    xTotDebDol = xTotDebDol + NulosN(RstDat("debedol"))
'                    xTotHabDol = xTotHabDol + NulosN(RstDat("haberdol"))
'                    xTotDebSol = xTotDebSol + NulosN(RstDat("debesol"))
'                    xTotHabSol = xTotHabSol + NulosN(RstDat("habersol"))
'
'
'                    RstDat.MoveNext
'                    If RstDat.EOF = True Then
'                        Fg1.Rows = Fg1.Rows + 1
'                        GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "TOTAL ==>", flexAlignLeftCenter, True, , &H0&, &HE2FEFB, True
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H0&, True, &HE2FEFB, Format(xTotDebDol, "0.00")
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &H0&, True, &HE2FEFB, Format(xTotHabDol, "0.00")
'
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H0&, True, &HE2FEFB, Format(xTotDebSol, "0.00")
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H0&, True, &HE2FEFB, Format(xTotHabSol, "0.00")
'
'                        ' ESCRIBIMOS EL SALDO DE CADA MONEDA
'                        Fg1.Rows = Fg1.Rows + 1
'                        GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "SALDO ==>", flexAlignLeftCenter, True, , &H0&, &HE2FEFB, True
'
'                        'COLOR PARA LOS DOLARES
'                        xTotEmpDol = xTotEmpDol + (NulosN(xTotDebDol) - NulosN(xTotHabDol))
'                        If NulosN(xTotDebDol) - NulosN(xTotHabDol) > 0 Then
'                            xColor = &HFF0000
'                        Else
'                            xColor = &HFF&
'                        End If
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, xColor, True, &HE2FEFB, Format(NulosN(xTotDebDol) - NulosN(xTotHabDol), "0.00")
'
'                        'COLOR PARA LOS SOLES
'                        xTotEmpSol = xTotEmpSol + (NulosN(xTotDebSol) - NulosN(xTotHabSol))
'                        If NulosN(xTotDebSol) - NulosN(xTotHabSol) > 0 Then
'                            xColor = &HFF0000
'                        Else
'                            xColor = &HFF&
'                        End If
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, xColor, True, &HE2FEFB, Format(xTotDebSol - xTotHabSol, "0.00")
'
'                        'MOSTRAMOS EL TOTAL DE LA EMPRESA
'                        Fg1.Rows = Fg1.Rows + 1
'
'                        GRID_COMBINAR Fg1, Fg1.Rows - 1, 7, Fg1.Rows - 1, 8, "TOTAL EMPRESA ==>", flexAlignLeftCenter, True, , &H0&, &HE2FEFB, True
'                        If xTotEmpDol > 0 Then
'                            xColor = &HFF0000
'                        Else
'                            xColor = &HFF&
'                        End If
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, xColor, True, &HE2FEFB, Format(NulosN(xTotEmpDol), "0.00")
'
'                        If xTotEmpSol > 0 Then
'                            xColor = &HFF0000
'                        Else
'                            xColor = &HFF&
'                        End If
'                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, xColor, True, &HE2FEFB, Format(NulosN(xTotEmpSol), "0.00")
'
'                        Fg1.Rows = Fg1.Rows + 1
'                        xTotDebDol = 0
'                        xTotHabDol = 0
'                        xTotDebSol = 0
'                        xTotHabSol = 0
'                        Exit For
'                    End If
'                Next B
'            End If
'
'            RstCli.MoveNext
'            If RstCli.EOF = True Then Exit For
'            Fg1.Rows = Fg1.Rows + 1
'        Next A
'    Else
'        MsgBox "No se ha encontrado registros en el periodo especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'    End If
'    Fg1.Cols = Fg1.Cols + 1
'    Fg1.ColDataType(Fg1.Cols - 1) = flexDTBoolean
'    Fg1.Editable = flexEDKbdMouse
'    If Fg1.Rows <> 2 Then
'        With Fg1
'            'VERDE
'            .Select 2, 9, Fg1.Rows - 1, 10
'            .FillStyle = flexFillRepeat
'            .CellBackColor = &HE6F8E0
'        End With
'    End If
'    If Fg1.Rows >= 3 Then Fg1.Select 2, 1
'
'    If Fg2.Rows = 0 Then Fg2.Rows = Fg2.Rows + 1
'
'End Sub

Sub CargarSelect()
    Dim A As Integer
    Dim xCadWhere As String
    
    Fg1.Rows = 2
    Fg3.Rows = 2
    
    If Check1.Value = 0 And Check2.Value = 0 Then
        MsgBox "No ha especificado que datos se van a mostrar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Check1.SetFocus
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
    
    ' ELIMINAMOS LAS FILAS EN BLANCO
    If Fg2.Rows <> 0 Then
        For A = 0 To Fg2.Rows - 1
            If NulosN(Fg2.TextMatrix(A, 2)) = 0 Then
                Fg2.RemoveItem A
            End If
        Next A
    End If
    
    Dim xSqlRes1 As String
    Dim xSqlRes2 As String
    Dim xSqlDet1 As String
    Dim xSqlDet2 As String
    
    'CONSULTA PARA TRAER EL RESUMEN DE LOS PENDIENTES DEL AÑO ACTUAL
    xSqlRes1 = "SELECT movimiento.idcli, mae_cliente.nombre, Sum([movimiento]![debedol]-[movimiento]![haberdol]) AS saldodol, " _
        & " Sum([movimiento]![debesol]-[movimiento]![habersol]) AS saldosol, movimiento.numerodocref, mae_docreferencia.descripcion AS docref, " _
        & " mae_cliente.numruc FROM ( " _
        & " ( " _
        & " SELECT vta_ventas.idcli, con_diario.idcue, vta_ventas.idmon, vta_ventas.tipdoc, vta_ventas.idtipdocref, vta_ventas.fchdoc AS fchemi, " _
        & " [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.imptotdoc AS impdoc, vta_ventas.tc, " _
        & " IIf([vta_ventas]![idmon]=2,[vta_ventas]![imptotdoc],0) AS debedol, 0 AS haberdol, IIf([vta_ventas]![idmon]=1,[vta_ventas]![imptotdoc],0) AS debesol, " _
        & " 0 AS habersol, vta_ventas.numerodocref FROM con_diario RIGHT JOIN vta_ventas ON con_diario.idmov = vta_ventas.id WHERE (((vta_ventas.tipdoc)<>7) " _
        & " AND ((con_diario.impdebsol)<>0) AND ((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((con_diario.idlib)=2) AND ((vta_ventas.anulado)=0)) OR (((vta_ventas.tipdoc)<>7) AND ((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=2) AND ((vta_ventas.anulado)=0) AND ((con_diario.impdebdol)<>0)) " _
        & " UNION " _
        & " SELECT vta_ventas.idcli, con_diario.idcue, vta_ventas.idmon, vta_ventas.tipdoc, vta_ventas.idtipdocref, vta_ventas.fchdoc AS fchemi, " _
        & " [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.imptotdoc AS impdoc, vta_ventas.tc, 0 AS debedol, " _
        & " IIf([vta_ventas]![idmon]=2,[vta_ventas]![imptotdoc],0) AS haberdol, 0 AS debesol, IIf([vta_ventas]![idmon]=1,[vta_ventas]![imptotdoc],0) AS habersol, " _
        & " vta_ventas.numerodocref FROM con_diario RIGHT JOIN vta_ventas ON con_diario.idmov = vta_ventas.id WHERE (((vta_ventas.tipdoc)=7) " _
        & " AND ((con_diario.imphabsol)<>0) AND ((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((con_diario.idlib)=2) AND ((vta_ventas.anulado)=0)) OR (((vta_ventas.tipdoc)=7) AND ((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=2) AND ((vta_ventas.anulado)=0) AND ((con_diario.imphabdol)<>0))"
        
        xSqlRes1 = xSqlRes1 & " UNION " _
        & " SELECT vta_gastodebito.idcli, con_diario.idcue, vta_gastodebito.idmon, vta_gastodebito.tipdoc, vta_gastodebito.idtipdocref, " _
        & " vta_gastodebito.fchemi AS fchemi, [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc] AS numdoc, vta_gastodebito.imptotdoc AS impdoc, " _
        & " vta_gastodebito.tc, IIf([vta_gastodebito]![idmon]=2,[vta_gastodebito]![imptotdoc],0) AS debedol, 0 AS haberdol, " _
        & " IIf([vta_gastodebito]![idmon]=1,[vta_gastodebito]![imptotdoc],0) AS debesol, 0 AS habersol, vta_gastodebito.numerodocref " _
        & " FROM con_diario RIGHT JOIN vta_gastodebito ON con_diario.idmov = vta_gastodebito.id WHERE (((vta_gastodebito.tipdoc)=120) AND " _
        & " ((con_diario.impdebsol)<>0) AND ((vta_gastodebito.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_gastodebito.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((con_diario.idlib)=41)) OR (((vta_gastodebito.tipdoc)=120) AND ((vta_gastodebito.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (vta_gastodebito.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=41) AND ((con_diario.impdebdol)<>0)) " _
        & " UNION " _
        & " SELECT vta_gastodebito.idcli, con_diario.idcue, vta_gastodebito.idmon, vta_gastodebito.tipdoc, vta_gastodebito.idtipdocref, " _
        & " vta_gastodebito.fchemi AS fchemi, [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc] AS numdoc, vta_gastodebito.imptotdoc AS imptotdoc, " _
        & " vta_gastodebito.tc, 0 AS debedol, IIf([vta_gastodebito]![idmon]=2,[vta_gastodebito]![imptotdoc],0) AS haberdol, 0 AS debesol, " _
        & " IIf([vta_gastodebito]![idmon]=1,[vta_gastodebito]![imptotdoc],0) AS habersol, vta_gastodebito.numerodocref FROM con_diario RIGHT JOIN " _
        & " vta_gastodebito ON con_diario.idmov = vta_gastodebito.id WHERE (((con_diario.imphabsol)<>0) AND ((vta_gastodebito.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (vta_gastodebito.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=41) AND ((vta_gastodebito.tipdoc)=126)) " _
        & " OR (((vta_gastodebito.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_gastodebito.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=41) " _
        & " AND ((vta_gastodebito.tipdoc)=126) AND ((con_diario.imphabdol)<>0)) "
        
        xSqlRes1 = xSqlRes1 & " UNION  " _
        & " SELECT let_letra.idclipro, con_diario.idcue, let_letra.idmon, let_letra.tipdoc, let_letra.idtipdocref, let_letradet.fchemi, " _
        & " [let_letradet]![numser] & '-' & [let_letradet]![numdoc] AS numdoc, let_letradet.implet AS impdoc, let_letra.tc, 0 AS debedol, " _
        & " IIf([let_letra]![idmon]=2,[let_letradet]![implet],0) AS haberdol, 0 AS debesol, IIf([let_letra]![idmon]=1,[let_letradet]![implet],0) AS habersol, " _
        & " [let_letra]![idaduana] & [let_letra]![idregimen] & [let_letra]![anoorden] & [let_letra]![numorden] AS numerodocref FROM let_letra LEFT JOIN " _
        & " (con_diario RIGHT JOIN let_letradet ON (con_diario.correlativo = let_letradet.corr) AND (con_diario.idmov = let_letradet.idlet)) " _
        & " ON let_letra.id = let_letradet.idlet WHERE (((let_letra.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (let_letra.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((con_diario.idlib)=37) AND ((con_diario.impdebsol)<>0)) OR (((let_letra.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (let_letra.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((con_diario.idlib)=37) AND ((con_diario.impdebdol)<>0)) " _
        & " ) AS movimiento " _
        & " LEFT JOIN mae_cliente ON movimiento.idcli = mae_cliente.id) " _
        & " LEFT JOIN mae_docreferencia ON movimiento.idtipdocref = mae_docreferencia.id GROUP BY movimiento.idcli, mae_cliente.nombre, movimiento.numerodocref, " _
        & " mae_docreferencia.descripcion, mae_cliente.numruc "
        
        '"Having (((Sum([movimiento]![debedol] - [movimiento]![haberdol])) <> 0)) " _
        & " Or (((Sum([movimiento]![debesol] - [movimiento]![habersol])) <> 0))"
             
        'ARMAMOS LAS SENTENCIA WHERE PARA LAS CONSULTAS
        xCadWhere = ""
        If Fg2.Rows <> 0 Then
            xCadWhere = " ("
            For A = 0 To Fg2.Rows - 1
                xCadWhere = xCadWhere & "(movimiento.idcli = " & NulosC(Fg2.TextMatrix(A, 2)) & ")"
                If A = Fg2.Rows - 1 Then
                    xCadWhere = xCadWhere & ")"
                    Exit For
                End If
                xCadWhere = xCadWhere & " OR "
            Next A
        End If
             
        If Fg2.Rows = 0 Then
            ' mostramos todos los clientes
            xSqlRes1 = xSqlRes1 & " Having (((Sum([movimiento]![debedol] - [movimiento]![haberdol])) <> 0)) " _
                & " Or (((Sum([movimiento]![debesol] - [movimiento]![habersol])) <> 0))"
        Else
            ' filtramos solo los clientes que se especifican el el grid de clientes
            xSqlRes1 = xSqlRes1 & " Having (" & xCadWhere _
                & " AND ((Sum([movimiento]![debedol]-[movimiento]![haberdol]))<>0)) " _
                & " OR (" & xCadWhere & " AND ((Sum([movimiento]![debesol]-[movimiento]![habersol]))<>0))"
        End If
        
        
        
    '-------------------------------------------------------------
    'XONSULTA PARA TRAER EL RESUMEN DE LOS PENDIENTES DEL APERTURA
    xSqlRes2 = "SELECT con_provicionesdetdoc.idclipro AS idcli, mae_cliente.nombre, (Sum(IIf([con_provicionesdetdoc]![idmon]=2, IIf([con_provicionesdet]![tipo]=-1, " _
        & " [con_provicionesdetdoc]![impdoc],0),0)))-(Sum(IIf([con_provicionesdetdoc]![idmon]=2,IIf([con_provicionesdet]![tipo]=0,[con_provicionesdetdoc]![impdoc],0),0))) AS saldodol, " _
        & " (Sum(IIf([con_provicionesdetdoc]![idmon]=1,IIf([con_provicionesdet]![tipo]=-1,[con_provicionesdetdoc]![impdoc],0),0)))-(Sum(IIf([con_provicionesdetdoc]![idmon]=1, " _
        & " IIf([con_provicionesdet]![tipo]=0,[con_provicionesdetdoc]![impdoc],0),0))) AS saldosol, con_provicionesdetdoc.numdocref as numerodocref, mae_docreferencia.descripcion AS docref, " _
        & " mae_cliente.numruc FROM con_proviciones RIGHT JOIN ((con_provicionesdet RIGHT JOIN (((mae_cliente RIGHT JOIN (con_provicionesdetdoc LEFT JOIN con_planctas " _
        & " ON con_provicionesdetdoc.idcue = con_planctas.id) ON mae_cliente.id = con_provicionesdetdoc.idclipro) LEFT JOIN mae_documento ON con_provicionesdetdoc.tipdoc = mae_documento.id) " _
        & " LEFT JOIN mae_moneda ON con_provicionesdetdoc.idmon = mae_moneda.id) ON (con_provicionesdet.id = con_provicionesdetdoc.idpro) AND " _
        & " (con_provicionesdet.idcuen = con_provicionesdetdoc.idcue)) LEFT JOIN mae_docreferencia ON con_provicionesdetdoc.idtipdocref = mae_docreferencia.id) " _
        & " ON con_proviciones.id = con_provicionesdet.id GROUP BY con_provicionesdetdoc.idclipro, mae_cliente.nombre, con_provicionesdetdoc.numdocref, " _
        & " mae_docreferencia.descripcion, mae_cliente.numruc, con_proviciones.idmes "
        
        'ARMAMOS LAS SENTENCIA WHERE PARA LAS CONSULTAS
        xCadWhere = ""
        If Fg2.Rows <> 0 Then
            xCadWhere = " ("
            For A = 0 To Fg2.Rows - 1
                xCadWhere = xCadWhere & "(con_provicionesdetdoc.idclipro = " & NulosC(Fg2.TextMatrix(A, 2)) & ")"
                If A = Fg2.Rows - 1 Then
                    xCadWhere = xCadWhere & ")"
                    Exit For
                End If
                xCadWhere = xCadWhere & " OR "
            Next A
        End If
        
        If Fg2.Rows = 0 Then
            ' mostramos todos los clientes
            xSqlRes2 = xSqlRes2 & " Having ((((Sum(IIf([con_provicionesdetdoc]![idmon] = 2, IIf([con_provicionesdet]![tipo] = -1, [con_provicionesdetdoc]![impdoc], 0), 0))) - (Sum(IIf([con_provicionesdetdoc]![idmon] = 2, " _
                & " IIf([con_provicionesdet]![tipo] = 0, [con_provicionesdetdoc]![impdoc], 0), 0)))) <> 0) And ((con_proviciones.idmes) = 0)) Or " _
                & " ((((Sum(IIf([con_provicionesdetdoc]![idmon] = 1, IIf([con_provicionesdet]![tipo] = -1, [con_provicionesdetdoc]![impdoc], 0), 0))) - (Sum(IIf([con_provicionesdetdoc]![idmon] = 1, " _
                & " IIf([con_provicionesdet]![tipo] = 0, [con_provicionesdetdoc]![impdoc], 0), 0)))) <> 0)) '"
                'ORDER BY mae_cliente.nombre, con_provicionesdetdoc.numdocref"
        Else
            ' filtramos solo los clientes que se especifican el el grid de clientes
            xSqlRes2 = xSqlRes2 & "Having (" & xCadWhere & " And (((Sum(IIf([con_provicionesdetdoc]![idmon] = 2, IIf([con_provicionesdet]![tipo] = -1, " _
                & " [con_provicionesdetdoc]![impdoc], 0), 0))) - (Sum(IIf([con_provicionesdetdoc]![idmon] = 2, IIf([con_provicionesdet]![tipo] = 0, " _
                & " [con_provicionesdetdoc]![impdoc], 0), 0)))) <> 0) And ((con_proviciones.idmes) = 0)) " _
                & " Or (" & xCadWhere & " And (((Sum(IIf([con_provicionesdetdoc]![idmon] = 1, IIf([con_provicionesdet]![tipo] = -1, " _
                & " [con_provicionesdetdoc]![impdoc], 0), 0))) - (Sum(IIf([con_provicionesdetdoc]![idmon] = 1, IIf([con_provicionesdet]![tipo] = 0, " _
                & " [con_provicionesdetdoc]![impdoc], 0), 0)))) <> 0)) "
                'ORDER BY mae_cliente.nombre, con_provicionesdetdoc.numdocref"
        End If

    '---------------------------------------------
    'CARGAMOS EL DETALLE DEL AÑO ACTUAL DE TRABAJO

    xSqlDet1 = "SELECT [movimiento].[idcli], [mae_cliente].[nombre], [movimiento].[idcue], [con_planctas].[cuenta], [con_planctas].[descripcion], " _
        & " [mae_documento].[abrev], [movimiento].[fchemi], [mae_moneda].[simbolo], [movimiento].[numdoc], [movimiento].[impdoc], " _
        & " IIf([movimiento].[tc]<>0,[movimiento].[tc],[con_tc]![impven]) AS tc, [movimiento].[debedol], [movimiento].[haberdol], [movimiento].[debesol], " _
        & " [movimiento].[habersol], [movimiento].[numerodocref], [mae_cliente].[numruc] FROM ((((( " _
        & " ( " _
        & " SELECT vta_ventas.idcli, con_diario.idcue, vta_ventas.idmon, vta_ventas.tipdoc, vta_ventas.idtipdocref, vta_ventas.fchdoc AS fchemi, " _
        & " [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.imptotdoc AS impdoc, vta_ventas.tc, 0 AS debedol, " _
        & " IIf([vta_ventas]![idmon]=2,[vta_ventas]![imptotdoc],0) AS haberdol, 0 AS debesol, IIf([vta_ventas]![idmon]=1,[vta_ventas]![imptotdoc],0) AS habersol, " _
        & " vta_ventas.numerodocref FROM con_diario RIGHT JOIN vta_ventas ON con_diario.idmov = vta_ventas.id WHERE (((vta_ventas.tipdoc)=7) " _
        & " AND ((con_diario.imphabsol)<>0) AND ((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND " _
        & " ((con_diario.idlib)=2) AND ((vta_ventas.anulado)=0)) OR (((vta_ventas.tipdoc)=7) AND ((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=2) AND ((vta_ventas.anulado)=0) AND ((con_diario.imphabdol)<>0))" _
        & " UNION " _
        & " SELECT vta_ventas.idcli, con_diario.idcue, vta_ventas.idmon, vta_ventas.tipdoc, vta_ventas.idtipdocref, vta_ventas.fchdoc AS fchemi, " _
        & " [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.imptotdoc AS impdoc, vta_ventas.tc, IIf([vta_ventas]![idmon]=2,[vta_ventas]![imptotdoc],0) AS debedol, " _
        & " 0 AS haberdol, IIf([vta_ventas]![idmon]=1,[vta_ventas]![imptotdoc],0) AS debesol, 0 AS habersol, vta_ventas.numerodocref FROM con_diario RIGHT JOIN " _
        & " vta_ventas ON con_diario.idmov = vta_ventas.id WHERE (((vta_ventas.tipdoc)<>7) AND ((con_diario.impdebsol)<>0) AND ((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=2) AND ((vta_ventas.anulado)=0)) OR (((vta_ventas.tipdoc)<>7) " _
        & " AND ((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=2) " _
        & " AND ((vta_ventas.anulado)=0) AND ((con_diario.impdebdol)<>0)) "

    xSqlDet1 = xSqlDet1 & " UNION " _
        & " SELECT vta_gastodebito.idcli, con_diario.idcue, vta_gastodebito.idmon, vta_gastodebito.tipdoc, vta_gastodebito.idtipdocref, vta_gastodebito.fchemi AS fchemi, " _
        & " [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc] AS numdoc, vta_gastodebito.imptotdoc AS impdoc, vta_gastodebito.tc, " _
        & " IIf([vta_gastodebito]![idmon]=2,[vta_gastodebito]![imptotdoc],0) AS debedol, 0 AS haberdol, IIf([vta_gastodebito]![idmon]=1,[vta_gastodebito]![imptotdoc],0) AS debesol, " _
        & " 0 AS habersol, vta_gastodebito.numerodocref FROM con_diario RIGHT JOIN vta_gastodebito ON con_diario.idmov = vta_gastodebito.id " _
        & " WHERE (((vta_gastodebito.tipdoc)=120) AND ((con_diario.impdebsol)<>0) AND ((vta_gastodebito.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (vta_gastodebito.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=41)) OR (((vta_gastodebito.tipdoc)=120) " _
        & " AND ((vta_gastodebito.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_gastodebito.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idlib)=41) " _
        & " AND ((con_diario.impdebdol)<>0)) " _
        & " UNION " _
        & " SELECT vta_gastodebito.idcli, con_diario.idcue, vta_gastodebito.idmon, vta_gastodebito.tipdoc, vta_gastodebito.idtipdocref, vta_gastodebito.fchemi AS fchemi, " _
        & " [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc] AS numdoc, vta_gastodebito.imptotdoc AS imptotdoc, vta_gastodebito.tc, 0 AS debedol, " _
        & " IIf([vta_gastodebito]![idmon]=2,[vta_gastodebito]![imptotdoc],0) AS haberdol, 0 AS debesol, IIf([vta_gastodebito]![idmon]=1,[vta_gastodebito]![imptotdoc],0) AS habersol, " _
        & " vta_gastodebito.numerodocref FROM con_diario RIGHT JOIN vta_gastodebito ON con_diario.idmov = vta_gastodebito.id WHERE (((con_diario.imphabsol)<>0) " _
        & " AND ((vta_gastodebito.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_gastodebito.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((con_diario.idlib)=41) AND ((vta_gastodebito.tipdoc)=126)) OR (((vta_gastodebito.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_gastodebito.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((con_diario.idlib)=41) AND ((vta_gastodebito.tipdoc)=126) AND ((con_diario.imphabdol)<>0)) "

    xSqlDet1 = xSqlDet1 & " UNION " _
        & " SELECT let_letra.idclipro, con_diario.idcue, let_letra.idmon, let_letra.tipdoc, let_letra.idtipdocref, let_letradet.fchemi, " _
        & " [let_letradet]![numser] & '-' & [let_letradet]![numdoc] AS numdoc, let_letradet.implet AS impdoc, let_letra.tc, 0 AS debedol, " _
        & " IIf([let_letra]![idmon]=2,[let_letradet]![implet],0) AS haberdol, 0 AS debesol, IIf([let_letra]![idmon]=1,[let_letradet]![implet],0) AS habersol, " _
        & " [let_letra]![idaduana] & [let_letra]![idregimen] & [let_letra]![anoorden] & [let_letra]![numorden] AS numerodocref " _
        & " FROM let_letra LEFT JOIN (con_diario RIGHT JOIN let_letradet ON (con_diario.correlativo = let_letradet.corr) AND (con_diario.idmov = let_letradet.idlet)) " _
        & " ON let_letra.id = let_letradet.idlet WHERE (((let_letra.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (let_letra.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((con_diario.idlib)=37) AND ((con_diario.impdebsol)<>0)) OR (((let_letra.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (let_letra.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((con_diario.idlib)=37) AND ((con_diario.impdebdol)<>0))) AS movimiento" _
        & " LEFT JOIN mae_cliente ON [movimiento].[idcli]=[mae_cliente].[id]) LEFT JOIN con_planctas ON [movimiento].[idcue]=[con_planctas].[id]) " _
        & " LEFT JOIN mae_moneda ON [movimiento].[idmon]=[mae_moneda].[id]) LEFT JOIN mae_documento ON [movimiento].[tipdoc]=[mae_documento].[id]) " _
        & " LEFT JOIN mae_docreferencia ON [movimiento].[idtipdocref]=[mae_docreferencia].[id]) LEFT JOIN con_tc ON [movimiento].[fchemi]=[con_tc].[fecha]"

    'xSqlDet2 = "SELECT con_provicionesdetdoc.idclipro, mae_cliente.nombre, con_provicionesdetdoc.idcue, con_planctas.cuenta, con_planctas.descripcion, " _
        & " mae_documento.abrev, con_provicionesdetdoc.fchemi, mae_moneda.simbolo, [con_provicionesdetdoc]![numser] & '-' & [con_provicionesdetdoc]![numdoc] AS numdoc, " _
        & " con_provicionesdetdoc.impdoc, 3.3333 AS tc, IIf([con_provicionesdetdoc]![idmon]=2,IIf([con_provicionesdet]![tipo]=-1,[con_provicionesdetdoc]![impdoc],0),0) AS debedol, " _
        & " IIf([con_provicionesdetdoc]![idmon]=2,IIf([con_provicionesdet]![tipo]=0,[con_provicionesdetdoc]![impdoc],0),0) AS haberdol, IIf([con_provicionesdetdoc]![idmon]=1, " _
        & " IIf([con_provicionesdet]![tipo]=-1,[con_provicionesdetdoc]![impdoc],0),0) AS debesol, IIf([con_provicionesdetdoc]![idmon]=1,IIf([con_provicionesdet]![tipo]=0, " _
        & " [con_provicionesdetdoc]![impdoc],0),0) AS habersol, con_provicionesdetdoc.numdocref, mae_cliente.numruc FROM con_proviciones RIGHT JOIN (con_provicionesdet " _
        & " RIGHT JOIN (((mae_cliente RIGHT JOIN (con_provicionesdetdoc LEFT JOIN con_planctas ON con_provicionesdetdoc.idcue = con_planctas.id) " _
        & " ON mae_cliente.id = con_provicionesdetdoc.idclipro) LEFT JOIN mae_documento ON con_provicionesdetdoc.tipdoc = mae_documento.id) LEFT JOIN mae_moneda " _
        & " ON con_provicionesdetdoc.idmon = mae_moneda.id) ON (con_provicionesdet.idcuen = con_provicionesdetdoc.idcue) AND (con_provicionesdet.id = con_provicionesdetdoc.idpro)) " _
        & " ON con_proviciones.id = con_provicionesdet.id Where (((con_proviciones.idmes) = 0)) ORDER BY con_provicionesdetdoc.numdocref"

    xSqlDet2 = "SELECT con_provicionesdetdoc.idclipro as idcli, mae_cliente.nombre, con_provicionesdetdoc.idcue, con_planctas.cuenta, con_planctas.descripcion, " _
        & " mae_documento.abrev, con_provicionesdetdoc.fchemi, mae_moneda.simbolo, [con_provicionesdetdoc]![numser] & '-' & [con_provicionesdetdoc]![numdoc] AS numdoc, " _
        & " con_provicionesdetdoc.impdoc, con_tc.impven AS tc, IIf([con_provicionesdetdoc]![idmon]=2,IIf([con_provicionesdet]![tipo]=-1,[con_provicionesdetdoc]![impdoc],0),0) AS debedol, " _
        & " IIf([con_provicionesdetdoc]![idmon]=2,IIf([con_provicionesdet]![tipo]=0,[con_provicionesdetdoc]![impdoc],0),0) AS haberdol, IIf([con_provicionesdetdoc]![idmon]=1," _
        & " IIf([con_provicionesdet]![tipo]=-1,[con_provicionesdetdoc]![impdoc],0),0) AS debesol, IIf([con_provicionesdetdoc]![idmon]=1,IIf([con_provicionesdet]![tipo]=0, " _
        & " [con_provicionesdetdoc]![impdoc],0),0) AS habersol, con_provicionesdetdoc.numdocref as numerodocref, mae_cliente.numruc FROM (con_proviciones RIGHT JOIN (con_provicionesdet " _
        & " RIGHT JOIN (((mae_cliente RIGHT JOIN (con_provicionesdetdoc LEFT JOIN con_planctas ON con_provicionesdetdoc.idcue = con_planctas.id) " _
        & " ON mae_cliente.id = con_provicionesdetdoc.idclipro) LEFT JOIN mae_documento ON con_provicionesdetdoc.tipdoc = mae_documento.id) LEFT JOIN mae_moneda " _
        & " ON con_provicionesdetdoc.idmon = mae_moneda.id) ON (con_provicionesdet.id = con_provicionesdetdoc.idpro) AND (con_provicionesdet.idcuen = con_provicionesdetdoc.idcue)) " _
        & " ON con_proviciones.id = con_provicionesdet.id) LEFT JOIN con_tc ON con_provicionesdetdoc.fchemi = con_tc.fecha Where (((con_proviciones.idmes) = 0)) "
        '& " ORDER BY con_provicionesdetdoc.numdocref"
    
    '-------------------------------------------
    ' CARGAMOS EL RESUMEN DE LA CUENTA CORRIENTE
    ' SI SE MUESTRA SOLO EL APERTURA
    If Check1.Value = 1 And Check2.Value = 0 Then
        RST_Busq RstRes1, xSqlRes2, xCon
    End If
    
    ' SI SE MUESTRA SOLO EL PERIODO ACTUAL
    If Check1.Value = 0 And Check2.Value = 1 Then
        RST_Busq RstRes1, xSqlRes1, xCon
    End If
    
    ' SI SE MUESTRA EL APERTURA Y EL PERIODO ACTUAL
    If Check1.Value = 1 And Check2.Value = 1 Then
        RST_Busq RstRes1, xSqlRes1 & Chr(13) & " UNION " & Chr(13) & xSqlRes2, xCon
    End If

    
    'RST_Busq RstRes1, xSqlRes1, xCon
    'RST_Busq RstRes2, xSqlRes2, xCon
    
    '-------------------------------------------
    ' CARGAMOS EL DETALLE DE LA CUENTA CORRIENTE
    ' SI SE MUESTRA SOLO EL APERTURA
    If Check1.Value = 1 And Check2.Value = 0 Then
        RST_Busq RstDet1, xSqlDet2, xCon
    End If
    
    ' SI SE MUESTRA SOLO EL PERIODO ACTUAL
    If Check1.Value = 0 And Check2.Value = 1 Then
        RST_Busq RstDet1, xSqlDet1, xCon
    End If
    
    ' SI SE MUESTRA EL APERTURA Y EL PERIODO ACTUAL
    If Check1.Value = 1 And Check2.Value = 1 Then
        RST_Busq RstDet1, xSqlDet1 & Chr(13) & " UNION " & Chr(13) & xSqlDet2, xCon
    End If

'    RST_Busq RstDet2, xSqlDet2, xCon
    
    RstRes1.ActiveConnection = Nothing
    'RstRes2.ActiveConnection = Nothing
    RstDet1.ActiveConnection = Nothing
    'RstDet2.ActiveConnection = Nothing
    
'    'MOSTRAMOS SOLO APERTURA
'    If Check1.Value = 1 And Check2.Value = 0 Then
'        If RstRes1.RecordCount <> 0 Then
'            RstRes1.MoveFirst
'            For A = 1 To RstRes1.RecordCount
'                RstRes1.Delete
'                RstRes1.MoveNext
'                If RstRes1.EOF = True Then Exit For
'            Next A
'
''            RstDet1.MoveFirst
''            For A = 1 To RstDet1.RecordCount
''                RstDet1.Delete
''                RstDet1.MoveNext
''                If RstDet1.EOF = True Then Exit For
''            Next A
'        End If
'    End If
    
'    'MOSTRAMOS SOLO EL PERIODO ACTUAL
'    If Check1.Value = 0 And Check2.Value = 1 Then
'        If RstRes2.RecordCount <> 0 Then
'            RstRes2.MoveFirst
'            For A = 1 To RstRes2.RecordCount
'                RstRes2.Delete
'                RstRes2.MoveNext
'                If RstRes2.EOF = True Then Exit For
'            Next A
'
''            RstDet2.MoveFirst
''            For A = 1 To RstDet2.RecordCount
''                RstDet2.Delete
''                RstDet2.MoveNext
''                If RstDet2.EOF = True Then Exit For
''            Next A
'        End If
'    End If
    
    Frame3.Visible = True
    
'    'PREGUTAMOS SI HAY MOVIMIENTOS DE APERTURA - CABECERA
'    If RstRes2.RecordCount <> 0 Then
'        RstRes2.MoveFirst
'
'        ProgressBar1.Max = RstRes2.RecordCount
'
'        ' SI HAY MOVIMIENTOS DE APERTURA, LOS AGREGAMOS AL PRIMER RECORSET
'        For A = 1 To RstRes2.RecordCount
'            ProgressBar1.Value = A
'            'Frame3.Refresh
'
'            RstRes1.AddNew
'            RstRes1("idcli") = RstRes2("idclipro")
'            RstRes1("nombre") = RstRes2("nombre")
'            RstRes1("saldodol") = RstRes2("saldodol")
'            RstRes1("saldosol") = RstRes2("saldosol")
'            RstRes1("numerodocref") = RstRes2("numdocref")
'            RstRes1("docref") = RstRes2("docref")
'            RstRes1("numruc") = RstRes2("numruc")
'            RstRes2.MoveNext
'            If RstRes2.EOF = True Then Exit For
'        Next A
'    End If
    
    
'    'PREGUTAMOS SI HAY MOVIMIENTOS DE APERTURA - DETALLE
'    If RstDet2.RecordCount <> 0 Then
''         CARGAR_RST_TMP RstDet1, RstDet2
'        RstDet2.MoveFirst
'        ProgressBar1.Max = RstDet2.RecordCount
'        ' SI HAY MOVIMIENTOS DE APERTURA, LOS AGREGAMOS AL PRIMER RECORSET
'        For A = 1 To RstDet2.RecordCount
'            ProgressBar1.Value = A
'            'Frame3.Refresh
'            RstDet1.AddNew
'            RstDet1("idcli") = RstDet2("idcli")
'            RstDet1("nombre") = RstDet2("nombre")
'            RstDet1("idcue") = RstDet2("idcue")
'            RstDet1("cuenta") = RstDet2("cuenta")
'            RstDet1("descripcion") = RstDet2("descripcion")
'            RstDet1("abrev") = RstDet2("abrev")
'            RstDet1("fchemi") = RstDet2("fchemi")
'            RstDet1("simbolo") = RstDet2("simbolo")
'            RstDet1("numdoc") = RstDet2("numdoc")
'            RstDet1("impdoc") = RstDet2("impdoc")
'            'RstDet1("tc") = NulosN(RstDet2("tc"))
'            RstDet1("debedol") = RstDet2("debedol")
'            RstDet1("haberdol") = RstDet2("haberdol")
'            RstDet1("debesol") = RstDet2("debesol")
'            RstDet1("habersol") = RstDet2("habersol")
'            RstDet1("numerodocref") = RstDet2("numerodocref")
'            RstDet1("numruc") = RstDet2("numruc")
'            RstDet2.MoveNext
'            If RstDet2.EOF = True Then Exit For
'        Next A
'    End If
    
    RstRes1.Sort = "nombre, docref, numerodocref"
    RstDet1.Sort = "nombre, numerodocref, fchemi"
    
    Fg1.Rows = 2
    Fg2.Rows = 2
    RstRes1.MoveFirst
    Dim xIdCli As Integer
    Dim xTotalDol As Double
    Dim xTotalSol As Double
    
    xIdCli = RstRes1("idcli")
    
    ProgressBar1.Max = RstRes1.RecordCount
    
    For A = 1 To RstRes1.RecordCount
        Fg1.Rows = Fg1.Rows + 1
        ProgressBar1.Value = A
        'Frame3.Refresh
        
        If A = 1 Then
            GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 2, "EMPRESA    :  " & RstRes1("nombre"), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
            GRID_COMBINAR Fg1, Fg1.Rows - 1, 3, Fg1.Rows - 1, 4, "R.U.C. Nº  :  " & RstRes1("numruc"), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
            
            AgregarCabeceraFG2 RstRes1("nombre"), RstRes1("numruc")
            ImprimirDetalleFG2 RstRes1("idcli")
            Fg1.Rows = Fg1.Rows + 1
        Else
            If RstRes1("idcli") <> xIdCli Then
                'ESCRIBIMOS EL TOTAL DE LA EMPRESA
                'Fg1.Rows = Fg1.Rows + 1
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 2, &H0&, True, &HE2FEFB, "TOTAL ==> "
                If xTotalDol < 0 Then
                    FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &HFF&, True, &HE2FEFB, Format(xTotalDol, "0.00")
                Else
                    FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &HFF0000, True, &HE2FEFB, Format(xTotalDol, "0.00")
                End If
                
                If xTotalSol < 0 Then
                    FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &HFF&, True, &HE2FEFB, Format(xTotalSol, "0.00")
                Else
                    FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &HFF0000, True, &HE2FEFB, Format(xTotalSol, "0.00")
                End If
                
                xTotalDol = 0
                xTotalSol = 0
                
                Fg1.Rows = Fg1.Rows + 2
                GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 2, "EMPRESA    :  " & RstRes1("nombre"), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                GRID_COMBINAR Fg1, Fg1.Rows - 1, 3, Fg1.Rows - 1, 4, "R.U.C. Nº  :  " & RstRes1("numruc"), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                
                AgregarCabeceraFG2 RstRes1("nombre"), RstRes1("numruc")
                ImprimirDetalleFG2 RstRes1("idcli")
                xIdCli = RstRes1("idcli")
                Fg1.Rows = Fg1.Rows + 1
            End If
        End If
        
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstRes1("docref"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstRes1("numerodocref"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(RstRes1("saldodol"), "0.00")
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(RstRes1("saldosol"), "0.00")
        
        xTotalDol = xTotalDol + RstRes1("saldodol")
        xTotalSol = xTotalSol + RstRes1("saldosol")
        RstRes1.MoveNext
        If RstRes1.EOF = True Then
            'ESCRIBIMOS EL TOTAL DE LA EMPRESA
            Fg1.Rows = Fg1.Rows + 1
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 2, &H0&, True, &HE2FEFB, "TOTAL ==> "
            If xTotalDol < 0 Then
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &HFF&, True, &HE2FEFB, Format(xTotalDol, "0.00")
            Else
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &HFF0000, True, &HE2FEFB, Format(xTotalDol, "0.00")
            End If
            
            If xTotalSol < 0 Then
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &HFF&, True, &HE2FEFB, Format(xTotalSol, "0.00")
            Else
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &HFF0000, True, &HE2FEFB, Format(xTotalSol, "0.00")
            End If
            
            Exit For
        End If
    Next A
    
    Frame3.Visible = False
End Sub

Sub AgregarCabeceraFG2(NombreCliente As String, NumRuc As String)
    Fg3.Rows = Fg3.Rows + 1
    
    GRID_COMBINAR Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, 6, "CLIENTE ==> " & NombreCliente, flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
    GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 8, "Nº DE RUC ==> ", flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
    GRID_COMBINAR Fg3, Fg3.Rows - 1, 9, Fg3.Rows - 1, 10, NumRuc, flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
    
End Sub

Sub ImprimirDetalleFG2(IdCliente As Integer)
    Dim B As Integer
    Dim xNumOrd As String
    Dim xForeColor As Long
    Dim xTotDebSol, xTotHabSol, xTotDebDol, xTotHabDol As Double
    Dim xxTotalSol, xxTotalDol As Double
    
    RstDet1.Filter = adFilterNone
    RstDet1.Filter = "idcli = " & IdCliente & ""
    
    If RstDet1.RecordCount <> 0 Then
        xNumOrd = NulosC(RstDet1("numerodocref"))
        
        For B = 1 To RstDet1.RecordCount
            If B = 1 Then
                Fg3.Rows = Fg3.Rows + 1
                'Fg3.TextMatrix(Fg3.Rows - 1, 1) = ""
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, 2, "DOC. REF. :  ORDEN DE DESPACHO  ", flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 9, "Nº DOC. REF. : " & NulosC(RstDet1("numerodocref")), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                
                Fg3.Rows = Fg3.Rows + 1
            Else
                If NulosC(RstDet1("numerodocref")) <> xNumOrd Then
                    Fg3.Rows = Fg3.Rows + 1
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 8, "TOTAL ==>", flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 9, &H0&, True, &HE2FEFB, Format(NulosN(xTotDebDol), "0.00")
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 10, &H0&, True, &HE2FEFB, Format(NulosN(xTotHabDol), "0.00")
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 11, &H0&, True, &HE2FEFB, Format(NulosN(xTotDebSol), "0.00")
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 12, &H0&, True, &HE2FEFB, Format(NulosN(xTotHabSol), "0.00")
                    
                    Fg3.Rows = Fg3.Rows + 1
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 8, "SALDO ==>", flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
                    If (xTotDebDol - xTotHabDol) < 0 Then xForeColor = &HFF&    'rojo
                    If (xTotDebDol - xTotHabDol) > 0 Then xForeColor = &HFF0000 'azul
                    If (xTotDebDol - xTotHabDol) = 0 Then xForeColor = &H0&     'negro
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 10, xForeColor, True, &HE2FEFB, Format(xTotDebDol - xTotHabDol, "0.00")
                    
                    If (xTotDebSol - xTotHabSol) < 0 Then xForeColor = &HFF&    'rojo
                    If (xTotDebSol - xTotHabSol) > 0 Then xForeColor = &HFF0000 'azul
                    If (xTotDebSol - xTotHabSol) = 0 Then xForeColor = &H0&     'negro
                    'Fg3.TextMatrix(Fg3.Rows - 1, 12) = Format(xTotDebSol - xTotHabSol, "0.00")
                    FORMATO_CELDA Fg3, Fg3.Rows - 1, 12, xForeColor, True, &HE2FEFB, Format(xTotDebSol - xTotHabSol, "0.00")
                    
                    xxTotalDol = xxTotalDol + (xTotDebDol - xTotHabDol)
                    xxTotalSol = xxTotalSol + (xTotDebSol - xTotHabSol)
                    
                    xTotDebDol = 0:    xTotHabDol = 0:     xTotDebSol = 0:     xTotHabSol = 0
                    Fg3.Rows = Fg3.Rows + 2
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, 2, "DOC. REF. :  ORDEN DE DESPACHO  ", flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 10, "Nº DOC. REF. : " & NulosC(RstDet1("numerodocref")), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                    
                    xNumOrd = NulosC(RstDet1("numerodocref"))
                End If
                Fg3.Rows = Fg3.Rows + 1
            End If
            Fg3.TextMatrix(Fg3.Rows - 1, 1) = RstDet1("cuenta")
            Fg3.TextMatrix(Fg3.Rows - 1, 2) = RstDet1("descripcion")
            Fg3.TextMatrix(Fg3.Rows - 1, 3) = RstDet1("fchemi")
            Fg3.TextMatrix(Fg3.Rows - 1, 4) = RstDet1("abrev")
            Fg3.TextMatrix(Fg3.Rows - 1, 5) = RstDet1("numdoc")
            Fg3.TextMatrix(Fg3.Rows - 1, 6) = RstDet1("simbolo")
            Fg3.TextMatrix(Fg3.Rows - 1, 7) = Format(NulosN(RstDet1("tc")), "0.000")
            Fg3.TextMatrix(Fg3.Rows - 1, 8) = Format(RstDet1("impdoc"), "0.00")
            
            If RstDet1("abrev") = "NC" Or RstDet1("abrev") = "LGC" Then
                Fg3.TextMatrix(Fg3.Rows - 1, 10) = Format(RstDet1("debedol"), "0.00")
                Fg3.TextMatrix(Fg3.Rows - 1, 9) = Format(RstDet1("haberdol"), "0.00")
                Fg3.TextMatrix(Fg3.Rows - 1, 12) = Format(RstDet1("debesol"), "0.00")
                Fg3.TextMatrix(Fg3.Rows - 1, 11) = Format(RstDet1("habersol"), "0.00")
                
                xTotDebDol = xTotDebDol + RstDet1("haberdol")
                xTotHabDol = xTotHabDol + RstDet1("debedol")
                xTotDebSol = xTotDebSol + RstDet1("habersol")
                xTotHabSol = xTotHabSol + RstDet1("debesol")
            Else
                Fg3.TextMatrix(Fg3.Rows - 1, 9) = Format(RstDet1("debedol"), "0.00")
                Fg3.TextMatrix(Fg3.Rows - 1, 10) = Format(RstDet1("haberdol"), "0.00")
                Fg3.TextMatrix(Fg3.Rows - 1, 11) = Format(RstDet1("debesol"), "0.00")
                Fg3.TextMatrix(Fg3.Rows - 1, 12) = Format(RstDet1("habersol"), "0.00")
                
                xTotDebDol = xTotDebDol + RstDet1("debedol")
                xTotHabDol = xTotHabDol + RstDet1("haberdol")
                xTotDebSol = xTotDebSol + RstDet1("debesol")
                xTotHabSol = xTotHabSol + RstDet1("habersol")
            End If
            RstDet1.MoveNext
            If RstDet1.EOF = True Then
                Fg3.Rows = Fg3.Rows + 1
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 8, "TOTAL ==>", flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 9, &H0&, True, &HE2FEFB, Format(NulosN(xTotDebDol), "0.00")
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 10, &H0&, True, &HE2FEFB, Format(NulosN(xTotHabDol), "0.00")
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 11, &H0&, True, &HE2FEFB, Format(NulosN(xTotDebSol), "0.00")
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 12, &H0&, True, &HE2FEFB, Format(NulosN(xTotHabSol), "0.00")
                
                Fg3.Rows = Fg3.Rows + 1
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 8, "SALDO ==>", flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
                
                If (xTotDebDol - xTotHabDol) < 0 Then xForeColor = &HFF&    'rojo
                If (xTotDebDol - xTotHabDol) > 0 Then xForeColor = &HFF0000 'azul
                If (xTotDebDol - xTotHabDol) = 0 Then xForeColor = &H0&     'negro
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 10, xForeColor, True, &HE2FEFB, Format(xTotDebDol - xTotHabDol, "0.00")
                
                If (xTotDebSol - xTotHabSol) < 0 Then xForeColor = &HFF&    'rojo
                If (xTotDebSol - xTotHabSol) > 0 Then xForeColor = &HFF0000 'azul
                If (xTotDebSol - xTotHabSol) = 0 Then xForeColor = &H0&     'negro
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 12, xForeColor, True, &HE2FEFB, Format(xTotDebSol - xTotHabSol, "0.00")
                
                xxTotalDol = xxTotalDol + (xTotDebDol - xTotHabDol)
                xxTotalSol = xxTotalSol + (xTotDebSol - xTotHabSol)
                
                Fg3.Rows = Fg3.Rows + 1
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 8, "TOTAL EMPRESA ==>", flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
                If (xxTotalDol < 0) Then xForeColor = &HFF&    'rojo
                If (xxTotalDol > 0) Then xForeColor = &HFF0000 'azul
                If (xxTotalDol = 0) Then xForeColor = &H0&     'negro
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 10, xForeColor, True, &HE2FEFB, Format(xxTotalDol, "0.00")
                
                If (xxTotalSol < 0) Then xForeColor = &HFF&    'rojo
                If (xxTotalSol > 0) Then xForeColor = &HFF0000 'azul
                If (xxTotalSol = 0) Then xForeColor = &H0&     'negro
                FORMATO_CELDA Fg3, Fg3.Rows - 1, 12, xForeColor, True, &HE2FEFB, Format(xxTotalSol, "0.00")
                                
                Fg3.Rows = Fg3.Rows + 1
                Exit For
            End If
        Next B
        
        RstDet1.MoveFirst
        For B = 1 To RstDet1.RecordCount
            RstDet1.Delete
            RstDet1.MoveNext
            If RstDet1.EOF = True Then Exit For
        Next B
    End If
End Sub
