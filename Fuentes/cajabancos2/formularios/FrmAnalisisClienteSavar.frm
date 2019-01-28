VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "Codejock.CommandBars.v12.0.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAnalisisClienteSavar 
   Caption         =   "Analisis de Cta Cte del Cliente  por Documento de Referencia"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14610
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   885
      Left            =   840
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   6180
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   405
         Left            =   60
         TabIndex        =   33
         Top             =   390
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   1
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
      _GridInfo       =   $"FrmAnalisisClienteSavar.frx":0000
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
         Begin VB.Frame FramResumen 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   4155
            Left            =   45
            TabIndex        =   7
            Top             =   45
            Width           =   14175
            Begin VSFlex7Ctl.VSFlexGrid Fg3 
               Height          =   4125
               Left            =   0
               TabIndex        =   8
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
               FormatString    =   $"FrmAnalisisClienteSavar.frx":0042
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
            TabIndex        =   5
            Top             =   45
            Width           =   14175
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   4110
               Left            =   0
               TabIndex        =   6
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
               FormatString    =   $"FrmAnalisisClienteSavar.frx":0118
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
         _GridInfo       =   $"FrmAnalisisClienteSavar.frx":01ED
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   1065
            Left            =   11280
            TabIndex        =   30
            Top             =   0
            Width           =   2985
            Begin VB.CommandButton Command2 
               Caption         =   "Eliminar OD No Marcadas"
               Height          =   360
               Left            =   105
               TabIndex        =   32
               Top             =   540
               Visible         =   0   'False
               Width           =   2070
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Agregar Documentos"
               Height          =   360
               Left            =   105
               TabIndex        =   31
               Top             =   180
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
               TabIndex        =   1
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
            TabIndex        =   22
            Top             =   0
            Width           =   3390
            Begin VB.Frame Frame66 
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   720
               Left            =   1755
               TabIndex        =   26
               Top             =   285
               Width           =   1575
               Begin VB.CheckBox Check2 
                  Caption         =   "Año Trabajo"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   28
                  Top             =   270
                  Width           =   1290
               End
               Begin VB.CheckBox Check1 
                  Caption         =   "Apertura"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   27
                  Top             =   45
                  Width           =   1170
               End
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Todos"
               Height          =   195
               Left            =   210
               TabIndex        =   25
               Top             =   780
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Cancelados"
               Height          =   195
               Left            =   210
               TabIndex        =   24
               Top             =   555
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Pendientes"
               Height          =   195
               Left            =   210
               TabIndex        =   23
               Top             =   330
               Visible         =   0   'False
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
               TabIndex        =   29
               Top             =   30
               Visible         =   0   'False
               Width           =   810
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   1065
            Left            =   3375
            TabIndex        =   19
            Top             =   0
            Width           =   4485
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   750
               Left            =   0
               TabIndex        =   20
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
               FormatString    =   $"FrmAnalisisClienteSavar.frx":024B
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
               TabIndex        =   21
               Top             =   30
               Width           =   600
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   1065
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   3360
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   870
               Locked          =   -1  'True
               TabIndex        =   14
               Text            =   "TxtIdMon"
               Top             =   690
               Width           =   615
            End
            Begin VB.CommandButton CmdBusMon 
               Enabled         =   0   'False
               Height          =   240
               Left            =   1230
               Picture         =   "FrmAnalisisClienteSavar.frx":029B
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   720
               Width           =   240
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   870
               TabIndex        =   12
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
               TabIndex        =   13
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
               TabIndex        =   18
               Top             =   690
               Width           =   1770
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Left            =   45
               TabIndex        =   17
               Top             =   720
               Width           =   585
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Venc."
               Height          =   195
               Left            =   45
               TabIndex        =   16
               Top             =   420
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Inicio"
               Height          =   195
               Left            =   45
               TabIndex        =   15
               Top             =   135
               Width           =   735
            End
         End
      End
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   4185
      Top             =   30
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmAnalisisClienteSavar.frx":03CD
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   3765
      Top             =   0
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "FrmAnalisisClienteSavar"
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
    FrmAgregaDoc.Show vbModal
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

'Private Sub Command3_Click()
'    Dim xCadWhere As String
'    Dim A As Integer
'    Fg1.Rows = 2
'    Fg3.Rows = 2
'
'    If Check1.Value = 0 And Check2.Value = 0 Then
'        MsgBox "No ha especificado que datos se van a mostrar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Check1.SetFocus
'        Exit Sub
'    End If
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
'    Fg1.MergeCells = flexMergeFixedOnly
'
'    TraerDatos
'End Sub

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
    FrameDetalle.BackColor = &H8000000F
    
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
        RST_Busq RstDet, "SELECT var_analisisctacte.*, mae_cliente.nombre, mae_cliente.numruc, mae_moneda.simbolo, mae_documento.abrev, mae_cliente.numruc, " _
            & " con_planctas.cuenta, con_planctas.descripcion FROM (((var_analisisctacte LEFT JOIN mae_cliente ON var_analisisctacte.idcli = mae_cliente.id) " _
            & " LEFT JOIN mae_moneda ON var_analisisctacte.idmon = mae_moneda.id) LEFT JOIN mae_documento ON var_analisisctacte.idtipdoc = mae_documento.id) " _
            & " LEFT JOIN con_planctas ON var_analisisctacte.idcue = con_planctas.id ORDER BY mae_cliente.nombre, var_analisisctacte.numdocref, var_analisisctacte.fchemi", xCon
    Else
        ' MOSTRAMOS SOLOS LOS CLIENTES ESPECIFICADOS
        Dim xCadWhere As String
        xCadWhere = "("
        For A = 0 To Fg2.Rows - 1
            xCadWhere = xCadWhere & "(var_analisisctacte.idcli = " & Fg2.TextMatrix(A, 2) & ")"
            If A = Fg2.Rows - 1 Then
                Exit For
            End If
            xCadWhere = xCadWhere & " OR "
        Next A
        xCadWhere = xCadWhere & ")"
        
        RST_Busq RstDet, "SELECT var_analisisctacte.*, mae_cliente.nombre, mae_cliente.numruc, mae_moneda.simbolo, mae_documento.abrev, mae_cliente.numruc, " _
            & " con_planctas.cuenta, con_planctas.descripcion FROM (((var_analisisctacte LEFT JOIN mae_cliente ON var_analisisctacte.idcli = mae_cliente.id) " _
            & " LEFT JOIN mae_moneda ON var_analisisctacte.idmon = mae_moneda.id) LEFT JOIN mae_documento ON var_analisisctacte.idtipdoc = mae_documento.id) " _
            & " LEFT JOIN con_planctas ON var_analisisctacte.idcue = con_planctas.id " _
            & " Where " & xCadWhere _
            & " ORDER BY mae_cliente.nombre, var_analisisctacte.numdocref, var_analisisctacte.fchemi", xCon

    End If


    Dim xFchIniPer As String
    xFchIniPer = "01/01/09"
    If Check1.Value = 1 And Check2.Value = 0 Then
        ' SOLO APERTURA
        RstDet.Filter = "fchemi < '" & xFchIniPer & "'"
    End If
    
    If Check1.Value = 0 And Check2.Value = 1 Then
        ' SOLO APERTURA
        RstDet.Filter = "fchemi >= '" & xFchIniPer & "'"
    End If
    
    ' MOSTRAMOS EL DETALLE DEL MOVIMIENTO
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        ImprimirDetalleFG2
    End If
    
    ' MOSTRAMOS EL RESUMEN
    Set RstDet = Nothing
    Dim xCadCondi As String
    
    If Check1.Value = 1 And Check2.Value = 0 Then
        ' SOLO APERTURA
        xCadCondi = " WHERE (var_analisisctacte.fchemi < cdate('" & xFchIniPer & "'))"
    End If
    
    If Check1.Value = 0 And Check2.Value = 1 Then
        ' SOLO DEL AÑO DE TRABAJO
        xCadCondi = " WHERE (var_analisisctacte.fchemi >= cdate('" & xFchIniPer & "'))"
    End If
    
    If Check1.Value = 1 And Check2.Value = 1 Then
        xCadCondi = ""
    End If
    
    If Fg2.Rows = 0 Then
        ' MOSTRAMOS EL RESUMEN DE TODOS LOS CLIENTES
        xCad = "SELECT var_analisisctacte.idcli, mae_cliente.nombre, mae_cliente.numruc, var_analisisctacte.numdocref, Sum(var_analisisctacte.debesol) AS SumaDedebesol, " _
            & " Sum(var_analisisctacte.habersol) AS SumaDehabersol, Sum(var_analisisctacte.debedol) AS SumaDedebedol, Sum(var_analisisctacte.haberdol) AS SumaDehaberdol, " _
            & " Sum([var_analisisctacte]![habersol]-[var_analisisctacte]![debesol]) AS saldosol, Sum([var_analisisctacte]![haberdol]-[var_analisisctacte]![debedol]) AS saldodol" _
            & " FROM var_analisisctacte LEFT JOIN mae_cliente ON var_analisisctacte.idcli = mae_cliente.id GROUP BY var_analisisctacte.idcli, mae_cliente.nombre, " _
            & " mae_cliente.numruc, var_analisisctacte.numdocref"

    End If
    
    If Fg2.Rows <> 0 Then
        ' MOSTRAMOS SOLOS LOS CLIENTES ESPECIFICADOS
        xCad = "SELECT var_analisisctacte.idcli, mae_cliente.nombre, mae_cliente.numruc, var_analisisctacte.numdocref, Sum(var_analisisctacte.debesol) AS SumaDedebesol, " _
            & " Sum(var_analisisctacte.habersol) AS SumaDehabersol, Sum(var_analisisctacte.debedol) AS SumaDedebedol, Sum(var_analisisctacte.haberdol) AS SumaDehaberdol, " _
            & " Sum([var_analisisctacte]![habersol]-[var_analisisctacte]![debesol]) AS saldosol, Sum([var_analisisctacte]![haberdol]-[var_analisisctacte]![debedol]) AS saldodol " _
            & " FROM var_analisisctacte LEFT JOIN mae_cliente ON var_analisisctacte.idcli = mae_cliente.id " _
            & xCadCondi _
            & " GROUP BY var_analisisctacte.idcli, mae_cliente.nombre, mae_cliente.numruc, var_analisisctacte.numdocref " _
            & " HAVING (" & xCadWhere & ")"
    End If
    
    RST_Busq RstDet, xCad, xCon
    RstDet.Filter = "idcli <>0"
    
    If RstDet.RecordCount <> 0 Then
        MostrarDetalle
    End If
    
    Exit Sub

LaCagada:
    MsgBox "Error Inesperado sucedio lo siguiente : " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Sub

Sub MostrarDetalle()
    Dim A As Double
    Dim xIdCli As Integer
    Dim xForeColor As Long
    Fg1.Rows = 2
    
    RstDet.MoveFirst
    xIdCli = RstDet("idcli")
    For A = 1 To RstDet.RecordCount
        Fg1.Rows = Fg1.Rows + 1
        'If RstDet("numdocref") = "118102008011655" Then
        '    MsgBox ""
        'End If
        
        If A = 1 Then
            GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 2, "CLIENTE : " & RstDet("nombre"), flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
            Fg1.Rows = Fg1.Rows + 1
        End If
        
        If RstDet("idcli") <> xIdCli Then
            Fg1.Rows = Fg1.Rows + 1
            GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 2, "CLIENTE : " & RstDet("nombre"), flexAlignLeftTop, True, , &H0&, &HE2FEFB, True
            xIdCli = RstDet("idcli")
            Fg1.Rows = Fg1.Rows + 1
        End If
        
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = "ORDEN DE DESPACHO"
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDet("numdocref"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(RstDet("sumadedebedol"), "0.00") 'Format((NulosN(RstDet("SumaDeimphabdol")) - NulosN(RstDet("SumaDeimpdebdol"))), "0.00")
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(RstDet("sumadehaberdol"), "0.00") 'Format((NulosN(RstDet("imphabsol2")) - NulosN(RstDet("impdebsol2"))), "0.00")
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(RstDet("sumadedebesol"), "0.00")
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(RstDet("sumadehabersol"), "0.00")
        
        If RstDet("saldodol") > 0 Then xForeColor = &HFF0000  'azul
        If RstDet("saldodol") < 0 Then xForeColor = &HFF&     'rojo
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, xForeColor, True, &HE2FEFB, Format(RstDet("saldodol"), "0.00")
        
        If RstDet("saldosol") > 0 Then xForeColor = &HFF0000  'azul
        If RstDet("saldosol") < 0 Then xForeColor = &HFF&     'rojo
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, xForeColor, True, &HE2FEFB, Format(RstDet("saldosol"), "0.00")
        
        Fg1.TextMatrix(Fg1.Rows - 1, 9) = RstDet("nombre")
        RstDet.MoveNext
        If RstDet.EOF = True Then Exit For
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
    'Dim xForeColor As Long
    
On Error GoTo LaCagada2

    Fg3.Rows = 2
    If RstDet.RecordCount <> 0 Then
        ProgressBar1.Max = RstDet.RecordCount
        Frame3.Left = ((Me.Width - Frame3.Width) / 2)
        Frame3.Top = ((Me.Height - Frame3.Height) / 2)
        Frame3.Visible = True
        Frame3.Refresh
        
        xNumOrd = NulosC(RstDet("numdocref"))
        xIdCliente = NulosN(RstDet("idcli"))
        xRucCli = NulosC(RstDet("numruc"))
        
        For B = 1 To RstDet.RecordCount
            ProgressBar1.Value = B
            
            If B = 1 Then
                Fg3.Rows = Fg3.Rows + 1
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, 2, "CLIENTE : " & RstDet("nombre"), flexAlignLeftCenter, True, , &HFF0000, &HE2FEFB, True
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 9, "Nº R.U.C. : " & xRucCli, flexAlignLeftCenter, True, , &HFF0000, &HE2FEFB, True
                
                Fg3.Rows = Fg3.Rows + 1
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, 2, "DOC. REF. :  ORDEN DE DESPACHO  ", flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 9, "Nº DOC. REF. : " & NulosC(RstDet("numdocref")), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                
                Fg3.Rows = Fg3.Rows + 1
            Else
                If (NulosC(RstDet("numdocref")) <> xNumOrd) Or (NulosC(RstDet("numdocref")) = xNumOrd And NulosC(RstDet("idcli")) <> xIdCliente) Then
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
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 10, "Nº DOC. REF. : " & NulosC(RstDet("numdocref")), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                    
                    xNumOrd = NulosC(RstDet("numdocref"))
                Else
                
                End If
                
                If NulosC(RstDet("idcli")) <> xIdCliente Then
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
                    
                    xIdCliente = NulosN(RstDet("idcli"))
                    xRucCli = Busca_Codigo(xIdCliente, "id", "numruc", "mae_cliente", "N", xCon)
                    
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, 2, "CLIENTE : " & RstDet("nombre"), flexAlignLeftCenter, True, , &HFF0000, &HE2FEFB, True
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 9, "Nº R.U.C. : " & xRucCli, flexAlignLeftCenter, True, , &HFF0000, &HE2FEFB, True
                    xIdCliente = RstDet("idcli")
                    Fg3.Rows = Fg3.Rows + 1
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, 2, "DOC. REF. :  ORDEN DE DESPACHO  ", flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                    GRID_COMBINAR Fg3, Fg3.Rows - 1, 7, Fg3.Rows - 1, 10, "Nº DOC. REF. : " & NulosC(RstDet("numdocref")), flexAlignLeftCenter, True, , &H800000, &HE2FEFB, True
                    
                End If
                
                Fg3.Rows = Fg3.Rows + 1
            End If
            
            Fg3.TextMatrix(Fg3.Rows - 1, 1) = NulosC(RstDet("cuenta"))
            Fg3.TextMatrix(Fg3.Rows - 1, 2) = NulosC(RstDet("descripcion"))
            Fg3.TextMatrix(Fg3.Rows - 1, 3) = NulosC(RstDet("numreg"))
            Fg3.TextMatrix(Fg3.Rows - 1, 4) = NulosC(RstDet("numdocref"))
            
            Fg3.TextMatrix(Fg3.Rows - 1, 5) = NulosC(RstDet("fchemi"))
            Fg3.TextMatrix(Fg3.Rows - 1, 6) = NulosC(RstDet("abrev"))
            Fg3.TextMatrix(Fg3.Rows - 1, 7) = NulosC(RstDet("numdoc"))
            Fg3.TextMatrix(Fg3.Rows - 1, 15) = NulosC(RstDet("glosa"))
            Fg3.TextMatrix(Fg3.Rows - 1, 16) = NulosC(RstDet("nombre"))
            Fg3.TextMatrix(Fg3.Rows - 1, 8) = NulosC(RstDet("simbolo"))
            Fg3.TextMatrix(Fg3.Rows - 1, 9) = Format(NulosN(RstDet("imptc")), "0.000")
            Fg3.TextMatrix(Fg3.Rows - 1, 10) = "0.00"
                        
            ' INVERTIMOS LA PRESENTACION DE LOS IMPORTES PORQUE EN LA CUENTA DE SAVAR SE MUESTRA AL REVES LOS MOVIMIENTOS
            Fg3.TextMatrix(Fg3.Rows - 1, 11) = Format(NulosN(RstDet("debedol")), "0.00")
            Fg3.TextMatrix(Fg3.Rows - 1, 12) = Format(NulosN(RstDet("haberdol")), "0.00")
            
            Fg3.TextMatrix(Fg3.Rows - 1, 13) = Format(NulosN(RstDet("debesol")), "0.00")
            Fg3.TextMatrix(Fg3.Rows - 1, 14) = Format(NulosN(RstDet("habersol")), "0.00")

            xTotDebDol = xTotDebDol + NulosN(RstDet("debedol"))
            xTotHabDol = xTotHabDol + NulosN(RstDet("haberdol"))
            xTotDebSol = xTotDebSol + NulosN(RstDet("debesol"))
            xTotHabSol = xTotHabSol + NulosN(RstDet("habersol"))

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
                
                xxTotalDol = xxTotalDol + (xTotHabDol - xTotDebDol)
                xxTotalSol = xxTotalSol + (xTotHabSol - xTotDebSol)
                
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
'    Resume
    MsgBox "En el Procedimiento ImprimirDetalleFG2 se produjo el siguiente error : " & Err.Description
    Resume Next
End Sub



