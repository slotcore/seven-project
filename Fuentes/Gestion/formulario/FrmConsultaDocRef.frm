VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "Codejock.CommandBars.v12.0.0.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.ocx"
Begin VB.Form FrmConsultaDocRef 
   Caption         =   "Gestion - Analisis x Documento de Referencia"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   14010
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   885
      Left            =   3090
      TabIndex        =   16
      Top             =   4830
      Visible         =   0   'False
      Width           =   6180
      Begin XtremeSuiteControls.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   90
         TabIndex        =   21
         Top             =   435
         Width           =   5985
         _Version        =   786432
         _ExtentX        =   10557
         _ExtentY        =   556
         _StockProps     =   93
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
         TabIndex        =   18
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
      Height          =   7275
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   465
      Width           =   13695
      _cx             =   24156
      _cy             =   12832
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
      BorderWidth     =   2
      ChildSpacing    =   2
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
      _GridInfo       =   $"FrmConsultaDocRef.frx":0000
      Begin VSFlex7Ctl.VSFlexGrid Fg1 
         Height          =   5550
         Left            =   30
         TabIndex        =   15
         Top             =   1695
         Width           =   13635
         _cx             =   24051
         _cy             =   9790
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsultaDocRef.frx":0045
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
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   1635
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   13635
         _cx             =   24051
         _cy             =   2884
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
         BorderWidth     =   2
         ChildSpacing    =   2
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
         _GridInfo       =   $"FrmConsultaDocRef.frx":011A
         Begin VB.Frame Frame4 
            Caption         =   "[ Tipo Doc. Referencia ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   9990
            TabIndex        =   13
            Top             =   30
            Width           =   3615
            Begin VSFlex7Ctl.VSFlexGrid Fg4 
               Height          =   1230
               Left            =   60
               TabIndex        =   14
               Top             =   270
               Width           =   2610
               _cx             =   4604
               _cy             =   2170
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
               Rows            =   1
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmConsultaDocRef.frx":0178
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
         Begin VB.Frame Frame3 
            Caption         =   "[ Origenes ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   7365
            TabIndex        =   4
            Top             =   30
            Width           =   2595
            Begin VSFlex7Ctl.VSFlexGrid Fg3 
               Height          =   1230
               Left            =   60
               TabIndex        =   12
               Top             =   270
               Width           =   2475
               _cx             =   4366
               _cy             =   2170
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
               Rows            =   5
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmConsultaDocRef.frx":01DD
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
         Begin VB.Frame Frame2 
            Caption         =   "[ Cliente ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   2685
            TabIndex        =   3
            Top             =   30
            Width           =   4650
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   1230
               Left            =   60
               TabIndex        =   11
               Top             =   270
               Width           =   4530
               _cx             =   7990
               _cy             =   2170
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
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmConsultaDocRef.frx":0272
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
         Begin VB.Frame Frame1 
            Caption         =   "[ Opciones ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   30
            TabIndex        =   2
            Top             =   30
            Width           =   2625
            Begin VB.Frame Frame6 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  'None
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
               Height          =   570
               Left            =   1200
               TabIndex        =   19
               Top             =   945
               Width           =   1380
               Begin VB.OptionButton Option4 
                  Caption         =   "Sin I.G.V."
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
                  Height          =   195
                  Left            =   135
                  TabIndex        =   17
                  Top             =   315
                  Width           =   1230
               End
               Begin VB.OptionButton Option3 
                  Caption         =   "Con I.G.V."
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
                  Height          =   195
                  Left            =   135
                  TabIndex        =   20
                  Top             =   45
                  Width           =   1230
               End
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Detallado"
               Height          =   195
               Left            =   135
               TabIndex        =   10
               Top             =   1260
               Width           =   1320
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Resumido"
               Height          =   195
               Left            =   135
               TabIndex        =   9
               Top             =   990
               Width           =   1320
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   1155
               TabIndex        =   5
               Top             =   285
               Width           =   1200
               _ExtentX        =   2117
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
               Left            =   1155
               TabIndex        =   6
               Top             =   600
               Width           =   1200
               _ExtentX        =   2117
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
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Final"
               Height          =   195
               Left            =   135
               TabIndex        =   8
               Top             =   645
               Width           =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Inicio"
               Height          =   195
               Left            =   135
               TabIndex        =   7
               Top             =   315
               Width           =   735
            End
         End
      End
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   2355
      Top             =   120
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmConsultaDocRef.frx":02C2
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   1935
      Top             =   90
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu_01 
         Caption         =   "Agregar Cliente"
      End
      Begin VB.Menu menu_02 
         Caption         =   "-"
      End
      Begin VB.Menu menu_03 
         Caption         =   "Eliminar Cliente"
      End
   End
End
Attribute VB_Name = "FrmConsultaDocRef"
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
    
Dim Rst As New ADODB.Recordset
Dim Rst2 As New ADODB.Recordset
Dim Rst3 As New ADODB.Recordset
Dim Rst4 As New ADODB.Recordset
Dim Rst5 As New ADODB.Recordset

Dim xTotComSol, xTotComDol, xTotVenSol, xTotVenDol As Double              ' variables para totalizar por documento de referencia
Dim xTotLGDSol, xTotLGDDol, xTotReeSol, xTotReeDol As Double

Dim xTotComSolTot, xTotComDolTot, xTotVenSolTot, xTotVenDolTot As Double  ' variables para totalizar por cliente
Dim xTotLGDSolTot, xTotLGDDolTot, xTotReeSolTot, xTotReeDolTot As Double

Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Id = 1 Then Procesar
    
    If Control.Id = 2 Then pExportar
    If Control.Id = 5 Then
        Unload Me
    End If
End Sub

Private Sub pExportar()
    Dim xFun As New SGI2_funciones.Formularios
    Dim Rst As New ADODB.Recordset
    
    If Fg1.Rows = 1 Then
        MsgBox "No hay registro para exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    GRID_EXPORTAR_MSEXCELTMP Fg1, xCon, flexFileCustomText, True, "Analisis x Documento de Referencia - DETALLADO", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en : Ambas Monedas"
    
    Set xFun = Nothing
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

Private Sub Fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu menu
    End If
End Sub

Private Sub Fg3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        
    End If
End Sub

Private Sub Fg3_EnterCell()
    If Fg3.Col = 2 Then
        Fg3.Editable = flexEDKbdMouse
    Else
        Fg3.Editable = flexEDNone
    End If
End Sub

Private Sub Fg4_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        If NulosN(Fg4.TextMatrix(Row, Col)) = -1 Then
            Fg3.Rows = 0
            Fg3.Rows = 5
            Fg3.TextMatrix(0, 1) = "VENTAS"
            Fg3.TextMatrix(1, 1) = "COMPRAS"
            Fg3.TextMatrix(2, 1) = "HONORARIOS"
            Fg3.TextMatrix(3, 1) = "REEMBOLSABLES"
            Fg3.TextMatrix(4, 1) = "L.G.D. / L.G.C."
        
            Fg3.TextMatrix(0, 2) = -1
            Fg3.TextMatrix(1, 2) = -1
            Fg3.TextMatrix(2, 2) = -1
            Fg3.TextMatrix(3, 2) = -1
            Fg3.TextMatrix(4, 2) = -1
        Else
            Fg3.Rows = 1
        End If
        Option2.Value = True
        Option2_Click
    End If
End Sub

Private Sub Fg4_EnterCell()
    If Fg4.Col = 2 Then
        Fg4.Editable = flexEDKbdMouse
    Else
        Fg4.Editable = flexEDNone
    End If
End Sub

Private Sub Form_Load()
    Me.WindowState = 2
    SeEjecuto = False
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
    Frame6.BackColor = &H8000000F
    SetearCuadricula Fg1, 7, xCon, 2, 2, False
    Fg1.BackColor = &HE2FEFB
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.BackColorSel = &H80&
    Fg1.Editable = flexEDNone
        
    Fg2.BackColor = &HE2FEFB
    Fg2.Rows = 0
    Fg2.Rows = Fg2.Rows + 1
    Fg2.ColComboList(1) = "|..."
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.ColWidth(2) = 0
    Fg2.Editable = flexEDKbdMouse
    
    Dim A As Integer
    For A = 0 To 4
        Fg3.TextMatrix(A, 2) = -1
    Next A
        
    Fg3.Editable = flexEDNone
    Fg3.SelectionMode = flexSelectionByRow
    
    Fg4.TextMatrix(0, 2) = -1
    Fg4.Editable = flexEDNone
    Fg4.SelectionMode = flexSelectionByRow
    
    Option2.Value = True   ' ponemos la opcion detallado por defecto
    Option3.Value = True   ' ponemos la opcion con IGV por defecto
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    CrearTool
    
    If TAMAÑO_TOOL = I16x16 Then EO.Top = 400: TopEO = 400
    If TAMAÑO_TOOL = I24x24 Then EO.Top = 520: TopEO = 520
    If TAMAÑO_TOOL = I32x32 Then EO.Top = 640: TopEO = 640
    If TAMAÑO_TOOL = I48x48 Then EO.Top = 890: TopEO = 890

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

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    EO.Width = Me.Width - 130
    If Me.Height <= (TopEO + 2375) Then
        Me.Height = (TopEO + 2375)
    Else
        EO.Height = (Me.Height - (TopEO + 400))
    End If
    
    Me.Refresh
End Sub

Sub Procesar()
    Dim A As Integer
    Dim xCadWhere As String
    
    If TxtFchIni.Valor = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If TxtFchFin.Valor = "" Then
        MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "Rango de fechas ingresado no valido", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    Dim xNumSel As Integer
    
    For A = 0 To Fg3.Rows - 1
        If NulosN(Fg3.TextMatrix(A, 2)) = -1 Then
            xNumSel = xNumSel + 1
        End If
    Next A
    
    If xNumSel = 0 Then
        MsgBox "Debe de seleccionar al menos un origen para realizar esta consulta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    ' ELIMINAMOS LAS FILAS EN BLANCO
    If Fg2.Rows <> 0 Then
        For A = 0 To Fg2.Rows - 1
            If NulosN(Fg2.TextMatrix(A, 2)) = 0 Then
                Fg2.RemoveItem A
            End If
        Next A
    End If

    If Fg2.Rows <> 0 Then
        xCadWhere = ""
        ' MOSTRAMOS SOLOS LOS CLIENTES ESPECIFICADOS

        xCadWhere = "("
        For A = 0 To Fg2.Rows - 1
            xCadWhere = xCadWhere & "(var_ordendespacho.idcli = " & Fg2.TextMatrix(A, 2) & ")"
            If A = Fg2.Rows - 1 Then
                Exit For
            End If
            xCadWhere = xCadWhere & " OR "
        Next A
        xCadWhere = xCadWhere & ") AND"
    Else
        xCadWhere = ""
    End If
    
    Dim xCadWhere2 As String
    
    If Fg2.Rows <> 0 Then
        xCadWhere2 = ""
        ' MOSTRAMOS SOLOS LOS CLIENTES ESPECIFICADOS

        xCadWhere2 = "("
        For A = 0 To Fg2.Rows - 1
            xCadWhere2 = xCadWhere2 & "(vta_ventas.idcli = " & Fg2.TextMatrix(A, 2) & ")"
            If A = Fg2.Rows - 1 Then
                Exit For
            End If
            xCadWhere2 = xCadWhere2 & " OR "
        Next A
        xCadWhere2 = xCadWhere2 & ") AND"
    Else
        xCadWhere2 = ""
    End If
    
    ' ARMAMOS LA CONSULTA PARA OBTENER LOS DOCUMENTOS DE REFERENCIA REQUERIDOS - ESTA CONSULTA CARGARA UN CURSOR CON LOS DATOS SOLICITADOS
    Dim xCursor As String
    xCursor = "SELECT DISTINCT var_ordendespacho.id AS idorddes, var_ordendespacho.idcli, mae_cliente.numruc, mae_cliente.nombre, var_ordendespacho.id AS iddocref, " _
        & " var_ordendespacho.idtipdoc AS tipdocref, mae_docreferencia.descripcion, var_ordendespacho.numerodoc, var_ordendespacho.fchemi " _
        & " FROM (var_ordendespacho LEFT JOIN mae_cliente ON var_ordendespacho.idcli = mae_cliente.id) LEFT JOIN mae_docreferencia ON var_ordendespacho.idtipdoc = mae_docreferencia.id " _
        & " WHERE ( " & xCadWhere _
        & " ((var_ordendespacho.fchemi)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (var_ordendespacho.fchemi)<=CDate('" & TxtFchFin.Valor & "'))) ORDER BY mae_cliente.nombre"
    
    If NulosN(Fg4.TextMatrix(0, 2)) = -1 Then
        If Option1.Value = True Then VerResumen xCursor
        If Option2.Value = True Then VerDetalle xCursor
    Else
        If Option1.Value = True Then VerSinDocRefResumen xCadWhere2
        'If Option2.Value = True Then VerSinDocRefDetalle
    End If
End Sub

Sub VerSinDocRefResumen(xCadWhere As String)
    Dim A As Integer
    
    Dim xSQL As String
    Dim xCursor As String
    
    ' creamos el cursor con los documentos sin documento de referencia
    'xCadWhere
    If Option4 = True Then
        ' sin I.G.V.
        xCursor = "SELECT vta_ventas.idcli, mae_cliente.numruc, mae_cliente.nombre, " _
            & " Sum(IIf(vta_ventas!idmon=1,vta_ventas!impbru+vta_ventas!impbru2+vta_ventas!impbru3+vta_ventas!impinaf, " _
            & " IIf(vta_ventas!tc<>0,(vta_ventas!impbru+vta_ventas!impbru2+vta_ventas!impbru3+vta_ventas!impinaf)*vta_ventas!tc, " _
            & " (vta_ventas!impbru+vta_ventas!impbru2+vta_ventas!impbru3+vta_ventas!impinaf)*con_tc!impven))) AS impsol, " _
            & " Sum(IIf(vta_ventas!idmon=2,vta_ventas!impbru+vta_ventas!impbru2+vta_ventas!impbru3+vta_ventas!impinaf, " _
            & " IIf(vta_ventas!tc<>0,(vta_ventas!impbru+vta_ventas!impbru2+vta_ventas!impbru3+vta_ventas!impinaf)/vta_ventas!tc, " _
            & " (vta_ventas!impbru+vta_ventas!impbru2+vta_ventas!impbru3+vta_ventas!impinaf)/con_tc!impven))) AS impdol" _
            & " FROM (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id=vta_ventas.idcli) LEFT JOIN con_tc ON vta_ventas.fchdoc=con_tc.fecha " _
            & " WHERE (" & xCadWhere _
            & " ((vta_ventas.iddocref)=0 Or (vta_ventas.iddocref) Is Null) AND ((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_ventas.anulado)=0) AND ((vta_ventas.numreg)<>'000001') AND ((vta_ventas.tipdoc)<>7)) " _
            & " GROUP BY vta_ventas.idcli, mae_cliente.numruc, mae_cliente.nombre " _
            & " UNION " _
            & " SELECT vta_ventas.idcli, mae_cliente.numruc, mae_cliente.nombre, " _
            & " Sum(IIf(vta_ventas!idmon=1,0-(vta_ventas!impbru+vta_ventas!impbru2+vta_ventas!impbru3+vta_ventas!impinaf), " _
            & " IIf(vta_ventas!tc<>0,0-((vta_ventas!impbru+vta_ventas!impbru2+vta_ventas!impbru3+vta_ventas!impinaf)*vta_ventas!tc)," _
            & " 0-((vta_ventas!impbru+vta_ventas!impbru2+vta_ventas!impbru3+vta_ventas!impinaf)*con_tc!impven)))) AS impsol, " _
            & " Sum(IIf(vta_ventas!idmon=2,0-(vta_ventas!impbru+vta_ventas!impbru2+vta_ventas!impbru3+vta_ventas!impinaf), " _
            & " IIf(vta_ventas!tc<>0,0-((vta_ventas!impbru+vta_ventas!impbru2+vta_ventas!impbru3+vta_ventas!impinaf)/vta_ventas!tc), " _
            & " 0-((vta_ventas!impbru+vta_ventas!impbru2+vta_ventas!impbru3+vta_ventas!impinaf)/con_tc!impven)))) AS impdol " _
            & " FROM (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id=vta_ventas.idcli) LEFT JOIN con_tc ON vta_ventas.fchdoc=con_tc.fecha " _
            & " WHERE (" & xCadWhere _
            & " ((vta_ventas.iddocref)=0 Or (vta_ventas.iddocref) Is Null) AND ((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_ventas.anulado)=0) AND ((vta_ventas.numreg)<>'000001') AND ((vta_ventas.tipdoc)=7))" _
            & " GROUP BY vta_ventas.idcli, mae_cliente.numruc, mae_cliente.nombre"
    Else
        'CON I.G.V.
        xCursor = "SELECT vta_ventas.idcli, mae_cliente.numruc, mae_cliente.nombre, " _
            & " Sum(IIf(vta_ventas!idmon=1,vta_ventas!imptotdoc, " _
            & " IIf(vta_ventas!tc<>0,(vta_ventas!imptotdoc)*vta_ventas!tc, " _
            & " (vta_ventas!imptotdoc)*con_tc!impven))) AS impsol, " _
            & " Sum(IIf(vta_ventas!idmon=2,vta_ventas!imptotdoc, " _
            & " IIf(vta_ventas!tc<>0,(vta_ventas!imptotdoc)/vta_ventas!tc, " _
            & " (vta_ventas!imptotdoc)/con_tc!impven))) AS impdol" _
            & " FROM (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id=vta_ventas.idcli) LEFT JOIN con_tc ON vta_ventas.fchdoc=con_tc.fecha " _
            & " WHERE (" & xCadWhere _
            & " ((vta_ventas.iddocref)=0 Or (vta_ventas.iddocref) Is Null) AND ((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_ventas.anulado)=0) AND ((vta_ventas.numreg)<>'000001') AND ((vta_ventas.tipdoc)<>7)) " _
            & " GROUP BY vta_ventas.idcli, mae_cliente.numruc, mae_cliente.nombre " _
            & " UNION " _
            & " SELECT vta_ventas.idcli, mae_cliente.numruc, mae_cliente.nombre, " _
            & " Sum(IIf(vta_ventas!idmon=1,0-(vta_ventas!imptotdoc), " _
            & " IIf(vta_ventas!tc<>0,0-((vta_ventas!imptotdoc)*vta_ventas!tc)," _
            & " 0-((vta_ventas!imptotdoc)*con_tc!impven)))) AS impsol, " _
            & " Sum(IIf(vta_ventas!idmon=2,0-(vta_ventas!imptotdoc), " _
            & " IIf(vta_ventas!tc<>0,0-((vta_ventas!imptotdoc)/vta_ventas!tc), " _
            & " 0-((vta_ventas!imptotdoc)/con_tc!impven)))) AS impdol " _
            & " FROM (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id=vta_ventas.idcli) LEFT JOIN con_tc ON vta_ventas.fchdoc=con_tc.fecha " _
            & " WHERE (" & xCadWhere _
            & " ((vta_ventas.iddocref)=0 Or (vta_ventas.iddocref) Is Null) AND ((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_ventas.anulado)=0) AND ((vta_ventas.numreg)<>'000001') AND ((vta_ventas.tipdoc)=7))" _
            & " GROUP BY vta_ventas.idcli, mae_cliente.numruc, mae_cliente.nombre"
    End If
    
    xSQL = "SELECT [union-res].idcli, [union-res].numruc, [union-res].nombre, Sum([union-res].impsol) AS SumaDeimpsol, Sum([union-res].impdol) AS SumaDeimpdol" _
        & " From " _
        & " ( " & xCursor _
        & " ) AS [union-res] " _
        & " GROUP BY [union-res].idcli, [union-res].numruc, [union-res].nombre"
   
    RST_Busq Rst, xSQL, xCon
    Rst.Sort = "nombre"
    Dim xTotSol, xTotDol As Double
    Fg1.Rows = 2
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("numruc")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("nombre")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(Rst("sumadeimpsol"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(Rst("sumadeimpdol"), "0.00")
            
            xTotSol = xTotSol + Rst("sumadeimpsol")
            xTotDol = xTotDol + Rst("sumadeimpdol")
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Fg1.Rows = Fg1.Rows + 1
                GRID_COMBINAR Fg1, Fg1.Rows - 1, 2, Fg1.Rows - 1, 3, "TOTAL ==> ", flexAlignLeftCenter, True, 1, &H800080, &HE2FEFB, True
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &H80000012, True, &HE2FEFB, Format(xTotSol, "0.00")
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H80000012, True, &HE2FEFB, Format(xTotDol, "0.00")
                
                Exit For
            End If
        Next A
    End If
    
End Sub

Sub VerResumen(xCursor As String)
    Dim A As Integer
    Dim xSQL  As String
    
    ' GENERAMOS LA TABLA VIRTUAL XSQL HACIENDO UNA UNION DE VARIAS CONSULTAS
       
    ' CON I.G.V.
    xSQL = ""
    If Option3.Value = True Then
        If NulosN(Fg3.TextMatrix(0, 2)) = -1 Then
            xSQL = "SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, vta_ventas.idtipdocref, [consulta OD].numerodoc AS numdocref, 0 AS comsol, " _
                & " 0 AS comdol, IIf(vta_ventas.idmon=1,vta_ventas.imptotdoc,IIf(vta_ventas!tc<>0,(vta_ventas.imptotdoc*vta_ventas!tc),vta_ventas.imptotdoc*con_tc!impven)) AS vensol, " _
                & " IIf(vta_ventas.idmon=2,vta_ventas.imptotdoc,IIf(vta_ventas!tc<>0,(vta_ventas.imptotdoc/vta_ventas!tc),vta_ventas.imptotdoc/con_tc!impven)) AS vendol, " _
                & " 0 AS lgdsol, 0 AS lgddol, 0 AS reesol, 0 AS reedol, 'V' AS tipo, vta_ventas.id" _
                & " FROM " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " LEFT JOIN (vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON [consulta OD].idorddes = vta_ventas.iddocref2 " _
                & " WHERE (((vta_ventas.idtipdocref)=4) AND ((vta_ventas.anulado)=0) AND ((vta_ventas.tipdoc)<>7))" _
                & " UNION " _
                & " SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, vta_ventas.idtipdocref, [consulta OD].numerodoc AS numdocref, 0 AS comsol, " _
                & " 0 AS comdol, IIf(vta_ventas.idmon=1,0-vta_ventas.imptotdoc,IIf(vta_ventas!tc<>0,0-(vta_ventas.imptotdoc*vta_ventas!tc),0-(vta_ventas.imptotdoc*con_tc!impven))) AS vensol, " _
                & " IIf(vta_ventas.idmon=2,0-vta_ventas.imptotdoc,IIf(vta_ventas!tc<>0,0-(vta_ventas.imptotdoc/vta_ventas!tc),0-(vta_ventas.imptotdoc/con_tc!impven))) AS vendol, " _
                & " 0 AS lgdsol, 0 AS lgddol, 0 AS reesol, 0 AS reedol, 'V' AS tipo, vta_ventas.id " _
                & " FROM " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " LEFT JOIN (vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON [consulta OD].idorddes = vta_ventas.iddocref2 " _
                & " WHERE (((vta_ventas.idtipdocref)=4) AND ((vta_ventas.anulado)=0) AND ((vta_ventas.tipdoc)=7))"
        End If
        
        If NulosN(Fg3.TextMatrix(1, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, com_compras.idtipdocref, [consulta OD].numerodoc AS numdocref, " _
                & " IIf(com_compras.idmon=1,com_compras.imptot,IIf(com_compras!tc<>0,(com_compras.imptot*com_compras!tc),com_compras.imptot*con_tc!impven)) AS comsol, " _
                & " IIf(com_compras.idmon=2,com_compras.imptot,IIf(com_compras!tc<>0,(com_compras.imptot/com_compras!tc),com_compras.imptot/con_tc!impven)) AS comdol, " _
                & " 0 AS vensol, 0 AS vendol, 0 AS lgdsol, 0 AS lgddol, 0 AS reesol, 0 AS reedol, 'C' AS tipo, com_compras.id " _
                & " FROM " _
                & " ( " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " LEFT JOIN com_compras ON [consulta OD].idorddes=com_compras.iddocref2) LEFT JOIN con_tc ON com_compras.fchdoc=con_tc.fecha " _
                & " WHERE (((com_compras.idtipdocref)=4) AND ((com_compras.tipdoc)<>7))" _
                & " UNION " _
                & " SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, com_compras.idtipdocref, [consulta OD].numerodoc AS numdocref, " _
                & " IIf([com_compras].[idmon]=1,0-[com_compras].[imptot],IIf([com_compras]![tc]<>0,0-([com_compras].[imptot]*[com_compras]![tc]),0-([com_compras].[imptot]*[con_tc]![impven]))) AS comsol, " _
                & " IIf([com_compras].[idmon]=2,0-[com_compras].[imptot],IIf([com_compras]![tc]<>0,0-([com_compras].[imptot]/[com_compras]![tc]),0-([com_compras].[imptot]/[con_tc]![impven]))) AS comdol, " _
                & " 0 AS vensol, 0 AS vendol, 0 AS lgdsol, 0 AS lgddol, 0 AS reesol, 0 AS reedol, 'C' AS tipo, com_compras.id " _
                & " FROM " _
                & " ( " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " LEFT JOIN com_compras ON [consulta OD].idorddes = com_compras.iddocref2) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
                & " WHERE (((com_compras.idtipdocref)=4) AND ((com_compras.tipdoc)=7))"
        End If
        
        If NulosN(Fg3.TextMatrix(2, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & "SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, com_honorarios.idtipdocref, [consulta OD].numerodoc, " _
                & " IIf(com_honorarios.idmon=1,com_honorarios.imptot,IIf(com_honorarios!tc<>0,(com_honorarios.imptot*com_honorarios!tc),com_honorarios.imptot*con_tc!impven)) AS comsol, " _
                & " IIf(com_honorarios.idmon=2,com_honorarios.imptot,IIf(com_honorarios!tc<>0,(com_honorarios.imptot/com_honorarios!tc),com_honorarios.imptot/con_tc!impven)) AS comdol, " _
                & " 0 AS vensol, 0 AS vendol, 0 AS ldgsol, 0 AS lgddol, 0 AS reesol, 0 AS reedol, 'H' AS tipo, com_honorarios.id " _
                & " FROM (com_honorarios LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) " _
                & " RIGHT JOIN " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " ON com_honorarios.iddocref2 = [consulta OD].idorddes WHERE (((com_honorarios.idtipdocref)=4)) "
        End If
        
        If NulosN(Fg3.TextMatrix(4, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, vta_gastodebito.idtipdocref, [consulta OD].numerodoc AS numdocref, " _
                & " 0 AS consol, 0 AS condol, 0 AS vensol, 0 AS vendol, IIf([vta_gastodebito].[idmon]=1,[vta_gastodebito].[imptot], " _
                & " IIf([vta_gastodebito]![tc]<>0,([vta_gastodebito].[imptot]*[vta_gastodebito]![tc]),[vta_gastodebito].[imptot]*[con_tc]![impven])) AS lgdsol, " _
                & " IIf([vta_gastodebito].[idmon]=2,[vta_gastodebito].[imptot],IIf([vta_gastodebito]![tc]<>0,([vta_gastodebito].[imptot]/[vta_gastodebito]![tc]),[vta_gastodebito].[imptot]/[con_tc]![impven])) AS lgddol, " _
                & " 0 AS reesol, 0 AS reedol," _
                & " 'L' AS tipo, vta_gastodebito.id " _
                & " FROM " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " LEFT JOIN (vta_gastodebito LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha) ON [consulta OD].idorddes = vta_gastodebito.iddocref2 " _
                & " WHERE (((vta_gastodebito.idtipdocref)=4))"
        End If
        
        If NulosN(Fg3.TextMatrix(3, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, com_reembolsables.idtipdocref, [consulta OD].numerodoc AS numdocref, " _
                & " 0 AS comsol, 0 AS comdol, 0 AS vensol, 0 AS vendol, 0 AS lgdsol, 0 AS lgddol, " _
                & " IIf([com_reembolsables].[idmon]=1,[com_reembolsables].[imptot], IIf([com_reembolsables]![tc]<>0,([com_reembolsables].[imptot]*[com_reembolsables]![tc])," _
                & " [com_reembolsables].[imptot]*[con_tc]![impven])) AS reesol, IIf([com_reembolsables].[idmon]=2,[com_reembolsables].[imptot],IIf([com_reembolsables]![tc]<>0, " _
                & " ([com_reembolsables].[imptot]/[com_reembolsables]![tc]), [com_reembolsables].[imptot]/[con_tc]![impven])) AS reedol, 'R' AS tipo, com_reembolsables.id " _
                & " FROM (com_reembolsables RIGHT JOIN " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " ON com_reembolsables.iddocref2 = [consulta OD].idorddes) LEFT JOIN con_tc ON com_reembolsables.fchdoc = con_tc.fecha " _
                & " WHERE (((com_reembolsables.idtipdocref)=4))"
        End If
    End If
   
   
    ' SIN I.G.V.
    If Option4.Value = True Then
        If NulosN(Fg3.TextMatrix(0, 2)) = -1 Then
            xSQL = "SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, vta_ventas.idtipdocref, [consulta OD].numerodoc AS numdocref, 0 AS comsol, 0 AS comdol, " _
                & " IIf(vta_ventas.idmon=1,(vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf),IIf(vta_ventas!tc<>0,((vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf)*vta_ventas!tc),(vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf)*con_tc!impven)) AS vensol, " _
                & " IIf(vta_ventas.idmon=2,(vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf),IIf(vta_ventas!tc<>0,((vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf)/vta_ventas!tc),(vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf)/con_tc!impven)) AS vendol, " _
                & " 0 AS lgdsol, 0 AS lgddol, 0 AS reesol, 0 AS reedol, 'V' AS tipo, vta_ventas.id " _
                & " FROM " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " LEFT JOIN (vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON [consulta OD].idorddes = vta_ventas.iddocref2 " _
                & " WHERE (((vta_ventas.idtipdocref)=4) AND ((vta_ventas.anulado)=0) AND ((vta_ventas.tipdoc)<>7))" _
                & " UNION " _
                & " SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, vta_ventas.idtipdocref, [consulta OD].numerodoc AS numdocref, 0 AS comsol, 0 AS comdol, " _
                & " IIf(vta_ventas.idmon=1,0-(vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf),IIf(vta_ventas!tc<>0,0-((vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf)*vta_ventas!tc),0-((vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf)*con_tc!impven))) AS vensol, " _
                & " IIf(vta_ventas.idmon=2,0-(vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf),IIf(vta_ventas!tc<>0,0-((vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf)/vta_ventas!tc),0-((vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf)/con_tc!impven))) AS vendol, " _
                & " 0 AS lgdsol, 0 AS lgddol, 0 AS reesol, 0 AS reedol, 'V' AS tipo, vta_ventas.id " _
                & " FROM " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " LEFT JOIN (vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON [consulta OD].idorddes = vta_ventas.iddocref2 " _
                & " WHERE (((vta_ventas.idtipdocref)=4) AND ((vta_ventas.anulado)=0) AND ((vta_ventas.tipdoc)=7))"
        End If
        
        If NulosN(Fg3.TextMatrix(1, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, com_compras.idtipdocref, [consulta OD].numerodoc AS numdocref, " _
                & " IIf(com_compras.idmon=1,(com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina),IIf(com_compras!tc<>0,((com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina)*com_compras!tc),(com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina)*con_tc!impven)) AS comsol, " _
                & " IIf(com_compras.idmon=2,(com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina),IIf(com_compras!tc<>0,((com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina)/com_compras!tc),(com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina)/con_tc!impven)) AS comdol, " _
                & " 0 AS vensol, 0 AS vendol, 0 AS lgdsol, 0 AS lgddol, 0 AS reesol, 0 AS reedol, 'C' AS tipo, com_compras.id " _
                & " FROM " _
                & " ( " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " LEFT JOIN com_compras ON [consulta OD].idorddes=com_compras.iddocref2) LEFT JOIN con_tc ON com_compras.fchdoc=con_tc.fecha " _
                & " WHERE (((com_compras.idtipdocref)=4) AND ((com_compras.tipdoc)<>7))" _
                & " UNION " _
                & " SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, com_compras.idtipdocref, [consulta OD].numerodoc AS numdocref, " _
                & " IIf([com_compras].[idmon]=1,0-(com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina),IIf([com_compras]![tc]<>0,0-((com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina)*[com_compras]![tc]),0-((com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina)*[con_tc]![impven]))) AS comsol, " _
                & " IIf([com_compras].[idmon]=2,0-(com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina),IIf([com_compras]![tc]<>0,0-((com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina)/[com_compras]![tc]),0-((com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina)/[con_tc]![impven]))) AS comdol, " _
                & " 0 AS vensol, 0 AS vendol, 0 AS lgdsol, 0 AS lgddol, 0 AS reesol, 0 AS reedol, 'C' AS tipo, com_compras.id " _
                & " FROM " _
                & " ( " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " LEFT JOIN com_compras ON [consulta OD].idorddes = com_compras.iddocref2) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
                & " WHERE (((com_compras.idtipdocref)=4) AND ((com_compras.tipdoc)=7))"
        End If
        
        If NulosN(Fg3.TextMatrix(2, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & "SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, com_honorarios.idtipdocref, [consulta OD].numerodoc, " _
                & " IIf(com_honorarios.idmon=1,(com_honorarios.impbru+com_honorarios.impina),IIf(com_honorarios!tc<>0,((com_honorarios.impbru+com_honorarios.impina)*com_honorarios!tc),(com_honorarios.impbru+com_honorarios.impina)*con_tc!impven)) AS comsol, " _
                & " IIf(com_honorarios.idmon=2,(com_honorarios.impbru+com_honorarios.impina),IIf(com_honorarios!tc<>0,((com_honorarios.impbru+com_honorarios.impina)/com_honorarios!tc),(com_honorarios.impbru+com_honorarios.impina)/con_tc!impven)) AS comdol, " _
                & " 0 AS vensol, 0 AS vendol, 0 AS ldgsol, 0 AS lgddol, 0 AS reesol, 0 AS reedol, 'H' AS tipo, com_honorarios.id " _
                & " FROM (com_honorarios LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) " _
                & " RIGHT JOIN " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " ON com_honorarios.iddocref2 = [consulta OD].idorddes WHERE (((com_honorarios.idtipdocref)=4)) "
        End If
        
        If NulosN(Fg3.TextMatrix(4, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, vta_gastodebito.idtipdocref, [consulta OD].numerodoc AS numdocref, " _
                & " 0 AS consol, 0 AS condol, 0 AS vensol, 0 AS vendol, " _
                & " IIf([vta_gastodebito].[idmon]=1,(vta_gastodebito.impbru+vta_gastodebito.impina), IIf([vta_gastodebito]![tc]<>0,((vta_gastodebito.impbru+vta_gastodebito.impina)*[vta_gastodebito]![tc]),(vta_gastodebito.impbru+vta_gastodebito.impina)*[con_tc]![impven])) AS lgdsol, " _
                & " IIf([vta_gastodebito].[idmon]=2,(vta_gastodebito.impbru+vta_gastodebito.impina), IIf([vta_gastodebito]![tc]<>0,((vta_gastodebito.impbru+vta_gastodebito.impina)/[vta_gastodebito]![tc]),(vta_gastodebito.impbru+vta_gastodebito.impina)/[con_tc]![impven])) AS lgddol, " _
                & " 0 AS reesol, 0 AS reedol,'L' AS tipo, vta_gastodebito.id " _
                & " FROM " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " LEFT JOIN (vta_gastodebito LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha) ON [consulta OD].idorddes = vta_gastodebito.iddocref2 " _
                & " WHERE (((vta_gastodebito.idtipdocref)=4))"
        End If
        
        If NulosN(Fg3.TextMatrix(3, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT [consulta OD].idcli AS idcliDR, [consulta OD].fchemi, [consulta OD].nombre AS nomcliDR, com_reembolsables.idtipdocref, [consulta OD].numerodoc AS numdocref, " _
                & " 0 AS comsol, 0 AS comdol, 0 AS vensol, 0 AS vendol, 0 AS lgdsol, 0 AS lgddol, " _
                & " IIf(com_reembolsables.idmon=1,(com_reembolsables.impbru+com_reembolsables.impina), IIf(com_reembolsables!tc<>0, ((com_reembolsables.impbru+com_reembolsables.impina)*com_reembolsables!tc), (com_reembolsables.impbru+com_reembolsables.impina)*con_tc!impven)) AS reesol, " _
                & " IIf(com_reembolsables.idmon=2,(com_reembolsables.impbru+com_reembolsables.impina), IIf(com_reembolsables!tc<>0, ((com_reembolsables.impbru+com_reembolsables.impina)/com_reembolsables!tc), (com_reembolsables.impbru+com_reembolsables.impina)/con_tc!impven)) AS reedol, " _
                & " 'R' AS tipo, com_reembolsables.id " _
                & " FROM (com_reembolsables RIGHT JOIN " _
                & " ( " & xCursor _
                & " ) AS [consulta OD] " _
                & " ON com_reembolsables.iddocref2 = [consulta OD].idorddes) LEFT JOIN con_tc ON com_reembolsables.fchdoc = con_tc.fecha " _
                & " WHERE (((com_reembolsables.idtipdocref)=4))"
        End If
    End If
    
    ' CARGAMOS LA CONSULTA FINAL CON LA TABLA VIRTUAL XSQL
    Set Rst = Nothing
    RST_Busq Rst, "SELECT RESUMEN.idcliDR, RESUMEN.fchemi,  RESUMEN.nomcliDR, RESUMEN.idtipdocref, RESUMEN.numdocref, Sum(RESUMEN.comsol) AS SumaDecomsol, " _
        & " Sum(RESUMEN.comdol) AS SumaDecomdol, Sum(RESUMEN.vensol) AS SumaDevensol, Sum(RESUMEN.vendol) AS SumaDevendol, Sum(RESUMEN.lgdsol) AS SumaDelgdsol, " _
        & " Sum(RESUMEN.lgddol) AS SumaDelgddol, Sum(RESUMEN.reesol) AS SumaDereesol, Sum(RESUMEN.reedol) AS SumaDereedol" _
        & " From " _
        & " ( " & xSQL _
        & " ) AS RESUMEN " _
        & " GROUP BY RESUMEN.idcliDR, RESUMEN.fchemi, RESUMEN.nomcliDR, RESUMEN.idtipdocref, RESUMEN.numdocref", xCon

    Rst.Sort = "nomcliDR, numdocref"
    
    Dim xTot1, xTot2, xTot3, xTot4, xTot5, xTot6 As Double
    Dim xTot7, xTot8 As Double
    
    Dim xGTot1, xGTot2, xGTot3, xGTot4, xGTot5, xGTot6 As Double
    Dim xGTot7, xGTot8 As Double
    Dim xNomCli As String
    
    Fg1.Rows = 2
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        xNomCli = Rst("nomcliDR")
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("nomcliDR")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("idtipdocref")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(Rst("numdocref"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format((Rst("fchemi")), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = (Date - Rst("fchemi"))
            
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(Rst("SumaDecomsol"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(Rst("SumaDevensol"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(Rst("SumaDereesol"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(Rst("SumaDelgdsol"), "0.00")
            
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(Rst("SumaDecomdol"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(Rst("SumaDevendol"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(Rst("SumaDereedol"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(Rst("SumaDelgddol"), "0.00")
            
            xTot1 = xTot1 + Rst("SumaDecomsol")
            xTot2 = xTot2 + Rst("SumaDevensol")
            xTot7 = xTot7 + Rst("SumaDereesol")
            xTot3 = xTot3 + Rst("SumaDelgdsol")
            
            xTot4 = xTot4 + Rst("SumaDecomdol")
            xTot5 = xTot5 + Rst("SumaDevendol")
            xTot8 = xTot8 + Rst("SumaDereedol")
            xTot6 = xTot6 + Rst("SumaDelgddol")
            
            Rst.MoveNext
            If Rst.EOF = True Then
                
                Fg1.Rows = Fg1.Rows + 1
                GRID_COMBINAR Fg1, Fg1.Rows - 1, 2, Fg1.Rows - 1, 3, "TOTAL CLIENTE ==> ", flexAlignLeftCenter, True, 1, &H800080, &HE2FEFB, True
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H80000012, True, &HE2FEFB, Format(xTot1, "0.00")
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, &H80000012, True, &HE2FEFB, Format(xTot2, "0.00")
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, &H80000012, True, &HE2FEFB, Format(xTot7, "0.00")
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H80000012, True, &HE2FEFB, Format(xTot3, "0.00")
                
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &H80000012, True, &HE2FEFB, Format(xTot4, "0.00")
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H80000012, True, &HE2FEFB, Format(xTot5, "0.00")
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H80000012, True, &HE2FEFB, Format(xTot8, "0.00")
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H80000012, True, &HE2FEFB, Format(xTot6, "0.00")
            
                Fg1.Rows = Fg1.Rows + 1
                
                xGTot1 = xGTot1 + xTot1
                xGTot2 = xGTot2 + xTot2
                xGTot3 = xGTot3 + xTot3
                xGTot4 = xGTot4 + xTot4
                xGTot5 = xGTot5 + xTot5
                xGTot6 = xGTot6 + xTot6
                xGTot7 = xGTot7 + xTot7
                xGTot8 = xGTot8 + xTot8
                
                xTot1 = 0
                xTot2 = 0
                xTot3 = 0
                xTot4 = 0
                xTot5 = 0
                xTot6 = 0
                xTot7 = 0
                xTot8 = 0
                
                Exit For
            End If
            
            If xNomCli <> Rst("nomcliDR") Then
                Fg1.Rows = Fg1.Rows + 1
                GRID_COMBINAR Fg1, Fg1.Rows - 1, 2, Fg1.Rows - 1, 3, "TOTAL CLIENTE ==> ", flexAlignLeftCenter, True, 1, &H800080, &HE2FEFB, True
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H80000012, True, &HE2FEFB, Format(xTot1, "0.00")
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, &H80000012, True, &HE2FEFB, Format(xTot2, "0.00")
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, &H80000012, True, &HE2FEFB, Format(xTot7, "0.00")
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H80000012, True, &HE2FEFB, Format(xTot3, "0.00")
                
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &H80000012, True, &HE2FEFB, Format(xTot4, "0.00")
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H80000012, True, &HE2FEFB, Format(xTot5, "0.00")
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H80000012, True, &HE2FEFB, Format(xTot8, "0.00")
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H80000012, True, &HE2FEFB, Format(xTot6, "0.00")
            
                Fg1.Rows = Fg1.Rows + 1
                
                xGTot1 = xGTot1 + xTot1
                xGTot2 = xGTot2 + xTot2
                xGTot3 = xGTot3 + xTot3
                xGTot4 = xGTot4 + xTot4
                xGTot5 = xGTot5 + xTot5
                xGTot6 = xGTot6 + xTot6
                xGTot7 = xGTot7 + xTot7
                xGTot8 = xGTot8 + xTot8
                
                xTot1 = 0
                xTot2 = 0
                xTot3 = 0
                xTot4 = 0
                xTot5 = 0
                xTot6 = 0
                xTot7 = 0
                xTot8 = 0
                
                xNomCli = Rst("nomcliDR")
            End If
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        GRID_COMBINAR Fg1, Fg1.Rows - 1, 2, Fg1.Rows - 1, 3, "GRAN TOTAL  ==> ", flexAlignLeftCenter, True, 1, &H800080, &HE2FEFB, True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H80000012, True, &HE2FEFB, Format(xGTot1, "0.00")
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, &H80000012, True, &HE2FEFB, Format(xGTot2, "0.00")
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, &H80000012, True, &HE2FEFB, Format(xGTot7, "0.00")
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H80000012, True, &HE2FEFB, Format(xGTot3, "0.00")
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &H80000012, True, &HE2FEFB, Format(xGTot4, "0.00")
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H80000012, True, &HE2FEFB, Format(xGTot5, "0.00")
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H80000012, True, &HE2FEFB, Format(xGTot8, "0.00")
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H80000012, True, &HE2FEFB, Format(xGTot6, "0.00")
    End If
End Sub

Sub VerDetalle(xCursor As String)
    Dim A As Double
    
    ' CARGAMOS TODAS LAS OPERACIONES
    Dim xSQL  As String
    
    ' CON I.G.V.
    If Option3.Value = True Then
        If NulosN(Fg3.TextMatrix(0, 2)) = -1 Then
            xSQL = "SELECT vta_ventas.idtipdocref, [consulta OD].fchemi, [consulta OD].idcli AS idcliDR, [consulta OD].nombre AS nomcli, [consulta OD].numerodoc AS numdocref, [consulta OD].idcli, [consulta OD].numruc, [consulta OD].nombre, " _
               & " Mid(vta_ventas!numreg,1,2) & mae_libros!codsun & Mid(vta_ventas!numreg,3,4) AS numreg, vta_ventas!numser & '-' & vta_ventas!numdoc AS numdoc, " _
               & " vta_ventas.fchdoc, mae_documento.abrev, vta_ventas.idmon, mae_moneda.simbolo, IIf(vta_ventas!tc<>0,vta_ventas!tc,con_tc!impven) AS tc, " _
               & " vta_ventas.imptotdoc From " _
               & " ( " & xCursor _
               & " ) as [consulta OD] " _
               & " LEFT JOIN ((((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) " _
               & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) " _
               & " ON [consulta OD].idorddes = vta_ventas.iddocref2 Where (((vta_ventas.idtipdocref) = 4) And ((vta_ventas.anulado) = 0)) ORDER BY [consulta OD].nombre"
        End If
        
        If NulosN(Fg3.TextMatrix(1, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT com_compras.idtipdocref, [consulta OD].fchemi, [consulta OD].idcli AS idcliDR, [consulta OD].nombre AS nomcli, [consulta OD].numerodoc AS numdocref, com_compras.idpro, mae_prov.numruc, mae_prov.nombre, " _
                & " Mid(com_compras!numreg,1,2) & mae_libros!codsun & Mid(com_compras!numreg,3,4) AS numreg, com_compras!numser & '-' & com_compras!numdoc AS numdoc, " _
                & " com_compras.fchdoc, mae_documento.abrev, com_compras.idmon, mae_moneda.simbolo, IIf(com_compras!tc<>0,com_compras!tc,con_tc!impven) AS tc, " _
                & " com_compras.imptot " _
                & " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN " _
                & " ((( " _
                & " ( " & xCursor _
                & " ) as [consulta OD] " _
                & " LEFT JOIN com_compras ON [consulta OD].idorddes = com_compras.iddocref2) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) " _
                & " LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) " _
                & " ON mae_prov.id = com_compras.idpro WHERE (((com_compras.idtipdocref)=4))"
        End If
        
        If NulosN(Fg3.TextMatrix(2, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT com_honorarios.idtipdocref, [consulta OD].fchemi, [consulta OD].idcli AS idcliDR, [consulta OD].nombre AS nomcli, [consulta OD].numerodoc AS numdocref, com_honorarios.idpro, mae_prov.numruc, mae_prov.nombre, " _
                & " Mid(com_honorarios!numreg,1,2) & mae_libros!codsun & Mid(com_honorarios!numreg,3,4) AS numreg, com_honorarios!numser & '-' & com_honorarios!numdoc AS numdoc, " _
                & " com_honorarios.fchdoc, mae_documento.abrev, com_honorarios.idmon, mae_moneda.simbolo, IIf(com_honorarios!tc<>0,com_honorarios!tc,con_tc!impven) AS tc, " _
                & " com_honorarios.imptot " _
                & " FROM " _
                & " ( " & xCursor _
                & " ) as [consulta OD] " _
                & " LEFT JOIN (((((com_honorarios LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) LEFT JOIN mae_documento ON com_honorarios.tipdoc = mae_documento.id) " _
                & " LEFT JOIN mae_prov ON com_honorarios.idpro = mae_prov.id) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) LEFT JOIN mae_moneda " _
                & " ON com_honorarios.idmon = mae_moneda.id) ON [consulta OD].idorddes = com_honorarios.iddocref2 WHERE (((com_honorarios.idtipdocref)=4))"
        End If
        
        If NulosN(Fg3.TextMatrix(3, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT com_reembolsables.idtipdocref, [consulta OD].fchemi, [consulta OD].idcli AS idcliDR, [consulta OD].nombre AS nomcli, [consulta OD].numerodoc AS numdocref, com_reembolsables.idpro, mae_prov.numruc, mae_prov.nombre, '00XX0000' AS numreg, " _
                & " com_reembolsables!numser & '-' & com_reembolsables!numdoc AS numdoc, com_reembolsables.fchdoc, mae_documento.abrev, com_reembolsables.idmon, " _
                & " mae_moneda.simbolo, IIf(com_reembolsables!tc<>0,com_reembolsables!tc,con_tc!impven) AS tc, com_reembolsables.imptot " _
                & " FROM ((((" _
                & " ( " & xCursor _
                & " ) as [consulta OD] " _
                & " LEFT JOIN com_reembolsables ON [consulta OD].idorddes = com_reembolsables.iddocref2) LEFT JOIN mae_documento ON com_reembolsables.tipdoc = mae_documento.id) " _
                & " LEFT JOIN mae_moneda ON com_reembolsables.idmon = mae_moneda.id) LEFT JOIN mae_prov ON com_reembolsables.idpro = mae_prov.id) LEFT JOIN con_tc " _
                & " ON com_reembolsables.fchdoc = con_tc.fecha WHERE (((com_reembolsables.idtipdocref)=4))"
        End If
        
        If NulosN(Fg3.TextMatrix(4, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT vta_gastodebito.idtipdocref, [consulta OD].fchemi, [consulta OD].idcli AS idcliDR, [consulta OD].nombre AS nomcli, [consulta OD].numerodoc, vta_gastodebito.idcli, mae_cliente.numruc, mae_cliente.nombre, '00LL0000' AS numreg, " _
                & " vta_gastodebito!numser & '-' & vta_gastodebito!numdoc AS numdoc, vta_gastodebito.fchemi AS fchdoc, mae_documento.abrev, vta_gastodebito.idmon, " _
                & " mae_moneda.simbolo, IIf(vta_gastodebito!tc<>0,vta_gastodebito!tc,con_tc!impven) AS tc, vta_gastodebito.imptot " _
                & " FROM ((((" _
                & " ( " & xCursor _
                & " ) as [consulta OD] " _
                & " LEFT JOIN vta_gastodebito ON [consulta OD].idorddes = vta_gastodebito.iddocref2) LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) " _
                & " LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha) LEFT JOIN mae_cliente " _
                & " ON vta_gastodebito.idcli = mae_cliente.id WHERE (((vta_gastodebito.idtipdocref)=4) AND ((vta_gastodebito.anulado)=0))"
        End If
    End If
        
    '************************************************************************************************************************************
    ' SIN I.G.V.
    If Option4.Value = True Then
        If NulosN(Fg3.TextMatrix(0, 2)) = -1 Then
            xSQL = "SELECT vta_ventas.idtipdocref, [consulta OD].fchemi, [consulta OD].idcli AS idcliDR, [consulta OD].nombre AS nomcli, [consulta OD].numerodoc AS numdocref, [consulta OD].idcli, [consulta OD].numruc, [consulta OD].nombre, " _
                & " Mid(vta_ventas!numreg,1,2) & mae_libros!codsun & Mid(vta_ventas!numreg,3,4) AS numreg, vta_ventas!numser & '-' & vta_ventas!numdoc AS numdoc, " _
                & " vta_ventas.fchdoc, mae_documento.abrev, vta_ventas.idmon, mae_moneda.simbolo, IIf(vta_ventas!tc<>0,vta_ventas!tc,con_tc!impven) AS tc, " _
                & " vta_ventas.impbru + vta_ventas.impbru2 + vta_ventas.impbru3 + vta_ventas.impinaf AS imptotdoc From " _
                & " ( " & xCursor _
                & " ) as [consulta OD] " _
                & " LEFT JOIN ((((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) " _
                & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) " _
                & " ON [consulta OD].idorddes = vta_ventas.iddocref2 Where (((vta_ventas.idtipdocref) = 4) And ((vta_ventas.anulado) = 0)) ORDER BY [consulta OD].nombre"
        End If
        
        If NulosN(Fg3.TextMatrix(1, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT com_compras.idtipdocref, [consulta OD].fchemi, [consulta OD].idcli AS idcliDR, [consulta OD].nombre AS nomcli, [consulta OD].numerodoc AS numdocref, com_compras.idpro, mae_prov.numruc, mae_prov.nombre, " _
                & " Mid(com_compras!numreg,1,2) & mae_libros!codsun & Mid(com_compras!numreg,3,4) AS numreg, com_compras!numser & '-' & com_compras!numdoc AS numdoc, " _
                & " com_compras.fchdoc, mae_documento.abrev, com_compras.idmon, mae_moneda.simbolo, IIf(com_compras!tc<>0,com_compras!tc,con_tc!impven) AS tc, " _
                & " com_compras.impbru + com_compras.impbru2 + com_compras.impbru3 + com_compras.impina  AS imptot " _
                & " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN " _
                & " ((( " _
                & " ( " & xCursor _
                & " ) as [consulta OD] " _
                & " LEFT JOIN com_compras ON [consulta OD].idorddes = com_compras.iddocref2) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) " _
                & " LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) " _
                & " ON mae_prov.id = com_compras.idpro WHERE (((com_compras.idtipdocref)=4))"
        End If
        
        If NulosN(Fg3.TextMatrix(2, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT com_honorarios.idtipdocref, [consulta OD].fchemi, [consulta OD].idcli AS idcliDR, [consulta OD].nombre AS nomcli, [consulta OD].numerodoc AS numdocref, com_honorarios.idpro, mae_prov.numruc, mae_prov.nombre, " _
                & " Mid(com_honorarios!numreg,1,2) & mae_libros!codsun & Mid(com_honorarios!numreg,3,4) AS numreg, com_honorarios!numser & '-' & com_honorarios!numdoc AS numdoc, " _
                & " com_honorarios.fchdoc, mae_documento.abrev, com_honorarios.idmon, mae_moneda.simbolo, IIf(com_honorarios!tc<>0,com_honorarios!tc,con_tc!impven) AS tc, " _
                & " com_honorarios.impbru + com_honorarios.impina AS imptot " _
                & " FROM " _
                & " ( " & xCursor _
                & " ) as [consulta OD] " _
                & " LEFT JOIN (((((com_honorarios LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) LEFT JOIN mae_documento ON com_honorarios.tipdoc = mae_documento.id) " _
                & " LEFT JOIN mae_prov ON com_honorarios.idpro = mae_prov.id) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) LEFT JOIN mae_moneda " _
                & " ON com_honorarios.idmon = mae_moneda.id) ON [consulta OD].idorddes = com_honorarios.iddocref2 WHERE (((com_honorarios.idtipdocref)=4))"
        End If
        
        If NulosN(Fg3.TextMatrix(3, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT com_reembolsables.idtipdocref, [consulta OD].fchemi, [consulta OD].idcli AS idcliDR, [consulta OD].nombre AS nomcli, [consulta OD].numerodoc AS numdocref, com_reembolsables.idpro, mae_prov.numruc, mae_prov.nombre, '00XX0000' AS numreg, " _
                & " com_reembolsables!numser & '-' & com_reembolsables!numdoc AS numdoc, com_reembolsables.fchdoc, mae_documento.abrev, com_reembolsables.idmon, " _
                & " mae_moneda.simbolo, IIf(com_reembolsables!tc<>0,com_reembolsables!tc,con_tc!impven) AS tc, " _
                & " com_reembolsables.impbru + com_reembolsables.impina AS imptot " _
                & " FROM ((((" _
                & " ( " & xCursor _
                & " ) as [consulta OD] " _
                & " LEFT JOIN com_reembolsables ON [consulta OD].idorddes = com_reembolsables.iddocref2) LEFT JOIN mae_documento ON com_reembolsables.tipdoc = mae_documento.id) " _
                & " LEFT JOIN mae_moneda ON com_reembolsables.idmon = mae_moneda.id) LEFT JOIN mae_prov ON com_reembolsables.idpro = mae_prov.id) LEFT JOIN con_tc " _
                & " ON com_reembolsables.fchdoc = con_tc.fecha WHERE (((com_reembolsables.idtipdocref)=4))"
        End If
    
        If NulosN(Fg3.TextMatrix(4, 2)) = -1 Then
            xSQL = xSQL & " UNION " _
                & " SELECT vta_gastodebito.idtipdocref, [consulta OD].fchemi, [consulta OD].idcli AS idcliDR, [consulta OD].nombre AS nomcli, [consulta OD].numerodoc, vta_gastodebito.idcli, mae_cliente.numruc, mae_cliente.nombre, '00LL0000' AS numreg, " _
                & " vta_gastodebito!numser & '-' & vta_gastodebito!numdoc AS numdoc, vta_gastodebito.fchemi AS fchdoc, mae_documento.abrev, vta_gastodebito.idmon, " _
                & " mae_moneda.simbolo, IIf(vta_gastodebito!tc<>0,vta_gastodebito!tc,con_tc!impven) AS tc, " _
                & " vta_gastodebito.impbru + vta_gastodebito.impina  AS imptot " _
                & " FROM ((((" _
                & " ( " & xCursor _
                & " ) as [consulta OD] " _
                & " LEFT JOIN vta_gastodebito ON [consulta OD].idorddes = vta_gastodebito.iddocref2) LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) " _
                & " LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha) LEFT JOIN mae_cliente " _
                & " ON vta_gastodebito.idcli = mae_cliente.id WHERE (((vta_gastodebito.idtipdocref)=4) AND ((vta_gastodebito.anulado)=0))"
        End If
    End If
    
    Set Rst = Nothing
    RST_Busq Rst, xSQL, xCon
    
    Rst.Sort = "nomcli, numdocref, fchdoc"
    
    Dim xIdCli As Integer
    Dim NumDocRef, xCliente As String
    
    If Rst.RecordCount <> 0 Then
        Frame5.Left = (Me.Width - Frame5.Width) / 2
        Frame5.Top = (Me.Height - Frame5.Height) / 2
        
        ProgressBar1.Max = Rst.RecordCount
        Frame5.Visible = True
        
        Rst.MoveFirst
        Fg1.Rows = 2
        xIdCli = Rst("idcliDR")
        NumDocRef = Rst("numdocref")
        
        xTotComSol = 0
        xTotVenSol = 0
        xTotComDol = 0
        xTotVenDol = 0
        xTotLGDSol = 0
        xTotLGDDol = 0
        xTotReeDol = 0
        xTotReeSol = 0
        
        xTotVenSolTot = 0
        xTotVenDolTot = 0
        xTotComSolTot = 0
        xTotComDolTot = 0
        xTotLGDSolTot = 0
        xTotLGDDolTot = 0
        xTotReeSolTot = 0
        xTotReeDolTot = 0
        
        For A = 1 To Rst.RecordCount
            ProgressBar1.Value = A
            
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("nomcli")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("idtipdocref")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Rst("numdocref")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(Rst("fchemi"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = (Date - Rst("fchemi"))
            
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Rst("numreg")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(Rst("numruc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(Rst("nombre"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(Rst("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosC(Rst("numdoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(Rst("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosC(Rst("simbolo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(Rst("imptotdoc"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(Rst("tc"), "0.000")
            
            If Rst("idmon") = 1 Then
                If Mid(Rst("numreg"), 3, 2) = "14" Then ' VENTAS
                    If Rst("abrev") = "NC" Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(0 - Rst("imptotdoc"), "0.00")
                        Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(0 - Rst("imptotdoc") / Rst("tc"), "0.00")
                        
                        xTotVenSol = xTotVenSol + (0 - Rst("imptotdoc"))
                        xTotVenDol = xTotVenDol + (0 - (Rst("imptotdoc") / Rst("tc")))
                        
                        xTotVenSolTot = xTotVenSolTot + (0 - Rst("imptotdoc")) '+ xTotVenSol
                        xTotVenDolTot = xTotVenDolTot + (0 - (Rst("imptotdoc") / Rst("tc"))) '+ xTotVenDol
                    Else
                        Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(Rst("imptotdoc"), "0.00")
                        Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(Rst("imptotdoc") / Rst("tc"), "0.00")
                        
                        xTotVenSol = xTotVenSol + Rst("imptotdoc")
                        xTotVenDol = xTotVenDol + (Rst("imptotdoc") / Rst("tc"))
                        
                        xTotVenSolTot = xTotVenSolTot + Rst("imptotdoc") '+ xTotVenSol
                        xTotVenDolTot = xTotVenDolTot + (Rst("imptotdoc") / Rst("tc")) '+ xTotVenDol
                    End If
                End If
                
                If Mid(Rst("numreg"), 3, 2) = "08" Or Mid(Rst("numreg"), 3, 2) = "50" Then  ' COMPRAS, HONORARIOS
                    If Rst("abrev") = "NC" Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(Rst("imptotdoc"), "0.00")
                        Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(Rst("imptotdoc") / Rst("tc"), "0.00")
                        
                        xTotComSol = xTotComSol + (0 - Rst("imptotdoc"))
                        xTotComDol = xTotComDol + (0 - (Rst("imptotdoc") / Rst("tc")))
                        
                        xTotComSolTot = xTotComSolTot + (0 - Rst("imptotdoc")) '+ xTotComSol
                        xTotComDolTot = xTotComDolTot + (0 - (Rst("imptotdoc") / Rst("tc"))) '+ xTotComDol
                    Else
                        Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(Rst("imptotdoc"), "0.00")
                        Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(Rst("imptotdoc") / Rst("tc"), "0.00")
                        
                        xTotComSol = xTotComSol + Rst("imptotdoc")
                        xTotComDol = xTotComDol + (Rst("imptotdoc") / Rst("tc"))
                        
                        xTotComSolTot = xTotComSolTot + Rst("imptotdoc") '+ xTotComSol
                        xTotComDolTot = xTotComDolTot + (Rst("imptotdoc") / Rst("tc")) '+ xTotComDol
                    End If
                End If
                
                If Mid(Rst("numreg"), 3, 2) = "XX" Then  ' REEMBOLSABLES
                    Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(Rst("imptotdoc"), "0.00")
                    Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(Rst("imptotdoc") / Rst("tc"), "0.00")
                    
                    xTotReeSol = xTotReeSol + Rst("imptotdoc")
                    xTotReeDol = xTotReeDol + (Rst("imptotdoc") / Rst("tc"))

                    xTotReeSolTot = xTotReeSolTot + Rst("imptotdoc") '+ xTotComSol
                    xTotReeDolTot = xTotReeDolTot + (Rst("imptotdoc") / Rst("tc")) '+ xTotComDol
                End If
                
                If Mid(Rst("numreg"), 3, 2) = "LL" Then ' LIQUIDACION GASTO DEBITO
                    Fg1.TextMatrix(Fg1.Rows - 1, 18) = Format(Rst("imptotdoc"), "0.00")
                    Fg1.TextMatrix(Fg1.Rows - 1, 22) = Format(Rst("imptotdoc") / Rst("tc"), "0.00")
                    
                    xTotLGDSol = xTotLGDSol + Rst("imptotdoc")
                    xTotLGDDol = xTotLGDDol + (Rst("imptotdoc") / Rst("tc"))
                                        
                    xTotLGDSolTot = xTotLGDSolTot + Rst("imptotdoc") '+ xTotLGDSol
                    xTotLGDDolTot = xTotLGDDolTot + (Rst("imptotdoc") / Rst("tc")) '+ xTotLGDDol
                End If
                                
            Else
                If Mid(Rst("numreg"), 3, 2) = "14" Then ' VENTAS
                    If Rst("abrev") = "NC" Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(Rst("imptotdoc") * Rst("tc"), "0.00")
                        Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(Rst("imptotdoc"), "0.00")
                    
                        xTotVenSol = xTotVenSol + (0 - (Rst("imptotdoc") * Rst("tc")))
                        xTotVenDol = xTotVenDol + (0 - Rst("imptotdoc"))
                        
                        xTotVenSolTot = xTotVenSolTot + (0 - (Rst("imptotdoc") * Rst("tc"))) ' + xTotVenSol
                        xTotVenDolTot = xTotVenDolTot + (0 - Rst("imptotdoc")) '+ xTotVenDol
                    Else
                        Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(Rst("imptotdoc") * Rst("tc"), "0.00")
                        Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(Rst("imptotdoc"), "0.00")
                    
                        xTotVenSol = xTotVenSol + (Rst("imptotdoc") * Rst("tc"))
                        xTotVenDol = xTotVenDol + Rst("imptotdoc")
                        
                        xTotVenSolTot = xTotVenSolTot + (Rst("imptotdoc") * Rst("tc")) ' + xTotVenSol
                        xTotVenDolTot = xTotVenDolTot + Rst("imptotdoc") '+ xTotVenDol
                    End If
                End If
                
                If Mid(Rst("numreg"), 3, 2) = "08" Or Mid(Rst("numreg"), 3, 2) = "50" Then  ' COMPRAS, HONORARIOS
                    If Rst("abrev") = "NC" Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(Rst("imptotdoc") * Rst("tc"), "0.00")
                        Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(Rst("imptotdoc"), "0.00")
                    
                        xTotComSol = xTotComSol + (0 - (Rst("imptotdoc") * Rst("tc")))
                        xTotComDol = xTotComDol + (0 - Rst("imptotdoc"))
                        
                        xTotComSolTot = xTotComSolTot + (0 - (Rst("imptotdoc") * Rst("tc"))) '+ xTotComSol
                        xTotComDolTot = xTotComDolTot + (0 - Rst("imptotdoc")) ' + xTotComDol
                    Else
                        Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(Rst("imptotdoc") * Rst("tc"), "0.00")
                        Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(Rst("imptotdoc"), "0.00")
                    
                        xTotComSol = xTotComSol + (Rst("imptotdoc") * Rst("tc"))
                        xTotComDol = xTotComDol + Rst("imptotdoc")
                        
                        xTotComSolTot = xTotComSolTot + (Rst("imptotdoc") * Rst("tc")) '+ xTotComSol
                        xTotComDolTot = xTotComDolTot + Rst("imptotdoc") ' + xTotComDol
                    End If
                End If
                
                If Mid(Rst("numreg"), 3, 2) = "XX" Then   ' REEMBOLSABLES
                    Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(Rst("imptotdoc") * Rst("tc"), "0.00")
                    Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(Rst("imptotdoc"), "0.00")
                
                    xTotReeSol = xTotReeSol + (Rst("imptotdoc") * Rst("tc"))
                    xTotReeDol = xTotReeDol + Rst("imptotdoc")
                    
                    xTotReeSolTot = xTotReeSolTot + (Rst("imptotdoc") * Rst("tc")) '+ xTotComSol
                    xTotReeDolTot = xTotReeDolTot + Rst("imptotdoc") ' + xTotComDol
                End If
                
                If Mid(Rst("numreg"), 3, 2) = "LL" Then ' LIQUIDACION GASTO DEBITO
                    Fg1.TextMatrix(Fg1.Rows - 1, 18) = Format(Rst("imptotdoc") * Rst("tc"), "0.00")
                    Fg1.TextMatrix(Fg1.Rows - 1, 22) = Format(Rst("imptotdoc"), "0.00")
                    
                    xTotLGDSol = xTotLGDSol + (Rst("imptotdoc") * Rst("tc"))
                    xTotLGDDol = xTotLGDDol + Rst("imptotdoc")
                    
                    xTotLGDSolTot = xTotLGDSolTot + (Rst("imptotdoc") * Rst("tc")) '+ xTotLGDSol
                    xTotLGDDolTot = xTotLGDDolTot + Rst("imptotdoc") '+ xTotLGDDol
                End If
                
            End If
            
            xCliente = Rst("nomcli")
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Totalizar
                Fg1.Rows = Fg1.Rows + 1
                TotalizarCliente
                Exit For
            End If
            
            ' si el numero de documento de referencia a cambiado
            If NumDocRef <> Rst("numdocref") Then
                Totalizar
                Fg1.Rows = Fg1.Rows + 1
            End If
            
            NumDocRef = Rst("numdocref")
            
            If xIdCli <> Rst("idcliDR") Then
                TotalizarCliente
                Fg1.Rows = Fg1.Rows + 2
                xIdCli = Rst("idcliDR")
            End If
        Next A
        
        Frame5.Visible = False
    End If

End Sub

Sub TotalizarCliente()
    'totalizamos
    Fg1.Rows = Fg1.Rows + 1
    
    GRID_COMBINAR Fg1, Fg1.Rows - 1, 8, Fg1.Rows - 1, 11, "TOTAL CLIENTE ==> ", flexAlignLeftCenter, True, 1, &H800080, &HE2FEFB, True
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 15, &H80000012, True, &HE2FEFB, Format(xTotComSolTot, "0.00")
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &H80000012, True, &HE2FEFB, Format(xTotVenSolTot, "0.00")
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 17, &H80000012, True, &HE2FEFB, Format(xTotReeSolTot, "0.00")
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &H80000012, True, &HE2FEFB, Format(xTotLGDSolTot, "0.00")
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 19, &H80000012, True, &HE2FEFB, Format(xTotComDolTot, "0.00")
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 20, &H80000012, True, &HE2FEFB, Format(xTotVenDolTot, "0.00")
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 21, &H80000012, True, &HE2FEFB, Format(xTotReeDolTot, "0.00")
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 22, &H80000012, True, &HE2FEFB, Format(xTotLGDDolTot, "0.00")
    
    Fg1.Rows = Fg1.Rows + 1
    GRID_COMBINAR Fg1, Fg1.Rows - 1, 8, Fg1.Rows - 1, 11, "SALDO CLIENTE ==> ", flexAlignLeftCenter, True, 1, &H800080, &HE2FEFB, True
    
    If (xTotVenSolTot - xTotComSolTot) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &HC0&, True, &HE2FEFB, Format(xTotVenSolTot - xTotComSolTot, "0.00")
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &HC00000, True, &HE2FEFB, Format(xTotVenSolTot - xTotComSolTot, "0.00")
    End If
    
    If (xTotVenDolTot - xTotComDolTot) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 20, &HC0&, True, &HE2FEFB, Format(xTotVenDolTot - xTotComDolTot, "0.00")
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 20, &HC00000, True, &HE2FEFB, Format(xTotVenDolTot - xTotComDolTot, "0.00")
    End If
    
    '----------------------------------------------------------------------------------------------------------------
    If (xTotLGDSolTot - xTotReeSolTot) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &HC0&, True, &HE2FEFB, Format(xTotLGDSolTot - xTotReeSolTot, "0.00")
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &HC00000, True, &HE2FEFB, Format(xTotLGDSolTot - xTotReeSolTot, "0.00")
    End If
    
    If (xTotLGDDolTot - xTotReeDolTot) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 22, &HC0&, True, &HE2FEFB, Format(xTotLGDDolTot - xTotReeDolTot, "0.00")
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 22, &HC00000, True, &HE2FEFB, Format(xTotLGDDolTot - xTotReeDolTot, "0.00")
    End If
    
    xTotComSolTot = 0
    xTotComDolTot = 0
    xTotVenSolTot = 0
    xTotVenDolTot = 0
    xTotLGDDolTot = 0
    xTotLGDSolTot = 0
    xTotReeDolTot = 0
    xTotReeSolTot = 0
End Sub

Sub Totalizar()
    'totalizamos
    Fg1.Rows = Fg1.Rows + 1
    
    GRID_COMBINAR Fg1, Fg1.Rows - 1, 8, Fg1.Rows - 1, 11, "TOTAL ORDEN ==> ", flexAlignLeftCenter, True, 1, &H800000, &HE2FEFB, True
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 15, &H80000012, True, &HE2FEFB, Format(xTotComSol, "0.00")
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &H80000012, True, &HE2FEFB, Format(xTotVenSol, "0.00")
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 17, &H80000012, True, &HE2FEFB, Format(xTotReeSol, "0.00")
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &H80000012, True, &HE2FEFB, Format(xTotLGDSol, "0.00")
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 19, &H80000012, True, &HE2FEFB, Format(xTotComDol, "0.00")
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 20, &H80000012, True, &HE2FEFB, Format(xTotVenDol, "0.00")
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 21, &H80000012, True, &HE2FEFB, Format(xTotReeDol, "0.00")
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 22, &H80000012, True, &HE2FEFB, Format(xTotLGDDol, "0.00")
    
    If xTotLGDSol <> 0 And xTotVenSol <> 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 23, &H80000012, True, &HE2FEFB, Format((xTotVenSol / xTotLGDSol) * 100, "0.00")
    End If
    
    Fg1.Rows = Fg1.Rows + 1
    
    GRID_COMBINAR Fg1, Fg1.Rows - 1, 8, Fg1.Rows - 1, 11, "SALDO ORDEN ==> ", flexAlignLeftCenter, True, 1, &H800000, &HE2FEFB, True
    
    If (xTotVenSol - xTotComSol) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &HFF&, True, &HE2FEFB, Format(xTotVenSol - xTotComSol, "0.00")
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &HFF0000, True, &HE2FEFB, Format(xTotVenSol - xTotComSol, "0.00")
    End If
    
    If (xTotVenDol - xTotComDol) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 20, &HFF&, True, &HE2FEFB, Format(xTotVenDol - xTotComDol, "0.00")
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 20, &HFF0000, True, &HE2FEFB, Format(xTotVenDol - xTotComDol, "0.00")
    End If
    
    '--------------------------------------------------------------------------------------------------------------
    If (xTotLGDSol - xTotReeSol) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &HFF&, True, &HE2FEFB, Format(xTotLGDSol - xTotReeSol, "0.00")
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &HFF0000, True, &HE2FEFB, Format(xTotLGDSol - xTotReeSol, "0.00")
    End If
    
    If (xTotLGDDol - xTotReeDol) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 22, &HFF&, True, &HE2FEFB, Format(xTotLGDDol - xTotReeDol, "0.00")
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 22, &HFF0000, True, &HE2FEFB, Format(xTotLGDDol - xTotReeDol, "0.00")
    End If
    xTotComSol = 0
    xTotVenSol = 0
    xTotComDol = 0
    xTotVenDol = 0
    xTotLGDSol = 0
    xTotLGDDol = 0
    xTotReeSol = 0
    xTotReeDol = 0
End Sub

Private Sub menu_01_Click()
    If Fg2.Rows = 0 Then
        Fg2.Rows = Fg2.Rows + 1
        Exit Sub
    End If
    
    If NulosC(Fg2.TextMatrix(Fg2.Rows - 1, 1)) <> "" Then
        Fg2.Rows = Fg2.Rows + 1
    End If
End Sub

Private Sub menu_03_Click()
    If Fg2.Rows <> 0 Then
        Fg2.RemoveItem Fg2.Row
    End If
    If Fg2.Rows = 0 Then
        Fg2.Rows = Fg2.Rows + 1
    End If
End Sub

Private Sub Option1_Click()
    If NulosN(Fg4.TextMatrix(0, 2)) = -1 Then
        If Option1.Value = True Then SetearCuadricula Fg1, 7, xCon, 2, 1, False
    Else
        If Option1.Value = True Then SetearCuadricula Fg1, 7, xCon, 2, 3, False
    End If
End Sub

Private Sub Option2_Click()
    If NulosN(Fg4.TextMatrix(0, 2)) = -1 Then
        If Option2.Value = True Then SetearCuadricula Fg1, 7, xCon, 2, 2, False
    Else
        If Option2.Value = True Then SetearCuadricula Fg1, 7, xCon, 2, 4, False
    End If
End Sub
