VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form FrmPuntoVenta 
   Caption         =   "Ventas - Punto de Venta"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   7170
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   11385
      _cx             =   20082
      _cy             =   12647
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
      BorderWidth     =   6
      ChildSpacing    =   4
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
      GridRows        =   4
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmPuntoVenta.frx":0000
      Begin SizerOneLibCtl.ElasticOne ElasticOne4 
         Height          =   750
         Left            =   90
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   6330
         Width           =   11205
         _cx             =   19764
         _cy             =   1323
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
         ChildSpacing    =   6
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
         GridCols        =   3
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmPuntoVenta.frx":005C
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame11"
            Height          =   690
            Left            =   7650
            TabIndex        =   57
            Top             =   30
            Width           =   3525
            Begin VB.CommandButton CmdCancel 
               Caption         =   "&Cancelar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   1800
               TabIndex        =   59
               Top             =   60
               Width           =   1400
            End
            Begin VB.CommandButton CmdSave 
               Caption         =   "&Grabar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   360
               TabIndex        =   58
               Top             =   60
               Width           =   1400
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            Height          =   690
            Left            =   3960
            TabIndex        =   56
            Top             =   30
            Width           =   3600
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "F4 = Cancelar Ingreso"
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
               Left            =   195
               TabIndex        =   62
               Top             =   465
               Width           =   1890
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "F3 = Agregar Cliente"
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
               Left            =   195
               TabIndex        =   61
               Top             =   255
               Width           =   1755
            End
            Begin VB.Label Label13 
               Caption         =   "F2 = Grabar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   195
               TabIndex        =   60
               Top             =   60
               Width           =   1155
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame9"
            Height          =   690
            Left            =   30
            TabIndex        =   55
            Top             =   30
            Width           =   3840
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   840
         Left            =   90
         TabIndex        =   21
         Top             =   5430
         Width           =   11205
         Begin SizerOneLibCtl.ElasticOne ElasticOne2 
            Height          =   810
            Left            =   15
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   15
            Width           =   11205
            _cx             =   19764
            _cy             =   1429
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
            BackColor       =   16777152
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   2
            ChildSpacing    =   4
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
            GridCols        =   2
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmPuntoVenta.frx":00AD
            Begin VB.Frame Frame6 
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   750
               Left            =   4485
               TabIndex        =   24
               Top             =   30
               Width           =   6690
               Begin VB.TextBox TxtImpTot 
                  Height          =   300
                  Left            =   5340
                  TabIndex        =   36
                  Text            =   "TxtImpTot"
                  Top             =   345
                  Width           =   1300
               End
               Begin VB.TextBox TxtImpIsc 
                  Height          =   300
                  Left            =   4005
                  TabIndex        =   31
                  Text            =   "TxtImpIsc"
                  Top             =   345
                  Width           =   1300
               End
               Begin VB.TextBox TxtImpIgv 
                  Height          =   300
                  Left            =   2685
                  TabIndex        =   30
                  Text            =   "TxtImpIgv"
                  Top             =   345
                  Width           =   1300
               End
               Begin VB.TextBox TxtImpIna 
                  Height          =   300
                  Left            =   1365
                  TabIndex        =   29
                  Text            =   "TxtImpIna"
                  Top             =   345
                  Width           =   1300
               End
               Begin VB.TextBox txtImpAfe 
                  Height          =   300
                  Left            =   45
                  TabIndex        =   28
                  Text            =   "txtImpAfe"
                  Top             =   345
                  Width           =   1300
               End
               Begin VB.Label Label7 
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
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   5355
                  TabIndex        =   37
                  Top             =   105
                  Width           =   450
               End
               Begin VB.Label Label6 
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
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   4020
                  TabIndex        =   35
                  Top             =   105
                  Width           =   495
               End
               Begin VB.Label Label5 
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
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   2700
                  TabIndex        =   34
                  Top             =   105
                  Width           =   510
               End
               Begin VB.Label Label4 
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
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   1395
                  TabIndex        =   33
                  Top             =   105
                  Width           =   1140
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Imp. Afecto"
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
                  Left            =   45
                  TabIndex        =   32
                  Top             =   105
                  Width           =   990
               End
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   750
               Left            =   30
               TabIndex        =   23
               Top             =   30
               Width           =   4395
               Begin VB.CommandButton CmdSelItem 
                  Caption         =   "Seleccionar Item"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   555
                  Left            =   1575
                  TabIndex        =   27
                  Top             =   90
                  Width           =   1290
               End
               Begin VB.CommandButton CmdDelItem 
                  Caption         =   "Eliminar Item"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   555
                  Left            =   2910
                  TabIndex        =   26
                  Top             =   90
                  Width           =   1290
               End
               Begin VB.CommandButton CmdAddItem 
                  Caption         =   "Agregar Item"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   555
                  Left            =   240
                  TabIndex        =   25
                  Top             =   90
                  Width           =   1290
               End
            End
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg1 
         Height          =   3660
         Left            =   90
         TabIndex        =   14
         Top             =   1710
         Width           =   11205
         _cx             =   19764
         _cy             =   6456
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
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmPuntoVenta.frx":00EB
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   1560
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   11205
         Begin VB.CommandButton CmdBusMoneda 
            Height          =   240
            Left            =   5715
            Picture         =   "FrmPuntoVenta.frx":024E
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   135
            Width           =   240
         End
         Begin VB.CommandButton CmdBusCli 
            Height          =   240
            Left            =   2355
            Picture         =   "FrmPuntoVenta.frx":0380
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   450
            Width           =   240
         End
         Begin XtremeSuiteControls.RadioButton RadioButton1 
            Height          =   255
            Left            =   210
            TabIndex        =   63
            Top             =   1275
            Width           =   1170
            _Version        =   786432
            _ExtentX        =   2064
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "BOLETA"
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Porcentaje"
            Height          =   195
            Left            =   6615
            TabIndex        =   20
            Top             =   1305
            Width           =   1140
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Valor"
            Height          =   195
            Left            =   5520
            TabIndex        =   19
            Top             =   1305
            Width           =   765
         End
         Begin VB.Frame Frame3 
            Caption         =   "[ Documento ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   8160
            TabIndex        =   15
            Top             =   15
            Width           =   3045
            Begin VB.Label LblNumDoc 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "LblNumDoc"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   360
               Left            =   855
               TabIndex        =   18
               Top             =   465
               Width           =   1995
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Nº -"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   360
               Left            =   120
               TabIndex        =   17
               Top             =   465
               Width           =   585
            End
            Begin VB.Label LblDocumento 
               Alignment       =   2  'Center
               Caption         =   "LblDocumento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   120
               TabIndex        =   16
               Top             =   210
               Width           =   2775
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "[ Tipo Cambio ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   8160
            TabIndex        =   12
            Top             =   915
            Width           =   3045
            Begin VB.Label LblTc 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblTc"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   780
               TabIndex        =   13
               Top             =   255
               Width           =   1530
            End
         End
         Begin VB.TextBox TxtIdMoneda 
            Height          =   300
            Left            =   5175
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "TxtIdMoneda"
            Top             =   105
            Width           =   810
         End
         Begin VB.TextBox TxtDirCli 
            Height          =   480
            Left            =   1035
            TabIndex        =   8
            Text            =   "TxtDirCli"
            Top             =   735
            Width           =   7095
         End
         Begin VB.TextBox TxtNomCli 
            Height          =   300
            Left            =   2700
            TabIndex        =   6
            Text            =   "TxtNomCli"
            Top             =   420
            Width           =   5415
         End
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1035
            TabIndex        =   4
            Text            =   "TxtNumRuc"
            Top             =   420
            Width           =   1590
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   1035
            TabIndex        =   2
            Top             =   105
            Width           =   1245
            _ExtentX        =   2196
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
         End
         Begin XtremeSuiteControls.RadioButton RadioButton2 
            Height          =   255
            Left            =   1680
            TabIndex        =   64
            Top             =   1275
            Width           =   1170
            _Version        =   786432
            _ExtentX        =   2064
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "FACTURA"
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadioButton3 
            Height          =   255
            Left            =   3165
            TabIndex        =   65
            Top             =   1275
            Width           =   1530
            _Version        =   786432
            _ExtentX        =   2699
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "COTIZACION"
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000009&
            X1              =   5205
            X2              =   5205
            Y1              =   1245
            Y2              =   1545
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00404040&
            X1              =   5190
            X2              =   5190
            Y1              =   1245
            Y2              =   1545
         End
         Begin VB.Label LblMoneda 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblMoneda"
            Height          =   300
            Left            =   6075
            TabIndex        =   11
            Top             =   105
            Width           =   2025
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda"
            Height          =   165
            Index           =   3
            Left            =   4305
            TabIndex        =   10
            Top             =   150
            Width           =   750
         End
         Begin VB.Label Label1 
            Caption         =   "Direccion"
            Height          =   165
            Index           =   2
            Left            =   60
            TabIndex        =   7
            Top             =   810
            Width           =   750
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
            Height          =   165
            Index           =   1
            Left            =   75
            TabIndex        =   5
            Top             =   465
            Width           =   750
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   165
            Index           =   0
            Left            =   75
            TabIndex        =   3
            Top             =   150
            Width           =   750
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne3 
         Height          =   1560
         Left            =   90
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   90
         Width           =   11205
         _cx             =   19764
         _cy             =   2752
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
         BackColor       =   16777152
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   2
         ChildSpacing    =   4
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
         GridCols        =   2
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmPuntoVenta.frx":04B2
         Begin VB.Frame Frame8 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   1500
            Left            =   30
            TabIndex        =   50
            Top             =   30
            Width           =   4395
            Begin VB.CommandButton Command3 
               Caption         =   "Agregar Item"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   240
               TabIndex        =   53
               Top             =   90
               Width           =   1290
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Eliminar Item"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   2910
               TabIndex        =   52
               Top             =   90
               Width           =   1290
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Seleccionar Item"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   1575
               TabIndex        =   51
               Top             =   90
               Width           =   1290
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            Height          =   1500
            Left            =   4485
            TabIndex        =   39
            Top             =   30
            Width           =   6690
            Begin VB.TextBox Text12 
               Height          =   300
               Left            =   45
               TabIndex        =   44
               Text            =   "Text3"
               Top             =   345
               Width           =   1300
            End
            Begin VB.TextBox Text11 
               Height          =   300
               Left            =   1365
               TabIndex        =   43
               Text            =   "Text3"
               Top             =   345
               Width           =   1300
            End
            Begin VB.TextBox Text10 
               Height          =   300
               Left            =   2685
               TabIndex        =   42
               Text            =   "Text3"
               Top             =   345
               Width           =   1300
            End
            Begin VB.TextBox Text9 
               Height          =   300
               Left            =   4005
               TabIndex        =   41
               Text            =   "Text3"
               Top             =   345
               Width           =   1300
            End
            Begin VB.TextBox Text8 
               Height          =   300
               Left            =   5340
               TabIndex        =   40
               Text            =   "Text3"
               Top             =   345
               Width           =   1300
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Afecto"
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
               Left            =   45
               TabIndex        =   49
               Top             =   105
               Width           =   990
            End
            Begin VB.Label Label11 
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
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   1395
               TabIndex        =   48
               Top             =   105
               Width           =   1140
            End
            Begin VB.Label Label10 
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
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   2700
               TabIndex        =   47
               Top             =   105
               Width           =   510
            End
            Begin VB.Label Label9 
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
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   4020
               TabIndex        =   46
               Top             =   105
               Width           =   495
            End
            Begin VB.Label Label8 
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
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   5355
               TabIndex        =   45
               Top             =   105
               Width           =   450
            End
         End
      End
   End
End
Attribute VB_Name = "FrmPuntoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xIdTipDocVen As Integer
Dim xNumSerieActual As String


Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        Dim xform As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
    
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        Dim xCampos(4, 4) As String
        
        xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "4800":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Unid.":         xCampos(1, 1) = "abrev":          xCampos(1, 2) = "500":     xCampos(1, 3) = "C"
        xCampos(2, 0) = "Stock":        xCampos(2, 1) = "stckact":        xCampos(2, 2) = "800":     xCampos(2, 3) = "N"
        xCampos(3, 0) = "Código":       xCampos(3, 1) = "codpro":         xCampos(3, 2) = "2000":    xCampos(3, 3) = "C"
        
        
        '*******************************************************************************************
        Dim nSQLId As String
        
        xform.SQLCad = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, pvt_items.precio, alm_inventario.idunimed, mae_unidades.abrev, " _
            & " mae_moneda.simbolo, alm_inventario.stckact, alm_inventario.idcuentaven FROM mae_unidades RIGHT JOIN ((pvt_items LEFT JOIN alm_inventario ON pvt_items.iditem = alm_inventario.id) " _
            & " LEFT JOIN mae_moneda ON alm_inventario.idmon = mae_moneda.id) ON mae_unidades.id = alm_inventario.idunimed Where (((pvt_items.activo) = -1)) " _
            & " ORDER BY alm_inventario.descripcion"
        
        xform.Titulo = "Buscando Productos"
        
        xform.FormaBusca = Principio
        xform.Criterio = ""
        
        Dim RstCamBus As New ADODB.Recordset
        RST_Busq RstCamBus, "SELECT var_opcionesformulario.idform, var_opcionesformulario.campobus From var_opcionesformulario " _
            & " WHERE (((var_opcionesformulario.idform)=78))", xCon
        
        If RstCamBus.RecordCount <> 0 Then
            xform.Ordenado = RstCamBus("campobus")
            xform.CampoBusca = RstCamBus("campobus")
        Else
            xform.Ordenado = "codpro"
            xform.CampoBusca = "codpro"
        End If
        
        
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        Dim A As Integer
        If xRs.State = 1 Then
            If NulosN(xRs("idcuentaven")) = 0 Then
                MsgBox "El item seleccionado no tiene una cuenta contable asignada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Set xform = Nothing
                Set xRs = Nothing
                Exit Sub
            End If
            
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 1) = NulosC(xRs("descripcion"))
                Fg1.TextMatrix(Fg1.Row, 2) = NulosC(xRs("abrev"))
                Fg1.TextMatrix(Fg1.Row, 3) = ""
                Fg1.TextMatrix(Fg1.Row, 4) = Format(xRs("precio"), "0.000000")
                Fg1.TextMatrix(Fg1.Row, 6) = Format(xRs("precio"), "0.000000")
                
                Fg1.TextMatrix(Fg1.Row, 8) = xRs("id")
                Fg1.TextMatrix(Fg1.Row, 9) = NulosN(xRs("idunimed"))
                Fg1.TextMatrix(Fg1.Row, 10) = NulosN(xRs("stckact"))
                Fg1.TextMatrix(Fg1.Row, 11) = NulosN(xRs("idcuentaven"))
            End If
        End If
    End If
    
    If Fg1.TextMatrix(Fg1.Row, 8) <> "" Then
        Fg1.Rows = Fg1.Rows + 1
    End If
    
End Sub

Private Sub Form_Load()
    
    Fg1.ColWidth(8) = 0
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
    
    Frame1.BackColor = &H8000000F
    Frame5.BackColor = &H8000000F
    Frame6.BackColor = &H8000000F
    Frame9.BackColor = &H8000000F
    Frame10.BackColor = &H8000000F
    Frame11.BackColor = &H8000000F
    
    xNumSerieActual = "0001"
    RadioButton1.Value = True
    RadioButton1_Click
    Fg1.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
    Blanquea
    TxtFchEmi.Valor = Date
    TxtIdMoneda.Text = "1"
    LblMoneda.Caption = Busca_Codigo(TxtIdMoneda.Text, "id", "descripcion", "mae_moneda", "N", xCon)
    'TxtNumRuc.SetFocus
    Fg1.Editable = flexEDKbdMouse
    Fg1.ColComboList(1) = "|..."
End Sub

Function GenereNumDoc(IdTipDoc As Integer, NumSer As String) As String
    GenereNumDoc = "000-0000000000"
End Function

Sub Blanquea()
    TxtIdMoneda.Text = ""
    LblMoneda.Caption = ""
    TxtFchEmi.Valor = ""
    TxtNumRuc.Text = ""
    TxtNomCli.Text = ""
    TxtDirCli.Text = ""
    txtImpAfe.Text = ""
    TxtImpIna.Text = ""
    TxtImpIgv.Text = ""
    TxtImpIsc.Text = ""
    TxtImpTot.Text = ""
End Sub

Private Sub RadioButton1_Click()
    RadioButton1.Value = True
    LblDocumento.Caption = "BOLETA"
    xIdTipDocVen = 3
    LblNumDoc.Caption = GenereNumDoc(xIdTipDocVen, xNumSerieActual)
End Sub

Private Sub RadioButton2_Click()
    RadioButton2.Value = True
    LblDocumento.Caption = "FACTURA"
    xIdTipDocVen = 1
    LblNumDoc.Caption = GenereNumDoc(xIdTipDocVen, xNumSerieActual)
End Sub

Private Sub RadioButton3_Click()
    RadioButton3.Value = True
    LblDocumento.Caption = "COTIZACION"
    xIdTipDocVen = 107
    LblNumDoc.Caption = GenereNumDoc(xIdTipDocVen, xNumSerieActual)
End Sub
