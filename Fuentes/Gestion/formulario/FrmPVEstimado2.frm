VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPVEstimado2 
   Caption         =   "Sistena de ventas - Estimado de Ventas"
   ClientHeight    =   7860
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7485
      Left            =   0
      TabIndex        =   4
      Top             =   375
      Width           =   11880
      _cx             =   20955
      _cy             =   13203
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
         Height          =   7065
         Left            =   45
         TabIndex        =   8
         Top             =   375
         Width           =   11790
         Begin VB.Frame FrmPorcentaje 
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            Height          =   2145
            Left            =   7000
            TabIndex        =   27
            Top             =   4500
            Visible         =   0   'False
            Width           =   3780
            Begin VB.TextBox TxtPorcentaje 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1890
               TabIndex        =   30
               Text            =   "TxtPorcentaje"
               Top             =   390
               Width           =   1785
            End
            Begin VB.CommandButton Command2 
               Caption         =   "&Cancelar"
               Height          =   400
               Left            =   2010
               TabIndex        =   29
               Top             =   1680
               Width           =   1650
            End
            Begin VB.CommandButton Command1 
               Caption         =   "&Aplicar"
               Height          =   400
               Left            =   300
               TabIndex        =   28
               Top             =   1680
               Width           =   1650
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "X"
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
               Left            =   3540
               TabIndex        =   37
               Top             =   60
               Width           =   150
            End
            Begin VB.Label LblAño 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblAño"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   1890
               TabIndex        =   36
               Top             =   810
               Width           =   1770
            End
            Begin VB.Label LblProd 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblProd"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   1890
               TabIndex        =   35
               Top             =   1215
               Width           =   1770
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Año Seleccionado"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   34
               Top             =   810
               Width           =   1560
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Prod. Seleccionado"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   1215
               Width           =   1680
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Procesando Porcentaje"
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
               Left            =   105
               TabIndex        =   32
               Top             =   75
               Width           =   1995
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000005&
               BorderWidth     =   2
               Index           =   1
               X1              =   0
               X2              =   5610
               Y1              =   15
               Y2              =   15
            End
            Begin VB.Line Line2 
               BorderColor     =   &H80000003&
               BorderWidth     =   2
               X1              =   3750
               X2              =   3750
               Y1              =   0
               Y2              =   2100
            End
            Begin VB.Line Line3 
               BorderColor     =   &H80000005&
               BorderWidth     =   2
               X1              =   15
               X2              =   15
               Y1              =   0
               Y2              =   1035
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000003&
               BorderWidth     =   2
               Index           =   0
               X1              =   30
               X2              =   3750
               Y1              =   2130
               Y2              =   2130
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H80000002&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H80000002&
               Height          =   300
               Left            =   30
               Top             =   30
               Width           =   3690
            End
            Begin VB.Label LblProcesa 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ingrese Porcentaje"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Top             =   420
               Width           =   1620
            End
         End
         Begin VB.Frame frmCronog 
            Caption         =   "[ Historico de Ventas ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   2385
            Left            =   0
            TabIndex        =   25
            Top             =   4620
            Width           =   11800
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   2055
               Left            =   30
               TabIndex        =   26
               Top             =   240
               Width           =   11685
               _cx             =   20611
               _cy             =   3625
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
               Rows            =   1
               Cols            =   15
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmPVEstimado2.frx":0000
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
               FrozenCols      =   1
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2565
            Left            =   45
            TabIndex        =   3
            Top             =   1260
            Width           =   11745
            _cx             =   20717
            _cy             =   4524
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
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmPVEstimado2.frx":0197
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
         Begin VB.CommandButton CmdProcesar 
            Caption         =   "&Procesar"
            Height          =   400
            Left            =   6900
            TabIndex        =   24
            Top             =   420
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.CommandButton CmdAddProy 
            Caption         =   "Agregar Plan de Proyección"
            Height          =   400
            Left            =   9060
            TabIndex        =   23
            Top             =   420
            Visible         =   0   'False
            Width           =   2160
         End
         Begin VB.TextBox TxtDesc 
            Height          =   300
            Left            =   1155
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "TxtDesc"
            Top             =   390
            Width           =   5250
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
            Height          =   300
            Left            =   1155
            TabIndex        =   1
            Top             =   705
            Width           =   1365
            _ExtentX        =   2408
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
            Enabled         =   0   'False
            Valor           =   "06/02/2006"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
            Height          =   300
            Left            =   5070
            TabIndex        =   2
            Top             =   705
            Width           =   1365
            _ExtentX        =   2408
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
            Enabled         =   0   'False
            Valor           =   "06/02/2006"
         End
         Begin VB.Label LblNumItem 
            Alignment       =   1  'Right Justify
            Caption         =   "LblNumItem"
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
            Height          =   210
            Left            =   11085
            TabIndex        =   22
            Top             =   3930
            Width           =   675
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nº Productos Procesados  :"
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
            Left            =   8640
            TabIndex        =   21
            Top             =   3930
            Width           =   2370
         End
         Begin VB.Label Label9 
            Caption         =   "Unidad Medida"
            Height          =   255
            Left            =   4725
            TabIndex        =   20
            Top             =   3900
            Width           =   1200
         End
         Begin VB.Label LblUniMed 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblUniMed"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   6000
            TabIndex        =   19
            Top             =   3885
            Width           =   1125
         End
         Begin VB.Label LblCodigo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblCodigo"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1080
            TabIndex        =   18
            Top             =   3930
            Width           =   2160
         End
         Begin VB.Label Label7 
            Caption         =   "Código"
            Height          =   255
            Left            =   60
            TabIndex        =   17
            Top             =   3960
            Width           =   1005
         End
         Begin VB.Label LblDesc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDesc"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1080
            TabIndex        =   16
            Top             =   4260
            Width           =   10620
         End
         Begin VB.Label Label5 
            Caption         =   "Descripción"
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   4290
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Inicio"
            Height          =   195
            Left            =   60
            TabIndex        =   13
            Top             =   735
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   60
            TabIndex        =   12
            Top             =   420
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle Proyección de Ventas"
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
            TabIndex        =   11
            Top             =   45
            Width           =   11610
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Productos "
            Height          =   195
            Left            =   60
            TabIndex        =   10
            Top             =   1035
            Width           =   765
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Término"
            Height          =   195
            Left            =   3900
            TabIndex        =   9
            Top             =   735
            Width           =   930
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7065
         Left            =   -12435
         TabIndex        =   5
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6570
            Left            =   30
            TabIndex        =   6
            Top             =   375
            Width           =   11790
            _ExtentX        =   20796
            _ExtentY        =   11589
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
            Columns(1).Caption=   "Nº Proyecto"
            Columns(1).DataField=   "id"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripción"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Ini"
            Columns(3).DataField=   "fchini"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Fin"
            Columns(4).DataField=   "fchfin"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Estado"
            Columns(5).DataField=   "estado"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2381"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2302"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=8202"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=8123"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1826"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1746"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1799"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1720"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1667"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1588"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HDBFDFD&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Consulta Proyección de Ventas"
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
            Left            =   120
            TabIndex        =   7
            Top             =   45
            Width           =   11610
         End
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado2.frx":0383
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado2.frx":08C7
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado2.frx":0A4B
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado2.frx":0E9F
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado2.frx":0FB7
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado2.frx":14FB
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado2.frx":1A3F
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado2.frx":1B53
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado2.frx":1C67
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado2.frx":20BB
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado2.frx":2227
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Proyeccion de Ventas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar Proyeccion de Ventas"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar Proyeccion de Ventas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Desactivar Proyeccion de Ventas"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar Promedio - Mes Actual"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "Agregar Promedio - Todos los meses"
      End
      Begin VB.Menu Menu1_5 
         Caption         =   "Agregar año seleccionado y aplicar porcentaje"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_6 
         Caption         =   "Exportar a Excel"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "Agregar Producto"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "Eliminar Producto"
      End
      Begin VB.Menu Menu2_4 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_5 
         Caption         =   "Ver Historico de Ventas"
      End
   End
End
Attribute VB_Name = "FrmPVEstimado2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMPVESTIMADO
'* Tipo             : FORMULARIO
'* Descripcion      : PERMITE GENERAR UNA PROYECCION DE LAS VENTAS, EN FUNCION A LOS DATOS HISTORICOS
'*                    DE VENTA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstPlanes As New ADODB.Recordset   ' RECORSET QUE ALMACENA LOS REGISTRO DE LA TABLA ges_ventaproy
Dim QueHace As Integer                 ' INDICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim SeEjecuto As Boolean               ' INDICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim xTitulo As String                  ' ALAMCENA EL TITULO DEL FORMULARIO
Dim xFilaActual As Integer             ' INDICA LA FILA ACTUAL PARA EL CONTROL FLEXGRID
Dim Agregando  As Boolean              ' VARIABLE QUE INDICA QUE SE ESTA AGREGANDO UNA FILA AL CONTROL FLEXGRID

Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO



'*****************************************************************************************************
'* Nombre Archivo   : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TEXTBOX DEL FORMULARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Bloquea()
    If QueHace = 1 Then CmdProcesar.Visible = True: CmdAddProy.Visible = True
    If QueHace = 2 Then CmdProcesar.Visible = True
    If QueHace = 3 Then CmdProcesar.Visible = False: CmdAddProy.Visible = False
    If QueHace = 3 Then FrmPorcentaje.Visible = False
    TxtDesc.Locked = Not TxtDesc.Locked
    TxtFchIni.Enabled = Not TxtFchIni.Enabled
    TxtFchFin.Enabled = Not TxtFchFin.Enabled
    'TxtPorcentaje.Locked = Not TxtPorcentaje.Locked
    Fg1.Rows = 1
    Fg2.Rows = 1
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UNA REGISTRO DE LA TABLA ges_ventaproy
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    TabOne1.CurrTab = 0
    If RstPlanes.RecordCount = 0 Then
        MsgBox "No hay registros para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de eliminar la proyección de ventas seleccionada?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM ges_ventaproy WHERE id = " & RstPlanes("id") & ""
        xCon.Execute "DELETE * FROM ges_ventaproydet2 WHERE id = " & RstPlanes("id") & ""
        xCon.Execute "DELETE * FROM ges_ventaproydet WHERE id = " & RstPlanes("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstPlanes("id") & " AND idform = " & IdMenuActivo
        
        
        MsgBox "La proyección de ventas se elimino con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstPlanes.Requery
        Dg1.Refresh
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LOS CONTROLES TEXTOBOX PARA EL INGRESO DE UN NUEVO REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Blanquea()
    TxtDesc.Text = ""
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    TxtPorcentaje.Text = ""
    LblUniMed.Caption = ""
    LblCodigo.Caption = ""
    LblDesc.Caption = ""
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA ges_ventaproy, ESTA FUNCION DEVUELVE VERDADERO
'*                    CUANDO TIENE EXITO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    If TxtDesc.Text = "" Then
        MsgBox "No ha especificado la descripcion del plan", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtDesc.SetFocus
        Exit Function
    End If
    
    If TxtFchIni.Valor = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    End If

    If TxtFchFin.Valor = "" Then
        MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Function
    End If

    Dim A As Integer
    
    'eliminar las filas que esten vacias
    For A = 1 To Fg1.Rows
        If Fg1.TextMatrix(A, 1) = "" Then
            Fg1.RemoveItem (A)
            A = A - 1
        End If
        
        If A = Fg1.Rows - 1 Then
            Exit For
        End If
    Next A

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId As Double
    Dim xFila As Integer
    Dim xCol As Integer
    
    On Error GoTo LaCague

    xCon.BeginTrans
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT * FROM ges_ventaproy", xCon
        RST_Busq RstDet, "SELECT * FROM ges_ventaproydet2", xCon
        
        xId = HallaCodigoTabla("ges_ventaproy", xCon, "id")
        RstCab.AddNew
        
        RstCab("id") = xId
    Else
        xId = RstPlanes("id")
        
        RST_Busq RstCab, "SELECT * FROM ges_ventaproy WHERE id = " & xId & "", xCon
        xCon.Execute "DELETE * FROM ges_ventaproydet2 WHERE id = " & xId & ""
        RST_Busq RstDet, "SELECT * FROM ges_ventaproydet2", xCon
        
        
    End If
    
    RstCab("descripcion") = NulosC(TxtDesc.Text)
    RstCab("fchini") = TxtFchIni.Valor
    RstCab("fchfin") = TxtFchFin.Valor
    RstCab.Update
    Dim idMesIni As Integer
    Dim xMes As Integer
    idMesIni = CInt(Mid(TxtFchIni.Valor, 4, 2))
    
    For A = 1 To Fg1.Rows - 1
        xMes = idMesIni
        For xCol = 6 To Fg1.Cols - 1
            RstDet.AddNew
            RstDet("id") = xId
            RstDet("idpro") = NulosN(Fg1.TextMatrix(A, 1))
            RstDet("codigo") = NulosC(Fg1.TextMatrix(A, 3))
            RstDet("idmes") = xMes
            RstDet("cantidad") = NulosN(Fg1.TextMatrix(A, xCol))
            xMes = xMes + 1
            If xMes > 12 Then xMes = 1
            RstDet.Update
        Next xCol
    Next A
    
    '-------------------
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    '-------------------
    
    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    
    MsgBox "El plan proyectado de ventas se guardo con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    
    Exit Function

LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

'*****************************************************************************************************
'* Nombre Archivo   : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE INGRESAR O MODIFICAR UN REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Bloquea
    Toolbar
    Label1.Caption = "Detalle Proyeccion de Ventas"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.ColComboList(1) = ""
    Fg1.Editable = flexEDNone
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.BackColorSel = &H80&
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    Label1.Caption = "Agregando Proyeccion de Ventas"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Toolbar
    Bloquea
    Blanquea
    Fg1.ColComboList(1) = "|..."
    'Fg1.Rows = Fg1.Rows + 1
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDNone
    
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
        
    TxtDesc.SetFocus
    
    Fg1.SelectionMode = flexSelectionFree
    Fg1.BackColorSel = &H80&
    LblNumItem.Caption = 0
    
    configurarVista
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA LA EDICION DE UN REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    xHorIni = Time
    Label1.Caption = "Modificando Proyeccion de Ventas"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Toolbar
    Bloquea
    'Blanquea
    Fg1.ColComboList(1) = "|..."
    Fg1.Rows = Fg1.Rows + 1
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDNone
    MuestraSegundoTab
    TxtDesc.SetFocus

    Fg1.SelectionMode = flexSelectionFree
    Fg1.BackColorSel = &H80&
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Toolbar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Toolbar()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
    
    Toolbar1.Buttons(14).Enabled = Not Toolbar1.Buttons(14).Enabled
End Sub

Private Function encontrarAniosTrabajo(ByRef rut() As String) As Boolean
    Dim xConAux As New ADODB.Connection
    Dim xFun As New eps_librerias.FuncionesData
    Dim Rst As New ADODB.Recordset
    Dim NumRUC As String
    Dim xCad As String
    Dim cSQL As String
    Dim rutas() As String
    Dim cant As Integer
    Dim A As Integer
    
    cSQL = "SELECT numruc FROM mae_empresa"
    RST_Busq Rst, cSQL, xCon
    
    NumRUC = Rst("numruc")
    xCad = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    
    xFun.F_BASEDATOS = xCad + "data.mdb"
    xFun.F_GRUPOTRABAJO = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS") + "seven.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    
    Set xConAux = xFun.AbrirConeccion
    
    cSQL = "SELECT mae_empresa.numruc, mae_empresa.ruta, mae_empresa.anotra, mae_empresa.activo " _
        + vbCr + "From mae_empresa " _
        + vbCr + "WHERE (((mae_empresa.numruc)= '" & NumRUC & "') AND ((mae_empresa.activo)=-1))"
    
    RST_Busq Rst, cSQL, xConAux
    If Rst.RecordCount <> 0 Then
        cant = Rst.RecordCount
        ReDim rutas(cant, 2) As String
        Rst.MoveFirst
        For A = 0 To cant - 1
            rutas(A, 1) = Rst("ruta")
            rutas(A, 0) = Rst("anotra")
            Rst.MoveNext
        Next A
        rut = rutas
        encontrarAniosTrabajo = True
    Else
        encontrarAniosTrabajo = False
    End If
End Function

Private Sub CmdAddProy_Click()
    'Se Busca un plan de Produccion
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCodItem As String
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "id":            xCampos(1, 2) = "2000":    xCampos(1, 3) = "N"
        
    xform.SQLCad = "SELECT ges_ventaproy.id, ges_ventaproy.descripcion, ges_ventaproy.fchini, ges_ventaproy.fchfin FROM ges_ventaproy;"
    
    xform.Titulo = "Buscando Proyeccion de Ventas"
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        Dim xId As Integer
        xId = xRs("id")
        Set xform = Nothing
        Set xRs = Nothing
    
        MostrarDetalleProyVtas xId
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Sub MostrarDetalleProyVtas(xIdProyVtas As Integer)
    Dim Rst As New ADODB.Recordset
    Dim RstPlaProy As New ADODB.Recordset
    Dim Rst2Aux As New ADODB.Recordset
    Dim A, B, xCol As Integer
    Dim Total As Double
    
    Dim idMesIni As Integer
    Dim idMesFin As Integer
    Dim idAñoIni As Integer
    Dim idAñoFin As Integer
    
    Dim xMes As Integer
    Dim xAño As Integer
    Dim bandera As Boolean
    Dim contador As Integer
    
    Dim xMesAux As String
    
    Fg1.Rows = 1
    Fg1.Cols = 3
    TxtFchIni.Valor = Format("28/12/11", "dd/mm/yyyy")
    TxtFchFin.Valor = Format("28/05/12", "dd/mm/yyyy")
    
    idMesIni = CInt(Mid(TxtFchIni.Valor, 4, 2))
    idMesFin = CInt(Mid(TxtFchFin.Valor, 4, 2))
    idAñoIni = CInt(Mid(TxtFchIni.Valor, 7, 4))
    idAñoFin = CInt(Mid(TxtFchFin.Valor, 7, 4))
'
    xMes = idMesIni
    xAño = idAñoIni
    
    Dim indicador As Integer
    indicador = (13 - idMesIni) + idMesFin
    If indicador > 12 Then indicador = 12
    
    Fg1.Cols = Fg1.Cols + indicador + 4
    
    Dim xCad As String
    
    xCad = "SELECT ges_ventaproy.id, ges_ventaproy.descripcion, ges_ventaproy.fchini, ges_ventaproy.fchfin " _
        + vbCr + "From ges_ventaproy " _
        + vbCr + "WHERE (((ges_ventaproy.id)=" & xIdProyVtas & "))"
    RST_Busq RstPlaProy, xCad, xCon
    
    TxtDesc.Text = RstPlaProy("descripcion")
    TxtFchIni.Valor = RstPlaProy("fchini")
    TxtFchFin.Valor = RstPlaProy("fchfin")
    
    xCad = "TRANSFORM First(ges_ventaProydet2.cantidad) AS PrimeroDecantidad " _
        + vbCr + "SELECT ges_ventaProydet2.idpro, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev, Sum(ges_ventaProydet2.cantidad) AS [Total de cantidad] " _
        + vbCr + "FROM ((ges_ventaProydet2 RIGHT JOIN ges_ventaproy ON ges_ventaProydet2.id = ges_ventaproy.id) LEFT JOIN alm_inventario ON ges_ventaProydet2.idpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        + vbCr + "Where (((ges_ventaProydet2.id) = " & xIdProyVtas & ")) " _
        + vbCr + "GROUP BY ges_ventaProydet2.idpro, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
        + vbCr + "PIVOT ges_ventaProydet2.idmes;"

    RST_Busq Rst, xCad, xCon
    
    Set Fg1.DataSource = Rst.DataSource
    
    configurarGrid Fg1, TxtFchIni.Valor, TxtFchFin.Valor
    Fg1.Editable = flexEDKbdMouse
End Sub

Private Sub mostrarHistoricoVtas()
    Dim RstRutas As New ADODB.Recordset
    Dim RstHisVenta As New ADODB.Recordset
    Dim A, B As Integer
    Dim NumAños As Integer
    Dim xTotal As Double
    Dim cont As Integer
    Dim var As Double
    Dim media As Double
    Dim X As Double
    Dim desv As Double
    Dim rutas() As String
    Dim indicador As Integer
    Dim xMes As Integer
    Dim xMesIni As Integer
    
    indicador = calcularIndicador(NulosC(TxtFchIni.Valor), NulosC(TxtFchFin.Valor))
    xMesIni = NulosN(Format(NulosC(TxtFchIni.Valor), "m"))
    configurarGrid2 Fg2, TxtFchIni.Valor, TxtFchFin.Valor
    Fg2.Rows = 1
    If encontrarAniosTrabajo(rutas) Then
        NumAños = UBound(rutas, 1)
        For A = 0 To UBound(rutas, 1) - 1
            Set RstHisVenta = MostrarAños(rutas(A, 0), NulosN(Fg1.TextMatrix(Fg1.Row, 1)), rutas(A, 1))
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = rutas(A, 0)
            If RstHisVenta.RecordCount <> 0 Then
                xMes = xMesIni
                For B = 2 To indicador + 1
                    Fg2.TextMatrix(Fg2.Rows - 1, B) = Format(NulosN(RstHisVenta("" & xMes & "")), "0.00")
                    xMes = xMes + 1
                    If xMes > 12 Then xMes = 1
                Next B
            End If
        Next A
    End If
    
    Fg2.Rows = Fg2.Rows + 3
    Fg2.TextMatrix(Fg2.Rows - 2, 1) = "Total ==>"
    Fg2.TextMatrix(Fg2.Rows - 1, 1) = "Promedio"
    
    For A = 2 To indicador + 1
        xTotal = 0
        cont = 0
        For B = 1 To Fg2.Rows - 2
            xTotal = NulosN(Fg2.TextMatrix(B, A)) + xTotal
            If Fg2.TextMatrix(B, A) <> "" Then cont = cont + 1
            If B = Fg2.Rows - 2 Then
                Exit For
            End If
        Next B
        Fg2.TextMatrix(Fg2.Rows - 2, A) = Format(xTotal, "0.00")
        If cont = 0 Then cont = 1
        Fg2.TextMatrix(Fg2.Rows - 1, A) = (Val(Fg2.TextMatrix(Fg2.Rows - 2, A)) / cont)
        Fg2.TextMatrix(Fg2.Rows - 1, A) = Format(Fg2.TextMatrix(Fg2.Rows - 1, A), "0.00")
    Next A
        
    Fg2.Rows = Fg2.Rows + 1
    Fg2.TextMatrix(Fg2.Rows - 1, 1) = "Desv.Estand"
    
    For A = 2 To indicador + 1
        cont = 0
        var = 0
        desv = 0
        For B = 1 To Fg2.Rows - 5
            If Fg2.TextMatrix(B, A) <> "" Then
                X = Fg2.TextMatrix(B, A)
                media = Fg2.TextMatrix(Fg2.Rows - 2, A)
                var = var + ((X - media) * (X - media))
                cont = cont + 1
            End If
            If B = Fg2.Rows - 3 Then
                Exit For
            End If
        Next B
        If cont = 0 Then cont = 1
        var = NulosN(var) / cont
        desv = Sqr(var)
        Fg2.TextMatrix(Fg2.Rows - 1, A) = Format(desv, "0.00")
    Next A
    
    With Fg2
        .Select 1, 1, Fg2.Rows - 1, 1
        .FillStyle = flexFillRepeat
        .CellBackColor = &HDDFFFF
        .Select Fg2.Rows - 3, 1, Fg2.Rows - 2, indicador + 1
        .FillStyle = flexFillRepeat
        .CellBackColor = &HEBD7BC
        .Select Fg2.Rows - 1, 1, Fg2.Rows - 1, indicador + 1
        .FillStyle = flexFillRepeat
        .CellBackColor = &H80000003
        .Select Fg2.Rows - 1, 2, Fg2.Rows - 1, 2
    End With
    
    
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : MostrarAños
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LA INFORMACION DE VENTAS DE TODOS LOS AÑOS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Año          |  String    |  ESPECIFICA EL AÑO DE TRABAJO
'*                    CodProducto  |  Integer   |  ESPECIFICA EL ID DEL PRODUCTO
'*                    RutaData     |  String    |  ESPECIFICA LA RUTA DE LA BASE DE DATOS
'* DEVUELVE         :
'*****************************************************************************************************
Function MostrarAños(Año As String, CodProducto As Integer, RutaData As String) As ADODB.Recordset
    Dim RstAño As New ADODB.Recordset
    Dim xCad As String
    
    Dim xFun As New eps_librerias.FuncionesData
    Dim xRutaData As String
    Dim xRst As New ADODB.Recordset
    Dim xCon2 As New ADODB.Connection
    Dim cSQL As String
    
    xCad = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    
    xFun.F_BASEDATOS = xCad + RutaData
    xFun.F_GRUPOTRABAJO = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS") + "seven.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    
    Set xCon2 = xFun.AbrirConeccion

    cSQL = "TRANSFORM Sum(vta_ventasdet.canpro) AS SumaDecanpro " _
        + vbCr + "SELECT vta_ventasdet.iditem, alm_inventario.descripcion, Sum(vta_ventasdet.canpro) AS total " _
        + vbCr + "FROM vta_ventas INNER JOIN (vta_ventasdet INNER JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta " _
        + vbCr + "Where (((vta_ventasdet.iditem) = " & CodProducto & ")) " _
        + vbCr + "GROUP BY vta_ventasdet.iditem, alm_inventario.descripcion " _
        + vbCr + "PIVOT Format([fchdoc],'m') In ('1','2','3','4','5','6','7','8','9','10','11','12');"
        
    RST_Busq RstAño, cSQL, xCon2

    Set MostrarAños = RstAño
End Function

'*****************************************************************************************************
'* Nombre           : calcularIndicador
'* Tipo             : Function
'* Descripcion      : Calcula el indicador de numero de meses a procesar
'* Creado por       : JOSE CHACON MANRIQUE
'* Modificado       :
'*****************************************************************************************************
Private Function calcularIndicador(fchIni As String, fchFin As String) As Integer
    Dim indicador As Integer
    Dim idMesIni As Integer
    Dim idMesFin As Integer
    Dim idAñoIni As Integer
    Dim idAñoFin As Integer
    
    idMesIni = NulosN(Format(fchIni, "m"))
    idMesFin = NulosN(Format(fchFin, "m"))
    idAñoIni = NulosN(Format(fchIni, "yyyy"))
    idAñoFin = NulosN(Format(fchFin, "yyyy"))
    
    If idMesIni <> 0 And idAñoIni <> 0 Then
        If idAñoFin > idAñoIni Then
            indicador = (13 - idMesIni) + idMesFin
        Else
            indicador = idMesFin - idMesIni + 1
        End If
        
        If indicador > 12 Then indicador = 12
    End If
    
    calcularIndicador = indicador
End Function

Private Sub configurarGrid2(fgx As VSFlexGrid, fchIni As String, fchFin As String)
    Dim Rst As New ADODB.Recordset
    Dim idMesIni As Integer
    Dim idAñoIni As Integer
    Dim A As Integer
    Dim xMes As Integer
    Dim xAño As Integer
    Dim indicador As Integer
    
    xMes = NulosN(Format(fchIni, "m"))
    
    xAño = NulosN(Format(fchIni, "yyyy"))
    indicador = calcularIndicador(NulosC(fchIni), NulosC(fchFin))
    
    If xMes <> 0 And indicador <> 0 Then
        fgx.Cols = 2 + indicador
        
        fgx.ColWidth(0) = 0
        fgx.TextMatrix(0, 1) = "Detalle"
        fgx.ColWidth(1) = 1500
        
        If fgx.Rows = 1 Then fgx.Rows = fgx.Rows + 1
        fgx.Select 1, 1, 1, 1
        fgx.FrozenCols = 1
        
        For A = 1 To indicador
            RST_Busq Rst, "SELECT DISTINCT con_meses.id, con_meses.descripcion " _
                        & "FROM con_meses " _
                        & "WHERE (((con_meses.id)=" & xMes & "))", xCon
            
            fgx.TextMatrix(0, A + 1) = Rst("descripcion")
            fgx.ColWidth(A + 1) = 1250
            xMes = xMes + 1
            If xMes > 12 Then xMes = 1
        Next A
        
        Set Rst = Nothing
    Else
        MsgBox "Las Fechas a procesar no son adecuadas"
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : configurarGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : Configura los Detalles del VsFlexGrid
'* Creado por       : JOSE CHACON MANRIQUE
'* Modificado       :
'*****************************************************************************************************
Sub configurarGrid(fgx As VSFlexGrid, fchIni As String, fchFin As String)
    Dim Rst As New ADODB.Recordset
    Dim idMesIni As Integer
    Dim idAñoIni As Integer
    Dim A As Integer
    Dim xMes As Integer
    Dim xAño As Integer
    Dim indicador As Integer
    
    xMes = NulosN(Format(fchIni, "m"))
    
    xAño = NulosN(Format(fchIni, "yyyy"))
    indicador = calcularIndicador(NulosC(fchIni), NulosC(fchFin))
    
    If xMes <> 0 And indicador <> 0 Then
        fgx.Cols = 6 + indicador
        
        fgx.ColWidth(0) = 0
        fgx.TextMatrix(0, 1) = "Id"
        fgx.ColWidth(1) = 0
        fgx.TextMatrix(0, 2) = "Producto"
        fgx.ColWidth(2) = 4500
        fgx.ColAlignment(2) = flexAlignLeftCenter
        fgx.TextMatrix(0, 3) = "Codigo"
        fgx.ColWidth(3) = 0
        fgx.TextMatrix(0, 4) = "Unidad"
        fgx.TextMatrix(0, 5) = "Programado"
        If QueHace = 3 Then fgx.ColWidth(5) = 1250 Else fgx.ColWidth(5) = 0
        
        If fgx.Rows = 1 Then fgx.Rows = fgx.Rows + 1
        fgx.Select 1, 1, 1, 1
        fgx.FrozenCols = 5
        
        For A = 1 To indicador
            RST_Busq Rst, "SELECT DISTINCT con_meses.id, con_meses.descripcion " _
                        & "FROM con_meses " _
                        & "WHERE (((con_meses.id)=" & xMes & "))", xCon
            
            fgx.TextMatrix(0, A + 5) = NulosC(Rst("descripcion")) & " " & xAño
            fgx.ColWidth(A + 5) = 1250
            xMes = xMes + 1
            If xMes > 12 Then xMes = 1: xAño = xAño + 1
        Next A
        
        With fgx
            .Select 1, 5, fgx.Rows - 1, 5
            .FillStyle = flexFillRepeat
            .CellBackColor = &HE0FEE7
            .Select 1, 1, 1, 1
        End With
        
        Set Rst = Nothing
        GRID_COMBOLIST fgx, 2
    Else
        MsgBox "Las Fechas a procesar no son adecuadas"
    End If
End Sub

Private Sub CmdProcesar_Click()
On Error GoTo TuMay
    If (TxtFchIni.Valor = "" Or TxtFchFin.Valor = "" Or CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor)) Then MsgBox "Ingrese correctamente las Fechas a Procesar": Exit Sub
    configurarGrid Fg1, TxtFchIni.Valor, TxtFchFin.Valor
    mostrarHistoricoVtas
    Exit Sub
TuMay:
    MsgBox "Ingrese correctamente las Fechas a Procesar"
End Sub

Private Sub Command1_Click()
    CompiarValoresAplicarPorcentaje
    FrmPorcentaje.Visible = False
End Sub

Private Sub Command2_Click()
    FrmPorcentaje.Visible = False
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstPlanes("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    ' EJECUTA LA BUSQUEDA DE UN PRODUCTO
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    If Col = 2 Then
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Producto":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5700":     xCampos(0, 3) = "C"
        xCampos(1, 0) = "Unidad":     xCampos(1, 1) = "abrev":         xCampos(1, 2) = "800":      xCampos(1, 3) = "C"
        xCampos(2, 0) = "Codigo":     xCampos(2, 1) = "codpro":        xCampos(2, 2) = "1700":     xCampos(2, 3) = "C"
        
        xform.SQLCad = "SELECT alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev, alm_inventario.idunimed, alm_inventario.id " _
            & " FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) = 3)) " _
            & " ORDER BY alm_inventario.descripcion"
        
        xform.Titulo = "Buscando Productos"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If BuscaItemGrid(xRs("codpro")) = False Then
                LblCodigo.Caption = xRs("codpro")
                LblDesc.Caption = xRs("descripcion")
                LblUniMed.Caption = Busca_Codigo(xRs("idunimed"), "id", "descripcion", "mae_unidades", "N", xCon)
                
                Fg1.TextMatrix(Fg1.Row, 1) = xRs("id")
                Fg1.TextMatrix(Fg1.Row, 2) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 3) = xRs("codpro")
                Fg1.TextMatrix(Fg1.Row, 4) = xRs("abrev")
                'PREGUNTAMOS SI LA ULTIMA FILA ESTA VACIA PARA AGREGARLE OTRO ITEM
                If Fg1.TextMatrix(Fg1.Rows - 1, 1) <> "" Then
                    Fg1.Rows = Fg1.Rows + 1
                End If
                Fg1.Select Fg1.Row, 1, Fg1.Row, 1
                Fg1_RowColChange
            End If
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
    Else
        If Fg1.Col > 2 Then
            Fg1.Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
        
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace = 3 Then Exit Sub
        PopupMenu menu2
    End If
End Sub

Private Sub configurarVista()
    If QueHace = 3 Then
        frmCronog.Visible = False
        
        Fg1.Height = 4850
        
        Label5.Top = 6600
        Label7.Top = 6210
        Label9.Top = 6210
        Label8.Top = 6210
        LblDesc.Top = 6560
        LblCodigo.Top = 6180
        LblUniMed.Top = 6180
        LblNumItem.Top = 6180
        
        Fg1.AutoSearch = flexSearchFromTop
        Fg1.ExplorerBar = flexExSortShowAndMove
        Fg2.AutoSearch = flexSearchFromTop
        Fg2.ExplorerBar = flexExSortShowAndMove
            
    Else
        frmCronog.Visible = True
        frmCronog.Top = 4250
        frmCronog.Left = 0
        frmCronog.Height = 2800
        frmCronog.Width = 11800
        
        Fg1.Height = 2350
        Fg2.Height = 2500
        
        Label5.Top = 3950
        Label7.Top = 3650
        Label9.Top = 3650
        Label8.Top = 3680
        LblDesc.Top = 3935
        LblCodigo.Top = 3635
        LblUniMed.Top = 3635
        LblNumItem.Top = 3680
        
        'Fg2.AllowUserResizing = flexResizeColumns
        Fg1.AutoSearch = flexSearchNone
        Fg1.ExplorerBar = flexExNone
        Fg2.AutoSearch = flexSearchNone
        Fg2.ExplorerBar = flexExNone
                
        
    End If
End Sub

Private Sub Fg1_RowColChange()
    If Agregando = True Then Exit Sub
    
    If Fg1.Rows = 1 Then Exit Sub
    
    If LblDesc.Caption <> Fg1.TextMatrix(Fg1.Row, 2) Then
        Fg2.Cols = 1
    End If
    
    LblDesc.Caption = Fg1.TextMatrix(Fg1.Row, 2)
    LblCodigo.Caption = Fg1.TextMatrix(Fg1.Row, 1)
    LblUniMed.Caption = Fg1.TextMatrix(Fg1.Row, 4)

End Sub

Private Sub Fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    If Button = 2 Then
        PopupMenu menu1
    End If
End Sub

Private Sub iniciarCampos()
    Fg1.AllowUserResizing = flexResizeColumns
    Fg1.AutoSearch = flexSearchFromTop
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ExplorerBar = flexExSortShow
    Fg1.ForeColorSel = &H80000005
    Fg1.BackColorSel = &H80&
    Fg1.Editable = flexEDNone
    Fg1.MergeCells = flexMergeSpill
    Fg1.BackColorSel = &H80&
    
    Fg2.AllowUserResizing = flexResizeColumns
    Fg2.AutoSearch = flexSearchFromTop
    Fg2.ExplorerBar = flexExSortShowAndMove
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.ForeColorSel = &H80000005
    Fg2.BackColorSel = &H80&
    
    
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Fg1.ColWidth(2) = 0
    Fg1.FrozenCols = 2
End Sub

Private Sub Form_Activate()
'Modificado: 08/01/11 Johan Castro
'            Agregar linea de codigo para bloquear accesos de usuarios

    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        '----------------------------------------------
        TxtFchIni.Valor = Date
        TxtFchFin.Valor = Date
        
        Dim Rpta As Integer
        RST_Busq RstPlanes, "SELECT ges_ventaproy.*, IIf([ges_ventaproy].[activo]=-1,'Activo','No Activo') AS estado " _
            & " FROM ges_ventaproy ORDER BY id DESC", xCon

        Set Dg1.DataSource = RstPlanes
        
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    TabOne1.CurrTab = 0
    QueHace = 3
    iniciarCampos
    SeEjecuto = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub Label15_Click()
    FrmPorcentaje.Visible = False
End Sub

Private Sub menu1_1_Click()
    If Fg1.TextMatrix(Fg1.Row, 2) = "" Then
        MsgBox "No ha seleccionado ningun Producto seleccione un Producto para aplicar esta opcion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If Fg2.TextMatrix(Fg2.Row, 1) = "" Or Fg2.TextMatrix(Fg2.Row, 1) = "Total ==>" Or Fg2.TextMatrix(Fg2.Row, 1) = "Promedio" Or Fg2.TextMatrix(Fg2.Row, 1) = "Desv.Estand" Then
        MsgBox "No ha seleccionado una fila valida del historico de ventas, " & Chr(13) _
            & "seleccione una fila valida para aplicar esta opcion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    CopiarValores
    'HallarTotales
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : CopiarValores
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : COPIA LOS VALORES DE UNA CELDA DEL CONTRO FLEXGRID Fg1, Fg2
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub CopiarValores()
    Fg1.TextMatrix(Fg1.Row, Fg2.Col + 4) = Val(Fg2.TextMatrix(Fg2.Rows - 2, Fg2.Col)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, Fg2.Col)) / 100) + 1)
    Fg1.TextMatrix(Fg1.Row, Fg2.Col + 4) = Format(Fg1.TextMatrix(Fg1.Row, Fg2.Col + 4), "0.00")
End Sub

Private Sub menu1_2_Click()
    Dim A As Integer
    Dim indicador As Integer
    
    If Fg1.TextMatrix(Fg1.Row, 2) = "" Then
        MsgBox "No ha seleccionado ningun Producto seleccione un Producto para aplicar esta opcion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If Fg2.TextMatrix(Fg2.Row, 1) = "" Or Fg2.TextMatrix(Fg2.Row, 1) = "Total ==>" Or Fg2.TextMatrix(Fg2.Row, 1) = "Promedio" Or Fg2.TextMatrix(Fg2.Row, 1) = "Desv.Estand" Then
        MsgBox "No ha seleccionado una fila valida del historico de ventas, " & Chr(13) _
            & "seleccione una fila valida para aplicar esta opcion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    indicador = calcularIndicador(TxtFchIni.Valor, TxtFchFin.Valor)
    For A = 1 To indicador
        Fg1.TextMatrix(Fg1.Row, A + 5) = Val(Fg2.TextMatrix(Fg2.Rows - 2, A + 1)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, A + 1)) / 100) + 1)
    Next A
End Sub

Private Sub Menu1_5_Click()
    If Fg1.TextMatrix(Fg1.Row, 2) = "" Then
        MsgBox "No ha seleccionado ningun Producto seleccione un Producto para aplicar esta opcion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If Fg2.TextMatrix(Fg2.Row, 1) = "" Or Fg2.TextMatrix(Fg2.Row, 1) = "Total ==>" Or Fg2.TextMatrix(Fg2.Row, 1) = "Promedio" Or Fg2.TextMatrix(Fg2.Row, 1) = "Desv.Estand" Then
        MsgBox "No ha seleccionado una fila valida del historico de ventas, " & Chr(13) _
            & "seleccione una fila valida para aplicar esta opcion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If QueHace <> 3 Then FrmPorcentaje.Visible = True Else FrmPorcentaje.Visible = False
    LblProd = Fg1.TextMatrix(Fg1.Row, 2)
    LblProd.ToolTipText = LblProd.Caption
    LblAño = Fg2.TextMatrix(Fg2.Row, 1)
    TxtPorcentaje.SetFocus
End Sub

Private Sub Menu1_6_Click()
    ExportarExcel
End Sub

Private Sub Menu2_1_Click()
    If Fg1.TextMatrix(Fg1.Rows - 1, 1) = "" Then Exit Sub
    
    Fg1.Rows = Fg1.Rows + 1
    Fg1.SetFocus
    Fg1.Select Fg1.Rows - 1, 1
    Fg1_CellButtonClick Fg1.Rows - 1, 1
End Sub

Private Sub Menu2_3_Click()
    ' ELIMINA UNA FILA DEL CONTROL Fg1
    Dim Rpta As Integer
    If Fg1.Row > 0 Then
        Rpta = MsgBox("¿Esta seguro de eliminar el producto seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbYes Then
            Fg1.RemoveItem (Fg1.Row)
        End If
    End If
End Sub

Private Sub Menu2_5_Click()
    mostrarHistoricoVtas
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 1 Then Exit Sub
        MuestraSegundoTab
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : CambiarEstado
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA UN REGISTRO DE LA TABLA ges_ventaproy
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Activado     |  Boolean   |  INDICA SI SE ACTIVA O DESACTIVA EL REGISTRO
'* DEVUELVE         :
'*****************************************************************************************************
Sub CambiarEstado(Activado As Boolean)
    Dim Rpta As Integer
    TabOne1.CurrTab = 0
    If Activado = False Then
        Rpta = MsgBox("Esta seguro de desactivar la proyeccion de ventas seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    Else
        Rpta = MsgBox("Esta seguro de activar la proyeccion de ventas seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    End If
    
    If Rpta = vbYes Then
        If Activado = False Then
            xCon.Execute "UPDATE ges_ventaproy SET ges_ventaproy.activo = 0 Where (((ges_ventaproy.id) = " & RstPlanes("id") & "))"
            MsgBox "La proyeccion de ventas se desactivo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Else
            xCon.Execute "UPDATE ges_ventaproy SET ges_ventaproy.activo = -1 Where (((ges_ventaproy.id) = " & RstPlanes("id") & "))"
            MsgBox "La proyeccion de ventas se activo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    RstPlanes.Requery
    Dg1.Refresh
End Sub

Private Sub configurarBotones()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo: CmdProcesar.Visible = True: CmdAddProy.Visible = True
    
    If Button.Index = 2 Then
        Modificar
    End If
    
    If Button.Index = 3 Then
        Eliminar
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstPlanes.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then
        Cancelar
    End If
    
    If Button.Index = 14 Then
        Unload Me
    End If
End Sub

Function BuscaItemGrid(CodigoProducto As String) As Boolean
    Dim A As Integer
    
    If Fg1.Rows > 2 Then
        For A = 1 To Fg1.Rows
            If CodigoProducto = Fg1.TextMatrix(A, 1) Then
                MsgBox "El producto ya fue seleccionado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                BuscaItemGrid = True
                Exit Function
            End If
            If A = Fg1.Rows - 1 Then
                Exit For
            End If
        Next A
    End If
    BuscaItemGrid = False
End Function

'*****************************************************************************************************
'* Nombre Archivo   : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTAÑA DETALLE DEL FORMULARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    'Dim A As Integer
    Dim cSQL As String
    
    Dim A, xCol, B As Integer
    Dim Total As Double
    
    Dim xMes As Integer
    Dim xMesAux As String
    Dim idMesIni As Integer
    Dim idAñoIni As Integer
    
    Agregando = True
    
    Blanquea
    
    TxtDesc.Text = RstPlanes("descripcion")
    TxtFchIni.Valor = RstPlanes("fchini")
    TxtFchFin.Valor = RstPlanes("fchfin")
    
    idMesIni = Format(TxtFchIni.Valor, "m")
    idAñoIni = Format(TxtFchIni.Valor, "yyyy")
    
    xMes = idMesIni
    
    Dim indicador As Integer
    indicador = calcularIndicador(CDate(RstPlanes("fchini")), CDate(RstPlanes("fchfin")))
    
    cSQL = "TRANSFORM First(ges_ventaproydet2.cantidad) AS PrimeroDecantidad " _
    + vbCr + "SELECT ges_ventaproydet2.idpro, alm_inventario.descripcion, ges_ventaproydet2.codigo, mae_unidades.abrev, Sum(ges_ventaproydet2.cantidad) AS [Total de cantidad] " _
    + vbCr + "FROM (ges_ventaproydet2 LEFT JOIN alm_inventario ON ges_ventaproydet2.idpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
    + vbCr + "Where (((ges_ventaproydet2.Id) =" & RstPlanes("id") & ")) " _
    + vbCr + "GROUP BY ges_ventaproydet2.idpro, alm_inventario.descripcion, ges_ventaproydet2.codigo, mae_unidades.abrev " _
    + vbCr + "PIVOT ges_ventaproydet2.idmes;"

    RST_Busq Rst, cSQL, xCon
    
    
    Fg1.Rows = 1
    Fg1.Cols = 20
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Total = 0
            Fg1.TextMatrix(A, 1) = NulosC(Rst("idpro"))
            Fg1.TextMatrix(A, 2) = NulosC(Rst("descripcion"))
            Fg1.TextMatrix(A, 3) = NulosC(Rst("codigo"))
            Fg1.TextMatrix(A, 4) = NulosC(Rst("abrev"))
            
            xMesAux = xMes
            For B = 1 To indicador
                Fg1.TextMatrix(A, B + 5) = Format(NulosN(Rst("" & xMesAux & "")), FORMAT_MONTO)
                Total = Total + NulosN(Rst("" & xMesAux & ""))
                xMesAux = xMesAux + 1
                If xMesAux > 12 Then xMesAux = 1
            Next B
            
            Fg1.TextMatrix(A, 5) = Format(Total, FORMAT_MONTO)
            Rst.MoveNext
        Next A
    End If
    
    configurarGrid Fg1, TxtFchIni.Valor, TxtFchFin.Valor
    
    LblNumItem.Caption = Rst.RecordCount
    Fg1_RowColChange
    
    configurarVista
    Agregando = False
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        If ButtonMenu.Index = 1 Then Modificar
        If ButtonMenu.Index = 2 Then CambiarEstado True
    End If
    If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then Eliminar
        If ButtonMenu.Index = 2 Then CambiarEstado False
    End If
End Sub

Private Sub TxtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtPorcentaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtPorcentaje.Text = Format(TxtPorcentaje.Text, "0.00")
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : CompiarValoresAplicarPorcentaje
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : COPIA LOS VALORES DE LA CELDA DE LOS CONTROLES Fg1 Y Fg2
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub CompiarValoresAplicarPorcentaje()
    Dim A As Integer
    Dim indicador As Integer
    
    If NulosC(TxtPorcentaje.Text) = "" Then
        MsgBox "No ha especificado el porcentaje de aumento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtPorcentaje.SetFocus
        Exit Sub
    End If
    
    indicador = calcularIndicador(TxtFchIni.Valor, TxtFchFin.Valor)
    For A = 1 To indicador
        Fg1.TextMatrix(Fg1.Row, A + 5) = Val(Fg2.TextMatrix(Fg2.Row, A + 1)) * ((Val(TxtPorcentaje.Text) / 100) + 1)
    Next A
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : ExportarExcel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A MS EXCEL LOS DATOS DEL CONTROL Fg1
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub ExportarExcel()
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    Dim xCad As String
    Dim objExcel As Object
    
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    objExcel.WindowState = 2
    objExcel.Workbooks.Add
   
    With objExcel.ActiveSheet
        xFilas = 1
        .Cells(xFilas, 2) = "Cronograma de Entregas "
        .Cells(xFilas, 4) = Fg1.TextMatrix(Fg1.Row, 2)
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Periodo "
        .Cells(xFilas, 3) = "Desde: "
        .Cells(xFilas, 4) = "'" + TxtFchIni.Valor
        .Cells(xFilas, 5) = "Hasta: "
        .Cells(xFilas, 6) = "'" + TxtFchFin.Valor
        
        xFilas = xFilas + 2
        For A = 0 To Fg2.Rows - 1
            For B = 1 To Fg2.Cols - 1
                If A = 0 Then
                    .Cells(xFilas, B + 1) = "'" + Fg2.TextMatrix(A, B)
                Else
                    If (B = 1) Then
                        .Cells(xFilas, B + 1) = Fg2.TextMatrix(A, B)
                    Else
                        .Cells(xFilas, B + 1) = NulosN(Fg2.TextMatrix(A, B))
                    End If
                End If
            Next B
            xFilas = xFilas + 1
        Next A
    End With
    
    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Reporte de Pedidos"
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
End Sub
