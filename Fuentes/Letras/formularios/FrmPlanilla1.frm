VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmPlanilla1 
   Caption         =   "Letras - Planilla de Letras"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   6780
      TabIndex        =   20
      Top             =   360
      Width           =   5040
      Begin VB.Label Label20 
         Caption         =   "Nº Registros : "
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   2130
         TabIndex        =   22
         Top             =   75
         Width           =   1920
      End
      Begin VB.Label LblNumRegistros 
         Alignment       =   2  'Center
         Caption         =   "LblNumRegistros"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   4110
         TabIndex        =   21
         Top             =   75
         Width           =   855
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7185
      Left            =   0
      TabIndex        =   10
      Top             =   375
      Width           =   11865
      _cx             =   20929
      _cy             =   12674
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
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "v"
         Height          =   6765
         Left            =   12510
         TabIndex        =   15
         Top             =   375
         Width           =   11775
         Begin VB.CheckBox ChkTC 
            Caption         =   "Check2"
            Enabled         =   0   'False
            Height          =   195
            Left            =   8025
            TabIndex        =   47
            Top             =   1050
            Width           =   195
         End
         Begin VB.TextBox TxtTC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   8220
            TabIndex        =   46
            Text            =   "TxtTC"
            Top             =   990
            Width           =   1065
         End
         Begin VB.TextBox TxtGlosa 
            Height          =   285
            Left            =   1290
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "TxtGlosa"
            Top             =   1980
            Width           =   10170
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   1740
            Picture         =   "FrmPlanilla1.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   1020
            Width           =   240
         End
         Begin VB.Frame Frame10 
            Caption         =   "[ Periodo ]"
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
            Height          =   780
            Left            =   9390
            TabIndex        =   38
            Top             =   840
            Width           =   2370
            Begin VB.Label LblMes1 
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
               Height          =   240
               Left            =   120
               TabIndex        =   39
               Top             =   330
               Width           =   2100
            End
         End
         Begin VB.Frame Frame3 
            Height          =   555
            Left            =   2820
            TabIndex        =   36
            Top             =   3150
            Visible         =   0   'False
            Width           =   7170
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   255
               Left            =   270
               TabIndex        =   37
               Top             =   210
               Width           =   6705
               _ExtentX        =   11827
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   1
               Scrolling       =   1
            End
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Generar Asientos"
            Height          =   525
            Left            =   6000
            TabIndex        =   33
            Top             =   420
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CommandButton CmdBusMod 
            Height          =   240
            Left            =   1740
            Picture         =   "FrmPlanilla1.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1365
            Width           =   240
         End
         Begin VB.CommandButton CmdBusBan 
            Height          =   240
            Left            =   1740
            Picture         =   "FrmPlanilla1.frx":0264
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1695
            Width           =   240
         End
         Begin VB.TextBox TxtIdModalidad 
            Height          =   300
            Left            =   1290
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "TxtModalidad"
            Top             =   1335
            Width           =   705
         End
         Begin VB.TextBox TxtIdBcoCta 
            Height          =   300
            Left            =   1290
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "TxtIdBan"
            Top             =   1650
            Width           =   705
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   1290
            TabIndex        =   1
            Top             =   660
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
            Locked          =   -1  'True
         End
         Begin VB.TextBox TxtNumPla 
            Height          =   300
            Left            =   1290
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "TxtNumPla"
            Top             =   345
            Width           =   1185
         End
         Begin VB.TextBox TxtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   10185
            TabIndex        =   24
            Text            =   "TxtTotal"
            Top             =   6180
            Width           =   1230
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   3630
            Left            =   75
            TabIndex        =   8
            Top             =   2340
            Width           =   11610
            _cx             =   20479
            _cy             =   6403
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
            BackColor       =   14810879
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   14810879
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
            Rows            =   10
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmPlanilla1.frx":0396
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
         Begin VB.Frame Frame4 
            Height          =   795
            Left            =   75
            TabIndex        =   23
            Top             =   5940
            Width           =   6180
            Begin VB.CommandButton CmdSelLet 
               Caption         =   "Seleccionar Letra"
               Height          =   420
               Left            =   1530
               TabIndex        =   7
               Top             =   240
               Width           =   1425
            End
            Begin VB.CommandButton CmdDelLet 
               Caption         =   "Eliminar Letra"
               Height          =   420
               Left            =   2985
               TabIndex        =   9
               Top             =   240
               Width           =   1425
            End
            Begin VB.CommandButton CmdAddLet 
               Caption         =   "Agregar Letra"
               Height          =   420
               Left            =   90
               TabIndex        =   6
               Top             =   240
               Width           =   1425
            End
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   1290
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   2
            Text            =   "TxtIdMon"
            Top             =   990
            Width           =   705
         End
         Begin VB.Label lblReg 
            Caption         =   "lblReg"
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
            Height          =   270
            Left            =   9480
            TabIndex        =   49
            Top             =   300
            Width           =   2250
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "T.C."
            Height          =   195
            Left            =   7710
            TabIndex        =   48
            Top             =   1035
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   10
            Left            =   105
            TabIndex        =   45
            Top             =   2010
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   105
            TabIndex        =   44
            Top             =   1056
            Width           =   585
         End
         Begin VB.Label LblDescMoneda 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescMoneda"
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
            Left            =   2040
            TabIndex        =   43
            Top             =   990
            Width           =   2595
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cuenta"
            Height          =   195
            Left            =   105
            TabIndex        =   41
            Top             =   1692
            Width           =   810
         End
         Begin VB.Label LblDescNumCta 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescNumCta"
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
            Left            =   2040
            TabIndex        =   40
            Top             =   1650
            Width           =   2595
         End
         Begin VB.Label LblDescModalidad 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescModalidad"
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
            Left            =   2025
            TabIndex        =   35
            Top             =   1365
            Width           =   2010
         End
         Begin VB.Label LblDescBan 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescBan"
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
            Left            =   5460
            TabIndex        =   34
            Top             =   1650
            Width           =   3825
         End
         Begin VB.Label LblIdModalidad 
            AutoSize        =   -1  'True
            Caption         =   "LblIdModalidad"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   6180
            TabIndex        =   32
            Top             =   1380
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label LblIdBcoCta 
            AutoSize        =   -1  'True
            Caption         =   "LblIdBanco"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   5010
            TabIndex        =   31
            Top             =   1020
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Modalidad"
            Height          =   195
            Index           =   2
            Left            =   105
            TabIndex        =   29
            Top             =   1374
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Index           =   1
            Left            =   4770
            TabIndex        =   28
            Top             =   1692
            Width           =   465
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Emisi{on"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   27
            Top             =   738
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nº Planilla"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   26
            Top             =   420
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "TOTAL ==>"
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
            Left            =   8985
            TabIndex        =   25
            Top             =   6225
            Width           =   990
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Planilla de Letras"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   60
            TabIndex        =   18
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   8010
            TabIndex        =   17
            Top             =   60
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6765
         Left            =   45
         TabIndex        =   11
         Top             =   375
         Width           =   11775
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   15
            TabIndex        =   12
            Top             =   330
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11324
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
            Columns(1).Caption=   "Nº Reg"
            Columns(1).DataField=   "numreg2"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Planilla"
            Columns(2).DataField=   "numdoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Nº Letras"
            Columns(3).DataField=   "numlet1"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Emi."
            Columns(4).DataField=   "fchemi1"
            Columns(4).NumberFormat=   "Short Date"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Modalidad"
            Columns(5).DataField=   "descmod"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Banco"
            Columns(6).DataField=   "descban"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Nro Cta"
            Columns(7).DataField=   "numcue"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "M"
            Columns(8).DataField=   "descmon"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Importe"
            Columns(9).DataField=   "imptot1"
            Columns(9).NumberFormat=   "0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Saldo"
            Columns(10).DataField=   "impsal1"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1561"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1482"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1482"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1402"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1402"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1323"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1640"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1561"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1879"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1799"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=4551"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=4471"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=2487"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2408"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=820"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=741"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=513"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=1852"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1773"
            Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(62)=   "Column(10).Width=1931"
            Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=1852"
            Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=514"
            Splits(0)._ColumnProps(67)=   "Column(10).Order=11"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=74,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=70,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.alignment=1"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
            _StyleDefs(80)  =   "Named:id=33:Normal"
            _StyleDefs(81)  =   ":id=33,.parent=0"
            _StyleDefs(82)  =   "Named:id=34:Heading"
            _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(84)  =   ":id=34,.wraptext=-1"
            _StyleDefs(85)  =   "Named:id=35:Footing"
            _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(87)  =   "Named:id=36:Selected"
            _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=37:Caption"
            _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(91)  =   "Named:id=38:HighlightRow"
            _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=39:EvenRow"
            _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(95)  =   "Named:id=40:OddRow"
            _StyleDefs(96)  =   ":id=40,.parent=33"
            _StyleDefs(97)  =   "Named:id=41:RecordSelector"
            _StyleDefs(98)  =   ":id=41,.parent=34"
            _StyleDefs(99)  =   "Named:id=42:FilterBar"
            _StyleDefs(100) =   ":id=42,.parent=33"
         End
         Begin VB.Label LblMes 
            AutoSize        =   -1  'True
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   9450
            TabIndex        =   50
            Top             =   75
            Width           =   825
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Planilla de Letras"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   14
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label lblperiodo 
            Caption         =   "lblperiodo(0)"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   0
            Left            =   9450
            TabIndex        =   13
            Top             =   75
            Width           =   1980
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   609
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
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Listado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Registro"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7410
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
               Picture         =   "FrmPlanilla1.frx":0495
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanilla1.frx":09D9
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanilla1.frx":0D6B
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanilla1.frx":0EEF
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanilla1.frx":1343
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanilla1.frx":145B
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanilla1.frx":199F
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanilla1.frx":1EE3
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanilla1.frx":1FF7
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanilla1.frx":210B
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanilla1.frx":255F
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanilla1.frx":26CB
               Key             =   "IMG11"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmPlanilla1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstLet As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim Agregando As Boolean


Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim mIdRegistro& '--identificador del registro
Dim mMesActivo As Integer '--indica el mes activo
Dim fCierrePeriodo As Boolean '--indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)

Dim xHorIni As Date
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Private Sub CmdAddLet_Click()
    pRegistroAdd False
End Sub

Private Sub CmdDelLet_Click()
    
    pRegistroDel
    
End Sub

Private Sub CmdSelLet_Click()
    pRegistroAdd True
End Sub

Private Sub Command1_Click()
    Dim Rpta As Integer
    Exit Sub
    
    Rpta = MsgBox("¿Generar Asientos Contables?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        GenerarAsiento1
    End If
End Sub


Sub GenerarAsiento1()
    Dim A, B As Integer
    RstLet.MoveFirst
    Dim xNumAsiento As String
    Dim RstDia As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim TC As Double
    Dim xFchIni, xFchFin As String
    
    xFchIni = "01/" & Format(mMesActivo, "00") & "/" & Format(AnoTra, "0000")
    
    'eliminamos los asientos contables generados
    xCon.Execute "DELETE * FROM con_diario WHERE (idlib = 42 ) AND (idmes =" & mMesActivo & ")"
    xCon.Execute "DELETE * FROM con_diario WHERE (idlib = 43 ) AND (idmes =" & mMesActivo & ")"
    xCon.Execute "UPDATE let_planilla SET let_planilla.numreg = '' WHERE (((let_planilla.fchreg)=CDate('" & xFchIni & "')));"

    Frame3.Visible = True
    
    'GRABAMOS EL ASIENTO DE LA PLANILLAS
    'RST_Busq RstDia, "SELECT * FROM con_diario", xCon
    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    ProgressBar1.Max = RstLet.RecordCount
    
    For A = 1 To RstLet.RecordCount
        
        ProgressBar1.value = A
        Frame3.Refresh
        xNumAsiento = NuevoNumAsiento(42, mMesActivo, xCon)
        
        TC = HallaTipoCambio(RstLet("fchemi"), 2, Venta, xCon)
        'grabamos la cabecera    DEBE
        RstDia.AddNew
        RstDia("año") = AnoTra
        RstDia("idmes") = mMesActivo
        RstDia("idlib") = 42
        RstDia("idmov") = RstLet("id")
        RstDia("numasi") = xNumAsiento
        RstDia("tc") = TC
        RstDia("fchasi") = "01/" + Format(mMesActivo, "00") + "/" + Trim(Str(AnoTra))
        RstDia("fchdoc") = RstLet("fchemi")
        
        If RstLet("idmon") = 1 Then
            RstDia("idcue") = RstLet("idcuensol")
            RstDia("impdebsol") = RstLet("imptot")
            RstDia("impdebdol") = 0
        Else
            RstDia("idcue") = RstLet("idcuendol")
            RstDia("impdebsol") = NulosN(RstLet("imptot")) * TC
            RstDia("impdebdol") = NulosN(RstLet("imptot"))
        End If
        RstDia.Update
    
        'grabamos el detalle DEBE
        Set RstDet = Nothing
        RST_Busq RstDet, "SELECT let_planilladet.* From let_planilladet WHERE (((let_planilladet.idpla)=" & RstLet("id") & "))", xCon
        
        If RstDet.RecordCount <> 0 Then
            RstDet.MoveFirst
            For B = 1 To RstDet.RecordCount
            
                RstDia.AddNew
                RstDia("año") = AnoTra
                RstDia("idmes") = mMesActivo
                RstDia("idlib") = 42
                RstDia("idmov") = RstLet("id")
                RstDia("numasi") = xNumAsiento
                RstDia("tc") = TC
                RstDia("fchasi") = "01/" + Format(mMesActivo, "00") + "/" + Trim(Str(AnoTra))
                RstDia("fchdoc") = RstLet("fchemi")
                
                If RstLet("idmon") = 1 Then
                    RstDia("idcue") = 73
                    RstDia("imphabsol") = RstDet("implet")
                    RstDia("imphabdol") = 0
                Else
                    RstDia("idcue") = 74
                    RstDia("imphabsol") = NulosN(RstDet("implet")) * TC
                    RstDia("imphabdol") = NulosN(RstDet("implet"))
                End If
                RstDia.Update
            
                RstDet.MoveNext
                If RstDet.EOF = True Then Exit For
            Next B
        End If

        xCon.Execute "UPDATE let_planilla SET let_planilla.numreg = '" & Format(mMesActivo, "00") & xNumAsiento & "' WHERE (((let_planilla.id)=" & RstLet("id") & "))"
        
        RstLet.MoveNext
        If RstLet.EOF = True Then Exit For
    Next A
    
    'GRABAMOS EL ASIENTO DEL ABONO
    RstLet.MoveFirst
    ProgressBar1.value = 0
    ProgressBar1.Max = RstLet.RecordCount
    
    For A = 1 To RstLet.RecordCount
    
        ProgressBar1.value = A
        Frame3.Refresh
    
        xNumAsiento = NuevoNumAsiento(43, mMesActivo, xCon)
        
        TC = HallaTipoCambio(RstLet("fchemi"), 2, Venta, xCon)
        
        
        'grabamos la cabecera    DEBE
        RstDia.AddNew
        RstDia("año") = AnoTra
        RstDia("idmes") = mMesActivo
        RstDia("idlib") = 43
        RstDia("idmov") = RstLet("id")
        RstDia("numasi") = xNumAsiento
        RstDia("tc") = TC
        RstDia("fchasi") = "01/" + Format(mMesActivo, "00") + "/" + Trim(Str(AnoTra))
        RstDia("fchdoc") = RstLet("fchemi")
        RstDia("idcue") = RstLet("ctactaban")
        
        If RstLet("idmon") = 1 Then
            RstDia("impdebsol") = RstLet("imptot")
            RstDia("impdebdol") = 0
        Else
            
            RstDia("impdebsol") = NulosN(RstLet("imptot")) * TC
            RstDia("impdebdol") = NulosN(RstLet("imptot"))
        End If
        RstDia.Update
        
        'grabamos la cabecera    HABER
        RstDia.AddNew
        RstDia("año") = AnoTra
        RstDia("idmes") = mMesActivo
        RstDia("idlib") = 43
        RstDia("idmov") = RstLet("id")
        RstDia("numasi") = xNumAsiento
        RstDia("tc") = TC
        RstDia("fchasi") = "01/" + Format(mMesActivo, "00") + "/" + Trim(Str(AnoTra))
        RstDia("fchdoc") = RstLet("fchemi")
        
        If RstLet("idmon") = 1 Then
            RstDia("idcue") = RstLet("idcuensol")
            RstDia("imphabsol") = RstLet("imptot")
            RstDia("imphabdol") = 0
        Else
            RstDia("idcue") = RstLet("idcuendol")
            RstDia("imphabsol") = NulosN(RstLet("imptot")) * TC
            RstDia("imphabdol") = NulosN(RstLet("imptot"))
        End If
        RstDia.Update
        
        RstLet.MoveNext
        If RstLet.EOF = True Then Exit For
    Next A
    Frame3.Visible = False
End Sub


Private Sub Dg1_DblClick()
    If TabOne1.CurrTab = 0 Then
        Exit Sub
    Else
        TabOne1.CurrTab = 1
    End If
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstLet
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)

    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstLet.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
    
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstLet("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col <> 2 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nTitulo As String
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Falta especificar la moneda", vbExclamation, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    Select Case Col
        Case 1 '--BUSCANDO CLIENTE
        Case 2 '--BUSCANDO LETRA
            ReDim xCampos(6, 5) As String
            xCampos(0, 0) = "Cliente":      xCampos(0, 1) = "nombre":    xCampos(0, 2) = "2800":     xCampos(0, 3) = "C":    xCampos(0, 4) = "N"
            xCampos(1, 0) = "N° Letra":     xCampos(1, 1) = "numletra":  xCampos(1, 2) = "1800":     xCampos(1, 3) = "C":    xCampos(1, 4) = "S"
            xCampos(2, 0) = "Fch Emi":      xCampos(2, 1) = "fchemi":    xCampos(2, 2) = "950":      xCampos(2, 3) = "F":    xCampos(2, 4) = "N"
            xCampos(3, 0) = "Fch Venc":     xCampos(3, 1) = "fchven":    xCampos(3, 2) = "950":      xCampos(3, 3) = "F":    xCampos(3, 4) = "N"
            xCampos(4, 0) = "M":            xCampos(4, 1) = "moneda":    xCampos(4, 2) = "500":      xCampos(4, 3) = "C":    xCampos(3, 4) = "N"
            xCampos(5, 0) = "Importe":      xCampos(5, 1) = "importe":   xCampos(5, 2) = "900":      xCampos(5, 3) = "N":    xCampos(3, 4) = "N"
        
        
            nSQL = "SELECT mae_cliente.nombre, let_letradet.numser, let_letradet.fchemi, let_letradet.fchven, let_letradet.numdoc, mae_moneda.simbolo AS moneda, " _
                & " [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser] AS numletra, let_letradet.implet AS importe,let_letradet.idlet,let_letradet.corr " _
                & " FROM ((let_letra LEFT JOIN mae_moneda ON let_letra.idmon = mae_moneda.id) RIGHT JOIN let_letradet ON let_letra.id = let_letradet.idlet) LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id " _
                & " WHERE (((let_letra.idmon)=" & NulosN(TxtIdMon.Text) & ")) "
            
            nTitulo = "Buscando Letras"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "numletra", Principio, ""
        Case Else
            Exit Sub
    End Select

    

    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    Agregando = True
    If Col = 1 Then '--DEL CLIENTE
        
    ElseIf Col = 2 Then '--DE LA LETRA

        Fg1.TextMatrix(Row, 1) = NulosC(xRs("nombre"))
        Fg1.TextMatrix(Row, 2) = NulosC(xRs("numletra"))
        Fg1.TextMatrix(Row, 3) = Format(NulosC(xRs("fchemi")), FORMAT_DATE)
        Fg1.TextMatrix(Row, 4) = Format(NulosC(xRs("fchven")), FORMAT_DATE)
        Fg1.TextMatrix(Row, 5) = NulosC(xRs("moneda"))
        Fg1.TextMatrix(Row, 6) = Format(xRs("importe"), FORMAT_MONTO)
        Fg1.TextMatrix(Row, 7) = NulosN(xRs("idlet"))
        Fg1.TextMatrix(Row, 8) = NulosN(xRs("corr"))
    End If
    Fg1.SetFocus
    Agregando = False
    Set xRs = Nothing
    
    '--TOTALIZANDO
    TxtTotal.Text = Format(GRID_SUMAR_COL(Fg1, 6), FORMAT_MONTO)
    
    Exit Sub
SALIR:
    Set xRs = Nothing
    Agregando = False
End Sub


Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row = 0 Then Exit Sub
    
    
    Select Case Col
        Case 6 '--IMPORTE
            If IsNumeric(Fg1.TextMatrix(Row, Col)) = False Then
                MsgBox "El valor ingresado no es numérico", vbCritical, xTitulo
                Fg1.TextMatrix(Row, Col) = ""
            Else
                Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), FORMAT_MONTO)
                '------------------------
            End If
            
            TxtTotal.Text = Format(GRID_SUMAR_COL(Fg1, 6), FORMAT_MONTO)
            
    End Select
    Exit Sub
error:
    
    SHOW_ERROR Me.Name, "Fg1_CellChanged"
End Sub

Private Sub Fg1_EnterCell()
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    
    If Fg1.Col = 2 Or Fg1.Col = 6 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Then KeyAscii = 0
    If KeyAscii = 13 Then Exit Sub
    '--validar los caracteres que se ingresan
    Select Case Col
        Case 6 '--importe
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub


Private Sub Form_Activate()
    If SeEjecuto = False Then
                
        SeEjecuto = True
        mMesActivo = xMes
    
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        Cargar
        
    End If
End Sub


Private Sub Form_Load()

    Dg1.Columns("fchemi1").NumberFormat = FORMAT_DATE
    
    Dg1.Columns("imptot1").NumberFormat = FORMAT_MONTO
    
    Dg1.Columns("impsal1").NumberFormat = FORMAT_MONTO
    
    Agregando = False

    SeEjecuto = False
    QueHace = 3
    Frame7.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    Fg1.ColWidth(7) = 0
    Fg1.ColWidth(8) = 0
    Fg1.Rows = 1
    TabOne1.CurrTab = 0
    Fg1.Editable = flexEDNone
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.BackColorSel = &H80&
    
    GRID_COMBOLIST Fg1, 2
    
End Sub

Sub Cargar()
    Dim nSQL As String
    
    
    
    '----------------
    OpcionesPeriodo
    '----------------
    
''    RST_Busq RstLet, "SELECT let_planilla.id, let_planilla.numdoc, let_planilla.numlet, let_planilla.imptot, let_planilla.fchemi, mae_bancos.descripcion AS descban, " _
''        & " let_modalidad.descripcion AS descmod, let_planilla.idmod, mae_banconumcta.idban, let_planilla.numreg, let_planilla.fchreg, mae_moneda.simbolo AS descmon, " _
''        & " let_planilla.idmon, let_modalidad_1.idcuensol, let_modalidad_1.idcuendol, mae_banconumcta.idcuen AS ctactaban " _
''        & " FROM ((((let_planilla LEFT JOIN mae_bancos ON let_planilla.idban = mae_bancos.id) LEFT JOIN let_modalidad ON let_planilla.idmod = let_modalidad.id) " _
''        & " LEFT JOIN mae_moneda ON let_planilla.idmon = mae_moneda.id) LEFT JOIN let_modalidad AS let_modalidad_1 ON let_planilla.idmod = let_modalidad_1.id) " _
''        & " LEFT JOIN mae_banconumcta ON let_planilla.idbcocta = mae_banconumcta.id " _
''        & " WHERE (((let_planilla.fchreg)>=CDate('" & xFchIni & "') And (let_planilla.fchreg)<=CDate('" & xFchFin & "')))", xCon
    
        '--limpiar los filtros
    TDB_FiltroLimpiar Dg1
    Set RstLet = Nothing
    Set Dg1.DataSource = Nothing
    DoEvents
    '----------------------
    
    nSQL = "SELECT let_planilla.*, mae_bancos.descripcion AS descban, let_modalidad.descripcion AS descmod, " _
        & " mae_banconumcta.idban, mae_moneda.simbolo AS descmon, mae_banconumcta.idcuen AS ctactaban,mae_banconumcta.numcue,IIF(let_planilla.anulado=-1,0,IIf([let_planilla].[tc]=0,[con_tc].[impven],[let_planilla].[tc])) & '' AS impven1, " _
        & " let_planilla.numlet & '' as numlet1 ,let_planilla.imptot & '' as imptot1, let_planilla.fchemi & '' as fchemi1,Mid([let_planilla]![numreg],1,2) & [mae_libros]![codsun] & Mid([let_planilla]![numreg],3,4) AS numreg2,let_planilla.impsal & '' as impsal1 " _
        & " FROM ((mae_bancos RIGHT JOIN (((let_planilla LEFT JOIN let_modalidad ON let_planilla.idmod = let_modalidad.id) LEFT JOIN mae_moneda ON let_planilla.idmon = mae_moneda.id)  " _
        & " LEFT JOIN mae_banconumcta ON let_planilla.idbcocta = mae_banconumcta.id) ON mae_bancos.id = mae_banconumcta.idban) LEFT JOIN mae_libros ON let_planilla.idlib = mae_libros.id) LEFT JOIN con_tc ON let_planilla.fchemi = con_tc.fecha " _
        & " WHERE let_planilla.idmes = " & mMesActivo

    RST_Busq RstLet, nSQL, xCon
    
    Set Dg1.DataSource = RstLet
    
    LblNumRegistros.Caption = RstLet.RecordCount
    
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 1 Then Exit Sub
        MuestraSegundoTab
    End If
End Sub

Sub MuestraSegundoTab()
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    Dim nSQL As String
    
    Blanquea
    If RstLet.EOF = True Or RstLet.BOF = True Or RstLet.RecordCount = 0 Then Exit Sub
    
    lblReg.Caption = "Nº Reg. " & NulosC(RstLet("numreg2"))
    
    TxtNumPla.Text = NulosC(RstLet("numdoc"))
    TxtFchEmi.Valor = NulosC(RstLet("fchemi"))
    
    TxtIdBcoCta.Text = NulosC(RstLet("idbcocta"))
    LblDescNumCta.Caption = NulosC(RstLet("numcue"))
    LblDescBan.Caption = NulosC(Busca_Codigo(NulosN(RstLet("idban")), "id", "descripcion", "mae_bancos", "N", xCon))
    
    TxtIdModalidad.Text = RstLet("idmod")
    LblDescModalidad.Caption = Busca_Codigo(NulosN(RstLet("idmod")), "id", "descripcion", "let_modalidad", "N", xCon)
    
    TxtIdMon.Text = RstLet("idmon")
    LblDescMoneda.Caption = Busca_Codigo(NulosN(RstLet("idmon")), "id", "descripcion", "mae_moneda", "N", xCon)
    
    
    
        '--tipo de cambio
    If NulosN(RstLet("tc")) = 0 Then
        ChkTC.value = 0
        TxtTC.Text = NulosN(RstLet("impven1"))
        TxtTC.BackColor = &H8000000F
        TxtTC.Enabled = False
    Else
        ChkTC.value = 1
        TxtTC.Text = NulosN(RstLet("tc"))
        TxtTC.BackColor = vbWhite
        TxtTC.Enabled = True
    End If
    If QueHace = 3 Then TxtTC.BackColor = &H8000000F
    
    
    nSQL = "SELECT let_planilladet.idpla, mae_cliente.nombre, let_letradet.numser, let_letradet.fchemi, let_letradet.fchven, let_letradet.numdoc, let_planilladet.implet, mae_moneda.simbolo AS moneda, [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser] AS numletra, let_letradet.idlet, let_letradet.corr " _
        + vbCr + " FROM ((let_letra LEFT JOIN mae_moneda ON let_letra.idmon = mae_moneda.id) RIGHT JOIN (let_planilladet INNER JOIN let_letradet ON (let_planilladet.idcorrlet = let_letradet.corr) AND (let_planilladet.idlet = let_letradet.idlet)) ON let_letra.id = let_letradet.idlet) LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id " _
        + vbCr + " WHERE (((let_planilladet.idpla)=" & NulosN(RstLet("id")) & " )) ORDER BY mae_cliente.nombre,let_letradet.numdoc "

    RST_Busq RstDet, nSQL, xCon
    
    If RstDet.RecordCount <> 0 Then
        Agregando = True
        RstDet.MoveFirst
        Fg1.Rows = 1
        
        For A = 1 To RstDet.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(RstDet("nombre"))
            Fg1.TextMatrix(A, 2) = NulosC(RstDet("numletra"))
            Fg1.TextMatrix(A, 3) = Format(NulosC(RstDet("fchemi")), FORMAT_DATE)
            Fg1.TextMatrix(A, 4) = Format(NulosC(RstDet("fchven")), FORMAT_DATE)
            Fg1.TextMatrix(A, 5) = NulosC(RstDet("moneda"))
            Fg1.TextMatrix(A, 6) = Format(RstDet("implet"), FORMAT_MONTO)
            Fg1.TextMatrix(A, 7) = NulosN(RstDet("idlet"))
            Fg1.TextMatrix(A, 8) = NulosN(RstDet("corr"))
            
            RstDet.MoveNext
            If RstDet.EOF = True Then Exit For
        Next A
        
        Agregando = False
        TxtTotal.Text = Format(GRID_SUMAR_COL(Fg1, 6), FORMAT_MONTO)
        
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar

    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstLet.Requery
            Dg1.Refresh
            
            If RstLet.RecordCount <> 0 Then RstLet.MoveFirst
            RstLet.Find "id = " & mIdRegistro & ""
            If RstLet.EOF = True Then RstLet.MoveFirst
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    
    If Button.Index = 11 Then CambiarMes
    
    If Button.Index = 9 Then
        '--limpiar los filtros
        RstLet.Filter = ""
        TDB_FiltroLimpiar Dg1
            
    End If
    
    If Button.Index = 15 Then
        Set RstLet = Nothing
        Unload Me
    End If
End Sub

Sub CambiarMes()
    mMesActivo = SeleccionaMes(xCon)
    Cargar
    TabOne1.CurrTab = 0
    Dg1.DataSource = RstLet
End Sub



Private Sub OpcionesPeriodo()
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    LblMes1.Caption = LblMes.Caption
    
    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    
End Sub



Sub Blanquea()
    
    lblReg.Caption = ""
    
    TxtNumPla.Text = ""
    TxtFchEmi.Valor = ""
    
    TxtIdMon.Text = ""
    LblDescMoneda.Caption = ""
    LblIdBcoCta.Caption = 0
    
    
    TxtIdBcoCta.Text = ""
    LblDescBan.Caption = ""
    LblDescNumCta.Caption = ""
    LblIdBcoCta.Caption = ""
    
    TxtIdModalidad.Text = ""
    LblDescModalidad.Caption = ""
    LblIdModalidad.Caption = ""
    
    TxtGlosa.Text = ""
    TxtTC.Text = ""
    TxtTotal.Text = ""
    
    Fg1.Rows = 1
    
End Sub


Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripción":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Código":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
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
            LblDescMoneda.Caption = xRs("descripcion")
            TxtIdModalidad.SetFocus
            End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub TxtFchEmi_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtFchEmi.Valor) <> "" Then
        If ChkTC.value = 0 Then TxtTC.Text = HallaTipoCambio(TxtFchEmi.Valor, 2, Venta, xCon)
    Else
        If ChkTC.value = 0 Then TxtTC.Text = "0.00"
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtIdMon.Text) = "" Then
        LblDescMoneda.Caption = ""
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    
    'buscamos el codigo de la moneda         digitada
    RST_Busq xRs1, "SELECT * FROM mae_moneda WHERE id = " & NulosN(TxtIdMon.Text) & "", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtIdMon.Text = ""
        LblDescMoneda.Caption = ""
    Else
        LblDescMoneda.Caption = Trim(xRs1("descripcion"))
        

    End If
    Set xRs1 = Nothing

End Sub



Private Sub CmdBusBan_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(6, 4) As String
    
    xCampos(0, 0) = "Banco":        xCampos(0, 1) = "banco":        xCampos(0, 2) = "3000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "N° Cta Cte":   xCampos(1, 1) = "numcue":       xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "M":            xCampos(2, 1) = "simbolo":      xCampos(2, 2) = "500":          xCampos(2, 3) = "C"
    xCampos(3, 0) = "N° Cta":       xCampos(3, 1) = "cuentanum":    xCampos(3, 2) = "900":         xCampos(3, 3) = "C"
    xCampos(4, 0) = "Nombre Cta":   xCampos(4, 1) = "cuentanom":    xCampos(4, 2) = "2300":         xCampos(4, 3) = "C"
    xCampos(5, 0) = "Id":           xCampos(5, 1) = "idbcocta":     xCampos(5, 2) = "500":          xCampos(5, 3) = "N"
    
    
    xform.SQLCad = "SELECT mae_banconumcta.id AS idbcocta, mae_bancos.descripcion AS banco, mae_banconumcta.numcue, mae_bancos.numruc, con_planctas.cuenta AS cuentanum, con_planctas.descripcion AS cuentanom, mae_moneda.simbolo " _
        & " FROM (((mae_banconumcta LEFT JOIN mae_bancos ON mae_banconumcta.idban = mae_bancos.id) RIGHT JOIN let_modalidadctabco ON mae_banconumcta.id = let_modalidadctabco.idbcocta) LEFT JOIN con_planctas ON let_modalidadctabco.idcuen = con_planctas.id) LEFT JOIN mae_moneda ON let_modalidadctabco.idmon = mae_moneda.id " _
        & " WHERE (((mae_banconumcta.id)<>0) AND ((let_modalidadctabco.idmon)=" & NulosN(TxtIdMon.Text) & ") AND ((let_modalidadctabco.idmod)=" & NulosN(TxtIdModalidad.Text) & ")) " _
        & " ORDER BY mae_bancos.descripcion; "

    
    xform.Titulo = "Buscando Cta Cte de Banco"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "banco"
    xform.CampoBusca = "numcue"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdBcoCta.Text = xRs("idbcocta")
            LblDescBan.Caption = NulosC(xRs("banco"))
            LblDescNumCta.Caption = NulosC(xRs("numcue"))
            TxtGlosa.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub TxtIdBcoCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdBcoCta_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusBan_Click
    End If
End Sub

Private Sub TxtIdBcoCta_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtIdBcoCta.Text) = "" Then
        TxtIdBcoCta.Text = ""
        LblDescBan.Caption = ""
        LblDescNumCta.Caption = ""
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    
    'buscamos el codigo de la moneda         digitada
    RST_Busq xRs1, "SELECT mae_banconumcta.id AS idbcocta, mae_bancos.descripcion AS banco, mae_banconumcta.numcue, mae_bancos.numruc " _
        & " FROM (mae_banconumcta LEFT JOIN mae_bancos ON mae_banconumcta.idban = mae_bancos.id) RIGHT JOIN let_modalidadctabco ON mae_banconumcta.id = let_modalidadctabco.idbcocta " _
        & " Where (((mae_banconumcta.id) = " & NulosN(TxtIdBcoCta.Text) & ")) " _
        & " ORDER BY mae_bancos.descripcion; ", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtIdBcoCta.Text = ""
        LblDescBan.Caption = ""
        LblDescNumCta.Caption = ""
    Else
        TxtIdBcoCta.Text = xRs1("idbcocta")
        LblDescBan.Caption = NulosC(xRs1("banco"))
        LblDescNumCta.Caption = NulosC(xRs1("numcue"))
        
    End If
    Set xRs1 = Nothing

End Sub




'*************************************************************
Private Sub CmdBusMod_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripción":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Código":         xCampos(1, 1) = "id":               xCampos(1, 2) = "500":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM let_modalidad ORDER BY descripcion"
    
    xform.Titulo = "Buscando Modalidad"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdModalidad.Text = xRs("id")
            LblDescModalidad.Caption = xRs("descripcion")
            CmdAddLet.SetFocus
            End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub TxtIdModalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdModalidad_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMod_Click
    End If
End Sub

Private Sub TxtIdModalidad_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtIdModalidad.Text) = "" Then
        LblDescMoneda.Caption = ""
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    
    'buscamos el codigo de la moneda         digitada
    RST_Busq xRs1, "SELECT * FROM let_modalidad WHERE id = " & NulosN(TxtIdModalidad.Text) & "", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtIdModalidad.Text = ""
        LblDescModalidad.Caption = ""
    Else
        LblDescModalidad.Caption = Trim(xRs1("descripcion"))

    End If
    Set xRs1 = Nothing

End Sub

'*************************************************************

Private Sub pRegistroAdd(Optional fSeleccionVarios As Boolean = False)
     
    If QueHace = 3 Then Exit Sub
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Falta especificar la moneda", vbExclamation, xTitulo
        Exit Sub
    End If
    
    Dim nSQLIdLetCorr As String '--almacenara el codigo del la letra que se encuentra en la grilla
    
    '--GENERAR EL WHERE DE LOS ID'S RECETA PARA QUE NO SE REPITAN
    nSQLIdLetCorr = GENERAR_SQL_ID(Fg1, 8, " AND let_letradet.corr", "NOT IN")
    
    '----
    On Error GoTo error
    Dim xRs  As New ADODB.Recordset
    
    Dim nSQL As String
    
    ReDim xCampos(6, 5) As String
    xCampos(0, 0) = "Cliente":      xCampos(0, 1) = "nombre":    xCampos(0, 2) = "3800":     xCampos(0, 3) = "C":    xCampos(0, 4) = "N"
    xCampos(1, 0) = "N° Letra":     xCampos(1, 1) = "numletra":  xCampos(1, 2) = "2200":     xCampos(1, 3) = "C":    xCampos(1, 4) = "S"
    xCampos(2, 0) = "Fch Emi":      xCampos(2, 1) = "fchemi":    xCampos(2, 2) = "1050":      xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
    xCampos(3, 0) = "Fch Venc":     xCampos(3, 1) = "fchven":    xCampos(3, 2) = "1050":      xCampos(3, 3) = "C":    xCampos(3, 4) = "C"
    xCampos(4, 0) = "M":            xCampos(4, 1) = "moneda":    xCampos(4, 2) = "500":      xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
    xCampos(5, 0) = "Importe":      xCampos(5, 1) = "importe":   xCampos(5, 2) = "1000":      xCampos(5, 3) = "N":    xCampos(5, 4) = "N"


    nSQL = "SELECT 0 as xsel, mae_cliente.nombre, let_letradet.numser, let_letradet.fchemi, let_letradet.fchven, let_letradet.numdoc, mae_moneda.simbolo AS moneda, " _
    & " [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser] AS numletra, let_letradet.implet AS importe,let_letradet.idlet,let_letradet.corr " _
    & " FROM ((let_letra LEFT JOIN mae_moneda ON let_letra.idmon = mae_moneda.id) RIGHT JOIN let_letradet ON let_letra.id = let_letradet.idlet) LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id " _
    & " WHERE (((let_letra.idmon)=" & NulosN(TxtIdMon.Text) & ")) and let_letradet.implet>0 "


    
    If fSeleccionVarios = True Then
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Letras"
    Else
        xCampos(0, 2) = "2500":
        xCampos(1, 2) = "1800":
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Letras", "nombre", "numcue", Principio
    End If
    
    Agregando = True
    Dim A As Integer
    Dim xFila As Integer
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    If fSeleccionVarios = True Then xRs.MoveFirst
   
    Do While Not xRs.EOF
        Fg1.Rows = Fg1.Rows + 1
        With Fg1
            Fg1.TextMatrix(.Rows - 1, 1) = NulosC(xRs("nombre"))
            Fg1.TextMatrix(.Rows - 1, 2) = NulosC(xRs("numletra"))
            Fg1.TextMatrix(.Rows - 1, 3) = Format(NulosC(xRs("fchemi")), FORMAT_DATE)
            Fg1.TextMatrix(.Rows - 1, 4) = Format(NulosC(xRs("fchven")), FORMAT_DATE)
            Fg1.TextMatrix(.Rows - 1, 5) = NulosC(xRs("moneda"))
            Fg1.TextMatrix(.Rows - 1, 6) = Format(xRs("importe"), FORMAT_MONTO)
            Fg1.TextMatrix(.Rows - 1, 7) = NulosN(xRs("idlet"))
            Fg1.TextMatrix(.Rows - 1, 8) = NulosN(xRs("corr"))
            
            If fSeleccionVarios = False Then Exit Do
        End With
        If fSeleccionVarios = False Then Exit Do
        xRs.MoveNext
    Loop
    
    TxtTotal.Text = Format(GRID_SUMAR_COL(Fg1, 6), FORMAT_MONTO)
    
SALIR:

    Agregando = False
    
    Set xRs = Nothing
    Fg1.SetFocus
    Exit Sub
error:
    Agregando = False
    Set xRs = Nothing
    MsgBox Err.Description + vbCr + Err.Source, vbCritical, xTitulo
    Err.Clear
    
End Sub



Private Sub pRegistroDel()
    If QueHace = 3 Then Exit Sub
    If Fg1.Row < 0 Then Exit Sub
    If Fg1.Row = 0 Then
        MsgBox "La fila no se puede eliminar" + vbCr + "Seleccione una correcta", vbExclamation
        Exit Sub
    End If
    If MsgBox("Seguro desea eliminar la Letra", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    '--ELIMINAR EL REGISTRO
    Fg1.RemoveItem (Fg1.Row)
    If Fg1.Rows > 1 Then Fg1.Row = 1
    TxtTotal.Text = Format(GRID_SUMAR_COL(Fg1, 6), FORMAT_MONTO)
End Sub



Sub Modificar()
    If RstLet.State = 0 Then Exit Sub
    If RstLet.EOF = True Or RstLet.BOF = True Or RstLet.RecordCount = 0 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If

    QueHace = 2
    
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Bloquea
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    Label5.Caption = "Modificando Planilla de Letras"
    GRID_COMBOLIST Fg1, 2
    TxtFchEmi.SetFocus
    
    xHorIni = Time
   
End Sub


Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub


Sub Bloquea()

    TxtNumPla.Locked = Not TxtNumPla.Locked
    TxtFchEmi.Locked = Not TxtFchEmi.Locked
    ChkTC.Enabled = Not ChkTC.Enabled
    TxtIdModalidad.Locked = Not TxtIdModalidad.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtIdBcoCta.Locked = Not TxtIdBcoCta.Locked
    TxtGlosa.Locked = Not TxtGlosa.Locked
    
End Sub

Sub Cancelar()
    QueHace = 3
    Bloquea
    ActivaTool
    Label5.Caption = "Detalle de Planilla de Letras"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
End Sub

Sub Nuevo()
    Bloquea
    Blanquea
    ActivaTool
    QueHace = 1
    Label5.Caption = "Agregando Planilla de Letras"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    
    xHorIni = Time
    
    TxtNumPla.SetFocus
End Sub




Function Grabar() As Boolean
    
    If NulosC(TxtNumPla.Text) = "" Then
        MsgBox "Falta especificar el Número de Planilla", vbExclamation, xTitulo
        TxtNumPla.SetFocus
        Exit Function
    End If
        
    If IsDate(TxtFchEmi.Valor) = False Then
        MsgBox "Falta especificar la Fecha de Emisión", vbExclamation, xTitulo
        TxtFchEmi.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Falta especificar la Modeda", vbExclamation, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtIdBcoCta.Text) = 0 Then
        MsgBox "Falta especificar la Cuenta Corriente de Banco", vbExclamation, xTitulo
        TxtIdBcoCta.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtIdModalidad.Text) = 0 Then
        MsgBox "Falta especificar la Modalidad", vbExclamation, xTitulo
        TxtIdModalidad.SetFocus
        Exit Function
    End If
    
    If Fg1.Rows = 0 Then
        MsgBox "Falta Especificar el detalle de las Letras", vbExclamation, xTitulo
        Exit Function
    End If
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el Costo", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo SALIR
       
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xCol&, xFil&, xCorr&
    Dim xId As Double
    Dim xNumAsiento As String
    
    On Error GoTo LaCague

    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM let_planilla ", xCon
        xId = HallaCodigoTabla("let_planilla", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstLet("id")
        RST_Busq RstCab, "SELECT * FROM let_planilla WHERE id =" & xId & "", xCon
        xCon.Execute "DELETE * FROM let_planilladet WHERE idpla = " & xId & ""
    End If
    '------------------------
    mIdRegistro = xId
    '------------------------
    RST_Busq RstDet, "SELECT top 1 * FROM let_planilladet", xCon

    RstCab("idlib") = 42
    RstCab("idmod") = NulosN(TxtIdModalidad.Text)
    RstCab("idbcocta") = NulosN(TxtIdBcoCta.Text)
    RstCab("tipdoc") = 130
    If IsDate(TxtFchEmi.Valor) = True Then RstCab("fchemi") = CDate(TxtFchEmi.Valor)
    RstCab("numser") = ""
    RstCab("numdoc") = NulosC(TxtNumPla.Text)
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    RstCab("numlet") = Fg1.Rows - 1
    RstCab("imptot") = NulosN(TxtTotal.Text)
    RstCab("glosa") = Trim(TxtGlosa.Text)
    
    ''RstCab("fchreg") =
    
    If xMes <> 0 And xMes <> 13 Then
        RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    End If
    
    If ChkTC.value = 1 Then RstCab("tc") = NulosN(TxtTC.Text)
    If ChkTC.value = 0 Then RstCab("tc") = 0
    
    RstCab.Update
    '--registrar las letras asociadas a la planilla
    For xFil = 1 To Fg1.Rows - 1
        '---------------------------------------------------------------------------------------------------
        RstDet.AddNew
        RstDet("idpla") = xId
        RstDet("idcli") = 0
        RstDet("idlet") = NulosN(Fg1.TextMatrix(xFil, 7))
        RstDet("idcorrlet") = NulosN(Fg1.TextMatrix(xFil, 8))
        RstDet("implet") = NulosN(Fg1.TextMatrix(xFil, 6))
        
        RstDet.Update
            
    Next xFil
    
    '--generamos es asiento
    xNumAsiento = GenerarAsiento(xCon, 42, CDbl(xId), AnoTra, mMesActivo, 1, 0)
    If xNumAsiento = "" Then GoTo LaCague
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    
    
    xCon.CommitTrans
    
    MsgBox "El movimiento se grabó con éxito" & vbCr & "Registro Nº: " & xNumAsiento, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo

    
    Grabar = True
SALIR:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing
    Exit Function
LaCague:
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" & vbCr & Trim(Err.Description)

    Grabar = False
End Function

Private Sub TxtNumPla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtNumPla_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtNumPla.Text) <> "" Then
    
        
        Dim Rst As New ADODB.Recordset
        Dim nSQLIdPla As String
        '--ver si existe el numero de doc
        If QueHace <> 1 Then nSQLIdPla = " and let_planilla.id <> " & NulosN(RstLet("id"))
        
        RST_Busq Rst, "SELECT mae_bancos.descripcion, mae_banconumcta.numcue, let_planilla.numdoc " _
            & " FROM let_planilla LEFT JOIN (mae_banconumcta LEFT JOIN mae_bancos ON mae_banconumcta.idban = mae_bancos.id) ON let_planilla.idbcocta = mae_banconumcta.id " _
            & " WHERE (((let_planilla.numdoc)='" & NulosC(TxtNumPla.Text) & "')) " & nSQLIdPla, xCon
                
        If Rst.RecordCount <> 0 Then
            
            MsgBox "El número de Planilla ya existe " & vbCr & "Será reemplazado por " + Trim(TxtNumPla.Text), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
        Set Rst = Nothing
        
    End If

End Sub





Private Sub ChkTC_Click()
    If QueHace = 3 Then Exit Sub
    
    If ChkTC.value = 0 Then
        TxtTC.BackColor = &H8000000F
        TxtTC.Enabled = False
        If IsDate(TxtFchEmi.Valor) = True Then
            TxtTC.Text = HallaTipoCambio(TxtFchEmi.Valor, 2, Venta, xCon)
        Else

            Exit Sub
        End If
    Else
        TxtTC.Enabled = True
        TxtTC.BackColor = vbWhite
        TxtTC.SetFocus
    End If
End Sub




Sub Eliminar()
    Dim Rpta As Integer
    Dim nSQL As String
    Dim xId As Double
    Dim rstBus As New ADODB.Recordset
    
    
    If RstLet.RecordCount = 0 Or RstLet.EOF = True Or RstLet.BOF = True Then
        MsgBox "No hay registro para eliminar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    xId = RstLet("id")
    
    '--erificar que el registro a eliminar no tenga movimientos en bancos
    nSQL = "SELECT tes_caja.numreg as registro, tes_caja.glosa " _
        + vbCr + " FROM tes_caja INNER JOIN tes_cajadestinodet ON tes_caja.id = tes_cajadestinodet.idtes " _
        + vbCr + " WHERE (((tes_cajadestinodet.idmod)=19) AND ((tes_cajadestinodet.iddoc)=" & xId & ")); "
    
    RST_Busq rstBus, nSQL, xCon
    
    If rstBus.RecordCount <> 0 Then
        MsgBox "El registro tiene movimiento en Bancos" & vbCr & "N° Reg: " & rstBus("registro") & vbCr & "Glosa: " & rstBus("glosa") & vbCr & "Para eliminar el registro, quite el vínculo de banco", vbInformation, xTitulo
        
        Exit Sub
    End If
    
    Set rstBus = Nothing
    
    
    Rpta = MsgBox("¿Esta seguro de eliminar el asiento seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        
        
        On Error GoTo LaCague
        xCon.BeginTrans
        
        
        'ELIMINAMOS EL ASIENTO REGISTRADO EN EL DIARIO
        xCon.Execute "DELETE * FROM con_diario WHERE ((idmes = " & mMesActivo & ") and (idlib = 42) AND (idmov = " & xId & ")) ;"
        xCon.Execute "DELETE * FROM let_planilladet WHERE idpla= " & xId & ""
        xCon.Execute "DELETE * FROM let_planilla WHERE id = " & xId & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & IdMenuActivo
        
        
        xCon.CommitTrans
        RstLet.Requery
        Dg1.Refresh
        MsgBox "El asiento fue eliminado con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TabOne1.CurrTab = 0
        
        RstLet.Filter = ""
        TDB_FiltroLimpiar Dg1
        
        If RstLet.RecordCount = 0 Then
            Rpta = MsgBox("No se ha registrado ninguna Planilla de Letras, ¿Desea agregar una ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            End If
        End If
    End If
    Exit Sub
LaCague:
    xCon.RollbackTrans
    SHOW_ERROR Me.Name, "Eliminar"
End Sub


