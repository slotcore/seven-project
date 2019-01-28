VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmRegPercepciones 
   Caption         =   "Contabilidad - Percepciones"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   14
      Top             =   375
      Width           =   11880
      _cx             =   20955
      _cy             =   12726
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
         BorderStyle     =   0  'None
         Caption         =   "Detalle de la Cuenta"
         Height          =   6795
         Left            =   12525
         TabIndex        =   18
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame6 
            Caption         =   "( Periodo )"
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
            Height          =   720
            Left            =   9450
            TabIndex        =   50
            Top             =   450
            Width           =   2010
            Begin VB.Label lblperiodo 
               Alignment       =   2  'Center
               Caption         =   "lblperiodo"
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
               Index           =   1
               Left            =   120
               TabIndex        =   51
               Top             =   330
               Width           =   1740
            End
         End
         Begin VB.TextBox TxtGlosa 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   9
            Text            =   "TxtGlosa"
            Top             =   3180
            Width           =   7845
         End
         Begin VB.CommandButton CmdIdPer 
            Height          =   240
            Left            =   2160
            Picture         =   "FrmRegPercepciones.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   2870
            Width           =   240
         End
         Begin VB.CommandButton CmdBusDoc 
            Height          =   240
            Left            =   2160
            Picture         =   "FrmRegPercepciones.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1890
            Width           =   240
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2685
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   6
            Text            =   "TxtNumDoc"
            Top             =   2190
            Width           =   1950
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   5
            Text            =   "TxtNumSer"
            Top             =   2190
            Width           =   900
         End
         Begin VB.Frame Fra_Tipo 
            Caption         =   "[ Tipo de Movimiento ]"
            Enabled         =   0   'False
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
            Height          =   600
            Left            =   165
            TabIndex        =   22
            Top             =   450
            Width           =   3090
            Begin VB.OptionButton Opt_Tipo 
               Caption         =   "Compra"
               Height          =   195
               Index           =   0
               Left            =   405
               TabIndex        =   0
               Top             =   315
               Value           =   -1  'True
               Width           =   1110
            End
            Begin VB.OptionButton Opt_Tipo 
               Caption         =   "Venta"
               Height          =   195
               Index           =   1
               Left            =   1755
               TabIndex        =   1
               Top             =   315
               Width           =   1110
            End
         End
         Begin VB.CommandButton CmdBusProv 
            Height          =   240
            Left            =   2880
            Picture         =   "FrmRegPercepciones.frx":0264
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1230
            Width           =   240
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   2160
            Picture         =   "FrmRegPercepciones.frx":0396
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   2550
            Width           =   240
         End
         Begin VB.Frame Frame3 
            Height          =   2535
            Left            =   9915
            TabIndex        =   19
            Top             =   3660
            Width           =   1725
            Begin VB.CommandButton CmdAddSel 
               Caption         =   "&Seleccionar Documentos"
               Height          =   585
               Left            =   150
               TabIndex        =   12
               Top             =   1035
               Width           =   1410
            End
            Begin VB.CommandButton CmdAdd 
               Caption         =   "&Agregar Documento"
               Height          =   585
               Left            =   135
               TabIndex        =   11
               Top             =   270
               Width           =   1410
            End
            Begin VB.CommandButton CmdDel 
               Caption         =   "&Eliminar Documento"
               Height          =   585
               Left            =   150
               TabIndex        =   13
               Top             =   1800
               Width           =   1410
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2445
            Left            =   75
            TabIndex        =   10
            Top             =   3750
            Width           =   9780
            _cx             =   17251
            _cy             =   4313
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
            ForeColorSel    =   16777215
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRegPercepciones.frx":04C8
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   1530
            TabIndex        =   3
            Top             =   1530
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
         End
         Begin VB.TextBox TxtRucPro 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   2
            Text            =   "TxtRucPro"
            Top             =   1200
            Width           =   1620
         End
         Begin VB.TextBox TxtMoneda 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   7
            Text            =   "TxtMoneda"
            Top             =   2505
            Width           =   900
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   150
            TabIndex        =   34
            Top             =   6195
            Width           =   11505
            Begin VB.TextBox TxtImpCob 
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
               Height          =   290
               Left            =   7170
               Locked          =   -1  'True
               TabIndex        =   38
               Text            =   "TxtImpCob"
               Top             =   195
               Width           =   1080
            End
            Begin VB.TextBox TxtImpRet 
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
               Height          =   290
               Left            =   6315
               Locked          =   -1  'True
               TabIndex        =   36
               Text            =   "TxtImpRet"
               Top             =   195
               Width           =   870
            End
            Begin VB.TextBox TxtImporte 
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
               Height          =   290
               Left            =   5340
               Locked          =   -1  'True
               TabIndex        =   35
               Text            =   "TxtImporte"
               Top             =   195
               Width           =   990
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total ==>"
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
               Left            =   4350
               TabIndex        =   37
               Top             =   225
               Width           =   825
            End
         End
         Begin VB.TextBox TxtIdDoc 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   4
            Text            =   "TxtIdDoc"
            Top             =   1860
            Width           =   900
         End
         Begin VB.TextBox TxtIdPer 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   8
            Text            =   "TxtIdPer"
            Top             =   2835
            Width           =   900
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
            Left            =   9420
            TabIndex        =   52
            Top             =   150
            Width           =   2250
         End
         Begin VB.Label LblTasa 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTasa"
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
            Height          =   300
            Left            =   8295
            TabIndex        =   47
            Top             =   2850
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tasa"
            Height          =   195
            Left            =   7785
            TabIndex        =   46
            Top             =   2925
            Width           =   360
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   4
            Left            =   165
            TabIndex        =   45
            Top             =   3255
            Width           =   405
         End
         Begin VB.Label LblPercepcion 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblPercepcion"
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
            Left            =   2445
            TabIndex        =   44
            Top             =   2850
            Width           =   4020
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Percepción"
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   43
            Top             =   2925
            Width           =   810
         End
         Begin VB.Label LblTipCam2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio"
            Height          =   195
            Left            =   7080
            TabIndex        =   42
            Top             =   2598
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
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   8295
            TabIndex        =   41
            Top             =   2520
            Width           =   1080
         End
         Begin VB.Label LblDocumento 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDocumento"
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
            Left            =   2445
            TabIndex        =   40
            Top             =   1860
            Width           =   4020
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   39
            Top             =   1944
            Width           =   825
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            Height          =   75
            Left            =   2475
            Top             =   2310
            Width           =   135
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   33
            Top             =   2271
            Width           =   1050
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Percepción"
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
            TabIndex        =   32
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblProveedor 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProveedor"
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
            Left            =   3165
            TabIndex        =   31
            Top             =   1200
            Width           =   6195
         End
         Begin VB.Label LblTitulo 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   165
            TabIndex        =   30
            Top             =   1290
            Width           =   735
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emisión"
            Height          =   195
            Index           =   3
            Left            =   165
            TabIndex        =   29
            Top             =   1617
            Width           =   1260
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   6
            Left            =   165
            TabIndex        =   28
            Top             =   2598
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
            Left            =   2445
            TabIndex        =   27
            Top             =   2520
            Width           =   4020
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   4560
            TabIndex        =   26
            Top             =   945
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label LblTituloDoc 
            AutoSize        =   -1  'True
            Caption         =   "LblTituloDoc"
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
            Left            =   165
            TabIndex        =   25
            Top             =   3525
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   45
         TabIndex        =   15
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   16
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11324
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Num.Reg."
            Columns(0).DataField=   "registro"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tip. Mov."
            Columns(1).DataField=   "tipmov"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº R.U.C."
            Columns(2).DataField=   "numruc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Proveedor / Cliente"
            Columns(3).DataField=   "nombre"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "T.D."
            Columns(4).DataField=   "docabrev"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fch Emi."
            Columns(5).DataField=   "fchdoc1"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Nº Documento"
            Columns(6).DataField=   "numdoc1"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "M"
            Columns(7).DataField=   "monabrev"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Imp. Perp."
            Columns(8).DataField=   "imptotper1"
            Columns(8).NumberFormat=   "0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Saldo"
            Columns(9).DataField=   "impsal1"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1535"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1455"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1349"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1270"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2249"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2170"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=4921"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=4842"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=953"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=873"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=1455"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1376"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=2646"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2566"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=635"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=556"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=513"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=1746"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1667"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(55)=   "Column(9).Width=1746"
            Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1667"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=74,.parent=13,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
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
         Begin VB.Label lblperiodo 
            AutoSize        =   -1  'True
            Caption         =   "lblperiodo"
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
            Index           =   0
            Left            =   9450
            TabIndex        =   49
            Top             =   75
            Width           =   1365
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Registro de Percepciones"
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
            TabIndex        =   17
            Top             =   30
            Width           =   11610
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
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
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6915
         Top             =   60
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
               Picture         =   "FrmRegPercepciones.frx":063C
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegPercepciones.frx":0B80
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegPercepciones.frx":0F12
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegPercepciones.frx":1096
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegPercepciones.frx":14EA
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegPercepciones.frx":1602
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegPercepciones.frx":1B46
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegPercepciones.frx":208A
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegPercepciones.frx":219E
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegPercepciones.frx":22B2
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegPercepciones.frx":2706
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegPercepciones.frx":2872
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegPercepciones.frx":2DBA
               Key             =   "IMG12"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "&Agregar Documento"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "&Seleccionar Documento"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "&Eliminar Documento"
      End
   End
End
Attribute VB_Name = "FrmRegPercepciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim RstPer As New ADODB.Recordset
Dim ValTipCam As Double
Dim xCuenPer As Integer
Dim Agregando As Boolean
'------------------
Dim xHorIni As Date

Dim mIdRegistro& '--identificador del registro
Dim mMesActivo As Integer '--indica el mes activo
Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta

Dim fCierrePeriodo As Boolean '--indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To 15
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Cancelar()
    Label5.Caption = "Detalle de la Percepcion"
    Bloquea
    ActivaTool
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
    QueHace = 3
End Sub

Private Sub CmdAdd_Click()
    If QueHace = 3 Then Exit Sub
    pRegistroAdd False
End Sub

Private Sub CmdAddSel_Click()
    If QueHace = 3 Then Exit Sub
    pRegistroAdd True
End Sub

Private Sub CmdBusDoc_Click()
    If QueHace = 3 Then Exit Sub
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Documento":    xCampos2(0, 1) = "descripcion":    xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Codigo":       xCampos2(1, 1) = "id":             xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"

    xform.SqlCad = "SELECT mae_documento.* From mae_documento WHERE (((mae_documento.id)=40 Or (mae_documento.id)=41))"

    xform.Titulo = "Buscando Tipo de Documento"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtIdDoc.Text = xRs("id")
        LblDocumento.Caption = NulosC(xRs("descripcion"))
        TxtNumSer.SetFocus
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Moneda":     xCampos2(0, 1) = "descripcion":    xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Codigo":      xCampos2(1, 1) = "id":          xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"

    xform.SqlCad = "SELECT * FROM mae_moneda"
    xform.Titulo = "Buscando Monedas"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtMoneda.Text = NulosC(xRs("id"))
        LblMoneda.Caption = NulosC(xRs("descripcion"))
        TxtIdPer.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusProv_Click()
    If QueHace = 3 Then Exit Sub
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    If Opt_Tipo(0).Value = True Then
        xCampos2(0, 0) = "Proveedor":   xCampos2(0, 1) = "nombre":       xCampos2(0, 2) = "6000":         xCampos2(0, 3) = "C"
        xform.SqlCad = "SELECT mae_prov.* From mae_prov where mae_prov.ageper = -1 and mae_prov.id<>0 ORDER BY mae_prov.nombre"
        xform.Titulo = "Buscando Proveedores"
    Else
        xCampos2(0, 0) = "Cliente":   xCampos2(0, 1) = "nombre":       xCampos2(0, 2) = "6000":         xCampos2(0, 3) = "C"
        xform.SqlCad = "SELECT mae_cliente.* From mae_cliente where mae_cliente.id<>0 ORDER BY mae_cliente.nombre"
        xform.Titulo = "Buscando Clientes"
    End If
    xCampos2(1, 0) = "Nº R.U.C.":   xCampos2(1, 1) = "numruc":       xCampos2(1, 2) = "1500":         xCampos2(1, 3) = "C"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtRucPro.Text = NulosC(xRs("numruc"))
        LblProveedor.Caption = NulosC(xRs("nombre"))
        LblIdProveedor.Caption = NulosC(xRs("id"))
        TxtFchEmi.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDel_Click()
    If QueHace = 3 Then Exit Sub
    If Fg1.Row < 1 Then Exit Sub
    If Fg1.Rows < 1 Then Exit Sub
    Fg1.RemoveItem (Fg1.Row)
    HallarTotales
End Sub

Private Sub CmdIdPer_Click()
    If QueHace = 3 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(5, 4) As String
    Dim nSQL As String
    
    xCampos2(0, 0) = "Descripcion":     xCampos2(0, 1) = "descripcion": xCampos2(0, 2) = "3200":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Tasa":            xCampos2(1, 1) = "tasa1":       xCampos2(1, 2) = "800":          xCampos2(1, 3) = "C"
    xCampos2(2, 0) = "Cta Numero.":     xCampos2(2, 1) = "ctanum":      xCampos2(2, 2) = "1200":         xCampos2(2, 3) = "C"
    xCampos2(3, 0) = "Cta Descripción": xCampos2(3, 1) = "ctadesc":     xCampos2(3, 2) = "2500":         xCampos2(3, 3) = "C"
    xCampos2(4, 0) = "Id":              xCampos2(4, 1) = "id":          xCampos2(4, 2) = "500":          xCampos2(4, 3) = "N"
    
    If Opt_Tipo(0).Value = True Then
        nSQL = "SELECT mae_percepcion.id, mae_percepcion.descripcion, mae_percepcion.tasa,mae_percepcion.idcuencom, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, Format([mae_percepcion].[tasa],'0.00') & '%' AS tasa1 " _
                & " FROM mae_percepcion LEFT JOIN con_planctas AS con_planctas ON mae_percepcion.idcuencom = con_planctas.id; "
    Else
        nSQL = "SELECT mae_percepcion.id, mae_percepcion.descripcion, mae_percepcion.tasa, mae_percepcion.idcuenven, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, Format([mae_percepcion].[tasa],'0.00') & '%' AS tasa1 " _
                & " FROM mae_percepcion LEFT JOIN con_planctas ON mae_percepcion.idcuenven = con_planctas.id; "
    End If
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos2(), "Buscando Percepciones", "descripcion", "descripcion", Principio
    
    If xRs.State = 1 Then
        TxtIdPer.Text = xRs("id")
        LblPercepcion.Caption = NulosC(xRs("descripcion"))
        LblTasa.Caption = Format(xRs("tasa"), "0.00")
        If Opt_Tipo(0).Value = True Then
            xCuenPer = NulosN(xRs("idcuencom"))
        Else
            xCuenPer = NulosN(xRs("idcuenven"))
        End If
        TxtGlosa.SetFocus
    End If
        
    Set xRs = Nothing
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstPer
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstPer.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstPer("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'    On Error GoTo error
'    If Col <> 1 Then Exit Sub
'    If QueHace = 3 Then Exit Sub
'    If NulosC(TxtRucPro.Text) = "" Then
'        MsgBox "Seleccione el " + LblTitulo.Caption, vbExclamation, xTitulo
'        CmdBusProv.SetFocus
'        Exit Sub
'    End If
'    If NulosN(TxtMoneda.Text) = 0 Then
'        MsgBox "Seleccione la Moneda", vbExclamation, xTitulo
'        CmdBusMon.SetFocus
'        Exit Sub
'    End If
'    Dim xCampos(5, 5) As String
'    Dim xRs As New ADODB.Recordset
'    Dim nSQLId As String '--almacenara los codigos de documentos ya seleccionados
'    Dim nSQL As String
'    Dim nTitulo As String
'    Dim nSQLNotInDocumentos As String
'
'    xCampos(0, 0) = "Num.Reg.":      xCampos(0, 1) = "registro":       xCampos(0, 2) = "1200":       xCampos(0, 3) = "C":    xCampos(0, 4) = "N"
'    xCampos(1, 0) = "Documento":      xCampos(1, 1) = "abrev":       xCampos(1, 2) = "1200":         xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
'    xCampos(2, 0) = "Nº Documento":   xCampos(2, 1) = "numdoc":      xCampos(2, 2) = "2000":         xCampos(2, 3) = "C":    xCampos(2, 4) = "S"
'    xCampos(3, 0) = "Fch. Emision":   xCampos(3, 1) = "fchdoc":      xCampos(3, 2) = "1200":         xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
'    xCampos(4, 0) = "Importe":        xCampos(4, 1) = "imptot":      xCampos(4, 2) = "1200":         xCampos(4, 3) = "N":    xCampos(4, 4) = "N"
'
'    '*************************************************************
'    nSQLNotInDocumentos = vbCr + " AND com_compras.id NOT IN (SELECT con_percepciondet1.iddoc " _
'        & " FROM con_percepcion AS con_percepcion1 INNER JOIN con_percepciondet AS con_percepciondet1 ON con_percepcion1.id = con_percepciondet1.id " _
'        & " WHERE (((con_percepcion1.tipo)=" + IIf(Opt_Tipo(0).Value = True, "1", "2") + ") AND ((con_percepcion1.idcli)=" & NulosN(LblIdProveedor.Caption) & "));)"
'        '--con_percepcion.tipo:=: 1::COMPRA, 2::VENTA
'    '*************************************************************
'    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 8, "com_compras.id", " NOT IN ")
'    '*************************************************************
'    If Opt_Tipo(0).Value = True Then
'        nSQL = "SELECT com_compras.id, mae_documento.abrev, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.fchdoc, com_compras.imptot, mae_documentocta.idmon, mae_documentocta.tipope, mae_documentocta.idcuen, Left([com_compras].[numreg],2) & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([com_compras].[numreg],3) AS registro, com_compras.idmon AS idmondoc " _
'            + vbCr + " FROM ((mae_documento RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) LEFT JOIN mae_documentocta ON mae_documento.id = mae_documentocta.iddoc) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id " _
'            + vbCr + " WHERE (((com_compras.idpro)=" & Val(LblIdProveedor.Caption) & ") AND ((mae_documentocta.idmon)=" & Val(TxtMoneda.Text) & ") AND ((mae_documentocta.tipope)=0)) "
'
'    Else
'        nSQL = "SELECT vta_ventas.id, mae_documento.abrev, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.fchdoc, vta_ventas.imptotdoc AS imptot, mae_documentocta.idcuen, Left([vta_ventas].[numreg],2) & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([vta_ventas].[numreg],3) AS registro, com_compras.idmon AS idmondoc " _
'            + vbCr + " FROM ((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_documentocta ON mae_documento.id = mae_documentocta.iddoc) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
'            + vbCr + " WHERE (((mae_documentocta.idmon)=" & Val(TxtMoneda.Text) & ") AND ((mae_documentocta.tipope)=-1) AND ((vta_ventas.idcli)=" & Val(LblIdProveedor.Caption) & ")) "
'            '*************************************************************
'            nSQLId = Replace(nSQLId, "com_compras.id", "vta_ventas.id")
'            '*************************************************************
'            nSQLNotInDocumentos = Replace(nSQLNotInDocumentos, "com_compras.id", "vta_ventas.id")
'            '*************************************************************
'    End If
'
'    nTitulo = "Buscando Documento del " + LblTitulo.Caption + ": " + LblProveedor.Caption
'    '*************************************************************
'    nSQL = nSQL + IIf(nSQLId = "", "", " AND " + nSQLId) + nSQLNotInDocumentos
'    '*************************************************************
'    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "numdoc", "numdoc", CualquierParte
'
'    If xRs.State = 1 Then
'        Agregando = True
'        Fg1.TextMatrix(Fg1.Row, 1) = NulosC(xRs("registro"))
'        Fg1.TextMatrix(Fg1.Row, 2) = NulosC(xRs("numdoc"))
'        Fg1.TextMatrix(Fg1.Row, 3) = NulosC(xRs("abrev"))
'        Fg1.TextMatrix(Fg1.Row, 4) = xRs("fchdoc")
'        'si el iporte es en dolares
'        Fg1.TextMatrix(Fg1.Row, 5) = Format(NulosN(xRs("imptot")), FORMAT_MONTO)
'
'        '-----------
'        If NulosN(Fg1.TextMatrix(Fg1.Row, 6)) = 0 Then
'            Fg1.TextMatrix(Fg1.Row, 6) = Format(NulosN(LblTasa.Caption), "0.00")
'        End If
'        Fg1.TextMatrix(Fg1.Row, 7) = NulosN(Fg1.TextMatrix(Fg1.Row, 5)) * NulosN(Fg1.TextMatrix(Fg1.Row, 6)) / 100
'        Fg1.TextMatrix(Fg1.Row, 8) = NulosN(Fg1.TextMatrix(Fg1.Row, 5)) + NulosN(Fg1.TextMatrix(Fg1.Row, 7))
'        Fg1.TextMatrix(Fg1.Row, 8) = Format(Fg1.TextMatrix(Fg1.Row, 8), FORMAT_MONTO)
'        '-----------
'        Fg1.TextMatrix(Fg1.Row, 9) = NulosN(xRs("id"))
'        Fg1.TextMatrix(Fg1.Row, 10) = NulosN(xRs("idcuen"))
'
'        Agregando = False
'        HallarTotales
'    End If
'    Set xRs = Nothing
'    Fg1.Row = Row: Fg1.Col = Col:    Fg1.SetFocus
'    Exit Sub
'error:
'    Set xRs = Nothing
'    HallarTotales
'    Agregando = False
'    SHOW_ERROR Me.Name, "Fg1_CellButtonClick"
End Sub

Private Sub fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then '--insert
        CmdAdd_Click
    End If
    
    If KeyCode = 46 Then '--delete
        CmdDel_Click
    End If
End Sub

Private Sub HallarTotales()
    
    TxtImporte.Text = Format(GRID_SUMAR_COL(Fg1, 6), FORMAT_MONTO) '--IMPORTE
    TxtImpRet.Text = Format(NulosN(LblTasa.Caption), FORMAT_MONTO) '--IMPORTE PERCEP
    TxtImpCob.Text = Format(GRID_SUMAR_COL(Fg1, 8), FORMAT_MONTO) '--TOTAL COBRADO

End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Col = 6 Then
'        Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), "0.00") '--porcentaje
'        Fg1.TextMatrix(Fg1.Row, 8) = NulosN(Fg1.TextMatrix(Fg1.Row, 6)) * NulosN(Fg1.TextMatrix(Fg1.Row, 7)) / 100
'        Fg1.TextMatrix(Fg1.Row, 9) = NulosN(Fg1.TextMatrix(Fg1.Row, 6)) + NulosN(Fg1.TextMatrix(Fg1.Row, 7))
'        Fg1.TextMatrix(Fg1.Row, 9) = Format(Fg1.TextMatrix(Fg1.Row, 9), FORMAT_MONTO)
        
        '--------------
        Fg1.TextMatrix(Fg1.Row, 8) = NulosN(Fg1.TextMatrix(Fg1.Row, 6)) * NulosN(Fg1.TextMatrix(Fg1.Row, 7)) / 100
        Fg1.TextMatrix(Fg1.Row, 9) = NulosN(Fg1.TextMatrix(Fg1.Row, 6)) + NulosN(Fg1.TextMatrix(Fg1.Row, 8))
        Fg1.TextMatrix(Fg1.Row, 9) = Format(Fg1.TextMatrix(Fg1.Row, 9), FORMAT_MONTO)
        
    End If
    
    If Col = 7 Then
'        Fg1.TextMatrix(Fg1.Row, 6) = Format(Fg1.TextMatrix(Fg1.Row, 7), "0.0000")
'        Fg1.TextMatrix(Fg1.Row, 7) = NulosN(Fg1.TextMatrix(Fg1.Row, 5)) + NulosN(Fg1.TextMatrix(Fg1.Row, 7))
'        Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 8), FORMAT_MONTO)
        '--------------
        Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), "0.0000")
        Fg1.TextMatrix(Fg1.Row, 8) = NulosN(Fg1.TextMatrix(Fg1.Row, 6)) * NulosN(Fg1.TextMatrix(Fg1.Row, 7)) / 100
        Fg1.TextMatrix(Fg1.Row, 9) = NulosN(Fg1.TextMatrix(Fg1.Row, 6)) + NulosN(Fg1.TextMatrix(Fg1.Row, 8))
        Fg1.TextMatrix(Fg1.Row, 9) = Format(Fg1.TextMatrix(Fg1.Row, 9), FORMAT_MONTO)
        
        
    End If
    HallarTotales
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone: Exit Sub
    End If
    If Fg1.Col = 6 Or Fg1.Col = 2 Or Fg1.Col = 7 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then PopupMenu Menu1
    End If
End Sub

Private Sub Form_Activate()
    
    If SeEjecuto = False Then
   
    SeEjecuto = True
    
    '--Almacenar temporalmente el codigo del menu
    IdMenuActivo = xIdMenu
    
    mMesActivo = xMes
    
    OpcionesPeriodo
    
    End If

End Sub

Private Sub Form_Load()
    QueHace = 3
    
    Dg1.Columns("fchdoc1").NumberFormat = FORMAT_DATE
    Dg1.Columns("imptotper1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("impsal1").NumberFormat = FORMAT_MONTO
    
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
   
    Dg1.Columns("fchdoc1").NumberFormat = FORMAT_DATE:
    Fg1.ColFormat(3) = FORMAT_DATE
    Fg1.SelectionMode = flexSelectionByRow
    Agregando = True
End Sub

Function Grabar() As Boolean
    If AnoTra = "" Then
        MsgBox "No hay año de trabajo", vbExclamation, xTitulo
        Exit Function
    End If

    If TxtRucPro.Text = "" Then
        MsgBox "No ha especificado el " + LblTitulo.Caption, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtRucPro.SetFocus
        Exit Function
    End If

    If IsDate(TxtFchEmi.Valor) = False Then
        MsgBox "No ha especificado la fecha de emisión de la percepción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchEmi.Valor = ""
        TxtFchEmi.SetFocus
        Exit Function
    End If

    If TxtNumSer.Text = "" Then
        MsgBox "No ha especificado el número de serie para el documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        Exit Function
    End If

    If TxtNumDoc.Text = "" Then
        MsgBox "No ha especificado el número de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        Exit Function
    End If

    If NulosN(TxtMoneda.Text) = 0 Then
        MsgBox "No ha especificado la moneda para el documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtMoneda.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtIdPer.Text) = 0 Then
        MsgBox "No ha especificado la percepción que se esta aplicando", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdPer.SetFocus
        Exit Function
    End If
    '---------------------------------------------------
    If Fg1.Rows <= 1 Then
        MsgBox "Ingrese los " + LblTituloDoc.Caption, vbExclamation, xTitulo
        CmdAddSel.SetFocus
        Exit Function
    End If
    '--VALIDAR EL INGRESO DE LOS DATOS
    Dim Q_ROW  As Long
    Dim Q_COL As Long '--COLUMNA A POSICIONAR SI FALTAN DATOS
    Q_COL = -1
    For Q_ROW = 1 To Fg1.Rows - 1
        If Fg1.TextMatrix(Q_ROW, 6) = "" Then
            MsgBox "Ingrese La tasa de la Percepción", vbExclamation, xTitulo
            Q_COL = 6:          Exit For
        ElseIf NulosN(Fg1.TextMatrix(Q_ROW, 7)) = 0 Then
            MsgBox "Ingrese un valor para el Importe de Percepción", vbExclamation, xTitulo
            Q_COL = 7:          Exit For
        End If
    Next Q_ROW
    If Q_COL <> -1 Then
        Agregando = True:  Fg1.Row = Q_ROW: Fg1.Col = Q_COL: Agregando = False
        Exit Function
    End If
    '-------------------------------------------------------
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modficar") + " la Percepción", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
''    Dim RstDia As New ADODB.Recordset
    
    Dim xId As Double
    Dim A&
    Dim xNumAsiento As String

On Error GoTo LaCague

    xCon.BeginTrans
   
    If QueHace = 1 Then
        xId = HallaCodigoTabla("con_percepcion", xCon, "id")
'''        xNumAsiento = NuevoNumAsiento(4, mMesActivo, xCon)
        
        RST_Busq RstCab, "SELECT TOP 1 * FROM con_percepcion", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstPer("id")
        RST_Busq RstCab, "SELECT * FROM con_percepcion WHERE id = " & xId & "", xCon
        xCon.Execute "DELETE * FROM con_percepciondet WHERE id = " & xId & ""
        
'''        xNumAsiento = DevuelveNumAsiento(4, RstPer("id"), mMesActivo, xCon)
'''
'''        If xNumAsiento = "" Then xNumAsiento = NuevoNumAsiento(4, mMesActivo, xCon)
'''
'''        xCon.Execute "DELETE * FROM con_diario WHERE ((idmes = " & mMesActivo & ") and (idlib = 4) AND (idmov = " & xId & "))"
        
    End If
    '------------------------------------------------
    RST_Busq RstDet, "SELECT TOP 1 * FROM con_percepciondet", xCon
'''    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    
    mIdRegistro = xId
    '------------------------------------------------
    RstCab("ano") = NulosN(AnoTra)
    RstCab("idmes") = mMesActivo
    RstCab("idlib") = 4
'''    RstCab("numreg") = Format(mMesActivo, "00") + xNumAsiento
    '---------
    If mMesActivo <> 0 And mMesActivo <> 13 Then
        RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    End If
    
    If Opt_Tipo(0).Value = True Then RstCab("tipo") = 1  'se esta registrando una compra
    If Opt_Tipo(1).Value = True Then RstCab("tipo") = 2  'se esta registrando una venta
    
    RstCab("idcli") = NulosN(LblIdProveedor.Caption)
    RstCab("tipdoc") = NulosN(TxtIdDoc.Text)
    RstCab("numser") = TxtNumSer.Text
    RstCab("numdoc") = TxtNumDoc.Text
    
    RstCab("idmon") = NulosN(TxtMoneda.Text)
    RstCab("idper") = NulosN(TxtIdPer.Text)
    RstCab("fchdoc") = TxtFchEmi.Valor
    RstCab("imptotper") = NulosN(TxtImpCob.Text)
    RstCab("impsal") = NulosN(TxtImpCob.Text)
    
    RstCab("glosa") = TxtGlosa.Text
    
    RstCab.Update
        
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("id") = xId
        RstDet("iddoc") = NulosN(Fg1.TextMatrix(A, 10))
        RstDet("porper") = NulosN(Fg1.TextMatrix(A, 7))
        RstDet("impper") = NulosN(Fg1.TextMatrix(A, 8))
        
        RstDet.Update
    Next A
    '---------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------
''''    'grabamos el la cuenta DEBE de la percepcion
''''    RstDia.AddNew
''''    RstDia("año") = NulosN(AnoTra)
''''    RstDia("idmes") = mMesActivo
''''    RstDia("idlib") = 4
''''    RstDia("idmov") = xId
''''    RstDia("numasi") = xNumAsiento
''''    RstDia("tc") = NulosN(LblTipoCambio.Caption)
''''    RstDia("idcue") = xCuenPer
''''    If mMesActivo = 0 Then
''''        RstDia("fchasi") = CDate("01/01/" + AnoTra)
''''    ElseIf mMesActivo = 13 Then
''''        RstDia("fchasi") = CDate("31/12/" + AnoTra)
''''    Else
''''        RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
''''    End If
''''    RstDia("fchdoc") = CDate(TxtFchEmi.Valor)
''''    If Opt_Tipo(0).Value = True Then '--COMPRA
''''        If TxtMoneda.Text = "1" Then
''''            RstDia("impdebsol") = NulosN(TxtImpCob.Text)
''''            RstDia("impdebdol") = 0
''''        Else
''''            RstDia("impdebsol") = NulosN(TxtImpCob.Text) * NulosN(LblTipoCambio.Caption)
''''            RstDia("impdebdol") = NulosN(TxtImpCob.Text)
''''        End If
''''    Else '--VENTA
''''        If TxtMoneda.Text = "1" Then
''''            RstDia("imphabsol") = NulosN(TxtImpCob.Text)
''''            RstDia("imphabdol") = 0
''''        Else
''''            RstDia("imphabsol") = NulosN(TxtImpCob.Text) * NulosN(LblTipoCambio.Caption)
''''            RstDia("imphabdol") = NulosN(TxtImpCob.Text)
''''        End If
''''    End If
''''    RstDia.Update
''''
''''    'GRABAMOS LA CUENTA HABER CON EL CODIGO DE LA CUENTA DE LOS DOCUMENTOS INVOLUCRADOS EN LA PERCEPCION
''''    Dim xIdCuen As Integer
''''    Dim xTotal As Double
''''    Dim Cambio As Boolean
''''
''''    A = 1
''''    xIdCuen = NulosN(Fg1.TextMatrix(A, 11))
''''    For A = 1 To Fg1.Rows - 1
''''        If xIdCuen = NulosN(Fg1.TextMatrix(A, 11)) Then
''''            xTotal = xTotal + NulosN(Fg1.TextMatrix(A, 8))
''''            Cambio = False
''''        Else
''''            RstDia.AddNew
''''            RstDia("idmes") = mMesActivo               'LLAVE - CODIGO DEL MES
''''            RstDia("idlib") = 4                  'LLAVE - CODIGO DEL LIBRO
''''            RstDia("idmov") = xId                'LLAVE - CODIGO DEL MOVIMIENTO
''''            RstDia("numasi") = xNumAsiento       'LLAVE - NUMERO DE ASIENTO
''''            RstDia("tc") = NulosN(LblTipoCambio.Caption)
''''            RstDia("idcue") = xIdCuen
''''
''''            If Opt_Tipo(0).Value = True Then '--COMPRA
''''                If TxtMoneda.Text = "1" Then
''''                    RstDia("impdebsol") = NulosN(TxtImpCob.Text)
''''                    RstDia("impdebdol") = 0
''''                Else
''''                    RstDia("impdebsol") = NulosN(TxtImpCob.Text) * NulosN(LblTipoCambio.Caption)
''''                    RstDia("impdebdol") = NulosN(TxtImpCob.Text)
''''                End If
''''            Else '--VENTA
''''                If TxtMoneda.Text = "1" Then
''''                    RstDia("imphabsol") = NulosN(TxtImpCob.Text)
''''                    RstDia("imphabdol") = 0
''''                Else
''''                    RstDia("imphabsol") = NulosN(TxtImpCob.Text) * NulosN(LblTipoCambio.Caption)
''''                    RstDia("imphabdol") = NulosN(TxtImpCob.Text)
''''                End If
''''            End If
''''            RstDia.Update
''''            Cambio = True
''''        End If
''''    Next A
''''
''''    If Cambio = False Then
''''        RstDia.AddNew
''''        RstDia("año") = NulosN(AnoTra)
''''        RstDia("idmes") = mMesActivo
''''        RstDia("idlib") = 4
''''        RstDia("idmov") = xId
''''        RstDia("numasi") = xNumAsiento
''''        RstDia("tc") = NulosN(LblTipoCambio.Caption)
''''        RstDia("idcue") = xIdCuen
''''        If mMesActivo = 0 Then
''''            RstDia("fchasi") = CDate("01/01/" + AnoTra)
''''        ElseIf mMesActivo = 13 Then
''''            RstDia("fchasi") = CDate("31/12/" + AnoTra)
''''        Else
''''            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
''''        End If
''''        RstDia("fchdoc") = CDate(TxtFchEmi.Valor)
''''        If Opt_Tipo(1).Value = True Then '--VENTA
''''            If TxtMoneda.Text = "1" Then
''''                RstDia("impdebsol") = NulosN(TxtImpCob.Text)
''''                RstDia("impdebdol") = 0
''''            Else
''''                RstDia("impdebsol") = NulosN(TxtImpCob.Text) * NulosN(LblTipoCambio.Caption)
''''                RstDia("impdebdol") = NulosN(TxtImpCob.Text)
''''            End If
''''        Else    '--COMPRA
''''            If TxtMoneda.Text = "1" Then
''''                RstDia("imphabsol") = NulosN(TxtImpCob.Text)
''''                RstDia("imphabdol") = 0
''''            Else
''''                RstDia("imphabsol") = NulosN(TxtImpCob.Text) * NulosN(LblTipoCambio.Caption)
''''                RstDia("imphabdol") = NulosN(TxtImpCob.Text)
''''            End If
''''        End If
''''
''''        RstDia.Update
''''    End If
    
    '---generar asiento
    xNumAsiento = GenerarAsiento(xCon, 4, xId, AnoTra, mMesActivo, 0)
    If xNumAsiento = "" Then GoTo LaCague
    '----------------------------------------------------------------------------------
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId

    
    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
'''    Set RstDia = Nothing
    

    MsgBox "La Percepción se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito" + vbCr + "Num.Reg.: " & xNumAsiento, vbInformation, xTitulo

    Grabar = True
    Exit Function
    
LaCague:
    'Resume
    Set RstCab = Nothing
    Set RstDet = Nothing
'''    Set RstDia = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar la Percepción por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando una Percepción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Menu1_1_Click()
    CmdAdd_Click
End Sub

Private Sub Menu1_2_Click()
    CmdAddSel_Click
End Sub

Private Sub Menu1_4_Click()
    CmdDel_Click
End Sub

Private Sub Opt_Tipo_Click(Index As Integer)
    If Index = 0 Then '--PROVEEDOR
        LblTitulo.Caption = "Proveedor"
        LblTituloDoc.Caption = "Documentos de Compra"
    Else
        LblTitulo.Caption = "Cliente"
        LblTituloDoc.Caption = "Documentos de Venta"
    End If
    TxtRucPro.Text = ""
    
End Sub

Private Sub Opt_Tipo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TxtRucPro.SetFocus
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If RstPer.RecordCount = 0 And QueHace = 3 Then
            Cancel = 1
            Exit Sub
        End If
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstPer.Requery
            Dg1.Refresh
            '-------
            If RstPer.RecordCount <> 0 Then
                RstPer.MoveFirst
                RstPer.Find "id=" & mIdRegistro
                If RstPer.EOF = True Then RstPer.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    If Button.Index = 8 Then Filtrar
    If Button.Index = 9 Then
        If RstPer.State = 0 Then Exit Sub
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstPer.Filter = adFilterNone
        RstPer.Requery
        Dg1.Refresh
    End If
    
    If Button.Index = 10 Then CambiarMes
    If Button.Index = 11 Then Buscar
    If Button.Index = 13 Then pExportar
    
    If Button.Index = 16 Then
        Set RstPer = Nothing
        Unload Me
    End If
End Sub

Sub Filtrar()
    'Dim xform As New eps_librerias.FormFiltrar
    Dim xform As New eps_librerias.FormFiltrar
    
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(5, 4) As String
    
    xCampos(0, 0) = "Num.Reg.":             xCampos(0, 1) = "numreg":        xCampos(0, 2) = "C":         xCampos(0, 3) = "4200"
    xCampos(1, 0) = "Tipo":                 xCampos(1, 1) = "tipo2":         xCampos(1, 2) = "C":         xCampos(1, 3) = "4200"
    xCampos(2, 0) = "Cliente / Proveedor":  xCampos(2, 1) = "nombre":        xCampos(2, 2) = "C":         xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Fch. Emision":         xCampos(3, 1) = "fchdoc":        xCampos(3, 2) = "F":         xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Nº R.U.C.":            xCampos(4, 1) = "numruc":        xCampos(4, 2) = "C":         xCampos(4, 3) = "1500"
    xCampos(5, 0) = "Nº Documento":         xCampos(5, 1) = "numdoc1":       xCampos(5, 2) = "C":         xCampos(5, 3) = "1500"
    
    
    Set xform.Coneccion = xCon
    Set xform.Rst = RstPer       'recorset que llena el grid
    Set RstPer = xform.FiltrarReg(xCampos)
    Set Dg1.DataSource = RstPer
    Dg1.Refresh
End Sub

Sub Modificar()
    If RstPer.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    QueHace = 2
    Label5.Caption = "Modificando Percepción"
    Bloquea
    ActivaTool
    
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    
    TabOne1.TabEnabled(0) = False
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    Agregando = False
    
    Fg1.ColFormat(3) = FORMAT_DATE
    xHorIni = Time
    TxtRucPro.SetFocus
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    If RstPer.RecordCount = 0 Or RstPer.EOF = True Or RstPer.BOF = True Then
        MsgBox "No hay registro para eliminar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    Rpta = MsgBox("Esta seguro de eliminar la percepción seleccionada", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        'ELIMINAMOS EL ASIENTO CONTABLE
        xCon.Execute "DELETE * FROM con_diario WHERE ((idmes = " & mMesActivo & ") and (idlib = 4) AND (idmov = " & RstPer("id") & "))"
        'ELIMINAMOS LA PERCEPCION
        xCon.Execute "DELETE * FROM con_percepciondet WHERE id = " & RstPer("id") & ""
        xCon.Execute "DELETE * FROM con_percepcion WHERE id = " & RstPer("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstPer("id") & " AND idform = " & IdMenuActivo
        
        
        RstPer.Requery
        Dg1.Refresh
        MsgBox "La percepción se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    
    TabOne1.CurrTab = 0
    
End Sub

Private Sub TxtFchEmi_Validate(Cancel As Boolean)
    If IsDate(TxtFchEmi.Valor) = True Then
        If NulosN(TxtMoneda.Text) = 0 Then
            LblTipoCambio.Caption = HallaTipoCambio(TxtFchEmi.Valor, 2, Venta, xCon)
        End If
    Else
        LblTipoCambio.Caption = ""
    End If
End Sub

Private Sub TxtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And CmdAdd.Enabled = True Then CmdAdd.SetFocus
End Sub

Private Sub TxtIdDoc_Change()
    If Trim(TxtIdDoc.Text) = "" Then LblDocumento.Caption = ""
End Sub

Private Sub TxtIdDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDoc_Click
    End If
End Sub

Private Sub TxtIdDoc_Validate(Cancel As Boolean)
    If NulosC(TxtIdDoc.Text) = "" Then Exit Sub
    LblDocumento.Caption = Busca_Codigo(TxtIdDoc.Text, "id", "descripcion", "mae_documento", "N", xCon)
    If NulosC(LblDocumento.Caption) = "" Then
        TxtIdDoc.Text = ""
    End If
End Sub

Private Sub TxtIdPer_Change()
    If Trim(TxtIdPer.Text) = "" Then LblPercepcion.Caption = ""
End Sub

Private Sub TxtIdPer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdPer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then CmdIdPer_Click
End Sub

Private Sub TxtIdPer_Validate(Cancel As Boolean)
    On Error GoTo error
    If QueHace = 3 Then Exit Sub
    If Trim(TxtIdPer.Text) = "" Then Exit Sub
    
    LblPercepcion.Caption = Busca_Codigo(NulosN(TxtIdPer.Text), "id", "descripcion", "mae_percepcion", "N", xCon)
    
    If LblPercepcion.Caption <> "" Then
        LblTasa.Caption = Busca_Codigo(NulosN(TxtIdPer.Text), "id", "tasa", "mae_percepcion", "N", xCon)
        LblTasa.Caption = Format(LblTasa.Caption, "0.00")
        
        If Opt_Tipo(0).Value = True Then
            xCuenPer = Busca_Codigo(NulosN(TxtIdPer.Text), "id", "idcuencom", "mae_percepcion", "N", xCon)
        Else
            xCuenPer = Busca_Codigo(NulosN(TxtIdPer.Text), "id", "idcuenven", "mae_percepcion", "N", xCon)
        End If
    End If
    Exit Sub
error:
    SHOW_ERROR Me.Name, "TxtIdPer_Validate"
End Sub

Private Sub TxtMoneda_Change()
    If QueHace = 3 Then Exit Sub
    If Trim(TxtMoneda.Text) = "" Then
        LblMoneda.Caption = ""
        Fg1.Rows = 1 '--limpiar la grilla, pues depende de la moneda
        HallarTotales
    End If
End Sub

Private Sub TxtMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtMoneda_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtMoneda_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtMoneda.Text) = "" Then Exit Sub
    LblMoneda.Caption = Busca_Codigo(TxtMoneda.Text, "id", "descripcion", "mae_moneda", "N", xCon)
    If NulosC(LblMoneda.Caption) = "" Then TxtMoneda.Text = ""
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
    If NulosC(TxtNumDoc.Text) <> "" Then
        TxtNumDoc.Text = Format(TxtNumDoc, "0000000000")
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    If NulosC(TxtNumSer.Text) <> "" Then
        TxtNumSer.Text = Format(TxtNumSer, "0000")
    End If
End Sub

Private Sub TxtRucPro_Change()
    If Trim(TxtRucPro) = "" Then
        LblProveedor.Caption = ""
        LblIdProveedor.Caption = ""
        Fg1.Rows = 1
        HallarTotales
    End If
End Sub

Private Sub TxtRucPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtRucPro_Validate True
        SendKeys vbTab
    End If
End Sub

Private Sub TxtRucPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub

Private Sub TxtRucPro_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If TxtRucPro.Text <> "" Then
        Dim Rst As New ADODB.Recordset
        If Opt_Tipo(0).Value = True Then
            RST_Busq Rst, "SELECT * FROM mae_prov WHERE numruc = '" & Trim(TxtRucPro.Text) & "'", xCon
            If Rst.RecordCount <> 0 Then
                Rst.MoveFirst
                TxtRucPro.Text = NulosC(Rst("numruc"))
                LblProveedor.Caption = NulosC(Rst("nombre"))
                LblIdProveedor.Caption = Rst("id")
            End If
        Else
            RST_Busq Rst, "SELECT * FROM mae_cliente WHERE numruc = '" & Trim(TxtRucPro.Text) & "'", xCon
            If Rst.RecordCount <> 0 Then
                Rst.MoveFirst
                TxtRucPro.Text = NulosC(Rst("numruc"))
                LblProveedor.Caption = NulosC(Rst("nombre"))
                LblIdProveedor.Caption = Rst("id")
            End If
        End If
        Set Rst = Nothing
    End If
End Sub

Sub Bloquea()
    TxtRucPro.Locked = Not TxtRucPro.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtMoneda.Locked = Not TxtMoneda.Locked
    TxtFchEmi.Locked = Not TxtFchEmi.Locked
    TxtIdDoc.Locked = Not TxtIdDoc.Locked
    TxtIdPer.Locked = Not TxtIdPer.Locked
    TxtGlosa.Locked = Not TxtGlosa.Locked
       
    Fra_Tipo.Enabled = Not Fra_Tipo.Enabled
End Sub

Sub Blanquea()

    LblTipoCambio.Caption = ""
    
    LblTasa.Caption = ""
    
    TxtRucPro.Text = ""
    LblProveedor.Caption = ""
    TxtFchEmi.Valor = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtMoneda.Text = ""
    TxtIdPer.Text = ""
    TxtIdDoc.Text = ""
    TxtGlosa.Text = ""
    LblIdProveedor.Caption = ""
    LblMoneda.Caption = ""
    
    LblDocumento.Caption = ""
    LblPercepcion.Caption = ""
    
    TxtImporte.Text = ""
    TxtImpRet.Text = ""
    TxtImpCob.Text = ""
    
    lblReg.Caption = ""
    
End Sub

Sub Nuevo()
    QueHace = 1
    Label5.Caption = "Agregando Percepción"
    Bloquea
    Blanquea
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Opt_Tipo(0).Value = True
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    Fg1.Editable = flexEDKbdMouse
    Fg1.Rows = 1
    Agregando = False
    LblTituloDoc.Caption = "Documentos de Compra"
    Fg1.ColFormat(3) = FORMAT_DATE
    xHorIni = Time
    Opt_Tipo(0).SetFocus
End Sub

Sub MuestraSegundoTab()

    Blanquea
    If RstPer.EOF = True Or RstPer.BOF = True Or RstPer.RecordCount = 0 Then Exit Sub
    
    lblReg.Caption = "Nº Reg. " & NulosC(RstPer("registro"))
    If RstPer("tipo") = 1 Then
        Opt_Tipo(0).Value = True
        LblTituloDoc.Caption = "Documentos de Compra"
    Else
        Opt_Tipo(1).Value = True
        LblTituloDoc.Caption = "Documentos de Venta"
    End If
    TxtRucPro.Text = NulosC(RstPer("numruc"))
    LblProveedor.Caption = NulosC(RstPer("nombre"))
    LblIdProveedor.Caption = NulosN(RstPer("idcli"))
    
    TxtFchEmi.Valor = RstPer("fchdoc")
    TxtFchEmi_Validate False
    
    TxtNumSer.Text = NulosC(RstPer("numser"))
    TxtNumDoc.Text = NulosC(RstPer("numdoc"))
    
    TxtIdDoc.Text = RstPer("tipdoc")
    LblDocumento.Caption = RstPer("docdesc")
    
    TxtIdPer.Text = NulosN(RstPer("idper"))
    LblPercepcion.Caption = NulosC(RstPer("percepdesc"))
    LblTasa.Caption = Format(NulosN(RstPer("perceptasa")), "0.00")
    
    TxtGlosa.Text = NulosC(RstPer("glosa"))
    
    TxtMoneda.Text = RstPer("monid")
    LblMoneda.Caption = NulosC(RstPer("mondesc"))
    
    If Opt_Tipo(0).Value = True Then
        'CUENTA CONTABLE DE LA RETENCION CUANDO SEA COMPRA
        xCuenPer = NulosN(RstPer("percepidcom"))
    Else
        'CUENTA CONTABLE DE LA RETENCION CUANDO SEA VENTA
        xCuenPer = NulosN(RstPer("percepidven"))
    End If
    
    
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Dim nSQL As String
    
    If Opt_Tipo(0).Value = True Then
        nSQL = "SELECT con_percepciondet.id, mae_documento.abrev, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.fchdoc, com_compras.imptot, con_percepciondet.porper, con_percepciondet.impper, con_percepciondet.iddoc, [com_compras]![imptot]+[con_percepciondet]![impper] AS impcob, mae_documentocta.idmon, mae_documentocta.tipope, mae_documentocta.idcuen,Left([com_compras].[numreg],2) & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([com_compras].[numreg],3) AS registro, com_compras.idmon AS idmondoc " _
            & " FROM ((mae_documento RIGHT JOIN (con_percepciondet LEFT JOIN com_compras ON con_percepciondet.iddoc = com_compras.id) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN mae_documentocta ON mae_documento.id = mae_documentocta.iddoc) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id " _
            & " WHERE (((con_percepciondet.id)=" & RstPer("id") & ") AND ((mae_documentocta.idmon)=" & Val(TxtMoneda.Text) & ") " _
            & " AND ((mae_documentocta.tipope)=0))"
        
    Else
        nSQL = "SELECT con_percepciondet.id, mae_documento.abrev, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.fchdoc, vta_ventas.imptotdoc AS imptot, con_percepciondet.porper, con_percepciondet.impper, con_percepciondet.iddoc, [vta_ventas]![imptotdoc]+[con_percepciondet]![impper] AS impcob, mae_documentocta.idmon, mae_documentocta.tipope, mae_documentocta.idcuen,Left([vta_ventas].[numreg],2) & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([vta_ventas].[numreg],3) AS registro " _
            & " FROM (((con_percepciondet LEFT JOIN vta_ventas ON con_percepciondet.iddoc = vta_ventas.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_documentocta ON mae_documento.id = mae_documentocta.iddoc) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
            & " WHERE (((con_percepciondet.id)=" & RstPer("id") & ") AND ((mae_documentocta.idmon)=" & Val(TxtMoneda.Text) & ") " _
            & " AND ((mae_documentocta.tipope)=-1))"
    End If
    RST_Busq Rst, nSQL, xCon
    If Rst.RecordCount <> 0 Then
        Fg1.Rows = 1
        Rst.MoveFirst
        Agregando = True
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(Rst("registro"))
            Fg1.TextMatrix(A, 2) = NulosC(Rst("numdoc"))
            Fg1.TextMatrix(A, 3) = NulosC(Rst("abrev"))
            Fg1.TextMatrix(A, 4) = NulosC(Rst("fchdoc"))
            
            If Rst("idmondoc") = 1 Then
                Fg1.TextMatrix(A, 5) = "0.00"
                Fg1.TextMatrix(A, 6) = Format(Rst("imptot"), FORMAT_MONTO)
            Else
                Fg1.TextMatrix(A, 5) = Format(Rst("imptot"), FORMAT_MONTO)
                Fg1.TextMatrix(A, 6) = Format(Rst("imptot") * NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            End If
            Fg1.TextMatrix(A, 7) = Format(Rst("porper"), FORMAT_MONTO)
            Fg1.TextMatrix(A, 8) = Format(Rst("impper"), "0.0000")
            Fg1.TextMatrix(A, 9) = Format(Rst("impcob"), FORMAT_MONTO)
            Fg1.TextMatrix(A, 10) = NulosN(Rst("iddoc"))
            Fg1.TextMatrix(A, 11) = NulosN(Rst("idcuen"))
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
        Agregando = False
    End If
    HallarTotales
End Sub

Private Sub CambiarMes()
    mMesActivo = SeleccionaMes(xCon)
    OpcionesPeriodo
    TabOne1.CurrTab = 0
End Sub

Private Sub pRegistroAdd(Optional F_SELECCION_VARIOS As Boolean = True)

    On Error GoTo error
    If NulosC(TxtRucPro.Text) = "" Then
        MsgBox "Seleccione el " + LblTitulo.Caption, vbExclamation, xTitulo
        CmdBusProv.SetFocus
        Exit Sub
    End If
    If NulosN(TxtMoneda.Text) = 0 Then
        MsgBox "Seleccione la Moneda", vbExclamation, xTitulo
        CmdBusMon.SetFocus
        Exit Sub
    End If
    Dim xCampos(6, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String '--almacenara los codigos de documentos ya seleccionados
    Dim nSQLNotInDocumentos As String
    Dim nSQL As String
    Dim nTitulo As String
    xCampos(0, 0) = "Num.Reg.":       xCampos(0, 1) = "registro":     xCampos(0, 2) = "1200":         xCampos(0, 3) = "C":    xCampos(0, 4) = "N"
    xCampos(1, 0) = "T.D.":           xCampos(1, 1) = "abrev":        xCampos(1, 2) = "700":          xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "Nº Documento":   xCampos(2, 1) = "numdoc":       xCampos(2, 2) = "2000":         xCampos(2, 3) = "C":    xCampos(2, 4) = "S"
    xCampos(3, 0) = "Fch. Emisión":   xCampos(3, 1) = "fchdoc1":      xCampos(3, 2) = "1200":         xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
    xCampos(4, 0) = "M":              xCampos(4, 1) = "simbolo":      xCampos(4, 2) = "550":          xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
    xCampos(5, 0) = "Importe":        xCampos(5, 1) = "imptot":       xCampos(5, 2) = "1200":         xCampos(5, 3) = "N":    xCampos(4, 4) = "N"
    '*************************************************************
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 8, "com_compras.id", " NOT IN ")
    '*************************************************************
    If Opt_Tipo(0).Value = True Then
        nSQL = "SELECT 0 as xsel,com_compras.id, mae_documento.abrev, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.fchdoc & '' as fchdoc1, com_compras.imptot, mae_documentocta.idmon, mae_documentocta.tipope, mae_documentocta.idcuen, Left([com_compras].[numreg],2) & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([com_compras].[numreg],3) AS registro, com_compras.idmon AS idmondoc, mae_moneda.simbolo  " _
            + vbCr + " FROM (((mae_documento RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) LEFT JOIN mae_documentocta ON mae_documento.id = mae_documentocta.iddoc) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN mae_moneda ON com_compras.idmon = mae_moneda.id " _
            + vbCr + " WHERE (((com_compras.idpro)=" & NulosN(LblIdProveedor.Caption) & ") AND ((mae_documentocta.idmon)=" & NulosN(TxtMoneda.Text) & ") AND ((mae_documentocta.tipope)=0)) "
    
    Else
        nSQL = "SELECT 0 as xsel,vta_ventas.id, mae_documento.abrev, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.fchdoc & '' as fchdoc1, vta_ventas.imptotdoc AS imptot, mae_documentocta.idcuen, Left([vta_ventas].[numreg],2) & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([vta_ventas].[numreg],3) AS registro, vta_ventas.idmon AS idmondoc, mae_moneda.simbolo  " _
            + vbCr + " FROM (((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_documentocta ON mae_documento.id = mae_documentocta.iddoc) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id " _
            + vbCr + " WHERE (((mae_documentocta.idmon)=" & NulosN(TxtMoneda.Text) & ") AND ((mae_documentocta.tipope)=-1) AND ((vta_ventas.idcli)=" & NulosN(LblIdProveedor.Caption) & ")) "
            
            '--consultar del registro de ventas
            '*************************************************************
            nSQLId = Replace(nSQLId, "com_compras.id", "vta_ventas.id")
            '*************************************************************
            nSQLNotInDocumentos = Replace(nSQLNotInDocumentos, "com_compras.id", "vta_ventas.id")
            '*************************************************************
    End If
    
    nTitulo = "Buscando Documento del " + LblTitulo.Caption + ": " + LblProveedor.Caption
    '*************************************************************
    nSQL = nSQL + IIf(nSQLId = "", "", " AND " + nSQLId) + nSQLNotInDocumentos
    '*************************************************************
    If F_SELECCION_VARIOS = True Then
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), nTitulo
    Else
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "numdoc", "numdoc", CualquierParte
    End If
    If xRs.State = 0 Then GoTo SALIR
    If xRs.RecordCount = 0 Then GoTo SALIR
    If F_SELECCION_VARIOS = True Then xRs.MoveFirst
    Agregando = True
    Do While Not xRs.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("registro"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRs("numdoc"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(xRs("abrev"))
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(xRs("fchdoc1"))
        
        If xRs("idmondoc") = 1 Then
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(xRs("imptot")), FORMAT_MONTO)
        Else
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(xRs("imptot")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(xRs("imptot") * NulosN(LblTipoCambio.Caption)), FORMAT_MONTO)
        End If
        '-----------
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(LblTasa.Caption), "0.00")
        Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 6)) * NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 7)) / 100
        Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 6)) + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 7))
        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(Fg1.TextMatrix(Fg1.Rows - 1, 9), FORMAT_MONTO)
        '-----------
        Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosN(xRs("id"))
        Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(xRs("idcuen"))
        If F_SELECCION_VARIOS = False Then Exit Do
        xRs.MoveNext
    Loop
    Agregando = False
    HallarTotales
    Fg1.Row = Fg1.Rows - 1: Fg1.Col = 6:  Fg1.SetFocus
SALIR:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    HallarTotales
    Agregando = False
    SHOW_ERROR Me.Name, "Fg1_CellButtonClick"
    CmdAdd.SetFocus
End Sub


Private Sub Buscar()
    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim xSQL As String
    ReDim xCampos(7, 4) As String
    
    xCampos(0, 0) = "Num.Reg.":             xCampos(0, 1) = "registro":    xCampos(0, 2) = "900":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Tipo":                 xCampos(1, 1) = "tipmov":    xCampos(1, 2) = "600":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "Cliente / Proveedor":  xCampos(2, 1) = "nombre":    xCampos(2, 2) = "2700":  xCampos(2, 3) = "C"
    xCampos(3, 0) = "Fch. Doc.":            xCampos(3, 1) = "fchdoc":    xCampos(3, 2) = "950":  xCampos(3, 3) = "F"
    xCampos(4, 0) = "Nº Documento":         xCampos(4, 1) = "numdoc1":   xCampos(4, 2) = "1500":  xCampos(4, 3) = "C"
    xCampos(5, 0) = "M":                    xCampos(5, 1) = "monabrev":  xCampos(5, 2) = "450":   xCampos(5, 3) = "C"
    xCampos(6, 0) = "Imp.Per.":             xCampos(6, 1) = "imptotper": xCampos(6, 2) = "900":  xCampos(6, 3) = "N"
    
    Set RstTmp = RstPer.Clone
    CARGAR_DLL_EPSBUSCAR xCon, xRs, "", xCampos(), "Buscando Percepciones", "numreg", "numreg", CualquierParte, , RstTmp
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    RstPer.MoveFirst
    RstPer.Find "id = " & xRs("id") & ""
SALIR:
    Set RstTmp = Nothing
    Set xRs = Nothing
error:
    Set RstTmp = Nothing
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub


Private Sub pExportar()
    If TabOne1.CurrTab = 0 Then
        Dim oExport As New SGI2_funciones.formularios
        Dim RstTmp  As New ADODB.Recordset
        Dim xCampos(10, 3) As String
        
        '0::Nombre a Mostrar;
        '1::nombre de Campo del Rst;
        '2::alineacion(0::derecha, 1::centro, 2::izquierda);
        '3::ancho de columna
        '--obs: el rst puede tener mas columnas solo se consideran los campos del array
        xCampos(0, 0) = "Nº. Reg":              xCampos(0, 1) = "registro":   xCampos(0, 2) = 1:    xCampos(0, 3) = "900"
        xCampos(1, 0) = "Tipo":                 xCampos(1, 1) = "tipmov":     xCampos(1, 2) = 0:    xCampos(1, 3) = "743"
        xCampos(2, 0) = "Nº RUC":               xCampos(2, 1) = "numruc":     xCampos(2, 2) = 0:    xCampos(2, 3) = "1229"
        xCampos(3, 0) = "Proveedor/Cliente":    xCampos(3, 1) = "nombre":     xCampos(3, 2) = 0:    xCampos(3, 3) = "3500"
        xCampos(4, 0) = "T.D.":                 xCampos(4, 1) = "docabrev":   xCampos(4, 2) = 1:    xCampos(4, 3) = "443"
        xCampos(5, 0) = "Fch.Emi":              xCampos(5, 1) = "fchdoc":     xCampos(5, 2) = 1:    xCampos(5, 3) = "1014"
        xCampos(6, 0) = "Nº.Documento":         xCampos(6, 1) = "numdoc1":    xCampos(6, 2) = 0:    xCampos(6, 3) = "1700"
        xCampos(7, 0) = "Glosa":                xCampos(7, 1) = "glosa":      xCampos(7, 2) = 0:    xCampos(7, 3) = "3000"
        xCampos(8, 0) = "M":                    xCampos(8, 1) = "monabrev":   xCampos(8, 2) = 1:    xCampos(8, 3) = "386"
        xCampos(9, 0) = "Importe":              xCampos(9, 1) = "imptotper":  xCampos(9, 2) = 2:    xCampos(9, 3) = "943"
        xCampos(10, 0) = "Saldo":               xCampos(10, 1) = "impsal":    xCampos(10, 2) = 2:   xCampos(10, 3) = "943"
        Set RstTmp = RstPer.Clone
        oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Registro de Percepciones", lblperiodo(0).Caption & " - " & AnoTra, "", "Registro de Percepciones", RstTmp, xCampos
        Set oExport = Nothing
        Set RstTmp = Nothing

        Exit Sub
    End If



    Dim A&, B&
    Dim xFilas&
    On Error GoTo error
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    'abre el Libro
    objExcel.WindowState = 1
    objExcel.Workbooks.Add
    With objExcel.ActiveSheet
        
        .Cells(1, 2) = NomEmp
        .Cells(1, 9) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        
        .Columns(2).ColumnWidth = Fg1.ColWidth(1) / 100
        .Columns(3).ColumnWidth = Fg1.ColWidth(2) / 100
        .Columns(4).ColumnWidth = Fg1.ColWidth(3) / 100
        .Columns(5).ColumnWidth = Fg1.ColWidth(4) / 100
        .Columns(6).ColumnWidth = Fg1.ColWidth(5) / 100
        .Columns(7).ColumnWidth = Fg1.ColWidth(6) / 100
        .Columns(8).ColumnWidth = Fg1.ColWidth(7) / 100
        .Columns(9).ColumnWidth = Fg1.ColWidth(8) / 100
                        
        '-----encabezado
        xFilas = 4
        .Cells(xFilas, 2) = "Registro de Percepciones"
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Tipo Mov."
        .Cells(xFilas, 3) = IIf(Opt_Tipo(0).Value = True, "Compra", "Venta")
        .Cells(xFilas, 8) = "Periodo"
        .Cells(xFilas, 9) = lblperiodo(1).Caption
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = LblTitulo.Caption
        .Cells(xFilas, 3) = LblProveedor.Caption
        
        .Cells(xFilas, 8) = "RUC"
        .Cells(xFilas, 9) = "'" + TxtRucPro.Text
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Fch.Emi."
        .Cells(xFilas, 3) = "'" & TxtFchEmi.Valor
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Documento"
        .Cells(xFilas, 3) = LblDocumento.Caption
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "N°.Doc."
        .Cells(xFilas, 3) = "'" & TxtNumSer.Text & "-" & TxtNumDoc.Text
        
        .Cells(xFilas, 8) = "T.C."
        .Cells(xFilas, 9) = NulosN(LblTipoCambio.Caption)
        
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Moneda"
        .Cells(xFilas, 3) = "'" & LblMoneda.Caption
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Percepción"
        .Cells(xFilas, 3) = "'" & LblPercepcion.Caption
            
        .Cells(xFilas, 2) = "Tasa"
        .Cells(xFilas, 3) = NulosN(LblTasa.Caption) & "%"
        '--titulo
        xFilas = xFilas + 2
        For A = 1 To 8
            .Cells(xFilas, A + 1) = Fg1.TextMatrix(0, A)
        Next A
       '--detalle
        xFilas = xFilas + 1
        For A = 1 To Fg1.Rows - 1
            For B = 1 To 8
                If B < 5 Then
                    .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                Else
                    .Cells(xFilas, B + 1) = NulosN(Fg1.TextMatrix(A, B))
                End If
            Next B
            xFilas = xFilas + 1
        Next A

        .Cells(xFilas, 4) = "Total =>"
        .Cells(xFilas, 6) = NulosN(TxtImporte.Text)
        .Cells(xFilas, 8) = NulosN(TxtImpRet.Text)
        .Cells(xFilas, 9) = NulosN(TxtImpCob.Text)
        
    End With
    
    MsgBox "El Registro se exportó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 1
    objExcel.Visible = True
    Set objExcel = Nothing
    Exit Sub
error:
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "Exportar", , IIf(Err.Number <> 50290, "", "No manipule el archivo hasta que termine de exportar!!!!")
End Sub





Private Sub OpcionesPeriodo()
     
     lblperiodo(0).Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
     lblperiodo(1).Caption = lblperiodo(0).Caption
    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    TDB_FiltroLimpiar Dg1
    Set RstPer = Nothing
    '------------------------------------------
    
    On Error GoTo error
    Dim nSQL  As String
    
    nSQL = "SELECT IIf([con_percepcion].[tipo]=1,'Compra','Venta') AS tipmov, IIf(con_percepcion.tipo=1,mae_prov.numruc,mae_cliente.numruc) AS numruc, IIf(con_percepcion.tipo=1,mae_prov.nombre,mae_cliente.nombre) AS nombre, mae_documento.abrev AS docabrev, mae_documento.descripcion AS docdesc, con_percepcion.idmon AS monid, mae_moneda.descripcion AS mondesc, mae_moneda.simbolo AS monabrev, con_percepcion.idper AS precepid, mae_percepcion.descripcion AS percepdesc, mae_percepcion.tasa AS perceptasa, mae_percepcion.idcuencom AS percepidcom, mae_percepcion.idcuenven AS percepidven, [con_percepcion]!numser+'-'+[con_percepcion]!numdoc AS numdoc1, con_percepcion.*, Format([con_percepcion].[idmes],'00') & [mae_libros].[codsun] & Right([con_percepcion].[numreg],4) AS registro, " _
        + vbCr + " con_percepcion.fchdoc & '' as fchdoc1,con_percepcion.imptotper & '' as imptotper1,con_percepcion.impsal & '' as impsal1 " _
        + vbCr + " FROM (mae_percepcion INNER JOIN ((((con_percepcion LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id) LEFT JOIN mae_documento ON con_percepcion.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON con_percepcion.idmon = mae_moneda.id) LEFT JOIN mae_cliente ON con_percepcion.idcli = mae_cliente.id) ON mae_percepcion.id = con_percepcion.idper) LEFT JOIN mae_libros ON con_percepcion.idlib = mae_libros.id " _
        + vbCr + " WHERE (((con_percepcion.ano) = " + CStr(AnoTra) + ") And ((con_percepcion.idmes) = " & mMesActivo & ")) " _
        + vbCr + " ORDER BY con_percepcion.numreg ;"

    Me.MousePointer = vbHourglass
    RST_Busq RstPer, nSQL, xCon
    
    Set Dg1.DataSource = RstPer
    Me.MousePointer = vbDefault
    
Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
    
End Sub



