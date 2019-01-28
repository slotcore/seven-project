VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmManOrdenCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras - Orden de Compra"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   0
      TabIndex        =   10
      Top             =   360
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
         TabIndex        =   16
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame7 
            Caption         =   "[ Estado ]"
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
            Height          =   630
            Left            =   8025
            TabIndex        =   55
            Top             =   1260
            Width           =   3600
            Begin VB.Label LblEstado 
               Alignment       =   2  'Center
               Caption         =   "LblEstado"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   105
               TabIndex        =   56
               Top             =   285
               Width           =   3360
            End
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   7650
            Picture         =   "FrmManOrdenCompra.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   990
            Width           =   240
         End
         Begin VB.Frame Frame4 
            Height          =   855
            Left            =   6480
            TabIndex        =   44
            Top             =   5940
            Width           =   5175
            Begin VB.Label Label2 
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
               Index           =   2
               Left            =   3465
               TabIndex        =   50
               Top             =   195
               Width           =   450
            End
            Begin VB.Label Label2 
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
               Index           =   1
               Left            =   1965
               TabIndex        =   49
               Top             =   195
               Width           =   510
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Importe Bruto"
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
               Index           =   0
               Left            =   435
               TabIndex        =   48
               Top             =   195
               Width           =   1155
            End
            Begin VB.Label LblIgv 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblIgv"
               Height          =   300
               Left            =   1950
               TabIndex        =   47
               Top             =   420
               Width           =   1350
            End
            Begin VB.Label LblTotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblTotal"
               Height          =   300
               Left            =   3465
               TabIndex        =   46
               Top             =   420
               Width           =   1350
            End
            Begin VB.Label LblBruto 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblBruto"
               Height          =   300
               Left            =   420
               TabIndex        =   45
               Top             =   420
               Width           =   1350
            End
         End
         Begin VB.CommandButton CmdBusCoti 
            Height          =   240
            Left            =   7650
            Picture         =   "FrmManOrdenCompra.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   600
            Width           =   240
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2415
            Locked          =   -1  'True
            TabIndex        =   2
            Text            =   "TxtNumDoc"
            Top             =   960
            Width           =   1875
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   1
            Text            =   "TxtNumSer"
            Top             =   960
            Width           =   900
         End
         Begin VB.CommandButton CmdBusArea 
            Height          =   240
            Left            =   2115
            Picture         =   "FrmManOrdenCompra.frx":0264
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1635
            Width           =   240
         End
         Begin VB.TextBox TxtIdArea 
            Height          =   300
            Left            =   1470
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   5
            Text            =   "TxtIdArea"
            Top             =   1605
            Width           =   915
         End
         Begin VB.CommandButton CmdBusSol 
            Height          =   240
            Left            =   2115
            Picture         =   "FrmManOrdenCompra.frx":0396
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   1320
            Width           =   240
         End
         Begin VB.TextBox TxtIdSol 
            Height          =   300
            Left            =   1470
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   4
            Text            =   "TxtIdSol"
            Top             =   1290
            Width           =   915
         End
         Begin VB.CommandButton CmdBusPro 
            Height          =   240
            Left            =   3030
            Picture         =   "FrmManOrdenCompra.frx":04C8
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2280
            Width           =   240
         End
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1470
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   8
            Text            =   "TxtNumRuc"
            Top             =   2250
            Width           =   1830
         End
         Begin VB.CommandButton CmdBusCondPag 
            Height          =   240
            Left            =   2115
            Picture         =   "FrmManOrdenCompra.frx":05FA
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   2595
            Width           =   240
         End
         Begin VB.Frame Frame5 
            Height          =   555
            Left            =   120
            TabIndex        =   27
            Top             =   375
            Width           =   3885
            Begin VB.OptionButton Option2 
               Caption         =   "Sin Cotizacion"
               Height          =   210
               Left            =   465
               TabIndex        =   29
               Top             =   225
               Width           =   1365
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Con Cotizacion"
               Height          =   210
               Left            =   2025
               TabIndex        =   28
               Top             =   225
               Width           =   1485
            End
         End
         Begin VB.TextBox TxtIdConPag 
            Height          =   300
            Left            =   1470
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   9
            Text            =   "TxtIdConPag"
            Top             =   2565
            Width           =   915
         End
         Begin VB.TextBox TxtNumCoti 
            Height          =   300
            Left            =   5310
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   0
            Text            =   "TxtNumCoti"
            Top             =   570
            Width           =   2610
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2685
            Left            =   105
            TabIndex        =   18
            Top             =   3225
            Width           =   11520
            _cx             =   20320
            _cy             =   4736
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
            FormatString    =   $"FrmManOrdenCompra.frx":072C
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
            Left            =   1470
            TabIndex        =   6
            Top             =   1920
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEnt 
            Height          =   300
            Left            =   6705
            TabIndex        =   7
            Top             =   1935
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
         Begin VB.Frame Frame3 
            Height          =   870
            Left            =   105
            TabIndex        =   40
            Top             =   5925
            Width           =   6210
            Begin VB.CommandButton CmdAddNewItem 
               Caption         =   "Agregar Nuevo Item"
               Height          =   525
               Left            =   3975
               TabIndex        =   43
               Top             =   225
               Width           =   1395
            End
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "&Eliminar Item"
               Height          =   525
               Left            =   2310
               TabIndex        =   42
               Top             =   225
               Width           =   1395
            End
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "Agregar Item"
               Height          =   525
               Left            =   870
               TabIndex        =   41
               Top             =   225
               Width           =   1395
            End
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   6990
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   3
            Text            =   "TxtI"
            Top             =   960
            Width           =   915
         End
         Begin VB.Label LblIdProv 
            Caption         =   "LblIdProv"
            ForeColor       =   &H000000FF&
            Height          =   165
            Left            =   8145
            TabIndex        =   58
            Top             =   2625
            Visible         =   0   'False
            Width           =   1005
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
            Left            =   7950
            TabIndex        =   54
            Top             =   960
            Width           =   3345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   6090
            TabIndex        =   53
            Top             =   1020
            Width           =   585
         End
         Begin VB.Label LblIdCotizacion 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCotizacion"
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
            Left            =   8250
            TabIndex        =   51
            Top             =   555
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº Orden Compra"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   1020
            Width           =   1245
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Entrega"
            Height          =   195
            Left            =   5595
            TabIndex        =   37
            Top             =   1980
            Width           =   915
         End
         Begin VB.Label LblArea 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblArea"
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
            Left            =   2415
            TabIndex        =   36
            Top             =   1605
            Width           =   5490
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Area"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   1650
            Width           =   330
         End
         Begin VB.Label LblSolicitante 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblSolicitante"
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
            Left            =   2415
            TabIndex        =   33
            Top             =   1290
            Width           =   5490
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   2310
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "[  Lista de Items  ]"
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
            Left            =   120
            TabIndex        =   26
            Top             =   2985
            Width           =   1560
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Orden de Compra"
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
            Left            =   90
            TabIndex        =   25
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Solicitante"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1335
            Width           =   735
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Nº Cotizacion"
            Height          =   195
            Left            =   4215
            TabIndex        =   23
            Top             =   600
            Width           =   960
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Emision"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   1965
            Width           =   900
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cond. de Pago"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   2640
            Width           =   1065
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
            Left            =   2415
            TabIndex        =   20
            Top             =   2565
            Width           =   5490
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
            Left            =   3345
            TabIndex        =   19
            Top             =   2250
            Width           =   8295
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   -12435
         TabIndex        =   11
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6480
            Left            =   30
            TabIndex        =   12
            Top             =   315
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11430
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
            Columns(1).Caption=   "Tipo"
            Columns(1).DataField=   "tipoorden"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numdoc2"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "M"
            Columns(3).DataField=   "simbolo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Emi."
            Columns(4).DataField=   "fchemi"
            Columns(4).NumberFormat=   "Short Date"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fch. Ent."
            Columns(5).DataField=   "fchent"
            Columns(5).NumberFormat=   "Short Date"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Nº Ord. Cotizacion"
            Columns(6).DataField=   "numordcot"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Solicitante"
            Columns(7).DataField=   "nomsol"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Condicion"
            Columns(8).DataField=   "descest"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   4
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Enviado"
            Columns(9).DataField=   "envcor"
            Columns(9).NumberFormat=   "General Number"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   397
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2143"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2064"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2778"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2699"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=900"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=820"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1720"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1640"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1693"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1614"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=3096"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=3016"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=4154"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=4075"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=1958"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1879"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=1376"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1296"
            Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=513"
            Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=86,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=83,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=84,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=85,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=17"
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
            TabIndex        =   15
            Top             =   30
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta Ordenes de Compra"
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
            TabIndex        =   14
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblPeriodo 
            Alignment       =   2  'Center
            Caption         =   "LblPeriodo"
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
            TabIndex        =   13
            Top             =   0
            Visible         =   0   'False
            Width           =   1860
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
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":0899
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":0DDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":116F
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":12F3
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":1747
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":185F
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":1DA3
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":22E7
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":23FB
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":250F
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":2963
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":2ACF
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":3017
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdenCompra.frx":33A9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   57
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
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Enviar por Correo Electronico"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir "
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
Attribute VB_Name = "FrmManOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim RstLista As New ADODB.Recordset
Dim CaracteresNumericos As String
Dim oPDF As cPDF
Dim Pagina  As Integer
Dim fOrdenLista As Boolean                 ' --especfica el orden de la lista de la consulta
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Sub ActivarTextos(Valor As Boolean)
    TxtidSol.Locked = Valor
    TxtIdArea.Locked = Valor
    TxtFchEmi.Locked = Valor
    TxtFchEnt.Locked = Valor
    TxtNumRuc.Locked = Valor
    'TxtIdConPag.Locked = Valor
    TxtIdMon.Locked = Valor
End Sub

Private Sub CmdAceptaOrd_Click()
    If RstLista("idcond") = 4 Then
        MsgBox "No se puede aprobar una orden de compra rechazada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idcond = 1 WHERE (((com_ordencompra.id)=" & RstLista("id") & "))"
    RstLista.Requery
    MsgBox "La orden de compra se aprobo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    TabOne1.CurrTab = 0
End Sub

Private Sub CmdAddItem_Click()
    If QueHace = 3 Then Exit Sub
    
    If Fg1.Rows = 1 Then
        Fg1.Rows = Fg1.Rows + 1
        Fg1_CellButtonClick Fg1.Rows - 1, 2
    Else
        If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 2)) <> "" Then
            Fg1.Rows = Fg1.Rows + 1
            Fg1.Select Fg1.Rows - 1, 1
            Fg1_CellButtonClick Fg1.Rows - 1, 2
        End If
    End If
End Sub

Private Sub CmdAddNewItem_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xFun As New Sgi2_Procesos.Procesos
    Dim xIdProducto As Integer
    Dim xRs As New ADODB.Recordset
    
    xIdProducto = xFun.IngRapidoItems(xCon)
    If xIdProducto <> 0 Then
        RST_Busq xRs, "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, mae_tipoproducto.descripcion AS desctippro, alm_inventario.id, " _
            & " alm_inventario.idunimed FROM mae_tipoproducto RIGHT JOIN (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) " _
            & " ON mae_tipoproducto.id = alm_inventario.tippro Where (((alm_inventario.activo) = -1) And ((alm_inventario.id) = " & xIdProducto & ")) ORDER BY alm_inventario.descripcion", xCon
    
        If xRs.RecordCount <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 2) = xRs("descripcion")
            Fg1.TextMatrix(Fg1.Row, 3) = xRs("abrev")
            Fg1.TextMatrix(Fg1.Row, 8) = xRs("id")
            Fg1.TextMatrix(Fg1.Row, 9) = xRs("idunimed")
        End If
        Set xRs = Nothing
        Fg1.SetFocus
    End If
End Sub

Private Sub CmdBusArea_Click()
    If QueHace = 3 Then Exit Sub
    If Option1.Value = True Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT * FROM mae_area ORDER BY descripcion"
    
    xForm.Titulo = "Buscando Area"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdArea.Text = xRs("id")
            LblArea.Caption = xRs("descripcion")
            TxtFchEmi.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusCondPag_Click()
    If QueHace = 3 Then Exit Sub
    'If Option1.Value = True Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion ":     xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":           xCampos(1, 1) = "id":             xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT * FROM mae_condpago ORDER BY descripcion"
    
    xForm.Titulo = "Buscando Condicion de Pago"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdConPag.Text = xRs("id")
            LblCondPag.Caption = xRs("descripcion")
            Fg1.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusCoti_Click()
    If QueHace = 3 Then Exit Sub
    'If Option1.Value = True Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(5, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Documento":             xCampos(0, 1) = "descdoc":    xCampos(0, 2) = "1400":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº Documento":          xCampos(1, 1) = "numdoc":     xCampos(1, 2) = "1400":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Proveedor":             xCampos(2, 1) = "descpro":    xCampos(2, 2) = "3000":    xCampos(2, 3) = "C"
    xCampos(3, 0) = "Fch. Emi.":             xCampos(3, 1) = "fchemi":     xCampos(3, 2) = "1000":    xCampos(3, 3) = "C"
    xCampos(4, 0) = "Nº Ord. Reqto.":        xCampos(4, 1) = "numordreq":  xCampos(4, 2) = "1400":    xCampos(4, 3) = "C"
    
    xForm.SQLCad = "SELECT mae_documento.descripcion AS descdoc, mae_prov.nombre AS descpro, com_ordencot.fchemi, " _
        & " [com_ordenreq]![numser] & '-' & [com_ordenreq]![numdoc] AS numordreq, [com_ordencot]![numser] & '-' & [com_ordencot]![numdoc] AS numdoc, " _
        & " com_ordencot.id FROM ((com_ordencot LEFT JOIN mae_prov ON com_ordencot.idpro = mae_prov.id) LEFT JOIN mae_documento " _
        & " ON com_ordencot.idtipdoc = mae_documento.id) LEFT JOIN com_ordenreq ON com_ordencot.idor = com_ordenreq.id Where (((com_ordencot.idest) = 2)) " _
        & " ORDER BY [com_ordenreq]![numser] & '-' & [com_ordenreq]![numdoc], [com_ordencot]![numser] & '-' & [com_ordencot]![numdoc]"

    xForm.Titulo = "Buscando Orden de Cotizacion"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "numdoc"
    xForm.CampoBusca = "numdoc"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumCoti.Text = xRs("numdoc")
            LblIdCotizacion.Caption = xRs("id")
            MuestraDatosOR xRs("id")
            TxtIdMon.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Sub MuestraDatosOR(idOR As Integer)
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    
    ' MOSTRAMOS LA CABECERA
    'RST_Busq xRs, "SELECT * FROM com_ordenreq WHERE id = " & idOR & "", xCon
    RST_Busq xRs, "SELECT com_ordencot.*, mae_prov.numruc, mae_prov.nombre AS nompro FROM com_ordencot LEFT JOIN mae_prov ON com_ordencot.idpro = mae_prov.id " _
        & " WHERE (((com_ordencot.id)=" & idOR & "))", xCon

    If xRs.RecordCount <> 0 Then
        TxtidSol.Text = xRs("idsol")
        TxtIdSol_Validate True
        TxtIdArea.Text = NulosN(xRs("idarea"))
        TxtIdArea_Validate True
        TxtFchEmi.Valor = xRs("fchemi")
        'TxtFchEnt.Valor = xRs("fchent")
        TxtIdMon.Text = xRs("idmon")
        TxtIdMon_Validate True
        LblIdProv.Caption = xRs("idpro")
        TxtNumRuc.Text = xRs("numruc")
        LblProveedor.Caption = xRs("nompro")
        
    End If
    Set xRs = Nothing
    
    ' MOSTRAMOS EL DETALLE
    RST_Busq xRs, "SELECT com_ordencotdet.*, alm_inventario.descripcion AS descitem, mae_unidades.descripcion AS descunimed, man_equipos.nombre AS nomequi, " _
        & " man_tipo.descripcion AS nomtipo FROM (((com_ordencotdet LEFT JOIN alm_inventario ON com_ordencotdet.iditem = alm_inventario.id) LEFT JOIN mae_unidades " _
        & " ON com_ordencotdet.idunimed = mae_unidades.id) LEFT JOIN man_equipos ON com_ordencotdet.idequi = man_equipos.id) LEFT JOIN man_tipo " _
        & " ON com_ordencotdet.idtip = man_tipo.id WHERE (((com_ordencotdet.idoc)=8))", xCon
    
    If xRs.RecordCount <> 0 Then
        Fg1.Rows = 1
        For A = 1 To xRs.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(xRs("nomtipo"))
            Fg1.TextMatrix(A, 2) = NulosC(xRs("nomequi"))
            Fg1.TextMatrix(A, 3) = xRs("descitem")
            Fg1.TextMatrix(A, 4) = xRs("descunimed")
            Fg1.TextMatrix(A, 5) = Format(xRs("cantidad"), "0.00")
            Fg1.TextMatrix(A, 6) = Format(xRs("precio"), "0.00")
            Fg1.TextMatrix(A, 7) = Format(xRs("cantidad") * xRs("precio"), "0.00")
            Fg1.TextMatrix(A, 8) = xRs("iditem")
            Fg1.TextMatrix(A, 9) = xRs("idunimed")
            Fg1.TextMatrix(A, 10) = NulosN(xRs("idtip"))
            Fg1.TextMatrix(A, 11) = NulosN(xRs("idequi"))
            xRs.MoveNext
            If xRs.EOF = True Then Exit For
        Next A
    End If
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub
    If Option1.Value = True Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT * FROM mae_moneda ORDER BY descripcion"
    
    xForm.Titulo = "Buscando Moneda"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdMon.Text = xRs("id")
            LblMoneda.Caption = xRs("descripcion")
            TxtidSol.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusPro_Click()
    If QueHace = 3 Then Exit Sub
    If Option1.Value = True Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre ":           xCampos(0, 1) = "nombre":         xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":         xCampos(1, 1) = "numruc":         xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT * FROM mae_prov WHERE activo = -1 ORDER BY nombre"
    
    xForm.Titulo = "Buscando Proveedor"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "nombre"
    xForm.CampoBusca = "nombre"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            LblIdProv.Caption = xRs("id")
            TxtNumRuc.Text = xRs("numruc")
            LblProveedor.Caption = xRs("nombre")
            TxtIdConPag.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusSol_Click()
    If QueHace = 3 Then Exit Sub
    If Option1.Value = True Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Empleado":    xCampos(0, 1) = "apenom":           xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":      xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT com_usuario.id, com_usuario.idper, UCase([pla_empleados]![apepat]) & ' ' & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS apenom " _
        & " FROM com_usuario LEFT JOIN pla_empleados ON com_usuario.idper = pla_empleados.id"

    xForm.Titulo = "Buscando Usuarios"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "apenom"
    xForm.CampoBusca = "apenom"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtidSol.Text = xRs("id")
            LblSolicitante.Caption = xRs("apenom")
            TxtIdArea.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdCancelarOrd_Click()
    If RstLista("idcond") = 1 Then
        MsgBox "No se puede rechazar una orden de compra aprobada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idcond = 4 WHERE (((com_ordencompra.id)=" & RstLista("id") & "))"
    MsgBox "La orden de compra se rechazo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    TabOne1.CurrTab = 0
    RstLista.Requery
End Sub

Private Sub CmdDelItem_Click()
    If QueHace = 3 Then Exit Sub
    
    If Fg1.Rows = 1 Then
        MsgBox "No hay items para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    Fg1.RemoveItem Fg1.Row
    Dim A As Integer
    
'    For A = 1 To Fg1.Rows - 1
'        Fg1.TextMatrix(A, 1) = Str(A)
'    Next A
End Sub


Private Sub Dg1_DblClick()
    ' MUESTRA INFORMACION EN LA PESTAÑA DETALLE
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstLista
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LAS COLUMNAS DEL CONTROL Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstLista.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstLista("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If QueHace = 3 Then Exit Sub
    If Option1.Value = True Then Exit Sub
    
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    Dim xCampos2(2, 4) As String
    Dim xCampos3(2, 4) As String
    Dim xCampos1(3, 4) As String
    'Dim xForm As New eps_librerias.FormBuscar
    'Dim xRs As New ADODB.Recordset
       
    If Col = 1 Then
        xCampos2(0, 0) = "Descripcion":    xCampos2(0, 1) = "descripcion":      xCampos2(0, 2) = "4000":         xCampos2(0, 3) = "C"
        xCampos2(1, 0) = "Codigo":         xCampos2(1, 1) = "id":               xCampos2(1, 2) = "1400":         xCampos2(1, 3) = "N"
        
        xForm.SQLCad = "SELECT * FROM man_equipotipo"
        xForm.Titulo = "Buscando Tipos"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos2)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 1) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 10) = xRs("id")
            End If
        End If
    End If
    
    If Col = 2 Then
        'If NulosN(Fg1.TextMatrix(Fg1.Row, 11)) = 0 Then Exit Sub
        xCampos1(0, 0) = "Descripcion":     xCampos1(0, 1) = "nombre":          xCampos1(0, 2) = "3500":   xCampos1(0, 3) = "C"
        xCampos1(1, 0) = "Caracteristicas": xCampos1(1, 1) = "caracteristicas": xCampos1(1, 2) = "5000":   xCampos1(1, 3) = "C"
        xCampos1(2, 0) = "Codigo":          xCampos1(2, 1) = "id":              xCampos1(2, 2) = "1000":   xCampos1(2, 3) = "N"
        
        xForm.SQLCad = "SELECT man_equipos.nombre, man_equipos.id, man_equipos.caracteristicas From man_equipos WHERE (((man_equipos.idtip)=" & NulosN(Fg1.TextMatrix(Fg1.Row, 10)) & "))"

        xForm.Titulo = "Buscando Equipos e Instalaciones"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "nombre"
        xForm.CampoBusca = "nombre"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos1)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 2) = xRs("nombre")
                Fg1.TextMatrix(Fg1.Row, 11) = xRs("id")
            End If
        End If
    End If

    If Col = 3 Then
        Dim xCampos(3, 4) As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "codpro":           xCampos(1, 2) = "1400":         xCampos(1, 3) = "c"
        xCampos(2, 0) = "Abreviatura":    xCampos(2, 1) = "abrev":            xCampos(2, 2) = "1000":         xCampos(2, 3) = "c"
        xCampos(3, 0) = "Tipo Producto":  xCampos(3, 1) = "desctippro":       xCampos(3, 2) = "1200":         xCampos(3, 3) = "c"
        
        xForm.SQLCad = "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, mae_tipoproducto.descripcion AS desctippro, " _
            & " alm_inventario.id, alm_inventario.idunimed FROM mae_tipoproducto RIGHT JOIN (mae_unidades RIGHT JOIN alm_inventario " _
            & " ON mae_unidades.id = alm_inventario.idunimed) ON mae_tipoproducto.id = alm_inventario.tippro Where (((alm_inventario.activo) = -1)) " _
            & " ORDER BY alm_inventario.descripcion"
        
        xForm.Titulo = "Buscando Items"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 3) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 4) = xRs("abrev")
                Fg1.TextMatrix(Fg1.Row, 8) = xRs("id")
                Fg1.TextMatrix(Fg1.Row, 9) = xRs("idunimed")
                'Fg1.TextMatrix(Fg1.Row, 1) = (NulosN(Fg1.TextMatrix(Fg1.Row - 1, 1)) + 1)
                
                If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 2)) <> "" Then
                    Fg1.Rows = Fg1.Rows + 1
                    'Fg1.TextMatrix(Fg1.Rows - 1, 1) = (NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 1)) + 1)
                End If
            End If
        End If
    End If
    
'    If Col = 7 Then
'        If NulosN(TxtIdArea.Text) = 0 Then
'            MsgBox "No ha especificado el area que solicita la orden de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            TxtIdArea.SetFocus
'            Exit Sub
'        End If
'
'        Dim xCampos2(2, 4) As String
'
'        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
'        xCampos2(0, 0) = "Codigo":        xCampos2(0, 1) = "codigo":          xCampos2(0, 2) = "1200":         xCampos2(0, 3) = "C"
'        xCampos2(1, 0) = "Descripcion":   xCampos2(1, 1) = "descripcion":     xCampos2(1, 2) = "6000":         xCampos2(1, 3) = "C"
'
'        xForm.SQLCad = "SELECT con_centrocosto.* FROM con_centocostoarea LEFT JOIN con_centrocosto ON con_centocostoarea.idcencos = con_centrocosto.id " _
'            & " Where (((con_centocostoarea.idarea) = " & NulosN(TxtIdArea.Text) & ")) ORDER BY con_centrocosto.codigo"
'
'        xForm.Titulo = "Buscando Centros de Costos"
'        xForm.FormaBusca = Principio
'        xForm.Criterio = ""
'        xForm.Ordenado = "codigo"
'        xForm.CampoBusca = "codigo"
'        Set xForm.Coneccion = xCon
'        Set xRs = xForm.BuscarReg(xCampos2)
'        If xRs.State = 1 Then
'            If xRs.RecordCount <> 0 Then
'                Fg1.TextMatrix(Fg1.Row, 7) = Trim(xRs("descripcion"))
'                Fg1.TextMatrix(Fg1.Row, 10) = xRs("id")
'            End If
'        End If
'    End If
    
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 5 Or Col = 6 Then
        Fg1.TextMatrix(Fg1.Row, 5) = Format(Fg1.TextMatrix(Fg1.Row, 5), "0.00")
        Fg1.TextMatrix(Fg1.Row, 6) = Format(Fg1.TextMatrix(Fg1.Row, 6), "0.00")
        Fg1.TextMatrix(Fg1.Row, 7) = NulosN(Fg1.TextMatrix(Fg1.Row, 5)) * NulosN(Fg1.TextMatrix(Fg1.Row, 6))
        Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), "0.00")
        SumarTotal
    End If
End Sub

Sub SumarTotal()
    Dim A As Integer
    Dim xTotal As Double
    For A = 1 To Fg1.Rows - 1
        xTotal = xTotal + NulosN(Fg1.TextMatrix(A, 7))
    Next A
    LblBruto.Caption = Format(xTotal, "0.00")
    LblIgv.Caption = Format(((xTotal * 1.19) - xTotal), "0.00")
    LblTotal.Caption = Format(xTotal * 1.19, "0.00")
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then Exit Sub
    If Fg1.Col = 1 Or Fg1.Col = 2 Or Fg1.Col = 3 Or Fg1.Col = 5 Or Fg1.Col = 6 Then
        'If Fg1.Col = 2 Or Fg1.Col = 7 Then
            If Option1.Value = True Then Exit Sub
        'End If
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 5 Or Col = 6 Then
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
            
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
       '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
                
        RST_Busq RstLista, "SELECT IIf([com_ordencompra]![tipo]=1,'Con Cotizacion','Sin Contizacion') AS tipoorden, com_ordencompra.*, mae_documento.descripcion AS descdoc, " _
            & " [com_ordencompra]![numser] & '-' & [com_ordencompra]![numdoc] AS numdoc2, UCase([pla_empleados]![apepat]) & ' ' & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS nomsol, " _
            & " mae_area.descripcion AS descarea, mae_condpago.descripcion AS desccondpag, mae_moneda.descripcion AS descmon, mae_prov.numruc, mae_prov.nombre, " _
            & " mae_estados.descripcion AS descest, [com_ordencot]![numser] & '-' & [com_ordencot]![numdoc] AS numordcot, mae_moneda.simbolo FROM ((((((((com_ordencompra " _
            & " LEFT JOIN mae_documento ON com_ordencompra.idtipdoc = mae_documento.id) LEFT JOIN com_usuario ON com_ordencompra.idsol = com_usuario.id) " _
            & " LEFT JOIN pla_empleados ON com_usuario.idper = pla_empleados.id) LEFT JOIN mae_area ON com_ordencompra.idare = mae_area.id) " _
            & " LEFT JOIN mae_condpago ON com_ordencompra.idconpag = mae_condpago.id) LEFT JOIN mae_moneda ON com_ordencompra.idmon = mae_moneda.id) " _
            & " LEFT JOIN mae_prov ON com_ordencompra.idpro = mae_prov.id) LEFT JOIN mae_estados ON com_ordencompra.idest = mae_estados.id) " _
            & " LEFT JOIN com_ordencot ON com_ordencompra.idcot = com_ordencot.id ORDER BY [com_ordencompra]![numser] & '-' & [com_ordencompra]![numdoc]", xCon
        
            Set Dg1.DataSource = RstLista
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
    CaracteresNumericos = "0123456789." & Chr(8) & Chr(13)
    Fg1.ColWidth(8) = 0
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
    'Fg1.ColWidth(10) = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Fg1.Editable = flexEDNone
    Fg1.SelectionMode = flexSelectionByRow
End Sub

Sub Blanquea()
    TxtNumCoti.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtidSol.Text = ""
    TxtIdArea.Text = ""
    TxtFchEmi.Valor = ""
    TxtFchEnt.Valor = ""
    TxtNumRuc.Text = ""
    TxtIdConPag.Text = ""
    TxtIdMon.Text = ""
    
    LblSolicitante.Caption = ""
    LblArea.Caption = ""
    LblProveedor.Caption = ""
    LblIdProv.Caption = ""
    LblCondPag.Caption = ""
    LblMoneda.Caption = ""
    
    LblBruto.Caption = ""
    LblIgv.Caption = ""
    LblTotal.Caption = ""
End Sub

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To 15
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Nuevo()
    Label5.Caption = "Agregando Orden de Compra"
    QueHace = 1
    xHorIni = Time
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Blanquea
    
    TxtNumSer.Text = "0001"
    TxtNumDoc.Text = HallaNumOrdenCompra(TxtNumSer.Text)
    Option2.Value = True
    Option2_Click
    Fg1.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
    Fg1.ColComboList(1) = "|..."
    Fg1.ColComboList(2) = "|..."
    Fg1.ColComboList(3) = "|..."
    LblEstado.Caption = ""
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    TxtIdMon.SetFocus
End Sub

Sub Modificar()
    Label5.Caption = "Modificando Orden de Compra"
    QueHace = 2
    xHorIni = Time
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Blanquea
    Fg1.Rows = 1
    Fg1.ColComboList(2) = "|..."
    Fg1.ColComboList(7) = "|..."
    
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    MuestraSegundoTab
    Fg1.Rows = Fg1.Rows + 1
    TxtIdMon.SetFocus
End Sub

Function HallaNumOrdenCompra(NumSer As String) As String
    Dim xRs As New ADODB.Recordset
    
    RST_Busq xRs, "SELECT * FROM com_ordencompra WHERE numser = '" & NumSer & "'", xCon
    
    If xRs.RecordCount = 0 Then
        HallaNumOrdenCompra = "0000000001"
    Else
        xRs.MoveLast
        HallaNumOrdenCompra = Format(Val(xRs("numdoc")) + 1, "0000000000")
    End If
    Set xRs = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando un registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Option1_Click()
    Fg1.Rows = 1
    TxtNumCoti.Visible = True
    CmdBusCoti.Visible = True
    Label13.Visible = True
    
    CmdAddItem.Enabled = False
    CmdDelItem.Enabled = False
    CmdAddNewItem.Enabled = False
    TxtIdConPag.Locked = False
    
    ActivarTextos True
End Sub

Private Sub Option2_Click()
    Fg1.Rows = 1
    TxtNumCoti.Visible = False
    CmdBusCoti.Visible = False
    Label13.Visible = False
    
    CmdAddItem.Enabled = True
    CmdDelItem.Enabled = True
    CmdAddNewItem.Enabled = True
    
    ActivarTextos False
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
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
            RstLista.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstLista.Filter = ""
    End If
    
    If Button.Index = 12 Then
        If RstLista("idcond") = 1 Then
            Imprimir RstLista("id"), 2
        Else
            MsgBox "No se puede enviar una orden de compra que no este aprobada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    
    If Button.Index = 13 Then Imprimir RstLista("id"), 1
    
    If Button.Index = 15 Then
        Set RstLista = Nothing
        Unload Me
    End If
End Sub

Sub Eliminar()
    Dim Rpta As Integer
        
    If RstLista("idest") = 3 Then
        MsgBox "No se puede eliminar una orden de compra procesada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
        
    Rpta = MsgBox("Esta seguro de eliminar el registro seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        
        ' PRGUNTAMOS SI LA ORDEN DE COMPRA ESTA VINCULADA A UNA ORDEN DE COTIZACION
        If RstLista("tipo") = 1 Then
            ' ACTUALIZAMOS LA ORDEN DE COTIZACION VINCULADA A PENDIENTE Y NO PROCESADA EN LOS CAMPOS idest y idsit
            xCon.Execute "UPDATE com_ordencot SET com_ordencot.idest = 1, com_ordencot.idsit = 0 WHERE (((com_ordencot.id)=" & RstLista("idcot") & "));"
        End If
        
        xCon.Execute "DELETE com_ordencompradet.* From com_ordencompradet WHERE (((com_ordencompradet.idoc)=" & RstLista("id") & "))"
        xCon.Execute "DELETE com_ordencompra.*  From com_ordencompra WHERE (((com_ordencompra.id)=" & RstLista("id") & "))"
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstLista("id") & " AND idform = " & IdMenuActivo

        MsgBox "El registro se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstLista.Requery
        Dg1.Refresh

    End If
End Sub

Private Sub TxtIdArea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdArea_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusArea_Click
    End If
End Sub

Private Sub TxtIdArea_Validate(Cancel As Boolean)
    If NulosN(TxtIdArea.Text) = 0 Then
        LblArea.Caption = ""
        Exit Sub
    End If
    
    LblArea.Caption = Busca_Codigo(TxtIdArea.Text, "id", "descripcion", "mae_area", "N", xCon)
    If NulosC(LblArea.Caption) = "" Then
        TxtIdArea.Text = ""
        LblArea.Caption = ""
    End If
End Sub

Private Sub TxtIdConPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdConPag_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCondPag_Click
    End If
End Sub

Private Sub TxtIdConPag_Validate(Cancel As Boolean)
    If NulosN(TxtIdConPag.Text) = 0 Then
        LblCondPag.Caption = ""
        TxtIdConPag.Text = ""
        Exit Sub
    End If
    
    LblCondPag.Caption = Busca_Codigo(TxtIdConPag.Text, "id", "descripcion", "mae_condpago", "N", xCon)
    If NulosC(LblCondPag.Caption) = "" Then
        TxtIdConPag.Text = ""
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If NulosN(TxtIdMon.Text) = 0 Then
        LblMoneda.Caption = ""
        Exit Sub
    End If
    
    LblMoneda.Caption = Busca_Codigo(TxtIdMon.Text, "id", "descripcion", "mae_moneda", "N", xCon)
    If NulosC(LblMoneda.Caption) = "" Then
        TxtIdMon.Text = ""
        LblMoneda.Caption = ""
    End If
End Sub

Private Sub TxtIdSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdSol_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusSol_Click
    End If
End Sub

Private Sub TxtIdSol_Validate(Cancel As Boolean)
    If NulosN(TxtidSol.Text) = 0 Then
        LblSolicitante.Caption = ""
        Exit Sub
    End If
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT com_usuario.id, com_usuario.idper, UCase([pla_empleados]![apepat]) & ' ' & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS apenom " _
        & " FROM com_usuario LEFT JOIN pla_empleados ON com_usuario.idper = pla_empleados.id WHERE (((com_usuario.id)=" & NulosN(TxtidSol.Text) & "))", xCon

    LblSolicitante.Caption = Rst("apenom")
    If NulosC(LblSolicitante.Caption) = "" Then
        TxtidSol.Text = ""
        LblSolicitante.Caption = ""
    End If
End Sub

Private Sub TxtNumCoti_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumCoti_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then

    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumRuc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusPro_Click
    End If
End Sub

Private Sub TxtNumRuc_Validate(Cancel As Boolean)
    If NulosC(TxtNumRuc.Text) = "" Then
        LblProveedor.Caption = ""
        LblIdProv.Caption = ""
        Exit Sub
    End If
    
    Dim xRs As New ADODB.Recordset
    
    RST_Busq xRs, "SELECT * FROM mae_prov WHERE numruc like '" & TxtNumRuc.Text & "%'", xCon
    If xRs.RecordCount <> 0 Then
        xRs.MoveFirst
        TxtNumRuc.Text = xRs("numruc")
        LblProveedor.Caption = xRs("nombre")
        LblIdProv.Caption = xRs("id")
    Else
        LblProveedor.Caption = ""
        LblIdProv.Caption = ""
        Exit Sub
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Function Grabar() As Boolean
    Dim xCampos(15, 5) As String
    Dim xCampos2(7, 5) As String
    Dim xId As Double
    Dim xEstado As Integer
    Dim A, B As Integer
    
    ' ELIMINAMOS LAS FILAS EN BLANCO DEL CONTROL Fg1
    For A = 1 To Fg1.Rows - 1
        If Fg1.TextMatrix(A, 3) = "" Then
            Fg1.RemoveItem A
        End If
    Next A
    
On Error GoTo LaCague
    xCon.BeginTrans
    If QueHace = 1 Then
        xId = HallaCodigoTabla("com_ordencompra", xCon, "id")
        xEstado = 1
    Else
        xId = RstLista("id")
        xEstado = RstLista("idest")
    End If
    
    'Columna    | Descripcion
    '------------------------
    '0          | campo
    '1          | Valor
    '2          | requerido
    '3          | tipo
    '4          | msj error
    '5          | INDICA QUE EL CAMPO ES INDICE Y NO SE ESCRIBIRA CUANDO SE MODIFIQUE EL REGISTRO
    '--------------------------------
    
    Dim xTipo As Integer
    If Option1.Value = True Then
        ' CON COTIZACION
        xTipo = 1
    End If
    If Option2.Value = True Then
        ' SIN COTIZACION
        xTipo = 2
    End If
    'GRABAMOS LA CABECERA DE LA ORDEN DE COMPRA
    xCampos(0, 0) = "id":         xCampos(0, 1) = Str(xId):                        xCampos(0, 2) = "S":    xCampos(0, 3) = "N":    xCampos(0, 4) = "":                                                                    xCampos(0, 5) = "S"
    xCampos(1, 0) = "tipo":       xCampos(1, 1) = xTipo:                           xCampos(1, 2) = "S":    xCampos(1, 3) = "N":    xCampos(1, 4) = "":                                                                    xCampos(1, 5) = ""
    xCampos(2, 0) = "idcot":      xCampos(2, 1) = NulosN(LblIdCotizacion.Caption): xCampos(2, 2) = "S":    xCampos(2, 3) = "N":    xCampos(2, 4) = "":                                                                    xCampos(2, 5) = ""
    xCampos(3, 0) = "idtipdoc":   xCampos(3, 1) = "92":                            xCampos(3, 2) = "S":    xCampos(3, 3) = "N":    xCampos(3, 4) = "":                                                                    xCampos(3, 5) = ""
    xCampos(4, 0) = "numser":     xCampos(4, 1) = NulosC(TxtNumSer.Text):          xCampos(4, 2) = "S":    xCampos(4, 3) = "C":    xCampos(4, 4) = "":                                                                    xCampos(4, 5) = ""
    xCampos(5, 0) = "numdoc":     xCampos(5, 1) = NulosC(TxtNumDoc.Text):          xCampos(5, 2) = "S":    xCampos(5, 3) = "C":    xCampos(5, 4) = "":                                                                    xCampos(5, 5) = ""
    xCampos(6, 0) = "idsol":      xCampos(6, 1) = NulosN(TxtidSol.Text):           xCampos(6, 2) = "S":    xCampos(6, 3) = "N":    xCampos(6, 4) = "No ha especificado el nombre del solicitante de la orden de compra":  xCampos(6, 5) = ""
    xCampos(7, 0) = "idare":      xCampos(7, 1) = NulosN(TxtIdArea.Text):          xCampos(7, 2) = "S":    xCampos(7, 3) = "N":    xCampos(7, 4) = "No ha especificado el area que solicita la orden de compra":          xCampos(7, 5) = ""
    xCampos(8, 0) = "fchemi":     xCampos(8, 1) = TxtFchEmi.Valor:                 xCampos(8, 2) = "S":    xCampos(8, 3) = "F":    xCampos(8, 4) = "No ha especificado la fecha de emision":                              xCampos(8, 5) = ""
    If Option1.Value = True Then
        xCampos(9, 0) = "fchent":     xCampos(9, 1) = TxtFchEnt.Valor:                 xCampos(9, 2) = "N":    xCampos(9, 3) = "F":    xCampos(9, 4) = "No ha especificado la fecha de entrega":                              xCampos(9, 5) = ""
    Else
        xCampos(9, 0) = "fchent":     xCampos(9, 1) = TxtFchEnt.Valor:                 xCampos(9, 2) = "S":    xCampos(9, 3) = "F":    xCampos(9, 4) = "No ha especificado la fecha de entrega":                              xCampos(9, 5) = ""
    End If
    xCampos(10, 0) = "idpro":     xCampos(10, 1) = NulosN(LblIdProv.Caption):      xCampos(10, 2) = "S":   xCampos(10, 3) = "N":   xCampos(10, 4) = "No ha especificado el nombre del proveedor":                         xCampos(10, 5) = ""
    xCampos(11, 0) = "idconpag":  xCampos(11, 1) = NulosN(TxtIdConPag.Text):       xCampos(11, 2) = "S":   xCampos(11, 3) = "N":   xCampos(11, 4) = "No ha especificado la condicion de pago":                            xCampos(11, 5) = ""
    xCampos(12, 0) = "idmon":     xCampos(12, 1) = NulosN(TxtIdMon.Text):          xCampos(12, 2) = "S":   xCampos(12, 3) = "N":   xCampos(12, 4) = "No ha especificado la moneda":                                       xCampos(12, 5) = ""
    xCampos(13, 0) = "idcond":    xCampos(13, 1) = "2":                            xCampos(13, 2) = "S":   xCampos(13, 3) = "N":   xCampos(13, 4) = "":                                                                   xCampos(13, 5) = ""
    xCampos(14, 0) = "envcor":    xCampos(14, 1) = 0:                              xCampos(14, 2) = "N":   xCampos(14, 3) = "N":   xCampos(14, 4) = "":                                                                   xCampos(14, 5) = ""
    xCampos(15, 0) = "idest":     xCampos(15, 1) = xEstado:                        xCampos(15, 2) = "N":   xCampos(15, 3) = "N":   xCampos(15, 4) = "":                                                                   xCampos(15, 5) = ""
    
    If QueHace = 1 Then
        If EscribirNuevoRegistro(xCampos, "com_ordencompra", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Else
        ' ELIMINAMOS LOS DETALLES DE LA ORDEN DE COMPRA
        xCon.Execute "DELETE * FROM com_ordencompradet WHERE idoc = " & RstLista("id") & ""
        
        ' MODIFICAMOS EL REGISTRO
        If ModificarRegistro(xCampos, "com_ordencompra", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    End If
    
    ' GRABAMOS EL DETALLE DE LA ORDEN DE REQUERIMIENTO
    For A = 1 To Fg1.Rows - 1
        xCampos2(0, 0) = "idoc":           xCampos2(0, 1) = Str(xId):                   xCampos2(0, 2) = "S":    xCampos2(0, 3) = "N":    xCampos2(0, 4) = "":      xCampos2(0, 5) = ""
        xCampos2(1, 0) = "iditem":         xCampos2(1, 1) = Fg1.TextMatrix(A, 8):       xCampos2(1, 2) = "S":    xCampos2(1, 3) = "N":    xCampos2(1, 4) = "":      xCampos2(1, 5) = ""
        xCampos2(2, 0) = "idcencos":       xCampos2(2, 1) = 0:                          xCampos2(2, 2) = "N":    xCampos2(2, 3) = "N":    xCampos2(2, 4) = "":      xCampos2(2, 5) = ""
        xCampos2(3, 0) = "idunimed":       xCampos2(3, 1) = Fg1.TextMatrix(A, 9):       xCampos2(3, 2) = "S":    xCampos2(3, 3) = "N":    xCampos2(3, 4) = "":      xCampos2(3, 5) = ""
        xCampos2(4, 0) = "cantidad":       xCampos2(4, 1) = Fg1.TextMatrix(A, 5):       xCampos2(4, 2) = "S":    xCampos2(4, 3) = "N":    xCampos2(4, 4) = "":      xCampos2(4, 5) = ""
        xCampos2(5, 0) = "impuni":         xCampos2(5, 1) = Fg1.TextMatrix(A, 6):       xCampos2(5, 2) = "S":    xCampos2(5, 3) = "N":    xCampos2(5, 4) = "":      xCampos2(5, 5) = ""
        xCampos2(6, 0) = "idtip":          xCampos2(6, 1) = Fg1.TextMatrix(A, 10):      xCampos2(6, 2) = "N":    xCampos2(6, 3) = "N":    xCampos2(6, 4) = "":      xCampos2(6, 5) = ""
        xCampos2(7, 0) = "idequi":         xCampos2(7, 1) = Fg1.TextMatrix(A, 11):      xCampos2(7, 2) = "N":    xCampos2(7, 3) = "N":    xCampos2(7, 4) = "":      xCampos2(7, 5) = ""
        
        If EscribirNuevoRegistro(xCampos2, "com_ordencompradet", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Next A
       
    If Option1.Value = True Then
        ' CON COTIZACION
        ' ACTUALIZAMOS EL CAMPO idest y idsit DE LA TABLA com_ordencot
        xCon.Execute " UPDATE com_ordencot SET com_ordencot.idest = 3, com_ordencot.idsit = 1 WHERE (((com_ordencot.id)=" & NulosN(LblIdCotizacion.Caption) & "))"
    End If
       
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId

    
    xCon.CommitTrans
    MsgBox "El registro se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function

LaCague:
    'Resume
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente motivo: " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = False
End Function

Sub Cancelar()
    QueHace = 3
    
    Label5.Caption = "Detalle de la Orden de Compra"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.Editable = flexEDNone
    Fg1.SelectionMode = flexSelectionByRow
    ActivaTool
    ActivarTextos False
End Sub

Sub MuestraSegundoTab()
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    
    If RstLista("tipo") = 1 Then
        Option1.Value = True
    Else
        Option2.Value = True
    End If
    
    If NulosN(RstLista("idcot")) <> 0 Then
        Dim xRs As New ADODB.Recordset
        
        RST_Busq xRs, "SELECT [com_ordencot]![numser] & '-' & [com_ordencot]![numdoc] AS numdoc2 From com_ordencot WHERE (((com_ordencot.id)=" & RstLista("idcot") & "))", xCon
        If xRs.RecordCount <> 0 Then
            TxtNumCoti.Text = xRs("numdoc2")
        End If
        LblIdCotizacion.Caption = RstLista("idcot")
    Else
        TxtNumCoti.Text = ""
        LblIdCotizacion.Caption = ""
    End If
        
    TxtNumSer.Text = RstLista("numser")
    TxtNumDoc.Text = RstLista("numdoc")
    TxtidSol.Text = RstLista("idsol")
    LblSolicitante.Caption = RstLista("nomsol")
    TxtIdArea.Text = RstLista("idare")
    LblArea.Caption = RstLista("descarea")
    TxtFchEmi.Valor = RstLista("fchemi")
    If IsNull(RstLista("fchent")) = True Then
        TxtFchEnt.Valor = ""
    Else
        TxtFchEnt.Valor = RstLista("fchent")
    End If
    LblIdProv.Caption = RstLista("idpro")
    TxtNumRuc.Text = RstLista("numruc")
    LblProveedor.Caption = RstLista("nombre")
    TxtIdConPag.Text = RstLista("idconpag")
    LblCondPag.Caption = Busca_Codigo(TxtIdConPag.Text, "id", "descripcion", "mae_condpago", "N", xCon)
    TxtIdMon.Text = RstLista("idmon")
    LblMoneda.Caption = NulosC(RstLista("descmon"))
    
    If RstLista("idest") = 1 Then
        LblEstado.Caption = "Pendiente"
        LblEstado.ForeColor = &H8000&        ' Verde
    End If
    
    If RstLista("idest") = 2 Then
        LblEstado.Caption = "Aprobada"
        LblEstado.ForeColor = &HC00000       ' Azul
    End If
        
        
    If RstLista("idest") = 4 Then
        LblEstado.Caption = "Rechazada"
        LblEstado.ForeColor = &HC0&          ' Rojo
    End If
        
    Fg1.Rows = 1
    RST_Busq RstDet, "SELECT com_ordencompradet.*, alm_inventario.descripcion AS descitem, con_centrocosto.descripcion AS desccencos, mae_unidades.abrev AS descunimed, " _
        & " man_tipo.descripcion AS desctipo, man_equipos.nombre AS descequi FROM ((((com_ordencompradet LEFT JOIN alm_inventario ON com_ordencompradet.iditem = alm_inventario.id) " _
        & " LEFT JOIN con_centrocosto ON com_ordencompradet.idcencos = con_centrocosto.id) LEFT JOIN mae_unidades ON com_ordencompradet.idunimed = mae_unidades.id) " _
        & " LEFT JOIN man_tipo ON com_ordencompradet.idtip = man_tipo.id) LEFT JOIN man_equipos ON com_ordencompradet.idequi = man_equipos.id " _
        & " Where (((com_ordencompradet.idoc) = 1)) ORDER BY alm_inventario.descripcion", xCon

    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        
        For A = 1 To RstDet.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(RstDet("desctipo"))
            Fg1.TextMatrix(A, 2) = NulosC(RstDet("descequi"))
            Fg1.TextMatrix(A, 3) = RstDet("descitem")
            Fg1.TextMatrix(A, 4) = RstDet("descunimed")
            Fg1.TextMatrix(A, 5) = Format(RstDet("cantidad"), "0.00")
            Fg1.TextMatrix(A, 6) = Format(RstDet("impuni"), "0.00")
            Fg1.TextMatrix(A, 7) = (RstDet("cantidad") * RstDet("impuni"))
            Fg1.TextMatrix(A, 7) = Format(Fg1.TextMatrix(A, 7), "0.00")
            Fg1.TextMatrix(A, 8) = RstDet("iditem")
            Fg1.TextMatrix(A, 9) = RstDet("idunimed")
            Fg1.TextMatrix(A, 10) = RstDet("idtip")
            Fg1.TextMatrix(A, 11) = RstDet("idequi")
            RstDet.MoveNext
            If RstDet.EOF = True Then Exit For
        Next A
    End If
    SumarTotal
End Sub

Sub Imprimir(IdCotizacion As Integer, Opcion As Integer)
    ' OPCION = 1 SE ABRE EL DOCUMENTO PDF
    ' OPCION = 2 SE ENVIA POR CORREO EL ARCHIVO PDF
    Dim Li As Integer
    Dim strSource As String
    Dim xArea, xEmp, xDir, xCuerpo, xCad  As String
    Dim xEmpleado As String
    Dim Pagina As Integer
    Dim Lineas As Integer
    Dim xNomPro, xNumRUCPro, xNumCot As String
    
    Set oPDF = New cPDF
    Pagina = 0
    xNomPro = Busca_Codigo(RstLista("idpro"), "id", "nombre", "mae_prov", "N", xCon)
    xNumRUC = Busca_Codigo(RstLista("idpro"), "id", "numruc", "mae_prov", "N", xCon)
    xNumCot = "0000001"
    If oPDF.PDFCreate(App.Path & "\OC" & RstLista("numdoc2") & ".pdf") = True Then
        oPDF.Fonts.Add "Tit", Times_BoldItalic, WinAnsiEncoding
        oPDF.Fonts.Add "Head", Times_Italic, WinAnsiEncoding
        oPDF.Fonts.Add "Cont", Courier, WinAnsiEncoding
        oPDF.Fonts.Add "Time", Times_Roman, WinAnsiEncoding
        
        CrearCabecera RstLista("numdoc2")
        xCad = xDisEmp & " " & Format(RstLista("fchemi"), "dd") & " de " & Format(RstLista("fchemi"), "mmmm") & " del " & Format(RstLista("fchemi"), "yyyy")
        
        oPDF.WTextBox 100, 55, 10, 420, xCad, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        
        oPDF.WTextBox 120, 55, 10, 50, "Proveedor", "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        oPDF.WTextBox 120, 107, 10, 373, ": " & xNomPro, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        
        oPDF.WTextBox 135, 55, 10, 50, "Nº R.U.C.", "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        oPDF.WTextBox 135, 107, 10, 200, ": " & xNumRUCPro, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        
        oPDF.WTextBox 135, 310, 10, 70, "Ref. A Cot. Nº", "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        oPDF.WTextBox 135, 385, 10, 95, ": " & xNumCot, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack

        oPDF.WTextBox 155, 55, 10, 50, "Nº R.U.C.", "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        oPDF.WTextBox 155, 107, 10, 200, ": " & xNumRUC, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        
        oPDF.WTextBox 170, 55, 10, 50, "Facturar A", "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        oPDF.WTextBox 170, 107, 10, 373, ": " & xNomEmp, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        
        oPDF.WTextBox 185, 55, 10, 50, "Direccion", "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        oPDF.WTextBox 185, 107, 10, 373, ": " & xDirEmp, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        
        oPDF.WTextBox 200, 55, 10, 50, "Fch. Emi.", "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        oPDF.WTextBox 200, 107, 10, 200, ": " & RstLista("fchemi"), "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        
        oPDF.WTextBox 200, 310, 10, 70, "Fch. Entº", "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        oPDF.WTextBox 200, 385, 10, 95, ": " & RstLista("fchent"), "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        
        oPDF.WTextBox 215, 55, 10, 50, "Moneda", "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        oPDF.WTextBox 215, 107, 10, 200, ": " & RstLista("descmon"), "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        
        oPDF.WTextBox 215, 310, 10, 70, "Condicion", "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        oPDF.WTextBox 215, 385, 10, 95, ": " & RstLista("desccondpag"), "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack

'
'        ' ESCRIBIMOS EL CONTENIDO DEL CUERPO
'        xCuerpo = "Por medio de la presente le saludamos y solicitamos nos envié en el mas breve plazo la cotización de los siguientes ítems"
'        oPDF.WTextBox 170, 55, 10, 420, xCuerpo, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
'
        oPDF.WTextBox 230, 55, 18, 19, "Item", "Time", 9, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oPDF.WTextBox 230, 76, 18, 250, "Descripcion", "Time", 9, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oPDF.WTextBox 230, 327, 18, 29, "Uni. Med.", "Time", 9, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oPDF.WTextBox 230, 358, 18, 38, "Cantidad", "Time", 9, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oPDF.WTextBox 230, 397, 18, 38, "P. Unitario", "Time", 9, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oPDF.WTextBox 230, 436, 18, 44, "Total", "Time", 9, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        
        Dim Rst As New ADODB.Recordset
        Dim A, Fila As Integer
        Dim xTotal As Double
        
        RST_Busq Rst, "SELECT com_ordencompradet.*, alm_inventario.descripcion AS descpro, mae_unidades.abrev AS descunimed " _
            & " FROM (com_ordencompradet LEFT JOIN alm_inventario ON com_ordencompradet.iditem = alm_inventario.id) LEFT JOIN mae_unidades " _
            & " ON com_ordencompradet.idunimed = mae_unidades.id WHERE (((com_ordencompradet.idoc)=" & RstLista("id") & "))", xCon

        
        If Rst.RecordCount <> 0 Then
            Fila = 250
            Rst.MoveFirst
            For A = 1 To 15 'Rst.RecordCount
                If Rst.EOF = False Then
                    oPDF.WTextBox Fila, 55, 10, 19, A, "Time", 9, hCenter, vMiddle, vbBlack, , vbBlack
                    oPDF.WTextBox Fila, 76, 10, 250, Rst("descpro"), "Time", 9, hLeft, vMiddle, vbBlack, , vbBlack
                    oPDF.WTextBox Fila, 327, 10, 29, Rst("descunimed"), "Time", 9, hCenter, vMiddle, vbBlack, , vbBlack
                    oPDF.WTextBox Fila, 358, 10, 38, Format(Rst("cantidad"), "0.00"), "Time", 9, hRight, vMiddle, vbBlack, , vbBlack
                    oPDF.WTextBox Fila, 397, 10, 38, Format(Rst("impuni"), "0.00"), "Time", 9, hRight, vMiddle, vbBlack, , vbBlack
                    oPDF.WTextBox Fila, 436, 10, 44, Format(Rst("cantidad") * Rst("impuni"), "0.00"), "Time", 9, hRight, vMiddle, vbBlack, , vbBlack
                    xTotal = xTotal + (Rst("impuni") * Rst("cantidad"))
                    Rst.MoveNext
                End If
                Fila = Fila + 10
                
                
                'If Rst.EOF = False Then Rst.MoveNext
            Next A
            Fila = Fila + 10
            oPDF.WTextBox Fila, 358, 10, 77, "Imp. Bruto", "Time", 9, hLeft, vMiddle, vbWhite, 1, vbBlack, True
            oPDF.WTextBox Fila, 436, 10, 44, Format(xTotal, "0.00"), "Time", 9, hRight, vMiddle, vbBlack, , vbBlack
            Fila = Fila + 10
            oPDF.WTextBox Fila, 358, 10, 77, "I.G.V.", "Time", 9, hLeft, vMiddle, vbWhite, 1, vbBlack, True
            oPDF.WTextBox Fila, 436, 10, 44, Format((xTotal * 1.19) - xTotal, "0.00"), "Time", 9, hRight, vMiddle, vbBlack, , vbBlack
            Fila = Fila + 10
            oPDF.WTextBox Fila, 358, 10, 77, "Imp. Total", "Time", 9, hLeft, vMiddle, vbWhite, 1, vbBlack, True
            oPDF.WTextBox Fila, 436, 10, 44, Format(xTotal * 1.19, "0.00"), "Time", 9, hRight, vMiddle, vbBlack, , vbBlack
        End If
        Fila = Fila + 20
        oPDF.WTextBox Fila, 55, 10, 425, NumeroLetra(Format(xTotal * 1.19, "0.00"), RstLista("idmon")), "Time", 9, hLeft, vMiddle, vbBlack, , vbBlack
        'Dim xCad As String
        
        xCad = "· Realice este pedido de acuerdo con los precios, términos, método de entrega y especificaciones enumeradas anteriormente."
        Fila = Fila + 20
        oPDF.WTextBox Fila, 55, 20, 425, xCad, "Time", 9, hLeft, vMiddle, vbBlack, , vbBlack
        
        xCad = "· Si no puede validar la orden de compra avisar por favor inmediatamente."
        Fila = Fila + 20
        oPDF.WTextBox Fila, 55, 10, 425, xCad, "Time", 9, hLeft, vMiddle, vbBlack, , vbBlack

        Fila = Fila + 20
        oPDF.WTextBox Fila, 55, 10, 425, "Atentamente", "Time", 9, hCenter, vMiddle, vbBlack, , vbBlack
        
'        ' ESCRIBIMOS LA FIRMA DEL ENCARGADO
        Fila = Fila + 60
        xCuerpo = "--------------------------------"
        oPDF.WTextBox Fila, 55, 10, 420, xCuerpo, "Time", 9, hCenter, vMiddle, vbBlack, , vbBlack
        Fila = Fila + 10
        xEmpleado = "Juan Perez Martinez"
        oPDF.WTextBox Fila, 55, 10, 420, xEmpleado, "Time", 9, hCenter, vMiddle, vbBlack, , vbBlack
        Fila = Fila + 10
        xCuerpo = "Jefe de Compras"
        oPDF.WTextBox Fila, 55, 10, 420, xCuerpo, "Time", 9, hCenter, vMiddle, vbBlack, , vbBlack
        
        oPDF.PDFClose
        Set oPDF = Nothing
        
        If Opcion = 1 Then
            Shell ("rundll32.exe url.dll,FileProtocolHandler " & Trim(App.Path) & ("\OC" & RstLista("numdoc2") & ".pdf")), vbMaximizedFocus
        End If
        
        If Opcion = 2 Then
            Dim xIdPro As Integer
            Dim eMail As String
            xIdPro = Busca_Codigo(IdCotizacion, "id", "idpro", "com_ordencompra", "N", xCon)
            eMail = NulosC(Busca_Codigo(xIdPro, "id", "email", "mae_prov", "N", xCon))
            If NulosC(eMail) = "" Then
                MsgBox "EL proveedor " & Trim(xEmp) & " no tiene correo electronico, agregue la direccion de correo electronico del proveedor para efectuar esta operacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
            Dim xFun As New eps_librerias.Correo
            Dim xAdjunto(2) As String
            xFun.ServidorSMTP = "mail.agro-vado.com"
            xFun.NomRemitente = "Sistema de Compras"
            xFun.MailRemitente = "seven@seven.com"
            xFun.MailDestino = eMail
            xFun.Asunto = "Orden de Compra Nº " & RstLista("numdoc2")
            xFun.Cuerpo = "Buenos dias remito orden de compra favor de enviar a la brevedad prosible"
            
            xAdjunto(0) = Trim(App.Path) & "\OC" & RstLista("numdoc2") & ".pdf"
            If xFun.EnviarCorreo(xAdjunto) = True Then
                xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.envcor = -1 WHERE (((com_ordencompra.id)=" & RstLista("id") & "))"
                RstLista.Requery
                Dg1.Refresh
            End If
        End If
    Else
        MsgBox "No se Puede Mostrar Documento", vbCritical, "Error"
    End If
End Sub

Sub CrearCabecera(NumDoc As String)
    Dim xTelEmp, xNumDoc As String
    
    xTelEmp = "Telf: 493-0808   Tele Fax: 295-6868"
    xNumDoc = NumDoc

    oPDF.NewPage UsarAnchoAlto, 525, 675
    Pagina = Pagina + 1
    oPDF.WTextBox 32, 55, 20, 250, xNomEmp, "Tit", 12, hCenter, vMiddle, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 55, 55, 10, 250, xDirEmp, "Head", 9, hCenter, vMiddle, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 65, 55, 10, 250, xTelEmp, "Head", 9, hCenter, vMiddle, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 75, 55, 10, 250, xPagEmp, "Head", 9, hCenter, vMiddle, RGB(0, 0, 128), , vbRed
    
    oPDF.WTextBox 46, 330, 10, 150, "ORDEN DE COMPRA", "Head", 10, hCenter, vMiddle, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 60, 330, 10, 150, "Nº " & xNumDoc, "Head", 10, hCenter, vMiddle, RGB(0, 0, 128), , vbRed
    
    oPDF.WRectangle 32, 330, 53, 150, 1.5, vbBlack
End Sub

