VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmProvisiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Asientos Diversos"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   11
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
         TabIndex        =   15
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdBusLib 
            Enabled         =   0   'False
            Height          =   240
            Left            =   2220
            Picture         =   "FrmProvisiones.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   420
            Width           =   210
         End
         Begin VB.CommandButton cb 
            Height          =   240
            Index           =   1
            Left            =   2220
            Picture         =   "FrmProvisiones.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   765
            Width           =   210
         End
         Begin VB.Frame Frame6 
            Caption         =   "[ Tipo de Cambio ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   690
            Left            =   9450
            TabIndex        =   38
            Top             =   1110
            Width           =   2250
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
               Left            =   120
               TabIndex        =   39
               Top             =   300
               Width           =   1980
            End
         End
         Begin VB.Frame Frame4 
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
            ForeColor       =   &H00400000&
            Height          =   720
            Left            =   9450
            TabIndex        =   32
            Top             =   345
            Width           =   2250
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
               TabIndex        =   33
               Top             =   330
               Width           =   1995
            End
         End
         Begin VB.CommandButton cb 
            Height          =   240
            Index           =   0
            Left            =   2220
            Picture         =   "FrmProvisiones.frx":0264
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   1080
            Width           =   210
         End
         Begin VB.CommandButton CmdBusPro 
            Height          =   240
            Left            =   8400
            Picture         =   "FrmProvisiones.frx":0396
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1260
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2640
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   5
            Text            =   "TxtNumDoc"
            Top             =   1995
            Width           =   1830
         End
         Begin VB.TextBox TxtSerDoc 
            Height          =   300
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   4
            Text            =   "TxtSerDoc"
            Top             =   1995
            Width           =   900
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmProvisiones.frx":04C8
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1710
            Width           =   210
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   3585
            Left            =   195
            TabIndex        =   7
            Top             =   2670
            Width           =   9840
            _cx             =   17357
            _cy             =   6324
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmProvisiones.frx":05FA
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
         Begin VB.Frame Frame3 
            Height          =   3675
            Left            =   10140
            TabIndex        =   18
            Top             =   2580
            Width           =   1560
            Begin VB.CommandButton Command1 
               Caption         =   "Agregar Documentos"
               Enabled         =   0   'False
               Height          =   690
               Left            =   120
               TabIndex        =   40
               Top             =   825
               Width           =   1305
            End
            Begin VB.CommandButton CmdAdd 
               Caption         =   "Agregar Cuenta"
               Enabled         =   0   'False
               Height          =   690
               Left            =   120
               TabIndex        =   20
               Top             =   1530
               Width           =   1305
            End
            Begin VB.CommandButton CmdDel 
               Caption         =   "Eliminar Cuenta"
               Enabled         =   0   'False
               Height          =   690
               Left            =   120
               TabIndex        =   19
               Top             =   2235
               Width           =   1305
            End
         End
         Begin VB.TextBox TxtGlosa 
            Height          =   300
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   6
            Text            =   "TxtGlosa"
            Top             =   2310
            Width           =   10125
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   1560
            TabIndex        =   2
            Top             =   1365
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
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   3
            Text            =   "TxtTipDoc"
            Top             =   1680
            Width           =   900
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   0
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "txt_cb(0)"
            ToolTipText     =   "Ingrese DNI del Supervisor"
            Top             =   1050
            Width           =   900
         End
         Begin VB.TextBox TxtNombre 
            Height          =   300
            Left            =   7575
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   10
            Text            =   "TxtNombre"
            Top             =   1230
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Frame Frame5 
            Height          =   570
            Left            =   195
            TabIndex        =   16
            Top             =   6225
            Width           =   11505
            Begin VB.TextBox TxtTotHab 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   8010
               Locked          =   -1  'True
               TabIndex        =   9
               Text            =   "TxtTotHab"
               Top             =   180
               Width           =   1110
            End
            Begin VB.TextBox TxtTotDeb 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   6900
               Locked          =   -1  'True
               TabIndex        =   8
               Text            =   "TxtTotDeb"
               Top             =   180
               Width           =   1110
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
               Left            =   5955
               TabIndex        =   17
               Top             =   210
               Width           =   825
            End
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   1
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   0
            Text            =   "txt_cb(1)"
            ToolTipText     =   "Ingrese DNI del Supervisor"
            Top             =   735
            Width           =   900
         End
         Begin VB.TextBox TxtIdLibro 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   46
            Text            =   "TxtIdLibro"
            Top             =   390
            Width           =   900
         End
         Begin VB.Label lbl_cb_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb_cod(0)"
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
            Index           =   0
            Left            =   4680
            TabIndex        =   29
            Top             =   1050
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label LblLibro 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblLibro"
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
            Left            =   2490
            TabIndex        =   48
            Top             =   390
            Width           =   4020
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Libro"
            Height          =   195
            Index           =   3
            Left            =   195
            TabIndex        =   47
            Top             =   480
            Width           =   345
         End
         Begin VB.Label lbl_cb_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb_cod(1)"
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
            Index           =   1
            Left            =   4680
            TabIndex        =   44
            Top             =   735
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label lbl_cb 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb(1)"
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
            Index           =   1
            Left            =   2490
            TabIndex        =   43
            Top             =   735
            Width           =   4020
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Sub Libro"
            Height          =   195
            Index           =   1
            Left            =   195
            TabIndex        =   42
            Top             =   800
            Width           =   675
         End
         Begin VB.Line Line2 
            BorderWidth     =   5
            X1              =   2550
            X2              =   2535
            Y1              =   2130
            Y2              =   2130
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   4
            Left            =   195
            TabIndex        =   37
            Top             =   2400
            Width           =   405
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Doc."
            Height          =   195
            Index           =   6
            Left            =   195
            TabIndex        =   36
            Top             =   1440
            Width           =   840
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
            Height          =   195
            Index           =   1
            Left            =   195
            TabIndex        =   35
            Top             =   1760
            Width           =   1185
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   34
            Top             =   2080
            Width           =   1050
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   0
            Left            =   195
            TabIndex        =   31
            Top             =   1120
            Width           =   585
         End
         Begin VB.Label lbl_cb 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb(0)"
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
            Index           =   0
            Left            =   2490
            TabIndex        =   30
            Top             =   1050
            Width           =   2055
         End
         Begin VB.Label LblIdCli 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCli"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   8745
            TabIndex        =   26
            Top             =   1335
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Index           =   5
            Left            =   6990
            TabIndex        =   24
            Top             =   1335
            Visible         =   0   'False
            Width           =   555
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
            Left            =   2505
            TabIndex        =   23
            Top             =   1680
            Width           =   4020
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Asiento"
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
            TabIndex        =   21
            Top             =   30
            Width           =   11610
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   45
         TabIndex        =   12
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6360
            Left            =   30
            TabIndex        =   13
            Top             =   405
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11218
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
            Columns(1).Caption=   "T.D."
            Columns(1).DataField=   "destipdoc"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numedoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "M"
            Columns(3).DataField=   "simbolo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fecha"
            Columns(4).DataField=   "fchdoc"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Sub Libro"
            Columns(5).DataField=   "sublibdesc"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Glosa"
            Columns(6).DataField=   "glosa"
            Columns(6).NumberFormat=   "0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Debe"
            Columns(7).DataField=   "totdeb"
            Columns(7).NumberFormat=   "0.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Haber"
            Columns(8).DataField=   "tothab"
            Columns(8).NumberFormat=   "0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1640"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=820"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=741"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=714"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=635"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1508"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1429"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2672"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2593"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=6033"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=5953"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=1879"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1799"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=514"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=1958"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1879"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
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
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.fgcolor=&HFFFFFF&"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=70,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=0"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
            _StyleDefs(72)  =   "Named:id=33:Normal"
            _StyleDefs(73)  =   ":id=33,.parent=0"
            _StyleDefs(74)  =   "Named:id=34:Heading"
            _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(76)  =   ":id=34,.wraptext=-1"
            _StyleDefs(77)  =   "Named:id=35:Footing"
            _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(79)  =   "Named:id=36:Selected"
            _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=37:Caption"
            _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(83)  =   "Named:id=38:HighlightRow"
            _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=39:EvenRow"
            _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(87)  =   "Named:id=40:OddRow"
            _StyleDefs(88)  =   ":id=40,.parent=33"
            _StyleDefs(89)  =   "Named:id=41:RecordSelector"
            _StyleDefs(90)  =   ":id=41,.parent=34"
            _StyleDefs(91)  =   "Named:id=42:FilterBar"
            _StyleDefs(92)  =   ":id=42,.parent=33"
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
            TabIndex        =   27
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
            Caption         =   "Consulta de Asientos Diversos"
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
            TabIndex        =   14
            Top             =   90
            Width           =   11595
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   49
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
         NumButtons      =   17
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
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
               Picture         =   "FrmProvisiones.frx":06B4
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones.frx":0BF8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones.frx":0F8A
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones.frx":110E
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones.frx":1562
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones.frx":167A
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones.frx":1BBE
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones.frx":2102
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones.frx":2216
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones.frx":232A
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones.frx":277E
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones.frx":28EA
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones.frx":2E32
               Key             =   "IMG12"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Eliminar             "
      End
   End
End
Attribute VB_Name = "FrmProvisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim CaracteresNumericos As String
Dim RstFrm As New ADODB.Recordset
Dim Agregando As Boolean
Dim xHorIni As Date

Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta


Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Bloquea()
'    TxtIdLibro.Locked = Not TxtIdLibro.Locked
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNombre.Locked = Not TxtNombre.Locked
    TxtSerDoc.Locked = Not TxtSerDoc.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    
    TxtFchEmi.Locked = Not TxtFchEmi.Locked
    TxtGlosa.Locked = Not TxtGlosa.Locked
    
    CmdAdd.Enabled = Not CmdAdd.Enabled
    CmdDel.Enabled = Not CmdDel.Enabled
    
    habilitar_Locked txt_cb, Not txt_cb(0).Locked
    
End Sub

Sub Blanquea()
    
    'TxtIdLibro.Text = ""
    TxtTipDoc.Text = ""
    TxtNombre.Text = ""
    TxtSerDoc.Text = ""
    TxtNumDoc.Text = ""

    TxtFchEmi.Valor = ""
    TxtGlosa.Text = ""
    

    LblLibro.Caption = ""
    LblDocumento.Caption = ""
    LblIdCli.Caption = ""
    TxtTotDeb.Text = ""
    TxtTotHab.Text = ""
    
    LimpiaText txt_cb, True
    
    LblTipoCambio.Caption = ""
    
End Sub

Private Sub CmdAdd_Click()
    If QueHace = 3 Then Exit Sub
    If fg1.TextMatrix(fg1.Rows - 1, 1) = "" Then
        fg1.Row = fg1.Rows - 1
        fg1.Col = 1
        fg1.SetFocus
        Exit Sub
    End If
    fg1.Rows = fg1.Rows + 1
    fg1.Row = fg1.Rows - 1:        fg1.Col = 1
    fg1.SetFocus
End Sub

Private Sub CmdBusLib_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Documento":    xCampos2(0, 1) = "descripcion":    xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Codigo":       xCampos2(1, 1) = "id":             xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"

    xform.SqlCad = "SELECT * FROM mae_libros"
    xform.Titulo = "Buscando Libros Contables"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtIdLibro.Text = NulosN(xRs("id"))
        LblLibro.Caption = NulosC(xRs("descripcion"))
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDoc_Click()
    If QueHace = 3 Then Exit Sub
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(3, 4) As String
    
    xCampos2(0, 0) = "Documento":    xCampos2(0, 1) = "descripcion":    xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Sigla":        xCampos2(1, 1) = "abrev":          xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"
    xCampos2(2, 0) = "Codigo":       xCampos2(2, 1) = "id":             xCampos2(2, 2) = "1000":         xCampos2(2, 3) = "N"

    xform.SqlCad = "SELECT * FROM mae_documento"
    xform.Titulo = "Buscando Documentos"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtTipDoc.Text = xRs("id")
        LblDocumento.Caption = xRs("descripcion")
        TxtSerDoc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDel_Click()
    If QueHace = 3 Then Exit Sub
    If fg1.Rows = 1 Then
        MsgBox "No hay cuentas para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If fg1.Row < 1 Then Exit Sub
    If MsgBox("Seguro desea elimiar el registro" + vbCr + "N° Cuenta: " + fg1.TextMatrix(fg1.Row, 1) + vbCr + "Descripción: " + fg1.TextMatrix(fg1.Row, 2), vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    fg1.RemoveItem fg1.Row
    HallarTotal
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 14, NulosN(RstFrm("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If QueHace = 3 Then Exit Sub
  
    If Col <> 1 Then Exit Sub

    Dim Rst As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLLike As String
    Dim nSQLIdCta As String
      
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "cuenta":             xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":        xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
    
    If NulosC(fg1.TextMatrix(fg1.Row, 1)) <> "" Then
        fg1.TextMatrix(fg1.Row, 1) = Replace(fg1.TextMatrix(fg1.Row, 1), "'", "")
        fg1.TextMatrix(fg1.Row, 1) = Replace(fg1.TextMatrix(fg1.Row, 1), "*", "")
        fg1.TextMatrix(fg1.Row, 1) = Replace(fg1.TextMatrix(fg1.Row, 1), "LIKE", "")
        
        nSQLLike = " and con_planctas.cuenta like '" + Trim(fg1.TextMatrix(fg1.Row, 1)) + "%' "
        
    End If
    
    
    
    nSQLIdCta = GRID_GENERAR_SQL_ID(fg1, 5, "con_planctas.id", " NOT IN ", True)
    If nSQLIdCta <> "" Then nSQLIdCta = " WHERE " + nSQLIdCta
    
    nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id " _
        + vbCr + " From con_planctas " + nSQLIdCta + nSQLLike + vbCr + "  ORDER BY con_planctas.cuenta"
    
    
    CARGAR_DLL_EPSBUSCAR xCon, Rst, nSQL, xCampos(), "Buscando Cuentas Contables", "cuenta", "cuenta", Principio
    
    If Rst.State = 0 Then GoTo SALIR
    If Rst.RecordCount = 0 Then GoTo SALIR
       
    RST_Busq xRs, "SELECT id, cuenta FROM con_planctas WHERE (((id)<>" + Trim(Rst("id")) + ") AND ((cuenta) Like '" + Trim(Rst("cuenta")) + "%'));", xCon
    If xRs.EOF = False Or xRs.BOF = False Or xRs.RecordCount <> 0 Then
        MsgBox "Cuenta no válida" + vbCr + "Seleccione una divisionaria", vbExclamation, xTitulo
        GoTo SALIR
        Exit Sub
    End If
    Agregando = True


    If GRID_BUSCAR_VALOR(fg1, 1, Trim(Rst("cuenta")), False, , Row) <> "-1" Then
        MsgBox "La Cuenta " + Trim(Rst("cuenta")) + " ya esta en la Lista" + vbCr + "Seleccione otra", vbExclamation, xTitulo
        GoTo SALIR
    End If
    fg1.TextMatrix(fg1.Row, 1) = NulosC(Rst("cuenta"))
    fg1.TextMatrix(fg1.Row, 2) = NulosC(Rst("descripcion"))
    fg1.TextMatrix(fg1.Row, 5) = NulosN(Rst("id"))

    Set Rst = Nothing
    Set xRs = Nothing
SALIR:
    
    Agregando = False
    Exit Sub
error:
    Set Rst = Nothing
    Set xRs = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "Fg1_CellButtonClick"
End Sub

Sub HallarTotal()
    TxtTotDeb.Text = Format(GRID_SUMAR_COL(fg1, 3), FORMAT_MONTO)
    TxtTotHab.Text = Format(GRID_SUMAR_COL(fg1, 4), FORMAT_MONTO)
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    If fg1.TextMatrix(Row, Col) = "" Then
        fg1.TextMatrix(Row, 2) = ""
        fg1.TextMatrix(Row, 5) = ""
        Exit Sub
    End If
    If Col = 1 Then
        If GRID_BUSCAR_VALOR(fg1, 1, Trim(fg1.TextMatrix(Row, Col)), False, -1, Row) <> "-1" Then
            MsgBox "El Num. Cuenta Contable ya existe" + vbCr + "Ingrese otro Num. Cuenta Contable", vbExclamation, xTitulo
            fg1.TextMatrix(Row, 1) = ""
            fg1.TextMatrix(Row, 5) = ""
            Exit Sub
        End If
        
    
        Dim Rst As New ADODB.Recordset
        RST_Busq Rst, "SELECT * FROM con_planctas WHERE cuenta = '" & NulosC(fg1.TextMatrix(Row, 1)) & "'", xCon
        If Rst.RecordCount = 1 Then
            fg1.TextMatrix(Row, 2) = NulosC(Rst("descripcion"))
            fg1.TextMatrix(Row, 5) = NulosN(Rst("id"))
        Else
            fg1.TextMatrix(Row, 2) = ""
            fg1.TextMatrix(Row, 5) = ""
        End If
        Set Rst = Nothing
    End If
    If Col = 3 Or Col = 4 Then
        If IsNumeric(fg1.TextMatrix(Row, Col)) = False Then
            MsgBox "El valor ingresado no es numérico", vbCritical, xTitulo
            fg1.TextMatrix(Row, Col) = ""
        Else
            If Col = 3 And NulosN(fg1.TextMatrix(Row, 4)) > 0 Then
                fg1.TextMatrix(Row, 4) = 0
            ElseIf Col = 4 And NulosN(fg1.TextMatrix(Row, 3)) > 0 Then
                fg1.TextMatrix(Row, 3) = 0
            End If
        End If

    End If
    HallarTotal
End Sub

Private Sub Fg1_EnterCell()
    If fg1.Col = 2 Then
        fg1.Editable = flexEDNone
    Else
        fg1.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Or Row < 1 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    Select Case Col
        Case 1
            
        Case 3, 4
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
    
    
End Sub

Private Sub fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        CmdAdd_Click
    End If
    
    If KeyCode = 46 Then
        CmdDel_Click
    End If
End Sub

Private Sub Fg1_LostFocus()
    HallarTotal
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
            PopupMenu menu1
        End If
    End If
End Sub


Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    SeEjecuto = False
    pCargarGrid
    SeEjecuto = True
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado ningún asiento, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            Nuevo
        End If
    End If
    
End Sub

Sub Nuevo()
    QueHace = 1
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Asiento"
    Blanquea
    Bloquea
    
    fg1.Rows = 1
    
    fg1.ColComboList(1) = "|..."
    fg1.SelectionMode = flexSelectionFree
    fg1.Editable = flexEDKbdMouse
    fg1.Rows = fg1.Rows + 1
    
    TxtIdLibro.Text = 3
    TxtIdLibro_Validate False
    
    xHorIni = Time
    txt_cb(1).SetFocus
End Sub

Sub Modificar()
    ActivaTool
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Asiento"
    Bloquea
    
    fg1.ColComboList(1) = "|..."
    fg1.SelectionMode = flexSelectionFree
    fg1.Editable = flexEDKbdMouse
    
    QueHace = 2
    
    xHorIni = Time
    txt_cb(1).SetFocus
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    
    If RstFrm.RecordCount = 0 Or RstFrm.EOF = True Or RstFrm.BOF = True Then
        MsgBox "No hay registro para eliminar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de eliminar el asiento seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        On Error GoTo LaCague
        xCon.BeginTrans
        'ELIMINAMOS EL ASIENTO REGISTRADO EN EL DIARIO
        xCon.Execute "DELETE * FROM con_diario WHERE ((idmes = " & xMes & ") and (idlib = 3) AND (idmov = " & RstFrm("id") & ")) ;"
        xCon.Execute "DELETE * FROM con_provicionesdet WHERE id = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM con_proviciones WHERE id = " & RstFrm("id") & ""
        xCon.CommitTrans
        RstFrm.Requery
        Dg1.Refresh
        MsgBox "El asiento fue eliminado con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TabOne1.CurrTab = 0
        If RstFrm.RecordCount = 0 Then
            Rpta = MsgBox("No se ha registrado ningún asiento, ¿Desea agregar una ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
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

Private Sub Form_Load()
    Agregando = False
    SeEjecuto = False
    QueHace = 3
    fg1.ColWidth(5) = 0
    TabOne1.CurrTab = 0
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    CaracteresNumericos = "0123456789." & Chr(8)
    
    fg1.ColFormat(3) = FORMAT_MONTO
    fg1.ColFormat(4) = FORMAT_MONTO
    
    Dg1.Columns("totdeb").NumberFormat = FORMAT_MONTO:
    Dg1.Columns("tothab").NumberFormat = FORMAT_MONTO:
    Dg1.Columns("fchdoc").NumberFormat = FORMAT_DATE:

    fg1.SelectionMode = flexSelectionByRow
'    Fg1.Editable = flexEDKbd
End Sub

Private Sub Menu1_1_Click()
    CmdAdd_Click
End Sub

Private Sub Menu1_3_Click()
    CmdDel_Click
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            MuestraSegundoTab
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstFrm.Requery
            Dg1.Refresh
        End If
    End If
    If Button.Index = 6 Then Cancelar
    If Button.Index = 8 Then Filtrar
    If Button.Index = 9 Then
        If xCon.State = 0 Then Exit Sub
        If RstFrm.State = 0 Then Exit Sub
        RstFrm.Filter = ""
    End If
    If Button.Index = 10 Then CambiarMes
    If Button.Index = 11 Then Buscar
    If Button.Index = 15 Then
        If TabOne1.CurrTab = 0 Then
            MsgBox "Para exportar el registro, primero muestre el detalle", vbExclamation, xTitulo
            Exit Sub
        End If
        pExportarMSExcel
    End If
    If Button.Index = 17 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Sub Cancelar()
    QueHace = 3
    fg1.SelectionMode = flexSelectionByRow
    Bloquea
    ActivaTool
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
End Sub

Function Grabar() As Boolean
    
    If fValidarDatos() = False Then Exit Function
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modficar") + " la Provición", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbNo Then Exit Function
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim xNumAsiento As String
    Dim xId, A As Integer
    
On Error GoTo LaCague

    xCon.BeginTrans
    If QueHace = 1 Then
        xNumAsiento = NuevoNumAsiento(3, xMes, xCon)
        xId = HallaCodigoTabla("con_proviciones", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM con_proviciones", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM con_proviciones WHERE id = " & xId & "", xCon
        'ELIMINAMOS EL DETALLE DE LA PROVICION
        xCon.Execute "DELETE * FROM con_provicionesdet WHERE id = " & xId & ""
        
        xNumAsiento = DevuelveNumAsiento(3, RstFrm("id"), xMes, xCon)
        If xNumAsiento = "" Then xNumAsiento = NuevoNumAsiento(3, xMes, xCon)
        'ELIMINAMOS EL ASIENTO REGISTRADO EN EL DIARIO
        xCon.Execute "DELETE * FROM con_diario WHERE ((idmes = " & xMes & ") and (idlib = 3) AND (idmov = " & xId & ")) ;"
    
    End If
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM con_provicionesdet", xCon
    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    
    RstCab("ano") = AnoTra
    RstCab("idmes") = xMes
    RstCab("numreg") = Format(xMes, "00") + xNumAsiento
    If xMes <> 0 And xMes <> 13 Then
        RstCab("fchreg") = CDate("01/" + Format(xMes, "00") + "/" + AnoTra)
    End If
    RstCab("idlib") = 3 '--proviciones diversas (libro diario)
    RstCab("idsublib") = NulosN(lbl_cb_cod(1).Caption)
    RstCab("idmon") = NulosN(lbl_cb_cod(0).Caption)
    RstCab("fchdoc") = CDate(TxtFchEmi.Valor)
    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("numser") = NulosC(TxtSerDoc.Text)
    RstCab("numdoc") = NulosC(TxtNumDoc.Text)
    RstCab("imp") = NulosN(TxtTotDeb.Text)
    RstCab("glosa") = NulosC(TxtGlosa.Text)
    
    RstCab.Update
    
    For A = 1 To fg1.Rows - 1
        RstDet.AddNew
        RstDet("id") = xId
        RstDet("idcuen") = NulosN(fg1.TextMatrix(A, 5))
        If NulosN(fg1.TextMatrix(A, 3)) <> 0 Then
            RstDet("tipo") = 0 '--debe
            RstDet("imp") = NulosN(fg1.TextMatrix(A, 3))
        End If
        If NulosN(fg1.TextMatrix(A, 4)) <> 0 Then
            RstDet("tipo") = -1 '--haber
            RstDet("imp") = NulosN(fg1.TextMatrix(A, 4))
        End If
        
        RstDet.Update
    Next A
    
    'grabamos el diario
    For A = 1 To fg1.Rows - 1
        RstDia.AddNew
        RstDia("año") = AnoTra
        RstDia("idmes") = xMes
        RstDia("idlib") = 3 'NulosN(TxtIdLibro.Text)
        RstDia("idmov") = xId
        RstDia("idcue") = NulosN(fg1.TextMatrix(A, 5))
        RstDia("numasi") = xNumAsiento
        RstDia("tc") = NulosN(LblTipoCambio.Caption)
        
        If NulosN(lbl_cb_cod(0).Caption) = 1 Then   '--soles
            RstDia("impdebsol") = NulosN(fg1.TextMatrix(A, 3))
            RstDia("imphabsol") = NulosN(fg1.TextMatrix(A, 4))
            
            If xMes = 0 Then
                RstDia("impdebdol") = (NulosN(fg1.TextMatrix(A, 3)) / NulosN(LblTipoCambio.Caption))
                RstDia("imphabdol") = (NulosN(fg1.TextMatrix(A, 4)) / NulosN(LblTipoCambio.Caption))
            Else
                RstDia("impdebdol") = 0
                RstDia("imphabdol") = 0
            End If
       
       Else '--dolares
            RstDia("impdebdol") = NulosN(fg1.TextMatrix(A, 3))
            RstDia("imphabdol") = NulosN(fg1.TextMatrix(A, 4))
            
            RstDia("impdebsol") = NulosN(fg1.TextMatrix(A, 3)) * NulosN(LblTipoCambio.Caption)
            RstDia("imphabsol") = NulosN(fg1.TextMatrix(A, 4)) * NulosN(LblTipoCambio.Caption)
        End If
        
        If xMes = 13 Then
            RstDia("fchasi") = CDate("31/12/" + AnoTra)
        Else
            RstDia("fchasi") = CDate("01/" + Format(xMes, "00") + "/" + AnoTra)
        End If
        RstDia("fchdoc") = CDate(TxtFchEmi.Valor)
        RstDia("prodiv") = -1
        RstDia.Update
    Next A
   
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, 14, QueHace, xHorIni, Time, Date, xCon, CDbl(xId)
   
    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDia = Nothing
    MsgBox "La Provición se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito" + vbCr + "Num.Reg. " + Format(xMes, "00") + xNumAsiento, vbInformation, xTitulo

    Grabar = True
    Exit Function
    
LaCague:
'    Resume
    Grabar = False
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDia = Nothing
    MsgBox "No se pudo guardar la provicion por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function



Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If txt_cb(Index).Text = "" Then Exit Sub
    If QueHace = 3 Then Exit Sub
    Dim RST_TMP As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--MONEDA
            nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion AS nombre, mae_moneda.id AS cod FROM mae_moneda WHERE (((mae_moneda.id)=" & NulosN(Trim(txt_cb(Index).Text)) & "))"
        Case 1 '--SUB LIBRO
            If NulosN(TxtIdLibro.Text) = 0 Then
                MsgBox "Falta especificar el libro", vbExclamation, xTitulo
                Exit Sub
            End If
            nSQL = "SELECT mae_librossub.id, mae_librossub.descripcion AS nombre, mae_librossub.id AS cod " _
                + vbCr + " FROM mae_librossub " _
                + vbCr + " WHERE (((mae_librossub.idlib) = " & NulosN(TxtIdLibro.Text) & ")) " _
                + vbCr + " AND (((mae_librossub.id)=" & NulosN(Trim(txt_cb(Index).Text)) & ")) ;"
                
    End Select
    If xCon.State = 0 Then Exit Sub
    RST_Busq RST_TMP, nSQL, xCon
    
    If RST_TMP.State = 0 Then GoTo SALIR
    If RST_TMP.RecordCount > 0 Then
        txt_cb(Index).Text = RST_TMP.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RST_TMP.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(Index).Caption = RST_TMP.Fields(2) & "" '--CODIGO
        
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cb_cod(Index).Caption = ""
        
    End If
    
SALIR:
    Set RST_TMP = Nothing
    Exit Sub
error:
    Set RST_TMP = Nothing
    SHOW_ERROR Me.Name, "txt_cb_KeyDown (" + CStr(Index) + ")"
End Sub

Private Sub TxtFchEmi_Validate(Cancel As Boolean)
    If IsDate(TxtFchEmi.Valor) = True Then
        If xMes = 0 Then
            'obtenemos el tipo de cambio del 31/12/xxxx para hacer la conversion a dolares, esto se hace solo para asientos de apertura
            Dim xMes1  As String
            xMes1 = "01/01/" + AnoTra
            LblTipoCambio.Caption = HallaTipoCambio(CDate(xMes1) - 1, 2, Venta, xCon)
            If NulosN(LblTipoCambio.Caption) = 0 Then
                MsgBox "No se ha especificado el tipo de cambio para el dia " & CDate(xMes1) - 1 & vbCr & "Ingrese el tipo de cambio para seguir con esta operación" + vbCr + "Ir a Menu Contabilidad/Tipo de Cambio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                TxtFchEmi.Valor = ""
                LblTipoCambio.Caption = ""
                TxtFchEmi.SetFocus
                Exit Sub
            End If
        Else
            LblTipoCambio.Caption = HallaTipoCambio(TxtFchEmi.Valor, 2, Venta, xCon)
        End If
    Else
        LblTipoCambio.Caption = ""
    End If
End Sub

Private Sub TxtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If fg1.Rows > 1 Then
             fg1.Row = 1
             fg1.Col = 1
             fg1.SetFocus
        Else
            CmdAdd.SetFocus
        End If
    End If
End Sub

Private Sub TxtIdLibro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdLibro_Validate(Cancel As Boolean)
    If NulosC(TxtIdLibro.Text) = "" Then Exit Sub
    LblLibro.Caption = Busca_Codigo(NulosN(TxtIdLibro.Text), "id", "descripcion", "mae_libros", "N", xCon)
    If LblLibro.Caption = "" Then
        TxtIdLibro.Text = ""
    End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
    If TxtNumDoc.Text = "" Then Exit Sub
    TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
End Sub

Private Sub TxtSerDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtSerDoc_Validate(Cancel As Boolean)
    If TxtSerDoc.Text = "" Then Exit Sub
    TxtSerDoc.Text = Format(TxtSerDoc.Text, "0000")
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDoc_Validate(Cancel As Boolean)
    If NulosC(TxtTipDoc.Text) = "" Then Exit Sub
    
    LblDocumento.Caption = Busca_Codigo(Val(TxtTipDoc.Text), "id", "descripcion", "mae_documento", "N", xCon)
    If LblDocumento.Caption = "" Then
        TxtTipDoc.Text = ""
        LblDocumento.Caption = ""
    End If
End Sub

Sub MuestraSegundoTab()
    On Error GoTo error
    fg1.Rows = 1
    Blanquea
    If RstFrm.RecordCount = 0 Then Exit Sub
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    QueHace = -1
    
    TxtIdLibro.Text = NulosN(RstFrm("idlib"))
    LblLibro.Caption = NulosC(RstFrm("desclib"))
    'LblDocumento.Caption = NulosC(RstFrm("descdoc"))
    TxtSerDoc.Text = NulosC(RstFrm("numser"))
    TxtNumDoc.Text = NulosC(RstFrm("numdoc"))
    TxtFchEmi.Valor = NulosC(RstFrm("fchdoc"))
    TxtFchEmi_Validate False
    
    TxtGlosa.Text = NulosC(RstFrm("glosa"))
    
    
    If NulosN(RstFrm.Fields("idsublib")) <> 0 Then
        txt_cb(1).Text = RstFrm.Fields("idsublib")
        txt_cb_Validate 1, False
    End If
    
    If NulosN(RstFrm.Fields("idmon")) <> 0 Then
        txt_cb(0).Text = RstFrm.Fields("idmon")
        txt_cb_Validate 0, False
    End If
    If NulosN(RstFrm("tipdoc")) <> 0 Then
        TxtTipDoc.Text = NulosN(RstFrm("tipdoc"))
        TxtTipDoc_Validate False
    End If
    
    QueHace = 3
   
    RST_Busq RstDet, "SELECT con_provicionesdet.id, con_provicionesdet.idcuen, con_planctas.cuenta, con_planctas.descripcion, " _
        & " IIf([con_provicionesdet].[tipo]=0,[con_provicionesdet]![imp],0) AS debe, IIf([con_provicionesdet].[tipo]=-1, " _
        & " [con_provicionesdet]![imp],0) AS haber FROM con_provicionesdet LEFT JOIN con_planctas ON con_provicionesdet.idcuen = con_planctas.id " _
        & " Where (((con_provicionesdet.id) = " & RstFrm("id") & ")) ORDER BY con_planctas.cuenta", xCon

    Agregando = True
    
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            fg1.Rows = fg1.Rows + 1
            fg1.TextMatrix(A, 1) = RstDet("cuenta") & ""
            fg1.TextMatrix(A, 2) = RstDet("descripcion") & ""
            fg1.TextMatrix(A, 3) = RstDet("debe") & ""
            fg1.TextMatrix(A, 4) = NulosN(RstDet("haber"))
            fg1.TextMatrix(A, 5) = NulosN(RstDet("idcuen"))
            
            RstDet.MoveNext
            If RstDet.EOF = True Then Exit For
        Next A
    End If
    
    HallarTotal
    
    Agregando = False
    Exit Sub
error:
'    Resume
    Agregando = False
    SHOW_ERROR Me.Name, "MuestraSegundoTab"
End Sub



Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL  As String
    
    LblPeriodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
    LblPeriodo(1).Caption = LblPeriodo(0).Caption

    nSQL = "SELECT con_proviciones.*, mae_libros.descripcion AS desclib, mae_documento.abrev AS destipdoc, con_meses.descripcion AS descmes, " _
        + vbCr + " (SELECT Sum([con_provicionesdet]![imp]) AS totdeb  FROM con_provicionesdet WHERE (((con_provicionesdet.id)=con_proviciones.id)  AND ((con_provicionesdet.tipo)=0))) AS totdeb, " _
        + vbCr + " (SELECT Sum([con_provicionesdet]![imp]) AS totdeb FROM con_provicionesdet WHERE  (((con_provicionesdet.id)=con_proviciones.id)   AND ((con_provicionesdet.tipo)=-1))) AS tothab, " _
        + vbCr + " mae_moneda.descripcion AS mondesc, mae_moneda.simbolo, [con_proviciones]![numser]+'-'+[con_proviciones]![numdoc] AS numedoc, mae_librossub.descripcion AS sublibdesc, " _
        + vbCr + " Format([con_proviciones].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([con_proviciones].[numreg],3) AS registro " _
        + vbCr + " FROM ((((con_proviciones LEFT JOIN mae_libros ON con_proviciones.idlib = mae_libros.id) LEFT JOIN con_meses ON con_proviciones.idmes = con_meses.id) LEFT JOIN mae_moneda ON con_proviciones.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_proviciones.tipdoc = mae_documento.id) LEFT JOIN mae_librossub ON con_proviciones.idsublib = mae_librossub.id " _
        + vbCr + " Where (((con_proviciones.ano) = " & AnoTra & "  ) And ((con_proviciones.idmes) = " & xMes & "  )) " _
        + vbCr + " ORDER BY con_proviciones.fchreg;"
    
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    '36
    Set Dg1.DataSource = RstFrm
    Me.MousePointer = vbDefault
Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

Private Sub CambiarMes()
    xMes = SeleccionaMes(xCon)
    pCargarGrid
    TabOne1.CurrTab = 0
End Sub


Private Function fValidarDatos() As Boolean
    '--VALIDAR QUE LA GRILLA DE ACTIVO Y PASIVO TENGAN VALORES TANTO DE ORDEN Y DESCRIPCION
    
    If NulosN(TxtIdLibro.Text) = 0 Then
        MsgBox "No ha especificado el libro contable", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If
    If NulosN(lbl_cb_cod(1).Caption) = 0 Then
        MsgBox "No ha especificado el Sub Libro para el libro " & LblLibro.Caption, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txt_cb(1).SetFocus
        Exit Function
    End If
    If NulosN(lbl_cb_cod(0).Caption) = 0 Then
        MsgBox "No ha especificado la Moneda", vbInformation, xTitulo
        txt_cb(0).SetFocus
        Exit Function
    End If
    
    If IsDate(TxtFchEmi.Valor) = False Then
        MsgBox "No ha especificado la fecha de movimiento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchEmi.SetFocus
        Exit Function
    End If
    
    If Trim(TxtGlosa.Text) = "" Then
        MsgBox "No ha especificado la glosa del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtGlosa.SetFocus
        Exit Function
    End If
    
    If fg1.Rows <= 1 Then
        MsgBox "Ingrese las Cuentas Contables", vbExclamation, xTitulo
        fg1.SetFocus
        Exit Function
    End If
    '--------------------------------
    HallarTotal
    
    If NulosN(TxtTotDeb.Text) <> NulosN(TxtTotHab.Text) Then
        MsgBox "Los totales del Debe y del Haber son diferentes" + vbCr + "Estos tienen que ser iguales", vbExclamation, xTitulo
        Exit Function
    End If
    '--------------------------------
    '--VALIDAR QUE EXISTA VALOR EN DEBE O HABER DE UAN FILA
    '--VALIDAR EL INGRESO DE LOS DATOS
    Dim mRow&
    Dim mCol& '--COLUMNA A POSICIONAR SI FALTAN DATOS
    mCol = -1
    For mRow = 1 To fg1.Rows - 1
        If fg1.TextMatrix(mRow, 1) = "" Then
            MsgBox "Ingrese La Cuenta Contable", vbExclamation, xTitulo
            mCol = 1:          Exit For
        ElseIf NulosN(fg1.TextMatrix(mRow, 3)) = 0 And NulosN(fg1.TextMatrix(mRow, 4)) = 0 Then
            MsgBox "Ingrese un valor en el Debe o Haber" + vbCr + "Luego Proceda", vbExclamation, xTitulo
            mCol = 3:          Exit For
        End If
    Next mRow
    If mCol <> -1 Then
        Agregando = True:  fg1.Row = mRow: fg1.Col = mCol: Agregando = False
        Exit Function
    End If
    '--solo para apertura
    If xMes = 0 And NulosN(lbl_cb_cod(0).Caption) = 1 Then '--es apertura y es soles
        If NulosN(LblTipoCambio.Caption) = 0 Then
            MsgBox "Falta ingresar el tipo de Cambio para el dia " & TxtFchEmi.Valor + vbCr + "Ir a Contabilidad/ Tipo de Cambio", vbExclamation, xTitulo
            Exit Function
        End If
    End If
    '-----
    fValidarDatos = True
End Function
 
Sub Filtrar()
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(6, 4) As String

    xCampos(0, 0) = "Num.Reg.":       xCampos(0, 1) = "registro":      xCampos(0, 2) = "C":          xCampos(0, 3) = "800"
    xCampos(1, 0) = "Libro":           xCampos(1, 1) = "desclib":     xCampos(1, 2) = "C":          xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Descripción":     xCampos(2, 1) = "glosa":       xCampos(2, 2) = "C":          xCampos(2, 3) = "6500"
    xCampos(3, 0) = "M":               xCampos(3, 1) = "mondesc":     xCampos(3, 2) = "C":          xCampos(3, 3) = "450"
    xCampos(4, 0) = "Debe":            xCampos(4, 1) = "totdeb":      xCampos(4, 2) = "N":          xCampos(4, 3) = "1000"
    xCampos(5, 0) = "Haber":           xCampos(5, 1) = "tothab":      xCampos(5, 2) = "N":          xCampos(5, 3) = "1000"

    TabOne1.CurrTab = 0
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1

End Sub

Sub Buscar()
    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim xCampos2(6, 4) As String
    
    xCampos2(0, 0) = "Num.Reg.":        xCampos2(0, 1) = "registro":      xCampos2(0, 2) = "900":          xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Sub Libro":       xCampos2(1, 1) = "sublibdesc":     xCampos2(1, 2) = "1900":         xCampos2(1, 3) = "C"
    xCampos2(2, 0) = "Descripción":     xCampos2(2, 1) = "glosa":       xCampos2(2, 2) = "3000":         xCampos2(2, 3) = "C"
    xCampos2(3, 0) = "M":               xCampos2(3, 1) = "simbolo":     xCampos2(3, 2) = "450":          xCampos2(3, 3) = "C"
    xCampos2(4, 0) = "Debe":            xCampos2(4, 1) = "totdeb":      xCampos2(4, 2) = "1000":         xCampos2(4, 3) = "N"
    xCampos2(5, 0) = "Haber":           xCampos2(5, 1) = "tothab":      xCampos2(5, 2) = "1000":         xCampos2(5, 3) = "N"
    
    nSQL = "SELECT con_proviciones.*, mae_libros.descripcion AS desclib, mae_documento.abrev AS destipdoc, con_meses.descripcion AS descmes, " _
        + vbCr + " (SELECT Sum([con_provicionesdet]![imp]) AS totdeb  FROM con_provicionesdet WHERE (((con_provicionesdet.id)=con_proviciones.id)  AND ((con_provicionesdet.tipo)=0))) AS totdeb, " _
        + vbCr + " (SELECT Sum([con_provicionesdet]![imp]) AS totdeb FROM con_provicionesdet WHERE  (((con_provicionesdet.id)=con_proviciones.id)   AND ((con_provicionesdet.tipo)=-1))) AS tothab, " _
        + vbCr + " mae_moneda.descripcion AS mondesc, mae_moneda.simbolo, [con_proviciones]![numser]+'-'+[con_proviciones]![numdoc] AS numedoc, mae_librossub.descripcion AS sublibdesc, " _
        + vbCr + " Format([con_proviciones].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([con_proviciones].[numreg],3) AS registro " _
        + vbCr + " FROM ((((con_proviciones LEFT JOIN mae_libros ON con_proviciones.idlib = mae_libros.id) LEFT JOIN con_meses ON con_proviciones.idmes = con_meses.id) LEFT JOIN mae_moneda ON con_proviciones.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_proviciones.tipdoc = mae_documento.id) LEFT JOIN mae_librossub ON con_proviciones.idsublib = mae_librossub.id " _
        + vbCr + " Where (((con_proviciones.ano) = " & AnoTra & "  ) And ((con_proviciones.idmes) = " & xMes & "  )) " _
        + vbCr + " ORDER BY con_proviciones.fchreg;"
    
    TabOne1.CurrTab = 0
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos2(), "Buscando Asientos Diversos", "glosa", "glosa", Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " & xRs("id") & ""
SALIR:
    Set xRs = Nothing
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub


Private Sub cb_Click(Index As Integer)
    On Error GoTo error
    Dim nSQL As String
    Dim xCampos() As String
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    
    If QueHace = 3 Then Exit Sub
    Select Case Index
    Case 0 '--MONEDA
    
        ReDim xCampos(2, 3) As String
        xCampos(0, 0) = "Moneda":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
        xCampos(1, 0) = "Símbolo":   xCampos(1, 1) = "simbolo":    xCampos(1, 2) = "1000":   xCampos(1, 3) = "C"
        nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion AS nombre, mae_moneda.id AS cod, mae_moneda.simbolo FROM mae_moneda ORDER BY mae_moneda.descripcion;"
        nTitulo = "Buscar Moneda"
    Case 1 '--SUBLIBRO
        If NulosN(TxtIdLibro.Text) = 0 Then
            MsgBox "Falta especificar el libro", vbExclamation, xTitulo
            
            Exit Sub
        End If
        
        ReDim xCampos(2, 3) As String
        xCampos(0, 0) = "Descripción":   xCampos(0, 1) = "nombre":  xCampos(0, 2) = "5000":     xCampos(0, 3) = "C"
        xCampos(1, 0) = "Id":            xCampos(1, 1) = "id":      xCampos(1, 2) = "500":      xCampos(1, 3) = "N"
        
        nSQL = "SELECT mae_librossub.id, mae_librossub.descripcion AS nombre, mae_librossub.id AS cod " _
            + vbCr + " From mae_librossub " _
            + vbCr + " Where (((mae_librossub.idlib) = " & NulosN(TxtIdLibro.Text) & ")) " _
            + vbCr + " ORDER BY mae_librossub.descripcion;"

        nTitulo = "Buscar Sub Libro"
    End Select

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
    If xRs.State = 0 Then GoTo SALIR
    If xRs.BOF = True Or xRs.EOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    If xRs.State = 0 Then GoTo SALIR
    
    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    Select Case Index
        Case 0 'MONEDA
             TxtFchEmi.SetFocus
        Case 1 'SUB LIBRO
            txt_cb(0).SetFocus
    End Select
   
   
SALIR:
    Set xRs = Nothing
Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click (" + CStr(Index) + ")"
End Sub



Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cb_cod(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub pExportarMSExcel()
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
        .Cells(1, 5) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        
        .Columns(2).ColumnWidth = fg1.ColWidth(1) / 100
        .Columns(3).ColumnWidth = fg1.ColWidth(2) / 100
        .Columns(4).ColumnWidth = fg1.ColWidth(3) / 100
        .Columns(5).ColumnWidth = fg1.ColWidth(4) / 100
                        
        '-----encabezado
        xFilas = 4
        .Cells(xFilas, 2) = "Provisiones Diversas"
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Libro"
        .Cells(xFilas, 3) = LblLibro.Caption
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Sub Libro"
        .Cells(xFilas, 3) = lbl_cb(1).Caption
        .Cells(xFilas, 4) = "Periodo"
        .Cells(xFilas, 5) = LblPeriodo(1).Caption
        
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Moneda"
        .Cells(xFilas, 3) = lbl_cb(0).Caption
        
        .Cells(xFilas, 4) = "T.C."
        .Cells(xFilas, 5) = NulosN(LblTipoCambio.Caption)
        
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Fch.Doc"
        .Cells(xFilas, 3) = "'" & TxtFchEmi.Valor
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Tipo Documento"
        .Cells(xFilas, 3) = "'" & LblDocumento.Caption
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "N° Documento"
        .Cells(xFilas, 3) = "'" & TxtSerDoc.Text & "-" & TxtNumDoc.Text
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Glosa"
        .Cells(xFilas, 3) = "'" & Trim(TxtGlosa.Text)

        xFilas = xFilas + 2
        
        .Cells(xFilas, 2) = "N° Cuenta"
        .Cells(xFilas, 3) = "Descripción"
        .Cells(xFilas, 4) = "Debe"
        .Cells(xFilas, 5) = "Haber"
        
        xFilas = xFilas + 1
        For A = 1 To fg1.Rows - 1
            .Cells(xFilas, 2) = "'" + fg1.TextMatrix(A, 1)
            .Cells(xFilas, 3) = "'" + fg1.TextMatrix(A, 2)
            .Cells(xFilas, 4) = NulosN(fg1.TextMatrix(A, 3))
            .Cells(xFilas, 5) = NulosN(fg1.TextMatrix(A, 4))
            xFilas = xFilas + 1
        Next A

        .Cells(xFilas, 3) = "Total"
        .Cells(xFilas, 4) = NulosN(TxtTotDeb.Text)
        .Cells(xFilas, 5) = NulosN(TxtTotHab.Text)
        
    End With
    
    MsgBox "El Registro se exportó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 1
    objExcel.Visible = True
    Set objExcel = Nothing
    Exit Sub
error:
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "ExportarExcelDetalle", , IIf(Err.Number <> 50290, "", "No manipule el archivo hasta que termine de exportar!!!!")
End Sub

