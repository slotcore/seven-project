VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#12.0#0"; "Codejock.Calendar.v12.0.0.ocx"
Begin VB.Form FrmCronoProduccion2 
   Caption         =   "Produccion - Cronograma de Produccion"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7125
      Left            =   15
      TabIndex        =   1
      Top             =   360
      Width           =   11850
      _cx             =   20902
      _cy             =   12568
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
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
      Caption         =   "  &Consulta  |   &Detalle   "
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
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   6690
         Left            =   45
         TabIndex        =   3
         Top             =   390
         Width           =   11760
         Begin VB.CommandButton CmdProcesar 
            Caption         =   "&Procesar"
            Height          =   570
            Left            =   9270
            TabIndex        =   49
            Top             =   780
            Visible         =   0   'False
            Width           =   2280
         End
         Begin VB.Frame FrmAdd 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   2450
            Left            =   3150
            TabIndex        =   30
            Top             =   2430
            Visible         =   0   'False
            Width           =   7800
            Begin MSComCtl2.DTPicker DTPHoras 
               Height          =   345
               Left            =   1800
               TabIndex        =   50
               Top             =   1170
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   609
               _Version        =   393216
               CustomFormat    =   "HH:mm"
               Format          =   58392579
               UpDown          =   -1  'True
               CurrentDate     =   40606
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   90
               TabIndex        =   38
               Top             =   1700
               Width           =   7575
               Begin VB.CommandButton CmdAgregar 
                  Caption         =   "&Aceptar"
                  Height          =   350
                  Left            =   2370
                  TabIndex        =   40
                  Top             =   180
                  Width           =   1155
               End
               Begin VB.CommandButton CmdAnular 
                  Caption         =   "&Cancelar"
                  Height          =   350
                  Left            =   3570
                  TabIndex        =   39
                  Top             =   180
                  Width           =   1155
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00FFFFFF&
               Height          =   1305
               Left            =   60
               TabIndex        =   34
               Top             =   300
               Width           =   1665
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Mat Prima/Producto"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   37
                  Top             =   180
                  Width           =   1425
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Cantidad"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   36
                  Top             =   540
                  Width           =   630
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "Hora de Inicio"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   35
                  Top             =   930
                  Width           =   990
               End
            End
            Begin VB.CommandButton CmdAddMatProd 
               Height          =   240
               Left            =   2550
               Picture         =   "FrmCronoProduccion2.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   450
               Width           =   225
            End
            Begin VB.TextBox TxtCant 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1800
               TabIndex        =   32
               Text            =   "TxtCant"
               Top             =   800
               Width           =   975
            End
            Begin VB.TextBox TxtMatProd 
               Height          =   300
               Left            =   1800
               TabIndex        =   31
               Text            =   "TxtMatProd"
               Top             =   420
               Width           =   1000
            End
            Begin VB.Label LblDia 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "LblDia"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   7140
               TabIndex        =   51
               Top             =   60
               Width           =   555
            End
            Begin VB.Label LblUnidad 
               BackColor       =   &H80000013&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblUnidad"
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
               Left            =   2850
               TabIndex        =   44
               Top             =   800
               Width           =   4840
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "HH/mm"
               Height          =   195
               Left            =   2850
               TabIndex        =   43
               Top             =   1260
               Width           =   555
            End
            Begin VB.Label LblMatProd 
               BackColor       =   &H80000013&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblMatProd"
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
               Left            =   2850
               TabIndex        =   42
               Top             =   420
               Width           =   4845
            End
            Begin VB.Line Line8 
               BorderColor     =   &H80000003&
               BorderWidth     =   2
               X1              =   0
               X2              =   7770
               Y1              =   2415
               Y2              =   2430
            End
            Begin VB.Line Line7 
               BorderColor     =   &H80000003&
               BorderWidth     =   2
               X1              =   7755
               X2              =   7770
               Y1              =   15
               Y2              =   2430
            End
            Begin VB.Line Line6 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   2
               X1              =   15
               X2              =   0
               Y1              =   0
               Y2              =   3315
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   2
               X1              =   15
               X2              =   6045
               Y1              =   15
               Y2              =   15
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Agregando Cronograma"
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
               Height          =   195
               Left            =   120
               TabIndex        =   41
               Top             =   60
               Width           =   1995
            End
            Begin VB.Shape Shape2 
               BackColor       =   &H80000002&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00800000&
               Height          =   255
               Left            =   25
               Top             =   45
               Width           =   7725
            End
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   3345
            Left            =   240
            TabIndex        =   19
            Top             =   2250
            Visible         =   0   'False
            Width           =   6105
            Begin VB.TextBox TxtMP 
               Height          =   300
               Left            =   1170
               TabIndex        =   24
               Text            =   "TxtMP"
               Top             =   360
               Width           =   4845
            End
            Begin VB.TextBox TxtCan 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1170
               TabIndex        =   23
               Text            =   "TxtCan"
               Top             =   660
               Width           =   1185
            End
            Begin VB.CommandButton CmAcepta 
               Caption         =   "&Aceptar"
               Height          =   350
               Left            =   1860
               TabIndex        =   22
               Top             =   2865
               Width           =   1155
            End
            Begin VB.CommandButton CmdCancelar 
               Caption         =   "&Cancelar"
               Height          =   350
               Left            =   3045
               TabIndex        =   21
               Top             =   2865
               Width           =   1155
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4020
               Locked          =   -1  'True
               TabIndex        =   20
               Text            =   "TxtTotal"
               Top             =   2460
               Width           =   945
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   1470
               Left            =   75
               TabIndex        =   25
               Top             =   990
               Width           =   5880
               _cx             =   10372
               _cy             =   2593
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
               BackColorSel    =   -2147483645
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
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmCronoProduccion2.frx":0132
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
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Materia Prima"
               Height          =   195
               Left            =   75
               TabIndex        =   29
               Top             =   390
               Width           =   960
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Seleccion de Productos"
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
               Height          =   195
               Left            =   120
               TabIndex        =   28
               Top             =   60
               Width           =   2040
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   2
               X1              =   15
               X2              =   6045
               Y1              =   15
               Y2              =   15
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   2
               X1              =   15
               X2              =   0
               Y1              =   0
               Y2              =   3315
            End
            Begin VB.Line Line3 
               BorderColor     =   &H80000003&
               BorderWidth     =   2
               X1              =   6060
               X2              =   6060
               Y1              =   15
               Y2              =   3330
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Cantidad"
               Height          =   195
               Left            =   75
               TabIndex        =   27
               Top             =   690
               Width           =   630
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000003&
               BorderWidth     =   2
               X1              =   15
               X2              =   6045
               Y1              =   3315
               Y2              =   3315
            End
            Begin VB.Label Label11 
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
               Height          =   165
               Left            =   3090
               TabIndex        =   26
               Top             =   2505
               Width           =   825
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H80000002&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00800000&
               Height          =   255
               Left            =   30
               Top             =   45
               Width           =   6015
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1245
            Left            =   0
            TabIndex        =   4
            Top             =   245
            Width           =   9060
            Begin VB.ComboBox ComboSemanas 
               Height          =   315
               ItemData        =   "FrmCronoProduccion2.frx":01CD
               Left            =   1020
               List            =   "FrmCronoProduccion2.frx":01CF
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   450
               Width           =   1000
            End
            Begin VB.CommandButton CmdBusTip 
               Enabled         =   0   'False
               Height          =   240
               Left            =   1770
               Picture         =   "FrmCronoProduccion2.frx":01D1
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   810
               Width           =   225
            End
            Begin VB.CommandButton CmdBusSup 
               Enabled         =   0   'False
               Height          =   240
               Left            =   1770
               Picture         =   "FrmCronoProduccion2.frx":0303
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   180
               Width           =   225
            End
            Begin VB.TextBox TxtIdSup 
               Height          =   300
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   6
               Text            =   "TxtIdSup"
               Top             =   150
               Width           =   1000
            End
            Begin VB.TextBox TxtTipPro 
               Height          =   300
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   5
               Text            =   "TxtTipPro"
               Top             =   780
               Width           =   1000
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
               Height          =   300
               Left            =   5415
               TabIndex        =   10
               Top             =   450
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
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   3000
               TabIndex        =   11
               Top             =   450
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
            Begin VB.Label LabelSemana 
               AutoSize        =   -1  'True
               Caption         =   "Semana"
               Height          =   195
               Left            =   60
               TabIndex        =   18
               Top             =   510
               Width           =   585
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Prod."
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   17
               Top             =   840
               Width           =   735
            End
            Begin VB.Label LblTipoProd 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblTipoProd"
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
               Left            =   2055
               TabIndex        =   16
               Top             =   780
               Width           =   6795
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Supervisor"
               Height          =   195
               Left            =   60
               TabIndex        =   15
               Top             =   195
               Width           =   750
            End
            Begin VB.Label LblSupervisor 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblSupervisor"
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
               Left            =   2055
               TabIndex        =   14
               Top             =   150
               Width           =   6795
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Final"
               Height          =   195
               Left            =   4530
               TabIndex        =   13
               Top             =   510
               Width           =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Inicio"
               Height          =   195
               Left            =   2055
               TabIndex        =   12
               Top             =   510
               Width           =   735
            End
         End
         Begin XtremeCalendarControl.CalendarControl CalendarControl1 
            Height          =   5175
            Left            =   0
            TabIndex        =   45
            Top             =   1500
            Width           =   11715
            _Version        =   786432
            _ExtentX        =   20664
            _ExtentY        =   9128
            _StockProps     =   64
            ViewType        =   2
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H8000000B&
            BackStyle       =   1  'Opaque
            Height          =   5175
            Left            =   0
            Top             =   1500
            Width           =   11715
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Cronograma"
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
            Height          =   315
            Left            =   0
            TabIndex        =   46
            Top             =   -10
            Width           =   11655
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   6690
         Left            =   -12405
         TabIndex        =   2
         Top             =   390
         Width           =   11760
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6285
            Left            =   30
            TabIndex        =   47
            Top             =   360
            Width           =   11700
            _ExtentX        =   20638
            _ExtentY        =   11086
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Fch. Ini."
            Columns(1).DataField=   "fchini"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Fin."
            Columns(2).DataField=   "fchfin"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tipo Produccion"
            Columns(3).DataField=   "destippro"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Programador"
            Columns(4).DataField=   "apenom"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1535"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1455"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2223"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2143"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2249"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2170"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=3757"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3678"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=9102"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=9022"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Named:id=33:Normal"
            _StyleDefs(57)  =   ":id=33,.parent=0"
            _StyleDefs(58)  =   "Named:id=34:Heading"
            _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(60)  =   ":id=34,.wraptext=-1"
            _StyleDefs(61)  =   "Named:id=35:Footing"
            _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(63)  =   "Named:id=36:Selected"
            _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=37:Caption"
            _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(67)  =   "Named:id=38:HighlightRow"
            _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=39:EvenRow"
            _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(71)  =   "Named:id=40:OddRow"
            _StyleDefs(72)  =   ":id=40,.parent=33"
            _StyleDefs(73)  =   "Named:id=41:RecordSelector"
            _StyleDefs(74)  =   ":id=41,.parent=34"
            _StyleDefs(75)  =   "Named:id=42:FilterBar"
            _StyleDefs(76)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Cronogramas"
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
            Height          =   315
            Left            =   0
            TabIndex        =   48
            Top             =   -10
            Width           =   11700
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   30
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
            Picture         =   "FrmCronoProduccion2.frx":0435
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.frx":0979
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.frx":0AFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.frx":0F51
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.frx":1069
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.frx":15AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.frx":1AF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.frx":1C05
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.frx":1D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.frx":216D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.frx":22D9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1005
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
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
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
            ImageIndex      =   11
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Recetas del producto"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Productos "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu_01 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu_01_02 
         Caption         =   "Agregar Producto"
      End
      Begin VB.Menu menu_01_04 
         Caption         =   "Modificar Producto"
      End
      Begin VB.Menu menu_01_03 
         Caption         =   "Eliminar Producto"
      End
      Begin VB.Menu menu_01_01 
         Caption         =   "Seleccionar Productos"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu menu2_1 
         Caption         =   "Agregar Producto"
      End
      Begin VB.Menu menu2_3 
         Caption         =   "Modificar Producto"
      End
      Begin VB.Menu menu2_2 
         Caption         =   "Eliminar Producto"
      End
   End
   Begin VB.Menu menu3 
      Caption         =   "Menu3"
      Visible         =   0   'False
      Begin VB.Menu menu3_1 
         Caption         =   "Ver Productos"
      End
   End
End
Attribute VB_Name = "FrmCronoProduccion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xNomMatPriPro As String
Dim QueHace As Integer
Dim Agregando As Boolean
Dim RstLis As New ADODB.Recordset
Dim RstMatPro As New ADODB.Recordset
Dim xIdMatPri As Integer
Dim xFchPro, xHorPro As Date

Dim oPDF As cPDF
Dim xNumPag As Integer
Dim xFilaInicial As Integer
Dim xHorIni As Date 'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer 'INDICA EL CODIGO DEL MENU ACTIVO
Dim fOrdenLista As Boolean ' especfica el orden de la lista de la consulta
Dim SeEjecuto As Boolean

Dim visEvent As Boolean
Dim modifEvent As Boolean
Dim agregEvent As Boolean

Dim mIdRegistro& 'identificador del registro

Dim OrigFX As Long
Dim OrigFY As Long

Dim HitTest As CalendarHitTestInfo
Dim c_Event As CalendarEvent

Dim cambio As Boolean

Private Sub CalendarControl1_DblClick()
    If QueHace <> 3 Then Exit Sub
    visEvent = True
    If TxtTipPro.Text = 1 Then
        Set HitTest = CalendarControl1.ActiveView.HitTest
        On Error Resume Next
        Set c_Event = HitTest.ViewEvent.Event
        menu_01_01_Click
        menu_01_04_Click
    Else
        Set HitTest = CalendarControl1.ActiveView.HitTest
        On Error Resume Next
        Set c_Event = HitTest.ViewEvent.Event
        Menu2_3_Click
    End If
End Sub

Private Sub CalendarControl1_KeyDown(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    On Error Resume Next
    'Se activa el detector para la vista activa del calendario
    Set HitTest = CalendarControl1.ActiveView.HitTest
    'Se agrega el evento del detector
    Set c_Event = HitTest.ViewEvent.Event
    
    If KeyCode = vbKeyInsert Then
        Menu2_1_Click
    End If
    If KeyCode = vbKeyDelete Then
        'Si el detector no tiene evento activo
        If HitTest.ViewEvent Is Nothing Then Exit Sub
        menu2_2_Click
    End If
    If KeyCode = 113 Then
        'Si el detector no tiene evento activo
        If HitTest.ViewEvent Is Nothing Then Exit Sub
        Menu2_3_Click
    End If
End Sub

Private Sub CalendarControl1_ViewChanged()
    'Si la vista del calendario es por dia se activa la barra de desplazamiento vertical
    If CalendarControl1.ViewType = xtpCalendarDayView Then CalendarControl1.DayView.EnableVScroll True
    'Si la vista del calendario es por semana se desactiva la barra de desplazamiento vertical
    If CalendarControl1.ViewType = xtpCalendarWeekView Then CalendarControl1.WeekView.EnableVScroll False
End Sub

Private Sub CalendarControl1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
                If NulosN(TxtTipPro.Text) = 1 Then
                    On Error Resume Next
                    Set HitTest = CalendarControl1.ActiveView.HitTest
                    Set c_Event = HitTest.ViewEvent.Event
                    PopupMenu menu_01
                    Set HitTest = Nothing
                Else
                    On Error Resume Next
                    Set HitTest = CalendarControl1.ActiveView.HitTest
                    Set c_Event = HitTest.ViewEvent.Event
                    PopupMenu menu2
                    Set HitTest = Nothing
                End If
        Else
            If NulosN(TxtTipPro.Text) = 1 Then
                On Error Resume Next
                Set HitTest = CalendarControl1.ActiveView.HitTest
                If HitTest.ViewEvent Is Nothing Then Exit Sub
                Set c_Event = HitTest.ViewEvent.Event
                PopupMenu menu3
                Set HitTest = Nothing
            End If
        End If
    End If
End Sub

'*****************************************************************************************************
'* Descripcion      : EVITA LA EDICION DEL CALENDARIO AL HACER CLIC
'* Modificacion     : 15/02/11 JOSE CHACON
'*****************************************************************************************************
Private Sub CalendarControl1_BeforeEditOperation(ByVal OpParams As XtremeCalendarControl.CalendarEditOperationParameters, bCancelOperation As Boolean)
    bCancelOperation = True
End Sub

Private Sub iniciarCampos()
    Dim pTema2007 As CalendarThemeOffice2007
    Dim A As Integer
    
    'Se guarda el tema del calendario activo
    Set pTema2007 = CalendarControl1.Theme
    'Se cambia el color de seleccion
    pTema2007.WeekView.Day.BackgroundSelectedColor = RGB(215, 215, 215)
    'Se inabilita los mensajes de ayuda
    CalendarControl1.EnableToolTips False
            
    Frame2.BackColor = &H8000000F
    TabOne1.CurrTab = 0
    
    Fg2.ColWidth(4) = 0
    'se cargan las semanas
    For A = 1 To 52
        ComboSemanas.AddItem A
    Next A
    'se cargan las horas
'    For A = 0 To 23
'        CmbHoras.AddItem Format(A, "0#")
'    Next A
'    'se cargan los minutos
'    For A = 0 To 59
'        CmbMinutos.AddItem Format(A, "0#")
'    Next A
    
'    UpDown1.Max = 23
'    UpDown1.Min = 0
'    UpDown2.Max = 59
'    UpDown2.Min = 0
    
    Me.Height = 8000
    Me.Width = 12000
End Sub




Private Sub CmdAddMatProd_Click()
    Dim xRs As New ADODB.Recordset
    Dim cSQL As String
    Dim titulo As String
    Dim xCampos(3, 4) As String
    
    If QueHace = 3 Then Exit Sub

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Uni. Med.":     xCampos(2, 1) = "abrev":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"

    If TxtTipPro.Text = "1" Then
        cSQL = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev " _
            + vbCr + "FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
            + vbCr + "Where (((alm_inventario.tippro) = 1)) " _
            + vbCr + "ORDER BY alm_inventario.descripcion"

        titulo = "Buscando Materia Prima"
    End If

    If TxtTipPro.Text = "3" Then
        cSQL = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.activo" _
            + vbCr + "FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
            + vbCr + "Where (((alm_inventario.tippro) = 3) And ((alm_inventario.activo) = -1)) " _
            + vbCr + "ORDER BY alm_inventario.descripcion"
            
        titulo = "Buscando Productos"
    End If
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos, titulo, "descripcion", "descripcion"
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            If xRs.RecordCount <> 0 Then
                TxtMatProd.Text = xRs("id")
                LblMatProd.Caption = xRs("descripcion")
                LblUnidad.Caption = xRs("abrev")
                TxtCant.SetFocus
            End If
        End If
    End If
    Set xRs = Nothing
End Sub

Private Sub CmdAgregar_Click()
    Dim horIni As Date
    Dim fIni As Date
    Dim fFin As Date
    Dim AllDay As Boolean

    If QueHace <> 3 Then
        If modifEvent Then CalendarControl1.DataProvider.DeleteEvent c_Event
        
        Set c_Event = CalendarControl1.DataProvider.CreateEvent
    
        CalendarControl1.ActiveView.GetSelection fIni, fFin, AllDay
        
        
        horIni = fIni & " " & Format(DTPHoras.Value, "HH:mm")
        
        c_Event.ScheduleID = NulosN(TxtMatProd.Text)
        c_Event.StartTime = horIni
        c_Event.EndTime = fFin
        c_Event.Subject = LblMatProd.Caption
        c_Event.Location = TxtCant.Text & " " & LblUnidad.Caption
        c_Event.ReminderSoundFile = NulosN(TxtCant.Text)
        'se coloca dentro del body la hora en formato 24 horas para evitar errores
        c_Event.Body = Format(DTPHoras.Value, "HH:mm")
        
        CalendarControl1.DataProvider.AddEvent c_Event
    End If
    
    FrmAdd.Visible = False
    modifEvent = False
    agregEvent = False
End Sub

Function DateFromString(DatePart As String, TimePart As String) As Date
    Dim dtDatePart As Date, dtTimePart As Date
    dtDatePart = DatePart
    dtTimePart = TimePart
    DateFromString = dtDatePart + dtTimePart
End Function

Private Sub CmdAnular_Click()
    CmdAgregar.Enabled = True
    visEvent = False
    FrmAdd.Visible = False
End Sub

Private Sub ComboSemanas_Click()
    If QueHace <> 3 Then
        Dim fechaI As Date
        Dim fechaF As Date
        calcularSemana ComboSemanas.Text, fechaI, fechaF
        cambio = True
        TxtFchIni.Valor = fechaI
        TxtFchFin.Valor = fechaF
        cambio = False
        CmdBusTip.SetFocus
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : LlenarDatos
'* Tipo             : SUB
'* Descripcion      : CARGA LOS DATOS AL CALENDARIO
'* Modificacion     : 15/02/11 JOSE CHACON
'*****************************************************************************************************
Sub LlenarDatos()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Dim cSQL As String
    
    cSQL = "SELECT pro_cronogramadet.*, alm_inventario.descripcion, mae_unidades.abrev " _
        + vbCr + "FROM (pro_cronogramadet LEFT JOIN alm_inventario ON pro_cronogramadet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        + vbCr + "WHERE (((pro_cronogramadet.id)=" & RstLis("id") & "))"

    RST_Busq Rst, cSQL, xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Dim xHorIni, horFin As String
        xHorIni = "08:00:00"
        'se crea un evento nuevo de calendario
        Dim eventoNuevo As CalendarEvent
        Set eventoNuevo = CalendarControl1.DataProvider.CreateEvent
        'se procede a llenar los detalles del evento
        For A = 1 To Rst.RecordCount
            eventoNuevo.ScheduleID = NulosN(Rst("iditem"))
            eventoNuevo.Subject = NulosC(Rst("descripcion"))
            eventoNuevo.Location = NulosC(Rst("cantidad")) & " " & NulosC(Rst("abrev"))
            eventoNuevo.ReminderSoundFile = NulosC(Rst("cantidad"))
            
            eventoNuevo.StartTime = Format(Rst("fchpro"), "dd/mm/yyyy") & " " & NulosC(Format(Rst("horpro"), "HH:mm"))
            eventoNuevo.Body = NulosC(Format(Rst("horpro"), "HH:mm"))
            
            eventoNuevo.EndTime = Format(Rst("fchpro"), "dd/mm/yyyy") & " " & NulosC(Format(Rst("horpro"), "HH:mm"))
            eventoNuevo.Importance = xtpCalendarImportanceHigh
            
            'se agrega el evento nuevo al calendario
            CalendarControl1.DataProvider.AddEvent eventoNuevo
            
            Rst.MoveNext
        Next A
    End If
End Sub

Sub calcularSemana(numSemana As Integer, ByRef fechaInicio As Date, ByRef fechaFin As Date)
    Dim fechaRef As Date
    fechaRef = CDate("01/01/" & AnoTra)
    
    'Buscamos el primer Lunes del Año
    While Weekday(fechaRef) <> vbMonday
        'Vamos sumando dia a dia, hasta encontrar el primer lunes
        fechaRef = fechaRef + 1
    Wend
    
    'Multiplicamos y obtenemos el rango inferior de la semana
    fechaInicio = fechaRef + (7 * (numSemana - 1))
    'Obtenemos el rango superior de la semana
    fechaFin = fechaInicio + 6
End Sub

Private Sub CmAcepta_Click()
    Dim B As Integer
    
    If NulosN(TxtTotal.Text) > NulosN(TxtCan.Text) Then
        MsgBox "El cantidad a procesar en productos es mayor a la cantidad de materia prima", vbInformation + vbOKOnly + vbDefaultButton1
        TxtTotal.SetFocus
        Exit Sub
    End If
    
    For B = 1 To Fg2.Rows - 1
        RstMatPro.Filter = adFilterNone
        If Abs(NulosN(Fg2.TextMatrix(B, 3))) = 1 Then
            RstMatPro.Filter = "iditem = " & xIdMatPri & " AND fchpro = " & xFchPro & " AND horpro = " & Format(xHorPro, "HH:mm") & " AND idpro = " & NulosN(Fg2.TextMatrix(B, 4)) & ""
            If RstMatPro.RecordCount = 0 Then
                RstMatPro.AddNew
                RstMatPro("id") = 0
                RstMatPro("iditem") = xIdMatPri
                RstMatPro("fchpro") = xFchPro
                RstMatPro("horpro") = xHorPro
                RstMatPro("idpro") = Fg2.TextMatrix(B, 4)
                RstMatPro("cantidad") = NulosN(Fg2.TextMatrix(B, 2))
            Else
                RstMatPro("cantidad") = NulosN(Fg2.TextMatrix(B, 2))
            End If
        Else
            RstMatPro.Filter = "iditem = " & xIdMatPri & " AND fchpro = " & xFchPro & " AND horpro = " & Format(xHorPro, "HH:mm") & " AND idpro = " & NulosN(Fg2.TextMatrix(B, 4)) & ""
            If RstMatPro.RecordCount <> 0 Then
                RstMatPro.Delete
            End If
        End If
    Next B
    
    CmdCancelar_Click
End Sub

Private Sub CmdBusSup_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Apellidos y Nombres":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":                xCampos(1, 1) = "id":        xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT pro_emp.*, pla_empleados.nombre FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) " _
        & "  LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id Where (((pro_empdet.idfun) = 2)) ORDER BY pla_empleados.nombre"
            
    xform.titulo = "Buscando Supervisores"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdSup.Text = xRs("id")
            LblSupervisor.Caption = xRs("nombre")
            TxtFchIni.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdBusTip_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_tipoproducto.id, mae_tipoproducto.descripcion FROM mae_tipoproducto"
    
    xform.titulo = "Buscando Tipo de Item"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipPro.Text = xRs("id")
            LblTipoProd.Caption = xRs("descripcion")
            CmdProcesar.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdCancelar_Click()
    TabOne1.Enabled = True
    Toolbar1.Enabled = True
    
    CmAcepta.Enabled = True
    visEvent = False
    Frame3.Visible = False
End Sub

Private Sub CmdProcesar_Click()
    If TxtFchIni.Valor = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1
        TxtFchIni.SetFocus
        Exit Sub
    End If

    If TxtFchFin.Valor = "" Then
        MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1
        TxtFchFin.SetFocus
        Exit Sub
    End If

    If NulosN(TxtTipPro.Text) = 0 Then
        MsgBox "No ha especificado el tipo de producto a procesar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipPro.SetFocus
        Exit Sub
    End If
    
    If QueHace = 1 Then CalendarControl1.Visible = True: CmdProcesar.Visible = False
    CalendarControl1.ActiveView.ShowDay (CDate(TxtFchIni.Valor))
    CalendarControl1.ViewType = xtpCalendarWeekView
End Sub

Function Grabar() As Boolean
    Dim A As Integer
    Dim xTot As Long
    
    Dim RstSolMat As New ADODB.Recordset
    Dim xIdSol As Double
    Dim RstSolMatDet As New ADODB.Recordset
    Dim numDoc As Double
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDet1 As New ADODB.Recordset
    Dim xId As Double
    Dim nSQL As String
    
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If TxtIdSup.Text = "" Then
        MsgBox "No ha especificado un Supervisor para el nuevo Cronograma, especifique uno", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        CmdBusSup.SetFocus
        Exit Function
    End If
    
    If ComboSemanas.Text = "" Then
        MsgBox "No ha especificado una fecha para el Cronograma, especifique una", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        ComboSemanas.SetFocus
        Exit Function
    End If
    
    If TxtTipPro.Text = "" Then
        MsgBox "No ha especificado un tipo de Producto para el Cronograma, especifique uno", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        CmdBusTip.SetFocus
        Exit Function
    End If
    
    On Error GoTo LaCague
    xCon.BeginTrans
    
    If QueHace = 1 Then
        ' SI ES UN NUEVO REGISTRO OBTENEMOS EL ULTIMO ID DE LA TABLA ped_pedido
        xId = HallaCodigoTabla("pro_cronograma", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pro_cronograma", xCon
        RstCab.AddNew
        RstCab("id") = xId
        
'        xIdSol = HallaCodigoTabla("pro_ordenprod", xCon, "id")
'        RST_Busq RstSolMat, "SELECT TOP 1 * FROM pro_ordenprod", xCon
'        RstSolMat.AddNew
'        RstSolMat("id") = xIdSol
    Else
        ' SI SE ESTA MOFIGICANDO UN REGISTRO OBTENEMOS EL ID DEL REGISTRO ACTUAL
        xId = RstLis("id")
        RST_Busq RstCab, "SELECT * FROM pro_cronograma WHERE id = " & xId & "", xCon
        ' Eliminamos el detalle
        xCon.Execute "DELETE * FROM pro_cronogramadet WHERE id  = " & xId & ""
        xCon.Execute "DELETE * FROM pro_cronogramadetprod WHERE id  = " & xId & ""
    End If
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM pro_cronogramadet", xCon
    RST_Busq RstDet1, "SELECT TOP 1 * FROM pro_cronogramadetprod", xCon
    
    mIdRegistro = xId
    
    RstCab("idsup") = NulosC(TxtIdSup.Text)
    RstCab("fchini") = NulosC(TxtFchIni.Valor)
    RstCab("fchfin") = NulosC(TxtFchFin.Valor)
    RstCab("idtippro") = NulosN(TxtTipPro.Text)
    RstCab.Update
    
    On Error Resume Next

    Dim pEvent As CalendarEvent
    Dim Events As CalendarEvents

    Set Events = CalendarControl1.DataProvider.GetAllEventsRaw
    
    For Each pEvent In Events
        RstDet.AddNew
        RstDet("id") = xId
        RstDet("fchpro") = NulosC(Mid(pEvent.StartTime, 1, 10))
        'Se graba la hora en formato 24 horas
        RstDet("horpro") = Format(pEvent.Body, "HH:mm")
        RstDet("iditem") = NulosC(pEvent.ScheduleID)
        RstDet("cantidad") = NulosN(pEvent.ReminderSoundFile)
        RstDet.Update
    Next
    
    RstMatPro.Filter = adFilterNone
    If RstMatPro.RecordCount <> 0 Then
        RstMatPro.MoveFirst
        For A = 1 To RstMatPro.RecordCount
            RstDet1.AddNew
            RstDet1("id") = xId
            RstDet1("iditem") = NulosN(RstMatPro("iditem"))
            RstDet1("fchpro") = NulosC(RstMatPro("fchpro"))
            RstDet1("horpro") = NulosC(RstMatPro("horpro"))
            RstDet1("idpro") = NulosC(RstMatPro("idpro"))
            RstDet1("cantidad") = NulosC(RstMatPro("cantidad"))
            RstDet1.Update
            RstMatPro.MoveNext
        Next A
    End If
        
    ' Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
   
    xCon.CommitTrans
    MsgBox "La operacion se registró con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDet1 = Nothing
    Grabar = True
    Exit Function
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDet1 = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
End Function

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstLis
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDETE LAS COLUMNAS DEL CONTROL Dg3
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstLis.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstLis("id")), xCon
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Fg2.Col = 2 Then
        Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, Col), "0.00")
        
        TxtTotal.Text = Format(GRID_SUMAR_COL(Fg2, 2, 1, Fg2.Rows - 1), "0.00")
        If NulosN(Fg2.TextMatrix(Row, Col)) <> 0 Then
            Fg2.TextMatrix(Row, 3) = 1
        Else
            Fg2.TextMatrix(Row, 3) = 0
        End If
    End If
    If Fg2.Col = 3 Then
        If NulosN(Fg2.TextMatrix(Row, Col)) = 0 Then
            Fg2.TextMatrix(Row, 2) = ""
        End If
        TxtTotal.Text = Format(GRID_SUMAR_COL(Fg2, 2, 1, Fg2.Rows - 1), "0.00")
    End If
End Sub

Private Sub Fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Then KeyAscii = 0
    
    If KeyAscii = 13 Then Exit Sub
    ' validar los caracteres que se ingresan
    Select Case Col
        Case 1, 3
            KeyAscii = 0
            
        Case 2
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub

Private Sub Form_Activate()
    Dim SeEjecuto As Boolean
    Dim Rpta As Integer
    Dim cSQL As String
    
    If SeEjecuto = False Then
    
        SeEjecuto = True
    
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        cSQL = "SELECT pro_cronograma.*, mae_tipoproducto.descripcion AS destippro, [pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] & ', ' & [pla_empleados]![nom] AS apenom " _
            + vbCr + "FROM (pla_empleados RIGHT JOIN (pro_cronograma LEFT JOIN pro_emp ON pro_cronograma.idsup = pro_emp.id) ON pla_empleados.id = pro_emp.idemp) LEFT JOIN mae_tipoproducto ON pro_cronograma.idtippro = mae_tipoproducto.id " _
            + vbCr + "ORDER BY pro_cronograma.fchini DESC , pro_cronograma.id DESC;"
            
        RST_Busq RstLis, cSQL, xCon
        
        Set Dg1.DataSource = RstLis
        
    End If
End Sub

Sub MuestraSegundoTab()
    Dim cSQL As String
    Dim Rst As New ADODB.Recordset

    TxtIdSup.Text = RstLis("idsup")
    TxtIdSup_Validate True
    TxtFchIni.Valor = RstLis("fchini")
    TxtFchFin.Valor = RstLis("fchfin")
    TxtTipPro.Text = RstLis("idtippro")
    TxtTipPro_Validate True
    
    CalendarControl1.ActiveView.ShowDay (CDate(TxtFchIni.Valor))
    CalendarControl1.ActiveView.EnableVScroll False
    CalendarControl1.ActiveView.EnableHScroll True
    CalendarControl1.ViewType = xtpCalendarWeekView
    
    CalendarControl1.DataProvider.RemoveAllEvents
    
    LlenarDatos
    
    cSQL = "SELECT pro_cronogramadetprod.*, alm_inventario.descripcion AS descpro " _
        + vbCr + "FROM pro_cronogramadetprod LEFT JOIN alm_inventario ON pro_cronogramadetprod.idpro = alm_inventario.id " _
        + vbCr + "WHERE (((pro_cronogramadetprod.id)=" & RstLis("id") & "))"
    
    RST_Busq RstMatPro, cSQL, xCon
    
    RstMatPro.ActiveConnection = Nothing
End Sub

Private Sub Form_Load()
    Agregando = False
    SeEjecuto = False
    QueHace = 3
    iniciarCampos
End Sub

Sub Modificar()
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Cronograma de Produccion"
    QueHace = 2
    xHorIni = Time
    ActivaTool
    Bloquea
    TxtIdSup.SetFocus
End Sub

Sub Nuevo()
    Dim cSQL As String
    
    QueHace = 1
    xHorIni = Time
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Cronograma de Produccion"
    ActivaTool
    Bloquea
    Blanquea
    
    cSQL = "SELECT pro_cronogramadetprod.*, alm_inventario.descripcion AS descpro " _
        + vbCr + "FROM pro_cronogramadetprod LEFT JOIN alm_inventario ON pro_cronogramadetprod.idpro = alm_inventario.id " _
        + vbCr + "WHERE (((pro_cronogramadetprod.id)=99999));"
    
    RST_Busq RstMatPro, cSQL, xCon
    
    RstMatPro.ActiveConnection = Nothing
    
    TxtIdSup.SetFocus
End Sub

Sub Bloquea()
    TxtIdSup.Locked = Not TxtIdSup.Locked
    TxtFchIni.Locked = Not TxtFchIni.Locked
    TxtFchFin.Locked = Not TxtFchFin.Locked
    TxtTipPro.Locked = Not TxtTipPro.Locked
    
    CmdBusSup.Enabled = Not CmdBusSup.Enabled
    CmdBusTip.Enabled = Not CmdBusTip.Enabled
    
    ComboSemanas.Locked = Not ComboSemanas.Locked
End Sub

Sub Blanquea()
    TxtIdSup.Text = ""
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    ComboSemanas.ListIndex = 0
    TxtTipPro.Text = ""
    LblSupervisor.Caption = ""
    LblTipoProd.Caption = ""
    CalendarControl1.DataProvider.RemoveAllEvents
    If QueHace = 1 Then CalendarControl1.Visible = False: CmdProcesar.Visible = True
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 8000
        Me.Width = 12000
    End If
End Sub

Sub Cancelar()
    QueHace = 3
    Bloquea
    Label5.Caption = "Consultando Cronograma de Produccion"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    CalendarControl1.Visible = True
    CmdProcesar.Visible = False
    ActivaTool
End Sub

Private Sub menu_01_01_Click()
    If HitTest.ViewEvent Is Nothing Then Exit Sub
    
    If QueHace = 3 Then
        CmAcepta.Enabled = False
        Fg2.SelectionMode = flexSelectionByRow
        Fg2.Editable = flexEDNone
    Else
        CmAcepta.Enabled = True
        Fg2.SelectionMode = flexSelectionFree
        Fg2.Editable = flexEDKbdMouse
    End If
    
    If TxtTipPro.Text <> 1 Then Exit Sub
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    Dim xMatPri As String
    
    Fg2.Rows = 1
    
    centrarFrm Frame3
    
    TxtMP.Text = c_Event.Subject
    TxtCan.Text = c_Event.ReminderSoundFile
    TxtCan.Text = Format(TxtCan.Text, "0.00")
    xMatPri = TxtMP.Text
        
    xFchPro = Mid(c_Event.StartTime, 1, 10)
    xHorPro = Mid(c_Event.StartTime, 11, 6)
    
    xIdMatPri = Busca_Codigo(xMatPri, "descripcion", "id", "alm_inventario", "C", xCon)
    
    If xIdMatPri = 0 Then
        MsgBox "La materia prima especificada no existe", vbInformation + vbOKOnly + vbDefaultButton1
        Exit Sub
    End If
    
    ' MOSTRAMOS TODOS LOS PRODUCTOS DE LA MATERIA PRIMA
    RST_Busq Rst, "SELECT pro_redimiento.iditem, pro_redimiento.idpro, alm_inventario.descripcion " _
        & " FROM pro_redimiento LEFT JOIN alm_inventario ON pro_redimiento.idpro = alm_inventario.id " _
        & " WHERE (((pro_redimiento.iditem)=" & xIdMatPri & "))", xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Fg2.Rows = 1
        For A = 1 To Rst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = Rst("descripcion")
            Fg2.TextMatrix(A, 2) = ""
            Fg2.TextMatrix(A, 3) = 0
            Fg2.TextMatrix(A, 4) = Rst("idpro")
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
        
        If Rst.RecordCount = 1 Then
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = Format(TxtCan.Text, "0.00")
            If QueHace = 3 Then
                Fg2.TextMatrix(Fg2.Rows - 1, 3) = 0
                Fg2.TextMatrix(Fg2.Rows - 1, 2) = ""
            Else
                Fg2.TextMatrix(Fg2.Rows - 1, 3) = 1
            End If
            Fg2.Editable = flexEDNone
        End If
    End If
    
    ' MOSTRAMOS EL CHECK DE LOS PRODUCTOS QUE SE VAYAN A DEFINIR
    
    RstMatPro.Filter = adFilterNone
    If RstMatPro.RecordCount <> 0 Then
        RstMatPro.Filter = "iditem =" & xIdMatPri & " AND fchpro = " & xFchPro & " AND horpro = " & Format(xHorPro, "HH:mm") & ""
        If RstMatPro.RecordCount <> 0 Then
            RstMatPro.MoveFirst
            For A = 1 To RstMatPro.RecordCount
                For B = 1 To Fg2.Rows - 1
                    If NulosN(Fg2.TextMatrix(B, 4)) = RstMatPro("idpro") Then
                        Fg2.TextMatrix(B, 3) = 1
                        Fg2.TextMatrix(B, 2) = Format(RstMatPro("cantidad"), "0.00")
                        Exit For
                    End If
                Next B
                RstMatPro.MoveNext
                If RstMatPro.EOF = True Then Exit For
            Next A
        End If
    End If
    TxtTotal.Text = Format(GRID_SUMAR_COL(Fg2, 2, 1, Fg2.Rows - 1), "0.00")
    
    
    If visEvent Then CmAcepta.Enabled = False
    Frame3.Visible = True
End Sub

Private Sub menu_01_02_Click()
    Dim fIni As Date
    Dim fFin As Date
    Dim AllDay As Boolean

    CalendarControl1.ActiveView.GetSelection fIni, fFin, AllDay

    agregEvent = True
    TxtMatProd.Text = ""
    LblMatProd.Caption = ""
    TxtCant.Text = ""
    LblUnidad.Caption = ""
    DTPHoras.Value = 0
    
    centrarFrm FrmAdd
    LblDia.Caption = fIni
    FrmAdd.Visible = True
End Sub

Private Sub centrarFrm(ByRef frm As Frame)
    With frm
        .Left = ((Me.Width - .Width) / 2)
        .Top = ((Me.Height - .Height) / 2)
    End With
End Sub

Private Sub menu_01_03_Click()
    If HitTest.ViewEvent Is Nothing Then Exit Sub
    CalendarControl1.DataProvider.DeleteEvent c_Event
End Sub

Private Sub menu_01_04_Click()
    If HitTest.ViewEvent Is Nothing Then Exit Sub
    modifEvent = True
    LblDia.Caption = Format(c_Event.StartTime, "dd/mm/yyyy")
    
    TxtMatProd.Text = c_Event.ScheduleID
    LblMatProd.Caption = c_Event.Subject
    TxtCant.Text = c_Event.ReminderSoundFile
    DTPHoras.Value = Format(c_Event.Body, "HH:mm")
    LblUnidad.Caption = encontrarUnidad(c_Event.ScheduleID)
    
    centrarFrm FrmAdd
    
    If visEvent Then CmdAgregar.Enabled = False
    FrmAdd.Visible = True
End Sub

Private Sub Menu2_1_Click()
    Dim fIni As Date
    Dim fFin As Date
    Dim AllDay As Boolean

    CalendarControl1.ActiveView.GetSelection fIni, fFin, AllDay
    
    agregEvent = True
    TxtMatProd.Text = ""
    LblMatProd.Caption = ""
    TxtCant.Text = ""
    LblUnidad.Caption = ""
    DTPHoras.Value = 0
    
    centrarFrm FrmAdd
    LblDia.Caption = fIni
    FrmAdd.Visible = True
End Sub

Private Sub menu2_2_Click()
    If HitTest.ViewEvent Is Nothing Then Exit Sub
    CalendarControl1.DataProvider.DeleteEvent c_Event
End Sub

Private Function encontrarUnidad(idProd As String) As String
    Dim codigo As String
    Dim unidad As String
    codigo = Busca_Codigo(idProd, "id", "idunimed", "alm_inventario", "N", xCon)
    If NulosC(codigo) <> "" Then
        unidad = Busca_Codigo(codigo, "id", "abrev", "mae_unidades", "N", xCon)
    Else
        unidad = ""
    End If
    encontrarUnidad = unidad
End Function

Private Sub Menu2_3_Click()
    If HitTest.ViewEvent Is Nothing Then Exit Sub
    modifEvent = True
    
    LblDia.Caption = Format(c_Event.StartTime, "dd/mm/yyyy")
    
    TxtMatProd.Text = c_Event.ScheduleID
    LblMatProd.Caption = c_Event.Subject
            
    TxtCant.Text = c_Event.ReminderSoundFile
    
    DTPHoras.Value = Format(c_Event.Body, "HH:mm")
    
'    CmbHoras.ListIndex = NulosN(Mid(c_Event.StartTime, 12, 2))
'    CmbMinutos.ListIndex = NulosN(Mid(c_Event.StartTime, 15, 2))
    
    LblUnidad.Caption = encontrarUnidad(c_Event.ScheduleID)
    
    centrarFrm FrmAdd
    
    If visEvent Then CmdAgregar.Enabled = False
    FrmAdd.Visible = True
End Sub

Private Sub Menu3_1_Click()
    If HitTest.ViewEvent Is Nothing Then Exit Sub
    
    If QueHace = 3 Then
        CmAcepta.Enabled = False
        Fg2.SelectionMode = flexSelectionByRow
        Fg2.Editable = flexEDNone
    Else
        CmAcepta.Enabled = True
        Fg2.SelectionMode = flexSelectionFree
        Fg2.Editable = flexEDKbdMouse
    End If
    
    If TxtTipPro.Text <> 1 Then Exit Sub
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    Dim xMatPri As String
    
    Toolbar1.Enabled = False
    Fg2.Rows = 1
    
    centrarFrm Frame3
    
    TxtMP.Text = c_Event.Subject
    TxtCan.Text = c_Event.ReminderSoundFile
    TxtCan.Text = Format(TxtCan.Text, "0.00")
    xMatPri = TxtMP.Text
        
    xFchPro = Mid(c_Event.StartTime, 1, 10)
    xHorPro = Mid(c_Event.StartTime, 11, 6)
    
    xIdMatPri = Busca_Codigo(xMatPri, "descripcion", "id", "alm_inventario", "C", xCon)
    
    If xIdMatPri = 0 Then
        MsgBox "La materia prima especificada no existe", vbInformation + vbOKOnly + vbDefaultButton1
        Exit Sub
    End If
    
    ' MOSTRAMOS TODOS LOS PRODUCTOS DE LA MATERIA PRIMA
    RST_Busq Rst, "SELECT pro_redimiento.iditem, pro_redimiento.idpro, alm_inventario.descripcion " _
        & " FROM pro_redimiento LEFT JOIN alm_inventario ON pro_redimiento.idpro = alm_inventario.id " _
        & " WHERE (((pro_redimiento.iditem)=" & xIdMatPri & "))", xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Fg2.Rows = 1
        For A = 1 To Rst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = Rst("descripcion")
            Fg2.TextMatrix(A, 2) = ""
            Fg2.TextMatrix(A, 3) = 0
            Fg2.TextMatrix(A, 4) = Rst("idpro")
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
        
        If Rst.RecordCount = 1 Then
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = Format(TxtCan.Text, "0.00")
            If QueHace = 3 Then
                Fg2.TextMatrix(Fg2.Rows - 1, 3) = 0
                Fg2.TextMatrix(Fg2.Rows - 1, 2) = ""
            Else
                Fg2.TextMatrix(Fg2.Rows - 1, 3) = 1
            End If
            Fg2.Editable = flexEDNone
        End If
    End If
    Frame3.Visible = True
    
    ' MOSTRAMOS EL CHECK DE LOS PRODUCTOS QUE SE VAYAN A DEFINIR
    
    RstMatPro.Filter = adFilterNone
    If RstMatPro.RecordCount <> 0 Then
        RstMatPro.Filter = "iditem =" & xIdMatPri & " AND fchpro = " & xFchPro & " AND horpro = " & Format(xHorPro, "HH:mm") & ""
        If RstMatPro.RecordCount <> 0 Then
            RstMatPro.MoveFirst
            For A = 1 To RstMatPro.RecordCount
                For B = 1 To Fg2.Rows - 1
                    If NulosN(Fg2.TextMatrix(B, 4)) = RstMatPro("idpro") Then
                        Fg2.TextMatrix(B, 3) = 1
                        Fg2.TextMatrix(B, 2) = Format(RstMatPro("cantidad"), "0.00")
                        Exit For
                    End If
                Next B
                RstMatPro.MoveNext
                If RstMatPro.EOF = True Then Exit For
            Next A
        End If
    End If
    TxtTotal.Text = Format(GRID_SUMAR_COL(Fg2, 2, 1, Fg2.Rows - 1), "0.00")
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    Else
        Frame3.Visible = False
        FrmAdd.Visible = False
    End If
End Sub

Sub Eliminar()
    TabOne1.CurrTab = 0
    Dim Rpta As Integer
    
    Rpta = MsgBox("¿Esta seguro de eliminar el cronograma seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_cronograma WHERE id = " & RstLis("id") & ""
        xCon.Execute "DELETE * FROM pro_cronogramadet WHERE id = " & RstLis("id") & ""
        xCon.Execute "DELETE * FROM pro_cronogramatarea WHERE id = " & RstLis("id") & ""
        xCon.Execute "DELETE * FROM pro_cronogramadetprod WHERE id = " & RstLis("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstLis("id") & " AND idform = " & IdMenuActivo
        
        
        RstLis.Requery
        Dg1.Refresh
        MsgBox "El cronograma se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstLis.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        If TabOne1.CurrTab = 0 Then RstLis.Filter = "": TDB_FiltroLimpiar Dg1
        If TabOne1.CurrTab = 1 Then CmdProcesar_Click
    End If
    
    If Button.Index = 12 Then Imprimir
    
    If Button.Index = 14 Then
        Set RstLis = Nothing
        Unload Me
    End If
End Sub

Sub Imprimir()
    Dim Rst As New ADODB.Recordset
    
    If NulosN(RstLis("idtippro")) = 1 Then
        RST_Busq Rst, "TRANSFORM sum(pro_cronogramadetprod.cantidad) AS PromedioDecantidad SELECT pro_cronogramadetprod.iditem, alm_inventario_1.descripcion AS desmatpri, " _
            & " mae_unidades.abrev, alm_inventario.descripcion AS descprod, Sum(pro_cronogramadetprod.cantidad) AS [TotalFila]" _
            & " FROM ((pro_cronogramadetprod LEFT JOIN alm_inventario ON pro_cronogramadetprod.idpro = alm_inventario.id) LEFT JOIN alm_inventario AS alm_inventario_1 " _
            & " ON pro_cronogramadetprod.iditem = alm_inventario_1.id) LEFT JOIN mae_unidades ON alm_inventario_1.idunimed = mae_unidades.id " _
            & " Where (((pro_cronogramadetprod.ID) = " & RstLis("id") & ")) GROUP BY pro_cronogramadetprod.iditem, alm_inventario_1.descripcion, mae_unidades.abrev, " _
            & " alm_inventario.descripcion, pro_cronogramadetprod.id ORDER BY alm_inventario_1.descripcion, alm_inventario.descripcion " _
            & " PIVOT Format([fchpro],'dd-mm-yy')", xCon
    Else
        RST_Busq Rst, "TRANSFORM Sum(pro_cronogramadet.cantidad) AS SumaDecantidad SELECT pro_cronogramadet.iditem, alm_inventario.descripcion, mae_unidades.abrev, " _
            & " Sum(pro_cronogramadet.cantidad) AS TotalFila FROM (pro_cronogramadet LEFT JOIN alm_inventario ON pro_cronogramadet.iditem = alm_inventario.id) " _
            & " LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id Where (((pro_cronogramadet.ID) = " & RstLis("id") & ")) " _
            & " GROUP BY pro_cronogramadet.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_cronogramadet.id PIVOT Format([fchpro],'dd-mm-yy')", xCon
    End If
    
    Dim Li As Integer
    Dim strSource As String
    Dim xArea, xEmp, xDir, xCuerpo, xCad  As String
    Dim xEmpleado As String
    Dim Pagina As Integer
    Dim Lineas As Integer
    
    Set oPDF = New cPDF
    Dim A, B, C As Integer
    xNumPag = 0
    Dim xTipPro As String
    
On Error GoTo Cerrado
    
    If oPDF.PDFCreate(App.Path & "\pro00001.pdf") = True Then
        
        oPDF.Fonts.Add "Tit", Times_BoldItalic, WinAnsiEncoding
        oPDF.Fonts.Add "Head", Times_Italic, WinAnsiEncoding
        oPDF.Fonts.Add "Cont", Courier, WinAnsiEncoding
        oPDF.Fonts.Add "CB", Courier_Bold, WinAnsiEncoding
        oPDF.Fonts.Add "Time", Times_Roman, WinAnsiEncoding
        
        CrearCabecera
        Dim xFilaAct As Integer
        Dim xPosX As Integer
        Dim xFch As Date
        
        oPDF.WTextBox 40, 30, 10, 750, "CRONOGRAMA DE PRODUCCION (" & RstLis("destippro") & ")", "CB", 10, hCenter, vMiddle, vbBlack, 0, vbRed
        oPDF.WTextBox 52, 30, 10, 750, "DEL " & RstLis("fchini") & " AL " & RstLis("fchfin"), "CB", 10, hCenter, vMiddle, vbBlack, 0, vbRed
        
        If NulosN(RstLis("idtippro")) = 1 Then
            oPDF.WTextBox 68, 30, 19, 150, "MATERIA PRIMA", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            oPDF.WTextBox 68, 180, 19, 30, "UNI. MED.", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            oPDF.WTextBox 68, 210, 19, 250, "PRODUCTO", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            xPosX = 460
        Else
            oPDF.WTextBox 68, 30, 19, 250, "PRODUCTO", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            oPDF.WTextBox 68, 280, 19, 30, "UNI. MED.", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            xPosX = 310
        End If
        
        ' IMPRIMIMOS EL ROTULO DE LAS FECHAS
        For xFch = RstLis("fchini") To RstLis("fchfin")
            oPDF.WTextBox 68, xPosX, 19, 45, Format(xFch, "dd/mm/yy"), "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            
            xPosX = xPosX + 45
        Next xFch
        
        ' IMPRIMIMOS EL ROTULO DEL TOTAL
        oPDF.WTextBox 68, xPosX, 19, 45, "TOTAL", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
                 
        xFilaInicial = 88
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                If NulosN(RstLis("idtippro")) = 1 Then
                    oPDF.WTextBox xFilaInicial, 30, 10, 150, Rst("desmatpri"), "CB", 8, hLeft, vMiddle, vbBlack, 0, vbBlack
                    oPDF.WTextBox xFilaInicial, 180, 10, 30, Rst("abrev"), "CB", 8, hCenter, vMiddle, vbBlack, 0, vbRed
                    oPDF.WTextBox xFilaInicial, 210, 10, 250, Rst("descprod"), "CB", 8, hLeft, vMiddle, vbBlack, 0, vbRed
                    xPosX = 460
                Else
                    oPDF.WTextBox xFilaInicial, 30, 10, 250, Rst("descripcion"), "CB", 8, hLeft, vMiddle, vbBlack, 0, vbBlack
                    oPDF.WTextBox xFilaInicial, 280, 10, 30, Rst("abrev"), "CB", 8, hCenter, vMiddle, vbBlack, 0, vbRed
                    xPosX = 310
                End If
                
                For xFch = RstLis("fchini") To RstLis("fchfin")
                    If RstRegistroBuscaCampo(Rst, Format(xFch, "dd-mm-yy")) = True Then
                        oPDF.WTextBox xFilaInicial, xPosX, 10, 45, Format(NulosN(Rst(Format(xFch, "dd-mm-yy"))), "0.00"), "CB", 8, hRight, vMiddle, vbBlack, 0, vbBlack
                    End If
                    xPosX = xPosX + 45
                Next xFch
                
                ' IMPRIMIMOS EL TOTAL DE LA FILA
                oPDF.WTextBox xFilaInicial, xPosX, 10, 45, Format(NulosN(Rst("TotalFila")), "0.00"), "CB", 8, hRight, vMiddle, vbBlack, 0, vbBlack
                xFilaInicial = xFilaInicial + 10
                
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If
        
        oPDF.PDFClose
        Set oPDF = Nothing
        Shell ("rundll32.exe url.dll,FileProtocolHandler " & Trim(App.Path) & ("\pro00001.pdf")), vbMaximizedFocus
    Else
        Set oPDF = Nothing
        MsgBox "No se Puede Mostrar Documento pro00001.pdf, psoblemente el archivo ya se encuentra abierto", vbCritical, "Error"
    End If
    Exit Sub
    
Cerrado:
    'Resume
    If Err.Number = 1 Then
    End If
End Sub

Sub CrearCabecera()
    Dim xTelEmp, xNumDoc As String
    
    'oPDF.NewPage A4_Vertical ', 525, 675
    oPDF.NewPage A4_Horizontal  ', 525, 675
    xNumPag = xNumPag + 1
    
    oPDF.WTextBox 15, 30, 8, 50, "EMPRESA", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 15, 105, 8, 5, ":", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 15, 111, 8, 150, NomEmp, "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    
    oPDF.WTextBox 23, 30, 8, 50, "Nº R.U.C.", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 23, 105, 8, 5, ":", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 23, 111, 8, 100, NumRUC, "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    
    oPDF.WTextBox 15, 700, 8, 50, "Nº PAGINA", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 15, 750, 8, 5, ":", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 15, 753, 8, 50, Format(xNumPag, "000"), "CB", 8, hRight, vTop, RGB(0, 0, 128), , vbRed
    
    oPDF.WTextBox 23, 700, 8, 50, "FCH. IMPR", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 23, 750, 8, 5, ":", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 23, 753, 8, 50, Format(Date, "dd/mm/yy"), "CB", 8, hRight, vTop, RGB(0, 0, 128), , vbRed
    
    oPDF.WLineTo 30, 36, 800, 36
    oPDF.LineStroke
End Sub

Private Sub TxtFchFin_Change()
    Dim fech As String
    If Not cambio Then
        If TxtFchFin.Valor <> "" Then
            fech = TxtFchFin.Valor
            ComboSemanas.Text = DatePart("ww", NulosC(CDate(fech)), vbMonday, vbFirstFullWeek)
        End If
    End If
End Sub

Private Sub TxtFchIni_Change()
    Dim fech As String
    If Not cambio Then
        If TxtFchIni.Valor <> "" Then
            fech = TxtFchIni.Valor
            ComboSemanas.Text = DatePart("ww", NulosC(CDate(fech)), vbMonday, vbFirstFullWeek)
        End If
    End If
End Sub

Private Sub TxtIdSup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtIdSup_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusSup_Click
    End If
End Sub

Private Sub TxtIdSup_Validate(Cancel As Boolean)
    If NulosN(TxtIdSup.Text) = 0 Then
        TxtIdSup.Text = ""
        Exit Sub
    Else
        Dim Rst As New ADODB.Recordset
        Dim xSqlCad As String
        xSqlCad = "SELECT pro_emp.*, pla_empleados.nombre, pro_emp.id " _
            & " FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id " _
            & " Where (((pro_empdet.idfun) = 2) And ((pro_emp.ID) = " & Val(TxtIdSup.Text) & ")) ORDER BY pla_empleados.nombre"

        Set Rst = BuscaConCriterio(xSqlCad, xCon)
        
        If Rst.RecordCount <> 0 Then
            LblSupervisor.Caption = Rst("nombre")
        Else
            TxtIdSup.Text = ""
            LblSupervisor.Caption = ""
        End If
        
        Set Rst = Nothing
    End If
End Sub

Private Sub TxtMatProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtMatProd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdAddMatProd_Click
    End If
End Sub

Private Sub TxtMatProd_Validate(Cancel As Boolean)
    If NulosN(TxtMatProd.Text) = 0 Then
        TxtMatProd.Text = ""
        Exit Sub
    Else
        Dim codigo As String
        LblMatProd.Caption = Busca_Codigo(TxtMatProd.Text, "id", "descripcion", "alm_inventario", "N", xCon)
        codigo = Busca_Codigo(TxtMatProd.Text, "id", "idunimed", "alm_inventario", "N", xCon)
        If NulosC(codigo) <> "" Then LblUnidad.Caption = Busca_Codigo(codigo, "id", "abrev", "mae_unidades", "N", xCon)
        If NulosC(LblMatProd.Caption) = "" Then
            TxtMatProd.Text = ""
            LblUnidad.Caption = ""
            TxtMatProd.SetFocus
        Else
            TxtCant.SetFocus
        End If
    End If
End Sub

Private Sub TxtTipPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtTipPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTip_Click
    End If
End Sub

Private Sub TxtTipPro_Validate(Cancel As Boolean)
    If NulosN(TxtTipPro.Text) = 0 Then
        TxtTipPro.Text = ""
        Exit Sub
    Else
        LblTipoProd.Caption = Busca_Codigo(TxtTipPro.Text, "id", "descripcion", " mae_tipoproducto", "N", xCon)
        If NulosC(LblTipoProd.Caption) = "" Then
            TxtTipPro.Text = ""
        End If
    End If
End Sub


'Metodos para arrastrar el Frame
''''''''''''''''''''''''''''''''

Private Sub FrmAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y - 1500
    FrmAdd.ZOrder 0
    FrmAdd.Drag
End Sub

Private Sub Frame3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y - 1500
    Frame3.ZOrder 0
    Frame3.Drag
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move X - OrigFX, Y - OrigFY
End Sub

Private Sub CalendarControl1_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move X - OrigFX, Y - OrigFY
End Sub
