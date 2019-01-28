VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmMantDocumCta 
   Caption         =   "Contabilidad - Asignar Cuenta Contable  a Documentos"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5775
      Left            =   -15
      TabIndex        =   0
      Top             =   375
      Width           =   8715
      _cx             =   15372
      _cy             =   10186
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
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5355
         Left            =   9360
         TabIndex        =   4
         Top             =   375
         Width           =   8625
         Begin VB.Frame Frame3 
            Height          =   2355
            Left            =   255
            TabIndex        =   21
            Top             =   1575
            Width           =   8130
            Begin VB.CommandButton CmdBusMon 
               Height          =   240
               Left            =   2415
               Picture         =   "FrmMantDocumCta.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   26
               Tag             =   "b"
               Top             =   915
               Width           =   240
            End
            Begin VB.CommandButton CmdBusTipDoc 
               Height          =   240
               Left            =   2415
               Picture         =   "FrmMantDocumCta.frx":0132
               Style           =   1  'Graphical
               TabIndex        =   23
               Tag             =   "b"
               Top             =   555
               Width           =   240
            End
            Begin VB.ComboBox CmbTipoOperac 
               Height          =   315
               ItemData        =   "FrmMantDocumCta.frx":0264
               Left            =   1770
               List            =   "FrmMantDocumCta.frx":026E
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   1575
               Width           =   1335
            End
            Begin VB.CommandButton CmbBusCta 
               Height          =   240
               Left            =   2415
               Picture         =   "FrmMantDocumCta.frx":0281
               Style           =   1  'Graphical
               TabIndex        =   22
               Tag             =   "b"
               Top             =   1260
               Width           =   240
            End
            Begin VB.TextBox TxtTipDoc 
               Height          =   300
               Left            =   1770
               Locked          =   -1  'True
               MaxLength       =   3
               TabIndex        =   24
               Tag             =   "a"
               Text            =   "TxtTipDoc"
               Top             =   525
               Width           =   915
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   1770
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   25
               Tag             =   "a"
               Text            =   "TxtIdMon"
               Top             =   885
               Width           =   915
            End
            Begin VB.TextBox TxtIdCta 
               Height          =   300
               Left            =   1770
               Locked          =   -1  'True
               TabIndex        =   27
               Tag             =   "a"
               Text            =   "TxtIdCta"
               Top             =   1230
               Width           =   915
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
               Left            =   2730
               TabIndex        =   36
               Tag             =   "a"
               Top             =   885
               Width           =   5160
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Index           =   5
               Left            =   225
               TabIndex        =   35
               Top             =   900
               Width           =   585
            End
            Begin VB.Label LblNomDoc 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblNomDoc"
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
               Left            =   2730
               TabIndex        =   34
               Tag             =   "a"
               Top             =   525
               Width           =   5160
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Documento"
               Height          =   195
               Index           =   9
               Left            =   225
               TabIndex        =   33
               Top             =   525
               Width           =   1185
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta"
               Height          =   195
               Index           =   0
               Left            =   225
               TabIndex        =   32
               Top             =   1230
               Width           =   510
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Operación"
               Height          =   195
               Index           =   7
               Left            =   225
               TabIndex        =   31
               Top             =   1590
               Width           =   1320
            End
            Begin VB.Label LblNumCta 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblNumCta"
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
               Left            =   2730
               TabIndex        =   30
               Tag             =   "a"
               Top             =   1230
               Width           =   1485
            End
            Begin VB.Label LblDesCta 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDesCta"
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
               Left            =   4260
               TabIndex        =   28
               Tag             =   "a"
               Top             =   1230
               Width           =   3630
            End
         End
         Begin VB.Frame Frame4 
            Height          =   840
            Left            =   240
            TabIndex        =   5
            Top             =   5955
            Width           =   11295
            Begin VB.TextBox TxtISC 
               Alignment       =   1  'Right Justify
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
               Left            =   8865
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   12
               TabStop         =   0   'False
               Text            =   "TxtISC"
               Top             =   420
               Width           =   1100
            End
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "Agregar Item"
               Enabled         =   0   'False
               Height          =   360
               Left            =   270
               TabIndex        =   11
               Top             =   285
               Width           =   1335
            End
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "Eliminar Item"
               Enabled         =   0   'False
               Height          =   360
               Left            =   1695
               TabIndex        =   10
               Top             =   285
               Width           =   1335
            End
            Begin VB.TextBox TxtInafecto 
               Alignment       =   1  'Right Justify
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
               Left            =   6390
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   9
               TabStop         =   0   'False
               Text            =   "TxtInafect"
               Top             =   420
               Width           =   1100
            End
            Begin VB.TextBox TxtBruto 
               Alignment       =   1  'Right Justify
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
               Left            =   5235
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   8
               TabStop         =   0   'False
               Text            =   "TxtBruto"
               Top             =   420
               Width           =   1100
            End
            Begin VB.TextBox TxtIGV 
               Alignment       =   1  'Right Justify
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
               Left            =   7545
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   7
               TabStop         =   0   'False
               Text            =   "TxtIGV"
               Top             =   420
               Width           =   1100
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
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
               Left            =   10035
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   6
               TabStop         =   0   'False
               Text            =   "TxtTotal"
               Top             =   420
               Width           =   1100
            End
            Begin VB.Label LblRotulo 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V. (        ) "
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
               Left            =   7545
               TabIndex        =   18
               Top             =   195
               Width           =   1260
            End
            Begin VB.Label Label1 
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
               Height          =   195
               Index           =   3
               Left            =   8865
               TabIndex        =   17
               Top             =   195
               Width           =   495
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000005&
               Index           =   1
               X1              =   3255
               X2              =   3255
               Y1              =   90
               Y2              =   870
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000003&
               Index           =   0
               X1              =   3240
               X2              =   3240
               Y1              =   105
               Y2              =   885
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Inafecto"
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
               Index           =   1
               Left            =   6390
               TabIndex        =   16
               Top             =   195
               Width           =   720
            End
            Begin VB.Label LblIgvTasa 
               Alignment       =   2  'Center
               Caption         =   "LblIgvTasa"
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
               Height          =   225
               Left            =   8190
               TabIndex        =   15
               Top             =   195
               Width           =   420
            End
            Begin VB.Label Label1 
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
               Height          =   195
               Index           =   0
               Left            =   5235
               TabIndex        =   14
               Top             =   195
               Width           =   990
            End
            Begin VB.Label Label1 
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
               Height          =   195
               Index           =   2
               Left            =   10035
               TabIndex        =   13
               Top             =   195
               Width           =   450
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Documento con Cta. Contable"
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
            Left            =   135
            TabIndex        =   19
            Top             =   45
            Width           =   8370
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5355
         Left            =   45
         TabIndex        =   1
         Top             =   375
         Width           =   8625
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   5010
            Left            =   30
            TabIndex        =   2
            ToolTipText     =   "Click derecho para Aceptar o Rechazar"
            Top             =   345
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   8837
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
            Columns(1).Caption=   "Documento"
            Columns(1).DataField=   "DESDOCU"
            Columns(1).NumberFormat=   "000000"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Moneda"
            Columns(2).DataField=   "DESMONEDA"
            Columns(2).NumberFormat=   "Short Date"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tipo Operacion"
            Columns(3).DataField=   "Tipo_Operacion"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Nº Cuenta"
            Columns(4).DataField=   "NUMCTA"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Des. Cuenta"
            Columns(5).DataField=   "DESPLANCTA"
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
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2963"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2884"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1535"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1455"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=3069"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2990"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1852"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1773"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=4710"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=4630"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta Documentos Con Cta Contable"
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
            Left            =   135
            TabIndex        =   3
            Top             =   45
            Width           =   8385
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6450
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
            Picture         =   "FrmMantDocumCta.frx":03B3
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumCta.frx":08F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumCta.frx":0C89
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumCta.frx":0E0D
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumCta.frx":1261
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumCta.frx":1379
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumCta.frx":18BD
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumCta.frx":1E01
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumCta.frx":1F15
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumCta.frx":2029
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumCta.frx":247D
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumCta.frx":25E9
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   8700
      _ExtentX        =   15346
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
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmMantDocumCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace  As Integer, vStr As String
Dim SeEjecuto As Boolean
Dim RstDocumCta As New ADODB.Recordset, RsConGen As New ADODB.Recordset
Dim RstTmp As New ADODB.Recordset
Dim CaracteresNumericos As String, CaracteresNumericos2 As String
Dim xHorIni As Date

Dim mIdRegistro& '--identificador del registro
Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Sub ActivaTool()
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
Private Function fVerifSiExisteReg(pIdDoc As Long, pIdMon As Long, pTipoOperac As Boolean) As Boolean
    Set RsConGen = New ADODB.Recordset
    RsConGen.CursorLocation = adUseClient
    vStr = "SELECT * FROM mae_documentocta WHERE iddoc = " & pIdDoc & " AND idmon = " & pIdMon & " "
    If pTipoOperac = False Then
        vStr = vStr & "AND tipope = FALSE"
    Else
        vStr = vStr & "AND tipope = TRUE"
    End If
    RsConGen.Open vStr, xCon, adOpenForwardOnly, adLockReadOnly
    If RsConGen.RecordCount > 0 Then
        fVerifSiExisteReg = True
    Else
        fVerifSiExisteReg = False
    End If
    Set RsConGen = Nothing
End Function
Function Grabar() As Boolean
    If NulosC(TxtTipDoc.Text) = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Function
    End If
    If NulosC(TxtIdMon.Text) = "" Then
        MsgBox "No ha especificado la moneda", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If
    If NulosC(TxtIdCta.Text) = "" Then
        MsgBox "No ha especificado la cuenta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdCta.SetFocus
        Exit Function
    End If
    If NulosC(CmbTipoOperac.Text) = "" Then
        MsgBox "No ha especificado el tipo de operacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        CmbTipoOperac.SetFocus
        Exit Function
    End If
    
    Dim RstGrabar As New ADODB.Recordset
    
    If QueHace = 1 Then
        RST_Busq RstGrabar, "SELECT mae_documentocta.iddoc, mae_documentocta.idmon, mae_documentocta.tipope " _
            & " From mae_documentocta " _
            & " WHERE (((mae_documentocta.iddoc)=" & NulosN(TxtTipDoc.Text) & ") AND ((mae_documentocta.idmon)=" & NulosN(TxtIdMon.Text) & ") AND ((mae_documentocta.tipope)=" & IIf(CmbTipoOperac.ListIndex = 0, 0, -1) & ")); ", xCon
        
        If RstGrabar.RecordCount <> 0 Then
            Set RstGrabar = Nothing
            MsgBox "No puede continuar, otro registro esta configurado de la misma manera", vbInformation, xTitulo
            Exit Function
        End If
        
    Else
        RST_Busq RstGrabar, "SELECT mae_documentocta.iddoc, mae_documentocta.idmon, mae_documentocta.tipope " _
            & " From mae_documentocta " _
            & " WHERE (((mae_documentocta.id)<>" & RstDocumCta("id") & ") AND ((mae_documentocta.iddoc)=" & NulosN(TxtTipDoc.Text) & ") AND ((mae_documentocta.idmon)=" & NulosN(TxtIdMon.Text) & ") AND ((mae_documentocta.tipope)=" & IIf(CmbTipoOperac.ListIndex = 0, 0, -1) & ")); ", xCon
        
        If RstGrabar.RecordCount <> 0 Then
            Set RstGrabar = Nothing
            MsgBox "No se puede registrar las modificaciones, otro registro esta configurado de la misma manera", vbInformation, xTitulo
            Exit Function
        End If
    End If
    
    Set RstGrabar = Nothing
    
    Dim xIdDocCta As Long, xIdMon As Long, xTipoOperac As Boolean
    Dim xId As Double
    On Error GoTo LaCague
    xCon.BeginTrans
    
    If QueHace = 1 Then
''        'VERIFICAR SI YA EXISTE REGISTRO
''        If fVerifSiExisteReg(NulosN(TxtTipDoc.Text), NulosN(TxtIdMon.Text), IIf(CmbTipoOperac.ListIndex = 0, False, True)) = True Then
''            MsgBox "El registro que esta tratando de ingresar ya existe...!", vbExclamation, "Mensaje...!"
''            Exit Function
''        End If
                
        xId = HallaCodigoTabla("mae_documentocta", xCon, "id")
        
        RST_Busq RstGrabar, "SELECT TOP 1 * FROM mae_documentocta", xCon
        
        RstGrabar.AddNew
        
        RstGrabar("id") = xId
        
    Else
        xId = RstDocumCta("id")

        RST_Busq RstGrabar, "SELECT * FROM mae_documentocta WHERE id = " & xId & "", xCon
        
'        xIdDocCta = NulosN(RstDocumCta("iddoc"))
'        xIdMon = NulosN(RstDocumCta("IDMONEDA"))
'        If RstDocumCta("tipope").Value = False Then
'            xTipoOperac = False
'            RST_Busq RstGrabar, "SELECT * FROM mae_documentocta WHERE iddoc = " & xIdDocCta & " AND idmon = " & xIdMon & " AND tipope = FALSE", xCon
'        Else
'            xTipoOperac = True
'            RST_Busq RstGrabar, "SELECT * FROM mae_documentocta WHERE iddoc = " & xIdDocCta & " AND idmon = " & xIdMon & " AND tipope = TRUE", xCon
'        End If
        
        
        
    End If
    
    mIdRegistro = xId
    
    RstGrabar("iddoc") = NulosN(TxtTipDoc.Text)
    RstGrabar("idmon") = NulosN(TxtIdMon.Text)
    RstGrabar("tipope") = IIf(CmbTipoOperac.ListIndex = 0, 0, -1)
    RstGrabar("idcuen") = NulosN(TxtIdCta.Text)
    RstGrabar.Update
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    
    MsgBox "El documento de cuenta se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstGrabar = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function
Sub Buscar()
    TabOne1.CurrTab = 0
    Dim vBoolTipoOper As Boolean
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    xRs.CursorLocation = adUseClient
    Dim xcant As Integer
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Documento":          xCampos(0, 1) = "Documento":      xCampos(0, 2) = "3500":  xCampos(0, 3) = "C"
    xCampos(1, 0) = "Moneda":             xCampos(1, 1) = "Moneda":         xCampos(1, 2) = "1500":  xCampos(1, 3) = "C"
    xCampos(2, 0) = "Tipo de Operacion":  xCampos(2, 1) = "Tipo_Operacion": xCampos(2, 2) = "2000":  xCampos(2, 3) = "C"
    xCampos(3, 0) = "Num Cta":            xCampos(3, 1) = "Num_Cta":        xCampos(3, 2) = "1000":  xCampos(3, 3) = "C"
    
'    xForm.SQLCad = "SELECT mae_prov.nombre AS nomprov, mae_prov.numruc, mae_trabajadores.apellnom, com_ordencompra.id " _
'        & " FROM mae_trabajadores RIGHT JOIN (mae_prov RIGHT JOIN com_ordencompra ON mae_prov.id = com_ordencompra.idpro) " _
'        & " ON mae_trabajadores.id = com_ordencompra.idaut ORDER BY mae_trabajadores.apellnom, com_ordencompra.id"
        
    xform.SqlCad = "SELECT mae_documento.descripcion AS Documento, mae_moneda.descripcion AS Moneda, IIF(mae_documentocta.tipope = FALSE, 'Compras', 'Ventas') AS Tipo_Operacion, mae_documentocta.tipope, con_planctas.descripcion AS Cuenta, con_planctas.cuenta AS Num_Cta, mae_documentocta.iddoc AS ID_DOC, mae_documentocta.idmon AS ID_MONEDA " _
        & "FROM mae_documento INNER JOIN (mae_moneda INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id=mae_documentocta.idcuen) ON mae_moneda.id=mae_documentocta.idmon) ON mae_documento.id=mae_documentocta.iddoc " _
        & "ORDER BY mae_documento.descripcion"
'    vStr = "SELECT mae_documento.descripcion AS Documento, mae_moneda.descripcion AS Moneda, IIF(mae_documentocta.tipope = FALSE, 'Compras', 'Ventas') AS Tipo_Operacion, mae_documentocta.tipope, con_planctas.descripcion AS Cuenta, con_planctas.cuenta AS Num_Cta, mae_documentocta.iddoc AS ID_DOC, mae_documentocta.idmon AS ID_MONEDA "
'    vStr = vStr & "FROM mae_documento INNER JOIN (mae_moneda INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id=mae_documentocta.idcuen) ON mae_moneda.id=mae_documentocta.idmon) ON mae_documento.id=mae_documentocta.iddoc "
'    vStr = vStr & "ORDER BY mae_documento.descripcion"
    
    
    xform.Titulo = "Buscando Documento de Cuenta"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "Documento"
    xform.CampoBusca = "Documento"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        RstDocumCta.MoveFirst
        vBoolTipoOper = xRs("tipope")
        If vBoolTipoOper = False Then
            RstDocumCta.Filter = "iddoc = " & NulosN(xRs("ID_DOC")) & " AND IDMONEDA = " & NulosN(xRs("ID_MONEDA")) & " AND tipope = false"
        Else
'            RstDocumCta.Find "iddoc = " & nulosn(xRs("ID_DOC")) & " AND IDMONEDA = " & nulosn(xRs("ID_MONEDA")) & ""
            RstDocumCta.Filter = "iddoc = " & NulosN(xRs("ID_DOC")) & " AND IDMONEDA = " & NulosN(xRs("ID_MONEDA")) & " AND tipope = TRUE"
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub
Sub MuestraSegundoTab()
    'SELECT mae_documentocta.iddoc, mae_documento.descripcion AS DESDOCU, mae_moneda.descripcion AS DESMONEDA, mae_documentocta.tipope, con_planctas.descripcion AS DESPLANCTA, con_planctas.cuenta AS NUMCTA, mae_moneda.simbolo, mae_documentocta.idcuen, mae_moneda.id AS IDMONEDA, mae_documento.id AS IDDOCUMENTO
    Blanquea
    TxtTipDoc.Text = Trim(RstDocumCta("iddoc"))
    LblNomDoc.Caption = Trim(RstDocumCta("DESDOCU"))
    TxtIdMon.Text = Trim(RstDocumCta("IDMONEDA"))
    LblMoneda.Caption = Trim(RstDocumCta("DESMONEDA"))
    TxtIdCta.Text = Trim(RstDocumCta("idcuen"))
    LblNumCta.Caption = Trim(RstDocumCta("NUMCTA"))
    LblDesCta.Caption = Trim(RstDocumCta("DESPLANCTA"))
    If RstDocumCta("tipope").Value = True Then
        CmbTipoOperac.ListIndex = 1
    Else
        CmbTipoOperac.ListIndex = 0
    End If
End Sub
Sub Cancelar()
    QueHace = 3
    ActivaTool
    Blanquea
    Bloquea False
    Label5.Caption = "Detalle de Documento con Cta. Contable"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
End Sub


Sub Eliminar()

    Dim Rpta As Integer
'RST_Busq RstDocumCta, "SELECT mae_documentocta.iddoc, mae_documento.descripcion AS DESDOCU, mae_moneda.descripcion AS DESMONEDA, IIF(mae_documentocta.tipope = FALSE, 'Compras', 'Ventas') AS Tipo_Pedido, mae_documentocta.tipope, con_planctas.descripcion AS DESPLANCTA, con_planctas.cuenta AS NUMCTA, mae_moneda.simbolo, mae_documentocta.idcuen, mae_moneda.id AS IDMONEDA, mae_documento.id AS IDDOCUMENTO " _
'& "FROM con_planctas INNER JOIN (mae_documento INNER JOIN (mae_moneda INNER JOIN mae_documentocta ON mae_moneda.id = mae_documentocta.idmon) ON mae_documento.id = mae_documentocta.iddoc) ON con_planctas.id = mae_documentocta.idcuen"


    Rpta = MsgBox("¿Esta seguro de eliminar el documento de cuenta seleccionada?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        TabOne1.CurrTab = 0
        xCon.Execute "DELETE * FROM mae_documentocta WHERE id = " & NulosN(RstDocumCta("id")) & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstDocumCta("id") & " AND idform = " & IdMenuActivo
                
        MsgBox "El documento con cuenta se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstDocumCta.Requery
        Dg1.Refresh
        Dg1.SetFocus
    End If
    
End Sub
Sub Modificar()
    QueHace = 2
    ActivaTool
    Blanquea
    Bloquea True
    TxtTipDoc.Locked = True: CmdBusTipDoc.Enabled = False
    TxtIdMon.Locked = True: CmdBusMon.Enabled = False
    
    CmbTipoOperac.Enabled = False
    Label5.Caption = "Modificando Documento con Cta. Contable"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    MuestraSegundoTab
    xHorIni = Time
    TxtTipDoc.SetFocus
End Sub
Sub Bloquea(pBool As Boolean)
    Dim obj As Object
    For Each obj In Me.Controls
        If obj.Tag = "a" And TypeName(obj) = "TextBox" Then
            obj.Locked = Not pBool
        ElseIf obj.Tag = "b" And TypeName(obj) = "CommandButton" Then
            obj.Enabled = True
        End If
    Next
    CmbTipoOperac.Enabled = pBool
End Sub
Sub Blanquea()
    Dim obj As Object
    For Each obj In Me.Controls
        If obj.Tag = "a" Then
            obj = ""
        End If
    Next
    CmbTipoOperac.ListIndex = -1
End Sub
Sub Nuevo()
    QueHace = 1
    ActivaTool
    Blanquea
    Bloquea True
    Label5.Caption = "Agregando Documento con Cta. Contable"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    xHorIni = Time
    TxtTipDoc.SetFocus
End Sub

Private Sub CmbBusCta_Click()
    If QueHace = 3 Then Exit Sub
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "N° Cuenta":     xCampos(0, 1) = "cuenta":       xCampos(0, 2) = "3000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":   xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "5000":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "Codigo":        xCampos(2, 1) = "id":           xCampos(2, 2) = "2000":   xCampos(2, 3) = "N"
    
    xform.SqlCad = "SELECT cuenta, descripcion, id FROM con_planctas ORDER BY cuenta"
    
    xform.Titulo = "Buscando Cuenta Contable"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdCta.Text = xRs("id")
            LblNumCta.Caption = xRs("cuenta")
            LblDesCta.Caption = xRs("descripcion")
            If CmbTipoOperac.Enabled = True Then
                CmbTipoOperac.SetFocus
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SqlCad = "SELECT * FROM mae_moneda"
    
    xform.Titulo = "Buscando Tipo de Moneda"
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
            TxtIdCta.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDoc_Click()
    If QueHace = 3 Then Exit Sub
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SqlCad = "SELECT descripcion, id FROM mae_documento"
    
    xform.Titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipDoc.Text = xRs("id")
            LblNomDoc.Caption = xRs("descripcion")
            TxtIdMon.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
    MuestraSegundoTab
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstDocumCta.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstDocumCta("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    'Modificado: 10/01/11 Johan Castro
    '            Agregar linea de codigo para bloquear accesos de usuarios


    If SeEjecuto = False Then
        Blanquea
'        Dim Rpta As Integer
        SeEjecuto = True
    
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        '----------------------------------------------

        vStr = "SELECT mae_documentocta.id, mae_documentocta.iddoc, mae_documento.descripcion AS DESDOCU, mae_moneda.descripcion AS DESMONEDA, IIF(mae_documentocta.tipope = FALSE, 'Compras', 'Ventas') AS Tipo_Operacion, mae_documentocta.tipope, con_planctas.descripcion AS DESPLANCTA, con_planctas.cuenta AS NUMCTA, mae_moneda.simbolo, mae_documentocta.idcuen, mae_moneda.id AS IDMONEDA, mae_documento.id AS IDDOCUMENTO "
        vStr = vStr & "FROM mae_documento INNER JOIN (mae_moneda INNER JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON mae_moneda.id = mae_documentocta.idmon) ON mae_documento.id = mae_documentocta.iddoc"
'        RST_Busq RstDocumCta, "SELECT mae_documentocta.iddoc, mae_documento.descripcion AS DESDOCU, mae_moneda.descripcion AS DESMONEDA, IIF(mae_documentocta.tipope = FALSE, 'Compras', 'Ventas') AS Tipo_Operacion, mae_documentocta.tipope, con_planctas.descripcion AS DESPLANCTA, con_planctas.cuenta AS NUMCTA, mae_moneda.simbolo, mae_documentocta.idcuen, mae_moneda.id AS IDMONEDA, mae_documento.id AS IDDOCUMENTO " _
'            & "FROM con_planctas INNER JOIN (mae_documento INNER JOIN (mae_moneda INNER JOIN mae_documentocta ON mae_moneda.id = mae_documentocta.idmon) ON mae_documento.id = mae_documentocta.iddoc) ON con_planctas.id = mae_documentocta.idcuen", xCon
        RST_Busq RstDocumCta, vStr, xCon

        Set Dg1.DataSource = RstDocumCta
''        If RstDocumCta.RecordCount = 0 Then
''            Rpta = MsgBox("No se ha registrado ningun documento de cuenta, ¿Desea agregar una ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
''            If Rpta = vbYes Then
''                Nuevo
''            Else
''                Set RstDocumCta = Nothing
''                Unload Me
''                Exit Sub
''            End If
''        End If
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    TabOne1.CurrTab = 0
    CaracteresNumericos = "0123456789." & Chr(8)
    CaracteresNumericos2 = "0123456789." & Chr(8) & Chr(13)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RstDocumCta = Nothing
    Set RsConGen = Nothing
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then
        Modificar
    End If
    If Button.Index = 3 Then
        Eliminar
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstDocumCta.Requery
            Dg1.Refresh
            Dg1.SetFocus
            '--------------------------
            If RstDocumCta.RecordCount <> 0 Then
                RstDocumCta.MoveFirst
                RstDocumCta.Find "iddoc=" & mIdRegistro
                If RstDocumCta.EOF = True Then RstDocumCta.MoveFirst
            End If
            '--------------------------
        End If
    End If
    If Button.Index = 6 Then Cancelar
    If Button.Index = 9 Then RstDocumCta.Filter = adFilterNone
    
    If Button.Index = 10 Then Buscar
        
'    If Button.Index = 12 Then 'ImprimirOrden
    
    If Button.Index = 14 Then
'        Set RstDocumCta = Nothing
        Unload Me
    End If
End Sub

Private Sub TxtIdCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If NulosC(TxtIdCta.Text) <> "" Then
            RstTmp.CursorLocation = adUseClient
            Set RstTmp = BuscaConCriterio("SELECT id, cuenta, descripcion FROM con_planctas WHERE id = " & NulosN(TxtIdCta.Text) & "", xCon)
            
            If RstTmp.RecordCount <> 0 Then
                LblNumCta.Caption = Trim(RstTmp("cuenta"))
                LblDesCta.Caption = Trim(RstTmp("descripcion"))
            Else
                TxtIdCta.Text = ""
                LblNumCta.Caption = "": LblDesCta.Caption = ""
            End If
        End If
        Set RstTmp = Nothing
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If NulosC(TxtIdMon.Text) <> "" Then
            RstTmp.CursorLocation = adUseClient
            Set RstTmp = BuscaConCriterio("SELECT id, descripcion FROM mae_moneda WHERE id = " & NulosN(TxtIdMon.Text) & "", xCon)
            
            If RstTmp.RecordCount <> 0 Then
                LblMoneda.Caption = RstTmp("descripcion")
            Else
                TxtIdMon.Text = ""
                LblMoneda.Caption = ""
            End If
        End If
        Set RstTmp = Nothing
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If NulosC(TxtTipDoc.Text) <> "" Then
            RstTmp.CursorLocation = adUseClient
            Set RstTmp = BuscaConCriterio("SELECT id, descripcion FROM mae_documento WHERE id = " & NulosN(TxtTipDoc.Text) & "", xCon)
            
            If RstTmp.RecordCount <> 0 Then
                LblNomDoc.Caption = RstTmp("descripcion")
            Else
                TxtTipDoc.Text = ""
                LblNomDoc.Caption = ""
            End If
        End If
        Set RstTmp = Nothing
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub
