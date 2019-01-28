VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManOrdProd 
   Caption         =   "Produccion - Orden de Producción"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
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
      ForeColor       =   &H80000001&
      Height          =   5040
      Index           =   0
      Left            =   30
      TabIndex        =   18
      Top             =   8010
      Visible         =   0   'False
      Width           =   4980
      Begin VB.Frame Frame12 
         Caption         =   "La tarea Empieza al : "
         Height          =   2085
         Left            =   150
         TabIndex        =   35
         Top             =   300
         Width           =   4660
         Begin VB.OptionButton OptTarea 
            Caption         =   "Segun Linea de Producción"
            Height          =   255
            Index           =   3
            Left            =   210
            TabIndex        =   40
            Top             =   1770
            Width           =   2325
         End
         Begin VB.OptionButton OptTarea 
            Caption         =   "Transcurrido los minutos de la tarea anterior"
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   39
            Top             =   1140
            Width           =   3855
         End
         Begin VB.TextBox TxtPctje 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   300
            Left            =   2145
            MaxLength       =   12
            TabIndex        =   38
            Text            =   "TxtPctje"
            Top             =   795
            Width           =   840
         End
         Begin VB.OptionButton OptTarea 
            Caption         =   "Transcurrir un porcentaje de la tarea anterior"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   37
            Top             =   510
            Width           =   4425
         End
         Begin VB.OptionButton OptTarea 
            Caption         =   "Finalizar la tarea anterior"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   36
            Top             =   270
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker DTPMinutos 
            Height          =   345
            Left            =   2160
            TabIndex        =   41
            Top             =   1410
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm"
            Format          =   16842755
            UpDown          =   -1  'True
            CurrentDate     =   40606
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "HH:mm"
            Height          =   195
            Index           =   4
            Left            =   3075
            TabIndex        =   45
            Top             =   1440
            Width           =   525
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Minutos"
            Height          =   195
            Index           =   7
            Left            =   1245
            TabIndex        =   44
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   2
            Left            =   3075
            TabIndex        =   43
            Top             =   840
            Width           =   120
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje"
            Height          =   195
            Index           =   6
            Left            =   1245
            TabIndex        =   42
            Top             =   840
            Width           =   765
         End
      End
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   2
         Left            =   4680
         Picture         =   "FrmManOrdProd.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   34
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VB.Frame Frame11 
         Caption         =   "Incluir Horas de Refrigerio?"
         Height          =   945
         Left            =   150
         TabIndex        =   25
         Top             =   2400
         Width           =   4660
         Begin VB.OptionButton OptHoras 
            Caption         =   "Si"
            Height          =   225
            Index           =   0
            Left            =   300
            TabIndex        =   27
            Top             =   450
            Width           =   555
         End
         Begin VB.OptionButton OptHoras 
            Caption         =   "No"
            Height          =   225
            Index           =   1
            Left            =   1000
            TabIndex        =   26
            Top             =   450
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPHorIni 
            Height          =   345
            Left            =   2700
            TabIndex        =   28
            Top             =   130
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm"
            Format          =   16842755
            UpDown          =   -1  'True
            CurrentDate     =   40606
         End
         Begin MSComCtl2.DTPicker DTPHorFin 
            Height          =   345
            Left            =   2700
            TabIndex        =   29
            Top             =   500
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm"
            Format          =   16842755
            UpDown          =   -1  'True
            CurrentDate     =   40606
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "HH:mm )"
            Height          =   195
            Index           =   8
            Left            =   3700
            TabIndex        =   33
            Top             =   230
            Width           =   615
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "( Inicio"
            Height          =   195
            Index           =   30
            Left            =   2100
            TabIndex        =   32
            Top             =   225
            Width           =   465
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "HH:mm )"
            Height          =   195
            Index           =   10
            Left            =   3705
            TabIndex        =   31
            Top             =   585
            Width           =   615
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "( Fin"
            Height          =   195
            Index           =   9
            Left            =   2100
            TabIndex        =   30
            Top             =   585
            Width           =   300
         End
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Cancelar"
         Height          =   345
         Index           =   17
         Left            =   3645
         TabIndex        =   24
         Top             =   4570
         Width           =   1155
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "Aceptar"
         Height          =   345
         Index           =   16
         Left            =   2430
         TabIndex        =   23
         Top             =   4570
         Width           =   1155
      End
      Begin VB.Frame Frame3 
         Caption         =   "Opciones Diversas"
         Height          =   1125
         Left            =   150
         TabIndex        =   19
         Top             =   3360
         Width           =   4665
         Begin VB.CheckBox cknumtar 
            Caption         =   "Limitar Numero de Tareas segun Linea"
            Height          =   195
            Left            =   180
            TabIndex        =   22
            Top             =   270
            Width           =   3195
         End
         Begin VB.CheckBox cknumper 
            Caption         =   "Limitar Numero de Personal segun Linea"
            Height          =   195
            Left            =   180
            TabIndex        =   21
            Top             =   540
            Width           =   3285
         End
         Begin VB.CheckBox ckperarea 
            Caption         =   "Limitar Seleccion de Personal por Area"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   810
            Width           =   3045
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   8295
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   2
         X1              =   4950
         X2              =   4950
         Y1              =   0
         Y2              =   5000
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   4950
         Y1              =   5000
         Y2              =   5000
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opciones de Procesado"
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
         Index           =   27
         Left            =   105
         TabIndex        =   46
         Top             =   45
         Width           =   2040
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   40
         Top             =   30
         Width           =   4860
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6570
      Top             =   45
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
            Picture         =   "FrmManOrdProd.frx":02EC
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdProd.frx":0830
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdProd.frx":0BC2
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdProd.frx":0D46
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdProd.frx":119A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdProd.frx":12B2
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdProd.frx":17F6
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdProd.frx":1D3A
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdProd.frx":1E4E
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdProd.frx":1F62
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdProd.frx":23B6
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdProd.frx":2522
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdProd.frx":2A6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Anular Registro"
               EndProperty
            EndProperty
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
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Materiales"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Linea"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7590
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   11925
      _cx             =   21034
      _cy             =   13388
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7170
         Left            =   45
         TabIndex        =   7
         Top             =   375
         Width           =   11835
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6555
            Left            =   30
            TabIndex        =   10
            Top             =   480
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11562
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
            Columns(1).Caption=   "Mes"
            Columns(1).DataField=   "desmes"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripcion"
            Columns(2).DataField=   "desitem"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Método Valorización"
            Columns(3).DataField=   "desmetval"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Aplica Gas.Fab."
            Columns(4).DataField=   "desaplgasfab"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Tip. Dist. Gas. Fab."
            Columns(5).DataField=   "destipdisgasfab"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2064"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1984"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=5345"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=5265"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=4974"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=4895"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=3493"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=3413"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=3731"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=3651"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=3"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14,.alignment=2"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15,.alignment=3"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=70,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14,.alignment=2"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(64)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(67)  =   ":id=35,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   "Named:id=36:Selected"
            _StyleDefs(69)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(70)  =   "Named:id=37:Caption"
            _StyleDefs(71)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(72)  =   "Named:id=38:HighlightRow"
            _StyleDefs(73)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(74)  =   "Named:id=39:EvenRow"
            _StyleDefs(75)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(76)  =   "Named:id=40:OddRow"
            _StyleDefs(77)  =   ":id=40,.parent=33"
            _StyleDefs(78)  =   "Named:id=41:RecordSelector"
            _StyleDefs(79)  =   ":id=41,.parent=34"
            _StyleDefs(80)  =   "Named:id=42:FilterBar"
            _StyleDefs(81)  =   ":id=42,.parent=33"
         End
         Begin VB.Label LblMes 
            AutoSize        =   -1  'True
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   10020
            TabIndex        =   9
            Top             =   90
            Width           =   720
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Libro de Costo de Producción"
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
            Left            =   45
            TabIndex        =   8
            Top             =   45
            Width           =   11685
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7170
         Left            =   12570
         TabIndex        =   5
         Top             =   375
         Width           =   11835
         Begin VB.Frame Frame9 
            Caption         =   "[ Gastos de Fábrica ]"
            ForeColor       =   &H00800000&
            Height          =   1035
            Left            =   5910
            TabIndex        =   55
            Top             =   330
            Width           =   5835
            Begin VB.CommandButton Cmd 
               Caption         =   "&Configurar"
               Enabled         =   0   'False
               Height          =   350
               Index           =   1
               Left            =   4320
               TabIndex        =   62
               ToolTipText     =   "Procesar las Tareas del Producto Seleccionado"
               Top             =   540
               Width           =   1400
            End
            Begin VB.OptionButton opttipdiscta 
               Caption         =   "Distribuida"
               Height          =   225
               Index           =   1
               Left            =   3120
               TabIndex        =   61
               Top             =   650
               Width           =   1065
            End
            Begin VB.OptionButton opttipdiscta 
               Caption         =   "Global"
               Height          =   225
               Index           =   0
               Left            =   2190
               TabIndex        =   60
               Top             =   650
               Width           =   885
            End
            Begin VB.OptionButton optdisgasfab 
               Caption         =   "Ventas"
               Height          =   225
               Index           =   1
               Left            =   1140
               TabIndex        =   58
               Top             =   650
               Width           =   795
            End
            Begin VB.OptionButton optdisgasfab 
               Caption         =   "Todos"
               Height          =   225
               Index           =   0
               Left            =   150
               TabIndex        =   56
               Top             =   650
               Width           =   795
            End
            Begin VB.Line Line4 
               X1              =   2090
               X2              =   2090
               Y1              =   210
               Y2              =   950
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000005&
               X1              =   2100
               X2              =   2100
               Y1              =   210
               Y2              =   950
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Distribubión de Cta.:"
               Height          =   195
               Left            =   2250
               TabIndex        =   59
               Top             =   330
               Width           =   2010
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Aplicar Distribucion a:"
               Height          =   195
               Left            =   120
               TabIndex        =   57
               Top             =   330
               Width           =   1530
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "[ Datos de Producción ]"
            Height          =   5685
            Left            =   0
            TabIndex        =   16
            Top             =   1380
            Width           =   11775
            Begin SizerOneLibCtl.TabOne TabOne2 
               Height          =   3735
               Left            =   60
               TabIndex        =   17
               Top             =   1830
               Width           =   11655
               _cx             =   20558
               _cy             =   6588
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
               Appearance      =   2
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   700
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FrontTabColor   =   -2147483633
               BackTabColor    =   12632256
               TabOutlineColor =   -2147483632
               FrontTabForeColor=   -2147483630
               Caption         =   "   &Mat. Pri.  |    &Man. Obr.   |  &Gas. fab.  "
               Align           =   0
               CurrTab         =   0
               FirstTab        =   0
               Style           =   0
               Position        =   1
               AutoSwitch      =   -1  'True
               AutoScroll      =   -1  'True
               TabPreview      =   -1  'True
               ShowFocusRect   =   -1  'True
               TabsPerPage     =   0
               BorderWidth     =   0
               BoldCurrent     =   -1  'True
               DogEars         =   -1  'True
               MultiRow        =   0   'False
               MultiRowOffset  =   200
               CaptionStyle    =   0
               TabHeight       =   0
               TabCaptionPos   =   4
               TabPicturePos   =   0
               CaptionEmpty    =   ""
               Separators      =   0   'False
               Begin VB.Frame Frame6 
                  Caption         =   "[ Personal ]"
                  Height          =   3360
                  Left            =   12600
                  TabIndex        =   49
                  Top             =   45
                  Width           =   11565
                  Begin VB.CommandButton Cmd 
                     Caption         =   "Eliminar Todos"
                     Enabled         =   0   'False
                     Height          =   330
                     Index           =   11
                     Left            =   10110
                     TabIndex        =   53
                     ToolTipText     =   "Eliminar Todos"
                     Top             =   2220
                     Width           =   1400
                  End
                  Begin VB.CommandButton Cmd 
                     Caption         =   "&Seleccionar"
                     Enabled         =   0   'False
                     Height          =   330
                     Index           =   7
                     Left            =   10110
                     TabIndex        =   52
                     TabStop         =   0   'False
                     ToolTipText     =   "Agregar Personal de una Lista"
                     Top             =   600
                     Width           =   1400
                  End
                  Begin VB.CommandButton Cmd 
                     Caption         =   "&Eliminar"
                     Enabled         =   0   'False
                     Height          =   330
                     Index           =   10
                     Left            =   10110
                     TabIndex        =   51
                     TabStop         =   0   'False
                     ToolTipText     =   "Eliminar Personal"
                     Top             =   1860
                     Width           =   1400
                  End
                  Begin VB.CommandButton Cmd 
                     Caption         =   "Agregar"
                     Enabled         =   0   'False
                     Height          =   330
                     Index           =   6
                     Left            =   10110
                     TabIndex        =   50
                     ToolTipText     =   "Agregar Personal"
                     Top             =   240
                     Width           =   1400
                  End
                  Begin VSFlex7Ctl.VSFlexGrid fg 
                     Height          =   2970
                     Index           =   2
                     Left            =   60
                     TabIndex        =   54
                     Top             =   270
                     Width           =   9945
                     _cx             =   17542
                     _cy             =   5239
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
                     AllowUserResizing=   1
                     SelectionMode   =   0
                     GridLines       =   1
                     GridLinesFixed  =   2
                     GridLineWidth   =   1
                     Rows            =   2
                     Cols            =   4
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   0
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmManOrdProd.frx":2DFC
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
               Begin VB.Frame Frame8 
                  Caption         =   "[ Tareas ]"
                  ForeColor       =   &H00800000&
                  Height          =   3360
                  Left            =   12300
                  TabIndex        =   48
                  Top             =   45
                  Width           =   11565
                  Begin VSFlex7Ctl.VSFlexGrid fg 
                     Height          =   3015
                     Index           =   4
                     Left            =   60
                     TabIndex        =   64
                     Top             =   240
                     Width           =   11415
                     _cx             =   20135
                     _cy             =   5318
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
                     AllowUserResizing=   1
                     SelectionMode   =   0
                     GridLines       =   1
                     GridLinesFixed  =   2
                     GridLineWidth   =   1
                     Rows            =   2
                     Cols            =   7
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   0
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmManOrdProd.frx":2E80
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
               Begin VB.Frame Frame7 
                  Caption         =   "[ Insumos ]"
                  ForeColor       =   &H00800000&
                  Height          =   3360
                  Left            =   45
                  TabIndex        =   47
                  Top             =   45
                  Width           =   11565
                  Begin VSFlex7Ctl.VSFlexGrid fg 
                     Height          =   3015
                     Index           =   3
                     Left            =   90
                     TabIndex        =   63
                     Top             =   240
                     Width           =   11370
                     _cx             =   20055
                     _cy             =   5318
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
                     Rows            =   2
                     Cols            =   8
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   0
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmManOrdProd.frx":2F53
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
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   1455
               Index           =   0
               Left            =   60
               TabIndex        =   65
               Top             =   300
               Width           =   11625
               _cx             =   20505
               _cy             =   2566
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
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   25
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManOrdProd.frx":3049
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
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   1
            Text            =   "TxtNumDoc"
            Top             =   690
            Width           =   4635
         End
         Begin VB.CommandButton Cmd 
            Height          =   240
            Index           =   3
            Left            =   1830
            Picture         =   "FrmManOrdProd.frx":330D
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1050
            Width           =   240
         End
         Begin VB.ComboBox cbMes 
            Height          =   315
            Left            =   1170
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   390
            Width           =   4665
         End
         Begin VB.TextBox txtidmetval 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   2
            Text            =   "txtidmetval"
            Top             =   1020
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Met. Val."
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   1065
            Width           =   630
         End
         Begin VB.Label lblmetval 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblmetval"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2115
            TabIndex        =   14
            Top             =   1020
            Width           =   3690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mes"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   12
            Top             =   420
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   765
            Width           =   840
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Libro de Costo de Producción"
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
            TabIndex        =   6
            Top             =   75
            Width           =   11685
         End
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Insertar"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Ver Receta"
      End
   End
End
Attribute VB_Name = "FrmManOrdProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------VARIABLES DE ESTADO DE FORMULARIO
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean               ' para saber cuando se este agregando FILAS AL CONTROL grid de productos
Dim IdMenuActivo As Integer            ' INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date                    ' ESPECIFICA LA HORA DE INICIO
Dim mIdRegistro&                       ' identificador del registro
Dim mMesActivo As Integer              ' indica el mes activo
Dim OrigFX As Long
Dim OrigFY As Long
Dim fOrdenLista As Boolean              ' especfica el orden de la lista de la consulta
'***********************************************
'-----------------------VARIABLES DE FORMULARIO
'***********************************************
Dim Rst As New ADODB.Recordset
Dim RstOrdProd As New ADODB.Recordset
Dim cSQL As String
Dim ESTADOANTERIOR_ As Double
Dim CORR_ As Double
'-----------------------PROPIEDADES DE PROCESADO
' -----ESTRUCTURA
Private Type PROPIEDADESPROCESADO_
    MODOTAREA_  As Integer
    PORCENTAJE_  As Double
    MINUTOS_ As Date
    INCLUIRREFRIGERIO_ As Boolean
    HORINIREFRIGERIO_ As Date
    HORFINREFRIGERIO_  As Date
    LIMITARNUMEROTAREAS_ As Boolean
    LIMITARNUMEROPERSONAL_ As Boolean
    LIMITARSELPERSONAL_ As Boolean
End Type
' -----TIPO
Dim PROPIEDADES_ As PROPIEDADESPROCESADO_
' ----------------------DEFINICION DE COLUMNAS
Private Enum COLUMNACABECERA_
    RECETA_ = 1
    UNIMED_
    CANTIDAD_
    LINEA_
    EFICIENCIA_
    HORINI_
    HORFIN_
    FCHFIN_
    NUMOPE_
    REPROC_
    IDRECETA_
    IDLINEA_
    IDUNIMED_
End Enum

Private Enum COLUMNADETALLETAREA_
    SEL_ = 1
    TAREA_
    DURACION_
    HORINI_
    HORFIN_
    NUMOP_
    CANTIDADSUM_
    CANTIDADPROC_
    FCHINI_
    FCHFIN_
    AREA_
    RESPONSABLE_
    IDTAR_
    IDAREA_
    IDRESP_
End Enum

Private Enum COLUMNADETALLEPERS_
    DNI_ = 1
    NOMBRE_
    IDPER_
End Enum

Private Enum COLUMNADETALLEREPROC_
    LOTE_ = 1
    ALMACEN_
    CANTIDAD_
    IDLOTE_
    IDLOTEDET_
    IDALM_
End Enum

Private Enum COLUMNADETALLEINSUMOS_
    INSUMO_ = 1
    UNIMED_
    CANTIDAD_
    IDINSUMO_
    IDUNIMED_
End Enum
' ----------------------DEFINICION DE ESTADOS
Const ESTADOPENDIENTE_ = 1
Const ESTADOPROCESADO_ = 2
Const ESTADOAPROBADO_ = 3
Const ESTADOANULADO_ = 4

Private Sub aplicarPropiedades(MODIFICAR_ As Boolean, Optional CARGAR_ As Boolean = False)
    If MODIFICAR_ Then
        If OptTarea(0).Value = True Then PROPIEDADES_.MODOTAREA_ = 0
        If OptTarea(1).Value = True Then PROPIEDADES_.MODOTAREA_ = 1
        If OptTarea(2).Value = True Then PROPIEDADES_.MODOTAREA_ = 2
        If OptTarea(3).Value = True Then PROPIEDADES_.MODOTAREA_ = 3
        
        If OptHoras(0).Value = True Then PROPIEDADES_.INCLUIRREFRIGERIO_ = True
        If OptHoras(1).Value = True Then PROPIEDADES_.INCLUIRREFRIGERIO_ = False
        
        PROPIEDADES_.PORCENTAJE_ = NulosN(TxtPctje.Text)
        PROPIEDADES_.MINUTOS_ = Format(DTPMinutos.Value, "HH:mm")
        PROPIEDADES_.HORINIREFRIGERIO_ = Format(DTPHorIni.Value, "HH:mm")
        PROPIEDADES_.HORFINREFRIGERIO_ = Format(DTPHorFin.Value, "HH:mm")
        PROPIEDADES_.LIMITARNUMEROPERSONAL_ = cknumper.Value
        PROPIEDADES_.LIMITARNUMEROTAREAS_ = cknumtar.Value
        PROPIEDADES_.LIMITARSELPERSONAL_ = ckperarea.Value
    End If
    
    If CARGAR_ Then
        OptTarea(PROPIEDADES_.MODOTAREA_).Value = True
        If PROPIEDADES_.INCLUIRREFRIGERIO_ Then
            OptHoras(0).Value = True
        Else
            OptHoras(1).Value = True
        End If
        TxtPctje.Text = PROPIEDADES_.PORCENTAJE_
        DTPMinutos.Value = PROPIEDADES_.MINUTOS_
        DTPHorIni.Value = PROPIEDADES_.HORINIREFRIGERIO_
        DTPHorFin.Value = PROPIEDADES_.HORFINREFRIGERIO_
        
        If PROPIEDADES_.LIMITARNUMEROPERSONAL_ Then cknumper.Value = 1 Else cknumper.Value = 0
        If PROPIEDADES_.LIMITARNUMEROTAREAS_ Then cknumtar.Value = 1 Else cknumtar.Value = 0
        If PROPIEDADES_.LIMITARSELPERSONAL_ Then ckperarea.Value = 1 Else ckperarea.Value = 0
    End If
End Sub

Private Function procesarLineaProduccion() As Boolean
    Dim xRs As New ADODB.Recordset
    Dim RECORDSET_ As New ADODB.Recordset
    Dim CANTIDADAPROCESAR_ As Double
    Dim CANTIDAD_ As Double
    Dim IDLINEA_ As Integer
    Dim IDITEM_ As Double
    Dim HORINI_ As String
    Dim FECHINI_ As Date
    Dim ESNUEVO_ As Boolean
    Dim PORCENTAJEAPLICADO_ As Double
    Dim A As Integer
            
    '*********************
    ' SE VERIFICAN CAMPOS
    '*********************
    ' ----------------------------------------------------FECHA DE PRODUCCION
    If Not IsDate(TxtFchPro.Valor) Then
        MsgBox "Ingrese Fecha de Programación", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchPro.SetFocus
        procesarLineaProduccion = False: Exit Function
    End If
    ' ----------------------------------------------------ITEM
    If NulosN(txtIdItem.Text) = 0 Then
        MsgBox "Ingrese Producto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txtIdItem.SetFocus
        procesarLineaProduccion = False: Exit Function
    End If
    ' ----------------------------------------------------RESPONSABLE
    If NulosN(TxtIdResp.Text) = 0 Then
        MsgBox "Ingrese Encargado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdResp.SetFocus
        procesarLineaProduccion = False: Exit Function
    End If
    ' ----------------------------------------------------CANTIDAD
    If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.CANTIDAD_)) = 0 Then
        MsgBox "Ingrese Cantidad", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).Col = COLUMNACABECERA_.CANTIDAD_
        fg(0).SetFocus
        procesarLineaProduccion = False: Exit Function
    End If
    ' ----------------------------------------------------HORA DE INICIO
    If fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.HORINI_) = "" Then
        MsgBox "Ingrese Hora de Inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).Col = COLUMNACABECERA_.HORINI_
        fg(0).SetFocus
        procesarLineaProduccion = False: Exit Function
    End If
    
    IDLINEA_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.IDLINEA_))
    HORINI_ = Format(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.HORINI_), "HH:mm")
    FECHINI_ = CDate(Format(TxtFchPro.Valor, "dd/mm/yyyy"))
    CANTIDAD_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.CANTIDAD_))
    PORCENTAJEAPLICADO_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.EFICIENCIA_))
    
    
    If xRs.State = 0 Then
        cSQL = "SELECT TOP 1 * FROM pro_ordenprodtar;"
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        DEFINIR_RST_TMP RECORDSET_, xRs
    End If
    
    With fg(1)
        For A = 1 To .Rows - 1
            RECORDSET_.AddNew
            RECORDSET_("idord") = 0
            RECORDSET_("idtar") = .TextMatrix(A, COLUMNADETALLETAREA_.IDTAR_)
            RECORDSET_("cantsum") = .TextMatrix(A, COLUMNADETALLETAREA_.CANTIDADSUM_)
            RECORDSET_("cantproc") = .TextMatrix(A, COLUMNADETALLETAREA_.CANTIDADPROC_)
            RECORDSET_("numop") = .TextMatrix(A, COLUMNADETALLETAREA_.NUMOP_)
            If Not IsDate(.TextMatrix(A, COLUMNADETALLETAREA_.FCHINI_)) Then
                RECORDSET_("fchini") = Null
            Else
                RECORDSET_("fchini") = .TextMatrix(A, COLUMNADETALLETAREA_.FCHINI_)
            End If
            If Not IsDate(.TextMatrix(A, COLUMNADETALLETAREA_.FCHFIN_)) Then
                RECORDSET_("fchfin") = Null
            Else
                RECORDSET_("fchfin") = .TextMatrix(A, COLUMNADETALLETAREA_.FCHFIN_)
            End If
            If Not IsDate(.TextMatrix(A, COLUMNADETALLETAREA_.HORINI_)) Then
                RECORDSET_("horini") = Null
            Else
                RECORDSET_("horini") = .TextMatrix(A, COLUMNADETALLETAREA_.HORINI_)
            End If
            If Not IsDate(.TextMatrix(A, COLUMNADETALLETAREA_.HORFIN_)) Then
                RECORDSET_("horfin") = Null
            Else
                RECORDSET_("horfin") = .TextMatrix(A, COLUMNADETALLETAREA_.HORFIN_)
            End If
            RECORDSET_("durtar") = .TextMatrix(A, COLUMNADETALLETAREA_.DURACION_)
            RECORDSET_("idsubresp") = .TextMatrix(A, COLUMNADETALLETAREA_.IDRESP_)
            RECORDSET_("idarea") = .TextMatrix(A, COLUMNADETALLETAREA_.IDAREA_)
            RECORDSET_("activo") = .TextMatrix(A, COLUMNADETALLETAREA_.SEL_)
            RECORDSET_.Update
        Next A
    End With
    
    CANTIDADAPROCESAR_ = CANTIDAD_ / caracteristicaLinea(2, IDLINEA_, , RECORDSET_)
    Set RECORDSET_ = procesarCronograma(0, IDLINEA_, RECORDSET_, CANTIDADAPROCESAR_, HORINI_, FECHINI_, PORCENTAJEAPLICADO_)
    
    fg(1).Rows = fg(1).FixedRows
    If RECORDSET_.State = 0 Then procesarLineaProduccion = False: Exit Function
    If RECORDSET_.RecordCount = 0 Then procesarLineaProduccion = False: Exit Function
    
    RECORDSET_.MoveFirst
    Agregando = True
    With fg(1)
        While Not RECORDSET_.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.SEL_) = NulosN(RECORDSET_("activo"))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.TAREA_) = UCase(Busca_Codigo(NulosN(RECORDSET_("idtar")), "id", "descripcion", "pro_tareas", "N", xCon))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.DURACION_) = Format(NulosC(RECORDSET_("durtar")), "HH:mm")
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.HORINI_) = Format(NulosC(RECORDSET_("horini")), FORMAT_HORA_SIN_SEGUNDO)
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.HORFIN_) = Format(NulosC(RECORDSET_("horfin")), FORMAT_HORA_SIN_SEGUNDO)
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.NUMOP_) = NulosN(RECORDSET_("numop"))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.CANTIDADSUM_) = Format(NulosN(RECORDSET_("cantsum")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.CANTIDADPROC_) = Format(NulosN(RECORDSET_("cantproc")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.FCHINI_) = Format(NulosC(RECORDSET_("fchini")), FORMAT_DATE)
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.FCHFIN_) = Format(NulosC(RECORDSET_("fchfin")), FORMAT_DATE)
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.AREA_) = UCase(Busca_Codigo(NulosN(RECORDSET_("idarea")), "id", "descripcion", "mae_area", "N", xCon))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.RESPONSABLE_) = UCase(Busca_Codigo(NulosN(RECORDSET_("idsubresp")), "id", "nombre", "pla_empleados", "N", xCon))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.IDTAR_) = NulosN(RECORDSET_("idtar"))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.IDAREA_) = NulosN(RECORDSET_("idarea"))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.IDRESP_) = NulosN(RECORDSET_("idsubresp"))
            RECORDSET_.MoveNext
        Wend
    End With
    
    Dim HORADEFIN_ As String
    Dim FECHADEFIN_ As String
    
    ' ------------------------------NUMERO DE OPERARIOS
    fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.NUMOPE_) = GRID_SUMAR_COL(fg(1), COLUMNADETALLETAREA_.NUMOP_)
    For A = 1 To fg(1).Rows - 1
        If NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.SEL_)) = -1 Then
            HORADEFIN_ = Format(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.HORFIN_), FORMAT_HORA_SIN_SEGUNDO)
            FECHADEFIN_ = Format(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.FCHFIN_), FORMAT_DATE)
        End If
    Next A
    ' ------------------------------HORA DE FIN DE TAREAS
    fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.HORFIN_) = HORADEFIN_
    ' ------------------------------FECHA DE FIN DE TAREAS
    fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.FCHFIN_) = FECHADEFIN_
    
    Agregando = False
    Set xRs = Nothing
    Set RECORDSET_ = Nothing
End Function

Private Sub cmd_Click(Index As Integer)
    Dim A As Integer
    Dim num As Integer
    Dim Rpta As Integer
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim MENSAJE_ As String
    Dim nSQLId As String
    Dim nSQLId2 As String
    Dim NUMEROMAXTRAB_ As Integer
    Dim NUMREGAAGREGAR_ As Integer
    
    If QueHace = 3 Then Exit Sub
            
    Select Case Index
        Case 0 ' AGREGAR ITEM
            ReDim xCampos(2, 4) As String
            'descripcion                     'campo                       'tamaño                         'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "desitem":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Uni. Med.":     xCampos(1, 1) = "desunimed":     xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
            
            cSQL = "SELECT pro_receta.iditem, alm_inventario.descripcion AS desitem, pro_receta.id AS idrec, pro_receta.codrec, pro_receta.idunimed, mae_unidades.abrev AS desunimed, pro_linea.id AS idlinea, pro_linea.descripcion AS deslinea " _
                + vbCr + "FROM ((((pro_receta LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id) LEFT JOIN pro_tiptrab ON pro_receta.idtiptrab = pro_tiptrab.id) LEFT JOIN pro_formapag ON pro_receta.idformapag = pro_formapag.id) LEFT JOIN pro_linea ON pro_receta.id = pro_linea.idrec " _
                + vbCr + "WHERE (((pro_linea.id) Is Not Null) AND ((pro_receta.prirec)=1) AND ((alm_inventario.activo)=-1));"
                
            nTitulo = "Buscando Ítems"
            
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "desitem", "desitem", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            fg(0).Rows = fg(0).FixedRows
            fg(1).Rows = fg(1).FixedRows
            fg(2).Rows = fg(2).FixedRows
            fg(3).Rows = fg(3).FixedRows
            
            lblItem.Caption = NulosC(xRs("desitem"))
            txtIdItem.Text = NulosN(xRs("iditem"))
            
            fg(0).Rows = fg(0).Rows + 1
            ' ----------------------------SE LLENA RECETA
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACABECERA_.IDRECETA_) = NulosN(xRs("idrec"))
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACABECERA_.RECETA_) = NulosC(xRs("codrec"))
            ' ----------------------------SE LLENA UNIDADES DE MEDIDA
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACABECERA_.IDUNIMED_) = NulosN(xRs("idunimed"))
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACABECERA_.UNIMED_) = NulosC(xRs("desunimed"))
            ' ----------------------------SE LLENA LA LINEA
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACABECERA_.IDLINEA_) = NulosN(xRs("idlinea"))
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACABECERA_.LINEA_) = NulosC(xRs("deslinea"))
            ' ----------------------------EFICIENCIA
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACABECERA_.EFICIENCIA_) = 100
            
            fg(0).Col = COLUMNACABECERA_.CANTIDAD_
            fg(0).SetFocus
            
        Case 1 ' TIPO DE DOCUMENTO DE REFERENCIA
            ReDim xCampos(2, 4) As String
            
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
                  
            nTitulo = "Buscando Tipo de Documento de Referencia"
            
            cSQL = "SELECT mae_documento.* FROM mae_documento WHERE (id In (115))"
            
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            TxtIdTipDocRef.Text = NulosN(xRs("id"))
            LblTipDocRef.Caption = UCase(NulosC(xRs("descripcion")))
            txtNumDocRef.SetFocus
            
        Case 2 ' NUMERO DE DOC REFERENCIA
            ReDim xCampos(6, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Fch.Doc.":     xCampos(0, 1) = "fchpro":          xCampos(0, 2) = "900":          xCampos(0, 3) = "C"
            xCampos(1, 0) = "Num.Doc":      xCampos(1, 1) = "numdoc":          xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Ítem":         xCampos(2, 1) = "desitem":         xCampos(2, 2) = "3200":         xCampos(2, 3) = "C"
            xCampos(3, 0) = "Receta":       xCampos(3, 1) = "codrec":          xCampos(3, 2) = "1000":         xCampos(3, 3) = "C"
            xCampos(4, 0) = "Cantidad":     xCampos(4, 1) = "cantidad":        xCampos(4, 2) = "900":         xCampos(4, 3) = "N"
            xCampos(5, 0) = "Hor.Ini.":     xCampos(5, 1) = "horini":          xCampos(5, 2) = "800":         xCampos(5, 3) = "C"
                  
            nTitulo = "Buscando Tipo de Documento de Referencia"
            cSQL = "SELECT pro_ordenprod.id, Format([pro_ordenprod].[fchpro],'dd/mm/yy') AS fchpro, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numdoc, alm_inventario.descripcion AS desitem, pro_receta.codrec, pro_ordenprod.cantidad, Format([pro_ordenprod].[horini],'Short Time') AS horini, Format([pro_ordenprod].[horfin],'Short Time') AS horfin, pro_ordenprod.estado, UCase([mae_estados].[descripcion]) AS desestado " _
                + vbCr + "FROM ((pro_ordenprod LEFT JOIN pro_receta ON pro_ordenprod.idrec = pro_receta.id) LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) LEFT JOIN mae_estados ON pro_ordenprod.estado = mae_estados.id " _
                + vbCr + "WHERE (((pro_ordenprod.estado) = " & ESTADOPROCESADO_ & ") And ((pro_ordenprod.ano) = " & AnoTra & ") And ((pro_ordenprod.idmes) in (" & mMesActivo & ", " & mMesActivo - 1 & ")));"
            
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "fchpro", "fchpro", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            lbliddocref.Caption = NulosN(xRs("id"))
            txtNumDocRef.Text = NulosC(xRs("numdoc"))
            
            TxtIdResp.SetFocus
            
        Case 3 ' RESPONSABLE
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "apenom":     xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":               xCampos(1, 1) = "id":         xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
            
            cSQL = "SELECT pla_empleados.nombre AS apenom, pla_empleados.id " _
                + vbCr + "FROM pla_empleados;"
            
            nTitulo = "Buscando Responsable"
                   
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "apenom", "apenom", Principio
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            TxtIdResp.Text = xRs("id")
            lblResponsable.Caption = xRs("apenom")
            txtIdItem.SetFocus
                    
        Case 4 ' PROCESAR
            procesarLineaProduccion
        
        Case 5 ' PROPIEDADES
            aplicarPropiedades False, True
            CentrarFrm frm(0)
            frm(0).ZOrder 0
            frm(0).Visible = True
            
        Case 6 ' AGREGAR PERSONAL
            ReDim xCampos(5, 4) As String
            
            xCampos(0, 0) = "DNI":                  xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Grupo":                xCampos(1, 1) = "grupo":       xCampos(1, 2) = "800":      xCampos(1, 3) = "N":    xCampos(1, 4) = "C"
            xCampos(2, 0) = "Apellidos y Nombres":  xCampos(2, 1) = "nombre":      xCampos(2, 2) = "3250":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
            xCampos(3, 0) = "Area":                 xCampos(3, 1) = "area":        xCampos(3, 2) = "1750":     xCampos(3, 3) = "C":    xCampos(3, 4) = "C"
            xCampos(4, 0) = "Fch. Ing.":            xCampos(4, 1) = "fching":      xCampos(4, 2) = "1000":     xCampos(4, 3) = "C":    xCampos(4, 4) = "C"
            
            NUMEROMAXTRAB_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.NUMOPE_))
            
            If PROPIEDADES_.LIMITARNUMEROPERSONAL_ Then
                If fg(2).Rows - 1 >= NUMEROMAXTRAB_ Then
                    MsgBox "La Orden de Producción actual no puede admitir mas personal", _
                            vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Exit Sub
                End If
            End If
                
            ' generar la lista de personal para no considerar en la lista
            nSQLId = GENERAR_SQL_ID(fg(2), COLUMNADETALLEPERS_.IDPER_, " AND pla_empleados.id", "NOT IN", True)
            nSQLId2 = GENERAR_SQL_ID(fg(1), COLUMNADETALLETAREA_.IDAREA_, " AND pla_empleados.idarea", "IN", True)

            If PROPIEDADES_.LIMITARSELPERSONAL_ Then
                ' generar la consulta
                cSQL = "SELECT pla_empleados.numdoc, pro_grupo.num AS grupo, pla_empleados.id AS idemp, pla_empleados.nombre, -1 AS activo, mae_area.descripcion AS area, pla_empleados.fching " _
                    + vbCr + "FROM (((pla_empleados LEFT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id) LEFT JOIN (pro_grupodet LEFT JOIN pro_grupo ON pro_grupodet.idgrupo = pro_grupo.id) ON pro_emp.id = pro_grupodet.idper " _
                    + vbCr + "WHERE (((pla_empleados.fchcese) Is Null) And ((pro_empdet.idfun) = 6)) " & nSQLId & nSQLId2 _
                    + vbCr + "ORDER BY pla_empleados.nombre;"
            Else
                ' generar la consulta
                cSQL = "SELECT pla_empleados.numdoc, pro_grupo.num AS grupo, pla_empleados.id AS idemp, pla_empleados.nombre, -1 AS activo, mae_area.descripcion AS area, pla_empleados.fching " _
                    + vbCr + "FROM (((pla_empleados LEFT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id) LEFT JOIN (pro_grupodet LEFT JOIN pro_grupo ON pro_grupodet.idgrupo = pro_grupo.id) ON pro_emp.id = pro_grupodet.idper " _
                    + vbCr + "WHERE (((pla_empleados.fchcese) Is Null) And ((pro_empdet.idfun) = 6)) " & nSQLId _
                    + vbCr + "ORDER BY pla_empleados.nombre;"
            End If

            nTitulo = "Buscando Personal"
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
                        
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub

            Agregando = True
            fg(2).Rows = fg(2).Rows + 1
            fg(2).TextMatrix(fg(2).Rows - 1, COLUMNADETALLEPERS_.DNI_) = NulosC(xRs("numdoc"))
            fg(2).TextMatrix(fg(2).Rows - 1, COLUMNADETALLEPERS_.NOMBRE_) = NulosC(xRs("nombre"))
            fg(2).TextMatrix(fg(2).Rows - 1, COLUMNADETALLEPERS_.IDPER_) = NulosN(xRs("idemp"))
            Agregando = False
        
        Case 7 ' SELECCIONAR PERSONAL
            ReDim xCampos(5, 4) As String
            
            NUMEROMAXTRAB_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.NUMOPE_))
            If PROPIEDADES_.LIMITARNUMEROPERSONAL_ Then
                NUMREGAAGREGAR_ = NUMEROMAXTRAB_ - (fg(2).Rows - 1)
                If NUMREGAAGREGAR_ <= 0 Then
                    MsgBox "La Linea de Producción actual no puede admitir mas personal", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Exit Sub
                End If
            End If
                        
            xCampos(0, 0) = "DNI":                  xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Grupo":                xCampos(1, 1) = "grupo":       xCampos(1, 2) = "800":      xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            xCampos(2, 0) = "Apellidos y Nombres":  xCampos(2, 1) = "nombre":      xCampos(2, 2) = "3250":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
            xCampos(3, 0) = "Area":                 xCampos(3, 1) = "area":        xCampos(3, 2) = "1750":     xCampos(3, 3) = "C":    xCampos(3, 4) = "C"
            xCampos(4, 0) = "Fch. Ing.":            xCampos(4, 1) = "fching":      xCampos(4, 2) = "1000":     xCampos(4, 3) = "D":    xCampos(4, 4) = "C"
                     
            ' generar la lista de personal para no considerar en la lista
            nSQLId = GENERAR_SQL_ID(fg(2), COLUMNADETALLEPERS_.IDPER_, " AND pla_empleados.id", "NOT IN", True)
            nSQLId2 = GENERAR_SQL_ID(fg(1), COLUMNADETALLETAREA_.IDAREA_, " AND pla_empleados.idarea", "IN", True)
            
            If PROPIEDADES_.LIMITARSELPERSONAL_ Then
                ' generar la consulta
                cSQL = "SELECT 0 AS xsel, pla_empleados.numdoc, pro_grupo.num AS grupo, pla_empleados.id AS idemp, pla_empleados.nombre, -1 AS activo, mae_area.descripcion AS area, pla_empleados.fching " _
                    + vbCr + "FROM (((pla_empleados LEFT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id) LEFT JOIN (pro_grupodet LEFT JOIN pro_grupo ON pro_grupodet.idgrupo = pro_grupo.id) ON pro_emp.id = pro_grupodet.idper " _
                    + vbCr + "WHERE (((pla_empleados.fchcese) Is Null) And ((pro_empdet.idfun) = 6)) " & nSQLId & nSQLId2 _
                    + vbCr + "ORDER BY pla_empleados.nombre;"
            Else
                ' generar la consulta
                cSQL = "SELECT 0 AS xsel, pla_empleados.numdoc, pro_grupo.num AS grupo, pla_empleados.id AS idemp, pla_empleados.nombre, -1 AS activo, mae_area.descripcion AS area, pla_empleados.fching " _
                    + vbCr + "FROM (((pla_empleados LEFT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id) LEFT JOIN (pro_grupodet LEFT JOIN pro_grupo ON pro_grupodet.idgrupo = pro_grupo.id) ON pro_emp.id = pro_grupodet.idper " _
                    + vbCr + "WHERE (((pla_empleados.fchcese) Is Null) And ((pro_empdet.idfun) = 6)) " & nSQLId _
                    + vbCr + "ORDER BY pla_empleados.nombre;"
            End If
                        
            xform.SqlCad = cSQL
            xform.Titulo = "Buscando Personal"
            Set xform.Coneccion = xCon
            Set xRs = Nothing
            Set xRs = xform.Seleccionar(xCampos)
            Set xform = Nothing
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
                          
            If Not PROPIEDADES_.LIMITARNUMEROPERSONAL_ Then NUMREGAAGREGAR_ = xRs.RecordCount
            Agregando = True
            For A = 1 To NUMREGAAGREGAR_
                fg(2).Rows = fg(2).Rows + 1
                fg(2).TextMatrix(fg(2).Rows - 1, COLUMNADETALLEPERS_.DNI_) = NulosC(xRs("numdoc"))
                fg(2).TextMatrix(fg(2).Rows - 1, COLUMNADETALLEPERS_.NOMBRE_) = NulosC(xRs("nombre"))
                fg(2).TextMatrix(fg(2).Rows - 1, COLUMNADETALLEPERS_.IDPER_) = NulosN(xRs("idemp"))
                xRs.MoveNext
                If xRs.EOF = True Then Exit For
            Next A
            Agregando = False
            
        Case 8 ' RANKING PERSONAL
            
        Case 9 ' GRUPO PERSONAL
            
        Case 10 ' ELIMINAR PERSONAL
            If fg(2).Rows <= 0 Then Exit Sub
            
            If NulosN(cbEstado.ItemData(cbEstado.ListIndex)) > ESTADOPENDIENTE_ Then
                MsgBox "El registro esta en un estado en el que no se puede modificar", vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
                Exit Sub
            End If
            
            Rpta = MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            
            If Rpta = vbYes Then fg(2).RemoveItem fg(2).Row
            
        Case 11 ' ELIMINAR TODOS PERSONAL
            If fg(2).Rows <= 0 Then Exit Sub
            
            If NulosN(cbEstado.ItemData(cbEstado.ListIndex)) > ESTADOPENDIENTE_ Then
                MsgBox "El registro esta en un estado en el que no se puede modificar", vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
                Exit Sub
            End If
            
            Rpta = MsgBox("¿Está seguro de eliminar todos los registros?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
        
            If Rpta = vbYes Then fg(2).Rows = fg(2).FixedRows
        
        Case 12 ' AGREGAR REPROCESO
            ReDim xCampos(4, 4) As String
            
            xCampos(0, 0) = "Lote":         xCampos(0, 1) = "deslote":      xCampos(0, 2) = "1500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Almacen":      xCampos(1, 1) = "desalm":       xCampos(1, 2) = "2500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            xCampos(2, 0) = "fch. Ing.":    xCampos(2, 1) = "fching":       xCampos(2, 2) = "1000":     xCampos(2, 3) = "D":    xCampos(2, 4) = "D"
            xCampos(3, 0) = "Stock":        xCampos(3, 1) = "cantidad":     xCampos(3, 2) = "1000":     xCampos(3, 3) = "N":    xCampos(3, 4) = "C"
                        
            ' generar la lista de personal para no considerar en la lista
            nSQLId = GENERAR_SQL_ID(fg(3), COLUMNADETALLEREPROC_.IDLOTEDET_, " AND alm_inventariolotedet.id", "NOT IN", True)
            
            cSQL = "SELECT alm_inventariolotedet.idlote, alm_inventariolotedet.id AS idlotedet, alm_inventariolote.fching, alm_inventariolote.descripcion AS deslote, alm_inventariolotedet.cantidad, alm_inventariolotedet.idalm, alm_almacenes.descripcion AS desalm " _
                + vbCr + "FROM (alm_inventariolote LEFT JOIN alm_inventariolotedet ON alm_inventariolote.id = alm_inventariolotedet.id) LEFT JOIN alm_almacenes ON alm_inventariolotedet.idalm = alm_almacenes.id " _
                + vbCr + "WHERE (((alm_inventariolotedet.cantidad)>0) AND ((alm_almacenes.tipo)=2) AND ((alm_inventariolote.iditem)=" & NulosN(txtIdItem.Text) & ")) " & nSQLId _
                        
            nTitulo = "Buscando Reprocesos"
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "fching", "fching", Principio
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            Agregando = True
            fg(3).Rows = fg(3).Rows + 1
            fg(3).TextMatrix(fg(3).Rows - 1, COLUMNADETALLEREPROC_.LOTE_) = NulosC(xRs("deslote"))
            fg(3).TextMatrix(fg(3).Rows - 1, COLUMNADETALLEREPROC_.ALMACEN_) = NulosC(xRs("desalm"))
            fg(3).TextMatrix(fg(3).Rows - 1, COLUMNADETALLEREPROC_.CANTIDAD_) = NulosN(xRs("cantidad"))
            fg(3).TextMatrix(fg(3).Rows - 1, COLUMNADETALLEREPROC_.IDLOTE_) = NulosN(xRs("idlote"))
            fg(3).TextMatrix(fg(3).Rows - 1, COLUMNADETALLEREPROC_.IDLOTEDET_) = NulosN(xRs("idlotedet"))
            fg(3).TextMatrix(fg(3).Rows - 1, COLUMNADETALLEREPROC_.IDALM_) = NulosN(xRs("idalm"))
            Agregando = False
            
        Case 13 ' SELECCIONAR REPROCESO
            ReDim xCampos(4, 4) As String
            
            xCampos(0, 0) = "Lote":         xCampos(0, 1) = "deslote":      xCampos(0, 2) = "1500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Almacen":      xCampos(1, 1) = "desalm":       xCampos(1, 2) = "2500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            xCampos(2, 0) = "fch. Ing.":    xCampos(2, 1) = "fching":       xCampos(2, 2) = "1000":     xCampos(2, 3) = "D":    xCampos(2, 4) = "D"
            xCampos(3, 0) = "Stock":        xCampos(3, 1) = "cantidad":     xCampos(3, 2) = "1000":     xCampos(3, 3) = "N":    xCampos(3, 4) = "C"
                        
            ' generar la lista de personal para no considerar en la lista
            nSQLId = GENERAR_SQL_ID(fg(3), COLUMNADETALLEREPROC_.IDLOTEDET_, " AND alm_inventariolotedet.id", "NOT IN", True)
            
            cSQL = "SELECT 0 AS xsel, alm_inventariolotedet.idlote, alm_inventariolotedet.id AS idlotedet, alm_inventariolote.fching, alm_inventariolote.descripcion AS deslote, alm_inventariolotedet.cantidad, alm_inventariolotedet.idalm, alm_almacenes.descripcion AS desalm " _
                + vbCr + "FROM (alm_inventariolote LEFT JOIN alm_inventariolotedet ON alm_inventariolote.id = alm_inventariolotedet.id) LEFT JOIN alm_almacenes ON alm_inventariolotedet.idalm = alm_almacenes.id " _
                + vbCr + "WHERE (((alm_inventariolotedet.cantidad)>0) AND ((alm_almacenes.tipo)=2) AND ((alm_inventariolote.iditem)=" & NulosN(txtIdItem.Text) & ")) " & nSQLId _
                            
            xform.SqlCad = cSQL
            xform.Titulo = "Buscando Reprocesos"
            Set xform.Coneccion = xCon
            Set xRs = Nothing
            Set xRs = xform.Seleccionar(xCampos)
            Set xform = Nothing
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
                    
            Agregando = True
            xRs.MoveFirst
            While Not xRs.EOF
                fg(3).Rows = fg(3).Rows + 1
                fg(3).TextMatrix(fg(3).Rows - 1, COLUMNADETALLEREPROC_.LOTE_) = NulosC(xRs("deslote"))
                fg(3).TextMatrix(fg(3).Rows - 1, COLUMNADETALLEREPROC_.ALMACEN_) = NulosC(xRs("desalm"))
                fg(3).TextMatrix(fg(3).Rows - 1, COLUMNADETALLEREPROC_.CANTIDAD_) = NulosN(xRs("cantidad"))
                fg(3).TextMatrix(fg(3).Rows - 1, COLUMNADETALLEREPROC_.IDLOTE_) = NulosN(xRs("idlote"))
                fg(3).TextMatrix(fg(3).Rows - 1, COLUMNADETALLEREPROC_.IDLOTEDET_) = NulosN(xRs("idlotedet"))
                fg(3).TextMatrix(fg(3).Rows - 1, COLUMNADETALLEREPROC_.IDALM_) = NulosN(xRs("idalm"))
                
                xRs.MoveNext
            Wend
            Agregando = False
            
        Case 14 ' ELIMINAR REPROCESO
            If fg(3).Rows <= 0 Then Exit Sub
            
            If NulosN(cbEstado.ItemData(cbEstado.ListIndex)) > ESTADOPENDIENTE_ Then
                MsgBox "El registro esta en un estado en el que no se puede modificar", vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
                Exit Sub
            End If
            
            Rpta = MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            
            If Rpta = vbYes Then fg(3).RemoveItem fg(3).Row
            
        Case 15 ' ELIMINAR TODOS REPROCESO
            If fg(3).Rows <= 0 Then Exit Sub
            
            If NulosN(cbEstado.ItemData(cbEstado.ListIndex)) > ESTADOPENDIENTE_ Then
                MsgBox "El registro esta en un estado en el que no se puede modificar", vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
                Exit Sub
            End If
            
            Rpta = MsgBox("¿Está seguro de eliminar todos los registros?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
        
            If Rpta = vbYes Then fg(3).Rows = fg(3).FixedRows
            
        Case 16 ' ACEPTAR OPCIONES PROCESADO
            aplicarPropiedades True
            frm(0).Visible = False
            
        Case 17 ' CANCELAR OPCIONES PROCESADO
            frm(0).Visible = False
        
    End Select
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstOrdProd("id")), xCon
    End If
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstOrdProd
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDETE LAS COLUMNAS DEL CONTROL Dg3
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstOrdProd.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub fg_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If NulosN(cbEstado.ItemData(cbEstado.ListIndex)) > ESTADOPENDIENTE_ Then Cancel = True
    
    Select Case Index
        Case 0
        
        Case 1 ' -----------------------GRID DE TAREAS
            Select Case Col
                Case COLUMNADETALLETAREA_.TAREA_ To COLUMNADETALLETAREA_.FCHFIN_
                    Cancel = True
                
            End Select
    End Select
End Sub

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    Dim xCampos() As String
    Dim TIPOPRODUCTO_ As Integer
    Dim IDITEM_ As Integer
    Dim IDTAR_ As Integer
    Dim nSQLId As String
    Dim nSQLId2 As String
    Dim Rpta As Integer
    
    If QueHace = 3 Then Exit Sub
    
    With fg(Index)
        Select Case Index
            Case 0 ' --------------------------------------GRID PRINCIPAL
                Select Case Col
                    Case COLUMNACABECERA_.RECETA_
                        ReDim xCampos(2, 4) As String
                        
                        IDITEM_ = NulosN(txtIdItem.Text)
                        
                        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
                        xCampos(1, 0) = "Receta":     xCampos(1, 1) = "codrec":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
                        
                        cSQL = "SELECT pro_receta.codrec, pro_receta.descripcion, pro_receta.prirec, pro_receta.id " _
                            + vbCr + "FROM pro_receta " _
                            + vbCr + "WHERE (((pro_receta.iditem) = " & IDITEM_ & ")) " _
                            + vbCr + "ORDER BY pro_receta.prirec;"
                            
                        nTitulo = "Buscando Recetas del Producto"
                           
                        Set xRs = Nothing
                        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
                        
                        If xRs.State = 0 Then Exit Sub
                        If xRs.RecordCount = 0 Then Exit Sub
                        
                        .TextMatrix(.Row, COLUMNACABECERA_.RECETA_) = NulosC(xRs("codrec"))
                        .TextMatrix(.Row, COLUMNACABECERA_.IDRECETA_) = NulosN(xRs("id"))
                        
                        ' ---------------------------------LINEA DE PRODUCCION
                        cSQL = "SELECT pro_linea.id AS idlineadet, pro_linea.descripcion " _
                                + vbCr + "From pro_linea " _
                                + vbCr + "WHERE (((pro_linea.idrec)=" & .TextMatrix(.Row, COLUMNACABECERA_.IDRECETA_) & ") AND ((pro_linea.activo)=-1));"
                           
                        Set xRs = Nothing
                        RST_Busq xRs, cSQL, xCon
                        
                        If xRs.State = 0 Then
                            MsgBox "Ha ocurrido un error verificar la Linea de Producción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                            Exit Sub
                        End If
                        
                        If xRs.RecordCount = 0 Then
                            MsgBox "El producto procesado no tiene Linea activa, procese una para calcular tiempos de Producción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                            Exit Sub
                        End If
                        
                        .TextMatrix(.Row, COLUMNACABECERA_.LINEA_) = NulosC(xRs("descripcion"))
                        .TextMatrix(.Row, COLUMNACABECERA_.IDLINEA_) = NulosN(xRs("idlineadet"))
                                                
                    Case COLUMNACABECERA_.UNIMED_
                        ReDim xCampos(2, 4) As String
                        
                        ' SE VERIFICA SI HAY RECETA
                        If NulosN(.TextMatrix(.Row, COLUMNACABECERA_.IDRECETA_)) = 0 Then
                            MsgBox "Agregue un ítem o una receta en su defecto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                            .Col = COLUMNACABECERA_.RECETA_
                            Exit Sub
                        End If
                        
                        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
                        xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "2500":         xCampos(0, 3) = "C"
                        xCampos(1, 0) = "Abrev.":           xCampos(1, 1) = "abrev":            xCampos(1, 2) = "1000":         xCampos(1, 3) = "D"
                                
                        nTitulo = "Buscando Unidades"
        
                        cSQL = "SELECT * " _
                            + vbCr + "FROM mae_unidades;"
                        
                        Set xRs = Nothing
                        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                        "descripcion", "descripcion", Principio, ""
                        
                        If xRs.State = 0 Then Exit Sub
                        If xRs.RecordCount = 0 Then Exit Sub
                        
                        .TextMatrix(.Row, COLUMNACABECERA_.IDUNIMED_) = NulosN(xRs("id"))
                        .TextMatrix(.Row, COLUMNACABECERA_.UNIMED_) = NulosC(xRs("abrev"))
                                                
                    Case COLUMNACABECERA_.LINEA_
                        ReDim xCampos(3, 4) As String
                        Dim IDREC_ As Double
                        
                        IDREC_ = NulosN(.TextMatrix(.Row, COLUMNACABECERA_.IDRECETA_))
                        
                        'descripcion                            'campo                          'tamaño                        'tipo = Numerico, caracter, fecha
                        xCampos(0, 0) = "Descripcion":          xCampos(0, 1) = "descline":     xCampos(0, 2) = "4000":        xCampos(0, 3) = "C"
                        xCampos(1, 0) = "Operarios":            xCampos(1, 1) = "numop":        xCampos(1, 2) = "1000":        xCampos(1, 3) = "N"
                        xCampos(2, 0) = "Eficiencia (%)":       xCampos(2, 1) = "efic":         xCampos(2, 2) = "1250":        xCampos(2, 3) = "N"
                     
                        cSQL = "SELECT pro_linea.descripcion AS descline, pro_linea.numop, pro_linea.efic, pro_linea.idlinea, pro_linea.id AS idlineadet " _
                            + vbCr + "From pro_linea " _
                            + vbCr + "WHERE (((pro_linea.idrec)=" & IDREC_ & "));"
                    
                        nTitulo = "Buscando Linea"
                        Set xRs = Nothing
                        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descline", "descline", Principio
                    
                        If xRs.State = 0 Then Exit Sub
                        If xRs.RecordCount = 0 Then
                            MsgBox "El producto procesado no tiene Linea activa, procese una para calcular tiempos de Producción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                            Exit Sub
                        End If
                        
                        If fg(1).Rows > fg(1).FixedRows Or _
                                    fg(2).Rows > fg(2).FixedRows Or fg(3).Rows > fg(3).FixedRows Then
                            Rpta = MsgBox("¿Se Eliminará todo el Personal y Tareas Relacionado a la linea Anterior; desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
                            If Rpta = vbNo Then Exit Sub
                        End If
                        
                        fg(1).Rows = fg(1).FixedRows
                        fg(2).Rows = fg(2).FixedRows
                        fg(3).Rows = fg(3).FixedRows
                        
                        .TextMatrix(.Row, COLUMNACABECERA_.LINEA_) = NulosC(xRs("descline"))
                        .TextMatrix(.Row, COLUMNACABECERA_.IDLINEA_) = NulosN(xRs("idlineadet"))
                        
                        
                End Select
            
            Case 1 ' -------------------------------------DRID DE TAREAS
                Select Case Col
                    Case COLUMNADETALLETAREA_.AREA_
                        IDTAR_ = NulosN(fg(1).TextMatrix(fg(1).Row, COLUMNADETALLETAREA_.IDTAR_))
                        
                        ReDim xCampos(2, 4) As String
                        xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
                        xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
                        
                        nTitulo = "Buscando Area"
                        
                        cSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea, pro_emp.id AS idper, pla_empleados.id AS idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado, pla_empleados.numdoc " _
                            + vbCr + "FROM (((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) RIGHT JOIN pro_area ON pro_emp.id = pro_area.idper) INNER JOIN mae_area ON pro_area.idarea = mae_area.id) LEFT JOIN pro_areadet ON pro_area.id = pro_areadet.idar " _
                            + vbCr + "WHERE (((pro_areadet.idtar)=" & IDTAR_ & ")); "
                        
                        Set xRs = Nothing
                        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                        "nombre", "nombre", Principio, ""
                                                                      
                        If xRs.State = 0 Then Exit Sub
                        If xRs.RecordCount = 0 Then Exit Sub
                        ' -------------------------AREA
                        fg(1).TextMatrix(fg(1).Row, COLUMNADETALLETAREA_.AREA_) = UCase(NulosC(xRs("nombre")))
                        fg(1).TextMatrix(fg(1).Row, COLUMNADETALLETAREA_.IDAREA_) = NulosN(xRs("id"))
                        ' -------------------------RESPONSABLE
                        fg(1).TextMatrix(fg(1).Row, COLUMNADETALLETAREA_.RESPONSABLE_) = NulosC(xRs("encargado"))
                        fg(1).TextMatrix(fg(1).Row, COLUMNADETALLETAREA_.IDRESP_) = NulosC(xRs("idemp"))
                    
                    Case COLUMNADETALLETAREA_.RESPONSABLE_
                        ReDim xCampos(2, 4) As String
                        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
                        xCampos(0, 0) = "Apellidos y Nombres":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
                        xCampos(1, 0) = "Codigo":                xCampos(1, 1) = "id":        xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
                                
                        cSQL = "SELECT pro_emp.idemp, pro_emp.sup, pro_emp.prog, pro_emp.res, pla_empleados.nombre " _
                            + vbCr + "FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id " _
                            + vbCr + "WHERE (((pro_empdet.idfun) = 3)) " _
                            + vbCr + "GROUP BY pro_emp.idemp, pro_emp.sup, pro_emp.prog, pro_emp.res, pla_empleados.nombre " _
                            + vbCr + "HAVING (((pla_empleados.nombre) Is Not Null)) " _
                            + vbCr + "ORDER BY pla_empleados.nombre;"
                            
                        nTitulo = "Buscando Personal Responsable"
                         
                        Set xRs = Nothing
                        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
                        
                        If xRs.State = 0 Then Exit Sub
                        If xRs.RecordCount = 0 Then Exit Sub
                        
                        ' -------------------------RESPONSABLE
                        fg(1).TextMatrix(fg(1).Row, COLUMNADETALLETAREA_.RESPONSABLE_) = UCase(NulosC(xRs("nombre")))
                        fg(1).TextMatrix(fg(1).Row, COLUMNADETALLETAREA_.IDRESP_) = NulosC(xRs("idemp"))
                        
                End Select
        End Select
    End With
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If QueHace = 3 Then Exit Sub
    If Agregando = True Then Exit Sub
    
    Select Case Index
        Case 0 ' -------------------------GRID CABECERA
            Select Case Col
                Case COLUMNACABECERA_.CANTIDAD_
                    fg(0).TextMatrix(Row, Col) = Format(fg(0).TextMatrix(Row, Col), FORMAT_CANTIDAD)
                    pLlenarInsumos NulosN(fg(0).TextMatrix(Row, COLUMNACABECERA_.IDRECETA_)), NulosN(fg(0).TextMatrix(Row, COLUMNACABECERA_.CANTIDAD_))
                    
                Case COLUMNACABECERA_.HORINI_
                    fg(0).TextMatrix(Row, Col) = Format(fg(0).TextMatrix(Row, Col), FORMAT_HORA_SIN_SEGUNDO)
            End Select
        
        Case 1 ' -------------------------GRID TAREAS
            Select Case Col
                Case COLUMNADETALLETAREA_.NUMOP_
                    fg(0).TextMatrix(fg(0).Col, COLUMNACABECERA_.NUMOPE_) = GRID_SUMAR_COL(fg(1), COLUMNADETALLETAREA_.NUMOP_)
                    
            End Select
        
        Case 4
            Select Case Col
                Case COLUMNADETALLEINSUMOS_.CANTIDAD_
                    pRecalcularReceta NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.IDRECETA_)), _
                                        NulosN(fg(4).TextMatrix(fg(4).Row, COLUMNADETALLEINSUMOS_.IDINSUMO_)), _
                                        NulosN(fg(4).TextMatrix(fg(4).Row, COLUMNADETALLEINSUMOS_.CANTIDAD_))
            End Select
        
    End Select
End Sub

Private Sub pLlenarInsumos(IDRECETA_ As Integer, CANTIDAD_ As Double)
    Dim xRs As New ADODB.Recordset
    
    cSQL = "SELECT alm_inventario.tippro AS idtippro, pro_recetains.iditem, [pro_recetains]![canpro]*" & CANTIDAD_ & " AS cantidad, pro_recetains.idunimed " _
        + vbCr + "FROM pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((pro_recetains.idrec)=" & IDRECETA_ & "));"
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    fg(4).Rows = fg(4).FixedRows
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    xRs.MoveFirst
    While Not xRs.EOF
        fg(4).Rows = fg(4).Rows + 1
        fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.IDINSUMO_) = NulosN(xRs("iditem"))
        fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.CANTIDAD_) = NulosN(xRs("cantidad"))
        fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.CANTIDAD_) = Format(fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.CANTIDAD_), FORMAT_CANTIDAD)
        fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.IDUNIMED_) = NulosN(xRs("idunimed"))
        fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.INSUMO_) = Busca_Codigo(NulosN(xRs("iditem")), "id", "descripcion", "alm_inventario", "N", xCon)
        fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.UNIMED_) = Busca_Codigo(NulosN(xRs("idunimed")), "id", "abrev", "mae_unidades", "N", xCon)
        xRs.MoveNext
    Wend
End Sub

Private Sub cbEstado_DropDown()
    If Agregando Then Exit Sub
    ESTADOANTERIOR_ = cbEstado.ItemData(cbEstado.ListIndex)
End Sub

Private Sub cbEstado_Click()
    Dim Rpta As Integer
    Dim IDORD_ As Double
    Dim MENSAJE_ As String
    Dim xRs As New ADODB.Recordset
    Dim RSTSOL_ As New ADODB.Recordset
    Dim IDRECETA_ As Integer
    Dim CANTIDAD_ As Double
    
    Dim IDSOL_ As Integer
    Dim FCHSOL_ As String
    Dim NUMSER_ As String
    Dim NUMERODOCUMENTO_ As Integer
    Dim NUMDOC_ As String
    Dim IDRESP_ As Integer
    Dim IDTIPDOCREF_ As Integer
    Dim IDDOCREF_ As Integer
    Dim IDITEM_ As Integer
    Dim IDALM_ As Integer
    Dim IDESTADO_ As Integer
    Dim A As Integer

    If Agregando Then Exit Sub
    If QueHace = 3 Then Exit Sub

    IDORD_ = NulosN(RstOrdProd("id"))

    Select Case cbEstado.ItemData(cbEstado.ListIndex)
        Case ESTADOPENDIENTE_ ' Pendiente
            If ESTADOANTERIOR_ > ESTADOPENDIENTE_ Then
                MsgBox "No se puede cambiar el estado a " & cbEstado.Text, vbInformation, xTitulo
                llenarEstado 0, 1, , cbEstado, ESTADOANTERIOR_
            End If
            Exit Sub

        Case ESTADOPROCESADO_ ' Procesado
            If ESTADOANTERIOR_ < ESTADOPROCESADO_ Then
                Rpta = MsgBox("Cambiar el estado a " & cbEstado.Text & " bloqueara el registro para su posterior modificación " _
                                    + vbCr + "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)

                If Rpta = vbYes Then
                    Rpta = MsgBox("?Desea generar la solicitud de materiales?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
                    If Rpta = vbNo Then Exit Sub
                    
                    ' -------------------------------------SE DEFINE EL RECORDSET
                    cSQL = "SELECT  TOP 1 * " _
                        + vbCr + "FROM pro_solicitudmatdet;"
                    Set xRs = Nothing
                    RST_Busq xRs, cSQL, xCon
                    DEFINIR_RST_TMP RSTSOL_, xRs
                    ' -------------------------------------SE BUSCAN LOS INSUMOS
                    IDRECETA_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.IDRECETA_))
                    CANTIDAD_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.CANTIDAD_))
                    
                    cSQL = "SELECT alm_inventario.tippro AS idtippro, pro_recetains.iditem, [pro_recetains]![canpro]*" & CANTIDAD_ & " AS cantidad, pro_recetains.idunimed " _
                        + vbCr + "FROM pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id " _
                        + vbCr + "WHERE (((pro_recetains.idrec)=" & IDRECETA_ & "));"
                    Set xRs = Nothing
                    RST_Busq xRs, cSQL, xCon
                    
                    If xRs.State = 0 Then Exit Sub
                    If xRs.RecordCount = 0 Then Exit Sub
                    FCHSOL_ = TxtFchPro.Valor
                    NUMSER_ = NulosC(TxtNumSer.Text)
                    NUMERODOCUMENTO_ = HallaCodigoTabla("pro_solicitudmat", xCon, "numdoc")
                    NUMDOC_ = Format(NUMERODOCUMENTO_, "0000000000")
                    IDRESP_ = NulosN(TxtIdResp.Text)
                    IDTIPDOCREF_ = 115
                    IDDOCREF_ = NulosN(RstOrdProd("id"))
                    IDESTADO_ = ESTADOPENDIENTE_
                    
                    ' ---------------SE LLENA LA MATERIA PRIMA
                    IDSOL_ = 0
                    IDALM_ = 1
                    xRs.Filter = "idtippro=1"
                    If xRs.RecordCount = 0 Then GoTo GRABARINSUMOS_
                    limpiarRST RSTSOL_
                    xRs.MoveFirst
                    While Not xRs.EOF
                        RSTSOL_.AddNew
                        RSTSOL_("iditem") = NulosN(xRs("iditem"))
                        RSTSOL_("cantidad") = NulosN(xRs("cantidad"))
                        RSTSOL_("idunimed") = NulosN(xRs("idunimed"))
                        RSTSOL_.Update
                        xRs.MoveNext
                    Wend
                    If grabarSolicitud(FCHSOL_, IDTIPDOCREF_, IDDOCREF_, IDRESP_, NUMDOC_, IDALM_, _
                                                    RSTSOL_, NUMSER_, IDSOL_, IDESTADO_, CInt(AnoTra), mMesActivo, 6) Then
                        NUMERODOCUMENTO_ = NUMERODOCUMENTO_ + 1
                    Else
                        MsgBox "Ocurrió un error al intentar grabar la solicitud de :" _
                        + vbCr + "Materia Prima", vbInformation, xTitulo
                    End If
GRABARINSUMOS_:
                    ' -----------------SE LLENA LOS INSUMOS
                    IDSOL_ = 0
                    IDALM_ = 6
                    NUMDOC_ = Format(NUMERODOCUMENTO_, "0000000000")
                    xRs.Filter = "idtippro=4"
                    If xRs.RecordCount = 0 Then GoTo GRABARINTERMEDIOS_
                    limpiarRST RSTSOL_
                    xRs.MoveFirst
                    While Not xRs.EOF
                        RSTSOL_.AddNew
                        RSTSOL_("iditem") = NulosN(xRs("iditem"))
                        RSTSOL_("cantidad") = NulosN(xRs("cantidad"))
                        RSTSOL_("idunimed") = NulosN(xRs("idunimed"))
                        RSTSOL_.Update
                        xRs.MoveNext
                    Wend
                    If grabarSolicitud(FCHSOL_, IDTIPDOCREF_, IDDOCREF_, IDRESP_, NUMDOC_, IDALM_, _
                                                    RSTSOL_, NUMSER_, IDSOL_, IDESTADO_, CInt(AnoTra), mMesActivo, 6) Then
                        NUMERODOCUMENTO_ = NUMERODOCUMENTO_ + 1
                    Else
                        MsgBox "Ocurrió un error al intentar grabar la solicitud de :" _
                        + vbCr + "Insumos", vbInformation, xTitulo
                    End If
GRABARINTERMEDIOS_:
                    ' ---------------SE LLENA LOS PRODUCTOS INTERMEDIOS
                    IDSOL_ = 0
                    IDALM_ = 2
                    NUMDOC_ = Format(NUMERODOCUMENTO_, "0000000000")
                    xRs.Filter = "idtippro=3"
                    If xRs.RecordCount = 0 Then GoTo SALIRGRABAR_
                    limpiarRST RSTSOL_
                    xRs.MoveFirst
                    While Not xRs.EOF
                        RSTSOL_.AddNew
                        RSTSOL_("iditem") = NulosN(xRs("iditem"))
                        RSTSOL_("cantidad") = NulosN(xRs("cantidad"))
                        RSTSOL_("idunimed") = NulosN(xRs("idunimed"))
                        RSTSOL_.Update
                        xRs.MoveNext
                    Wend
                    If grabarSolicitud(FCHSOL_, IDTIPDOCREF_, IDDOCREF_, IDRESP_, NUMDOC_, IDALM_, _
                                                    RSTSOL_, NUMSER_, IDSOL_, IDESTADO_, CInt(AnoTra), mMesActivo, 6) Then
                        NUMERODOCUMENTO_ = NUMERODOCUMENTO_ + 1
                    Else
                        MsgBox "Ocurrió un error al intentar grabar la solicitud de: " _
                        + vbCr + "P. Intermedios", vbInformation, xTitulo
                    End If
SALIRGRABAR_:
                    Grabar
                    RstOrdProd.Requery
                    Dg1.Refresh
                    RstOrdProd.Find "id=" & IDORD_
            
                Else
                    Agregando = True
                    llenarEstado 0, 1, , cbEstado, ESTADOANTERIOR_
                    Agregando = False
                End If
            Else
                MsgBox "No se puede pasar a un estado " & cbEstado.Text, vbInformation, xTitulo
                Agregando = True
                llenarEstado 0, 1, , cbEstado, ESTADOANTERIOR_
                Agregando = False
            End If
            Exit Sub

        Case ESTADOANULADO_ ' Anulada
            If ESTADOANTERIOR_ = ESTADOPROCESADO_ Then
                If Not verificarCambioEstado(NulosN(RstOrdProd("id")), MENSAJE_) Then
                    MsgBox "No se puede pasar a un estado " & cbEstado.Text, vbInformation, xTitulo
                    Agregando = True
                    llenarEstado 0, 1, , cbEstado, ESTADOANTERIOR_
                    Agregando = False
                Else
                    ' -------------SE CAMBIA DE ESTADO A LA SOLICITUD DE MATERIALES
                    cSQL = "UPDATE pro_solicitudmat SET pro_solicitudmat.estado = 2 " _
                        + vbCr + "WHERE (((pro_solicitudmat.idtipdocref)=115) AND ((pro_solicitudmat.iddocref)=" & NulosN(RstOrdProd("id")) & "));"
                    ' --------------EJECUTA COMANDO
                    xCon.Execute cSQL
                    ' --------------ACTUALIZA VAR_EDICION
                    cSQL = "SELECT pro_solicitudmat.id " _
                        + vbCr + "FROM pro_solicitudmat " _
                        + vbCr + "WHERE (((pro_solicitudmat.idtipdocref)=115) And ((pro_solicitudmat.iddocref)=" & NulosN(RstOrdProd("id")) & "))"
                    
                    Set xRs = Nothing
                    RST_Busq xRs, cSQL, xCon
                    If xRs.State = 0 Then Exit Sub
                    If xRs.RecordCount = 0 Then Exit Sub
                    xRs.MoveFirst
                    While Not xRs.EOF
                        GrabarOperacion xIdUsuario, 54, 7, xHorIni, Time, Date, xCon, NulosN(xRs("id"))
                        xRs.MoveNext
                    Wend
                End If
            Else
                MsgBox "No se puede cambiar el estado a " & cbEstado.Text, vbInformation, xTitulo
                Agregando = True
                llenarEstado 0, 1, , cbEstado, ESTADOANTERIOR_
                Agregando = False
            End If
            Exit Sub

    End Select
End Sub

Private Sub anular()
    Dim MENSAJE_ As String
    Dim xRs As New ADODB.Recordset
    
    If verificarCambioEstado(NulosN(RstOrdProd("id")), MENSAJE_) Then
        ' ----------------------------------------SE CAMBIA DE ESTADO A LA SOLICITUD DE MATERIALES
        cSQL = "UPDATE pro_solicitudmat SET pro_solicitudmat.estado = " & ESTADOANULADO_ & " " _
            + vbCr + "WHERE (((pro_solicitudmat.idtipdocref)=115) AND ((pro_solicitudmat.iddocref)=" & NulosN(RstOrdProd("id")) & "));"
        ' --------------EJECUTA COMANDO
        xCon.Execute cSQL
        ' --------------ACTUALIZA VAR_EDICION
        cSQL = "SELECT pro_solicitudmat.id " _
            + vbCr + "FROM pro_solicitudmat " _
            + vbCr + "WHERE (((pro_solicitudmat.idtipdocref)=115) And ((pro_solicitudmat.iddocref)=" & NulosN(RstOrdProd("ID")) & "))"
        
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        xRs.MoveFirst
        While Not xRs.EOF
            GrabarOperacion xIdUsuario, 54, 7, xHorIni, Time, Date, xCon, NulosN(xRs("id"))
            xRs.MoveNext
        Wend
        ' ----------------------------------------SE CAMBIA DE ESTADO AL REGISTRO
        xCon.Execute "UPDATE pro_ordenprod SET pro_ordenprod.estado = " & ESTADOANULADO_ & " WHERE (((pro_ordenprod.id) = " & NulosN(RstOrdProd("id")) & "))"
        MsgBox "El registro se anuló con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstOrdProd.Requery
        Dg1.Refresh
    Else
        MsgBox MENSAJE_, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Function verificarCambioEstado(IDORD_ As Integer, ByRef MENSAJE_ As String) As Boolean
    Dim xRs As New ADODB.Recordset
    
    ' -------------------------------------SOLICITUD DE MATERIALES
    cSQL = "SELECT * " _
        + vbCr + "FROM pro_solicitudmat " _
        + vbCr + "WHERE (((pro_solicitudmat.idtipdocref)=115) AND ((pro_solicitudmat.iddocref)=" & IDORD_ & ") AND ((pro_solicitudmat.estado)=" & ESTADOPROCESADO_ & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    MENSAJE_ = "Solicitud de Materiales"
    
    If xRs.State = 0 Then verificarCambioEstado = False: GoTo SALIR_
    If xRs.RecordCount > 0 Then verificarCambioEstado = False: GoTo SALIR_
    
    verificarCambioEstado = True
    Exit Function
    
SALIR_:
    MENSAJE_ = "Se han encontrado " & MENSAJE_ & " que se encuentran en un estado no modificable; " _
    & vbCr & "verifique la condición de dichos Registros para completar esta acción."
End Function

Private Function cambiarEstadoRelacionados(IDORDDET_ As Double, ESTADO_ As Double) As Boolean
    Dim ID_ As Double
    
    On Error GoTo ERROR_
    ' Salidas de Almacen
    cSQL = "UPDATE alm_ingreso SET alm_ingreso.estado = " & ESTADO_ & " " _
        + vbCr + "WHERE (((alm_ingreso.idorddet)=" & IDORDDET_ & "));"

    xCon.Execute cSQL
    
    ' GRABAMOS LOS MOVIMIENTOS
    ' INGRESOS Y SALIDAS DE ALMACEN
    ID_ = Busca_Codigo(IDORDDET_, "idorddet", "id", "alm_ingreso", "N", xCon)
    GrabarOperacion xIdUsuario, 8, 7, xHorIni, Time, Date, xCon, ID_
        
    cambiarEstadoRelacionados = True
    Exit Function
    
ERROR_:
    MsgBox "Ha ocurrido un error al tratar de cambiar de estado", vbInformation, xTitulo
    cambiarEstadoRelacionados = False
End Function

Private Sub Fg_EnterCell(Index As Integer)
    If QueHace = 3 Then
        fg(Index).Editable = flexEDNone
        fg(Index).SelectionMode = flexSelectionByRow
        Exit Sub
    End If
    fg(Index).Editable = flexEDKbdMouse
    fg(Index).SelectionMode = flexSelectionFree
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Index
        Case 0 ' --------------------GRID CABECERA
            Select Case Col
                Case COLUMNACABECERA_.RECETA_, COLUMNACABECERA_.UNIMED_, COLUMNACABECERA_.LINEA_ _
                                , COLUMNACABECERA_.HORFIN_, COLUMNACABECERA_.FCHFIN_, COLUMNACABECERA_.NUMOPE_
                    KeyAscii = 0
                    
                Case COLUMNACABECERA_.CANTIDAD_
                    If IsNumeric(KeyAscii) = False Then KeyAscii = 0
                    
            End Select
        
        Case 1 ' --------------------GRID DE TAREAS
            Select Case Col
                Case COLUMNADETALLETAREA_.AREA_, COLUMNADETALLETAREA_.RESPONSABLE_
                    KeyAscii = 0
                
            End Select
    End Select
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        If QueHace = 3 Then Exit Sub
        If Button = 2 Then
            PopupMenu menu1
        End If
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    TabOne1.CurrTab = 0
    SeEjecuto = False
    Agregando = False
    iniciarCampos
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDOS E CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        mMesActivo = xMes
            
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        pCargarGrid
    End If
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub

    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 4000 Then Me.Height = 4000

    ' Se dimensiona el Contenido
    TabOne1.Width = Me.Width - 150
    TabOne1.Height = Me.Height - 750
    
    Label4(0).Width = Me.Width - 100
    LblMes.Left = TabOne1.Width - 1200
    Dg1.Width = TabOne1.Width - 100
    Dg1.Height = TabOne1.Height - 1000
    
    ' Se dimensiona el Detalle
    Label5.Width = Me.Width - 100
    
    Frame4.Width = TabOne1.Width - 150
    Frame4.Height = TabOne1.Height - 2115
    
    fg(0).Width = Frame4.Width - 150
    TabOne2.Width = Frame4.Width - 120
    TabOne2.Height = Frame4.Height - 1110
    
    fg(1).Width = TabOne2.Width - 1710
    fg(1).Height = TabOne2.Height - 735
    fg(2).Width = TabOne2.Width - 1710
    fg(2).Height = TabOne2.Height - 735
    fg(3).Width = TabOne2.Width - 1710
    fg(3).Height = TabOne2.Height - 735
    
    Cmd(4).Left = TabOne2.Width - 1575
    Cmd(5).Left = TabOne2.Width - 1575
    Cmd(6).Left = TabOne2.Width - 1575
    Cmd(7).Left = TabOne2.Width - 1575
    Cmd(8).Left = TabOne2.Width - 1575
    Cmd(9).Left = TabOne2.Width - 1575
    Cmd(10).Left = TabOne2.Width - 1575
    Cmd(11).Left = TabOne2.Width - 1575
    Cmd(12).Left = TabOne2.Width - 1575
    Cmd(13).Left = TabOne2.Width - 1575
    Cmd(14).Left = TabOne2.Width - 1575
    Cmd(15).Left = TabOne2.Width - 1575
    
End Sub

Private Sub iniciarCampos()
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
    CORR_ = -666
    
    ' -------------------------PROPIEDADES DE PROCESADO
    PROPIEDADES_.MODOTAREA_ = 3
    PROPIEDADES_.PORCENTAJE_ = 10
    PROPIEDADES_.MINUTOS_ = "00:10"
    PROPIEDADES_.INCLUIRREFRIGERIO_ = True
    PROPIEDADES_.HORINIREFRIGERIO_ = "13:00"
    PROPIEDADES_.HORFINREFRIGERIO_ = "14:00"
    PROPIEDADES_.LIMITARNUMEROPERSONAL_ = True
    PROPIEDADES_.LIMITARNUMEROTAREAS_ = True
    PROPIEDADES_.LIMITARSELPERSONAL_ = True
    
    '**********************
    ' CONFIGURACIONES GRID
    '**********************
    ' -------------------------------------------PROPIEDADES GRID
    fg(0).AllowUserResizing = flexResizeColumns
        
    fg(1).AllowUserResizing = flexResizeColumns
    fg(1).ExplorerBar = flexExSortShow
    fg(1).SelectionMode = flexSelectionByRow
    fg(1).ForeColorSel = &H80000005
    fg(1).BackColorSel = &H80&
            
    fg(2).AllowUserResizing = flexResizeColumns
    fg(2).ExplorerBar = flexExSortShow
    fg(2).SelectionMode = flexSelectionByRow
    fg(2).ForeColorSel = &H80000005
    fg(2).BackColorSel = &H80&
    
    fg(3).AllowUserResizing = flexResizeColumns
    fg(3).ExplorerBar = flexExSortShow
    fg(3).SelectionMode = flexSelectionByRow
    fg(3).ForeColorSel = &H80000005
    fg(3).BackColorSel = &H80&
    
    fg(4).AllowUserResizing = flexResizeColumns
    fg(4).ExplorerBar = flexExSortShow
    fg(4).SelectionMode = flexSelectionByRow
    fg(4).ForeColorSel = &H80000005
    fg(4).BackColorSel = &H80&
    ' -------------------------------------------TAMAÑOS GRID
    fg(0).ColWidth(COLUMNACABECERA_.IDRECETA_) = 0
    fg(0).ColWidth(COLUMNACABECERA_.IDLINEA_) = 0
    fg(0).ColWidth(COLUMNACABECERA_.IDUNIMED_) = 0

    fg(1).ColWidth(COLUMNADETALLETAREA_.IDTAR_) = 0
    fg(1).ColWidth(COLUMNADETALLETAREA_.IDAREA_) = 0
    fg(1).ColWidth(COLUMNADETALLETAREA_.IDRESP_) = 0

    fg(2).ColWidth(COLUMNADETALLEPERS_.IDPER_) = 0

    fg(3).ColWidth(COLUMNADETALLEREPROC_.IDLOTE_) = 0
    fg(3).ColWidth(COLUMNADETALLEREPROC_.IDLOTEDET_) = 0
    fg(3).ColWidth(COLUMNADETALLEREPROC_.IDALM_) = 0
    
    fg(4).ColWidth(COLUMNADETALLEINSUMOS_.IDINSUMO_) = 0
    fg(4).ColWidth(COLUMNADETALLEINSUMOS_.IDUNIMED_) = 0
    ' ------------------------------------------COMBOLIST GRID
    GRID_COMBOLIST fg(0), COLUMNACABECERA_.RECETA_
    GRID_COMBOLIST fg(0), COLUMNACABECERA_.UNIMED_
    GRID_COMBOLIST fg(0), COLUMNACABECERA_.LINEA_
    
    GRID_COMBOLIST fg(1), COLUMNADETALLETAREA_.AREA_
    GRID_COMBOLIST fg(1), COLUMNADETALLETAREA_.RESPONSABLE_
    
    GRID_COMBOLIST fg(2), COLUMNADETALLEPERS_.DNI_
    GRID_COMBOLIST fg(2), COLUMNADETALLEREPROC_.LOTE_
    ' ------------------------------------------FORMATOS GRID
    fg(0).ColEditMask(COLUMNACABECERA_.HORINI_) = "##:##"
    
    Dg1.Columns("numdoc").Alignment = dbgCenter
    
    ' Se agrega el mes Activo
    mMesActivo = xMes
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
       
    ESTADOANTERIOR_ = ESTADOPENDIENTE_
End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub pCargarGrid()
    Dim cSQL  As String
    Dim Rpta As Integer
    
    TDB_FiltroLimpiar Dg1
    
    cSQL = "SELECT con_librocosto.*, con_meses.descripcion AS desmes, mae_metodoval.descripcion AS desmetval, IIf([con_librocosto].[aplvtas]=0,'TODOS','VENTAS') AS desaplgasfab, IIf([con_librocosto].[tipo]=0,'GLOBAL','DISTRIBUIDO') AS destipdisgasfab " _
        + vbCr + "FROM (con_librocosto LEFT JOIN mae_metodoval ON con_librocosto.idmetodoval = mae_metodoval.id) LEFT JOIN con_meses ON con_librocosto.idmes = con_meses.id " _
        + vbCr + "ORDER BY con_librocosto.idmes DESC;"
        
    Me.MousePointer = vbHourglass
    
    RST_Busq RstOrdProd, cSQL, xCon
    Set Dg1.DataSource = RstOrdProd
    
    Me.MousePointer = vbDefault
    
    If RstOrdProd.State = 0 Then Exit Sub
End Sub

Private Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim Rpta As Integer
    
    Agregando = True
    Blanquea
    'llenarEstados
    If QueHace = 3 Then llenarEstado 1, 1, , cbEstado, , , True
    
    If RstOrdProd.RecordCount = 0 Then Exit Sub
    If RstOrdProd.EOF = True Then Exit Sub
     
    Set xRs = Nothing
    Agregando = True
    
    ' CABECERA
    cSQL = "SELECT * " _
        + vbCr + "FROM pro_ordenprod " _
        + vbCr + "WHERE (((pro_ordenprod.id)=" & NulosN(RstOrdProd("id")) & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    seleccionarIndiceCombo NulosN(xRs("estado")), cbEstado
    
    TxtFchPro.Valor = NulosC(xRs("fchpro"))
    TxtNumSer.Text = NulosC(xRs("numser"))
    TxtNumDoc.Text = NulosC(xRs("numdoc"))
    TxtIdTipDocRef.Text = NulosN(xRs("idtipdocref"))
    LblTipDocRef.Caption = UCase(Busca_Codigo(NulosN(xRs("idtipdocref")), "id", "descripcion", "mae_documento", "N", xCon))
    lbliddocref.Caption = NulosN(xRs("iddocref"))
    txtNumDocRef.Text = ""
    TxtIdResp.Text = NulosN(xRs("idresp"))
    lblResponsable.Caption = UCase(Busca_Codigo(NulosN(xRs("idresp")), "id", "nombre", "pla_empleados", "N", xCon))
    txtIdItem.Text = UCase(Busca_Codigo(NulosN(xRs("idrec")), "id", "iditem", "pro_receta", "N", xCon))
    lblItem.Caption = UCase(Busca_Codigo(NulosN(txtIdItem.Text), "id", "descripcion", "alm_inventario", "N", xCon))
    
    With fg(0)
        .Rows = .FixedRows
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, COLUMNACABECERA_.RECETA_) = UCase(Busca_Codigo(NulosN(xRs("idrec")), "id", "codrec", "pro_receta", "N", xCon))
        .TextMatrix(.Rows - 1, COLUMNACABECERA_.UNIMED_) = UCase(Busca_Codigo(NulosN(xRs("idunimed")), "id", "abrev", "mae_unidades", "N", xCon))
        .TextMatrix(.Rows - 1, COLUMNACABECERA_.CANTIDAD_) = Format(NulosN(xRs("cantidad")), FORMAT_CANTIDAD)
        .TextMatrix(.Rows - 1, COLUMNACABECERA_.LINEA_) = UCase(Busca_Codigo(NulosN(xRs("idlinea")), "id", "descripcion", "pro_linea", "N", xCon))
        .TextMatrix(.Rows - 1, COLUMNACABECERA_.EFICIENCIA_) = NulosN(xRs("efic"))
        .TextMatrix(.Rows - 1, COLUMNACABECERA_.HORINI_) = Format(NulosC(xRs("horini")), FORMAT_HORA_SIN_SEGUNDO)
        .TextMatrix(.Rows - 1, COLUMNACABECERA_.HORFIN_) = Format(NulosC(xRs("horfin")), FORMAT_HORA_SIN_SEGUNDO)
        .TextMatrix(.Rows - 1, COLUMNACABECERA_.FCHFIN_) = Format(NulosC(xRs("fchfin")), FORMAT_DATE)
        .TextMatrix(.Rows - 1, COLUMNACABECERA_.NUMOPE_) = NulosN(xRs("numop"))
        .TextMatrix(.Rows - 1, COLUMNACABECERA_.REPROC_) = NulosN(xRs("reproc"))
        .TextMatrix(.Rows - 1, COLUMNACABECERA_.IDRECETA_) = NulosN(xRs("idrec"))
        .TextMatrix(.Rows - 1, COLUMNACABECERA_.IDLINEA_) = NulosN(xRs("idlinea"))
        .TextMatrix(.Rows - 1, COLUMNACABECERA_.IDUNIMED_) = NulosN(xRs("idunimed"))
    End With
    
    pLlenarInsumos NulosN(xRs("idrec")), NulosN(xRs("cantidad"))
         
    ' DETALLE
    ' --------------------------------TAREAS
    cSQL = "SELECT pro_ordenprodtar.* " _
        + vbCr + "FROM pro_ordenprodtar " _
        + vbCr + "WHERE (((pro_ordenprodtar.idord)=" & NulosN(RstOrdProd("id")) & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    fg(1).Rows = fg(1).FixedRows
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then GoTo LLENARPERSONAL_
    
    xRs.MoveFirst
    With fg(1)
        While Not xRs.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.SEL_) = NulosN(xRs("activo"))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.TAREA_) = UCase(Busca_Codigo(NulosN(xRs("idtar")), "id", "descripcion", "pro_tareas", "N", xCon))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.DURACION_) = Format(NulosC(xRs("durtar")), "HH:mm")
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.HORINI_) = Format(NulosC(xRs("horini")), FORMAT_HORA_SIN_SEGUNDO)
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.HORFIN_) = Format(NulosC(xRs("horfin")), FORMAT_HORA_SIN_SEGUNDO)
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.NUMOP_) = NulosN(xRs("numop"))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.CANTIDADSUM_) = Format(NulosN(xRs("cantsum")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.CANTIDADPROC_) = Format(NulosN(xRs("cantproc")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.FCHINI_) = Format(NulosC(xRs("fchini")), FORMAT_DATE)
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.FCHFIN_) = Format(NulosC(xRs("fchfin")), FORMAT_DATE)
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.AREA_) = UCase(Busca_Codigo(NulosN(xRs("idarea")), "id", "descripcion", "mae_area", "N", xCon))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.RESPONSABLE_) = UCase(Busca_Codigo(NulosN(xRs("idsubresp")), "id", "nombre", "pla_empleados", "N", xCon))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.IDTAR_) = NulosN(xRs("idtar"))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.IDAREA_) = NulosN(xRs("idarea"))
            .TextMatrix(.Rows - 1, COLUMNADETALLETAREA_.IDRESP_) = NulosN(xRs("idsubresp"))
            xRs.MoveNext
        Wend
    End With
    
LLENARPERSONAL_:
    ' --------------------------------PERSONAL
    cSQL = "SELECT pro_ordenprodpers.* " _
        + vbCr + "FROM pro_ordenprodpers " _
        + vbCr + "WHERE (((pro_ordenprodpers.idord)=" & NulosN(RstOrdProd("id")) & "))"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    fg(2).Rows = fg(2).FixedRows
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then GoTo LLENARREPROCESO_
    
    xRs.MoveFirst
    With fg(2)
        While Not xRs.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COLUMNADETALLEPERS_.DNI_) = Busca_Codigo(NulosN(xRs("idper")), "id", "numdoc", "pla_empleados", "N", xCon)
            .TextMatrix(.Rows - 1, COLUMNADETALLEPERS_.NOMBRE_) = UCase(Busca_Codigo(NulosN(xRs("idper")), "id", "nombre", "pla_empleados", "N", xCon))
            .TextMatrix(.Rows - 1, COLUMNADETALLEPERS_.IDPER_) = NulosN(xRs("idper"))
            xRs.MoveNext
        Wend
    End With
   
LLENARREPROCESO_:
    ' --------------------------------REPROCESO
    cSQL = "SELECT pro_ordenprodreproc.* " _
        + vbCr + "FROM pro_ordenprodreproc " _
        + vbCr + "WHERE (((pro_ordenprodreproc.idord)=" & NulosN(RstOrdProd("id")) & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    fg(3).Rows = fg(3).FixedRows
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then GoTo SALIR_
    
    xRs.MoveFirst
    With fg(3)
        While Not xRs.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COLUMNADETALLEREPROC_.LOTE_) = Busca_Codigo(NulosN(xRs("idlote")), "id", "descripcion", "alm_lote", "N", xCon)
            .TextMatrix(.Rows - 1, COLUMNADETALLEREPROC_.CANTIDAD_) = Format(NulosC(xRs("cantidad")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, COLUMNADETALLEREPROC_.IDLOTE_) = NulosN(xRs("idlote"))
            .TextMatrix(.Rows - 1, COLUMNADETALLEREPROC_.IDLOTEDET_) = NulosN(xRs("idlotedet"))
            .TextMatrix(.Rows - 1, COLUMNADETALLEREPROC_.IDALM_) = NulosN(Busca_Codigo(NulosN(xRs("idlotedet")), "id", "idalm", "alm_lotedet", "N", xCon))
            .TextMatrix(.Rows - 1, COLUMNADETALLEREPROC_.ALMACEN_) = UCase(Busca_Codigo(NulosN(.TextMatrix(.Rows - 1, COLUMNADETALLEREPROC_.IDALM_)), "id", "descripcion", "alm_almacenes", "N", xCon))
            xRs.MoveNext
        Wend
    End With
SALIR_:
    Agregando = False
End Sub

Sub Cancelar()
    Bloquea
    Label5.Caption = "Detalle de Orden de Producción"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
     
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
End Sub

Sub Nuevo()
    'llenarEstados
    llenarEstado 1, 1, , cbEstado, , , False, ESTADOPENDIENTE_ & "," & ESTADOPROCESADO_
    
    QueHace = 1
    xHorIni = Time
    Bloquea
    Blanquea
    fg(0).Rows = 2
    fg(1).Rows = 1
    fg(2).Rows = 1
    fg(3).Rows = 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    
    Label5.Caption = "Agregando Orden de Producción"
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
    TxtFchPro.Valor = Date
    TxtNumSer.SetFocus
End Sub

Sub Bloquea()
    cbEstado.Locked = Not cbEstado.Locked
    TxtFchPro.Locked = Not TxtFchPro.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtIdResp.Locked = Not TxtIdResp.Locked
    txtIdItem.Locked = Not txtIdItem.Locked
    TxtIdTipDocRef.Locked = Not TxtIdTipDocRef.Locked
    txtNumDocRef.Locked = Not txtNumDocRef.Locked
    habilitar Cmd, Not TxtFchPro.Locked
End Sub

Sub Blanquea()
    TxtFchPro.Valor = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtIdResp.Text = ""
    lblResponsable.Caption = ""
    txtIdItem.Text = ""
    lblItem.Caption = ""
    TxtIdTipDocRef.Text = ""
    LblTipDocRef.Caption = ""
    lbliddocref.Caption = ""
    txtNumDocRef.Text = ""
End Sub

Function Grabar() As Boolean
    Dim IDORD_ As Integer
    Dim FCHORD_ As String
    Dim NUMSER_ As String
    Dim NUMDOC_ As String
    Dim IDRESP_ As Integer
    Dim IDTIPDOCREF_ As Integer
    Dim IDDOCREF_ As Integer
    Dim IDREC_ As Integer
    Dim IDUNIMED_ As Integer
    Dim CANTIDAD_ As Double
    Dim IDLINEA_ As Integer
    Dim EFIC_ As Integer
    Dim HORINI_ As String
    Dim HORFIN_ As String
    Dim FCHFIN_ As String
    Dim NUMOP_ As Integer
    Dim REPROC_ As Boolean
    Dim IDESTADO_ As Integer
    Dim xRs As New ADODB.Recordset
    
    Dim xRsTar As New ADODB.Recordset
    Dim xRsPer As New ADODB.Recordset
    Dim xRsRep As New ADODB.Recordset
    
    Dim xRsAux As New ADODB.Recordset
    Dim A As Integer
    
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If TxtFchPro.Valor = "" Then
        MsgBox "No ha especificado fecha de producción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchPro.SetFocus
        Exit Function
    End If
    
    If TxtIdResp.Text = "" Then
        MsgBox "No ha especificado un encargado para la solicitud", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdResp.SetFocus
        Exit Function
    End If
    
    If TxtNumSer.Text = "" Then
        MsgBox "No ha especificado el número de serie", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        Exit Function
    End If
    
    If TxtNumDoc.Text = "" Then
        MsgBox "No ha especificado el número de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        Exit Function
    End If
    
    If txtIdItem.Text = "" Then
        MsgBox "No ha especificado el Ítem para la producción actual", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txtIdItem.SetFocus
        Exit Function
    End If
    
    If fg(0).Rows = fg(0).FixedRows Then
        MsgBox "No ha especificado una descripción para la producción actual", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).Rows = fg(0).Rows + 1
        fg(0).SetFocus
        Exit Function
    End If
        
    ' Se llenan los detalles
    If QueHace = 1 Then IDORD_ = 0 Else IDORD_ = NulosN(RstOrdProd("id"))
    NUMSER_ = NulosC(TxtNumSer.Text)
    NUMDOC_ = NulosC(TxtNumDoc.Text)
    IDTIPDOCREF_ = NulosN(TxtIdTipDocRef.Text)
    IDDOCREF_ = NulosN(lbliddocref.Caption)
    FCHORD_ = Format(NulosC(TxtFchPro.Valor), "dd/mm/yyyy")
    IDRESP_ = NulosN(TxtIdResp.Text)
    IDREC_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.IDRECETA_))
    IDUNIMED_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.IDUNIMED_))
    CANTIDAD_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.CANTIDAD_))
    IDLINEA_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.IDLINEA_))
    EFIC_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.EFICIENCIA_))
    HORINI_ = Format(NulosC(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.HORINI_)), "HH:mm")
    HORFIN_ = Format(NulosC(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.HORFIN_)), "HH:mm")
    FCHFIN_ = Format(NulosC(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.FCHFIN_)), "dd/mm/yyyy")
    NUMOP_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.NUMOPE_))
    REPROC_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.REPROC_))
    IDESTADO_ = NulosN(cbEstado.ItemData(cbEstado.ListIndex))
    
    ' ------------------------------------------RECORDSET DE TAREAS
    If xRsTar.State = 0 Then
        cSQL = "SELECT TOP 1 * FROM pro_ordenprodtar;"
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        DEFINIR_RST_TMP xRsTar, xRs
    End If
    limpiarRST xRsTar
    For A = 1 To fg(1).Rows - 1
        xRsTar.AddNew
        xRsTar("idtar") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.IDTAR_))
        xRsTar("cantsum") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.CANTIDADSUM_))
        xRsTar("cantproc") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.CANTIDADPROC_))
        xRsTar("numop") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.NUMOP_))
        If NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.FCHINI_)) = "" Then
            xRsTar("fchini") = Null
        Else
            xRsTar("fchini") = NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.FCHINI_))
        End If
        If NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.FCHFIN_)) = "" Then
            xRsTar("fchfin") = Null
        Else
            xRsTar("fchfin") = NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.FCHFIN_))
        End If
        If NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.HORINI_)) = "" Then
            xRsTar("horini") = Null
        Else
            xRsTar("horini") = NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.HORINI_))
        End If
        If NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.HORFIN_)) = "" Then
            xRsTar("horfin") = Null
        Else
            xRsTar("horfin") = NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.HORFIN_))
        End If
        xRsTar("durtar") = NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.DURACION_))
        xRsTar("idsubresp") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.IDRESP_))
        xRsTar("idarea") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.IDAREA_))
        xRsTar("activo") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.SEL_))
        xRsTar.Update
    Next A
    ' ------------------------------------------RECORDSET DE PERSONAS
    If xRsPer.State = 0 Then
        cSQL = "SELECT TOP 1 * FROM pro_ordenprodpers;"
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        DEFINIR_RST_TMP xRsPer, xRs
    End If
    limpiarRST xRsPer
    For A = 1 To fg(2).Rows - 1
        xRsPer.AddNew
        xRsPer("idper") = NulosN(fg(2).TextMatrix(A, COLUMNADETALLEPERS_.IDPER_))
        xRsPer.Update
    Next A
    ' ------------------------------------------RECORDSET DE REPROCESOS
    If xRsRep.State = 0 Then
        cSQL = "SELECT TOP 1 * FROM pro_ordenprodreproc;"
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        DEFINIR_RST_TMP xRsRep, xRs
    End If
    limpiarRST xRsRep
    For A = 1 To fg(3).Rows - 1
        xRsRep.AddNew
        xRsRep("idlotedet") = NulosN(fg(3).TextMatrix(A, COLUMNADETALLEREPROC_.IDLOTEDET_))
        xRsRep("cantidad") = NulosN(fg(3).TextMatrix(A, COLUMNADETALLEREPROC_.CANTIDAD_))
        xRsRep.Update
    Next A
    
    ' Se graba el movimiento
    Grabar = grabarOrdProd(FCHORD_, IDTIPDOCREF_, IDDOCREF_, IDRESP_, IDREC_, IDUNIMED_, CANTIDAD_, _
                                    IDLINEA_, EFIC_, HORINI_, HORFIN_, FCHFIN_, NUMOP_, REPROC_, NUMDOC_, _
                                    xRsTar, xRsPer, xRsRep, NUMSER_, IDORD_, IDESTADO_, CInt(AnoTra), mMesActivo, QueHace)

    mIdRegistro = IDORD_
End Function

Private Sub pRecalcularReceta(IDRECETA_ As Integer, IDINSUMO_ As Integer, CANINSUMO_ As Double)
    Dim xRs As New ADODB.Recordset
    Dim CANTIDADTOTAL As Double
    
    cSQL = "SELECT pro_recetains.iditem AS idins, pro_receta.cantidad AS canitem, pro_recetains.canpro AS canins, pro_receta.idunimed " _
        + vbCr + "FROM pro_receta INNER JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec " _
        + vbCr + "WHERE (((pro_receta.id)=" & IDRECETA_ & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    xRs.Filter = "idins=" & IDINSUMO_
    If xRs.RecordCount = 0 Then Exit Sub
    If NulosN(xRs("canins")) = 0 Then Exit Sub
    
    CANTIDADTOTAL = (CANINSUMO_ * NulosN(xRs("canitem"))) / NulosN(xRs("canins"))
    
    xRs.Filter = adFilterNone
    xRs.MoveFirst
    fg(4).Rows = fg(4).FixedRows
    While Not xRs.EOF
        fg(4).Rows = fg(4).Rows + 1
        fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.IDINSUMO_) = NulosN(xRs("idins"))
        fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.CANTIDAD_) = (NulosN(xRs("canins")) * CANTIDADTOTAL) / NulosN(xRs("canitem"))
        fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.CANTIDAD_) = Format(fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.CANTIDAD_), FORMAT_CANTIDAD)
        fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.IDUNIMED_) = NulosN(xRs("idunimed"))
        fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.INSUMO_) = Busca_Codigo(NulosN(xRs("idins")), "id", "descripcion", "alm_inventario", "N", xCon)
        fg(4).TextMatrix(fg(4).Rows - 1, COLUMNADETALLEINSUMOS_.UNIMED_) = Busca_Codigo(NulosN(xRs("idunimed")), "id", "abrev", "mae_unidades", "N", xCon)
        xRs.MoveNext
    Wend
    
    fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.CANTIDAD_) = Format(CANTIDADTOTAL, FORMAT_CANTIDAD)
End Sub

Sub Modificar()
    llenarEstado 1, 1, , cbEstado, , , False, ESTADOPENDIENTE_ & "," & ESTADOPROCESADO_ 'llenarEstados
    
    If NulosN(RstOrdProd("estado")) > ESTADOPENDIENTE_ Then
        MsgBox "El registro está en un estado no modificable", vbInformation, xTitulo
        Exit Sub
    End If
            
    If RstOrdProd.RecordCount = 0 Then
        MsgBox "No hay Registros para Modificar", vbInformation, xTitulo
        Exit Sub
    End If
   
    QueHace = 2
    xHorIni = Time
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    
    Label5.Caption = "Modificando Solicitud de Materiales"
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
    
    xHorIni = Time
    TxtIdResp.SetFocus
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    Dim xRs As New ADODB.Recordset
    
    If RstOrdProd.RecordCount = 0 Then
        MsgBox "No hay documentos para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    TabOne1.CurrTab = 0
    Rpta = MsgBox("¿ Esta seguro de eliminar el Registro seleccionado ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_ordenprodreproc WHERE idord = " & NulosN(RstOrdProd("id"))
        xCon.Execute "DELETE * FROM pro_ordenprodpers WHERE idord = " & NulosN(RstOrdProd("id"))
        xCon.Execute "DELETE * FROM pro_ordenprodtar WHERE idord = " & NulosN(RstOrdProd("id"))
        xCon.Execute "DELETE * FROM pro_ordenprod WHERE id = " & NulosN(RstOrdProd("id"))
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & NulosN(RstOrdProd("id")) & " AND idform = " & IdMenuActivo
        
        MsgBox "El registro se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstOrdProd.Requery
        Dg1.Refresh
    End If
End Sub

Sub ExportarExcel()
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    Dim xCad As String
    Dim objExcel As Object
    
    Dim oExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    
    Set oExcel = New Excel.Application
    Set oWBook = oExcel.Workbooks.Add
    Screen.MousePointer = vbHourglass

    oExcel.WindowState = 2
    For A = 1 To fg(0).Rows - 1
        oWBook.ActiveSheet.Name = fg(0).TextMatrix(A, 10)
        With oWBook.ActiveSheet
            'Se llena cabecera
            .Cells(1, 2) = "SOLICITUD DE MATERIALES Nº:   " + "0001" & "-" & fg(0).TextMatrix(A, 10)
            .Range("B1", "H1").Merge
            .Cells(1, 2).HorizontalAlignment = xlHAlignCenterAcrossSelection
            .Cells(1, 2).Font.Bold = True
            .Cells(1, 2).Rows(1).Font.Size = 12
            
            .Cells(3, 2) = "Producción    Nº " + fg(0).TextMatrix(A, 6)
            .Cells(3, 2).Font.Bold = True
            .Cells(4, 2) = "Fch. Prod. :   " + TxtFchPro.Valor
            .Cells(4, 2).Font.Bold = True
            .Cells(5, 2) = "Producto :   " + fg(0).TextMatrix(A, 2)
            .Cells(5, 2).Font.Bold = True
            .Cells(6, 2) = "Receta :   " + fg(0).TextMatrix(A, 5)
            .Cells(6, 2).Font.Bold = True
            
            .Cells(7, 2) = "Cantidad  :   " + fg(0).TextMatrix(A, 4)
            .Cells(7, 2).Font.Bold = True
            
            .Cells(9, 2) = "Item"
            .Cells(9, 2).Font.Bold = True
            .Cells(9, 3) = "INSUMO / PRODUCTO / MP"
            .Cells(9, 3).Font.Bold = True
            .Cells(9, 4) = "Uni. Med."
            .Cells(9, 4).Font.Bold = True
            .Cells(9, 5) = "Cantidad Teorica"
            .Cells(9, 5).Font.Bold = True
            .Cells(9, 6) = "Cantidad Real"
            .Cells(9, 6).Font.Bold = True
            .Cells(9, 7) = "Adicional"
            .Cells(9, 7).Font.Bold = True
            .Cells(9, 8) = "Devolucion"
            .Cells(9, 8).Font.Bold = True
            
            Dim Rst As New ADODB.Recordset
            
            RST_Busq Rst, "SELECT pro_recetains.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, [pro_recetains]![canpro]*" & NulosN(fg(0).TextMatrix(A, 3)) & " AS canreq " _
                + vbCr + "FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
                + vbCr + "WHERE (((pro_recetains.idrec)=" & NulosN(fg(0).TextMatrix(A, 8)) & "))", xCon
        
            If Rst.RecordCount <> 0 Then
                Dim xFila As Integer
                xFila = 10
                For B = 1 To Rst.RecordCount
                    .Cells(xFila, 2) = Format(B, "00")
                    .Cells(xFila, 3) = Rst("descripcion")
                    .Cells(xFila, 4) = Rst("abrev")
                    .Cells(xFila, 5) = Format(Rst("canreq"), FORMAT_CANTIDADDECIMAL)
    
                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                    xFila = xFila + 1
                Next B
            End If
            .Cells(xFila + 5, 5) = "VºBº Ger. Prod. "
            .Cells(xFila + 5, 5).Font.Bold = True
            .Cells(xFila + 5, 7) = "Entregado Por "
            .Cells(xFila + 5, 7).Font.Bold = True
        End With
        If A < fg(0).Rows - 1 Then oWBook.Sheets.Add
    Next A
    
    oExcel.Visible = True
    
    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Reporte de Pedidos"
    oExcel.WindowState = 1
    
    Set oExcel = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub limpiarRST(Rst As ADODB.Recordset, Optional TODO As Boolean = True)
    With Rst
        If .State <> 0 Then
            If TODO Then .Filter = adFilterNone
            If .RecordCount <> 0 Then
                .MoveFirst
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
            End If
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If CAMBIOGRABAR_ = -1 Then
'        MsgBox "No se puede Cancelar la operación; Grabe los registros para continuar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Cancel = 1
'    End If
End Sub

Private Sub LblDetTrab_Click()
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 1 Then Exit Sub
        MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then
        If RstOrdProd.RecordCount = 0 Then
            MsgBox "No se han registardos ventas para realizar esta opcion", vbInformation, Me.Caption
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.Index = 3 Then
        If RstOrdProd.RecordCount = 0 Then
            MsgBox "No se han registrados Pedidos para realizar esta opción", vbInformation, Me.Caption
            Exit Sub
        End If
        Eliminar
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstOrdProd.Requery
            Dg1.Refresh
            If RstOrdProd.RecordCount <> 0 Then
                RstOrdProd.MoveFirst
                RstOrdProd.Find "id=" & mIdRegistro
                If RstOrdProd.EOF = True Then RstOrdProd.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        If TabOne1.CurrTab = 0 Then RstOrdProd.Filter = "": TDB_FiltroLimpiar Dg1
    End If
    
    If Button.Index = 12 Then CambiarMes
    
    If Button.Index = 14 Then ExportarExcel
    If Button.Index = 15 Then imprimir 0
    If Button.Index = 17 Then Unload Me
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        If ButtonMenu.Index = 1 Then ' ANULAR REGISTRO
            If TabOne1.CurrTab = 1 Then TabOne1.CurrTab = 0
            anular
        End If
    End If
End Sub

Private Sub imprimir(TIPO_ As Integer)
    'TIPO_ = 0:LINEA
    'TIPO_ = 1:ACABADO
    'TIPO_ = 2:REPORTE
    Dim xLinea As Integer
    Dim xform As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim nSQLFiltro As String '--Almacenara el filtro por movimiento
    Dim xCampos(6, 5) As String
    Dim xRsTar As New ADODB.Recordset
    Dim xRsPer As New ADODB.Recordset
    Dim A As Integer
    
    Select Case TIPO_
        Case 0
            If NulosN(RstOrdProd("estado")) = ESTADOPENDIENTE_ Then
                MsgBox "El registro actual no se puede imprimir debido a su estado", vbInformation, xTitulo
                Exit Sub
            End If
            ' ------------------------------------------RECORDSET DE TAREAS
            If xRsTar.State = 0 Then
                cSQL = "SELECT TOP 1 * FROM pro_ordenprodtar;"
                Set xRs = Nothing
                RST_Busq xRs, cSQL, xCon
                DEFINIR_RST_TMP xRsTar, xRs
            End If
            limpiarRST xRsTar
            For A = 1 To fg(1).Rows - 1
                If NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.SEL_)) = -1 Then
                    xRsTar.AddNew
                    xRsTar("idtar") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.IDTAR_))
                    xRsTar("cantsum") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.CANTIDADSUM_))
                    xRsTar("cantproc") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.CANTIDADPROC_))
                    xRsTar("numop") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.NUMOP_))
                    If NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.FCHINI_)) = "" Then
                        xRsTar("fchini") = Null
                    Else
                        xRsTar("fchini") = NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.FCHINI_))
                    End If
                    If NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.FCHFIN_)) = "" Then
                        xRsTar("fchfin") = Null
                    Else
                        xRsTar("fchfin") = NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.FCHFIN_))
                    End If
                    If NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.HORINI_)) = "" Then
                        xRsTar("horini") = Null
                    Else
                        xRsTar("horini") = NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.HORINI_))
                    End If
                    If NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.HORFIN_)) = "" Then
                        xRsTar("horfin") = Null
                    Else
                        xRsTar("horfin") = NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.HORFIN_))
                    End If
                    xRsTar("durtar") = NulosC(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.DURACION_))
                    xRsTar("idsubresp") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.IDRESP_))
                    xRsTar("idarea") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.IDAREA_))
                    xRsTar("activo") = NulosN(fg(1).TextMatrix(A, COLUMNADETALLETAREA_.SEL_))
                    xRsTar.Update
                End If
            Next A
            ' ------------------------------------------RECORDSET DE PERSONAS
            If xRsPer.State = 0 Then
                cSQL = "SELECT TOP 1 * FROM pro_ordenprodpers;"
                Set xRs = Nothing
                RST_Busq xRs, cSQL, xCon
                DEFINIR_RST_TMP xRsPer, xRs
            End If
            limpiarRST xRsPer
            For A = 1 To fg(2).Rows - 1
                xRsPer.AddNew
                xRsPer("idper") = NulosN(fg(2).TextMatrix(A, COLUMNADETALLEPERS_.IDPER_))
                xRsPer.Update
            Next A
            ' ------------------------------------------IMPRESION
            With FrmVsPrinter.Vs
                .StartDoc
                Me.MousePointer = vbHourglass
                ImprimirLinea NulosC(TxtNumSer.Text) & "-" & NulosC(TxtNumDoc.Text), NulosC(lblItem.Caption), _
                            NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.IDRECETA_)), _
                            NulosC(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.RECETA_)), _
                            NulosC(TxtFchPro.Valor), NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.CANTIDAD_)), _
                            NulosC(lblResponsable.Caption), NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.UNIMED_)), _
                            NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.NUMOPE_)), xRsTar, xRsPer
                Me.MousePointer = vbDefault
                .EndDoc
            End With
            'Muestra la preimagen de la impresion
            FrmVsPrinter.WindowState = 2
            FrmVsPrinter.Show
        
        Case 1
        Case 2
            
    End Select
End Sub

Private Sub CambiarMes()
    TabOne1.CurrTab = 0
    mMesActivo = SeleccionaMes(xCon)
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    pCargarGrid
End Sub

Private Sub txtIdItem_KeyPress(KeyAscii As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub txtIdItem_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 0
    End If
End Sub

Private Sub TxtIdResp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 3
    End If
End Sub

Private Sub TxtIdTipDocRef_KeyPress(KeyAscii As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdTipDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 1
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If NulosC(TxtNumSer.Text) = "" Then
        MsgBox "Ingrese un número de serie", vbInformation, Me.Caption
        TxtNumDoc.Text = ""
        TxtNumSer.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtNumDocRef_KeyPress(KeyAscii As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub txtNumDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 2
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtNumSer.Text) <> "" Then
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
        TxtNumDoc.Text = hallarNumDoc("pro_ordenprod", "'" & NulosC(TxtNumSer.Text) & "'", "numser")
    End If
End Sub
