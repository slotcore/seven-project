VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmAnalisisCuenta1 
   Caption         =   "Contabilidad - Analisis de Cuentas"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7665
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      Caption         =   "[ Expresado en ]"
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
      Height          =   555
      Left            =   0
      TabIndex        =   25
      Top             =   1650
      Width           =   3075
      Begin VB.CommandButton CmdBusMon 
         Height          =   230
         Left            =   1035
         Picture         =   "FrmAnalisisCuenta1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   225
         Width           =   210
      End
      Begin VB.TextBox TxtIdMon 
         Height          =   300
         Left            =   720
         MaxLength       =   1
         TabIndex        =   27
         Text            =   "TxtIdMon"
         Top             =   195
         Width           =   555
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
         Height          =   285
         Left            =   1260
         TabIndex        =   29
         Top             =   210
         Width           =   1755
      End
      Begin VB.Label LblTipCam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   28
         Top             =   270
         Width           =   585
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5460
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":0132
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":0676
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":0A08
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":0B8C
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":0FE0
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":10F8
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":163C
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":1B80
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":1C94
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":1DA8
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":21FC
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":2368
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":28B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":2BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":2F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisisCuenta1.frx":32EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   5295
      Left            =   0
      TabIndex        =   18
      Top             =   2370
      Width           =   11880
      _cx             =   20955
      _cy             =   9340
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
      BackColor       =   14745342
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   128
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14745342
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   16
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmAnalisisCuenta1.frx":3680
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
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   1245
         Left            =   3435
         TabIndex        =   19
         Top             =   1485
         Visible         =   0   'False
         Width           =   5010
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   330
            Left            =   120
            TabIndex        =   20
            Top             =   630
            Width           =   4770
            _ExtentX        =   8414
            _ExtentY        =   582
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exportando a Excel"
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
            Left            =   150
            TabIndex        =   21
            Top             =   105
            Width           =   1665
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000002&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000002&
            Height          =   315
            Left            =   45
            Top             =   45
            Width           =   4935
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            BorderWidth     =   2
            Index           =   1
            X1              =   15
            X2              =   15
            Y1              =   0
            Y2              =   1200
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            Index           =   0
            X1              =   4995
            X2              =   4995
            Y1              =   30
            Y2              =   1230
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            BorderWidth     =   2
            Index           =   1
            X1              =   0
            X2              =   4995
            Y1              =   15
            Y2              =   15
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            Index           =   0
            X1              =   15
            X2              =   5010
            Y1              =   1230
            Y2              =   1230
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   5
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
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Configurar Formatos"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar a PDT"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Ordenado Por ]"
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
      Height          =   1290
      Left            =   6855
      TabIndex        =   6
      Top             =   330
      Width           =   2910
      Begin VB.OptionButton OptSort4 
         Caption         =   "Fch. Emision y Nº de Documento"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   1020
         Width           =   2730
      End
      Begin VB.OptionButton OptSort3 
         Caption         =   "Nº Registro"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   765
         Width           =   2010
      End
      Begin VB.OptionButton OptSort1 
         Caption         =   "Fecha  de Emision"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   270
         Width           =   2010
      End
      Begin VB.OptionButton OptSort2 
         Caption         =   "Nº de Documento"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   510
         Width           =   2010
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   0
      TabIndex        =   11
      Top             =   330
      Width           =   6825
      Begin VB.CommandButton CmdBusProv 
         Height          =   230
         Left            =   6450
         Picture         =   "FrmAnalisisCuenta1.frx":3875
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   270
         Width           =   210
      End
      Begin VB.TextBox TxtFormato 
         Height          =   300
         Left            =   1050
         TabIndex        =   0
         Text            =   "TxtFormato"
         Top             =   240
         Width           =   5640
      End
      Begin VB.OptionButton OptDol 
         Caption         =   "Analitico"
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
         Left            =   2520
         TabIndex        =   4
         Top             =   975
         Width           =   1245
      End
      Begin VB.OptionButton OptSol 
         Caption         =   "Tributario"
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
         Left            =   1035
         TabIndex        =   3
         Top             =   975
         Width           =   1245
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   1725
         TabIndex        =   1
         Top             =   570
         Width           =   1305
         _ExtentX        =   2302
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
         Valor           =   "11/09/2008"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   3780
         TabIndex        =   2
         Top             =   570
         Width           =   1305
         _ExtentX        =   2302
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
         Valor           =   "11/09/2008"
      End
      Begin VB.Label LblIdFormato 
         Caption         =   "LblIdFormato"
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   5790
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Formato"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   22
         Top             =   255
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Index           =   2
         Left            =   3225
         TabIndex        =   14
         Top             =   615
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   1065
         TabIndex        =   13
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Periodo :"
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
         Left            =   105
         TabIndex        =   12
         Top             =   585
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Datos ]"
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
      Height          =   1290
      Left            =   9810
      TabIndex        =   15
      Top             =   330
      Width           =   2055
      Begin VB.Label LblNumreg 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblNumreg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   180
         TabIndex        =   17
         Top             =   630
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Registros :"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   390
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmAnalisisCuenta1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean

Private Sub CmdBusProv_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "5500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM con_analisis ORDER BY descripcion "
    
    xform.Titulo = "Buscando Formatos del Analisis del Cuenta"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        TxtFormato.Text = NulosC(xRs("descripcion"))
        LblIdFormato.Caption = NulosN(xRs("id"))
        
        SetearCuadricula Fg1, NulosN(xRs("id")), xCon, 2
        TxtFchIni.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        OptSol.Value = True
        TxtFchIni.SetFocus
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    OptSol.Value = True
    'OptOpc11.Value = True
    OptSort3.Value = True
    LblNumreg.Caption = ""
    TxtFormato.Text = ""
    LblIdFormato.Caption = ""
    LblIdFormato.Caption = "1"
    TxtFormato.Text = Busca_Codigo(NulosN(LblIdFormato.Caption), "id", "descripcion", "con_analisis", "N", xCon)
    SetearCuadricula Fg1, 1, xCon, 2, , False
End Sub

Sub MostrarFormato312Tribu()
    Dim Rst As New ADODB.Recordset
    Dim xTotal As Double
    Dim A As Integer

    Me.MousePointer = vbHourglass
    Fg1.Rows = 2
    DoEvents
    
    RST_Busq Rst, "SELECT com_compras.idpro, mae_dociden.codsun, mae_prov.numruc, mae_prov.nombre, Sum(IIf([com_compras].[idmon]=2,[impsal],0)) AS saldodol, " _
        & " Sum(IIf([com_compras].[idmon]=1,[impsal],0)) AS saldosol, Sum(IIf([com_compras].[idmon]=1,[impsal],[impsal]*[con_tc]![impven])) AS saldototal" _
        & " FROM (mae_prov LEFT JOIN mae_dociden ON mae_prov.idtipdoc = mae_dociden.id) RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) " _
        & " ON mae_prov.id = com_compras.idpro WHERE (((com_compras.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (com_compras.fchreg)<=CDate('" & TxtFchFin.Valor & "'))) " _
        & " GROUP BY com_compras.idpro, mae_dociden.codsun, mae_prov.numruc, mae_prov.nombre ORDER BY mae_prov.nombre" _
        & " Union " _
        & " SELECT com_honorarios.idpro, mae_dociden.codsun, mae_prov.numruc, mae_prov.nombre, Sum(IIf([com_honorarios].[idmon]=2,[impsal],0)) AS saldodol, " _
        & " Sum(IIf([com_honorarios].[idmon]=1,[impsal],0)) AS saldosol, Sum(IIf([com_honorarios].[idmon]=1,[impsal],[impsal]*[con_tc]![impven])) AS saldototal " _
        & " FROM ((com_honorarios LEFT JOIN mae_prov ON com_honorarios.idpro = mae_prov.id) LEFT JOIN mae_dociden ON mae_prov.idtipdoc = mae_dociden.id) " _
        & " LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha WHERE (((com_honorarios.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (com_honorarios.fchreg)<=CDate('" & TxtFchFin.Valor & "'))) " _
        & " GROUP BY com_honorarios.idpro, mae_dociden.codsun, mae_prov.numruc, mae_prov.nombre", xCon
      
    If Rst.RecordCount <> 0 Then
        LblNumreg.Caption = Rst.RecordCount
        Rst.MoveFirst
        Rst.Sort = "nombre"
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("codsun")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("numruc")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Rst("nombre")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosN(Rst("saldodol")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(Rst("saldosol")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(Rst("saldototal")), FORMAT_MONTO)
            Rst.MoveNext
            
            If Rst.EOF = True Then Exit For
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(GRID_SUMAR_COL(Fg1, 4), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(GRID_SUMAR_COL(Fg1, 5), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(GRID_SUMAR_COL(Fg1, 6), FORMAT_MONTO)
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &HC00000, True, &HE0FEFE, "TOTALES ==>"
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &HC00000, True, &HE0FEFE
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &HC00000, True, &HE0FEFE
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &HC00000, True, &HE0FEFE
    End If
    Set Rst = Nothing
    Me.MousePointer = vbDefault
End Sub

Sub MostrarFormato33Tribu()
    Dim Rst As New ADODB.Recordset
    Dim xTotal As Double
    Dim A As Integer
    Me.MousePointer = vbHourglass
    Fg1.Rows = 2
    DoEvents
    
    RST_Busq Rst, "SELECT vta_ventas.idcli, mae_dociden.codsun, mae_cliente.numruc, mae_cliente.nombre, IIf([vta_ventas].[idmon]=2,[impsal],0) AS saldodol, " _
        & " Sum(IIf([vta_ventas].[idmon]=1,[impsal],0)) AS saldosol, Sum(IIf([vta_ventas].[idmon]=1,[impsal],[impsal]*[con_tc].[impven])) AS saldototal " _
        & " FROM (mae_cliente LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id) RIGHT JOIN (vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) " _
        & " ON mae_cliente.id = vta_ventas.idcli WHERE (((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "'))) " _
        & " GROUP BY vta_ventas.idcli, mae_dociden.codsun, mae_cliente.numruc, mae_cliente.nombre, IIf([vta_ventas].[idmon]=2,[impsal],0), mae_cliente.nombre " _
        & " ORDER BY mae_cliente.nombre", xCon

    If Rst.RecordCount <> 0 Then
        LblNumreg.Caption = Rst.RecordCount
        Rst.MoveFirst
        
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("codsun")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("numruc")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Rst("nombre")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosN(Rst("saldodol")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(Rst("saldosol")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(Rst("saldototal")), FORMAT_MONTO)
            
            Rst.MoveNext
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(GRID_SUMAR_COL(Fg1, 4), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(GRID_SUMAR_COL(Fg1, 5), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(GRID_SUMAR_COL(Fg1, 6), FORMAT_MONTO)
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &HC00000, True, &HE0FEFE, "TOTALES ==>"
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &HC00000, True, &HE0FEFE
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &HC00000, True, &HE0FEFE
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &HC00000, True, &HE0FEFE
    End If
    Set Rst = Nothing
    Me.MousePointer = vbDefault
End Sub

Sub MostrarFormato33Anal()
    Dim Rst As New ADODB.Recordset
    Dim xTotal As Double
    Dim A, xIdCliente As Integer

    Me.MousePointer = vbHourglass
    Fg1.Rows = 2
    DoEvents
    
    RST_Busq Rst, "SELECT vta_ventas.idcli, mae_dociden.codsun, mae_cliente.numruc, IIf(vta_ventas.idmon=2,[impsal],0) AS saldodol, " _
        & " IIf(vta_ventas.idmon=1,[impsal],0) AS saldosol, IIf(vta_ventas.idmon=1,[impsal],[impsal]*con_tc.impven) AS saldototal, " _
        & " Mid(vta_ventas!numreg,1,2) & mae_libros!codsun & Mid(vta_ventas!numreg,3,4) AS numasi, mae_cliente.nombre, " _
        & " vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc, vta_ventas.imptotdoc, mae_moneda.simbolo, mae_docreferencia.descripcion AS tipdocref, " _
        & " vta_ventas.idtipdocref, IIf([vta_ventas]![idtipdocref]=5,(SELECT vta_pedido.numcen AS numdoc FROM vta_pedido " _
        & " WHERE (vta_pedido.id=vta_ventas.iddocref2)),(SELECT [var_ordendespacho]![anno] & [var_ordendespacho]![idaduana] & [var_ordendespacho]![idregimen] & [var_ordendespacho]![numdoc] AS Expr2 " _
        & " FROM var_ordendespacho WHERE (((var_ordendespacho.id)=vta_ventas.iddocref2)))) AS numdocref, vta_ventas.iddocref2, vta_ventas.fchdoc, " _
        & " vta_ventas.fchven, mae_documento.abrev FROM ((mae_cliente LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id) RIGHT JOIN " _
        & " ((((vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN " _
        & " mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_docreferencia ON vta_ventas.idtipdocref = mae_docreferencia.id) " _
        & " ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id WHERE " _
        & " (((IIf([vta_ventas].[idmon]=1,[impsal],[impsal]*[con_tc].[impven]))<>0) AND ((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_ventas.anulado)=0)) ORDER BY mae_cliente.nombre, vta_ventas!numser+'-'+vta_ventas!numdoc", xCon
        
    If Rst.RecordCount <> 0 Then
        LblNumreg.Caption = Rst.RecordCount
        Rst.MoveFirst
        
        xIdCliente = Rst("idcli")
        Fg1.Rows = Fg1.Rows + 1
        GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 12, "CLIENTE   :  " + Rst("nombre"), flexAlignLeftCenter, True, flexMergeFree, &HC00000, &HE0FEFE, True
        
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("numasi")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("abrev")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(Rst("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Rst("simbolo")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Rst("numdoc")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(Rst("fchven"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(Rst("imptotdoc")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(Rst("tipdocref"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(Rst("numdocref"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(Rst("saldodol"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(Rst("saldosol"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(Rst("saldototal"), FORMAT_MONTO)
            
            Rst.MoveNext
            
            If Rst.EOF = True Then Exit For
            
            If xIdCliente <> Rst("idcli") Then
                xIdCliente = Rst("idcli")
                Fg1.Rows = Fg1.Rows + 2
                GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 12, "CLIENTE   :  " + Rst("nombre"), flexAlignLeftCenter, True, flexMergeFree, &HC00000, &HE0FEFE, True
            End If
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        
        Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(GRID_SUMAR_COL(Fg1, 10), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(GRID_SUMAR_COL(Fg1, 11), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(GRID_SUMAR_COL(Fg1, 12), FORMAT_MONTO)
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &HC00000, True, &HE0FEFE, "TOTALES ==>"
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &HC00000, True, &HE0FEFE
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &HC00000, True, &HE0FEFE
    End If
    Set Rst = Nothing
    
    Me.MousePointer = vbDefault
    
End Sub

Sub MostrarFormato32()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Dim xTotal As Double
    
    Me.MousePointer = vbHourglass
    DoEvents
    RST_Busq Rst, "SELECT con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion AS desccta, mae_bancos.descripcion AS descban, mae_banconumcta.numcue, " _
        & " mae_banconumcta.idmon, mae_moneda.simbolo, Sum(IIf([con_diario]![impdebdol]<>0,[con_diario]![impdebdol]*[con_tc]![impven],[con_diario]![impdebsol])) AS TotDeb, " _
        & " Sum(IIf([con_diario]![imphabdol]<>0,[con_diario]![imphabdol]*[con_tc]![impven],[con_diario]![imphabsol])) AS TotHab, " _
        & " Sum(IIf([mae_banconumcta].[idmon]=2,[con_diario]![impdebdol],0)) AS TotDebDol, Sum(IIf([mae_banconumcta].[idmon]=2,[con_diario]![impHabdol],0)) AS TotHabDol " _
        & " FROM (con_planctas RIGHT JOIN (mae_bancos RIGHT JOIN (mae_banconumcta LEFT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) " _
        & " ON mae_banconumcta.idcuen = con_diario.idcue) ON mae_bancos.id = mae_banconumcta.idban) ON con_planctas.id = mae_banconumcta.idcuen) " _
        & " LEFT JOIN mae_moneda ON mae_banconumcta.idmon = mae_moneda.id GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion, " _
        & " mae_bancos.descripcion, mae_banconumcta.numcue, mae_banconumcta.idmon, mae_moneda.simbolo, con_diario.fchasi " _
        & " HAVING (((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "')))", xCon

    Fg1.Rows = 2
    DoEvents
    
    If Rst.RecordCount <> 0 Then
        LblNumreg.Caption = Rst.RecordCount
        Rst.MoveFirst
        
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("cuenta")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Rst("desccta"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(Rst("descban"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(Rst("numcue"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(Rst("simbolo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(Rst("totdebdol")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(Rst("tothabdol")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(Rst("totdeb")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(NulosN(Rst("tothab")), FORMAT_MONTO)
            
            If (Rst("totdeb") - Rst("tothab")) > 0 Then
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format((Rst("totdeb") - Rst("tothab")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 11) = "0.00"
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = "0.00"
                Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format((Rst("totdeb") - Rst("tothab")), FORMAT_MONTO)
            End If
            xTotal = (xTotal + (NulosN(Rst("totdeb")) - NulosN(Rst("tothab"))))
            Rst.MoveNext
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        
        Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(GRID_SUMAR_COL(Fg1, 10), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(GRID_SUMAR_COL(Fg1, 11), FORMAT_MONTO)
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &HC00000, True, &HE0FEFE, "TOTALES"
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &HC00000, True, &HE0FEFE
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &HC00000, True, &HE0FEFE
    End If
    Set Rst = Nothing
    Me.MousePointer = vbDefault
End Sub

Sub MostrarRegistros()
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "No ha especificado la fecha de final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    LblNumreg.Caption = "0"
    
    fGenerarConsulta
    
'    If LblIdFormato.Caption = 1 Then
'        MostrarFormato32
'    End If
'    If LblIdFormato.Caption = 2 Then
'        If OptSol.Value = True Then
'            MostrarFormato33Tribu
'        Else
'            MostrarFormato33Anal
'        End If
'    End If
'
'    If LblIdFormato.Caption = 3 Then
'        If OptSol.Value = True Then
'            MostrarFormato312Tribu
'        Else
'            MostrarFormato312Anal
'        End If
'    End If
    
    Exit Sub
End Sub

Private Sub OptDol_Click()
    LblNumreg.Caption = ""
    SetearCuadricula Fg1, NulosN(LblIdFormato.Caption), xCon, 2, 2, False
End Sub

Private Sub OptSol_Click()
    LblNumreg.Caption = ""
    SetearCuadricula Fg1, NulosN(LblIdFormato.Caption), xCon, 2, 1, False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        If TxtFchIni.Valor = "" And TxtFchFin.Valor = "" Then
            MsgBox "No ha especificado el periodo de la consulta", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
            TxtFchIni.SetFocus
            Exit Sub
        End If

        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
            TxtFchIni.SetFocus
            Exit Sub
        End If
        
        MostrarRegistros
    End If
    
    If Button.Index = 3 Then
        If Fg1.Rows = 2 Then
            MsgBox "No hay registro que exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Dim xFun As New SGI2_funciones.formularios
        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "ANALISIS DE CUENTA", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, TxtFormato.Text, "analisis.xls"
        Set xFun = Nothing
    End If
    
    If Button.Index = 4 Then
        Dim xMoneda As String
        Dim nPeriodo As String
        Dim xPrint As New SGI2_funciones.formularios

        If OptSol.Value = True Then
            xMoneda = "Nuevos Soles"
        Else
            xMoneda = "Dolares Americanos"
        End If
        
        nPeriodo = "Del " & TxtFchIni.Valor & " Al " & TxtFchFin.Valor
        Me.MousePointer = vbHourglass
        xPrint.Imprimir_x_VSFlexGrid Fg1, "ANALISIS DE CUENTA ", TxtFormato.Text, nPeriodo, False, True
        Set xPrint = Nothing
        Me.MousePointer = vbDefault
    End If
    
    If Button.Index = 5 Then Configurar
    
    'If Button.Index = 6 Then ExportarPDT
        
    If Button.Index = 8 Then
        Unload Me
    End If
End Sub

Sub Configurar()
    Dim xform As New SGI2_funciones.Varias
    If xform.CambioOpcionLiro(1, xCon, 2) = True Then
        SetearCuadricula Fg1, 1, xCon, 2
        If TxtFchIni.Valor = "" And TxtFchFin.Valor = "" Then
            MsgBox "No ha especificado el periodo de la consulta", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
            TxtFchIni.SetFocus
            Exit Sub
        End If

        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
            TxtFchIni.SetFocus
            Exit Sub
        End If
        MostrarRegistros
    End If
    Set xform = Nothing
End Sub

Sub MostrarFormato312Anal()
    Dim Rst As New ADODB.Recordset
    Dim xTotal As Double
    Dim A As Integer
    Dim xIdProveedor As Integer
    
    Me.MousePointer = vbHourglass
    Fg1.Rows = 2
    DoEvents
    
    RST_Busq Rst, "SELECT com_honorarios.idpro, mae_dociden.codsun, mae_prov.numruc, Mid([com_honorarios]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_honorarios]![numreg],3,4) AS numasi, " _
        & " mae_prov.nombre, [com_honorarios]![numser] & '-' & [com_honorarios]![numdoc] AS numdoc, com_honorarios.imptot, mae_moneda.simbolo, mae_docreferencia.descripcion AS tipdocref, " _
        & " com_honorarios.iddocref2, com_honorarios.fchdoc, com_honorarios.fchven, mae_documento.abrev, IIf([com_honorarios]![idmon]=2,[com_honorarios]![impsal],0) AS saldodol, " _
        & " IIf([com_honorarios]![idmon]=1,[com_honorarios]![impsal],0) AS saldosol, IIf([com_honorarios]![idmon]=1,[com_honorarios]![impsal],[com_honorarios]![impsal]*[con_tc].[impven]) AS saldototal, " _
        & " '' AS numdocref FROM (((((com_honorarios LEFT JOIN mae_moneda ON com_honorarios.id = mae_moneda.id) LEFT JOIN mae_documento ON com_honorarios.tipdoc = mae_documento.id) " _
        & " LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) LEFT JOIN mae_docreferencia ON com_honorarios.idtipdocref = mae_docreferencia.id) LEFT JOIN con_tc ON " _
        & " com_honorarios.fchdoc = con_tc.fecha) LEFT JOIN (mae_dociden RIGHT JOIN mae_prov ON mae_dociden.id = mae_prov.idtipdoc) ON com_honorarios.idpro = mae_prov.id " _
        & " WHERE (((com_honorarios.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (com_honorarios.fchreg)<=CDate('" & TxtFchFin.Valor & "'))) " _
        & " Union  " _
        & " SELECT com_compras.idpro, mae_dociden.codsun, mae_prov.numruc, Mid([com_compras]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_compras]![numreg],3,4) AS numasi, " _
        & " mae_prov.nombre, [com_compras]![numser] & '-' & [com_compras]![numdoc] AS numdoc, com_compras.imptot, mae_moneda.simbolo, mae_docreferencia.descripcion AS tipdocref, " _
        & " com_compras.iddocref2, com_compras.fchdoc, com_compras.fchven, mae_documento.abrev, IIf([com_compras]![idmon]=2,[com_compras]![impsal],0) AS saldodol, " _
        & " IIf([com_compras]![idmon]=1,[com_compras]![impsal],0) AS saldosol, IIf([com_compras]![idmon]=1,[com_compras]![impsal],[com_compras]![impsal]*[con_tc].[impven]) AS saldototal, '' AS numdocref " _
        & " FROM (mae_dociden RIGHT JOIN (((mae_documento RIGHT JOIN (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN com_compras ON mae_prov.id = com_compras.idpro) ON mae_moneda.id = com_compras.idmon) " _
        & " ON mae_documento.id = com_compras.tipdoc) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_dociden.id = mae_prov.idtipdoc) " _
        & " LEFT JOIN mae_docreferencia ON com_compras.idtipdocref = mae_docreferencia.id WHERE (((com_compras.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (com_compras.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((IIf([com_compras]![idmon]=1,[com_compras]![impsal],[com_compras]![impsal]*[con_tc].[impven]))<>0)) " _
        & " ORDER BY mae_prov.nombre", xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.Sort = "nombre, fchdoc"
        LblNumreg.Caption = Rst.RecordCount
        Rst.MoveFirst
        
        xIdProveedor = Rst("idpro")
        Fg1.Rows = Fg1.Rows + 1
        GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 12, "PROVEEDOR   :  " + Rst("nombre"), flexAlignLeftCenter, True, flexMergeFree, &HC00000, &HE0FEFE, True
        
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst("numasi"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Rst("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(Rst("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(Rst("simbolo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(Rst("numdoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(Rst("fchven"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(Rst("imptot")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(Rst("tipdocref"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(Rst("numdocref"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(Rst("saldodol")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(NulosN(Rst("saldosol")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(NulosN(Rst("saldototal")), FORMAT_MONTO)
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
            
            If xIdProveedor <> Rst("idpro") Then
                xIdProveedor = Rst("idpro")
                Fg1.Rows = Fg1.Rows + 2
                GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 12, "PROVEEDOR   :  " + Rst("nombre"), flexAlignLeftCenter, True, flexMergeFree, &HC00000, &HE0FEFE, True
            End If
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        
        Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(GRID_SUMAR_COL(Fg1, 10), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(GRID_SUMAR_COL(Fg1, 11), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(GRID_SUMAR_COL(Fg1, 12), FORMAT_MONTO)
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &HC00000, True, &HE0FEFE, "TOTALES ==>"
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &HC00000, True, &HE0FEFE
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &HC00000, True, &HE0FEFE
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &HC00000, True, &HE0FEFE
    End If
    Set Rst = Nothing
    Me.MousePointer = vbDefault
End Sub



Function fGenerarConsulta(Optional Tipo As Integer = 1) As String
    '===================================================================================================
    'Creado :  03/01/11 Por: Johan Castro
    'Propósito: Definir las consultas de todos los formatos del analisis de cuenta
    '
    'Entradas:  Tipo: Indica el tipo de formato
    '
    'Resultados: Sentencia SQL
    '
    'Otros:
    '
    'Modificado :
    
    '===================================================================================================


    'LEYENDA:
    'SI: Saldos Iniciales
    'MP: Movimientos del Periodo
    'SM: Sumas del Mayor
    'SA: Saldos Finales Al



    Dim nSQL As String
    Dim Rst As New ADODB.Recordset
    Dim nSQLIdLibro As String
    Dim nSQLTipoPersona As String
    Dim nSQLAjuste  As String '--sentencia sql para considera los registros del diario se ajuste por diferencia de cambio
    Dim nSQLCierre As String '--sentencia sql para no mostrar el cierre
    
    
    
    Dim nSQLCampos As String '--Relacion de campos a mostrar, obtenido de tabla: con_formatostipodet
    
    
    'muestra barra de progreso
    Frame4.Left = 3120
    Frame4.Top = 3930
    Frame4.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Value = 0
    '--Reiniciar la grilla
    Fg1.Rows = Fg1.FixedRows
    '--Refrescar
    DoEvents
   
    '--obtener el orden de presentacion de los campos
    nSQLCampos = fSetearCuadriculaColumna(xCon, 1, 2)
    '--verificar si hay campos seleccionados para mostrar el reporte
    If nSQLCampos = "" Then Exit Function

    
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " AND (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) "
    '-----------------------------------------------
    nSQLCierre = " AND (con_diario.idmes<>13) "
    '-----------------------------------------------
    
    
    DoEvents
    
    '**********************************************************************************************************************
    nSQL = "SELECT con_planctas.cuenta AS ctanum, con_planctas.descripcion as ctadesc,  " _
        + vbCr + " IIF(((IIF(SaldosIni.Deb Is Null,0,SaldosIni.Deb))-(IIF(SaldosIni.Hab Is Null,0,SaldosIni.Hab)))>0,((IIF(SaldosIni.Deb Is Null,0,SaldosIni.Deb))-(IIF(SaldosIni.Hab Is Null,0,SaldosIni.Hab))),0) AS SIDeb, " _
        + vbCr + " IIF(((IIF(SaldosIni.Hab Is Null,0,SaldosIni.Hab))-(IIF(SaldosIni.Deb Is Null,0,SaldosIni.Deb)))>0,((IIF(SaldosIni.Hab Is Null,0,SaldosIni.Hab))-(IIF(SaldosIni.Deb Is Null,0,SaldosIni.Deb))),0) AS SIHab, " _
        + vbCr + " IIF(MovPeriodo.Deb Is Null,0,MovPeriodo.Deb) AS MPDeb, " _
        + vbCr + " IIF(MovPeriodo.Hab Is Null,0,MovPeriodo.Hab) AS MPHab, " _
        + vbCr + " [SIDeb]+[MPDeb] AS SMDeb, " _
        + vbCr + " [SIHab]+[MPHab] AS SMHab, " _
        + vbCr + " IIF((SMDeb-SMHab)>0,(SMDeb-SMHab),0) AS SADeb, " _
        + vbCr + " IIF((SMHab-SMDeb)>0,(SMHab-SMDeb),0) AS SAHab, " _
        + vbCr + " con_planctas.iddes,con_planctas.iddes2,con_planctas.id AS IdCta, mae_bancos_2.numruc AS bcoruc, mae_bancos_2.descripcion AS bconombre, mae_banconumcta.numcue AS bcoctacte, mae_moneda_2.simbolo AS bcomon, mae_moneda_2.id as tmonsun "
    
    
    '**********************************************************************************************************************
    '--movimientos del periodo
    nSQL = nSQL _
        + vbCr + " FROM ((((mae_banconumcta LEFT JOIN mae_moneda AS mae_moneda_2 ON mae_banconumcta.idmon = mae_moneda_2.id) LEFT JOIN mae_bancos AS mae_bancos_2 ON mae_banconumcta.idban = mae_bancos_2.id) RIGHT JOIN (con_planctas INNER JOIN con_conceptodet ON con_planctas.id = con_conceptodet.idref) ON mae_banconumcta.idcuen = con_planctas.id) " _
        + vbCr + " LEFT JOIN " _
        + vbCr + " ( " _
        + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, "
        
        '--expresadoe en moneda nacional
        If NulosN(TxtIdMon) = 1 Then
            nSQL = nSQL _
                + vbCr + " Sum(IIF(con_diario.idmon=2,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS Deb, " _
                + vbCr + " Sum(IIF(con_diario.idmon=2,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS Hab "
        '--expresado en moneda extranjera
        ElseIf NulosN(TxtIdMon) = 2 Then
            nSQL = nSQL _
                + vbCr + " Sum(IIF(con_diario.idmon=1,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS Deb, " _
                + vbCr + " Sum(IIF(con_diario.idmon=1,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As Hab "
        '--expresado en moneda origen
        Else
            nSQL = nSQL _
                + vbCr + " Sum(IIF(con_diario.idmon=1,con_diario.impdebsol,con_diario.impdebdol)) AS Deb, " _
                + vbCr + " Sum(IIF(con_diario.idmon=1,con_diario.imphabsol,con_diario.imphabdol)) AS Hab "
        End If
        
        nSQL = nSQL _
        + vbCr + " FROM ((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) INNER JOIN con_conceptodet ON con_diario.idcue = con_conceptodet.idref " _
        + vbCr + " WHERE (con_conceptodet.idcpto=1) and (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " & nSQLIdLibro & nSQLAjuste _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
        + vbCr + " ORDER BY con_planctas.cuenta " _
        + vbCr + " ) AS MovPeriodo ON con_planctas.id = MovPeriodo.IdCta) " _
        + vbCr + " LEFT JOIN "
        
    '**********************************************************************************************************************
    '--saldos iniciales
    nSQL = nSQL _
        + vbCr + " (SELECT SaldosIniDet.IdCta,SaldosIniDet.cuenta,SaldosIniDet.descripcion,sum(SaldosIniDet.impdeb) as deb,sum(SaldosIniDet.imphab) as hab " _
        + vbCr + " FROM ( " _
        + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion,iif(con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) as tipcam, "
        
        '--expresadoe en moneda nacional
        If NulosN(TxtIdMon) = 1 Then
            nSQL = nSQL _
                + vbCr + " IIF(con_diario.idmon=2,IIF(tipcam=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS ImpDeb, " _
                + vbCr + " IIF(con_diario.idmon=2,IIF(tipcam=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS ImpHab "
        '--expresado en moneda extranjera
        ElseIf NulosN(TxtIdMon) = 2 Then
            nSQL = nSQL _
                + vbCr + " IIF(con_diario.idmon=1,IIF(tipcam=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/tipcam),con_diario.impdebdol) AS ImpDeb, " _
                + vbCr + " IIF(con_diario.idmon=1,IIF(tipcam=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/tipcam),con_diario.imphabdol) As ImpHab "
        '--expresado en moneda origen
        '--el saldo inicial se mostrara segun su moneda origen; Si la operacion esta en otra moneda, esta operacion se transformara utilizando el T/C.
        Else
            nSQL = nSQL _
                + vbCr + " IIF(con_diario.idmon=1,con_diario.impdebsol,con_diario.impdebdol) AS DebReal, " _
                + vbCr + " IIF(con_diario.idmon=1,con_diario.imphabsol,con_diario.imphabdol) AS HabReal, " _
                + vbCr + " IIF(con_diario.idmon=2,IIF(tipcam=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS Debsol, " _
                + vbCr + " IIF(con_diario.idmon=2,IIF(tipcam=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS Habsol, " _
                + vbCr + " IIF(con_diario.idmon=1,IIF(tipcam=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/tipcam),con_diario.impdebdol) AS Debdol, " _
                + vbCr + " IIF(con_diario.idmon=1,IIF(tipcam=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/tipcam),con_diario.imphabdol) As Habdol , " _
                + vbCr + " con_diario.idmon AS monope, mae_banconumcta.idmon AS moncta, " _
                + vbCr + " IIF(monope=moncta or moncta is null, debreal,IIF(moncta=1, debsol,debdol  ) ) as ImpDeb, " _
                + vbCr + " IIF(monope = moncta Or moncta Is Null, habreal, IIF(moncta = 1, habsol, habdol)) As ImpHab "
        End If
        
        nSQL = nSQL _
        + vbCr + " FROM (((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) INNER JOIN con_conceptodet ON con_diario.idcue = con_conceptodet.idref) LEFT JOIN mae_banconumcta ON con_diario.idcue = mae_banconumcta.idcuen " _
        + vbCr + " WHERE (con_conceptodet.idcpto=1) and ((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')) " & nSQLIdLibro & nSQLAjuste & nSQLCierre _
        + vbCr + " ) AS SaldosIniDet" _
        + vbCr + " GROUP BY SaldosIniDet.IdCta,SaldosIniDet.cuenta,SaldosIniDet.descripcion  " _
        + vbCr + " ORDER BY SaldosIniDet.cuenta ) as SaldosIni " _

    '**********************************************************************************************************************
    '--filtro para el where
    nSQLIdLibro = nSQLIdLibro & nSQLAjuste & " AND (  (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))  OR  (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')  OR  (con_diario.fchasi) is null  )"

    nSQL = nSQL _
        + vbCr + " ON con_planctas.id = SaldosIni.IdCta " _
        + vbCr + " WHERE (((con_conceptodet.idcpto)=1)) " _
        + vbCr + " ORDER BY con_planctas.cuenta asc"
       
    Me.MousePointer = vbHourglass
    
      
    DoEvents
    
    '--armar la sentencia SQL
    nSQL = "Select " & nSQLCampos & _
            vbCr + " from ( " _
            + vbCr + nSQL _
            + vbCr + ") as consulta "
    
    '--ejecutar la consulta
    RST_Busq Rst, nSQL, xCon
    
    '--Salir si hay error en la consulta
    If Rst.State = 0 Then GoTo LaCague
    
   '**************************************************************************************************
    '--obtener las posiciones de las columnas
    Dim mColCampo As Integer
    Dim mCol As Integer '--indica la posicion del campo
    '--definir el array por defecto a 15 campos
    Dim ArrCampos(15) As Integer
    '--posicionar la variable a la primera columna
    mCol = 0
    '--obtener la posicion de los campos de la consulta en el arreglo
    For mColCampo = 0 To Rst.Fields.Count - 1
        Select Case LCase(Rst.Fields(mColCampo).Name)
            Case "sideb":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "sihab":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "mpdeb":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "mphab":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "smdeb":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "smhab":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "sadeb":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "sahab":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            
        End Select
        
    Next mColCampo
    '**************************************************************************************************
    
    If Rst.RecordCount <> 0 Then ProgressBar1.Max = Rst.RecordCount
    
    '--proceder a cargar los datos
    Do While Not Rst.EOF
        DoEvents
''        ProgressBar1.Value = Rst.Bookmark
        
        '-----------------------------------------------
        Fg1.Rows = Fg1.Rows + 1
        
        For mCol = 0 To Rst.Fields.Count - 1
        
            Select Case LCase(Rst.Fields(mCol).Name)
                
                Case "sideb", "sihab", "mpdeb", "mphab", "smdeb", "smhab", "sadeb", "sahab"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_MONTO)
                
                Case Else
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
            End Select
            
        Next mCol

            
        Rst.MoveNext
    Loop

    '**************************************************************************************************
    '--verificamos si se suman las columnas
    If ArrCampos(0) <> 0 Then
            
        '--sumando las columnas
        Fg1.Rows = Fg1.Rows + 1
        'ArrCampos(1) - 3 < 0:: Verificamos que la columna exista
        FORMATO_CELDA Fg1, Fg1.Rows - 1, IIf(ArrCampos(1) - 3 < 0, 1, ArrCampos(1) - 3), &H800000, False, , "TOTAL ==>"
        
        For mCol = 0 To UBound(ArrCampos())
            If ArrCampos(mCol) <> 0 Then
                FORMATO_CELDA Fg1, Fg1.Rows - 1, ArrCampos(mCol) + 1, &H800000, False, , Format(GRID_SUMAR_COL(Fg1, ArrCampos(mCol) + 1), FORMAT_MONTO)
            End If
        Next mCol
        
    End If
    '**************************************************************************************************


LaCague:

    Set Rst = Nothing
    '--Ocultar barra de progreso
    Frame4.Visible = False
    '--restablecer cursor
    Me.MousePointer = vbDefault
    
End Function






Function fGenerarConsultaBK(Optional Tipo As Integer = 1) As String
    '===================================================================================================
    'Creado :  03/01/11 Por: Johan Castro
    'Propósito: Definir las consultas de todos los formatos del analisis de cuenta
    '
    'Entradas:  Tipo: Indica el tipo de formato
    '
    'Resultados: Sentencia SQL
    '
    'Otros:
    '
    'Modificado :
    
    '===================================================================================================


    'LEYENDA:
    'SI: Saldos Iniciales
    'MP: Movimientos del Periodo
    'SM: Sumas del Mayor
    'SA: Saldos Finales Al



    Dim nSQL As String
    Dim Rst As New ADODB.Recordset
    Dim nSQLIdLibro As String
    Dim nSQLTipoPersona As String
    Dim nSQLAjuste  As String '--sentencia sql para considera los registros del diario se ajuste por diferencia de cambio
    Dim nSQLCierre As String '--sentencia sql para no mostrar el cierre
    
    
    
    Dim nSQLCampos As String '--Relacion de campos a mostrar, obtenido de tabla: con_formatostipodet
    
    
    'muestra barra de progreso
    Frame4.Left = 3120
    Frame4.Top = 3930
    Frame4.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Value = 0
    '--Reiniciar la grilla
    Fg1.Rows = Fg1.FixedRows
    '--Refrescar
    DoEvents
   
    '--obtener el orden de presentacion de los campos
    nSQLCampos = fSetearCuadriculaColumna(xCon, 1, 2)
    '--verificar si hay campos seleccionados para mostrar el reporte
    If nSQLCampos = "" Then Exit Function

    
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " AND (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) "
    '-----------------------------------------------
    nSQLCierre = " AND (con_diario.idmes<>13) "
    '-----------------------------------------------
    
    
    DoEvents
    
    '**********************************************************************************************************************
    nSQL = "SELECT con_planctas.cuenta AS ctanum, con_planctas.descripcion as ctadesc,  " _
        + vbCr + " IIF(((IIF(SaldosIni.Deb Is Null,0,SaldosIni.Deb))-(IIF(SaldosIni.Hab Is Null,0,SaldosIni.Hab)))>0,((IIF(SaldosIni.Deb Is Null,0,SaldosIni.Deb))-(IIF(SaldosIni.Hab Is Null,0,SaldosIni.Hab))),0) AS SIDeb, " _
        + vbCr + " IIF(((IIF(SaldosIni.Hab Is Null,0,SaldosIni.Hab))-(IIF(SaldosIni.Deb Is Null,0,SaldosIni.Deb)))>0,((IIF(SaldosIni.Hab Is Null,0,SaldosIni.Hab))-(IIF(SaldosIni.Deb Is Null,0,SaldosIni.Deb))),0) AS SIHab, " _
        + vbCr + " IIF(MovPeriodo.Deb Is Null,0,MovPeriodo.Deb) AS MPDeb, " _
        + vbCr + " IIF(MovPeriodo.Hab Is Null,0,MovPeriodo.Hab) AS MPHab, " _
        + vbCr + " [SIDeb]+[MPDeb] AS SMDeb, " _
        + vbCr + " [SIHab]+[MPHab] AS SMHab, " _
        + vbCr + " IIF((SMDeb-SMHab)>0,(SMDeb-SMHab),0) AS SADeb, " _
        + vbCr + " IIF((SMHab-SMDeb)>0,(SMHab-SMDeb),0) AS SAHab, " _
        + vbCr + " con_planctas.iddes,con_planctas.iddes2,con_planctas.id AS IdCta, mae_bancos_2.numruc AS bcoruc, mae_bancos_2.descripcion AS bconombre, mae_banconumcta.numcue AS bcoctacte, mae_moneda_2.simbolo AS bcomon, mae_moneda_2.id as tmonsun "
    
    
    '**********************************************************************************************************************
    '--movimientos del periodo
    nSQL = nSQL _
        + vbCr + " FROM ((((mae_banconumcta LEFT JOIN mae_moneda AS mae_moneda_2 ON mae_banconumcta.idmon = mae_moneda_2.id) LEFT JOIN mae_bancos AS mae_bancos_2 ON mae_banconumcta.idban = mae_bancos_2.id) RIGHT JOIN (con_planctas INNER JOIN con_conceptodet ON con_planctas.id = con_conceptodet.idref) ON mae_banconumcta.idcuen = con_planctas.id) " _
        + vbCr + " LEFT JOIN " _
        + vbCr + " ( " _
        + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, "
        
        '--expresadoe en moneda nacional
        If NulosN(TxtIdMon) = 1 Then
            nSQL = nSQL _
                + vbCr + " Sum(IIF(con_diario.idmon=2,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS Deb, " _
                + vbCr + " Sum(IIF(con_diario.idmon=2,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS Hab "
        '--expresado en moneda extranjera
        ElseIf NulosN(TxtIdMon) = 2 Then
            nSQL = nSQL _
                + vbCr + " Sum(IIF(con_diario.idmon=1,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS Deb, " _
                + vbCr + " Sum(IIF(con_diario.idmon=1,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As Hab "
        '--expresado en moneda origen
        Else
            nSQL = nSQL _
                + vbCr + " Sum(IIF(con_diario.idmon=1,con_diario.impdebsol,con_diario.impdebdol)) AS Deb, " _
                + vbCr + " Sum(IIF(con_diario.idmon=1,con_diario.imphabsol,con_diario.imphabdol)) AS Hab "
        End If
        
        nSQL = nSQL _
        + vbCr + " FROM ((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) INNER JOIN con_conceptodet ON con_diario.idcue = con_conceptodet.idref " _
        + vbCr + " WHERE (con_conceptodet.idcpto=1) and (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " & nSQLIdLibro & nSQLAjuste _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
        + vbCr + " ORDER BY con_planctas.cuenta " _
        + vbCr + " ) AS MovPeriodo ON con_planctas.id = MovPeriodo.IdCta) " _
        + vbCr + " LEFT JOIN "
        
    '**********************************************************************************************************************
    '--saldos iniciales
    nSQL = nSQL _
        + vbCr + " ( " _
        + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, "
        
        '--expresadoe en moneda nacional
        If NulosN(TxtIdMon) = 1 Then
            nSQL = nSQL _
                + vbCr + " Sum(IIF(con_diario.idmon=2,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS Deb, " _
                + vbCr + " Sum(IIF(con_diario.idmon=2,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS Hab "
        '--expresado en moneda extranjera
        ElseIf NulosN(TxtIdMon) = 2 Then
            nSQL = nSQL _
                + vbCr + " Sum(IIF(con_diario.idmon=1,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS Deb, " _
                + vbCr + " Sum(IIF(con_diario.idmon=1,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As Hab "
        '--expresado en moneda origen
        '--el saldo inicial se mostrara segun su moneda origen; Si la operacion esta en otra moneda, esta operacion se transformara utilizando el T/C.
        Else
            nSQL = nSQL _
                + vbCr + " Sum(IIF(con_diario.idmon=1,con_diario.impdebsol,con_diario.impdebdol)) AS DebReal, " _
                + vbCr + " Sum(IIF(con_diario.idmon=1,con_diario.imphabsol,con_diario.imphabdol)) AS HabReal, " _
                + vbCr + " Sum(IIF(con_diario.idmon=2,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS Debsol, " _
                + vbCr + " Sum(IIF(con_diario.idmon=2,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS Habsol, " _
                + vbCr + " Sum(IIF(con_diario.idmon=1,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS Debdol, " _
                + vbCr + " Sum(IIF(con_diario.idmon=1,IIF(IIF( con_diario.aplicatc=-1,con_diario.tc,IIF(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(IIF( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As Habdol , " _
                + vbCr + " con_diario.idmon AS monope, mae_banconumcta.idmon AS moncta, " _
                + vbCr + " IIF(monope=moncta or moncta is null, debreal,IIF(moncta=1, debsol,debdol  ) ) as Deb, " _
                + vbCr + " IIF(monope = moncta Or moncta Is Null, habreal, IIF(moncta = 1, habsol, habdol)) As Hab "
        End If
        
        nSQL = nSQL _
        + vbCr + " FROM (((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) INNER JOIN con_conceptodet ON con_diario.idcue = con_conceptodet.idref) LEFT JOIN mae_banconumcta ON con_diario.idcue = mae_banconumcta.idcuen " _
        + vbCr + " WHERE (con_conceptodet.idcpto=1) and ((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')) " & nSQLIdLibro & nSQLAjuste & nSQLCierre _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion, con_diario.idmon, mae_banconumcta.idmon " _
        + vbCr + " ORDER BY con_planctas.cuenta " _
        + vbCr + " ) AS SaldosIni "
        
    '**********************************************************************************************************************
    '--filtro para el where
    nSQLIdLibro = nSQLIdLibro & nSQLAjuste & " AND (  (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))  OR  (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')  OR  (con_diario.fchasi) is null  )"

    nSQL = nSQL _
        + vbCr + " ON con_planctas.id = SaldosIni.IdCta " _
        + vbCr + " WHERE (((con_conceptodet.idcpto)=1)) " _
        + vbCr + " ORDER BY con_planctas.cuenta "
       
    Me.MousePointer = vbHourglass
    
      
    DoEvents
    
    '--armar la sentencia SQL
    nSQL = "Select " & nSQLCampos & _
            vbCr + " from ( " _
            + vbCr + nSQL _
            + vbCr + ") as consulta "
    
    '--ejecutar la consulta
    RST_Busq Rst, nSQL, xCon
    
    '--Salir si hay error en la consulta
    If Rst.State = 0 Then GoTo LaCague
    
    '--obtener las posiciones de las columnas
    Dim mColCampo As Integer
    Dim mCol As Integer '--indica la posicion del campo
    
    If Rst.RecordCount <> 0 Then ProgressBar1.Max = Rst.RecordCount
    
    '--proceder a cargar los datos
    Do While Not Rst.EOF
        DoEvents
''        ProgressBar1.Value = Rst.Bookmark
        
        '-----------------------------------------------
        Fg1.Rows = Fg1.Rows + 1
        
        For mCol = 0 To Rst.Fields.Count - 1
        
            Select Case LCase(Rst.Fields(mCol).Name)
                
                Case "sideb", "sihab", "mpdeb", "mphab", "smdeb", "smhab", "sadeb", "sahab"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_MONTO)
                
                Case Else
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
            End Select
            
        Next mCol

            
        Rst.MoveNext
    Loop


LaCague:

    Set Rst = Nothing
    '--Ocultar barra de progreso
    Frame4.Visible = False
    '--restablecer cursor
    Me.MousePointer = vbDefault
    
End Function




