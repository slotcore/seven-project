VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmAgregaDoc 
   Caption         =   "Ingreso de Documentos de Apertura"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   13080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmIngreso 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H00800000&
      Height          =   3765
      Left            =   1380
      TabIndex        =   8
      Top             =   1290
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox TxtNumReg 
         Height          =   300
         Left            =   1650
         MaxLength       =   8
         TabIndex        =   15
         Text            =   "TxtNumReg"
         Top             =   2220
         Width           =   1485
      End
      Begin VB.CommandButton Command4 
         Height          =   240
         Left            =   2880
         Picture         =   "FrmAgregaDoc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1950
         Width           =   240
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Cancelar"
         Height          =   405
         Left            =   5250
         TabIndex        =   21
         Top             =   3150
         Width           =   1305
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Height          =   405
         Left            =   3900
         TabIndex        =   20
         Top             =   3150
         Width           =   1305
      End
      Begin VB.TextBox TxtNumDocRef 
         Height          =   300
         Left            =   1650
         TabIndex        =   13
         Text            =   "TxtNumDocRef"
         Top             =   1620
         Width           =   1905
      End
      Begin VB.TextBox TxtHaberSol 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   5400
         TabIndex        =   19
         Text            =   "TxtHaberSol"
         Top             =   2820
         Width           =   1455
      End
      Begin VB.TextBox TxtDebeSol 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   3870
         TabIndex        =   18
         Text            =   "TxtDebeSol"
         Top             =   2820
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Height          =   240
         Left            =   2460
         Picture         =   "FrmAgregaDoc.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   750
         Width           =   240
      End
      Begin VB.TextBox TxtIdMoneda 
         Height          =   300
         Left            =   1650
         TabIndex        =   10
         Text            =   "TxtIdMoneda"
         Top             =   720
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
         Height          =   240
         Left            =   2460
         Picture         =   "FrmAgregaDoc.frx":0264
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   450
         Width           =   240
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
         Height          =   300
         Left            =   1650
         TabIndex        =   12
         Top             =   1320
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
      End
      Begin VB.TextBox TxtTipDoc 
         Height          =   300
         Left            =   1650
         TabIndex        =   9
         Text            =   "TxtTipDoc"
         Top             =   420
         Width           =   1065
      End
      Begin VB.TextBox TxtHaberDol 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1800
         TabIndex        =   17
         Text            =   "TxtHaberDol"
         Top             =   2820
         Width           =   1455
      End
      Begin VB.TextBox TxtDebeDol 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   270
         TabIndex        =   16
         Text            =   "TxtDebeDol"
         Top             =   2820
         Width           =   1455
      End
      Begin VB.TextBox TxtNumDoc 
         Height          =   300
         Left            =   1650
         TabIndex        =   11
         Text            =   "TxtNumDoc"
         Top             =   1020
         Width           =   1905
      End
      Begin VB.TextBox TxtCuenta 
         Height          =   300
         Left            =   1650
         TabIndex        =   14
         Text            =   "TxtCuenta"
         Top             =   1920
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº Registro"
         Height          =   195
         Index           =   10
         Left            =   270
         TabIndex        =   43
         Top             =   2250
         Width           =   810
      End
      Begin VB.Label LblIdCuenta 
         Caption         =   "LblIdCuenta"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5220
         TabIndex        =   42
         Top             =   1650
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº Cuenta"
         Height          =   195
         Index           =   9
         Left            =   270
         TabIndex        =   41
         Top             =   1950
         Width           =   735
      End
      Begin VB.Label LblCuenta 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCuenta"
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
         Left            =   3180
         TabIndex        =   40
         Top             =   1920
         Width           =   6975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agregando Documentos"
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
         Left            =   210
         TabIndex        =   38
         Top             =   90
         Width           =   2040
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Left            =   30
         Top             =   30
         Width           =   10395
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   10440
         X2              =   10440
         Y1              =   30
         Y2              =   3750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   10440
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   10440
         Y1              =   3750
         Y2              =   3750
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
         Left            =   2760
         TabIndex        =   34
         Top             =   720
         Width           =   4245
      End
      Begin VB.Label LblTipDoc 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblTipDoc"
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
         Left            =   2760
         TabIndex        =   33
         Top             =   420
         Width           =   4245
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Doc. Ref."
         Height          =   255
         Index           =   8
         Left            =   270
         TabIndex        =   32
         Top             =   1650
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Haber S/."
         Height          =   195
         Index           =   7
         Left            =   5400
         TabIndex        =   31
         Top             =   2610
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Debe S/."
         Height          =   195
         Index           =   6
         Left            =   3870
         TabIndex        =   30
         Top             =   2610
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Haber US $"
         Height          =   195
         Index           =   5
         Left            =   1830
         TabIndex        =   29
         Top             =   2610
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Debe US $"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   28
         Top             =   2610
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   26
         Top             =   750
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   25
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Fch. Emi."
         Height          =   255
         Index           =   1
         Left            =   270
         TabIndex        =   23
         Top             =   1350
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Documento"
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   22
         Top             =   1050
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Buscar Datos"
      Height          =   420
      Left            =   9585
      TabIndex        =   7
      Top             =   30
      Width           =   1950
   End
   Begin VB.CommandButton CmdBusCli 
      Height          =   240
      Left            =   1335
      Picture         =   "FrmAgregaDoc.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   165
      Width           =   240
   End
   Begin VB.TextBox TxtCliente 
      Height          =   300
      Left            =   1620
      TabIndex        =   3
      Text            =   "TxtCliente"
      Top             =   135
      Width           =   7815
   End
   Begin VB.TextBox TxtIdCliente 
      Height          =   300
      Left            =   660
      TabIndex        =   1
      Text            =   "TxtIdCliente"
      Top             =   135
      Width           =   945
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   5325
      Left            =   30
      TabIndex        =   0
      Top             =   495
      Width           =   13005
      _cx             =   22939
      _cy             =   9393
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
      BackColor       =   14744827
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14744827
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmAgregaDoc.frx":04C8
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
      Height          =   780
      Left            =   30
      TabIndex        =   5
      Top             =   5760
      Width           =   13020
      Begin VB.CommandButton CmdDel 
         Caption         =   "Eliminar Documento"
         Height          =   510
         Left            =   4080
         TabIndex        =   37
         Top             =   195
         Width           =   1365
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Agregar Documento"
         Height          =   510
         Left            =   1260
         TabIndex        =   36
         Top             =   195
         Width           =   1365
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Modificar Documento"
         Height          =   510
         Left            =   2670
         TabIndex        =   35
         Top             =   195
         Width           =   1365
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Salir"
         Height          =   510
         Left            =   9885
         TabIndex        =   6
         Top             =   195
         Width           =   1365
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Left            =   45
      TabIndex        =   2
      Top             =   165
      Width           =   480
   End
End
Attribute VB_Name = "FrmAgregaDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim QueHace As Integer
Dim SeEjecuto  As Boolean

Private Sub CmdAdd_Click()
    If NulosN(TxtIdCliente.Text) = 0 Then
        MsgBox "No ha especificado el cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdCliente.SetFocus
        Exit Sub
    End If
    
    QueHace = 1
    Fg1.Enabled = False
    TxtIdCliente.Enabled = False
    TxtCliente.Enabled = False
    Frame1.Enabled = False
    
    FrmIngreso.Left = 1380
    FrmIngreso.Top = 1320
    Blanquea
    FrmIngreso.Visible = True
    TxtTipDoc.SetFocus
End Sub


Private Sub CmdBusCli_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Cliente":       xCampos(0, 1) = "nombre":           xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "abrev":            xCampos(1, 2) = "800":          xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Documento":  xCampos(2, 1) = "numruc":           xCampos(2, 2) = "1500":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Tipo Empresa":  xCampos(3, 1) = "tipemp":           xCampos(3, 2) = "1500":         xCampos(3, 3) = "C"
    
    xform.SQLCad = "SELECT mae_cliente.nombre, mae_dociden.abrev, mae_tipoempresa.descripcion AS tipemp, mae_cliente.numruc, " _
        & " mae_cliente.id FROM (mae_cliente LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id) LEFT JOIN mae_tipoempresa " _
        & " ON mae_cliente.tipper = mae_tipoempresa.id Where (((mae_cliente.activo) = -1)) ORDER BY mae_cliente.nombre"
    
    xform.Titulo = "Buscando Cliente"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtCliente.Text = xRs("nombre")
            TxtIdCliente.Text = xRs("id")
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDel_Click()
    If Fg1.Rows = 1 Then Exit Sub
    Dim Rpta As Integer
    
    If NulosN(Fg1.TextMatrix(Fg1.Row, 13)) <> 999 Then
        MsgBox "Solo puede eliminar registros del apertura de la cuenta corriente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("Esta seguro de eliminar el documento seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM var_analisisctacte WHERE idlib = 999 AND idope = " & Fg1.TextMatrix(Fg1.Row, 12) & ""
        Fg1.RemoveItem Fg1.Row
    End If
    
End Sub

Private Sub CmdGrabar_Click()
    Dim xCampos(14, 4) As String
    Dim A As Integer
    Dim xImpTc As Double
    
On Error GoTo Lacague
    
    xCon.BeginTrans
    
    Dim xIdDoc As Double
    Dim Rst As New ADODB.Recordset
    Dim xTc As Double
    
    xTc = HallaTipoCambio(TxtFchEmi.Valor, 2, Venta, xCon)
    
    If QueHace = 1 Then
        RST_Busq Rst, "SELECT * FROM var_analisisctacte WHERE idlib = 999 ORDER BY idope", xCon
        
        If Rst.RecordCount = 0 Then
            xIdDoc = 1
        Else
            Rst.MoveLast
            xIdDoc = NulosN(Rst("idope")) + 1
        End If
        Set Rst = Nothing
    Else
        xIdDoc = NulosN(Fg1.TextMatrix(Fg1.Row, 12))
        xCon.Execute "DELETE * FROM var_analisisctacte WHERE idlib = 999  AND idope = " & xIdDoc & ""
    End If
    
    xCampos(0, 0) = "idlib":         xCampos(0, 1) = 999:                          xCampos(0, 2) = "S":     xCampos(0, 3) = "N":     xCampos(0, 4) = ""
    xCampos(1, 0) = "idope":         xCampos(1, 1) = xIdDoc:                       xCampos(1, 2) = "S":     xCampos(1, 3) = "N":     xCampos(1, 4) = ""
    xCampos(2, 0) = "numdocref":     xCampos(2, 1) = TxtNumDocRef:                 xCampos(2, 2) = "N":     xCampos(2, 3) = "C":     xCampos(2, 4) = "No ha especificado el numero de documento de referencia"
    xCampos(3, 0) = "idcli":         xCampos(3, 1) = TxtIdCliente.Text:            xCampos(3, 2) = "S":     xCampos(3, 3) = "N":     xCampos(3, 4) = ""
    xCampos(4, 0) = "idtipdoc":      xCampos(4, 1) = TxtTipDoc.Text:               xCampos(4, 2) = "S":     xCampos(4, 3) = "N":     xCampos(4, 4) = "No ha especificado el tipo de documento"
    xCampos(5, 0) = "numdoc":        xCampos(5, 1) = TxtNumDoc.Text:               xCampos(5, 2) = "S":     xCampos(5, 3) = "C":     xCampos(5, 4) = "No ha especificado el numero de documento"
    xCampos(6, 0) = "fchemi":        xCampos(6, 1) = TxtFchEmi.Valor:              xCampos(6, 2) = "S":     xCampos(6, 3) = "F":     xCampos(6, 4) = "No ha especificado la fecha de emision"
    xCampos(7, 0) = "idmon":         xCampos(7, 1) = TxtIdMoneda.Text:             xCampos(7, 2) = "S":     xCampos(7, 3) = "N":     xCampos(7, 4) = "No ha especificado el tipo de moneda"
    xCampos(8, 0) = "imptc":         xCampos(8, 1) = xTc:                          xCampos(8, 2) = "S":     xCampos(8, 3) = "N":     xCampos(8, 4) = ""
    xCampos(9, 0) = "debesol":       xCampos(9, 1) = NulosN(TxtDebeSol.Text):      xCampos(9, 2) = "N":     xCampos(9, 3) = "N":     xCampos(9, 4) = ""
    xCampos(10, 0) = "habersol":     xCampos(10, 1) = NulosN(TxtHaberSol.Text):    xCampos(10, 2) = "N":    xCampos(10, 3) = "N":    xCampos(10, 4) = ""
    xCampos(11, 0) = "debedol":      xCampos(11, 1) = NulosN(TxtDebeDol.Text):     xCampos(11, 2) = "N":    xCampos(11, 3) = "N":    xCampos(11, 4) = ""
    xCampos(12, 0) = "haberdol":     xCampos(12, 1) = NulosN(TxtHaberDol.Text):    xCampos(12, 2) = "N":    xCampos(12, 3) = "N":    xCampos(12, 4) = ""
    xCampos(13, 0) = "idcue":        xCampos(13, 1) = NulosN(LblIdCuenta.Caption): xCampos(13, 2) = "N":    xCampos(13, 3) = "N":    xCampos(13, 4) = ""
    xCampos(14, 0) = "numreg":       xCampos(14, 1) = TxtNumReg.Text:              xCampos(14, 2) = "S":    xCampos(14, 3) = "C":    xCampos(14, 4) = "No ha especificado el numero de registro"
    
    If EscribirNuevoRegistro(xCampos, "var_analisisctacte", xCon) = False Then
        xCon.RollbackTrans
        Exit Sub
    End If
    
    xCon.CommitTrans
    MsgBox "El documento se guardo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        
    Fg1.Enabled = True
    TxtIdCliente.Enabled = True
    TxtCliente.Enabled = True
    Frame1.Enabled = True
    QueHace = 3
    FrmIngreso.Visible = False
    Command1_Click
    Exit Sub
    
Lacague:
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Exit Sub
End Sub

Private Sub CmdSalir_Click()
    Fg1.Enabled = True
    TxtIdCliente.Enabled = True
    TxtCliente.Enabled = True
    Frame1.Enabled = True
    QueHace = 3
    FrmIngreso.Visible = False
End Sub

'Private Sub CmdSave_Click()
'    If NulosN(TxtIdCliente.Text) = 0 Then
'        MsgBox "No ha especificado el cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtIdCliente.SetFocus
'        Exit Sub
'    End If
'
'    If Fg1.Rows = 1 Then
'        MsgBox "Debe de ingresar al menos un documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Fg1.SetFocus
'        Exit Sub
'    End If
'
'    Dim xCampos(19, 4) As String
'    Dim A As Integer
'
'    For A = 1 To Fg1.Rows - 1
'        If NulosC(Fg1.TextMatrix(A, 1)) = "" Then
'            Fg1.RemoveItem (A)
'            A = A - 1
'        End If
'
'        If A = Fg1.Rows - 1 Then Exit For
'    Next A
'
'On Error GoTo Lacague
'
'    xCon.BeginTrans
'
'    Dim xIdDoc As Double
'    xIdDoc = HallaCodigoTabla("con_provicionesdetdoc_AP", xCon, "id")
'
'    For A = 1 To Fg1.Rows - 1
'        xCampos(0, 0) = "idpro":         xCampos(0, 1) = 2:                           xCampos(0, 2) = "S":     xCampos(0, 3) = "N":    xCampos(0, 4) = ""
'        xCampos(1, 0) = "idcue":         xCampos(1, 1) = Fg1.TextMatrix(A, 11):       xCampos(1, 2) = "S":     xCampos(1, 3) = "N":    xCampos(1, 4) = ""
'        xCampos(2, 0) = "id":            xCampos(2, 1) = xIdDoc:                      xCampos(2, 2) = "S":     xCampos(2, 3) = "N":    xCampos(2, 4) = ""
'        xCampos(3, 0) = "idmod":         xCampos(3, 1) = Fg1.TextMatrix(A, 12):       xCampos(3, 2) = "S":     xCampos(3, 3) = "N":    xCampos(3, 4) = "No ha especificado el modulo"
'        xCampos(4, 0) = "idtipper":      xCampos(4, 1) = 2:                           xCampos(4, 2) = "S":     xCampos(4, 3) = "N":    xCampos(4, 4) = ""
'        xCampos(5, 0) = "idclipro":      xCampos(5, 1) = TxtIdCliente.Text:           xCampos(5, 2) = "S":     xCampos(5, 3) = "N":    xCampos(5, 4) = "No ha especificado el id del cliente"
'        xCampos(6, 0) = "cliente":       xCampos(6, 1) = TxtCliente:                  xCampos(6, 2) = "S":     xCampos(6, 3) = "C":    xCampos(6, 4) = "No ha especificado el nombre del cliente"
'        xCampos(7, 0) = "numser":        xCampos(7, 1) = Fg1.TextMatrix(A, 4):        xCampos(7, 2) = "S":     xCampos(7, 3) = "C":    xCampos(7, 4) = "No ha especificado el numero de serie"
'        xCampos(8, 0) = "numdoc":        xCampos(8, 1) = Fg1.TextMatrix(A, 5):        xCampos(8, 2) = "S":     xCampos(8, 3) = "C":    xCampos(8, 4) = "No ha especificado el numero de documento"
'        xCampos(9, 0) = "fchemi":        xCampos(9, 1) = Fg1.TextMatrix(A, 6):        xCampos(9, 2) = "S":     xCampos(9, 3) = "F":    xCampos(9, 4) = "No ha especificado la fecha de emision"
'        xCampos(10, 0) = "idmon":        xCampos(10, 1) = Fg1.TextMatrix(A, 13):      xCampos(10, 2) = "S":    xCampos(10, 3) = "N":    xCampos(10, 4) = "No ha especificado la moneda"
'        xCampos(11, 0) = "tipdoc":       xCampos(11, 1) = Fg1.TextMatrix(A, 14):      xCampos(11, 2) = "S":    xCampos(11, 3) = "N":    xCampos(11, 4) = "No ha especificado el tipo de documento"
'        xCampos(12, 0) = "impdoc":       xCampos(12, 1) = Fg1.TextMatrix(A, 9):       xCampos(12, 2) = "S":    xCampos(12, 3) = "N":    xCampos(12, 4) = "No ha especificado el importe del documento"
'        xCampos(13, 0) = "impsal":       xCampos(13, 1) = Fg1.TextMatrix(A, 9):       xCampos(13, 2) = "S":    xCampos(13, 3) = "N":    xCampos(13, 4) = ""
'        xCampos(14, 0) = "idtipdocref":  xCampos(14, 1) = 4:                          xCampos(14, 2) = "S":    xCampos(14, 3) = "N":    xCampos(14, 4) = ""
'        xCampos(15, 0) = "numorden":     xCampos(15, 1) = Fg1.TextMatrix(A, 10):      xCampos(15, 2) = "S":    xCampos(15, 3) = "C":    xCampos(15, 4) = "No ha especificado el numero de documento de referencia"
'
'        If NulosN(Fg1.TextMatrix(A, 14)) = 3 Or NulosN(Fg1.TextMatrix(A, 14)) = 1 Or NulosN(Fg1.TextMatrix(A, 14)) = 8 Or NulosN(Fg1.TextMatrix(A, 14)) = 120 Then
'            If Fg1.TextMatrix(A, 13) = 2 Then
'                xCampos(16, 0) = "abonodol":       xCampos(16, 1) = Fg1.TextMatrix(A, 9):       xCampos(16, 2) = "N":    xCampos(16, 3) = "N":    xCampos(16, 4) = ""
'                xCampos(17, 0) = "cargodol":       xCampos(17, 1) = 0:                          xCampos(17, 2) = "N":    xCampos(17, 3) = "N":    xCampos(17, 4) = ""
'                xCampos(18, 0) = "abonosol":       xCampos(18, 1) = 0:                          xCampos(18, 2) = "N":    xCampos(18, 3) = "N":    xCampos(18, 4) = ""
'                xCampos(19, 0) = "cargosol":       xCampos(19, 1) = 0:                          xCampos(19, 2) = "N":    xCampos(19, 3) = "N":    xCampos(19, 4) = ""
'            Else
'                xCampos(16, 0) = "abonodol":       xCampos(16, 1) = 0:                          xCampos(16, 2) = "N":    xCampos(16, 3) = "N":    xCampos(16, 4) = ""
'                xCampos(17, 0) = "cargodol":       xCampos(17, 1) = 0:                          xCampos(17, 2) = "N":    xCampos(17, 3) = "N":    xCampos(17, 4) = ""
'                xCampos(18, 0) = "abonosol":       xCampos(18, 1) = Fg1.TextMatrix(A, 9):       xCampos(18, 2) = "N":    xCampos(18, 3) = "N":    xCampos(18, 4) = ""
'                xCampos(19, 0) = "cargosol":       xCampos(19, 1) = 0:                          xCampos(19, 2) = "N":    xCampos(19, 3) = "N":    xCampos(19, 4) = ""
'            End If
'        Else
'            If Fg1.TextMatrix(A, 13) = 2 Then
'                xCampos(16, 0) = "abonodol":       xCampos(16, 1) = 0:                          xCampos(16, 2) = "N":    xCampos(16, 3) = "N":    xCampos(16, 4) = ""
'                xCampos(17, 0) = "cargodol":       xCampos(17, 1) = Fg1.TextMatrix(A, 9):       xCampos(17, 2) = "N":    xCampos(17, 3) = "N":    xCampos(17, 4) = ""
'                xCampos(18, 0) = "abonosol":       xCampos(18, 1) = 0:                          xCampos(18, 2) = "N":    xCampos(18, 3) = "N":    xCampos(18, 4) = ""
'                xCampos(19, 0) = "cargosol":       xCampos(19, 1) = 0:                          xCampos(19, 2) = "N":    xCampos(19, 3) = "N":    xCampos(19, 4) = ""
'
'            Else
'                xCampos(16, 0) = "abonodol":       xCampos(16, 1) = 0:                          xCampos(16, 2) = "N":    xCampos(16, 3) = "N":    xCampos(16, 4) = ""
'                xCampos(17, 0) = "cargodol":       xCampos(17, 1) = 0:                          xCampos(17, 2) = "N":    xCampos(17, 3) = "N":    xCampos(17, 4) = ""
'                xCampos(18, 0) = "abonosol":       xCampos(18, 1) = 0:                          xCampos(18, 2) = "N":    xCampos(18, 3) = "N":    xCampos(18, 4) = ""
'                xCampos(19, 0) = "cargosol":       xCampos(19, 1) = Fg1.TextMatrix(A, 9):       xCampos(19, 2) = "N":    xCampos(19, 3) = "N":    xCampos(19, 4) = ""
'            End If
'        End If
'
'        If EscribirNuevoRegistro(xCampos, "con_provicionesdetdoc_AP", xCon) = False Then
'            xCon.RollbackTrans
'            Exit Sub
'        End If
'
'        xIdDoc = xIdDoc + 1
'    Next A
'    xCon.CommitTrans
'    MsgBox "Los documentos se guaradaron con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'
'    Fg1.Rows = 1
'    TxtIdCliente.Text = ""
'    TxtCliente.Text = ""
'    Exit Sub
'
'Lacague:
'    xCon.RollbackTrans
'    MsgBox "No se pudo guardar el registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'    Exit Sub
'End Sub

Private Sub Command1_Click()
    If NulosC(TxtIdCliente.Text) = "" Then
        MsgBox "No ha especificado un cliente, seleccione un cliente para utilizar esta opcion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdCliente.SetFocus
        Exit Sub
    End If
    BuscaDatosCliente NulosN(TxtIdCliente.Text)
End Sub

Private Sub Command4_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Nº Cuenta":       xCampos(0, 1) = "cuenta":           xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":     xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "5000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT * FROM con_planctas ORDER BY cuenta"
    
    xform.Titulo = "Buscando Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtCuenta.Text = NulosC(xRs("cuenta"))
            LblCuenta.Caption = NulosC(xRs("descripcion"))
            LblIdCuenta.Caption = xRs("id")
            TxtNumReg.SetFocus
        End If
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command2_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Documento":     xCampos(0, 1) = "descripcion":           xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":                    xCampos(1, 2) = "800":          xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_documento"
    
    xform.Titulo = "Buscando Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            LblTipDoc.Caption = NulosC(xRs("descripcion"))
            TxtTipDoc.Text = NulosN(xRs("id"))
            TxtIdMoneda.SetFocus
        End If
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Command3_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Documento":     xCampos(0, 1) = "descripcion":           xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":                    xCampos(1, 2) = "800":          xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_moneda"
    
    xform.Titulo = "Buscando Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            LblMoneda.Caption = xRs("descripcion")
            TxtIdMoneda.Text = xRs("id")
            TxtNumDoc.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub Command6_Click()
    If NulosN(TxtIdCliente.Text) = 0 Then
        MsgBox "No ha especificado el cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdCliente.SetFocus
        Exit Sub
    End If
    
    If NulosN(Fg1.TextMatrix(Fg1.Row, 13)) <> 999 Then
        MsgBox "Solo puede modificar registros del apertura de la cuenta corriente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    QueHace = 2
    MuestraDatos
    Fg1.Enabled = False
    TxtIdCliente.Enabled = False
    TxtCliente.Enabled = False
    Frame1.Enabled = False
    
    FrmIngreso.Left = 1380
    FrmIngreso.Top = 1320
    'Blanquea
    FrmIngreso.Visible = True
    TxtTipDoc_Validate True
    TxtIdMoneda_Validate True
    TxtTipDoc.SetFocus

End Sub

Sub MuestraDatos()
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT var_analisisctacte.*, con_planctas.cuenta, con_planctas.descripcion FROM var_analisisctacte LEFT JOIN con_planctas ON " _
        & " var_analisisctacte.idcue = con_planctas.id WHERE (((var_analisisctacte.idlib)=999) AND ((var_analisisctacte.idope)=" & NulosN(Fg1.TextMatrix(Fg1.Row, 12)) & "))", xCon

    If Rst.RecordCount <> 0 Then
        TxtTipDoc.Text = Rst("idtipdoc")
        TxtIdMoneda.Text = Rst("idmon")
        TxtNumDoc.Text = Rst("numdoc")
        TxtFchEmi.Valor = Rst("fchemi")
        TxtNumDocRef.Text = Rst("numdocref")
        TxtNumReg.Text = NulosC(Rst("numreg"))
        LblIdCuenta.Caption = Rst("idcue")
        TxtCuenta.Text = Rst("cuenta")
        LblCuenta.Caption = Rst("descripcion")
        TxtDebeSol.Text = Format(Rst("debesol"), "0.00")
        TxtHaberSol.Text = Format(Rst("habersol"), "0.00")
        TxtDebeDol.Text = Format(Rst("debedol"), "0.00")
        TxtHaberDol.Text = Format(Rst("haberdol"), "0.00")
    End If
    
    Set Rst = Nothing
    
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    If Col = 1 Then
        Dim xCampos2(4, 4) As String
        xCampos2(0, 0) = "Cuenta":         xCampos2(0, 1) = "cuenta":         xCampos2(0, 2) = "1000":    xCampos2(0, 3) = "C"
        xCampos2(1, 0) = "Descripcion":    xCampos2(1, 1) = "descripcion":    xCampos2(1, 2) = "4000":    xCampos2(1, 3) = "C"
        xCampos2(2, 0) = "Tipo":           xCampos2(2, 1) = "tipo":           xCampos2(2, 2) = "1000":    xCampos2(2, 3) = "C"
        xCampos2(3, 0) = "Importe":        xCampos2(3, 1) = "imp":            xCampos2(3, 2) = "1000":    xCampos2(3, 3) = "N"
        
        ' MOSTRAMOS LAS CUENTAS CONTABLES DEL ASIENTO DE APERTURA
        xform.SQLCad = "SELECT con_planctas.id, con_planctas.cuenta, con_planctas.descripcion, con_provicionesdet.[imp], " _
            & " IIf([con_provicionesdet].[tipo]=0,'DEBE','HABER') AS tipo FROM con_provicionesdet LEFT JOIN con_planctas " _
            & " ON con_provicionesdet.idcuen = con_planctas.id"
        
        xform.Titulo = "Buscando Cuenta Contable"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "cuenta"
        xform.CampoBusca = "cuenta"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos2)
        
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 1) = xRs("cuenta")
                Fg1.TextMatrix(Fg1.Row, 2) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 11) = xRs("id")
            End If
        End If
    End If
    
    If Col = 3 Then
        xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "4000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":             xCampos(1, 2) = "1000":    xCampos(1, 3) = "N"
        
        xform.SQLCad = "SELECT * FROM tes_modulos ORDER BY descripcion"
        
        xform.Titulo = "Buscando Modulos"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 3) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 12) = xRs("id")
            End If
        End If
    End If
    
    If Col = 7 Then
        xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "4000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":             xCampos(1, 2) = "1000":    xCampos(1, 3) = "N"
        
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
                Fg1.TextMatrix(Fg1.Row, 7) = NulosC(xRs("descripcion"))
                Fg1.TextMatrix(Fg1.Row, 13) = NulosN(xRs("id"))
            End If
        End If
    End If
    
    If Col = 8 Then
        xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "4000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":             xCampos(1, 2) = "1000":    xCampos(1, 3) = "N"
        
        xform.SQLCad = "SELECT * FROM mae_documento ORDER BY descripcion"
        
        xform.Titulo = "Buscando Documento"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 8) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 14) = xRs("id")
            End If
        End If
    End If
    
    If NulosC(Fg1.TextMatrix(Fg1.Row, 1)) <> "" And NulosC(Fg1.TextMatrix(Fg1.Row, 3)) <> "" And _
         NulosC(Fg1.TextMatrix(Fg1.Row, 7)) <> "" And NulosC(Fg1.TextMatrix(Fg1.Row, 8)) <> "" Then
        Fg1.Rows = Fg1.Rows + 1
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 4 Then
        Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), "0000")
    End If
    If Col = 5 Then
        Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), "0000000000")
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then  ' ELIMINAMOS UN REGISTRO AL PRESIONAR LA TECLA DELETE
        If Fg1.Rows = 1 Then Exit Sub
        Fg1.RemoveItem Fg1.Row
        If Fg1.Rows = 1 Then
        Else
            Fg1.Select 1, 1
        End If
        Fg1.SetFocus
    End If
    
    If KeyCode = 45 Then  ' INSERTAMOS UN REGISTRO AL PRESIONAR LA TECLA INSERT
        If NulosC(Fg1.TextMatrix(Fg1.Row, 1)) <> "" And NulosC(Fg1.TextMatrix(Fg1.Row, 3)) <> "" And _
            NulosC(Fg1.TextMatrix(Fg1.Row, 7)) <> "" And NulosC(Fg1.TextMatrix(Fg1.Row, 8)) <> "" Then
            Fg1.Rows = Fg1.Rows + 1
        End If
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        TxtIdCliente.SetFocus
    End If
End Sub


Private Sub Form_Load()
    SeEjecuto = False
    TxtIdCliente.Text = ""
    TxtCliente.Text = ""
    
    Fg1.SelectionMode = flexSelectionFree
    Fg1.Editable = flexEDKbdMouse
    Fg1.ColWidth(12) = 0
    Fg1.ColWidth(13) = 0
    'Fg1.ColWidth(14) = 0
    Fg1.BackColorSel = &H80&
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
    QueHace = 3
End Sub

Private Sub TxtCuenta_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Command4_Click
    End If
End Sub

Private Sub TxtDebeDol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDebeDol_Validate(Cancel As Boolean)
    'If KeyAscii = 13 Then
        TxtDebeDol.Text = Format(TxtDebeDol.Text, "0.00")
    'End If
End Sub

Private Sub TxtDebeSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDebeSol_Validate(Cancel As Boolean)
    'If KeyAscii = 13 Then
        TxtDebeSol.Text = Format(TxtDebeSol.Text, "0.00")
    'End If
End Sub

Private Sub TxtHaberDol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtHaberDol_Validate(Cancel As Boolean)
    'If KeyAscii = 13 Then
        TxtHaberDol.Text = Format(TxtHaberDol.Text, "0.00")
    'End If
End Sub

Private Sub TxtHaberSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtHaberSol_Validate(Cancel As Boolean)
    'If KeyAscii = 13 Then
        TxtHaberSol.Text = Format(TxtHaberSol.Text, "0.00")
    'End If
End Sub

Private Sub TxtIdCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Sub Blanquea()
    TxtTipDoc.Text = ""
    TxtIdMoneda.Text = ""
    TxtNumDoc.Text = ""
    TxtFchEmi.Valor = ""
    TxtNumDocRef.Text = ""
    TxtDebeDol.Text = ""
    TxtHaberDol.Text = ""
    TxtDebeSol.Text = ""
    TxtHaberSol.Text = ""
    TxtCuenta.Text = ""
    TxtNumReg.Text = ""
    
    LblTipDoc.Caption = ""
    LblMoneda.Caption = ""
    LblCuenta.Caption = ""
    LblIdCuenta.Caption = ""
End Sub
Sub BuscaDatosCliente(xIdCli As Integer)
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq Rst, "SELECT var_analisisctacte.*, mae_documento.abrev, mae_moneda.simbolo AS descmon, mae_documento.abrev AS descdoc, con_planctas.cuenta AS numcue, " _
        & " con_planctas.descripcion AS desccue FROM ((var_analisisctacte LEFT JOIN mae_documento ON var_analisisctacte.idtipdoc = mae_documento.id) " _
        & " LEFT JOIN mae_moneda ON var_analisisctacte.idmon = mae_moneda.id) LEFT JOIN con_planctas ON var_analisisctacte.idcue = con_planctas.id " _
        & " Where (((var_analisisctacte.idcli) = " & xIdCli & ")) ORDER BY var_analisisctacte.numdocref, var_analisisctacte.fchemi", xCon

    If Rst.RecordCount <> 0 Then
        Fg1.Rows = 1
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(Rst("numcue"))
            Fg1.TextMatrix(A, 2) = NulosC(Rst("desccue"))
            Fg1.TextMatrix(A, 3) = NulosC(Rst("descdoc"))
            Fg1.TextMatrix(A, 4) = Rst("numdoc")
            Fg1.TextMatrix(A, 5) = Rst("fchemi")
            Fg1.TextMatrix(A, 6) = Rst("descmon")
            Fg1.TextMatrix(A, 7) = Rst("debedol")
            Fg1.TextMatrix(A, 8) = Rst("haberdol")
            Fg1.TextMatrix(A, 9) = Rst("debesol")
            Fg1.TextMatrix(A, 10) = Rst("habersol")
            Fg1.TextMatrix(A, 11) = NulosC(Rst("numdocref"))
            Fg1.TextMatrix(A, 12) = Rst("idope")
            Fg1.TextMatrix(A, 13) = Rst("idlib")
            
            Rst.MoveNext
            
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

Private Sub TxtIdCliente_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCli_Click
    End If
End Sub

Private Sub TxtIdCliente_Validate(Cancel As Boolean)
    If NulosN(TxtIdCliente.Text) = 0 Then
        TxtIdCliente.Text = ""
        TxtCliente.Text = ""
        Exit Sub
    End If
    
    TxtCliente.Text = Busca_Codigo(TxtIdCliente.Text, "id", "nombre", "mae_cliente", "N", xCon)
    If TxtCliente.Text = "" Then
        TxtIdCliente.Text = ""
    End If
End Sub

Private Sub TxtIdMod_Change()

End Sub

Private Sub TxtIdMod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdMoneda_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Command3_Click
    End If
End Sub

Private Sub TxtIdMoneda_Validate(Cancel As Boolean)
    If NulosN(TxtIdMoneda.Text) <> 0 Then
        LblMoneda.Caption = Busca_Codigo(TxtIdMoneda.Text, "id", "descripcion", "mae_moneda", "N", xCon)
        If NulosC(LblMoneda.Caption) = "" Then
            TxtIdMoneda.Text = ""
        End If
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumReg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Command2_Click
    End If
End Sub

Private Sub TxtTipDoc_Validate(Cancel As Boolean)
    If NulosN(TxtTipDoc.Text) <> 0 Then
        LblTipDoc.Caption = Busca_Codigo(TxtTipDoc.Text, "id", "descripcion", "mae_documento", "N", xCon)
        If NulosC(LblTipDoc.Caption) = "" Then
            TxtTipDoc.Text = ""
        End If
    End If
End Sub
