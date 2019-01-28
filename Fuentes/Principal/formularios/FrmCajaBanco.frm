VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmCajaBancos 
   Caption         =   "Form2"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   LinkTopic       =   "Form2"
   ScaleHeight     =   7575
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   225
      Left            =   1695
      TabIndex        =   26
      Top             =   60
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   7125
      Left            =   15
      TabIndex        =   9
      Top             =   375
      Width           =   10980
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   570
         Left            =   9345
         TabIndex        =   38
         Top             =   6090
         Width           =   1380
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   570
         Left            =   9405
         TabIndex        =   37
         Top             =   3975
         Width           =   1380
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg1 
         Height          =   1170
         Left            =   330
         TabIndex        =   35
         Top             =   3975
         Width           =   8985
         _cx             =   15849
         _cy             =   2064
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
         Rows            =   50
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCajaBanco.frx":0000
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
      Begin VB.CommandButton CmdBusProCli 
         Height          =   240
         Left            =   3255
         Picture         =   "FrmCajaBanco.frx":0103
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3150
         Width           =   240
      End
      Begin VB.TextBox TxtNumRuc 
         Height          =   300
         Left            =   1905
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   8
         Text            =   "TxtNumRuc"
         Top             =   3120
         Width           =   1620
      End
      Begin VB.CommandButton CmdBusCuentaBanco 
         Height          =   240
         Left            =   2550
         Picture         =   "FrmCajaBanco.frx":0235
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1065
         Width           =   240
      End
      Begin VB.CommandButton CmdBusMon 
         Height          =   240
         Left            =   2550
         Picture         =   "FrmCajaBanco.frx":0367
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2055
         Width           =   240
      End
      Begin VB.CommandButton CmdBuscaMovimiento 
         Cancel          =   -1  'True
         Height          =   240
         Index           =   3
         Left            =   2550
         Picture         =   "FrmCajaBanco.frx":0499
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1740
         Width           =   240
      End
      Begin VB.TextBox TxtIdCuenta 
         Height          =   300
         Left            =   1905
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "TxtIdCuenta"
         Top             =   1035
         Width           =   915
      End
      Begin VB.TextBox TxtIdMoneda 
         Height          =   300
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "TxtIdMov"
         Top             =   2025
         Width           =   915
      End
      Begin VB.TextBox TxtIdMov 
         Height          =   300
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "TxtIdMov"
         Top             =   1710
         Width           =   915
      End
      Begin VB.TextBox TxtImporte 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "TxtImporte"
         Top             =   2655
         Width           =   1275
      End
      Begin VB.TextBox TxtTipOpe 
         Height          =   300
         Left            =   1905
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   1
         Text            =   "TxtTipOpe"
         Top             =   720
         Width           =   915
      End
      Begin VB.TextBox TxtTipMov 
         Height          =   300
         Left            =   1905
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   3
         Text            =   "TxtTipMov"
         Top             =   1395
         Width           =   915
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
         Height          =   300
         Left            =   1905
         TabIndex        =   0
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.TextBox TxtGlosa 
         Height          =   300
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "TxtGlosa"
         Top             =   2340
         Width           =   8115
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   1170
         Left            =   330
         TabIndex        =   36
         Top             =   5520
         Width           =   8985
         _cx             =   15849
         _cy             =   2064
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
         Rows            =   50
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCajaBanco.frx":05CB
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Documentos x Pagar"
         Height          =   195
         Index           =   10
         Left            =   330
         TabIndex        =   34
         Top             =   5265
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Documentos Pendientes"
         Height          =   195
         Index           =   9
         Left            =   375
         TabIndex        =   33
         Top             =   3720
         Width           =   1740
      End
      Begin VB.Label LblidCli 
         Caption         =   "LblidCli"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   8430
         TabIndex        =   31
         Top             =   2880
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label LblNomCli 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblNomCli"
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
         Left            =   3600
         TabIndex        =   30
         Top             =   3120
         Width           =   6375
      End
      Begin VB.Label LblDato 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Left            =   360
         TabIndex        =   29
         Top             =   3150
         Width           =   600
      End
      Begin VB.Label LblTipoCambio 
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   6660
         TabIndex        =   28
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio"
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
         Left            =   5130
         TabIndex        =   27
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº Cuenta"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   22
         Top             =   1065
         Width           =   735
      End
      Begin VB.Label LblCuentaBan 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCuentaBan"
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
         Left            =   2865
         TabIndex        =   21
         Top             =   1035
         Width           =   4935
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
         Left            =   2865
         TabIndex        =   20
         Top             =   2025
         Width           =   4935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   19
         Top             =   2040
         Width           =   585
      End
      Begin VB.Label LblMovimiento 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblMovimiento"
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
         Left            =   2865
         TabIndex        =   18
         Top             =   1710
         Width           =   4935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   17
         Top             =   1725
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   16
         Top             =   2655
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Glosa"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   15
         Top             =   2370
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Movimiento"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Top             =   1410
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Operacion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   750
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   390
         Width           =   1305
      End
      Begin VB.Label LblTipoOpera 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblTipoOpera"
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
         Left            =   2865
         TabIndex        =   11
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label LblTipMov 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblTipMov"
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
         Left            =   2865
         TabIndex        =   10
         Top             =   1395
         Width           =   4935
      End
   End
End
Attribute VB_Name = "FrmCajaBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CaracteresNumericos As String
Dim CaracteresNumericos2 As String
Dim QueHace As Integer
Dim xRs1 As New ADODB.Recordset

Private Sub CmdAgregar_Click()
    AgregarParaPago
End Sub

Private Sub CmdBuscaMovimiento_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub

    Dim xForm As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Descripcion":                xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº Cuenta":                  xCampos(1, 1) = "cuenta":        xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Descripcion de la Cuenta":   xCampos(2, 1) = "descue":        xCampos(2, 2) = "3000":         xCampos(2, 3) = "C"
    
    xForm.SQLCad = "SELECT con_cajabanmovi.*, con_planctas.cuenta, con_cajabanmovi.descripcion AS descue, " _
        & " con_cajabanmovi.tipmov, con_cajabanmovi.tipope FROM con_planctas RIGHT JOIN con_cajabanmovi " _
        & " ON con_planctas.id = con_cajabanmovi.idcue WHERE (((con_cajabanmovi.tipmov)=" & Val(TxtTipMov.Text) & ") AND ((con_cajabanmovi.tipope)=1))"

    xForm.Titulo = "Buscando Movimientos"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdMov.Text = xRs("id")
        LblMovimiento.Caption = Trim(xRs("descripcion"))
        TxtIdMoneda.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusCuentaBanco_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Banco":      xCampos(0, 1) = "desban":        xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº Cuenta":  xCampos(1, 1) = "numcue":        xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Moneda":     xCampos(2, 1) = "desmon":        xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    
    xForm.SQLCad = "SELECT mae_bancos.descripcion AS desban, con_bancocuenta.*, mae_moneda.descripcion AS desmon" _
        & " FROM (mae_bancos INNER JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban) " _
        & " LEFT JOIN mae_moneda ON con_bancocuenta.idmon = mae_moneda.id ORDER BY mae_bancos.descripcion"
    
    xForm.Titulo = "Buscando Cuentas de Banco"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "desban"
    xForm.CampoBusca = "desban"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdCuenta.Text = xRs("id")
        LblCuentaBan.Caption = Trim(xRs("desban")) + "   Cuenta Nº " & xRs("numcue")
        TxtTipMov.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusProCli_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Nombre":          xCampos(0, 1) = "nombre":   xCampos(0, 2) = "6000":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":       xCampos(1, 1) = "numruc":   xCampos(1, 2) = "1500":     xCampos(1, 3) = "C"
        

    If TxtTipMov = 1 Then
        xForm.SQLCad = "SELECT mae_cliente.nombre, mae_cliente.numruc, mae_cliente.id " _
            & " From mae_cliente ORDER BY mae_cliente.nombre"
        xForm.Titulo = "Buscando Clientes"
    Else
        xForm.SQLCad = "SELECT mae_prov.id, mae_prov.nombre, mae_prov.numruc " _
            & "  From mae_prov ORDER BY mae_prov.nombre"
        xForm.Titulo = "Buscando Proveedores"
    End If
    
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "nombre"
    xForm.CampoBusca = "nombre"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtNumRuc.Text = xRs("numruc")
        LblNomCli.Caption = xRs("nombre")
        LblidCli.Caption = xRs("id")
        MuestraDocumentos
        Fg1.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Sub MuestraDocumentos()
    Dim rst As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq rst, "SELECT com_compras.idpro, mae_documento.abrev, com_compras.fchdoc, mae_moneda.simbolo, " _
        & " [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.imptot, com_compras.impsal, " _
        & " mae_documento.idcuen, com_compras.id " _
        & " FROM mae_moneda INNER JOIN (mae_documento INNER JOIN (mae_prov INNER JOIN com_compras ON mae_prov.id = com_compras.idpro) " _
        & " ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon " _
        & " WHERE  ((com_compras.idpro=" & Val(LblidCli.Caption) & ") AND (com_compras.impsal<>0))", xCon
    
    If rst.RecordCount <> 0 Then
        Fg1.Rows = 1
        
        rst.MoveFirst
        For A = 1 To rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = rst("abrev")
            Fg1.TextMatrix(A, 2) = rst("fchdoc")
            Fg1.TextMatrix(A, 3) = rst("simbolo")
            Fg1.TextMatrix(A, 4) = rst("numdoc")
            Fg1.TextMatrix(A, 5) = Format(rst("imptot"), "0.00")
            Fg1.TextMatrix(A, 6) = Format(rst("impsal"), "0.00")
            Fg1.TextMatrix(A, 7) = rst("id")
            Fg1.TextMatrix(A, 8) = rst("idcuen")
            rst.MoveNext
            
            If rst.EOF = True Then Exit For
        Next A
    Else
        MsgBox "El proveedor seleccionado no tiene documentos pendientes de pago", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Sub CmdEliminar_Click()
    Fg2.RemoveItem Fg2.Row
End Sub

Private Sub Command1_Click()
    Blanquea
    QueHace = 1
    Bloquea
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Fg2.Col = 7 Then
        Fg2.TextMatrix(Fg2.Row, 7) = Format(Fg2.TextMatrix(Fg2.Row, 7), "0.00")
        Fg2.TextMatrix(Fg2.Row, 8) = Val(Fg2.TextMatrix(Fg2.Row, 6)) - Val(Fg2.TextMatrix(Fg2.Row, 7))
        Fg2.TextMatrix(Fg2.Row, 8) = Format(Fg2.TextMatrix(Fg2.Row, 8), "0.00")
    End If
End Sub

Private Sub Fg2_EnterCell()
    If Fg2.Col = 7 Then
        Fg2.Editable = flexEDKbdMouse
    Else
        Fg2.Editable = flexEDNone
    End If
End Sub

Private Sub Form_Load()
    CaracteresNumericos = "0123456789." & Chr(8)
    CaracteresNumericos2 = "12" & Chr(8)
    QueHace = 3
    Fg1.ColWidth(7) = 0
    Fg1.ColWidth(8) = 0
    
    Fg2.ColWidth(9) = 0
    Fg2.ColWidth(10) = 0
    
    Fg1.Rows = 1
    Fg2.Rows = 1
End Sub

Private Sub TxtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        
        If NulosC(TxtIdCuenta.Text) = "" Then Exit Sub
        
        RST_Busq xRs1, "SELECT mae_bancos.descripcion AS desban, con_bancocuenta.*, mae_moneda.descripcion AS desmon" _
            & " FROM (mae_bancos INNER JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban) " _
            & " LEFT JOIN mae_moneda ON con_bancocuenta.idmon = mae_moneda.id WHERE con_bancocuenta.id = " & Val(TxtIdCuenta.Text) & " " _
            & " ORDER BY mae_bancos.descripcion", xCon
        
        If xRs1.RecordCount = 0 Then
            TxtIdCuenta.Text = ""
            LblCuentaBan.Caption = ""
        Else
            LblCuentaBan.Caption = Trim(xRs1("desban")) + " Nº Cuenta " & Trim(xRs1("numcue"))
        End If
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdMoneda_Change()
    If TxtTipOpe.Text = "1" Then
        LblMoneda.Caption = "Soles"
    End If
    If TxtTipOpe.Text = "2" Then
        LblMoneda.Caption = "Dolares"
    End If
End Sub

Private Sub TxtIdMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos2, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdMov_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        
        If NulosC(TxtIdMov.Text) = "" Then Exit Sub
        
        RST_Busq xRs1, "SELECT con_cajabanmovi.*, con_planctas.cuenta, con_cajabanmovi.descripcion AS descue, " _
        & " con_cajabanmovi.tipmov, con_cajabanmovi.tipope FROM con_planctas RIGHT JOIN con_cajabanmovi " _
        & " ON con_planctas.id = con_cajabanmovi.idcue WHERE (((con_cajabanmovi.tipmov)=1) " _
        & " AND ((con_cajabanmovi.tipope)=1))", xCon

        LblMovimiento.Caption = Trim(xRs1("descripcion"))
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipMov_Change()
    If TxtTipMov.Text = "1" Then
        LblTipMov.Caption = "Ingreso"
        LblDato.Caption = "Clientes"
    End If
    
    If TxtTipMov.Text = "2" Then
        LblTipMov.Caption = "Egreso"
        LblDato.Caption = "Proveedores"
    End If
End Sub

Private Sub TxtTipMov_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos2, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipOpe_Change()
    If TxtTipOpe.Text = "1" Then
        LblTipoOpera.Caption = "Caja"
        
        TxtIdCuenta.Text = ""
        LblCuentaBan.Caption = ""
        TxtIdCuenta.Locked = True
        CmdBusCuentaBanco.Enabled = False
    End If
    If TxtTipOpe.Text = "2" Then
        LblTipoOpera.Caption = "Bancos"
        TxtIdCuenta.Locked = False
        CmdBusCuentaBanco.Enabled = True
    End If
End Sub

Private Sub TxtTipOpe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos2, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Sub Blanquea()
    TxtFecha.Valor = ""
    TxtTipOpe.Text = ""
    TxtIdCuenta.Text = ""
    TxtTipMov.Text = ""
    TxtIdMov.Text = ""
    TxtIdMoneda = ""
    TxtGlosa.Text = ""
    TxtImporte.Text = ""
    TxtNumRuc.Text = ""
    
    LblTipoOpera.Caption = ""
    LblCuentaBan.Caption = ""
    LblTipMov.Caption = ""
    LblMovimiento.Caption = ""
    LblMoneda.Caption = ""
End Sub

Sub Bloquea()
    TxtFecha.Locked = Not TxtFecha.Locked
    TxtTipOpe.Locked = Not TxtTipOpe.Locked
    TxtIdCuenta.Locked = Not TxtIdCuenta.Locked
    TxtIdMov.Locked = Not TxtIdMov.Locked
    TxtIdMoneda.Locked = Not TxtIdMoneda.Locked
    TxtGlosa.Locked = Not TxtGlosa.Locked
    TxtTipMov.Locked = Not TxtTipMov.Locked
    
End Sub

Sub AgregarParaPago()
    Dim A As Integer
    
    If Fg1.Rows = 1 Then
        MsgBox "No hay documentos para agregar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    For A = 1 To Fg2.Rows - 1
        If Val(Fg2.TextMatrix(A, 7)) = Val(Fg1.TextMatrix(Fg1.Row, 7)) Then
            MsgBox "El documento seleccionado ya esta agregado para cancelacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
    Next A
    
    Fg2.Rows = Fg2.Rows + 1
    Fg2.TextMatrix(Fg2.Rows - 1, 1) = Fg1.TextMatrix(Fg1.Row, 1)
    Fg2.TextMatrix(Fg2.Rows - 1, 2) = Fg1.TextMatrix(Fg1.Row, 2)
    Fg2.TextMatrix(Fg2.Rows - 1, 3) = Fg1.TextMatrix(Fg1.Row, 3)
    Fg2.TextMatrix(Fg2.Rows - 1, 4) = Fg1.TextMatrix(Fg1.Row, 4)
    
    Fg2.TextMatrix(Fg2.Rows - 1, 5) = Fg1.TextMatrix(Fg1.Row, 5)
    Fg2.TextMatrix(Fg2.Rows - 1, 6) = Fg1.TextMatrix(Fg1.Row, 6)
    Fg2.TextMatrix(Fg2.Rows - 1, 9) = Fg1.TextMatrix(Fg1.Row, 7)
    Fg2.TextMatrix(Fg2.Rows - 1, 10) = Fg1.TextMatrix(Fg1.Row, 8)
End Sub
