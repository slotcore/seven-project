VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmTransOpe 
   Caption         =   "Operaciones - Mover Operaciones a Otro Periodo"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBusMesFin 
      Height          =   240
      Left            =   3060
      Picture         =   "FrmTransOpe.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   795
      Width           =   240
   End
   Begin VB.CommandButton CmdBusMesIni 
      Height          =   240
      Left            =   3060
      Picture         =   "FrmTransOpe.frx":0132
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   480
      Width           =   240
   End
   Begin VB.CommandButton CmdBusModulo 
      Height          =   240
      Left            =   2205
      Picture         =   "FrmTransOpe.frx":0264
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   150
      Width           =   240
   End
   Begin VB.Frame Frame1 
      Height          =   1080
      Left            =   5730
      TabIndex        =   8
      Top             =   -30
      Width           =   4425
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Command1"
         Height          =   720
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Eliminar Documento"
         Top             =   240
         Width           =   810
      End
      Begin VB.CommandButton CmdAñadir 
         Caption         =   "Command1"
         Height          =   720
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Agregar Documento"
         Top             =   240
         Width           =   810
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Command1"
         Height          =   720
         Left            =   3285
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   810
      End
      Begin VB.CommandButton CmdTransferir 
         Caption         =   "Command1"
         Height          =   720
         Left            =   2475
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Transferir a Otro Mes"
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.TextBox TxtMesFin 
      Height          =   300
      Left            =   1725
      TabIndex        =   2
      Text            =   "TxtMesFin"
      Top             =   765
      Width           =   1605
   End
   Begin VB.TextBox TxtMesIni 
      Height          =   300
      Left            =   1725
      TabIndex        =   1
      Text            =   "TxtMesIni"
      Top             =   450
      Width           =   1605
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   4845
      Left            =   45
      TabIndex        =   5
      Top             =   1125
      Width           =   10125
      _cx             =   17859
      _cy             =   8546
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
      BackColor       =   14876414
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   128
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14876414
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmTransOpe.frx":0396
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
   Begin VB.TextBox TxtidModulo 
      Height          =   300
      Left            =   1725
      TabIndex        =   0
      Text            =   "TxtIdModulo"
      Top             =   120
      Width           =   750
   End
   Begin VB.Label LblIdMesDes 
      AutoSize        =   -1  'True
      Caption         =   "LblIdMesDes"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3465
      TabIndex        =   14
      Top             =   810
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label LblIdMesOri 
      AutoSize        =   -1  'True
      Caption         =   "LblIdMesOri"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3525
      TabIndex        =   13
      Top             =   525
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mes de Origen"
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   7
      Top             =   495
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mes de Transferencia"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   6
      Top             =   825
      Width           =   1545
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Modulo"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   150
      Width           =   525
   End
   Begin VB.Label LblDescModulo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblDescModulo"
      Height          =   300
      Left            =   2535
      TabIndex        =   3
      Top             =   120
      Width           =   3120
   End
End
Attribute VB_Name = "FrmTransOpe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CaracteresNumericos As String
Dim SeEjecuto As Boolean

Private Sub CmdAñadir_Click()
    If TxtidModulo.Text = "" Then
        MsgBox "No ha especificado de que modulo se hara la transferencia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtidModulo.SetFocus
        Exit Sub
    End If
    If NulosN(LblIdMesOri.Caption) = 0 Then
        MsgBox "No ha especificado el mes de origen àra realizar esta operacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtMesIni.SetFocus
        Exit Sub
    End If
    
    Dim xCampos(5, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLId As String
    Dim A&
    
    xCampos(0, 0) = "NºRegistro":      xCampos(0, 1) = "numreg":        xCampos(0, 2) = "1000":         xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
    xCampos(1, 0) = "T.D":             xCampos(1, 1) = "abrev":         xCampos(1, 2) = "800":          xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "Fch. Emi.":       xCampos(2, 1) = "fchdoc":        xCampos(2, 2) = "900":          xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
    xCampos(3, 0) = "Nº Documento":    xCampos(3, 1) = "numdoc":        xCampos(3, 2) = "1200":         xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
    xCampos(4, 0) = "Proveedor":       xCampos(4, 1) = "nombre":        xCampos(4, 2) = "3000":         xCampos(4, 3) = "C":    xCampos(4, 4) = "N"

    '*******************************************************************************************
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 9, "alm_inventario.id", " NOT IN ", True)
    '*******************************************************************************************

    If NulosN(TxtidModulo.Text) = 1 Then
        nSQL = "SELECT Mid([com_compras]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_compras]![numreg],3,4) AS numreg, mae_documento.abrev, " _
            & " com_compras.fchven, com_compras.fchdoc, mae_prov.nombre, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.id, " _
            & " Format([fchreg],'mm') AS idmes FROM mae_documento RIGHT JOIN (mae_prov RIGHT JOIN (com_compras LEFT JOIN mae_libros " _
            & " ON com_compras.idlib = mae_libros.id) ON mae_prov.id = com_compras.idpro) ON mae_documento.id = com_compras.tipdoc " _
            & " WHERE (((Format([fchreg],'mm'))='" & Format(NulosN(LblIdMesOri.Caption), "00") & "')) ORDER BY Mid([com_compras]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_compras]![numreg],3,4)"
    End If

   '*******************************************************************************************
    CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Documentos"
    '*******************************************************************************************
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then xRs.MoveFirst
        Do While Not xRs.EOF
            Fg1.Rows = Fg1.Rows + 1
            
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("numreg"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = xRs("abrev")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(xRs("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(xRs("nombre"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(xRs("numdoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(xRs("fchven"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = xRs("id")
            xRs.MoveNext
        Loop
    End If
    Set xRs = Nothing
    
End Sub

Private Sub CmdBusMesFin_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM con_meses ORDER BY descripcion"
    
    xform.Titulo = "Buscando Modulos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs("id") = NulosN(LblIdMesOri.Caption) Then
            MsgBox "No puede especificar el mismo mes para el mes de transferencia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        If xRs.RecordCount <> 0 Then
            TxtMesFin.Text = xRs("descripcion")
            LblIdMesDes.Caption = xRs("id")
            CmdAñadir.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMesIni_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM con_meses ORDER BY descripcion"
    
    xform.Titulo = "Buscando Modulos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtMesIni.Text = xRs("descripcion")
            LblIdMesOri.Caption = xRs("id")
            TxtMesFin.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusModulo_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
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
            TxtidModulo.Text = xRs("id")
            LblDescModulo.Caption = xRs("descripcion")
            TxtMesIni.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdTransferir_Click()
    If NulosC(TxtMesFin.Text) = "" Then
        MsgBox "No ha especificado el mes de transferencia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtMesFin.SetFocus
        Exit Sub
    End If
    
    If NulosN(TxtidModulo.Text) = 1 Then TransferirCompras
    
    MsgBox "Los documentos se tranfirieron con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Fg1.Rows = 1
    Blanquea
    TxtidModulo.SetFocus
End Sub

Sub TransferirCompras()
    Dim A As Integer
    Dim xNumReg, xNumReg2, xCodSunLib As String
    Dim xMesReg As String
    'AnoTra = AnoTra
    xMesReg = "01/" & Format(LblIdMesDes.Caption, "00") & "/" & Format(AnoTra, "0000")
    xCodSunLib = Busca_Codigo(1, "id", "codsun", "mae_libros", "N", xCon)
    For A = 1 To Fg1.Rows - 1
        xNumReg = NuevoNumAsiento(1, NulosN(LblIdMesDes.Caption), xCon)
        
        xCon.Execute "UPDATE com_compras SET com_compras.fchreg = CDate('" & xMesReg & "'), com_compras.numreg = '" & Format(LblIdMesDes.Caption, "00") & xNumReg & "'" _
            & " WHERE (((com_compras.id)=" & NulosN(Fg1.TextMatrix(A, 7)) & "))"

        xNumReg2 = Format(NulosN(LblIdMesDes.Caption), "00") & xCodSunLib & xNumReg
        xCon.Execute "UPDATE con_diario SET con_diario.numasi = '" & xNumReg & "', con_diario.idmes = " & NulosN(LblIdMesDes.Caption) & ", " _
            & " con_diario.fchasi = CDate('" & xMesReg & "'), con_diario.rregistro = '" & xNumReg2 & "'" _
            & " WHERE (((con_diario.idlib)=1) AND ((con_diario.idmov)=" & NulosN(Fg1.TextMatrix(A, 7)) & "))"
    Next A
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        TxtidModulo.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim Ruta As String
    
    Ruta = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
    Me.ScaleMode = 3
    
    CaracteresNumericos = "0123456789." & Chr(8)

    Blanquea
    CmdAñadir.Caption = "Añadir Documento"
    CmdEliminar.Caption = "Eliminar Documento"
    CmdTransferir.Caption = ""
    CmdSalir.Caption = ""
    CmdTransferir.Picture = LeerIcono(Ruta + "toolbar\18.ico", T32x32, Me, Me.BackColor)
    CmdSalir.Picture = LeerIcono(Ruta + "toolbar\16.ico", T32x32, Me, Me.BackColor)
    
    Fg1.ColWidth(7) = 0
    Fg1.Editable = flexEDNone
    Fg1.Rows = 1
    Fg1.SelectionMode = flexSelectionByRow
End Sub

Sub Blanquea()
    TxtidModulo.Text = ""
    TxtMesIni.Text = ""
    TxtMesFin.Text = ""
    LblIdMesOri.Caption = ""
    LblIdMesDes.Caption = ""
    LblDescModulo.Caption = ""
End Sub

Private Sub TxtidModulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtidModulo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusModulo_Click
    End If
End Sub

Private Sub TxtidModulo_Validate(Cancel As Boolean)
    If NulosN(TxtidModulo.Text) <> 0 Then
        LblDescModulo.Caption = Busca_Codigo(TxtidModulo.Text, "id", "descripcion", "tes_modulos", "N", xCon)
        If NulosC(LblDescModulo.Caption) = "" Then
            TxtidModulo.Text = ""
        End If
    End If
End Sub

Private Sub TxtMesFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtMesFin_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMesFin_Click
    End If
End Sub

Private Sub TxtMesIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtMesIni_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMesIni_Click
    End If
End Sub
