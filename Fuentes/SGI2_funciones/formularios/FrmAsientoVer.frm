VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmAsientoVer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vista de Asiento"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   8730
      TabIndex        =   4
      Top             =   180
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   10365
      Begin VB.Label LblRegistro 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblRegistro"
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
         Height          =   285
         Left            =   3930
         TabIndex        =   6
         Top             =   180
         Width           =   1575
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
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   570
         TabIndex        =   5
         Top             =   180
         Width           =   2625
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº. Reg"
         Height          =   195
         Left            =   3270
         TabIndex        =   3
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Libro"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   240
         Width           =   345
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   2685
      Left            =   30
      TabIndex        =   0
      Top             =   600
      Width           =   10320
      _cx             =   18203
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmAsientoVer.frx":0000
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
      ShowComboButton =   -1  'True
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
Attribute VB_Name = "FrmAsientoVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim SeEjecuto As Boolean

Dim SGI_JC As New SGI2_funciones.JC_Varios
Dim SGI_JC1 As New SGI2_funciones.JC_VSFlexGrid

Dim RstFrm As New ADODB.Recordset


Public Sub pRecibeLink(xCon As ADODB.Connection, NumRegistro As String)
    '---------------------------
    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
    
    Configurar_Grilla
    LblLibro.Caption = ""
    LblRegistro.Caption = ""
    '--si numero de registro es vacio salir
    If NumRegistro = "" Then Exit Sub
    
    DoEvents
    
    
    
    nSQL = "SELECT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, mae_libros.descripcion AS libro, " _
        + vbCr + " iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tipcam, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
        + vbCr + " iif(con_diario.ajuste=2,0, IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.impdebdol* tipcam ),con_diario.impdebsol) ) AS impdebesol, " _
        + vbCr + " iif(con_diario.ajuste=2,0, IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.imphabdol* tipcam ),con_diario.imphabsol) ) AS imphabersol, " _
        + vbCr + " iif(con_diario.ajuste=1,0, IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/ tipcam ))) ) AS impdebedol, " _
        + vbCr + " iif(con_diario.ajuste=1,0, IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(tipcam=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/ tipcam ))) ) As imphaberdol " _
        + vbCr + " FROM ((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
        + vbCr + " WHERE (((Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi])='" & NumRegistro & "')) " _
        + vbCr + " ORDER BY con_planctas.cuenta; "

    RST_Busq Rst, nSQL, xCon
    
    
    If Rst.RecordCount <> 0 Then
    
        LblLibro.Caption = NulosC(Rst("libro"))
        LblRegistro.Caption = NulosC(Rst("registro"))
    
        Do While Not Rst.EOF
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst("ctanum"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Rst("ctadesc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosN(Rst("tipcam"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosN(Rst("impdebesol"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosN(Rst("imphabersol"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosN(Rst("impdebedol"))
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosN(Rst("imphaberdol"))
                        
            Rst.MoveNext
        Loop
    
        Fg1.Rows = Fg1.Rows + 1
        
            SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 2, &H800000, True, , "TOTAL =>"
            
            SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 4)
            SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 5)
            SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 6)
            SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 7)
    End If
   
   
   
    Me.MousePointer = vbDefault
    
    DoEvents
    
End Sub


Public Sub pRecibeLinkTmp(xCon As ADODB.Connection, xRst As ADODB.Recordset, idlib As Integer, idmov As Double)
    '---------------------------
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    Set RstFrm = xRst
    
    Configurar_Grilla
    LblLibro.Caption = ""
    LblRegistro.Caption = ""
    
    DoEvents
    '--sentencia SQL para determinar el numero de registro de la operacion,
    '--dependera del libro y el movimiento
    nSQL = "SELECT TOP 1 Format(con_diario.idmes,'00') & Format(mae_libros.codsun,'00') & Format(con_diario.numasi,'0000') AS registro " _
        + vbCr + " FROM con_diario LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id " _
        + vbCr + " WHERE (((con_diario.idlib)=" & idlib & ") AND ((con_diario.idmov)=" & idmov & "));"
        
    RST_Busq RstTmp, nSQL, xCon
    
    

    If RstFrm.RecordCount <> 0 Then
        RstFrm.MoveFirst
        LblLibro.Caption = Busca_Codigo(idlib, "id", "descripcion", "mae_libros", "N", xCon)
        
        '--mostrando el numero de registro
        If RstTmp.RecordCount <> 0 Then LblRegistro.Caption = NulosC(RstTmp("registro"))
        '--liberando rst
        Set xRst = Nothing
    
    
        Do While Not RstFrm.EOF
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstFrm("ctanum"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstFrm("ctadesc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosN(RstFrm("tipcam"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosN(RstFrm("impdebmn")), SGI_JC1.FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(RstFrm("imphabmn")), SGI_JC1.FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(RstFrm("impdebme")), SGI_JC1.FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(RstFrm("imphabme")), SGI_JC1.FORMAT_MONTO)
                        
            RstFrm.MoveNext
        Loop
    
        SGI_JC1.GRID_ORDENAR Fg1, 1, 1, , , flexSortGenericAscending
        
        Fg1.Rows = Fg1.Rows + 1
        
        SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 2, &H800000, True, , "TOTAL =>"
        
        SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 4)
        SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 5)
        SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 6)
        SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 7)
        
        '--ajustando el ancho de la columna
        Fg1.AutoSize 4
        Fg1.AutoSize 5
        Fg1.AutoSize 6
        Fg1.AutoSize 7
        
    End If
   

    DoEvents
End Sub

Public Sub fDefinirRst(xRst As ADODB.Recordset)
    Set xRst = Nothing
    Set xRst = New ADODB.Recordset
    
    xRst.Fields.Append "IdCue", adNumeric '--codigo de cuenta
    xRst.Fields.Append "Importe", adDouble '--importe de la cuenta
    xRst.Fields.Append "tipo", adVarChar, 2 '--
    
    xRst.Open

End Sub

Private Sub pCargarDatos()
    '--cargar los datos de la grilla




End Sub


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Fg1_EnterCell()
Fg1.SelectionMode = flexSelectionByRow
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub

    SeEjecuto = True
    
End Sub

Private Sub Form_Deactivate()
    
    On Error Resume Next
    
    Err.Clear
    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Form_Load()
    SeEjecuto = False
    SGI_JC.CentrarFrm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SGI_JC = Nothing
End Sub





Private Sub Configurar_Grilla()

    With Fg1
        '-----
        .Rows = 2
        .FixedRows = 2
        .Cols = 13
        
        .ColWidth(0) = 200
        '--DATOS DE FILA
        
        SGI_JC1.GRID_COMBINAR Fg1, 0, 1, 0, 7, "DATOS DE LA OPERACIÓN", flexAlignCenterCenter, True, , vbBlack, &HC8D0D4, True
        SGI_JC1.GRID_COMBINAR Fg1, 0, 8, 0, 12, "DATOS DE REFERENCIA", flexAlignCenterCenter, True, , vbBlack, &HC8D0D4, True
        
        .FrozenCols = 7
       
        .TextMatrix(1, 1) = "Nª Cuenta":                .ColWidth(1) = 1000:  .ColAlignment(1) = flexAlignLeftCenter:   .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter:       .CellFontBold = True
        .TextMatrix(1, 2) = "Descripción de la Cuenta": .ColWidth(2) = 3500:  .ColAlignment(2) = flexAlignLeftCenter:   .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter:       .CellFontBold = True
        .TextMatrix(1, 3) = "T.C.":                     .ColWidth(3) = 600:   .ColAlignment(3) = flexAlignRightCenter:  .Row = 1: .Col = 3: .CellAlignment = flexAlignRightCenter:      .CellFontBold = True
        .TextMatrix(1, 4) = "Debe MN":                  .ColWidth(4) = 1100: .ColAlignment(4) = flexAlignRightCenter:   .Row = 1: .Col = 4: .CellAlignment = flexAlignRightCenter:      .CellFontBold = True
        .TextMatrix(1, 5) = "Haber MN":                 .ColWidth(5) = 1100: .ColAlignment(5) = flexAlignRightCenter:   .Row = 1: .Col = 5: .CellAlignment = flexAlignRightCenter:      .CellFontBold = True
        .TextMatrix(1, 6) = "Debe ME":                  .ColWidth(6) = 1100: .ColAlignment(6) = flexAlignRightCenter:   .Row = 1: .Col = 6: .CellAlignment = flexAlignRightCenter:      .CellFontBold = True
        .TextMatrix(1, 7) = "Haber ME":                 .ColWidth(7) = 1100: .ColAlignment(7) = flexAlignRightCenter:   .Row = 1: .Col = 7: .CellAlignment = flexAlignRightCenter:      .CellFontBold = True
        
        
        .TextMatrix(1, 8) = "Num.Reg.":                 .ColWidth(8) = 900:   .ColAlignment(8) = flexAlignLeftCenter:   .Row = 1: .Col = 8: .CellAlignment = flexAlignLeftCenter:       .CellFontBold = True
        .TextMatrix(1, 9) = "T.D.":                     .ColWidth(9) = 750:  .ColAlignment(9) = flexAlignLeftCenter:    .Row = 1: .Col = 9: .CellAlignment = flexAlignLeftCenter:       .CellFontBold = True
        .TextMatrix(1, 10) = "Nº.Doc":                   .ColWidth(10) = 1000:  .ColAlignment(10) = flexAlignLeftCenter:   .Row = 1: .Col = 10: .CellAlignment = flexAlignLeftCenter:   .CellFontBold = True
        .TextMatrix(1, 11) = "Fch.Doc":                  .ColWidth(11) = 800:  .ColAlignment(11) = flexAlignLeftCenter:   .Row = 1: .Col = 11: .CellAlignment = flexAlignLeftCenter:    .CellFontBold = True
        .TextMatrix(1, 12) = "Proveedor/Cliente/Otros": .ColWidth(12) = 1800:  .ColAlignment(12) = flexAlignLeftCenter:   .Row = 1: .Col = 12: .CellAlignment = flexAlignLeftCenter:    .CellFontBold = True
        
        Fg1.ColFormat(3) = "0.000"
                
        Fg1.ColFormat(11) = SGI_JC1.FORMAT_DATE
        
        SGI_JC1.OCULTAR_COL Fg1, 8, 12
        
        

    End With
    DoEvents
End Sub

