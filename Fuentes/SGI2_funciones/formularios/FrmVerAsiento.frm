VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmVerAsiento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vista de Asiento"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   10395
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
      Rows            =   20
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmVerAsiento.frx":0000
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
Attribute VB_Name = "FrmVerAsiento"
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
    Dim rst As New ADODB.Recordset
    Dim nSQL As String
    
    Configurar_Grilla
    LblLibro.Caption = ""
    LblRegistro.Caption = ""
    
    DoEvents
    
    
'    nSQL = "SELECT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, mae_libros.descripcion AS libro, con_diario.fchdoc AS fchope, con_diario.rregistro AS registroref, IIf(con_diario.ridtipper=0,tes_documentos.abrev,mae_documento.abrev) AS tdocdesc, con_diario.rfchope AS fchdoc, con_diario.rnumerodoc AS numdoc, IIf([con_diario].[ridtipper]=1,[mae_prov].[nombre],IIf([con_diario].[ridtipper]=2,[mae_cliente].[nombre],IIf([con_diario].[ridtipper]=3,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ', ' & [pla_empleados].[nom],IIf([con_diario].[ridtipper]=5,[mae_bancos].[descripcion],'')))) AS apenom, con_tc.impven AS tc, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
'        + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
'        + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
'        + vbCr + " IIf(con_diario.idmon = 2, con_diario.imphabdol, IIf(con_tc.impven Is Null Or con_diario.imphabsol = 0, 0, (con_diario.imphabsol / con_tc.impven))) As imphaberdol " _
'        + vbCr + " FROM ((pla_empleados RIGHT JOIN (mae_cliente RIGHT JOIN (((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN mae_prov ON con_diario.ridper = mae_prov.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON mae_cliente.id = con_diario.ridper) ON pla_empleados.id = con_diario.ridper) LEFT JOIN tes_documentos ON con_diario.rtipdoc = tes_documentos.id) LEFT JOIN mae_bancos ON con_diario.ridper = mae_bancos.id " _
'        + vbCr + " WHERE (((Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi])='" & NumRegistro & "')) " _
'        + vbCr + " ORDER BY con_planctas.cuenta;"
    
    
    nSQL = "SELECT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, mae_libros.descripcion AS libro, con_tc.impven AS tc, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
        + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
        + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
        + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
        + vbCr + " IIf(con_diario.idmon = 2, con_diario.imphabdol, IIf(con_tc.impven Is Null Or con_diario.imphabsol = 0, 0, (con_diario.imphabsol / con_tc.impven))) As imphaberdol " _
        + vbCr + " FROM ((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
        + vbCr + " WHERE (((Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi])='" & NumRegistro & "')) " _
        + vbCr + " ORDER BY con_planctas.cuenta; "

RST_Busq rst, nSQL, xCon


If rst.RecordCount <> 0 Then

LblLibro.Caption = NulosC(rst("registro"))
LblRegistro.Caption = NulosC(rst("libro"))



Do While Not rst.EOF
    Fg1.Rows = Fg1.Rows + 1
    Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(rst("ctanum"))
    Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(rst("ctadesc"))
    Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosN(rst("tc"))
    Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosN(rst("impdebesol"))
    Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosN(rst("imphabersol"))
    Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosN(rst("impdebedol"))
    Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosN(rst("imphaberdol"))
    
'    Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(rst("registroref"))
'    Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(rst("tdocdesc"))
'    Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosC(rst("numdoc"))
'    Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosC(rst("fchdoc"))
'    Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosC(rst("apenom"))
    
    rst.MoveNext
Loop



    Fg1.Rows = Fg1.Rows + 1

'''    Fg1.TextMatrix(Fg1.Rows - 1, 4) = SGI_JC1.GRID_SUMAR_COL(Fg1, 4)
'''    Fg1.TextMatrix(Fg1.Rows - 1, 5) = SGI_JC1.GRID_SUMAR_COL(Fg1, 5)
'''    Fg1.TextMatrix(Fg1.Rows - 1, 6) = SGI_JC1.GRID_SUMAR_COL(Fg1, 6)
'''    Fg1.TextMatrix(Fg1.Rows - 1, 7) = SGI_JC1.GRID_SUMAR_COL(Fg1, 7)

    SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 2, &H800000, True, , "TOTAL =>"
    SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 4)
    SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 5)
    SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 6)
    SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 7)

End If
   
    
    DoEvents
End Sub

Public Sub pRecibeLinkTmp(xCon As ADODB.Connection, xRst As ADODB.Recordset, idlib As Integer, idmov As Double)
    '---------------------------
    Set RstFrm = xRst
    
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
        
        SGI_JC1.GRID_COMBINAR Fg1, 0, 1, 0, 7, "DATOS DE LA OPERACIÓN", flexAlignCenterCenter, True, , vbBlack, &HC8D0D4
        SGI_JC1.GRID_COMBINAR Fg1, 0, 8, 0, 12, "DATOS DE REFERENCIA", flexAlignCenterCenter, True, , vbBlack, &HC8D0D4
        
        .FrozenCols = 7
       
        .TextMatrix(1, 1) = "Nª Cuenta":                .ColWidth(1) = 1000:  .ColAlignment(1) = flexAlignLeftCenter:   .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 2) = "Descripción de la Cuenta": .ColWidth(2) = 3200:  .ColAlignment(2) = flexAlignLeftCenter:   .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 3) = "T.C.":                     .ColWidth(3) = 600:   .ColAlignment(3) = flexAlignRightCenter:  .Row = 1: .Col = 3: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 4) = "Debe MN":                  .ColWidth(4) = 1100: .ColAlignment(4) = flexAlignRightCenter:   .Row = 1: .Col = 4: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 5) = "Haber MN":                 .ColWidth(5) = 1100: .ColAlignment(5) = flexAlignRightCenter:   .Row = 1: .Col = 5: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 6) = "Debe ME":                  .ColWidth(6) = 1100: .ColAlignment(6) = flexAlignRightCenter:   .Row = 1: .Col = 6: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 7) = "Haber ME":                 .ColWidth(7) = 1100: .ColAlignment(7) = flexAlignRightCenter:   .Row = 1: .Col = 7: .CellAlignment = flexAlignRightCenter
        
        
        .TextMatrix(1, 8) = "Num.Reg.":                 .ColWidth(8) = 900:   .ColAlignment(8) = flexAlignLeftCenter:   .Row = 1: .Col = 8: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 9) = "T.D.":                     .ColWidth(9) = 750:  .ColAlignment(9) = flexAlignLeftCenter:    .Row = 1: .Col = 9: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 10) = "Nº.Doc":                   .ColWidth(10) = 1000:  .ColAlignment(10) = flexAlignLeftCenter:   .Row = 1: .Col = 10: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 11) = "Fch.Doc":                  .ColWidth(11) = 800:  .ColAlignment(11) = flexAlignLeftCenter:   .Row = 1: .Col = 11: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 12) = "Proveedor/Cliente/Otros": .ColWidth(12) = 1800:  .ColAlignment(12) = flexAlignLeftCenter:   .Row = 1: .Col = 12: .CellAlignment = flexAlignLeftCenter
        
        Fg1.ColFormat(3) = "0.000"
        
        Fg1.ColFormat(4) = SGI_JC1.FORMAT_MONTO
        Fg1.ColFormat(5) = SGI_JC1.FORMAT_MONTO
        Fg1.ColFormat(6) = SGI_JC1.FORMAT_MONTO
        Fg1.ColFormat(7) = SGI_JC1.FORMAT_MONTO
                
        Fg1.ColFormat(11) = SGI_JC1.FORMAT_DATE
        
        SGI_JC1.OCULTAR_COL Fg1, 8, 12
    End With
    DoEvents
End Sub

