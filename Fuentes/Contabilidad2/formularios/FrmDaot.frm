VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmDaot 
   Caption         =   "Contabilidad - DAOT"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   11625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   900
      Left            =   0
      TabIndex        =   1
      Top             =   -60
      Width           =   11625
      Begin VB.CommandButton Command1 
         Height          =   570
         Left            =   8175
         Picture         =   "FrmDaot.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Mostrar"
         Top             =   210
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clientes"
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
         Height          =   225
         Left            =   285
         TabIndex        =   7
         Top             =   240
         Width           =   2250
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Proveedores"
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
         Height          =   225
         Left            =   285
         TabIndex        =   6
         Top             =   525
         Width           =   2250
      End
      Begin VB.CommandButton Command3 
         Height          =   570
         Left            =   8820
         Picture         =   "FrmDaot.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Height          =   570
         Left            =   10755
         Picture         =   "FrmDaot.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton CmdExp 
         Height          =   570
         Left            =   9465
         Picture         =   "FrmDaot.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exportar a Excel"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton CmdExpPDT 
         Height          =   570
         Left            =   10110
         Picture         =   "FrmDaot.frx":1560
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exportar a PDT"
         Top             =   210
         Width           =   615
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6510
      Left            =   0
      TabIndex        =   0
      Top             =   855
      Width           =   11610
      _cx             =   20479
      _cy             =   11483
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
      Rows            =   50
      Cols            =   13
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmDaot.frx":2022
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
Attribute VB_Name = "FrmDaot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim SeEjecuto As Boolean

Private Sub CmdExp_Click()
    On Error GoTo ERROR
    Dim X_PRINT As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    'X_PRINT.Imprimir_x_VSFlexGrid Fg1, "Declaracion Anual Operaciones con Terceros DAOT", "Declaracion Anual", "", False, True
    X_PRINT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Declaracion Anual Operaciones con Terceros DAOT", "Todo el Periodo", "Declaracion Anual", "daot0001.xls"
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
ERROR:
    Me.MousePointer = vbDefault
    SHOW_ERROR
End Sub

Private Sub CmdExpPDT_Click()
    Exportar
End Sub

Sub Exportar()
    Dim NomArch, xCad As String
    Dim A As Integer
    
    If Rst.RecordCount <> 0 Then
        NomArch = "costos.txt"
        Open Trim(App.Path) + "\" + NomArch For Output As #1
    
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            xCad = ""
            xCad = xCad + Trim(Str(A)) + "|"
            xCad = xCad + Rst("tipdocdecl") + "|"
            xCad = xCad + Rst("numdocdecla") + "|"
            xCad = xCad + Rst("anodecla") + "|"
            xCad = xCad + Rst("tipperpro") + "|"
            
            xCad = xCad + Rst("tipdocpro") + "|"
            xCad = xCad + Rst("numruc") + "|"
            xCad = xCad + Format(Rst("Total"), "0") + "|"
            xCad = xCad + Mid(NulosC(Rst("apepro1")), 1, 20) + "|"
            xCad = xCad + Mid(NulosC(Rst("apepro2")), 1, 20) + "|"
            xCad = xCad + Mid(NulosC(Rst("nompro1")), 1, 20) + "|"
            xCad = xCad + Mid(NulosC(Rst("nompro2")), 1, 20) + "|"
            xCad = xCad + Mid(Rst("nomprovedor"), 1, 40) + "|"
            
            Print #1, Trim(xCad)
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
    
    Close #1
    MsgBox "El archivo para el DAOT se genero con exito", vbInformation + vbOKCancel + vbDefaultButton1, xTitulo
    
End Sub

Private Sub Command1_Click()
    If Option1.Value = True Then CargarDAOTCliente
    If Option2.Value = True Then CargarDAOTProveedor
End Sub

Private Sub Command2_Click()
    Set Rst = Nothing
    Unload Me
End Sub

Sub CargarDAOTCliente()
    Dim A, xFila As Integer
    Dim xTotal As Double
'
'    RST_Busq Rst, "SELECT DISTINCT '6' AS tipdocdecl, '99999999999' AS numdocdecla, '2005' AS anodecla, mae_tipoempresa.codsun AS tipperpro, mae_dociden.codsun AS tipdocpro, mae_cliente.numruc, mae_cliente.apecli1, mae_cliente.apecli2, mae_cliente.nomcli1, mae_cliente.nomcli2, IIf([mae_cliente]![tipper]=1,'',[mae_cliente].[nombre]) AS nomprovedor, Sum(IIf([vta_ventas].[idmon]=1,[vta_ventas]![impbru]+[vta_ventas]![impinaf],([vta_ventas]![impbru]*[con_tc].[impven])+([vta_ventas]![impinaf]*[con_tc].[impven]))) AS total"
'FROM (mae_dociden RIGHT JOIN (mae_tipoempresa RIGHT JOIN (mae_cliente LEFT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON mae_tipoempresa.id = mae_cliente.tipper) ON mae_dociden.id = mae_cliente.tipdoc) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha
'WHERE (((vta_ventas.numreg)<>'000001'))
'Group By '6', mae_tipoempresa.codsun, mae_dociden.codsun, mae_cliente.numruc, mae_cliente.apecli1, mae_cliente.apecli2, mae_cliente.nomcli1, mae_cliente.nomcli2, IIf([mae_cliente]![tipper]=1,'',[mae_cliente].[nombre]), vta_ventas.tipdoc
'HAVING (((Sum(IIf([vta_ventas].[idmon]=1,[vta_ventas]![impbru]+[vta_ventas]![impinaf],([vta_ventas]![impbru]*[con_tc].[impven])+([vta_ventas]![impinaf]*[con_tc].[impven]))))>=6900));
'
    
    Fg1.Rows = 2
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        xFila = 2
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(xFila, 1) = Rst("tipdocdecl")
            Fg1.TextMatrix(xFila, 2) = Rst("numdocdecla")
            Fg1.TextMatrix(xFila, 3) = Rst("anodecla")
            
            Fg1.TextMatrix(xFila, 4) = Rst("tipperpro")
            Fg1.TextMatrix(xFila, 5) = Rst("tipdocpro")
            Fg1.TextMatrix(xFila, 6) = Rst("numruc")
            Fg1.TextMatrix(xFila, 7) = NulosC(Rst("apepro1"))
            Fg1.TextMatrix(xFila, 8) = NulosC(Rst("apepro2"))
            Fg1.TextMatrix(xFila, 9) = NulosC(Rst("nompro1"))
            Fg1.TextMatrix(xFila, 10) = NulosC(Rst("nompro2"))
            Fg1.TextMatrix(xFila, 11) = NulosC(Rst("nomprovedor"))
            Fg1.TextMatrix(xFila, 12) = Format(Rst("Total"), "0")
            xTotal = xTotal + NulosN(Format(Rst("Total"), "0"))
            Rst.MoveNext
            
            If Rst.EOF = True Then Exit For
            xFila = xFila + 1
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 11) = "TOTAL ==>"
        Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(xTotal, "0")
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H800000, True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H800000, True
    End If
End Sub

Sub CargarDAOTProveedor()
    Dim A, xFila As Integer
    Dim xTotal As Double
    
    Fg1.Rows = 2
    DoEvents
    
    RST_Busq Rst, "SELECT '" & TIPODOCUMENTOIDEN & "' AS tipdocdecl, '" & NumRUC & "' AS numdocdecla, '" & AnoTra & "' AS anodecla, mae_tipoempresa.codsun AS tipperpro, mae_dociden.codsun AS tipdocpro, " _
        & " mae_prov.numruc, IIf([mae_prov]![tipper]=2,'',IIf([mae_prov]![apepro1] Is Null Or [mae_prov]![apepro1]='',[mae_prov]![nombre],[mae_prov]![apepro1])) AS apepro1, mae_prov.apepro2, mae_prov.nompro1, mae_prov.nompro2, IIf(mae_prov!tipper=1,'',mae_prov!nombre) AS nomprovedor, " _
        & " Sum(IIf(com_compras!idmon=1,com_compras!impina+com_compras!impbru,(com_compras!impina*con_tc!impven)+(com_compras!impbru*con_tc!impven))) AS Total " _
        & " FROM mae_dociden INNER JOIN ((mae_tipoempresa RIGHT JOIN mae_prov ON mae_tipoempresa.id = mae_prov.tipper) LEFT JOIN (com_compras LEFT JOIN con_tc " _
        & " ON com_compras.fchdoc = con_tc.fecha) ON mae_prov.id = com_compras.idpro) ON mae_dociden.id = mae_prov.idtipdoc WHERE (((com_compras.tipdoc)<>2) " _
        & " AND ((com_compras.numreg) not in ('000001','0001'))) " _
        & " GROUP BY  mae_tipoempresa.codsun, mae_dociden.codsun, mae_prov.numruc, IIf([mae_prov]![tipper]=2,'',IIf([mae_prov]![apepro1] Is Null Or [mae_prov]![apepro1]='',[mae_prov]![nombre],[mae_prov]![apepro1])) , " _
        & " mae_prov.apepro2, mae_prov.nompro1, mae_prov.nompro2, IIf(mae_prov!tipper=1,'',mae_prov!nombre) Having (((Sum(IIf([com_compras]![idmon] = 1, " _
        & " [com_compras]![impina] + [com_compras]![impbru], ([com_compras]![impina] * [con_tc]![impven]) + ([com_compras]![impbru] * [con_tc]![impven])))) >= 7100)) " _
        & " ORDER BY Sum(IIf(com_compras!idmon=1,com_compras!impina+com_compras!impbru,(com_compras!impina*con_tc!impven)+(com_compras!impbru*con_tc!impven)))", xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        xFila = 2
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(xFila, 1) = NulosC(Rst("tipdocdecl"))
            Fg1.TextMatrix(xFila, 2) = NulosC(Rst("numdocdecla"))
            Fg1.TextMatrix(xFila, 3) = NulosC(Rst("anodecla"))
            
            Fg1.TextMatrix(xFila, 4) = NulosC(Rst("tipperpro"))
            Fg1.TextMatrix(xFila, 5) = NulosC(Rst("tipdocpro"))
            Fg1.TextMatrix(xFila, 6) = NulosC(Rst("numruc"))
            Fg1.TextMatrix(xFila, 7) = NulosC(Rst("apepro1"))
            Fg1.TextMatrix(xFila, 8) = NulosC(Rst("apepro2"))
            Fg1.TextMatrix(xFila, 9) = NulosC(Rst("nompro1"))
            Fg1.TextMatrix(xFila, 10) = NulosC(Rst("nompro2"))
            Fg1.TextMatrix(xFila, 11) = NulosC(Rst("nomprovedor"))
            Fg1.TextMatrix(xFila, 12) = Format(NulosN(Rst("Total")), FORMAT_MONTO)
            xTotal = xTotal + NulosN(Rst("Total"))
            
            Rst.MoveNext
            
            If Rst.EOF = True Then Exit For
            xFila = xFila + 1
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 11) = "TOTAL ==>"
        Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(xTotal, FORMAT_MONTO)
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H800000, True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H800000, True
    End If
End Sub

Private Sub Command3_Click()
    On Error GoTo ERROR
    Dim X_PRINT As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, "Declaracion Anual Operaciones con Terceros DAOT", "Declaracion Anual", "", False, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
ERROR:
    Me.MousePointer = vbDefault
    SHOW_ERROR
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Option2.Value = True
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    UNIR_CELDAS Fg1, 0, 1, 0, 3, "Datos del Declarante", flexAlignCenterCenter, True
    UNIR_CELDAS Fg1, 0, 4, 0, 12, "Datos del Declarado", flexAlignCenterCenter, True
    Fg1.Rows = 2
    Fg1.SelectionMode = flexSelectionByRow
End Sub

