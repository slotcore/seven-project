VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmDaot1 
   Caption         =   "Contabilidad - DAOT"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   12630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   900
      Left            =   0
      TabIndex        =   8
      Top             =   -60
      Width           =   11625
      Begin VB.ComboBox CbSimbolo 
         Height          =   315
         ItemData        =   "FrmDaot1.frx":0000
         Left            =   4140
         List            =   "FrmDaot1.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   480
         Width           =   645
      End
      Begin VB.Frame Frame3 
         Caption         =   "[  Seleccionar  ]"
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
         Height          =   735
         Left            =   1530
         TabIndex        =   15
         Top             =   120
         Width           =   1605
         Begin VB.OptionButton OptSel2 
            Caption         =   "Seleccionar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   17
            Top             =   465
            Width           =   1140
         End
         Begin VB.OptionButton OptSel1 
            Caption         =   "Todos"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   16
            Top             =   270
            Value           =   -1  'True
            Width           =   840
         End
      End
      Begin VB.CommandButton CmdBusCliPro 
         Enabled         =   0   'False
         Height          =   240
         Left            =   7920
         Picture         =   "FrmDaot1.frx":0028
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   210
         Width           =   210
      End
      Begin VB.TextBox txtBase 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         TabIndex        =   1
         Text            =   "txtBase"
         Top             =   510
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   570
         Left            =   8355
         Picture         =   "FrmDaot1.frx":015A
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Left            =   105
         TabIndex        =   0
         Top             =   240
         Width           =   1200
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
         Left            =   105
         TabIndex        =   9
         Top             =   525
         Width           =   1440
      End
      Begin VB.CommandButton Command3 
         Height          =   570
         Left            =   9000
         Picture         =   "FrmDaot1.frx":059C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Imprimir"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Height          =   570
         Left            =   10935
         Picture         =   "FrmDaot1.frx":08A6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton CmdExp 
         Height          =   570
         Left            =   9645
         Picture         =   "FrmDaot1.frx":0BB0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exportar a Excel"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton CmdExpPDT 
         Height          =   570
         Left            =   10290
         Picture         =   "FrmDaot1.frx":16BA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exportar a PDT"
         Top             =   210
         Width           =   615
      End
      Begin VB.TextBox TxtCliPro 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   4155
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "TxtCliPro"
         Top             =   165
         Width           =   4005
      End
      Begin VB.Line Line2 
         X1              =   8250
         X2              =   8250
         Y1              =   840
         Y2              =   150
      End
      Begin VB.Line Line1 
         X1              =   3090
         X2              =   3090
         Y1              =   720
         Y2              =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
         Height          =   195
         Index           =   0
         Left            =   3210
         TabIndex        =   14
         Top             =   270
         Width           =   630
      End
      Begin VB.Label LblIdCliPro 
         Caption         =   "LblIdCliPro"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7815
         TabIndex        =   13
         Top             =   60
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Base en MN"
         Height          =   195
         Index           =   1
         Left            =   3210
         TabIndex        =   10
         Top             =   555
         Width           =   885
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6510
      Left            =   0
      TabIndex        =   7
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
      Cols            =   17
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmDaot1.frx":217C
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
Attribute VB_Name = "FrmDaot1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--modificado 08/02/10 por Johan Castro
'                      modificar la consulta para compras
'                      agregar consulta para ventas
'                      se agrega columna de correlativo, importes en MN y ME, total expresado a MN y ME
'                      agregar filtro por proveedor o cliente, por importe base
                       'Nota:Para obtener el importe base se suma [impbru; impbru2; impbru3; impinaf; impisc]
                       'Las personas no domicialiados no se consideran en este reporte

Option Explicit

Dim Rst As New ADODB.Recordset
Dim SeEjecuto As Boolean

Private Sub CmdExp_Click()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    'X_PRINT.Imprimir_x_VSFlexGrid Fg1, "Declaracion Anual Operaciones con Terceros DAOT", "Declaracion Anual", "", False, True
    X_PRINT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Declaracion Anual Operaciones con Terceros DAOT - " & IIf(Option1.Value = True, "Clientes", "Proveedores"), "Ejercicio " & AnoTra, "", "daot0001.xls"
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR
End Sub

Private Sub CmdExpPDT_Click()
    Exportar
End Sub

Sub Exportar()
    Dim NomArch, xCad As String
    Dim A As Integer
    
    If Rst.State = 0 Then Exit Sub
    
    If Rst.RecordCount = 0 Then
        MsgBox "No hay registros para exportar", vbInformation, xTitulo
        Exit Sub
    End If
    
    If Rst.RecordCount <> 0 Then
        If Option1.Value = True Then
            NomArch = "ingresos.txt"
        Else
            NomArch = "costos.txt"
        End If
        
        Open Trim(App.Path) + "\" + NomArch For Output As #1
    
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            xCad = ""
            xCad = xCad & Trim(Str(A)) + "|"
            xCad = xCad & NulosN(Rst("tipdocdecl")) & "|"
            xCad = xCad & NulosC(Rst("numdocdecla")) & "|"
            xCad = xCad & NulosC(Rst("anodecla")) & "|"
            xCad = xCad & NulosC(Rst("tipperpro")) & "|"
            
            xCad = xCad & NulosN(Rst("tipdocpro")) & "|"
            xCad = xCad & NulosC(Rst("numruc")) & "|"
            xCad = xCad & Format(NulosN(Rst("imptotexpmn")), "0") & "|"
            xCad = xCad & Mid(NulosC(Rst("apepro1")), 1, 20) & "|"
            xCad = xCad & Mid(NulosC(Rst("apepro2")), 1, 20) & "|"
            xCad = xCad & Mid(NulosC(Rst("nompro1")), 1, 20) & "|"
            xCad = xCad & Mid(NulosC(Rst("nompro2")), 1, 20) & "|"
            xCad = xCad & Mid(Rst("nomprovedor"), 1, 40) & "|"
            
            Print #1, Trim(xCad)
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
    
    Close #1
    
    MsgBox "El archivo para el DAOT se generó con éxito" & vbCr & "El archivo se grabó en el sgte. directorio: " & vbCr & Trim(App.Path) + "\" + NomArch, vbInformation, xTitulo
    
End Sub

Private Sub Command1_Click()
    
    '--verificar si han ingresado el importe base
    If NulosN(txtBase.Text) = 0 Then
        MsgBox "Falta especificar la Base", vbInformation, xTitulo
        txtBase.SetFocus
        Exit Sub
    End If
    
    If CbSimbolo.ListIndex = -1 Then
        MsgBox "Falta seleccionar el símbolo", vbInformation, xTitulo
        CbSimbolo.SetFocus
        Exit Sub
    End If
    
    Set Rst = Nothing
    Fg1.Rows = Fg1.FixedRows
    
    DoEvents

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
    Dim nSQL As String
    Dim nSQLIdCliente As String '--sentencia sql para filtrar por Cliente

    
'
'    RST_Busq Rst, "SELECT DISTINCT '6' AS tipdocdecl, '99999999999' AS numdocdecla, '2005' AS anodecla, mae_tipoempresa.codsun AS tipperpro, mae_dociden.codsun AS tipdocpro, mae_cliente.numruc, mae_cliente.apecli1, mae_cliente.apecli2, mae_cliente.nomcli1, mae_cliente.nomcli2, IIf([mae_cliente]![tipper]=1,'',[mae_cliente].[nombre]) AS nomprovedor, Sum(IIf([vta_ventas].[idmon]=1,[vta_ventas]![impbru]+[vta_ventas]![impinaf],([vta_ventas]![impbru]*[con_tc].[impven])+([vta_ventas]![impinaf]*[con_tc].[impven]))) AS total"
'FROM (mae_dociden RIGHT JOIN (mae_tipoempresa RIGHT JOIN (mae_cliente LEFT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON mae_tipoempresa.id = mae_cliente.tipper) ON mae_dociden.id = mae_cliente.tipdoc) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha
'WHERE (((vta_ventas.numreg)<>'000001'))
'Group By '6', mae_tipoempresa.codsun, mae_dociden.codsun, mae_cliente.numruc, mae_cliente.apecli1, mae_cliente.apecli2, mae_cliente.nomcli1, mae_cliente.nomcli2, IIf([mae_cliente]![tipper]=1,'',[mae_cliente].[nombre]), vta_ventas.tipdoc
'HAVING (((Sum(IIf([vta_ventas].[idmon]=1,[vta_ventas]![impbru]+[vta_ventas]![impinaf],([vta_ventas]![impbru]*[con_tc].[impven])+([vta_ventas]![impinaf]*[con_tc].[impven]))))>=6900));
'
    
    Me.MousePointer = vbHourglass
    
    '--consulta de compras
        
    If NulosN(LblIdCliPro.Caption) <> 0 Then nSQLIdCliente = " and vta_ventas.idcli= " & NulosN(LblIdCliPro.Caption) & " "
        
        
    nSQL = "SELECT '" & TIPODOCUMENTOIDEN & "' AS tipdocdecl, '" & NumRUC & "' AS numdocdecla, '" & AnoTra & "' AS anodecla, mae_tipoempresa.codsun AS tipperpro, mae_dociden.codsun AS tipdocpro, mae_cliente.numruc, IIf([mae_cliente]![tipper]=2,'',IIf([mae_cliente]![apecli1] Is Null Or [mae_cliente]![apecli1]='',[mae_cliente]![nombre],[mae_cliente]![apecli1])) AS apepro1, mae_cliente.apecli2 as apepro2, mae_cliente.nomcli1 as nompro1, mae_cliente.nomcli2 as nompro2, IIf(mae_cliente!tipper=1,'',mae_cliente!nombre) AS nomprovedor, Sum(compra.impmn) AS imptotmn, Sum(compra.impme) AS imptotme, Sum(compra.impexpmn) AS imptotexpmn, Sum(compra.impexpme) AS imptotexpme " _
        + vbCr + " FROM (mae_tipoempresa RIGHT JOIN (mae_dociden INNER JOIN mae_cliente ON mae_dociden.id = mae_cliente.idtipdoc) ON mae_tipoempresa.id = mae_cliente.tipper)  " _
        + vbCr + " LEFT JOIN " _
        + vbCr + "( SELECT vta_ventas.idcli, vta_ventas.numreg, mae_documento.abrev, vta_ventas.numser, vta_ventas.numdoc, mae_moneda.simbolo, vta_ventas.impbru, vta_ventas.impbru2, vta_ventas.impbru3, vta_ventas.impinaf,vta_ventas.impisc,(vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf+vta_ventas.impisc) AS impreal, IIf(vta_ventas.tipdoc<>7,impreal,-1*impreal) AS impbase, IIf(vta_ventas.tc=0 Or vta_ventas.tc Is Null,con_tc.impven,vta_ventas.tc) AS tipcam, IIf(vta_ventas.idmon=1,impbase,0) AS impmn, IIf(vta_ventas.idmon=2,impbase,0) AS impme, IIf([vta_ventas].[idmon]=1,[impbase],[impbase]*[tipcam]) AS impexpmn, IIf([vta_ventas].[idmon]=2,[impbase],IIf([tipcam]=0,0,[impbase]/[tipcam])) AS impexpme  " _
                    + vbCr + " FROM mae_moneda RIGHT JOIN ((vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_moneda.id = vta_ventas.idmon " _
                    + vbCr + " WHERE (((vta_ventas.tipdoc)<>2) AND ((vta_ventas.numreg) Not In ('000001','0001'))) " & nSQLIdCliente & " ) as compra  ON mae_cliente.id = compra.idcli " _
        + vbCr + " GROUP BY mae_tipoempresa.codsun, mae_dociden.codsun, mae_cliente.numruc, IIf([mae_cliente]![tipper]=2,'',IIf([mae_cliente]![apecli1] Is Null Or [mae_cliente]![apecli1]='',[mae_cliente]![nombre],[mae_cliente]![apecli1])), mae_cliente.apecli2, mae_cliente.nomcli1, mae_cliente.nomcli2, IIf(mae_cliente!tipper=1,'',mae_cliente!nombre) " _
        + vbCr + " Having (((Sum(compra.impexpmn)) " & CbSimbolo.Text & " " & NulosN(txtBase.Text) & "))" _
        + vbCr + " ORDER BY Sum(compra.impexpmn);"
    
        
        RST_Busq Rst, nSQL, xCon
    
    
    
    
    Fg1.Rows = 2
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        xFila = 2
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(xFila, 1) = A
            
            Fg1.TextMatrix(xFila, 2) = Rst("tipdocdecl")
            Fg1.TextMatrix(xFila, 3) = Rst("numdocdecla")
            Fg1.TextMatrix(xFila, 4) = Rst("anodecla")
            
            Fg1.TextMatrix(xFila, 5) = NulosN(Rst("tipperpro"))
            Fg1.TextMatrix(xFila, 6) = NulosN(Rst("tipdocpro"))
            Fg1.TextMatrix(xFila, 7) = NulosC(Rst("numruc"))
            Fg1.TextMatrix(xFila, 8) = NulosC(Rst("apepro1"))
            Fg1.TextMatrix(xFila, 9) = NulosC(Rst("apepro2"))
            Fg1.TextMatrix(xFila, 10) = NulosC(Rst("nompro1"))
            Fg1.TextMatrix(xFila, 11) = NulosC(Rst("nompro2"))
            Fg1.TextMatrix(xFila, 12) = NulosC(Rst("nomprovedor"))

            Fg1.TextMatrix(xFila, 13) = Format(NulosN(Rst("imptotmn")), FORMAT_MONTO)
            Fg1.TextMatrix(xFila, 14) = Format(NulosN(Rst("imptotme")), FORMAT_MONTO)
            Fg1.TextMatrix(xFila, 15) = Format(NulosN(Rst("imptotexpmn")), FORMAT_MONTO)
            Fg1.TextMatrix(xFila, 16) = Format(NulosN(Rst("imptotexpme")), FORMAT_MONTO)


            Rst.MoveNext
            
            If Rst.EOF = True Then Exit For
            xFila = xFila + 1
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 12) = "TOTAL ==>"
        
        Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(GRID_SUMAR_COL(Fg1, 13), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(GRID_SUMAR_COL(Fg1, 14), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(GRID_SUMAR_COL(Fg1, 15), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(GRID_SUMAR_COL(Fg1, 16), FORMAT_MONTO)
        
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H800000, True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H800000, True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, &H800000, True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 15, &H800000, True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &H800000, True
        
        '--poner color en el fondo
        GRID_COLOR_FONDO Fg1, 2, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HE0FEFE
        GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 12, Fg1.Rows - 1, 16, &HFFFFFF
        

        '--ajustando las columnas de acuerdo a los importes
        Fg1.AutoSizeMode = flexAutoSizeColWidth
        Fg1.AutoSize 13
        Fg1.AutoSize 14
        Fg1.AutoSize 15
        Fg1.AutoSize 16
    End If
    
    Me.MousePointer = vbDefault
    
End Sub

Sub CargarDAOTProveedor()
    Dim A, xFila As Integer
    Dim xTotal As Double
    Dim nSQL As String
    Dim nSQLIdProveedor As String '--sentencia sql para filtrar por proveeedor
    Fg1.Rows = 2
    DoEvents
''
''    RST_Busq Rst, "SELECT '" & TIPODOCUMENTOIDEN & "' AS tipdocdecl, '" & NumRUC & "' AS numdocdecla, '" & AnoTra & "' AS anodecla, mae_tipoempresa.codsun AS tipperpro, mae_dociden.codsun AS tipdocpro, " _
''        & " mae_prov.numruc, IIf([mae_prov]![tipper]=2,'',IIf([mae_prov]![apepro1] Is Null Or [mae_prov]![apepro1]='',[mae_prov]![nombre],[mae_prov]![apepro1])) AS apepro1, mae_prov.apepro2, mae_prov.nompro1, mae_prov.nompro2, IIf(mae_prov!tipper=1,'',mae_prov!nombre) AS nomprovedor, " _
''        & " Sum(IIf(com_compras!idmon=1,com_compras!impina+com_compras!impbru,(com_compras!impina*con_tc!impven)+(com_compras!impbru*con_tc!impven))) AS Total " _
''        & " FROM mae_dociden INNER JOIN ((mae_tipoempresa RIGHT JOIN mae_prov ON mae_tipoempresa.id = mae_prov.tipper) LEFT JOIN (com_compras LEFT JOIN con_tc " _
''        & " ON com_compras.fchdoc = con_tc.fecha) ON mae_prov.id = com_compras.idpro) ON mae_dociden.id = mae_prov.idtipdoc WHERE (((com_compras.tipdoc)<>2) " _
''        & " AND ((com_compras.numreg) not in ('000001','0001'))) " _
''        & " GROUP BY  mae_tipoempresa.codsun, mae_dociden.codsun, mae_prov.numruc, IIf([mae_prov]![tipper]=2,'',IIf([mae_prov]![apepro1] Is Null Or [mae_prov]![apepro1]='',[mae_prov]![nombre],[mae_prov]![apepro1])) , " _
''        & " mae_prov.apepro2, mae_prov.nompro1, mae_prov.nompro2, IIf(mae_prov!tipper=1,'',mae_prov!nombre) Having (((Sum(IIf([com_compras]![idmon] = 1, " _
''        & " [com_compras]![impina] + [com_compras]![impbru], ([com_compras]![impina] * [con_tc]![impven]) + ([com_compras]![impbru] * [con_tc]![impven])))) >= 7100)) " _
''        & " ORDER BY Sum(IIf(com_compras!idmon=1,com_compras!impina+com_compras!impbru,(com_compras!impina*con_tc!impven)+(com_compras!impbru*con_tc!impven)))", xCon
''
''
    
    Me.MousePointer = vbHourglass
    
    If NulosN(LblIdCliPro.Caption) <> 0 Then nSQLIdProveedor = " and com_compras.idpro= " & NulosN(LblIdCliPro.Caption) & " "
    
        '--consulta de compras
    nSQL = "SELECT '" & TIPODOCUMENTOIDEN & "' AS tipdocdecl, '" & NumRUC & "' AS numdocdecla, " & AnoTra & " AS anodecla, mae_tipoempresa.codsun AS tipperpro, mae_dociden.codsun AS tipdocpro, mae_prov.numruc, IIf([mae_prov]![tipper]=2,'',IIf([mae_prov]![apepro1] Is Null Or [mae_prov]![apepro1]='',[mae_prov]![nombre],[mae_prov]![apepro1])) AS apepro1, mae_prov.apepro2, mae_prov.nompro1, mae_prov.nompro2, IIf(mae_prov!tipper=1,'',mae_prov!nombre) AS nomprovedor, Sum(compra.impmn) AS imptotmn, Sum(compra.impme) AS imptotme, Sum(compra.impexpmn) AS imptotexpmn, Sum(compra.impexpme) AS imptotexpme " _
        + vbCr + " FROM (mae_tipoempresa RIGHT JOIN (mae_dociden INNER JOIN mae_prov ON mae_dociden.id = mae_prov.idtipdoc) ON mae_tipoempresa.id = mae_prov.tipper)  " _
        + vbCr + " LEFT JOIN " _
        + vbCr + "( SELECT com_compras.idpro, com_compras.numreg, mae_documento.abrev, com_compras.numser, com_compras.numdoc, mae_moneda.simbolo, com_compras.impbru, com_compras.impbru2, com_compras.impbru3, com_compras.impina, com_compras.impisc, (com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina+com_compras.impisc) AS impreal, IIf(com_compras.tipdoc<>7,impreal,-1*impreal) AS impbase, IIf(com_compras.tc=0 Or com_compras.tc Is Null,con_tc.impven,com_compras.tc) AS tipcam, IIf(com_compras.idmon=1,impbase,0) AS impmn, IIf(com_compras.idmon=2,impbase,0) AS impme, IIf([com_compras].[idmon]=1,[impbase],[impbase]*[tipcam]) AS impexpmn, IIf([com_compras].[idmon]=2,[impbase],IIf([tipcam]=0,0,[impbase]/[tipcam])) AS impexpme  " _
                    + vbCr + " FROM mae_moneda RIGHT JOIN ((com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) ON mae_moneda.id = com_compras.idmon " _
                    + vbCr + " WHERE (((com_compras.tipdoc)<>2) AND ((com_compras.numreg) Not In ('000001','0001'))) " & nSQLIdProveedor & " ) as compra  ON mae_prov.id = compra.idpro " _
        + vbCr + " GROUP BY mae_tipoempresa.id,mae_tipoempresa.codsun, mae_dociden.codsun, mae_prov.numruc, IIf([mae_prov]![tipper]=2,'',IIf([mae_prov]![apepro1] Is Null Or [mae_prov]![apepro1]='',[mae_prov]![nombre],[mae_prov]![apepro1])), mae_prov.apepro2, mae_prov.nompro1, mae_prov.nompro2, IIf(mae_prov!tipper=1,'',mae_prov!nombre) " _
        + vbCr + " HAVING (((mae_tipoempresa.id)<>3) AND ((Sum(compra.impexpmn)) " & CbSimbolo.Text & " " & NulosN(txtBase.Text) & ")) " _
        + vbCr + " ORDER BY Sum(compra.impexpmn);"
        
        RST_Busq Rst, nSQL, xCon
        
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        xFila = 2
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(xFila, 1) = A
            
            Fg1.TextMatrix(xFila, 2) = NulosC(Rst("tipdocdecl"))
            Fg1.TextMatrix(xFila, 3) = NulosC(Rst("numdocdecla"))
            Fg1.TextMatrix(xFila, 4) = NulosC(Rst("anodecla"))
            
            Fg1.TextMatrix(xFila, 5) = NulosC(Rst("tipperpro"))
            Fg1.TextMatrix(xFila, 6) = NulosC(Rst("tipdocpro"))
            Fg1.TextMatrix(xFila, 7) = NulosC(Rst("numruc"))
            Fg1.TextMatrix(xFila, 8) = NulosC(Rst("apepro1"))
            Fg1.TextMatrix(xFila, 9) = NulosC(Rst("apepro2"))
            Fg1.TextMatrix(xFila, 10) = NulosC(Rst("nompro1"))
            Fg1.TextMatrix(xFila, 11) = NulosC(Rst("nompro2"))
            Fg1.TextMatrix(xFila, 12) = NulosC(Rst("nomprovedor"))
            
            Fg1.TextMatrix(xFila, 13) = Format(NulosN(Rst("imptotmn")), FORMAT_MONTO)
            Fg1.TextMatrix(xFila, 14) = Format(NulosN(Rst("imptotme")), FORMAT_MONTO)
            Fg1.TextMatrix(xFila, 15) = Format(NulosN(Rst("imptotexpmn")), FORMAT_MONTO)
            Fg1.TextMatrix(xFila, 16) = Format(NulosN(Rst("imptotexpme")), FORMAT_MONTO)
            
            Rst.MoveNext
            
            If Rst.EOF = True Then Exit For
            xFila = xFila + 1
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 12) = "TOTAL ==>"
        
        Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(GRID_SUMAR_COL(Fg1, 13), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(GRID_SUMAR_COL(Fg1, 14), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(GRID_SUMAR_COL(Fg1, 15), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(GRID_SUMAR_COL(Fg1, 16), FORMAT_MONTO)
        
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H800000, True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H800000, True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, &H800000, True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 15, &H800000, True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &H800000, True
        
        '--poner color en el fondo
        GRID_COLOR_FONDO Fg1, 2, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HE0FEFE
        GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 12, Fg1.Rows - 1, 16, &HFFFFFF
        
        '--ajustando las columnas de acuerdo a los importes
        Fg1.AutoSizeMode = flexAutoSizeColWidth
        
        Fg1.AutoSize 13
        Fg1.AutoSize 14
        Fg1.AutoSize 15
        Fg1.AutoSize 16
        
    End If
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Command3_Click()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, "Declaración Anual Operaciones con Terceros DAOT - " & IIf(Option1.Value = True, "Clientes", "Proveedores"), " ", "Ejercicio " & AnoTra, False, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        txtBase.Text = "0.00"
        txtBase.SetFocus
        SeEjecuto = True
        Option2.Value = True
        CbSimbolo.ListIndex = 0
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    UNIR_CELDAS Fg1, 0, 1, 0, 4, "DATOS DEL DECLARANTE", flexAlignCenterCenter, True
    UNIR_CELDAS Fg1, 0, 5, 0, 12, "DATOS DEL DECLARADO", flexAlignCenterCenter, True
    UNIR_CELDAS Fg1, 0, 13, 0, 14, "TOTALES", flexAlignCenterCenter, True
    UNIR_CELDAS Fg1, 0, 15, 0, 16, "EXPRESADO EN", flexAlignCenterCenter, True
    
    Fg1.Rows = 2
    Fg1.SelectionMode = flexSelectionByRow
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
End Sub


Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Height > 3000 Then
        Fg1.Top = 855
        Fg1.Width = Me.Width - 150
        Fg1.Height = Me.Height - 1250
    End If
End Sub

Private Sub Option1_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
End Sub

Private Sub Option2_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
End Sub

Private Sub txtBase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    
    Select Case KeyAscii
        Case 45
            
        Case Else
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub

Private Sub txtBase_Validate(Cancel As Boolean)
    If IsNumeric(txtBase.Text) = False Then
        MsgBox "Valor incorrecto", vbInformation, xTitulo
        txtBase.Text = "0.00"
        Exit Sub
    End If
    
    txtBase.Text = Format(txtBase.Text, FORMAT_MONTO)
    
End Sub








Private Sub CmdBusCliPro_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    If Option1.Value = True Then
        xform.Titulo = "Buscando Clientes"
        xform.SQLCad = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id From mae_cliente ORDER BY mae_cliente.nombre"
        xCampos(0, 0) = "Cliente":       xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    ElseIf Option2.Value = True Then
        xform.Titulo = "Buscando Proveedores"
        xform.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov ORDER BY mae_prov.nombre"
        xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
        
    End If

    xCampos(1, 0) = "Nº R.U.C.":     xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"

    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCliPro.Text = xRs("nombre")
        LblIdCliPro.Caption = xRs("id")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub OptSel1_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    TxtCliPro.Enabled = False
    CmdBusCliPro.Enabled = False
End Sub

Private Sub OptSel2_Click()
    TxtCliPro.Enabled = True
    CmdBusCliPro.Enabled = True
End Sub
