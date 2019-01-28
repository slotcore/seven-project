VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmResuMov 
   Caption         =   "Contabilidad - Kardex Resumen"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3180
      TabIndex        =   11
      Top             =   3315
      Visible         =   0   'False
      Width           =   5760
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   210
         Left            =   165
         TabIndex        =   12
         Top             =   390
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
         Caption         =   "Interrumpir = ESC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   4170
         TabIndex        =   19
         Top             =   120
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Registros"
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
         Left            =   165
         TabIndex        =   13
         Top             =   150
         Width           =   1875
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5745
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   705
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   0
         X1              =   5745
         X2              =   5745
         Y1              =   15
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   5745
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6285
      Left            =   15
      TabIndex        =   2
      Top             =   1305
      Width           =   11865
      _cx             =   20929
      _cy             =   11086
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
      BackColorSel    =   8388608
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
      Rows            =   5
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmResuMov.frx":0000
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
      Height          =   1290
      Left            =   15
      TabIndex        =   0
      Top             =   -30
      Width           =   11865
      Begin VB.CheckBox chkActivos 
         Caption         =   "Solo Activos"
         Height          =   195
         Left            =   195
         TabIndex        =   18
         Top             =   990
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Height          =   600
         Left            =   8730
         Picture         =   "FrmResuMov.frx":013F
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Kardex"
         Top             =   405
         Width           =   620
      End
      Begin VB.CommandButton Command5 
         Height          =   600
         Left            =   10050
         Picture         =   "FrmResuMov.frx":0449
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Exportar a Ecxel"
         Top             =   405
         Width           =   620
      End
      Begin VB.CommandButton Command4 
         Height          =   600
         Left            =   7875
         Picture         =   "FrmResuMov.frx":0F53
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Mostrar Resumen"
         Top             =   405
         Width           =   620
      End
      Begin VB.CommandButton Command3 
         Height          =   600
         Left            =   9390
         Picture         =   "FrmResuMov.frx":1395
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir Resumen"
         Top             =   405
         Width           =   620
      End
      Begin VB.CommandButton Command2 
         Height          =   600
         Left            =   10710
         Picture         =   "FrmResuMov.frx":169F
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir"
         Top             =   405
         Width           =   620
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   990
         Left            =   2970
         TabIndex        =   1
         Top             =   210
         Width           =   2685
         _cx             =   4736
         _cy             =   1746
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
         BackColorSel    =   8388608
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmResuMov.frx":19A9
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
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   1020
         TabIndex        =   3
         Top             =   270
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
         Valor           =   "23/03/2007"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   1020
         TabIndex        =   4
         Top             =   600
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
         Valor           =   "23/03/2007"
      End
      Begin VB.Frame Frame3 
         Height          =   1080
         Left            =   5700
         TabIndex        =   15
         Top             =   120
         Width           =   1245
         Begin VB.CommandButton CmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   375
            Left            =   105
            TabIndex        =   17
            Top             =   600
            Width           =   1020
         End
         Begin VB.CommandButton CmdAdd 
            Caption         =   "&Agregar"
            Height          =   375
            Left            =   105
            TabIndex        =   16
            Top             =   195
            Width           =   1020
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   7350
         X2              =   7350
         Y1              =   180
         Y2              =   1230
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   7365
         X2              =   7365
         Y1              =   180
         Y2              =   1230
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   2550
         X2              =   2550
         Y1              =   180
         Y2              =   1230
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   2565
         X2              =   2565
         Y1              =   180
         Y2              =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Final"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   615
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   " Inicial"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   315
         Width           =   450
      End
   End
End
Attribute VB_Name = "FrmResuMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FrmResuMov.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : MUESTRA UN RESUMEN DE LOS MOVIMIENTOS DE INGRESO Y SALIDA DE LOS ITEMS
'*                    REGISTRADOS EN EL ALMACEN, DESDE ESTE FORMULARIO SE INVOCA AL FORMULARIO KARDEX
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 22/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim SeEjecuto As Boolean     ' CONTROLA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE

'*****************************************************************************************************
'* Nombre           : Cargar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA EN EL CONTROL FG1 EL RESUMEN DE MOVIMIENTOS DE LOS ITEMS ESPECIFICADOS POR
'*                    TIPO DE ITEM
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cargar()
    Dim Rst As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim A&, B&, C&
    Dim xTotal As Double
    Dim xCad As String
    Dim nSQLActivos As String       ' Para mostrar solo items activos
    
    Dim StockIni As Double          '--stock incial, depende de la fecha de inicio de consulta
    
    
    Frame2.Left = 3180
    Frame2.Top = 3315
    Frame2.Visible = True
    Fg1.Rows = 1
    DoEvents
    Fg1.Rows = Fg1.Rows + 1
    
    If chkActivos.Value = 1 Then nSQLActivos = " and alm_inventario.activo =-1 "
    Dim xPrecioPromedio As Double
    
    BAND_INTERRUMPIR = False
    
    
    For A = 1 To Fg2.Rows - 1
'        If NulosN(fg2.TextMatrix(A, 2)) = 1 Or NulosN(fg2.TextMatrix(A, 2)) = 4 Or NulosN(fg2.TextMatrix(A, 2)) = 5 Or NulosN(fg2.TextMatrix(A, 2)) = 7 Then
'            RST_Busq Rst, "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckini, " _
'                & " (SELECT Max([preuni]) AS precio From com_comprasdet WHERE (((com_comprasdet.iditem)=alm_inventario.id))) AS precio FROM mae_unidades RIGHT JOIN " _
'                & " alm_inventario ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) = " & NulosN(fg2.TextMatrix(A, 2)) & ")) " & nSQLActivos & " ORDER BY alm_inventario.descripcion", xCon
'        ElseIf NulosN(fg2.TextMatrix(A, 2)) = 2 Or NulosN(fg2.TextMatrix(A, 2)) = 3 Then
'            RST_Busq Rst, "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckini, " _
'                & " (SELECT Max([preuni]) AS precio From vta_ventasdet WHERE (((vta_ventasdet.iditem)=[alm_inventario].[id]))) AS precio FROM mae_unidades RIGHT JOIN " _
'                & " alm_inventario ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) = " & NulosN(fg2.TextMatrix(A, 2)) & "))  " & nSQLActivos & " ORDER BY alm_inventario.descripcion", xCon
'        End If
        
            RST_Busq Rst, "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckini " _
                & " FROM mae_unidades RIGHT JOIN " _
                & " alm_inventario ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) = " & NulosN(Fg2.TextMatrix(A, 2)) & "))  " & nSQLActivos & " ORDER BY alm_inventario.descripcion", xCon
        
        
        If Rst.RecordCount <> 0 Then
            If A = 2 Then Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = "TIPO PRODUCTO  : " + NulosC(Fg2.TextMatrix(A, 1))
            
            Rst.MoveFirst
            
            ProgressBar1.Max = Rst.RecordCount
            For B = 1 To Rst.RecordCount
                xPrecioPromedio = 0
                DoEvents
                '--Validar la interrupcion de la consulta
                If BAND_INTERRUMPIR = True Then GoTo xSalir
                
                Label1.Caption = "Procesando  : " + UCase(NulosC(Fg2.TextMatrix(A, 1)))
                
                ProgressBar1.Value = B
                
                Fg1.Rows = Fg1.Rows + 1
                
                Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst("codpro"))
                Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Rst("descripcion"))
                Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(Rst("abrev"))
                

                '--obtener el saldo inicial
                If CDate(TxtFchIni.Valor) <> CDate("01/01/" & AnoTra) Then
                    StockIni = SaldoActual(NulosN(Rst("id")), NulosC("01/01/" & AnoTra), NulosC(CDate(TxtFchIni.Valor) - 1), xCon)
                Else
'                    StockIni = NulosN(Busca_Codigo("id", NulosN(Rst("id")), "stckini", "alm_inventario", "N", xCon))
                    StockIni = NulosN(Rst("stckini"))
                End If
                
                '--Hallar precio promedio
                xPrecioPromedio = HallarPrecioPromedio(Rst("id"), TxtFchIni.Valor, TxtFchFin.Valor)
                
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(StockIni, "0.00")
                
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(xPrecioPromedio, "0.000000")
                
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = Rst("id")
                
                ' CARGAMOS TODAS LAS ENTRADAS
'''                xCad = "SELECT Sum([canpro]) AS total FROM alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc " _
'''                    & " ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) " _
'''                    & " ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem WHERE (((com_comprasdet.iditem)=" & Rst("id") & ") " _
'''                    & " AND ((com_compras.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (com_compras.fchdoc)<=CDate('" & TxtFchFin.Valor & "')) AND ((com_compras.tipcom)=1)) " _
'''                    & " Union " _
'''                    & " SELECT Sum([cantidad]) AS total FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN " _
'''                    & " (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id " _
'''                    & " WHERE (((alm_ingresodet.iditem)=" & Rst("id") & ") AND ((alm_ingreso.fching)>=CDate('" & TxtFchIni.Valor & "') And (alm_ingreso.fching)<=CDate('" & TxtFchFin.Valor & "')) " _
'''                    & " AND ((alm_ingreso.tipmov)=-1)) " _
'''                    & " Union " _
'''                    & " SELECT Sum([cantidad]) AS total FROM pro_produccion LEFT JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro WHERE  " _
'''                    & " (((pro_producciondet.iditem)=" & Rst("id") & ") AND ((pro_produccion.dia)>=CDate('" & TxtFchIni.Valor & "') And (pro_produccion.dia)<=CDate('" & TxtFchFin.Valor & "')))" _
'''                    & " UNION " _
'''                    & " SELECT Sum(vta_ventasdet.canpro) AS SumaDecanpro FROM vta_ventas RIGHT JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta " _
'''                    & " GROUP BY vta_ventasdet.iditem, vta_ventas.tipdoc, vta_ventas.idmotnotcre, vta_ventas.fchdoc HAVING (((vta_ventasdet.iditem)=" & Rst("id") & ") " _
'''                    & " AND ((vta_ventas.tipdoc)=7) AND ((vta_ventas.idmotnotcre)=4) AND ((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And " _
'''                    & " (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "')))"
'''
'''                xTotal = 0
'''                RST_Busq RstDet, xCad, xCon
'''                If RstDet.RecordCount <> 0 Then
'''                    RstDet.MoveFirst
'''                    For C = 1 To RstDet.RecordCount
'''                        xTotal = xTotal + NulosN(RstDet("total"))
'''                        RstDet.MoveNext
'''                        If Rst.EOF = True Then
'''                            Exit For
'''                        End If
'''                    Next C
'''                End If

                Fg1.TextMatrix(Fg1.Rows - 1, 5) = SaldoActual(Rst("id"), TxtFchIni.Valor, TxtFchFin.Valor, xCon, 1)

                'CARGAMOS TODAS LAS SALIDAS
'''                xCad = " SELECT Sum(vta_ventasdet.canpro) AS total FROM (mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) " _
'''                    & " ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN (vta_ventasdet LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) " _
'''                    & " ON vta_ventas.id = vta_ventasdet.idvta WHERE (((vta_ventasdet.iditem)=" & Rst("id") & ") AND ((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') " _
'''                    & " And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_ventas.oriitem)=1) AND ((vta_ventas.iddocref)=0 Or (vta_ventas.iddocref) Is Null))" _
'''                    & " Union " _
'''                    & " SELECT Sum(alm_ingresodet.cantidad) AS total  FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) " _
'''                    & " LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario  ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id  " _
'''                    & " WHERE (((alm_ingresodet.iditem)=" & Rst("id") & ") AND ((alm_ingreso.fching)>=CDate('" & TxtFchIni.Valor & "') And (alm_ingreso.fching)<=CDate('" & TxtFchFin.Valor & "')) " _
'''                    & " AND ((alm_ingreso.tipmov)=0)) " _
'''                    & " Union " _
'''                    & " SELECT Sum([canpro]) AS total FROM vta_guia LEFT JOIN vta_guiadet ON vta_guia.id = vta_guiadet.idgui WHERE (((vta_guiadet.iditem)=" & Rst("id") & ") " _
'''                    & " AND ((vta_guia.fecgiro)>=CDate('" & TxtFchIni.Valor & "') And (vta_guia.fecgiro)<=CDate('" & TxtFchFin.Valor & "')))" _
'''                    & " UNION " _
'''                    & " SELECT Sum([canutil]) AS total FROM pro_produccion LEFT JOIN pro_producciondetins ON pro_produccion.id = pro_producciondetins.idpro " _
'''                    & " WHERE (((pro_producciondetins.iditem)=" & Rst("id") & ") AND ((pro_produccion.dia)>=CDate('" & TxtFchIni.Valor & "') And (pro_produccion.dia)<=CDate('" & TxtFchFin.Valor & "')))"
'''
'''                xTotal = 0
'''                RST_Busq RstDet, xCad, xCon
'''                If RstDet.RecordCount <> 0 Then
'''                    RstDet.MoveFirst
'''                    For C = 1 To RstDet.RecordCount
'''                        xTotal = xTotal + NulosN(RstDet("total"))
'''                        RstDet.MoveNext
'''                        If Rst.EOF = True Then
'''                            Exit For
'''                        End If
'''                    Next C
'''                End If
                
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(SaldoActual(Rst("id"), TxtFchIni.Valor, TxtFchFin.Valor, xCon, 2), FORMAT_MONTO)
                
                '--Stock Actual
                Fg1.TextMatrix(Fg1.Rows - 1, 7) = (NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 4)) + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 5))) - NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 6))
                Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(Fg1.TextMatrix(Fg1.Rows - 1, 7), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosN(xPrecioPromedio) * NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 7))
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(Fg1.TextMatrix(Fg1.Rows - 1, 9), FORMAT_MONTO)
                
                ' actualizamos el stock actual en la tabla alm_inventario
                xCon.Execute "UPDATE alm_inventario SET alm_inventario.stckact = " & NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 7)) & "" _
                    & " WHERE (((alm_inventario.id)=" & NulosN(Rst("id")) & "))"

                Rst.MoveNext
                If Rst.EOF = True Then
                    Exit For
                End If
            Next B
        End If
    Next A
xSalir:
    Set Rst = Nothing
    Set RstDet = Nothing
    Frame2.Visible = False
End Sub

Private Sub CmdAdd_Click()
    ' AGREGA UNA FILA AL CONTROL Fg2
    If Fg2.Rows = 1 Then
        Fg2.Rows = Fg2.Rows + 1
    Else
        If NulosC(Fg2.TextMatrix(Fg2.Rows - 1, 1)) = "" Then
            Exit Sub
        Else
            Fg2.Rows = Fg2.Rows + 1
        End If
    End If
End Sub

Private Sub CmdEliminar_Click()
    ' ELIMINA UNA FILA AL CONTROL Fg2
    If Fg2.Rows = 1 Then Exit Sub
    Fg2.RemoveItem Fg2.Row
End Sub

Private Sub Command1_Click()
    ' MUESTRA EL KARDEX
    If Fg1.Rows = 1 Then
        FrmVerKardex.Show
        Exit Sub
    End If
    
    Unload FrmVerKardex2
    
    FrmVerKardex.txtCodItem.Text = Fg1.TextMatrix(Fg1.Row, 1)
    FrmVerKardex.LblIdProducto.Caption = Fg1.TextMatrix(Fg1.Row, 10)
    FrmVerKardex.TxtFchIni.Valor = TxtFchIni.Valor
    FrmVerKardex.TxtFchFin.Valor = TxtFchFin.Valor
    FrmVerKardex.Show
End Sub

Private Sub Command2_Click()
    ' SALE DEL FORMULARIO
    Unload Me
End Sub

Private Sub Command3_Click()
    ' MANDA A LA IMPRESORA EL RESUMEN
    If Fg1.Rows = 1 Then
        MsgBox "No hay registros para imprimir", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    FrmPrintKardex.Cargar
    Me.MousePointer = vbDefault
    FrmPrintKardex.Show
End Sub

Private Sub Command4_Click()
    ' MUESTRA INFORMACION DE LOS MOVIMIENTOS DE LOS ITEMS EN EL CONTROL Fg1
    If TxtFchIni.Valor = "" Then
        MsgBox "No ha especificado la fecha de inicio de la consulta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If TxtFchFin.Valor = "" Then
        MsgBox "No ha especificado la fecha final de la consulta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    Dim A As Integer
    
    For A = 1 To Fg2.Rows - 1
        If Fg2.TextMatrix(A, 1) = "" Then
            Fg2.RemoveItem A
        End If
    Next A
    
    If Fg2.Rows = 1 Then
        MsgBox "No ha especificado el tipo de producto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Cargar
End Sub

Private Sub Command5_Click()
    ' EXPORTA A EXCEL LOS DATOS DEL CONTROL Fg1
    If Fg1.Rows = 1 Then
        MsgBox "No se ha mostrado el movimiento de ningun item", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg2.SetFocus
        Exit Sub
    End If
    ExportarExcel
End Sub

Private Sub Fg1_DblClick()
    Command1_Click
End Sub

Private Sub Fg1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1_Click
    End If
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        Dim xform As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        Dim xCampos(2, 4) As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
        
        xform.SqlCad = "SELECT mae_tipoproducto.* FROM mae_tipoproducto"
        
        xform.Titulo = "Buscando Tipo de Producto"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            Fg2.TextMatrix(Fg2.Row, 1) = xRs("descripcion")
            Fg2.TextMatrix(Fg2.Row, 2) = xRs("id")
            
            If Fg2.TextMatrix(Fg2.Row, 1) <> "" Then
                Fg2.Rows = Fg2.Rows + 1
            End If
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        Fg2.Rows = Fg2.Rows + 1
    End If

    If KeyCode = 46 Then
        Fg2.RemoveItem Fg2.Row
    End If
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        If MostrarValorizado = False Then
            Me.Caption = "Almacén - Kardex Resumen"
        Else
            Me.Caption = "Contabilidad - Kardex Resumen"
        End If
        SeEjecuto = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        '--interrumpir
        BAND_INTERRUMPIR = True
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    Fg2.ColWidth(2) = 0
    Fg2.ColComboList(1) = "|..."
    
    Fg2.Rows = 1
    Fg2.Rows = Fg2.Rows + 1
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.Editable = flexEDKbdMouse
    
    Fg1.Editable = flexEDNone
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.AutoSearch = flexSearchFromTop
    Fg1.Rows = 1
    
    
    
    TxtFchIni.Valor = CDate("01/01/" & Year(Date))
    TxtFchFin.Valor = Date
    
    If MostrarValorizado = True Then
        Fg1.ColWidth(8) = 690
        Fg1.ColWidth(9) = 1100
    Else
        Fg1.ColWidth(8) = 0
        Fg1.ColWidth(9) = 0
    End If
    Fg1.ColWidth(10) = 0
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
End Sub

'*****************************************************************************************************
'* Nombre           : ExportarExcel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A EXCEL LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ExportarExcel()
    If Fg1.Rows = 1 Then
        MsgBox "No se ha registrado compras para exportar", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
        Exit Sub
    End If
    
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    ' determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    ' abre el Libro
    objExcel.WindowState = 1
    objExcel.Workbooks.Add
    
    Frame2.Left = 3180
    Frame2.Top = 3315
    Label1.Caption = "Exportando Documentos"
    Frame2.Visible = True
    
    ProgressBar1.Max = Fg1.Rows - 1
    Dim xCadTipItem As String
    
    xCadTipItem = ""
    For A = 1 To Fg2.Rows - 1
        xCadTipItem = xCadTipItem + UCase(Fg2.TextMatrix(A, 1))
        If A = Fg2.Rows - 1 Then Exit For
        xCadTipItem = xCadTipItem & ", "
    Next A
    
    With objExcel.ActiveSheet
        .Cells(1, 2) = NomEmp
        .Cells(1, 13) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        
        .Cells(3, 2) = "PERIODO  :  DEL " & TxtFchIni.Valor & " AL " & TxtFchFin.Valor
        
        .Cells(5, 2) = "RESUMEN DE " & xCadTipItem
        xFilas = 7
        For B = 1 To Fg1.Cols - 1
            .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(0, B)
        Next B
        
        xFilas = xFilas + 1
        For A = 1 To Fg1.Rows - 1
            ProgressBar1.Value = A
            Frame2.Refresh
            
            For B = 1 To Fg1.Cols - 1
                If B <= 3 Then
                    .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                Else
                    If B = 8 Or B = 9 Then
                        If MostrarValorizado = True Then
                            .Cells(xFilas, B + 1) = NulosN(Fg1.TextMatrix(A, B))
                        End If
                    Else
                        .Cells(xFilas, B + 1) = NulosN(Fg1.TextMatrix(A, B))
                    End If
                End If
            Next B
            xFilas = xFilas + 1
        Next A
    End With
    
    Frame2.Visible = False
    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Registro de Compras y Ventas"
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
End Sub




Function HallarPrecioPromedio(IdItem As Double, xFchIni As String, xFchFin As String) As Double
    Dim Rst As New ADODB.Recordset
    Dim xCadSQL As String
    '--alm_ingreso.tipmov -1=Ingreso; 0=Salida
    
    ' PREPARAMOS LA SELECT PARA ARMAR EL KARDEX
''''    xCadSQL = "SELECT alm_ingreso.id, alm_ingresodet.iditem, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS numdoc, " _
''''        & " alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, 'AI' AS tipo, alm_ingreso.nombre AS entidad, 0 AS aa, " _
''''        & " (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos ,'Almacén' & iif(cstr(numdocumentos) ='', ' - Compras','') as modulo  " _
''''        & " FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) " _
''''        & " ON alm_ingreso.id = alm_ingresodet.id WHERE (((alm_ingresodet.iditem)=" & IdItem & ") AND ((alm_ingreso.fching)>=CDate('" & xFchIni & "') " _
''''        & " And (alm_ingreso.fching)<=CDate('" & xFchFin & "')) AND ((alm_ingreso.tipmov)=-1)) " _
''''        & " Union " _
''''        & " SELECT alm_ingreso.id, alm_ingresodet.iditem, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS numdoc, " _
''''        & " alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, 'AS' AS tipo, alm_ingreso.nombre AS entidad, 0 AS aa, " _
''''        & " (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos ,'Almacén' & iif(cstr(numdocumentos) ='', ' - Compras','') as modulo " _
''''        & " FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) " _
''''        & " ON alm_ingreso.id = alm_ingresodet.id WHERE (((alm_ingresodet.iditem)=" & IdItem & ") AND ((alm_ingreso.fching)>=CDate('" & xFchIni & "') " _
''''        & " And (alm_ingreso.fching)<=CDate('" & xFchFin & "')) AND ((alm_ingreso.tipmov)=0))" _
''''        & " Union " _
''''        & " SELECT com_compras.id, com_comprasdet.iditem, alm_inventario.descripcion, com_compras.fchdoc, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, " _
''''        & " com_comprasdet.canpro, IIf([com_compras]![idmon]=2,[com_comprasdet]![preuni]*[con_tc]![impcom],[com_comprasdet]![preuni]) AS preuni, mae_documento.abrev AS descdoc, " _
''''        & " 'C' AS Tipo, mae_prov.nombre AS entidad, 0 AS aa, 0 AS numdocumentos,'Compras' as modulo FROM alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN (com_compras " _
''''        & " LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) " _
''''        & " ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem WHERE (((com_comprasdet.iditem)=" & IdItem & ") AND " _
''''        & " ((com_compras.fchdoc)>=CDate('" & xFchIni & "') And (com_compras.fchdoc)<=CDate('" & xFchFin & "')) AND ((com_compras.tipcom)=1))"
''''
''''    xCadSQL = xCadSQL + "Union " _
''''        & " SELECT vta_guia.id, vta_guiadet.iditem, alm_inventario.descripcion, vta_guia.fecgiro, [vta_guia]![numser]+'-'+[vta_guia]![numdoc] AS numdoc, vta_guiadet.canpro, " _
''''        & " 0 AS preuni, mae_documento.abrev AS desdoc, 'GR' AS tipo, mae_cliente.nombre AS entidad, 0 AS aa, IIf([vta_guia]![iddocven]<>0,1,0) AS numdocumentos,'Guia de Remisión' as modulo " _
''''        & " FROM ((mae_cliente RIGHT JOIN vta_guia ON mae_cliente.id = vta_guia.idcli) LEFT JOIN mae_documento ON vta_guia.tipdoc = mae_documento.id) LEFT JOIN (vta_guiadet " _
''''        & " LEFT JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id) ON vta_guia.id = vta_guiadet.idgui WHERE (((vta_guiadet.iditem)=" & IdItem & ") " _
''''        & " AND ((vta_guia.fecgiro)>=CDate('" & xFchIni & "') And (vta_guia.fecgiro)<=CDate('" & xFchFin & "'))) " _
''''        & " Union " _
''''        & " SELECT pro_produccion.id, pro_producciondetins.iditem, alm_inventario.descripcion, pro_produccion.dia, pro_producciondetins.numparte, pro_producciondetins.canutil, " _
''''        & " 0 AS preuni, 'SM' AS desdoc, 'PS' AS tipo, [alm_inventario_1].[descripcion] AS entidad, pro_producciondet.iditem AS aa, 0 AS numdocumentos,'Producción' as modulo " _
''''        & " FROM (((pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN (pro_producciondetins LEFT JOIN alm_inventario ON pro_producciondetins.iditem = alm_inventario.id) ON (pro_producciondet.idrec = pro_producciondetins.idrec) " _
''''        & " AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro)) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN alm_inventario AS alm_inventario_1 ON pro_receta.iditem = alm_inventario_1.id " _
''''        & " WHERE (((pro_producciondetins.iditem)=" & IdItem & ") AND ((pro_produccion.dia)>=CDate('" & xFchIni & "') And (pro_produccion.dia)<=CDate('" & xFchFin & "')))" _
''''        & " Union " _
''''        & " SELECT pro_produccion.id, pro_producciondet.iditem, alm_inventario.descripcion, pro_produccion.dia, pro_producciondet.numparte, pro_producciondet.cantidad, " _
''''        & " 0 AS preuni, 'PP' AS desdoc, 'P' AS tipo, 'Producción' AS entidad, pro_producciondet.iditem AS aa, 0 AS numdocumentos ,'Producción' as modulo " _
''''        & " FROM pro_produccion INNER JOIN (pro_producciondet LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) ON pro_produccion.id = pro_producciondet.idpro " _
''''        & " WHERE (((pro_producciondet.iditem)=" & IdItem & ") AND ((pro_produccion.dia)>=CDate('" & xFchIni & "') And (pro_produccion.dia)<=CDate('" & xFchFin & "'))) "
''''
''''    xCadSQL = xCadSQL + "Union " _
''''        & " SELECT vta_ventas.id, vta_ventasdet.iditem, alm_inventario.descripcion, vta_ventas.fchdoc, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, " _
''''        & " vta_ventasdet.canpro, vta_ventasdet.preuni, mae_documento.abrev AS descdoc, 'V' AS tipo, mae_cliente.nombre AS entidad, 0 AS aa, 0 AS numdocumentos,'Ventas' as modulo " _
''''        & " FROM mae_cliente RIGHT JOIN ((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) RIGHT JOIN (vta_ventasdet  " _
''''        & " LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta) ON mae_cliente.id = vta_ventas.idcli " _
''''        & " WHERE (((vta_ventasdet.iditem)=" & IdItem & ") " _
''''        & " AND ((vta_ventas.fchdoc)>=CDate('" & xFchIni & "') And (vta_ventas.fchdoc)<=CDate('" & xFchFin & "')) AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) " _
''''        & " AND ((vta_ventas.iddocref) Is Null Or (vta_ventas.iddocref)=0) )" _
''''        & " UNION " _
''''        & " SELECT vta_ventas.id, vta_ventasdet.iditem, alm_inventario.descripcion, vta_ventas.fchdoc, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, " _
''''        & " vta_ventasdet.canpro, vta_ventasdet.preuni, mae_documento.abrev AS descdoc, 'V' AS tipo, mae_cliente.nombre AS entidad, 0 AS aa, 0 AS numdocumentos, " _
''''        & " 'Ventas NC' AS modulo FROM (mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) " _
''''        & " RIGHT JOIN (vta_ventasdet LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta " _
''''        & " WHERE (((vta_ventasdet.iditem)=" & IdItem & ") AND ((vta_ventas.fchdoc)>=CDate('" & xFchIni & "') And (vta_ventas.fchdoc)<=CDate('" & xFchFin & "')) " _
''''        & " AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) AND ((vta_ventas.iddocref)<>0) AND ((vta_ventas.idmotnotcre)=4))"

    xCadSQL = KardexMovimientoSQL(IdItem, 0, TxtFchIni.Valor, TxtFchFin.Valor)
    
    RST_Busq Rst, xCadSQL, xCon
    
    Rst.Sort = "fchdoc"
    Dim xPreIni As Double
    Dim xStckIni As Double
    
    xPreIni = Busca_Codigo(IdItem, "id", "preini", "alm_inventario", "N", xCon)
    xStckIni = Busca_Codigo(IdItem, "id", "stckini", "alm_inventario", "N", xCon)
    If Rst.RecordCount = 0 Then
        HallarPrecioPromedio = xPreIni
    Else
        Dim xTotal, xNuevoTotal, xNuevoSaldoUni As Double
        Dim NuevoPrecio As Double
        Dim A As Integer
        Rst.MoveFirst
        
        xTotal = xPreIni * xStckIni   ' el valor inicial total
        NuevoPrecio = xPreIni
        xNuevoSaldoUni = xStckIni
        For A = 1 To Rst.RecordCount
            If Rst("tipo") = "V" Then
                ' DETERMINAMOS EL NUEVO SALDO DEL ITEM
                xNuevoSaldoUni = xNuevoSaldoUni - Rst("canpro")
                ' CALCULAMOS EL VALOR TOTAL DEL ITEM
                xTotal = xNuevoSaldoUni * NuevoPrecio
                'xTotal = Abs(xTotal)
                ' CALCULAMOS EL NUEVO PRECIO
                If xTotal <> 0 Then
                    NuevoPrecio = (xTotal / xNuevoSaldoUni)
                End If
            Else
                ' DETERMINAMOS EL NUEVO SALDO DEL ITEM
                xNuevoSaldoUni = xNuevoSaldoUni + Rst("canpro")
                
                ' CALCULAMOS EL VALOR TOTAL DEL ITEM
                xNuevoTotal = Rst("canpro") * Rst("preuni")
            
                ' SUMAMOS EL NUEVO VALOR DEL ITEM AL VALOR ANTERIOR
                xTotal = xTotal + xNuevoTotal
                
                ' CALCULAMOS EL NUEVO PRECIO
                If xNuevoSaldoUni = 0 Then
                    NuevoPrecio = 0
                Else
                    NuevoPrecio = (xTotal / xNuevoSaldoUni)
                End If
                xNuevoTotal = 0
            End If
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
        
        HallarPrecioPromedio = NuevoPrecio
    End If
End Function




Sub Cargar_110215()
    Dim Rst As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim A&, B&, C&
    Dim xTotal As Double
    Dim xCad As String
    Dim nSQLActivos As String     ' Para mostrar solo items activos
    
    Frame2.Left = 3180
    Frame2.Top = 3315
    Frame2.Visible = True
    Fg1.Rows = 1
    DoEvents
    Fg1.Rows = Fg1.Rows + 1
    
    If chkActivos.Value = 1 Then nSQLActivos = " and alm_inventario.activo =-1 "
    Dim xPrecioPromedio As Double
    
    For A = 1 To Fg2.Rows - 1
        If NulosN(Fg2.TextMatrix(A, 2)) = 1 Or NulosN(Fg2.TextMatrix(A, 2)) = 4 Or NulosN(Fg2.TextMatrix(A, 2)) = 5 Or NulosN(Fg2.TextMatrix(A, 2)) = 7 Then
            RST_Busq Rst, "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckini, " _
                & " (SELECT Max([preuni]) AS precio From com_comprasdet WHERE (((com_comprasdet.iditem)=alm_inventario.id))) AS precio FROM mae_unidades RIGHT JOIN " _
                & " alm_inventario ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) = " & NulosN(Fg2.TextMatrix(A, 2)) & ")) " & nSQLActivos & " ORDER BY alm_inventario.descripcion", xCon
        End If
        
        If NulosN(Fg2.TextMatrix(A, 2)) = 2 Or NulosN(Fg2.TextMatrix(A, 2)) = 3 Then
            RST_Busq Rst, "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckini, " _
                & " (SELECT Max([preuni]) AS precio From vta_ventasdet WHERE (((vta_ventasdet.iditem)=[alm_inventario].[id]))) AS precio FROM mae_unidades RIGHT JOIN " _
                & " alm_inventario ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) = " & NulosN(Fg2.TextMatrix(A, 2)) & "))  " & nSQLActivos & " ORDER BY alm_inventario.descripcion", xCon
        End If
        
        If Rst.RecordCount <> 0 Then
            If A = 2 Then Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = "TIPO PRODUCTO  : " + NulosC(Fg2.TextMatrix(A, 1))
            
            Rst.MoveFirst
            
            ProgressBar1.Max = Rst.RecordCount
            For B = 1 To Rst.RecordCount
                xPrecioPromedio = 0
                DoEvents
                Label1.Caption = "Procesando  : " + UCase(NulosC(Fg2.TextMatrix(A, 1)))
                ProgressBar1.Value = B
                
                xPrecioPromedio = HallarPrecioPromedio(Rst("id"), TxtFchIni.Valor, TxtFchFin.Valor)
                Fg1.Rows = Fg1.Rows + 1
                'xPrecioPromedio = hallarpreciopromedio()
                Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst("codpro"))
                Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Rst("descripcion"))
                Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(Rst("abrev"))
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(Rst("stckini"), "0.00")
                
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(xPrecioPromedio, "0.000000")
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = Rst("id")
                
                ' CARGAMOS TODAS LAS ENTRADAS
                xCad = "SELECT Sum([canpro]) AS total FROM alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc " _
                    & " ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) " _
                    & " ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem WHERE (((com_comprasdet.iditem)=" & Rst("id") & ") " _
                    & " AND ((com_compras.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (com_compras.fchdoc)<=CDate('" & TxtFchFin.Valor & "')) AND ((com_compras.tipcom)=1)) " _
                    & " Union " _
                    & " SELECT Sum([cantidad]) AS total FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN " _
                    & " (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id " _
                    & " WHERE (((alm_ingresodet.iditem)=" & Rst("id") & ") AND ((alm_ingreso.fching)>=CDate('" & TxtFchIni.Valor & "') And (alm_ingreso.fching)<=CDate('" & TxtFchFin.Valor & "')) " _
                    & " AND ((alm_ingreso.tipmov)=-1)) " _
                    & " Union " _
                    & " SELECT Sum([cantidad]) AS total FROM pro_produccion LEFT JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro WHERE  " _
                    & " (((pro_producciondet.iditem)=" & Rst("id") & ") AND ((pro_produccion.dia)>=CDate('" & TxtFchIni.Valor & "') And (pro_produccion.dia)<=CDate('" & TxtFchFin.Valor & "')))" _
                    & " UNION " _
                    & " SELECT Sum(vta_ventasdet.canpro) AS SumaDecanpro FROM vta_ventas RIGHT JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta " _
                    & " GROUP BY vta_ventasdet.iditem, vta_ventas.tipdoc, vta_ventas.idmotnotcre, vta_ventas.fchdoc HAVING (((vta_ventasdet.iditem)=" & Rst("id") & ") " _
                    & " AND ((vta_ventas.tipdoc)=7) AND ((vta_ventas.idmotnotcre)=4) AND ((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And " _
                    & " (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "')))"
                
                xTotal = 0
                RST_Busq RstDet, xCad, xCon
                If RstDet.RecordCount <> 0 Then
                    RstDet.MoveFirst
                    For C = 1 To RstDet.RecordCount
                        xTotal = xTotal + NulosN(RstDet("total"))
                        RstDet.MoveNext
                        If Rst.EOF = True Then
                            Exit For
                        End If
                    Next C
                End If
                Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(xTotal, FORMAT_MONTO)
                
                'CARGAMOS TODAS LAS SALIDAS
                xCad = " SELECT Sum(vta_ventasdet.canpro) AS total FROM (mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) " _
                    & " ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN (vta_ventasdet LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) " _
                    & " ON vta_ventas.id = vta_ventasdet.idvta WHERE (((vta_ventasdet.iditem)=" & Rst("id") & ") AND ((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') " _
                    & " And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_ventas.oriitem)=1) AND ((vta_ventas.iddocref)=0 Or (vta_ventas.iddocref) Is Null))" _
                    & " Union " _
                    & " SELECT Sum(alm_ingresodet.cantidad) AS total  FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) " _
                    & " LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario  ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id  " _
                    & " WHERE (((alm_ingresodet.iditem)=" & Rst("id") & ") AND ((alm_ingreso.fching)>=CDate('" & TxtFchIni.Valor & "') And (alm_ingreso.fching)<=CDate('" & TxtFchFin.Valor & "')) " _
                    & " AND ((alm_ingreso.tipmov)=0)) " _
                    & " Union " _
                    & " SELECT Sum([canpro]) AS total FROM vta_guia LEFT JOIN vta_guiadet ON vta_guia.id = vta_guiadet.idgui WHERE (((vta_guiadet.iditem)=" & Rst("id") & ") " _
                    & " AND ((vta_guia.fecgiro)>=CDate('" & TxtFchIni.Valor & "') And (vta_guia.fecgiro)<=CDate('" & TxtFchFin.Valor & "')))" _
                    & " UNION " _
                    & " SELECT Sum([canutil]) AS total FROM pro_produccion LEFT JOIN pro_producciondetins ON pro_produccion.id = pro_producciondetins.idpro " _
                    & " WHERE (((pro_producciondetins.iditem)=" & Rst("id") & ") AND ((pro_produccion.dia)>=CDate('" & TxtFchIni.Valor & "') And (pro_produccion.dia)<=CDate('" & TxtFchFin.Valor & "')))"
                
                xTotal = 0
                RST_Busq RstDet, xCad, xCon
                If RstDet.RecordCount <> 0 Then
                    RstDet.MoveFirst
                    For C = 1 To RstDet.RecordCount
                        xTotal = xTotal + NulosN(RstDet("total"))
                        RstDet.MoveNext
                        If Rst.EOF = True Then
                            Exit For
                        End If
                    Next C
                End If
                
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(xTotal, FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 7) = (NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 4)) + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 5))) - NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 6))
                Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(Fg1.TextMatrix(Fg1.Rows - 1, 7), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosN(xPrecioPromedio) * NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 7))
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(Fg1.TextMatrix(Fg1.Rows - 1, 9), FORMAT_MONTO)
                
                ' actualizamos el stock actual en la tabla alm_inventario
                xCon.Execute "UPDATE alm_inventario SET alm_inventario.stckact = " & NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 7)) & "" _
                    & " WHERE (((alm_inventario.id)=" & NulosN(Rst("id")) & "))"

                Rst.MoveNext
                If Rst.EOF = True Then
                    Exit For
                End If
            Next B
        End If
    Next A
    Frame2.Visible = False
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Height > 3000 Then
        Fg1.Top = 1305
        Fg1.Width = Me.Width - 150
        Fg1.Height = Me.Height - 1700 '+300
    End If

End Sub
