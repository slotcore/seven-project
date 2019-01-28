VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmImportaCtaCte 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Importar Documento"
      Height          =   405
      Left            =   8835
      TabIndex        =   12
      Top             =   4590
      Width           =   1755
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   360
      Index           =   1
      Left            =   6105
      TabIndex        =   8
      Top             =   4740
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command3"
      Height          =   360
      Left            =   6105
      TabIndex        =   7
      Top             =   4290
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Importar Ordenes"
      Height          =   450
      Left            =   3660
      TabIndex        =   6
      Top             =   4530
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mostrar Ordenes"
      Height          =   480
      Left            =   2295
      TabIndex        =   5
      Top             =   4530
      Width           =   1335
   End
   Begin VB.CommandButton CmdBusMon 
      Height          =   240
      Left            =   5115
      Picture         =   "FrmImportaCtaCte.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   330
      Width           =   240
   End
   Begin VB.TextBox TxtCliente 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "TxtCliente"
      Top             =   300
      Width           =   5145
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   3435
      Left            =   180
      TabIndex        =   0
      Top             =   990
      Width           =   3630
      _cx             =   6403
      _cy             =   6059
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmImportaCtaCte.frx":0132
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
   Begin VSFlex7Ctl.VSFlexGrid Fg2 
      Height          =   3435
      Left            =   4035
      TabIndex        =   9
      Top             =   990
      Width           =   7665
      _cx             =   13520
      _cy             =   6059
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmImportaCtaCte.frx":01AA
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
      Caption         =   "Documentos de la Orden de Despacho"
      Height          =   195
      Left            =   5460
      TabIndex        =   11
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Ordenes de Despacho"
      Height          =   195
      Left            =   210
      TabIndex        =   10
      Top             =   735
      Width           =   1815
   End
   Begin VB.Label LblIdCliente 
      AutoSize        =   -1  'True
      Caption         =   "LblIdCliente"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3435
      TabIndex        =   4
      Top             =   60
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   210
      Left            =   300
      TabIndex        =   2
      Top             =   45
      Width           =   1455
   End
End
Attribute VB_Name = "FrmImportaCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SeEjecuto As Boolean

Private Sub CmdBusMon_Click()
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
            LblIdCliente.Caption = xRs("id")
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Command1_Click()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Fg1.Rows = 1
    RST_Busq Rst, "SELECT DISTINCT ctacte_8888.idcli, ctacte_8888.cliente, ctacte_8888.orden, ctacte_8888.ordenx From ctacte_8888 " _
        & " Where (((ctacte_8888.idcli) = " & NulosN(LblIdCliente.Caption) & ")) ORDER BY ctacte_8888.orden", xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst("ordenx"))
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

Private Sub Command2_Click()
    Dim A  As Integer
    Dim B As Integer
    Dim RstDet As New ADODB.Recordset
    Dim RstTabla As New ADODB.Recordset
    Dim xId As Double
    Dim xTipDoc As Integer
    
    For A = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(A, 2)) = -1 Then
            Dim RstLGD As New ADODB.Recordset
            Set RstDet = Nothing
            Set RstTabla = Nothing
            Set RstLGD = Nothing
            RST_Busq RstDet, "SELECT ctacte_8888.* From ctacte_8888 WHERE (((ctacte_8888.idcli)=" & NulosN(LblIdCliente.Caption) & ") " _
                & " AND ((ctacte_8888.ordenx)='" & Fg1.TextMatrix(A, 1) & "'))", xCon
            
            RST_Busq RstTabla, "SELECT * FROM  vta_ventas", xCon
            RST_Busq RstLGD, "SELECT * FROM vta_gastodebito", xCon
            If RstDet.RecordCount <> 0 Then
                'IMPORTAMOS TODOS LOS DOCUMENTOS MENOS LAS LETRAS
                For B = 1 To RstDet.RecordCount
                    xTipDoc = 0
                    If RstDet("tipdoc") = "FACTURA DE VENTA" Or RstDet("tipdoc") = "BOLETA DE VENTA" _
                        Or RstDet("tipdoc") = "NOTA DE CREDITO DE VENTA" Or RstDet("tipdoc") = "NOTA DE DEBITOS DE VENTA" Then
                        
                        If RstDet("tipdoc") = "FACTURA DE VENTA" Then xTipDoc = 1
                        If RstDet("tipdoc") = "BOLETA DE VENTA" Then xTipDoc = 3
                        If RstDet("tipdoc") = "NOTA DE CREDITO DE VENTA" Then xTipDoc = 7
                        If RstDet("tipdoc") = "NOTA DE DEBITOS DE VENTA" Then xTipDoc = 8
                        
                        xId = HallaCodigoTabla("vta_ventas", xCon, "id")
                        
                        RstTabla.AddNew
                        RstTabla("id") = xId:                  RstTabla("idlib") = 2:                       RstTabla("idtipo") = 5:              RstTabla("idcli") = NulosN(LblIdCliente.Caption)
                        RstTabla("tipdoc") = xTipDoc:          RstTabla("numdoc") = Mid(RstDet("numdoc"), 1, 7):     RstTabla("fchreg") = "01/01/08":     RstTabla("fchdoc") = RstDet("fchdoc")
                        RstTabla("fchven") = RstDet("fchdoc"): RstTabla("idconpag") = 1:                   RstTabla("idmon") = RstDet("idmon"): RstTabla("imptotdoc") = RstDet("importe")
                        RstTabla("numreg") = "0001":           RstTabla("numerodocref") = RstDet("ordenx"): RstTabla("idmes") = 0:               RstTabla("tc") = 0
                        RstTabla("impsal") = RstDet("importe")
                        RstTabla("numser") = Mid(RstDet("numdoc"), 9, 3)
                        RstTabla.Update
                        GenerarAsiento xCon, 2, xId, 2008, 0
                    End If
                    
                    If RstDet("tipdoc") = "LIQUIDACION GASTOS CREDITO" Or RstDet("tipdoc") = "LIQUIDACION GASTOS DEBITO" Then
                        If RstDet("tipdoc") = "LIQUIDACION GASTOS CREDITO" Then xTipDoc = 126
                        If RstDet("tipdoc") = "LIQUIDACION GASTOS DEBITO" Then xTipDoc = 120
                        
                        xId = HallaCodigoTabla("vta_gastodebito", xCon, "id")
                        
                        RstLGD.AddNew
                        RstLGD("id") = xId:                  RstLGD("tipdoc") = xTipDoc:                RstLGD("numdoc") = Mid(RstDet("numdoc"), 1, 7): RstLGD("fchemi") = RstDet("fchdoc"):
                        RstLGD("idcli") = RstDet("idcli"):   RstLGD("idmon") = RstDet("idmon"):         RstLGD("imptot") = RstDet("importe"): RstLGD("idmes") = 0:
                        RstLGD("idlib") = 41:                RstLGD("fchreg") = "01/01/08":             RstLGD("impsal") = RstDet("importe"):
                        RstLGD("numerodocref") = RstDet("ordenx"): RstLGD("imptotdoc") = RstDet("importe")
                        RstLGD("numser") = Mid(RstDet("numdoc"), 9, 3)
                        RstLGD.Update
                        GenerarAsiento xCon, 41, xId, 2008, 0
                    End If
                    
                    RstDet.MoveNext
                    If RstDet.EOF = True Then Exit For
                Next B
            
            Dim X As Integer
            
                
                Dim RstLet As New ADODB.Recordset
                Dim RstLetDet As New ADODB.Recordset
                
                
                'IMPORTAMOS SOLO LAS LETRAS
            For X = 1 To 2
                Set RstLet = Nothing
                Set RstLetDet = Nothing
            
                Dim C As Integer
                RST_Busq RstLet, "SELECT * FROM let_letra", xCon
                RST_Busq RstLetDet, "SELECT * FROM let_letradet", xCon
                
                RST_Busq RstDet, "SELECT ctacte_8888.* From ctacte_8888 WHERE ((ctacte_8888.idcli=" & NulosN(LblIdCliente.Caption) & ") " _
                    & " AND (ctacte_8888.ordenx='" & Fg1.TextMatrix(A, 1) & "') AND (ctacte_8888.documento='LETRAS POR COBRAR') And (ctacte_8888.idmon=" & X & "))", xCon
                
                If RstDet.RecordCount <> 0 Then
                    xId = HallaCodigoTabla("let_letra", xCon, "id")
                    
                    RstLet.AddNew
                    RstLet("id") = xId:                RstLet("idlib") = 37: RstLet("idtipdoc") = 95:     RstLet("ano") = "2008"
                    RstLet("idmes") = 0:               RstLet("fchemi") = RstDet("fchdoc"):               RstLet("fchini") = RstDet("fchdoc"):
                    RstLet("tiplet") = 1:              RstLet("idclipro") = NulosN(LblIdCliente.Caption): RstLet("numlet") = RstDet.RecordCount
                    RstLet("idmon") = RstDet("idmon"): RstLet("impcap") = 0:                              RstLet("fchreg") = "01/01/08"
                    RstLet("numreg") = "0001":         RstLet("numrefjunto") = RstDet("ordenx"):          RstLet("tc") = 0
                    RstLet("idaduana") = Mid(RstDet("ordenx"), 1, 3)
                    RstLet("idregimen") = Mid(RstDet("ordenx"), 4, 2)
                    RstLet("anoorden") = Mid(RstDet("ordenx"), 6, 4):
                    RstLet("numorden") = Mid(RstDet("ordenx"), 10, 6)
                    
                    RstLet.Update
                     
                    RstDet.MoveFirst
                    For C = 1 To RstDet.RecordCount
                        RstLetDet.AddNew
                        RstLetDet("idlet") = xId:                           RstLetDet("corr") = C:                  RstLetDet("numser") = Mid(RstDet("numdoc"), 14, 3):
                        RstLetDet("numdoc") = Mid(RstDet("numdoc"), 1, 12): RstLetDet("fchemi") = RstDet("fchdoc"): RstLetDet("fchven") = RstDet("fchdoc")
                        RstLetDet("implet") = RstDet("importe")
                        RstLetDet.Update
                        
                        RstDet.MoveNext
                        If RstDet.EOF = True Then Exit For
                    Next C
                End If
            Next X
            
            
            
                'IMPORTAMOS SOLO LAS DEPOSITO ANTICIPO'
                'Dim RstLet As New ADODB.Recordset
                'Dim RstLetDet As New ADODB.Recordset
                        
            For X = 1 To 2
                Set RstLet = Nothing
                Set RstLetDet = Nothing
                'Dim C As Integer
                RST_Busq RstLet, "SELECT * FROM let_letra", xCon
                RST_Busq RstLetDet, "SELECT * FROM let_letradet", xCon
                
                RST_Busq RstDet, "SELECT ctacte_8888.* From ctacte_8888 WHERE (((ctacte_8888.idcli)=" & NulosN(LblIdCliente.Caption) & ") " _
                    & " AND ((ctacte_8888.ordenx)='" & Fg1.TextMatrix(A, 1) & "') AND ((ctacte_8888.documento)='DEPOSITO ANTICIPO') " _
                    & " AND ((ctacte_8888.idmon)=" & X & "))", xCon

                
                '"SELECT ctacte_8888.* From ctacte_8888 WHERE (((ctacte_8888.idcli)=" & NulosN(LblIdCliente.Caption) & ") " _
                    & " AND ((ctacte_8888.ordenx)='" & Fg1.TextMatrix(A, 1) & "') AND ((ctacte_8888.documento)='DEPOSITO ANTICIPO'))", xCon
                
                If RstDet.RecordCount <> 0 Then
                    xId = HallaCodigoTabla("let_letra", xCon, "id")
                    
                    RstLet.AddNew
                    RstLet("id") = xId:                RstLet("idlib") = 37: RstLet("idtipdoc") = 132:     RstLet("ano") = "2008"
                    RstLet("idmes") = 0:               RstLet("fchemi") = RstDet("fchdoc"):               RstLet("fchini") = RstDet("fchdoc"):
                    RstLet("tiplet") = 1:              RstLet("idclipro") = NulosN(LblIdCliente.Caption): RstLet("numlet") = RstDet.RecordCount
                    RstLet("idmon") = RstDet("idmon"): RstLet("impcap") = 0:                              RstLet("fchreg") = "01/01/08"
                    RstLet("numreg") = "0001":         RstLet("numrefjunto") = RstDet("ordenx"):          RstLet("tc") = 0
                    RstLet("idaduana") = Mid(RstDet("ordenx"), 1, 3):       RstLet("idregimen") = Mid(RstDet("ordenx"), 4, 2):
                    RstLet("anoorden") = Mid(RstDet("ordenx"), 6, 4):       RstLet("numorden") = Mid(RstDet("ordenx"), 10, 6)
                    
                    RstLet.Update
                     
                    RstDet.MoveFirst
                    For C = 1 To RstDet.RecordCount
                        RstLetDet.AddNew
                        RstLetDet("idlet") = xId:                           RstLetDet("corr") = C:                  RstLetDet("numser") = "-":
                        RstLetDet("numdoc") = RstDet("numdoc"):             RstLetDet("fchemi") = RstDet("fchdoc"): RstLetDet("fchven") = RstDet("fchdoc")
                        RstLetDet("implet") = RstDet("importe")
                        RstLetDet.Update
                        
                        RstDet.MoveNext
                        If RstDet.EOF = True Then Exit For
                    Next C
                End If
            Next X
                
                'IMPORTAMOS SOLO LAS DEPOSITO CHEQUES MISMO BANCO'
                Set RstLet = Nothing
                Set RstLetDet = Nothing
                
                'Dim C As Integer
            For X = 1 To 2
                Set RstLet = Nothing
                Set RstLetDet = Nothing
            
                RST_Busq RstLet, "SELECT * FROM let_letra", xCon
                RST_Busq RstLetDet, "SELECT * FROM let_letradet", xCon
                
                RST_Busq RstDet, "SELECT ctacte_8888.* From ctacte_8888 WHERE (((ctacte_8888.idcli)=" & NulosN(LblIdCliente.Caption) & ") " _
                    & " AND (ctacte_8888.ordenx='" & Fg1.TextMatrix(A, 1) & "') AND (ctacte_8888.documento='DEPOSITO CHEQUES MISMO BANCO') AND idmon = " & X & ")", xCon
                
                If RstDet.RecordCount <> 0 Then
                    xId = HallaCodigoTabla("let_letra", xCon, "id")
                    
                    RstLet.AddNew
                    RstLet("id") = xId:                RstLet("idlib") = 37: RstLet("idtipdoc") = 133:     RstLet("ano") = "2008"
                    RstLet("idmes") = 0:               RstLet("fchemi") = RstDet("fchdoc"):               RstLet("fchini") = RstDet("fchdoc"):
                    RstLet("tiplet") = 1:              RstLet("idclipro") = NulosN(LblIdCliente.Caption): RstLet("numlet") = RstDet.RecordCount
                    RstLet("idmon") = RstDet("idmon"): RstLet("impcap") = 0:                              RstLet("fchreg") = "01/01/08"
                    RstLet("numreg") = "0001":         RstLet("numrefjunto") = RstDet("ordenx"):          RstLet("tc") = 0
                    RstLet("idaduana") = Mid(RstDet("ordenx"), 1, 3):       RstLet("idregimen") = Mid(RstDet("ordenx"), 4, 2):
                    RstLet("anoorden") = Mid(RstDet("ordenx"), 6, 4):       RstLet("numorden") = Mid(RstDet("ordenx"), 10, 6)
                    
                    RstLet.Update
                     
                    RstDet.MoveFirst
                    For C = 1 To RstDet.RecordCount
                        RstLetDet.AddNew
                        RstLetDet("idlet") = xId:                           RstLetDet("corr") = C:                  RstLetDet("numser") = "-":
                        RstLetDet("numdoc") = RstDet("numdoc"):             RstLetDet("fchemi") = RstDet("fchdoc"): RstLetDet("fchven") = RstDet("fchdoc")
                        RstLetDet("implet") = RstDet("importe")
                        RstLetDet.Update
                        
                        RstDet.MoveNext
                        If RstDet.EOF = True Then Exit For
                    Next C
                End If
            Next X
            
            
                'IMPORTAMOS SOLO LAS DEPOSITO COMPROBANTE DE DETRACCION'
                Set RstLet = Nothing
                Set RstLetDet = Nothing
            
            For X = 1 To 2
                'Dim C As Integer
                Set RstLet = Nothing
                Set RstLetDet = Nothing
                RST_Busq RstLet, "SELECT * FROM let_letra", xCon
                RST_Busq RstLetDet, "SELECT * FROM let_letradet", xCon
                
                RST_Busq RstDet, "SELECT ctacte_8888.* From ctacte_8888 WHERE (((ctacte_8888.idcli)=" & NulosN(LblIdCliente.Caption) & ") " _
                    & " AND (ctacte_8888.ordenx='" & Fg1.TextMatrix(A, 1) & "') AND (ctacte_8888.documento='COMPROBANTE DE DETRACCION') AND idmon = " & X & " )", xCon
                
                If RstDet.RecordCount <> 0 Then
                    xId = HallaCodigoTabla("let_letra", xCon, "id")
                    
                    RstLet.AddNew
                    RstLet("id") = xId:                RstLet("idlib") = 37: RstLet("idtipdoc") = 134:     RstLet("ano") = "2008"
                    RstLet("idmes") = 0:               RstLet("fchemi") = RstDet("fchdoc"):               RstLet("fchini") = RstDet("fchdoc"):
                    RstLet("tiplet") = 1:              RstLet("idclipro") = NulosN(LblIdCliente.Caption): RstLet("numlet") = RstDet.RecordCount
                    RstLet("idmon") = RstDet("idmon"): RstLet("impcap") = 0:                              RstLet("fchreg") = "01/01/08"
                    RstLet("numreg") = "0001":         RstLet("numrefjunto") = RstDet("ordenx"):          RstLet("tc") = 0
                    RstLet("idaduana") = Mid(RstDet("ordenx"), 1, 3):       RstLet("idregimen") = Mid(RstDet("ordenx"), 4, 2):
                    RstLet("anoorden") = Mid(RstDet("ordenx"), 6, 4):       RstLet("numorden") = Mid(RstDet("ordenx"), 10, 6)
                    
                    RstLet.Update
                     
                    RstDet.MoveFirst
                    For C = 1 To RstDet.RecordCount
                        RstLetDet.AddNew
                        RstLetDet("idlet") = xId:                           RstLetDet("corr") = C:                  RstLetDet("numser") = "-":
                        RstLetDet("numdoc") = RstDet("numdoc"):             RstLetDet("fchemi") = RstDet("fchdoc"): RstLetDet("fchven") = RstDet("fchdoc")
                        RstLetDet("implet") = RstDet("importe")
                        RstLetDet.Update
                        
                        RstDet.MoveNext
                        If RstDet.EOF = True Then Exit For
                    Next C
                End If
            Next X
            
            
                'IMPORTAMOS SOLO LAS DEPOSITO COMPROBANTE DE RETENCION'
                Set RstLet = Nothing
                Set RstLetDet = Nothing
            
            For X = 1 To 2
                'Dim C As Integer
                Set RstLet = Nothing
                Set RstLetDet = Nothing
                
                RST_Busq RstLet, "SELECT * FROM let_letra", xCon
                RST_Busq RstLetDet, "SELECT * FROM let_letradet", xCon
                
                RST_Busq RstDet, "SELECT ctacte_8888.* From ctacte_8888 WHERE (((ctacte_8888.idcli)=" & NulosN(LblIdCliente.Caption) & ") " _
                    & " AND (ctacte_8888.ordenx='" & Fg1.TextMatrix(A, 1) & "') AND (ctacte_8888.documento='COMPROBANTE DE RETENCION') AND idmon = " & X & " )", xCon
                
                If RstDet.RecordCount <> 0 Then
                    xId = HallaCodigoTabla("let_letra", xCon, "id")
                    
                    RstLet.AddNew
                    RstLet("id") = xId:                RstLet("idlib") = 37: RstLet("idtipdoc") = 20:     RstLet("ano") = "2008"
                    RstLet("idmes") = 0:               RstLet("fchemi") = RstDet("fchdoc"):               RstLet("fchini") = RstDet("fchdoc"):
                    RstLet("tiplet") = 1:              RstLet("idclipro") = NulosN(LblIdCliente.Caption): RstLet("numlet") = RstDet.RecordCount
                    RstLet("idmon") = RstDet("idmon"): RstLet("impcap") = 0:                              RstLet("fchreg") = "01/01/08"
                    RstLet("numreg") = "0001":         RstLet("numrefjunto") = RstDet("ordenx"):          RstLet("tc") = 0
                    RstLet("idaduana") = Mid(RstDet("ordenx"), 1, 3):       RstLet("idregimen") = Mid(RstDet("ordenx"), 4, 2):
                    RstLet("anoorden") = Mid(RstDet("ordenx"), 6, 4):       RstLet("numorden") = Mid(RstDet("ordenx"), 10, 6)
                    
                    RstLet.Update
                     
                    RstDet.MoveFirst
                    For C = 1 To RstDet.RecordCount
                        RstLetDet.AddNew
                        RstLetDet("idlet") = xId:                           RstLetDet("corr") = C:                  RstLetDet("numser") = "-":
                        RstLetDet("numdoc") = RstDet("numdoc"):             RstLetDet("fchemi") = RstDet("fchdoc"): RstLetDet("fchven") = RstDet("fchdoc")
                        RstLetDet("implet") = RstDet("importe")
                        RstLetDet.Update
                        
                        RstDet.MoveNext
                        If RstDet.EOF = True Then Exit For
                    Next C
                End If
            Next X
            
                ' IMPORTANDO DEPOSITO CHEQUES OTRO BANCO
                Set RstLet = Nothing
                Set RstLetDet = Nothing
            
            For X = 1 To 2
                'Dim C As Integer
                Set RstLet = Nothing
                Set RstLetDet = Nothing
                
                RST_Busq RstLet, "SELECT * FROM let_letra", xCon
                RST_Busq RstLetDet, "SELECT * FROM let_letradet", xCon
                
                RST_Busq RstDet, "SELECT ctacte_8888.* From ctacte_8888 WHERE (((ctacte_8888.idcli)=" & NulosN(LblIdCliente.Caption) & ") " _
                    & " AND (ctacte_8888.ordenx='" & Fg1.TextMatrix(A, 1) & "') AND (ctacte_8888.documento='DEPOSITO CHEQUES OTRO BANCO') AND idmon = " & X & ")", xCon
                
                If RstDet.RecordCount <> 0 Then
                    xId = HallaCodigoTabla("let_letra", xCon, "id")
                    
                    RstLet.AddNew
                    RstLet("id") = xId:                RstLet("idlib") = 37: RstLet("idtipdoc") = 135:     RstLet("ano") = "2008"
                    RstLet("idmes") = 0:               RstLet("fchemi") = RstDet("fchdoc"):               RstLet("fchini") = RstDet("fchdoc"):
                    RstLet("tiplet") = 1:              RstLet("idclipro") = NulosN(LblIdCliente.Caption): RstLet("numlet") = RstDet.RecordCount
                    RstLet("idmon") = RstDet("idmon"): RstLet("impcap") = 0:                              RstLet("fchreg") = "01/01/08"
                    RstLet("numreg") = "0001":         RstLet("numrefjunto") = RstDet("ordenx"):          RstLet("tc") = 0
                    RstLet("idaduana") = Mid(RstDet("ordenx"), 1, 3):       RstLet("idregimen") = Mid(RstDet("ordenx"), 4, 2):
                    RstLet("anoorden") = Mid(RstDet("ordenx"), 6, 4):       RstLet("numorden") = Mid(RstDet("ordenx"), 10, 6)
                    
                    RstLet.Update
                     
                    RstDet.MoveFirst
                    For C = 1 To RstDet.RecordCount
                        RstLetDet.AddNew
                        RstLetDet("idlet") = xId:                           RstLetDet("corr") = C:                  RstLetDet("numser") = "-":
                        RstLetDet("numdoc") = RstDet("numdoc"):             RstLetDet("fchemi") = RstDet("fchdoc"): RstLetDet("fchven") = RstDet("fchdoc")
                        RstLetDet("implet") = RstDet("importe")
                        RstLetDet.Update
                        
                        RstDet.MoveNext
                        If RstDet.EOF = True Then Exit For
                    Next C
                End If
            Next X
            
            
                ' IMPORTANDO DEPOSITO EFECTIVO
                Set RstLet = Nothing
                Set RstLetDet = Nothing
                
            For X = 1 To 2
                'Dim C As Integer
                Set RstLet = Nothing
                Set RstLetDet = Nothing
                RST_Busq RstLet, "SELECT * FROM let_letra", xCon
                RST_Busq RstLetDet, "SELECT * FROM let_letradet", xCon
                
                RST_Busq RstDet, "SELECT ctacte_8888.* From ctacte_8888 WHERE (((ctacte_8888.idcli)=" & NulosN(LblIdCliente.Caption) & ") " _
                    & " AND (ctacte_8888.ordenx='" & Fg1.TextMatrix(A, 1) & "') AND (ctacte_8888.documento='DEPOSITO EFECTIVO') AND idmon = " & X & ")", xCon
                
                If RstDet.RecordCount <> 0 Then
                    xId = HallaCodigoTabla("let_letra", xCon, "id")
                    
                    RstLet.AddNew
                    RstLet("id") = xId:                RstLet("idlib") = 37: RstLet("idtipdoc") = 136:     RstLet("ano") = "2008"
                    RstLet("idmes") = 0:               RstLet("fchemi") = RstDet("fchdoc"):               RstLet("fchini") = RstDet("fchdoc"):
                    RstLet("tiplet") = 1:              RstLet("idclipro") = NulosN(LblIdCliente.Caption): RstLet("numlet") = RstDet.RecordCount
                    RstLet("idmon") = RstDet("idmon"): RstLet("impcap") = 0:                              RstLet("fchreg") = "01/01/08"
                    RstLet("numreg") = "0001":         RstLet("numrefjunto") = RstDet("ordenx"):          RstLet("tc") = 0
                    RstLet("idaduana") = Mid(RstDet("ordenx"), 1, 3):       RstLet("idregimen") = Mid(RstDet("ordenx"), 4, 2):
                    RstLet("anoorden") = Mid(RstDet("ordenx"), 6, 4):       RstLet("numorden") = Mid(RstDet("ordenx"), 10, 6)
                    
                    RstLet.Update
                     
                    RstDet.MoveFirst
                    For C = 1 To RstDet.RecordCount
                        RstLetDet.AddNew
                        RstLetDet("idlet") = xId:                           RstLetDet("corr") = C:                  RstLetDet("numser") = "-":
                        RstLetDet("numdoc") = RstDet("numdoc"):             RstLetDet("fchemi") = RstDet("fchdoc"): RstLetDet("fchven") = RstDet("fchdoc")
                        RstLetDet("implet") = RstDet("importe")
                        RstLetDet.Update
                        
                        RstDet.MoveNext
                        If RstDet.EOF = True Then Exit For
                    Next C
                End If
            Next X
            
            '''''''''''''-------------
                ' IMPORTANDO CHEQUE GIRADO POR DEVOLUCION
                Set RstLet = Nothing
                Set RstLetDet = Nothing
                
            For X = 1 To 2
                'Dim C As Integer
                Set RstLet = Nothing
                Set RstLetDet = Nothing
                RST_Busq RstLet, "SELECT * FROM let_letra", xCon
                RST_Busq RstLetDet, "SELECT * FROM let_letradet", xCon
                
                RST_Busq RstDet, "SELECT ctacte_8888.* From ctacte_8888 WHERE (((ctacte_8888.idcli)=" & NulosN(LblIdCliente.Caption) & ") " _
                    & " AND (ctacte_8888.ordenx='" & Fg1.TextMatrix(A, 1) & "') AND (ctacte_8888.documento='CHEQUE GIRADO POR DEVOLUCION') AND idmon = " & X & ")", xCon
                
                If RstDet.RecordCount <> 0 Then
                    xId = HallaCodigoTabla("let_letra", xCon, "id")
                    
                    RstLet.AddNew
                    RstLet("id") = xId:                RstLet("idlib") = 37: RstLet("idtipdoc") = 137:     RstLet("ano") = "2008"
                    RstLet("idmes") = 0:               RstLet("fchemi") = RstDet("fchdoc"):               RstLet("fchini") = RstDet("fchdoc"):
                    RstLet("tiplet") = 1:              RstLet("idclipro") = NulosN(LblIdCliente.Caption): RstLet("numlet") = RstDet.RecordCount
                    RstLet("idmon") = RstDet("idmon"): RstLet("impcap") = 0:                              RstLet("fchreg") = "01/01/08"
                    RstLet("numreg") = "0001":         RstLet("numrefjunto") = RstDet("ordenx"):          RstLet("tc") = 0
                    RstLet("idaduana") = Mid(RstDet("ordenx"), 1, 3):       RstLet("idregimen") = Mid(RstDet("ordenx"), 4, 2):
                    RstLet("anoorden") = Mid(RstDet("ordenx"), 6, 4):       RstLet("numorden") = Mid(RstDet("ordenx"), 10, 6)
                    
                    RstLet.Update
                     
                    RstDet.MoveFirst
                    For C = 1 To RstDet.RecordCount
                        RstLetDet.AddNew
                        RstLetDet("idlet") = xId:                           RstLetDet("corr") = C:                  RstLetDet("numser") = "-":
                        RstLetDet("numdoc") = RstDet("numdoc"):             RstLetDet("fchemi") = RstDet("fchdoc"): RstLetDet("fchven") = RstDet("fchdoc")
                        RstLetDet("implet") = RstDet("importe")
                        RstLetDet.Update
                        
                        RstDet.MoveNext
                        If RstDet.EOF = True Then Exit For
                    Next C
                End If
            Next X
            
            End If
        End If
    Next A
    MsgBox "Las ordenes de despacho fueron importadas con exito"
End Sub

'Private Sub Command3_Click()
'
'    'GenerarAsiento xCon, 2, 18140, 2008, 0
'End Sub

Private Sub Command4_Click()
    Dim A As Integer
    Dim xNum As Integer
    Dim xNum2 As Double
    xNum2 = 10997
    For A = 1 To 13
        GenerarAsiento xCon, 41, xNum2, 2008, 0
        xNum2 = xNum2 + 1
    Next A
End Sub

Private Sub Command5_Click()
    Dim xTipDoc As Integer
    Dim RstTabla As New ADODB.Recordset
    Dim RstLGD As New ADODB.Recordset
    Dim xId As Double
    
    RST_Busq RstTabla, "SELECT * FROM  vta_ventas", xCon
    RST_Busq RstLGD, "SELECT * FROM vta_gastodebito", xCon
    
    If Fg2.TextMatrix(Fg2.Row, 2) = "FACTURA DE VENTA" Or Fg2.TextMatrix(Fg2.Row, 2) = "BOLETA DE VENTA" _
        Or Fg2.TextMatrix(Fg2.Row, 2) = "NOTA DE CREDITO DE VENTA" Or Fg2.TextMatrix(Fg2.Row, 2) = "NOTA DE DEBITOS DE VENTA" Then
        
        If Fg2.TextMatrix(Fg2.Row, 2) = "FACTURA DE VENTA" Then xTipDoc = 1
        If Fg2.TextMatrix(Fg2.Row, 2) = "BOLETA DE VENTA" Then xTipDoc = 3
        If Fg2.TextMatrix(Fg2.Row, 2) = "NOTA DE CREDITO DE VENTA" Then xTipDoc = 7
        If Fg2.TextMatrix(Fg2.Row, 2) = "NOTA DE DEBITOS DE VENTA" Then xTipDoc = 8
        
        xId = HallaCodigoTabla("vta_ventas", xCon, "id")
        
        RstTabla.AddNew
        RstTabla("id") = xId:                               RstTabla("idlib") = 2:                       RstTabla("idtipo") = 5:
        RstTabla("idcli") = NulosN(LblIdCliente.Caption)
        RstTabla("tipdoc") = xTipDoc:                                        RstTabla("fchdoc") = Fg2.TextMatrix(Fg2.Row, 3)
        RstTabla("numdoc") = Mid(Fg2.TextMatrix(Fg2.Row, 1), 1, 7):          RstTabla("numser") = Mid(Fg2.TextMatrix(Fg2.Row, 1), 9, 7):
        RstTabla("fchven") = Fg2.TextMatrix(Fg2.Row, 3):                     RstTabla("idconpag") = 1:
        RstTabla("idmon") = Fg2.TextMatrix(Fg2.Row, 4):                      RstTabla("imptotdoc") = Fg2.TextMatrix(Fg2.Row, 5)
        RstTabla("numreg") = "0001":                                         RstTabla("numerodocref") = Fg1.TextMatrix(Fg1.Row, 1):
        RstTabla("tc") = 0:                                                  RstTabla("impsal") = Fg2.TextMatrix(Fg2.Row, 5)
        
        'If Format(Fg2.TextMatrix(Fg2.Row, 3), "yyyy") = 2008 Then
        '    RstTabla("idmes") = Format(Fg2.TextMatrix(Fg2.Row, 3), "mm")
        '    RstTabla("fchreg") = "01/" & Format(Fg2.TextMatrix(Fg2.Row, 3), "mm") & "/" & Format(Fg2.TextMatrix(Fg2.Row, 3), "yyyy")
        'Else
            RstTabla("idmes") = 0:
            RstTabla("fchreg") = "01/01/08"
        'End If
        
        RstTabla.Update
        'If Format(Fg2.TextMatrix(Fg2.Row, 3), "yyyy") = 2008 Then
            'GenerarAsiento xCon, 2, xId, 2008,
        'Else
            GenerarAsiento xCon, 2, xId, 2008, 0
        'End If
    End If
    
'    If RstDet("tipdoc") = "LIQUIDACION GASTOS CREDITO" Or RstDet("tipdoc") = "LIQUIDACION GASTOS DEBITO" Then
'        If RstDet("tipdoc") = "LIQUIDACION GASTOS CREDITO" Then xTipDoc = 126
'        If RstDet("tipdoc") = "LIQUIDACION GASTOS DEBITO" Then xTipDoc = 120
'
'        xId = HallaCodigoTabla("vta_gastodebito", xCon, "id")
'
'        RstLGD.AddNew
'        RstLGD("id") = xId:                  RstLGD("tipdoc") = xTipDoc:                RstLGD("numdoc") = RstDet("numdoc"):  RstLGD("fchemi") = RstDet("fchdoc"):
'        RstLGD("idcli") = RstDet("idcli"):   RstLGD("idmon") = RstDet("idmon"):         RstLGD("imptot") = RstDet("importe"): RstLGD("idmes") = 0:
'        RstLGD("idlib") = 41:                RstLGD("fchreg") = "01/01/08":             RstLGD("impsal") = RstDet("importe"):
'        RstLGD("numerodocref") = RstDet("ordenx"): RstLGD("imptotdoc") = RstDet("importe")
'
'        RstLGD.Update
'        GenerarAsiento xCon, 41, xId, 2008, 0
'    End If

End Sub

Private Sub Fg1_RowColChange()
    If Fg1.Rows = 1 Then Exit Sub
    MostrarDocumentos
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        LblIdCliente.Visible = False
        TxtCliente.Text = Busca_Codigo(LblIdCliente.Caption, "id", "nombre", "mae_cliente", "N", xCon)
        Command1_Click
    End If
End Sub

Private Sub Form_Load()
    Fg1.Editable = flexEDKbdMouse
    SeEjecuto = False
    Fg1.Rows = 1
End Sub


Sub MostrarDocumentos()
    Dim rstDoc As New ADODB.Recordset
    Dim A As Integer
    RST_Busq rstDoc, "SELECT ctacte_8888.ordenx, ctacte_8888.numdoc, ctacte_8888.fchdoc, ctacte_8888.importe, ctacte_8888.tipdoc, ctacte_8888.idmon " _
        & " From ctacte_8888 WHERE (((ctacte_8888.ordenx)='" & Fg1.TextMatrix(Fg1.Row, 1) & "') AND ((ctacte_8888.idcli)=" & NulosN(LblIdCliente.Caption) & "))" _
        & " ORDER BY fchdoc ", xCon
    
    Fg2.Rows = 1
    If rstDoc.RecordCount <> 0 Then
        rstDoc.MoveFirst
        For A = 1 To rstDoc.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = rstDoc("numdoc")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = rstDoc("tipdoc")
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = rstDoc("fchdoc")
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = rstDoc("idmon")
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(rstDoc("importe"), "0.00")
            rstDoc.MoveNext
            
            If rstDoc.EOF = True Then Exit For
        Next A
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SeEjecuto = False
End Sub
