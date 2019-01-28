VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInfoKardexValorizado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Informe Kardex Valorizado"
   ClientHeight    =   2160
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   7470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "[ Seleccionar ]"
      Height          =   1710
      Left            =   50
      TabIndex        =   1
      Top             =   400
      Width           =   7380
      Begin VB.CommandButton cmd 
         Height          =   240
         Index           =   1
         Left            =   2100
         Picture         =   "FrmInfoKardexValorizado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1140
         Width           =   240
      End
      Begin VB.CommandButton cmd 
         Height          =   240
         Index           =   0
         Left            =   1950
         Picture         =   "FrmInfoKardexValorizado.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   780
         Width           =   240
      End
      Begin VB.ComboBox cbMes 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   390
         Width           =   2865
      End
      Begin VB.TextBox txtalm 
         Height          =   300
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   5
         Text            =   "txtalm"
         Top             =   750
         Width           =   915
      End
      Begin VB.TextBox txtIdItem 
         Height          =   300
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   9
         Text            =   "txtIdItem"
         Top             =   1110
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ítem"
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   11
         Top             =   1170
         Width           =   300
      End
      Begin VB.Label IdItemLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IdItemLabel"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5520
         TabIndex        =   10
         Top             =   1065
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblalm 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblalm"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2220
         TabIndex        =   7
         Top             =   750
         Width           =   4950
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Almacen"
         Height          =   195
         Left            =   285
         TabIndex        =   6
         Top             =   795
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   285
         TabIndex        =   3
         Top             =   420
         Width           =   540
      End
      Begin VB.Label lblItem 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblItem"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2355
         TabIndex        =   12
         Top             =   1110
         Width           =   4815
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":0264
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":07A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":0B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":0C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":1026
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":11AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":15FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":1716
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":1C5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":219E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":22B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":23C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":281A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInfoKardexValorizado.frx":2986
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu00 
         Caption         =   "Insertar Ítem"
      End
      Begin VB.Menu separador 
         Caption         =   "-"
      End
      Begin VB.Menu menu01 
         Caption         =   "Eliminar Ítem"
      End
   End
End
Attribute VB_Name = "FrmInfoKardexValorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FrmVerKardex.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : MUESTRA EL VINCAR DEL ITEM SELECCIONADO, ADEMAS PERMITE COSTEAS LAS SALIDAS
'*                    MEDIANTE EL METODO PROMEDIO PONDERADO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 23/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim SeEjecuto As Boolean                  ' VARIABLE QUE CONTROLARA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim cSQL As String
Dim F As New SistemaLogica.Funciones
Dim RstInfo As New ADODB.Recordset

Private Sub pIniciarCampos()
    Llenar_Mes cbMes
    Blanquea
End Sub

Sub Blanquea()
    IdItemLabel.Caption = 0
    txtIdItem.Text = ""
    lblItem.Caption = ""
    lblalm.Caption = ""
    txtalm.Text = ""
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim nTitulo As String
    Dim mRecord As New ADODB.Recordset
    Dim xCampos() As String
    Dim xRs As New ADODB.Recordset
     
    Select Case Index
        Case 0 ' Almacen
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
            
            nTitulo = "Buscando Almacenes"
            cSQL = "SELECT alm_almacenes.* FROM alm_almacenes"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            txtalm.Text = NulosN(xRs("id"))
            lblalm.Caption = UCase(NulosC(xRs("descripcion")))
            txtIdItem.SetFocus
            Set xRs = Nothing
            
        Case 1 ' ITEM
            ReDim xCampos(2, 4) As String
            xCampos(0, 0) = "Código":       xCampos(0, 1) = "codpro":       xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "4000":    xCampos(1, 3) = "C"
            
            cSQL = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion " _
                + vbCr + "FROM alm_inventario " _
                + vbCr + "WHERE (((alm_inventario.activo)=-1))"
                             
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando " & nTitulo, "codpro", "codpro", Principio
    
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            IdItemLabel.Caption = F.NuloNumeric(xRs("id"))
            txtIdItem.Text = F.NuloString(xRs("codpro"))
            lblItem.Caption = F.NuloString(xRs("descripcion"))
            lblItem.ToolTipText = F.NuloString(xRs("descripcion"))
    End Select
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    pIniciarCampos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pCargarDatos
    
    If Button.Index = 5 Then
        Unload Me
    End If
End Sub


Sub CrearRecordSet(ByRef RST_ As ADODB.Recordset)
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(14, 3) As String
    ' N: Numerico
    ' D: Double
    ' F: Fecha
    ' C: Caracter
    ' L: Logico
    xCampos(0, 0) = "FECHA":            xCampos(0, 1) = "C":      xCampos(0, 2) = "100"
    xCampos(1, 0) = "TIPO":             xCampos(1, 1) = "C":      xCampos(1, 2) = "100"
    xCampos(2, 0) = "SERIE":            xCampos(2, 1) = "C":      xCampos(2, 2) = "100"
    xCampos(3, 0) = "NÚMERO":           xCampos(3, 1) = "C":      xCampos(3, 2) = "100"
    xCampos(4, 0) = "TIPO OPERACION":    xCampos(4, 1) = "C":      xCampos(4, 2) = "100"
    xCampos(5, 0) = "ENTRADAS - CANTIDAD":      xCampos(5, 1) = "D":      xCampos(5, 2) = ""
    xCampos(6, 0) = "ENTRADAS - COSTO UNITARIO":      xCampos(6, 1) = "D":      xCampos(6, 2) = ""
    xCampos(7, 0) = "ENTRADAS - COSTO TOTAL":         xCampos(7, 1) = "D":      xCampos(7, 2) = ""
    xCampos(8, 0) = "SALIDAS - CANTIDAD":      xCampos(8, 1) = "D":      xCampos(8, 2) = ""
    xCampos(9, 0) = "SALIDAS - COSTO UNITARIO":      xCampos(9, 1) = "D":      xCampos(9, 2) = ""
    xCampos(10, 0) = "SALIDAS - COSTO TOTAL":         xCampos(10, 1) = "D":      xCampos(10, 2) = ""
    xCampos(11, 0) = "SALDO FINAL - CANTIDAD":    xCampos(11, 1) = "D":      xCampos(11, 2) = ""
    xCampos(12, 0) = "SALDO FINAL - COSTO UNITARIO":    xCampos(12, 1) = "D":      xCampos(12, 2) = ""
    xCampos(13, 0) = "SALDO FINAL - COSTO TOTAL":       xCampos(13, 1) = "D":      xCampos(13, 2) = ""
    'xCampos(14, 0) = "SALDO FINAL - COSTO UNITARIO PROMEDIO":     xCampos(14, 1) = "C":      xCampos(14, 2) = "100"
    
    Set RST_ = xFun.CrearRstTMP(xCampos)
    RST_.Open
End Sub

Sub LLenarRecordSet(IdAlmacen As Long, _
                            IdItem As Long, _
                            FechaInicio As Date, _
                            FechaFin As Date)
    Dim xCadSQL As String
    Dim UltPreCosto As Double
    Dim xPrecioUni As Double
    Dim xPrecioUniProm As Double
    Dim StockIni As Double
    Dim xPrecioIni As Double
    Dim rst As New ADODB.Recordset
    Dim xSaldo As Double
    Dim xSaldoImp As Double
    Dim A&
    Dim xFila As Integer
    Dim xTotSal, xTotEnt As Double
    
    Set RstInfo = Nothing
    CrearRecordSet RstInfo
    '********************
    ' Saldo Inicial
    '********************
    xCadSQL = F.SQL_MovHistoricoTotalizado(F.NuloNumeric(IdAlmacen), FechaInicio - 1, CStr(IdItem), xCon, True)
    Set rst = Nothing
    Set rst = F.GeneraRstSQL(xCadSQL, xCon)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        
        RstInfo.AddNew
        RstInfo("FECHA") = ""
        RstInfo("TIPO") = "SALDO INICIAL"
        RstInfo("SERIE") = ""
        RstInfo("NÚMERO") = ""
        RstInfo("TIPO OPERACION") = ""
        RstInfo("ENTRADAS - CANTIDAD") = Format(NulosN(rst("canini")) + NulosN(rst("canent")) - NulosN(rst("cansal")), FORMAT_MONTO)
        RstInfo("ENTRADAS - COSTO UNITARIO") = Format(NulosN(rst("costouniprom")), FORMAT_IMPORTEKARDEX)
        RstInfo("ENTRADAS - COSTO TOTAL") = Format(NulosN(rst("costoini")) + NulosN(rst("costoent")) - NulosN(rst("costosal")), FORMAT_IMPORTEKARDEX)
        RstInfo("SALIDAS - CANTIDAD") = 0
        RstInfo("SALIDAS - COSTO UNITARIO") = 0
        RstInfo("SALIDAS - COSTO TOTAL") = 0
        RstInfo("SALDO FINAL - CANTIDAD") = Format((rst("canini")) + NulosN(rst("canent")) - NulosN(rst("cansal")), FORMAT_IMPORTEKARDEX)
        RstInfo("SALDO FINAL - COSTO UNITARIO") = Format(NulosN(rst("costouniprom")), FORMAT_IMPORTEKARDEX)
        RstInfo("SALDO FINAL - COSTO TOTAL") = Format(NulosN(rst("costoini")) + NulosN(rst("costoent")) - NulosN(rst("costosal")), FORMAT_IMPORTEKARDEX)
        
        StockIni = NulosN(rst("canini")) + NulosN(rst("canent")) - NulosN(rst("cansal"))
        xPrecioIni = NulosN(rst("costouniprom"))
    Else
        StockIni = 0
        xPrecioIni = 0
    End If
    
    '*************
    ' Movimientos
    '*************
    xCadSQL = F.SQL_MovDetallado(CStr(IdItem), F.NuloNumeric(IdAlmacen), FechaInicio, FechaFin, xCon, , False, , , True)
    Set rst = Nothing
    Set rst = F.GeneraRstSQL(xCadSQL, xCon)
    
    
    UltPreCosto = xPrecioIni
    xPrecioUniProm = xPrecioIni
    xPrecioUni = xPrecioIni
    xSaldo = StockIni
    xSaldoImp = xSaldo * xPrecioIni
    xTotEnt = xTotEnt + StockIni
            
    If rst.RecordCount <> 0 Then
        rst.MoveFirst

        For A = 1 To rst.RecordCount
        
            RstInfo.AddNew
            RstInfo("FECHA") = Format(rst("fchmov"), "dd/mm/yy")
            RstInfo("TIPO") = NulosC(rst("doc"))
            RstInfo("SERIE") = NulosC(rst("numser"))
            RstInfo("NÚMERO") = NulosC(rst("numdoc"))
        
            If NulosN(rst("tipmov") = 0) Then
                RstInfo("TIPO OPERACION") = "ALMACEN SALIDA"
            Else
                RstInfo("TIPO OPERACION") = "ALMACEN INGRESO"
            End If
            
            ' ----------------------------------------------INGRESOS
            If NulosN(rst("tipmov")) = -1 Then
        
                RstInfo("SALIDAS - CANTIDAD") = 0
                RstInfo("SALIDAS - COSTO UNITARIO") = 0
                RstInfo("SALIDAS - COSTO TOTAL") = 0
        
                RstInfo("ENTRADAS - CANTIDAD") = Format(NulosN(rst("cantidad")), FORMAT_MONTO)
                xSaldo = xSaldo + NulosN(rst("cantidad"))
                xTotEnt = xTotEnt + NulosN(rst("cantidad"))
                                
                RstInfo("SALDO FINAL - CANTIDAD") = Format(xSaldo, FORMAT_MONTO)
                
                If F.NuloNumeric(rst("cantidad")) > 0 Then
                    xPrecioUni = F.NuloNumeric(rst("costo")) / F.NuloNumeric(rst("cantidad"))
                Else
                    xPrecioUni = 0
                End If
                                                
                ' --------------------------------PRECIO UNITARIO
                RstInfo("ENTRADAS - COSTO UNITARIO") = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
                
                RstInfo("ENTRADAS - COSTO TOTAL") = Format(NulosN(rst("costo")), FORMAT_IMPORTEKARDEX)
                xSaldoImp = xSaldoImp + NulosN(rst("costo"))
                
                RstInfo("SALDO FINAL - COSTO UNITARIO") = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
                RstInfo("SALDO FINAL - COSTO TOTAL") = Format(xSaldoImp, FORMAT_IMPORTEKARDEX)
                
                ' --------------------------------PRECIO PROMEDIO
                If xSaldo > 0 Then
                    xPrecioUniProm = xSaldoImp / xSaldo
                Else
                    xPrecioUniProm = 0
                End If
                'RstInfo("SALDO FINAL - COSTO UNITARIO PROMEDIO") = Format(xPrecioUniProm, FORMAT_IMPORTEKARDEX)
                UltPreCosto = xPrecioUni
            
            ' ----------------------------------------------------------SALIDAS
            Else
            
                RstInfo("ENTRADAS - CANTIDAD") = 0
                RstInfo("ENTRADAS - COSTO UNITARIO") = 0
                RstInfo("ENTRADAS - COSTO TOTAL") = 0
            
                If F.NuloNumeric(rst("cantidad")) > 0 Then
                    xPrecioUni = F.NuloNumeric(rst("costo")) / F.NuloNumeric(rst("cantidad"))
                Else
                    xPrecioUni = 0
                End If
                
                RstInfo("SALIDAS - CANTIDAD") = Format(NulosN(rst("cantidad")), FORMAT_MONTO)
                xSaldo = xSaldo - NulosN(rst("cantidad"))
                xTotSal = xTotSal + NulosN(rst("cantidad"))
                
                '--saldo x cantidad
                RstInfo("SALDO FINAL - CANTIDAD") = Format(xSaldo, FORMAT_MONTO)
                    
                ' ----------------------PRECIO UNITARIO
                RstInfo("SALDO FINAL - COSTO UNITARIO") = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
                RstInfo("SALIDAS - COSTO UNITARIO") = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
                
                RstInfo("SALIDAS - COSTO TOTAL") = Format(NulosN(rst("costo")), FORMAT_IMPORTEKARDEX)
                xSaldoImp = xSaldoImp - (NulosN(rst("cantidad")) * xPrecioUni)
                '--saldo
                RstInfo("SALDO FINAL - COSTO TOTAL") = Format(xSaldoImp, FORMAT_IMPORTEKARDEX)
                
                ' -------------PRECIO PROMEDIO
                'RstInfo("SALDO FINAL - COSTO UNITARIO PROMEDIO") = Format(xPrecioUniProm, FORMAT_IMPORTEKARDEX)
                
                If xSaldo = 0 Then
                    xPrecioUniProm = 0
                End If
            End If
            rst.MoveNext
            If rst.EOF = True Then
                Exit For
            End If
            
            xFila = xFila + 1
        Next A
    End If
End Sub


Private Sub pCargarDatos()
    Dim mRecord As New ADODB.Recordset
    Dim mDataBase As New SistemaData.EDataBase
    Dim oExport As New SGI2_funciones.formularios
    Dim xCampos() As String
    Dim mUltimoDiaMes As Date
    Dim mPrimerDiaMes As Date
    Dim mMesActual As Integer
    
    If fValidarDatos() = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    Set mDataBase.Connection = xCon
    
    mMesActual = cbMes.ListIndex + 1
    mPrimerDiaMes = F.RetornarPrimerDiaMes(CDate("01/" & mMesActual & "/" & AnoTra & ""))
    mUltimoDiaMes = F.RetornarUltimoDiaMes(mPrimerDiaMes)
    
    LLenarRecordSet F.NuloNumeric(txtalm.Text), F.NuloNumeric(IdItemLabel.Caption), mPrimerDiaMes, mUltimoDiaMes
            
    ' Exportar a excel el recordset
    F.ExportarExcelRecordSet RstInfo
    
'    Set RptLibroKardexDetallado.DataSource = RstInfo
'    RptLibroKardexDetallado.Show

'    ' Exportar Excel Reporte
'    RptLibroKardexDetallado.ExportReport rptKeyHTML, "c:\report.html"
'    Dim Abrir_Excel As Object
'    Set Abrir_Excel = CreateObject("Excel.Application")
'    Abrir_Excel.Visible = True
'    Abrir_Excel.Workbooks.Open ("c:\report.html")
'    Abrir_Excel.Windows("report.html").Activate
'    Abrir_Excel.Sheets("report").Select
    
    Me.MousePointer = vbDefault
End Sub

'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : VERIFICA QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    If F.NuloNumeric(txtalm.Text) = 0 Then
        F.MostrarMensajeError "Debe de seleccionar un almacén", "Error"
        txtalm.SetFocus
        fValidarDatos = False
        Exit Function
    End If
    If F.NuloNumeric(IdItemLabel.Caption) = 0 Then
        F.MostrarMensajeError "Debe de seleccionar un item", "Error"
        txtIdItem.SetFocus
        fValidarDatos = False
        Exit Function
    End If
    fValidarDatos = True
End Function
