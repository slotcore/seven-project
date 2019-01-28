VERSION 5.00
Begin VB.Form frmImprimirItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Items"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmImprimirItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3030
      TabIndex        =   12
      Top             =   2850
      Width           =   1440
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   1530
      TabIndex        =   11
      Top             =   2850
      Width           =   1440
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   30
      TabIndex        =   4
      Top             =   -45
      Width           =   5955
      Begin VB.CommandButton cmdFam 
         Height          =   240
         Left            =   5220
         Picture         =   "frmImprimirItem.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   735
         Width           =   240
      End
      Begin VB.CommandButton CmdBusTipiTem 
         Height          =   240
         Left            =   5220
         Picture         =   "frmImprimirItem.frx":043C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   420
         Width           =   240
      End
      Begin VB.TextBox txtFam 
         Height          =   300
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   9
         Top             =   705
         Width           =   3450
      End
      Begin VB.TextBox txtTipIte 
         Height          =   300
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   6
         Top             =   390
         Width           =   3450
      End
      Begin VB.Label lblIdFam 
         Caption         =   "Label1"
         Height          =   195
         Left            =   5535
         TabIndex        =   14
         Top             =   780
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblIdItem 
         Caption         =   "Label1"
         Height          =   195
         Left            =   5520
         TabIndex        =   13
         Top             =   435
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
         Height          =   195
         Index           =   1
         Left            =   435
         TabIndex        =   10
         Top             =   735
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Item"
         Height          =   195
         Index           =   0
         Left            =   435
         TabIndex        =   7
         Top             =   420
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1530
      Left            =   30
      TabIndex        =   0
      Top             =   1245
      Width           =   5940
      Begin VB.OptionButton Option2 
         Caption         =   "Items con Stock Miaximo critico"
         Height          =   255
         Left            =   270
         TabIndex        =   16
         Top             =   1140
         Width           =   3030
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Items con Stock Minimo critico"
         Height          =   255
         Left            =   270
         TabIndex        =   15
         Top             =   855
         Width           =   3030
      End
      Begin VB.OptionButton optInventario 
         Caption         =   "Listar Inventario"
         Height          =   255
         Left            =   4125
         TabIndex        =   3
         Top             =   315
         Width           =   1515
      End
      Begin VB.OptionButton optstock 
         Caption         =   "Listar Stock"
         Height          =   255
         Left            =   2460
         TabIndex        =   2
         Top             =   315
         Width           =   1410
      End
      Begin VB.OptionButton optPrecioStock 
         Caption         =   "Listar Precio y Stock"
         Height          =   255
         Left            =   270
         TabIndex        =   1
         Top             =   315
         Width           =   1800
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   180
         X2              =   5715
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         X1              =   180
         X2              =   5715
         Y1              =   720
         Y2              =   720
      End
   End
End
Attribute VB_Name = "frmImprimirItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMIMPRIMIRITEM.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : EMITE LOS SIGUIENTES REPORTES: LISTAR PRECIO Y STOCK, LISTAR STOCK, LISTAR
'*                    INVENTARIO, USA EL CONTROL DataReport PARA EMITIR LOS REPORTES, DEBERIA DE USARSE
'*                    EL REPORTEADOR DEL COMPONENT ONE
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 08/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Private Sub CmdBusTipiTem_Click()
    'BUSCAMOS EL TIPO DE PRODUCTO QUE SE VA A LISTAR
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_tipoproducto.* FROM mae_tipoproducto"
    
    xform.Titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        txtTipIte.Text = xRs("descripcion")
        lblIdItem.Caption = xRs("id")
        txtFam.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdFam_Click()
    ' SI NO SE HA SELECCIONADO UN TIPO DE ITEM, EL SISTEMA PEDIDIRA EL INGRESO DE UN TIPO DE ITEM
    If txtTipIte.Text = "" Then
        MsgBox "Seleccione en tipo de item", vbInformation, "Mensaje"
        txtTipIte.SetFocus
        Exit Sub
    End If
    
    ' BUSCAMOS LA FAMILIA DEL TIPO DE PRODUCTO QUE SE VA A LISTAR
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_familia.* FROM mae_familia"
    
    xform.Titulo = "Buscando Familia"
    xform.FormaBusca = Principio
    xform.criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        txtFam.Text = xRs("descripcion")
        lblIdFam.Caption = xRs("id")
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

'*****************************************************************************************************
'* Nombre           : Reporte
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL REPORTE EN EL CONTROL DataReport
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Reporte()
    Dim Rst As New ADODB.Recordset
    Dim criterio As String
    
    If Me.txtTipIte.Text = "" Then
        MsgBox "Seleccione el tipo de item", vbInformation, "Mensaje"
        txtTipIte.SetFocus
        Exit Sub
    End If
    
    If txtFam.Text = "" Then
        lblIdFam.Caption = ""
    End If
 
    rptItem.Sections("Sección2").Controls("lblEmp").Caption = NomEmp
    rptItem.Sections("Sección2").Controls("lblRUC").Caption = NumRUC
    rptItem.Sections("Sección2").Controls("lblItem").Caption = txtTipIte.Text
    rptItem.Sections("Sección2").Controls("lblFamilia").Caption = txtFam.Text
    
    ' CARGAMOS LOS DATOS UN FUNCION A LOS CRITERIOS DE BUSQUEDA
    If txtFam.Text = "" Then
        RST_Busq Rst, "SELECT alm_inventario.codpro AS Codigo, alm_inventario.descripcion AS Item, " _
            & " mae_unidades.abrev AS UM, alm_inventario.preuni AS Precio, alm_inventario.stckact AS StockActual, " _
            & " alm_inventario.tippro, alm_inventario.idfam FROM mae_familia RIGHT JOIN " _
            & " (mae_tipoproducto RIGHT JOIN (mae_unidades RIGHT JOIN alm_inventario " _
            & " ON mae_unidades.id = alm_inventario.idunimed) ON mae_tipoproducto.id = alm_inventario.tippro) " _
            & " ON mae_familia.id = alm_inventario.idfam WHERE alm_inventario.tippro= " & Val(lblIdItem.Caption) & " " _
            & " AND alm_inventario.activo = -1 ORDER BY alm_inventario.descripcion", xCon
    Else
        RST_Busq Rst, "SELECT alm_inventario.codpro AS Codigo, alm_inventario.descripcion AS Item, " _
            & " mae_unidades.abrev AS UM, alm_inventario.preuni AS Precio, alm_inventario.stckact AS StockActual, " _
            & " alm_inventario.tippro, alm_inventario.idfam FROM mae_unidades RIGHT JOIN (mae_tipoproducto " _
            & " RIGHT JOIN (mae_familia RIGHT JOIN alm_inventario ON mae_familia.id = alm_inventario.idfam) " _
            & " ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed " _
            & " WHERE alm_inventario.tippro= " & Val(lblIdItem.Caption) & " AND alm_inventario.idfam=" & Val(lblIdFam.Caption) & " " _
            & " AND alm_inventario.activo = -1  ORDER BY alm_inventario.descripcion", xCon
    End If
    
    If optPrecioStock.Value = True Then
        
    End If
    
    If optstock.Value = True Then
        rptItem.Sections("Sección2").Controls("Etiqueta11").Visible = False
        rptItem.Sections("Sección1").Controls("Texto5").Visible = False
    End If
    
    If optInventario.Value = True Then
        rptItem.Sections("Sección2").Controls("Etiqueta11").Caption = "Nuevo Stock"
        rptItem.Sections("Sección1").Controls("Texto5").Visible = False
        'rptItem.Sections("Sección1").Controls("Texto4").Visible = False
        rptItem.Sections("Sección1").Controls("lblStockNuevo").Left = 9480
        rptItem.Sections("Sección1").Controls("lblStockNuevo").Top = 15
        rptItem.Sections("Sección1").Controls("lblStockNuevo").Width = 990
        rptItem.Sections("Sección1").Controls("lblStockNuevo").Visible = -1
        rptItem.Sections("Sección1").Controls("lblStockNuevo").Caption = "__________"
    End If
    
    ' CARGAMOS EL DataReport CON LOS DATOS ESPECIFICADOS
    Set rptItem.DataSource = Rst
    rptItem.Width = 11865
    rptItem.Height = 7980
    rptItem.Show vbModal
End Sub

Private Sub cmdImprimir_Click()
    
    If Option1.Value = True Or Option2.Value = True Then
        ImprimirItemsStock 1, "ITEMS CON STCOK MINIMO CRITICO", "productos"
    Else
        Reporte
    End If
    
End Sub

Sub ImprimirItemsStock(xTipo As Integer, xTitulo As String, xTipoProducto As String)
    ' xTipo = 1 Stock minimo critico
    ' xTipo = 2 Stock maximo critico
    Dim Rst As New ADODB.Recordset
    Dim xFilaInicial As Integer
    
    If xTipo = 1 Then
        RST_Busq Rst, "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckmin, " _
            & " alm_inventario.stckmax, alm_inventario.stckact, alm_inventario.tippro, [alm_inventario]![stckact]-[alm_inventario]![stckmin] AS diferencia, " _
            & " IIf([alm_inventario]![stckact]<=[alm_inventario]![stckmin],1,0) AS critico, alm_inventario.activo " _
            & " FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
            & " Where (((alm_inventario.tippro) = 3) And ((IIf([alm_inventario]![stckact] <= [alm_inventario]![stckmin], 1, 0)) = 1) " _
            & " And ((alm_inventario.activo) = -1)) ORDER BY alm_inventario.descripcion", xCon
    Else
    End If
    
    Dim xFila As Integer
    With FrmPrinter.VS
        .MarginTop = 1000
        .MarginRight = 900
        .BrushColor = &H80000005
        .StartDoc
        CrearCabeceraVS
        
        
        xFilaInicial = 1600
        xFila = xFilaInicial
        .FontSize = 10
        .TextAlign = taCenterMiddle
        FrmPrinter.VS.TextBox xTitulo, 1000, xFila, 10000, 250, True, False, True
        
        .FontSize = 8
        xFila = xFila + 300
        .TextAlign = taLeftMiddle
        FrmPrinter.VS.TextBox "TIPO DE PRODUCTO : " & UCase(xTipoProducto), 1000, xFila, 10000, 200, True, False, True
        
        xFila = xFila + 200
        
        .EndDoc
    End With
    FrmPrinter.Show vbModal
End Sub

Sub CrearCabeceraVS()
    Dim xCad As String
    
    FrmPrinter.VS.TextAlign = taLeftTop
    FrmPrinter.VS.FontName = "Courier New"
    FrmPrinter.VS.FontBold = True
    FrmPrinter.VS.FontSize = 9
    
    FrmPrinter.VS.CurrentX = 1000:      FrmPrinter.VS.CurrentY = 1000
    FrmPrinter.VS.Paragraph = "EMPRESA   : " & NomEmp
    
    FrmPrinter.VS.CurrentX = 8800:      FrmPrinter.VS.CurrentY = 1000
    FrmPrinter.VS.Paragraph = "FECHA     : " & Format(Date, "dd/mm/yy")
    
    FrmPrinter.VS.CurrentX = 1000:      FrmPrinter.VS.CurrentY = 1200
    FrmPrinter.VS.Paragraph = "Nº R.U.C. :  " & NumRUC
    
    FrmPrinter.VS.CurrentX = 8800:      FrmPrinter.VS.CurrentY = 1200
    FrmPrinter.VS.Paragraph = "Nº Pagina : " & "0001"
    
    FrmPrinter.VS.DrawLine 1000, 1450, 11000, 1450
End Sub

