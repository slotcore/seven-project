VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form FrmPrintAsignacPrec 
   Caption         =   "Reporte de asignación de precios"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Height          =   2160
      Left            =   270
      TabIndex        =   0
      Top             =   2865
      Width           =   8610
      _cx             =   15187
      _cy             =   3810
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   8.01424755120214
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
End
Attribute VB_Name = "FrmPrintAsignacPrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMPRINTASIGNACPREC.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA IMPRIMIR LOS PRECIOS ASIGNADOS A LOS ITEMS, A ESTE FORMULARIO SE
'*                    LE LLAMA DESDE FrmMantComPrecios
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 18/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit
Const vConsLeftInic As Long = 1500      ' ALMACENA LA POCISION DERECHA DEL REPORTE
Const vConsTopInic As Long = 2000       ' ALMACENA LA POSICION ARRIBA DEL REPORTE
Dim vTopIni_General As Long
Dim m_Titulo1 As String

Public Property Let propTitulo1(pdata As String)
    m_Titulo1 = pdata
End Property

'*****************************************************************************************************
'* Nombre           : Encabezado
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MANDA AL OBJETO VsPrinter Vp LOS DATOS DE LA CABECERA Y COLUMNAS A IMPRIMIR
'* Paranetros       : NOMBRE    |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pGrid     |  OBJECT      |  ESPECIFIVA EL OBJETO DataGrid QUE SE IMPRIMIRA
'*                    pArr      |  STRING      |  ESPECIFICA EL ENCABEZADO Y LAS COLUMNAS DEL REPORTE
'* Devuelve         :
'*****************************************************************************************************
Sub Encabezado(pGrid As Object, pArr() As String)
    Dim i_colum As Long, vLefInic_Var As Long, vTopInic_Var As Long
    vLefInic_Var = vConsLeftInic
    vTopInic_Var = vConsTopInic
    
    ' IMPRIMIMOS EL ENCABEZADO DEL REPORTE
    vp.FontSize = 8
    vp.TextAlign = taLeftMiddle
    vp.TextBox NomEmp, vConsLeftInic, 600, 2000, 250
    vp.TextBox NumRUC, vConsLeftInic, 800, 2000, 250
    vp.TextAlign = taRightMiddle
    vp.TextBox NomSIS, vConsLeftInic, 600, 10100, 250
    vp.TextBox Date, vConsLeftInic, 800, 10100, 250
    vp.TextAlign = taCenterMiddle
    vp.FontSize = 11
    vp.TextBox m_Titulo1, vConsLeftInic, 1000, 10100, 300
    
    ' IMPRIMIMOS LAS COLUMNAS DEL REPORTE
    vp.TextAlign = taLeftMiddle
    vp.FontSize = 9
    For i_colum = LBound(pArr) To UBound(pArr)
        Select Case pArr(i_colum, 3)
            Case "Descripción"
                vLefInic_Var = vConsLeftInic
                vp.TextBox pArr(i_colum, 3), vLefInic_Var, vTopInic_Var, pArr(i_colum, 4), 300
                vLefInic_Var = vLefInic_Var + pArr(i_colum, 4) + 150
                'valor
                vp.TextBox pArr(i_colum, 1), vLefInic_Var, vTopInic_Var, pArr(i_colum, 2), 300
                vTopInic_Var = vTopInic_Var + 300
            Case "Uni. Medida"
                vLefInic_Var = vConsLeftInic
                vp.TextBox pArr(i_colum, 3), vLefInic_Var, vTopInic_Var, pArr(i_colum, 4), 300
                vLefInic_Var = vLefInic_Var + pArr(i_colum, 4) + 150
                'valor
                vp.TextBox pArr(i_colum, 1), vLefInic_Var, vTopInic_Var, pArr(i_colum, 2), 300
                vTopInic_Var = vTopInic_Var + 300
            Case "Precio Tope"
                vLefInic_Var = vConsLeftInic
                vp.TextBox pArr(i_colum, 3), vLefInic_Var, vTopInic_Var, pArr(i_colum, 4), 300
                vLefInic_Var = vLefInic_Var + pArr(i_colum, 4) + 150
                'valor
                vp.TextBox pArr(i_colum, 1), vLefInic_Var, vTopInic_Var, pArr(i_colum, 2), 300
'                vTopInic_Var = vTopInic_Var + 300
            Case "Tope Max"
                vLefInic_Var = vLefInic_Var + pArr(i_colum, 4) + 150
                vp.TextAlign = taRightMiddle
                vp.TextBox pArr(i_colum, 3), vLefInic_Var, vTopInic_Var, pArr(i_colum, 4), 300
                vLefInic_Var = vLefInic_Var + pArr(i_colum, 4) + 150
                'valor
                vp.TextAlign = taLeftMiddle
                vp.TextBox pArr(i_colum, 1), vLefInic_Var, vTopInic_Var, pArr(i_colum, 2), 300
                vTopInic_Var = vTopInic_Var + 300
            Case "Stock Maximo"
                vLefInic_Var = vConsLeftInic
                vp.TextBox pArr(i_colum, 3), vLefInic_Var, vTopInic_Var, pArr(i_colum, 4), 300
                vLefInic_Var = vLefInic_Var + pArr(i_colum, 4) + 150
                'valor
                vp.TextBox pArr(i_colum, 1), vLefInic_Var, vTopInic_Var, pArr(i_colum, 2), 300
                vTopInic_Var = vTopInic_Var + 300
            Case "Stock Mínimo"
                vLefInic_Var = vConsLeftInic
                vp.TextBox pArr(i_colum, 3), vLefInic_Var, vTopInic_Var, pArr(i_colum, 4), 300
                vLefInic_Var = vLefInic_Var + pArr(i_colum, 4) + 150
                'valor
                vp.TextBox pArr(i_colum, 1), vLefInic_Var, vTopInic_Var, pArr(i_colum, 2), 300
                vTopInic_Var = vTopInic_Var + 500
        End Select
    Next
    vLefInic_Var = vConsLeftInic
    For i_colum = 1 To pGrid.Cols - 1
        Select Case pGrid.TextMatrix(0, i_colum)
            Case "Fecha Reg."
                vp.TextBox pGrid.TextMatrix(0, i_colum), vLefInic_Var, vTopInic_Var, pGrid.ColWidth(i_colum), 250
                vLefInic_Var = vLefInic_Var + pGrid.ColWidth(i_colum) + 200
            Case "Prec. Tope", "Precio", "Dif. de Prec."
                vp.TextBox pGrid.TextMatrix(0, i_colum), vLefInic_Var, vTopInic_Var, pGrid.ColWidth(i_colum), 250
                vLefInic_Var = vLefInic_Var + pGrid.ColWidth(i_colum) + 200
            Case "Proveedor"
                vp.TextBox pGrid.TextMatrix(0, i_colum), vLefInic_Var, vTopInic_Var, pGrid.ColWidth(i_colum), 250
        End Select
    Next
    vp.DrawLine vConsLeftInic, vTopInic_Var, 11000, vTopInic_Var
    vp.DrawLine vConsLeftInic, vTopInic_Var + 300, 11000, vTopInic_Var + 300
    vTopIni_General = vTopInic_Var
    'DIBUJAR LINEA
End Sub

'*****************************************************************************************************
'* Nombre           : Detalle
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MANDA AL OBJETO VsPrinter Vp EL DETALLE DEL REPORTE
'* Paranetros       : NOMBRE    |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pGrid     |  OBJECT      |  SPECIFIVA EL OBJETO DataGrid QUE SE IMPRIMIRA
'*                    pArr      |  STRING      |  ESPECIFICA EL ENCABEZADO Y LAS COLUMNAS DEL REPORTE
'* Devuelve         :
'*****************************************************************************************************
Sub Detalle(pGrid As Object, pArr() As String)
    Dim i_det As Long, i_colum As Integer
    Dim vLefInic_Var As Long, vTopInic_Var As Long
    With pGrid
        vp.PageBorder = pbNone
        vp.StartDoc
        vp.PaperSize = pprA4
        Encabezado pGrid, pArr
        vp.FontSize = 8
        
        ' IMPRIMIMOS EL DETALLE DEL CONTROL DataGrid
        vTopInic_Var = vTopIni_General + 300
        For i_det = 1 To pGrid.Rows - 1
            For i_colum = 1 To pGrid.Cols - 1
                Select Case pGrid.TextMatrix(0, i_colum)
                    Case "Fecha Reg."
                        vp.TextAlign = taLeftMiddle
                        vLefInic_Var = vConsLeftInic
                        vp.TextBox pGrid.TextMatrix(i_det, i_colum), vLefInic_Var, vTopInic_Var, pGrid.ColWidth(i_colum), 250
                        vLefInic_Var = vLefInic_Var + pGrid.ColWidth(i_colum) + 200
                    Case "Prec. Tope"
                        vp.TextAlign = taRightMiddle
                        vp.TextBox pGrid.TextMatrix(i_det, i_colum), vLefInic_Var, vTopInic_Var, pGrid.ColWidth(i_colum), 250
                        vLefInic_Var = vLefInic_Var + pGrid.ColWidth(i_colum) + 200
                    Case "Precio"
                        vp.TextAlign = taRightMiddle
                        vp.TextBox pGrid.TextMatrix(i_det, i_colum), vLefInic_Var, vTopInic_Var, pGrid.ColWidth(i_colum), 250
                        vLefInic_Var = vLefInic_Var + pGrid.ColWidth(i_colum) + 200
                    Case "Dif. de Prec."
                        vp.TextAlign = taRightMiddle
                        vp.TextBox pGrid.TextMatrix(i_det, i_colum), vLefInic_Var, vTopInic_Var, pGrid.ColWidth(i_colum), 250
                        vLefInic_Var = vLefInic_Var + pGrid.ColWidth(i_colum) + 200
                    Case "Proveedor"
                        vp.TextAlign = taLeftMiddle
                        vp.TextBox pGrid.TextMatrix(i_det, i_colum), vLefInic_Var, vTopInic_Var, pGrid.ColWidth(i_colum), 250
                End Select
            Next
            vTopInic_Var = vTopInic_Var + 100

            If vTopInic_Var >= 14500 Then
                .NewPage
                Encabezado pGrid, pArr
                vTopInic_Var = 2300
            Else
                vTopInic_Var = vTopInic_Var + 200
            End If
        Next
    End With
    vp.EndDoc
    vp.ScrollIntoView 0, 0, 0, 0
End Sub

Private Sub Form_Resize()
    ' RECONFIGURA EL TAMÑO DEL FORMULARIO
    vp.Top = 0: vp.Left = 0
    On Error Resume Next
    vp.Height = Me.Height - 500
    vp.Width = Me.Width
End Sub
