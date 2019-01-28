VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form FrmPrintConsComAdmin 
   Caption         =   "Form3"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12345
   LinkTopic       =   "Form3"
   ScaleHeight     =   7215
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VSPrinter7LibCtl.VSPrinter Vp 
      Height          =   4275
      Left            =   195
      TabIndex        =   0
      Top             =   2205
      Width           =   11655
      _cx             =   20558
      _cy             =   7541
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
      Zoom            =   22.6692836113837
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
Attribute VB_Name = "FrmPrintConsComAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMPRINTCONSCOMADMIN.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA IMPRIMIR INFORMACION DEL FORMULARIO FrmComsCompra_Administ, ESTE
'*                    FOMULARIO SE INVOCA DESDE FrmComsCompra_Administ
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 18/09/09
'* VERSION          : 1.0
'*****************************************************************************************************

Option Explicit
' VARIABLES PARA EL REPORTE
Dim Consposinileft As Long
Dim Consposinitop As Long

' SOLO PARA LA PARTE SUPERIOR DE LA PAGINA
Const consposinittop_enc As Long = 600
Const vsepentrecol As Integer = 50
Const vsepentrerow As Integer = 200
Const cons_ancho_pagVert As Long = 12000
Const cons_ancho_pagHor As Long = 15800

Dim arrayenc() As String, vContCol_Usar As Integer
Dim vLonLinea As Long, vOrientPag As Integer
Dim vindenc As Boolean                                ' PARA INDICAR SI VA CON ENCABEZADO SI NO NO
Dim xfrm As Form
Dim m_Titulo1 As String, m_titulo2 As String          ' VARIBLES PARA LAS PROPIEDADES

Public Property Let propTitulo1(pdata As String)
    m_Titulo1 = pdata
End Property

Public Property Let proptitulo2(pdata As String)
    m_titulo2 = pdata
End Property

'*****************************************************************************************************
'* Nombre           : fdetertipdatoNumeric
'* Tipo             : FUNCION
'* Descripcion      : VERIFICA SI UN DATO ES NUMERICO, DEVUELVE VERDADERO SI EL DATOS ES NUMERICO
'* Paranetros       : NOMBRE    |  TIPO         |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pvalor    |  VARIANT      |  VALOR A COMPARAR
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function fdetertipdatoNumeric(pvalor As Variant) As Boolean
    If IsNumeric(pvalor) = True Then
        fdetertipdatoNumeric = True
    Else
        fdetertipdatoNumeric = False
    End If
End Function

'*****************************************************************************************************
'* Nombre           : inicializar_var_decgen
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : INICIALIZA LAS VARIABLES DE POSICION
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub inicializar_var_decgen()
    Consposinileft = 0: Consposinitop = 0
    vContCol_Usar = 0
    vLonLinea = 0: vOrientPag = 0: vindenc = True
End Sub

Function capform(pform As Form)
    Set xfrm = pform
End Function

Private Function fdevuelvepos(pindexlimite As Integer) As Long
    Dim vdevpos As Long, i_devpos
    vdevpos = Consposinileft
    For i_devpos = 1 To pindexlimite - 1
        vdevpos = vdevpos + arrayenc(i_devpos, 2) + vsepentrecol
    Next
    fdevuelvepos = vdevpos
End Function

'*****************************************************************************************************
'* Nombre           : Detalle
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME EL DETALLE DEL REPORTE
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Detalle()
    Dim i_det As Long, x_det As Long
    vp.FontSize = 7
    Dim arrdet() As String, vposinitop As Long
    Dim vposinterm As Long, vdeterfila As Long, vposinileft As Long
    Dim vindexlimite As Integer
    vposinitop = Consposinitop + 500
    For i_det = 2 To FrmConsCompra.Fg1.Rows - 1
        vposinileft = Consposinileft
        arrdet = fDeterFila(i_det)
        If arrdet(1, 2) <> "" Then
            vindexlimite = CInt(arrdet(1, 2))
        End If
        If arrdet(1, 1) = "TOT" Then
            vposinterm = fdevuelvepos(vindexlimite)
            vdeterfila = arrdet(1, 2)
            For x_det = vdeterfila To vContCol_Usar
                If x_det = vdeterfila Then
                    vp.TextBox "Total: ", vposinterm, vposinitop, arrayenc(x_det, 2), 200
                Else
                    If Trim(FrmConsCompra.Fg1.TextMatrix(i_det, arrayenc(x_det, 3))) <> "" Then
                        If fdetertipdatoNumeric(FrmConsCompra.Fg1.TextMatrix(i_det, arrayenc(x_det, 3))) = True Then
                            vp.TextAlign = taRightMiddle
                        End If
                        vp.TextBox FrmConsCompra.Fg1.TextMatrix(i_det, arrayenc(x_det, 3)), vposinterm, vposinitop, arrayenc(x_det, 2), 200
                        vp.TextAlign = taLeftMiddle
                    End If
                End If
                vposinterm = vposinterm + arrayenc(x_det, 2) + vsepentrecol
            Next
        ElseIf arrdet(1, 1) = "TOTGEN" Then
            vposinterm = fdevuelvepos(vindexlimite)
            vdeterfila = arrdet(1, 2)
            For x_det = vdeterfila To vContCol_Usar
                If x_det = vdeterfila Then
                    vp.TextBox "Tot. Gen.: ", vposinterm, vposinitop, arrayenc(x_det, 2), 200
                Else
                    If Trim(FrmConsCompra.Fg1.TextMatrix(i_det, arrayenc(x_det, 3))) <> "" Then
                        If fdetertipdatoNumeric(FrmConsCompra.Fg1.TextMatrix(i_det, arrayenc(x_det, 3))) = True Then
                            vp.TextAlign = taRightMiddle
                        End If
                        vp.TextBox FrmConsCompra.Fg1.TextMatrix(i_det, arrayenc(x_det, 3)), vposinterm, vposinitop, arrayenc(x_det, 2), 200
                        vp.TextAlign = taLeftMiddle
                    End If
                End If
                vposinterm = vposinterm + arrayenc(x_det, 2) + vsepentrecol
            Next
        ElseIf arrdet(1, 1) = "ENC" Then
            vp.TextBox FrmConsCompra.Fg1.TextMatrix(i_det, arrdet(1, 2)), Consposinileft, vposinitop, 2000, 200
        ElseIf arrdet(1, 1) = "" Then
            For x_det = 1 To vContCol_Usar
                If fdetertipdatoNumeric(FrmConsCompra.Fg1.TextMatrix(i_det, arrayenc(x_det, 3))) = True Then
                    vp.TextAlign = taRightMiddle
                End If
                vp.TextBox FrmConsCompra.Fg1.TextMatrix(i_det, arrayenc(x_det, 3)), vposinileft, vposinitop, arrayenc(x_det, 2), 200
                vposinileft = vposinileft + arrayenc(x_det, 2) + vsepentrecol
                vp.TextAlign = taLeftMiddle
            Next
        End If
        If vposinitop >= 10900 Then
            vp.NewPage
            Encabezado True
            vindenc = True
            If vindenc = False Then
                vposinitop = consposinittop_enc + 500
            Else
                vposinitop = Consposinitop + 500
            End If
        Else
            vposinitop = vposinitop + vsepentrerow
        End If
    Next
End Sub

'*****************************************************************************************************
'* Nombre           : fDeterFila
'* Tipo             : FUNCION
'* Descripcion      : ***********************************************************, DEVUELVE UNA CADENA
'* Paranetros       : NOMBRE   |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pFila    |  LONG        |  ESPECIFICA EL ID DE LA FILA
'* Devuelve         : STRING
'*****************************************************************************************************
Private Function fDeterFila(pFila As Long) As String()
    Dim y_detfil As Long
    Dim arreglo(1 To 1, 1 To 2) As String, vcont As Integer, vOpc As Integer
    For y_detfil = 1 To vContCol_Usar
        If InStr(1, UCase(Trim(FrmConsCompra.Fg1.TextMatrix(pFila, arrayenc(y_detfil, 3)))), "TOTAL") > 0 Then
            arreglo(1, 1) = "TOT"
            arreglo(1, 2) = CStr(y_detfil)
            fDeterFila = arreglo
            Exit Function
        ElseIf InStr(1, UCase(Trim(FrmConsCompra.Fg1.TextMatrix(pFila, arrayenc(y_detfil, 3)))), "TOT. GEN.") > 0 Then
            arreglo(1, 1) = "TOTGEN"
            arreglo(1, 2) = CStr(y_detfil)
            fDeterFila = arreglo
            Exit Function
        Else
            If Trim(FrmConsCompra.Fg1.TextMatrix(pFila, arrayenc(y_detfil, 3))) <> "" Then
                vcont = vcont + 1
                vOpc = y_detfil
            End If
        End If
    Next
    If vcont >= 1 And vcont <= 4 Then
        arreglo(1, 1) = "ENC"
        arreglo(1, 2) = CStr(vOpc)
    Else
        arreglo(1, 1) = ""
        arreglo(1, 2) = ""
    End If
    fDeterFila = arreglo
End Function


'*****************************************************************************************************
'* Nombre           : CalLonLinea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CALCULA LA LONGITUD DE LA LINEA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CalLonLinea()
    Dim i_callonlin As Long
    For i_callonlin = 1 To vContCol_Usar
        If i_callonlin < vContCol_Usar Then
            vLonLinea = vLonLinea + arrayenc(i_callonlin, 2) + vsepentrecol
        Else
            vLonLinea = vLonLinea + arrayenc(i_callonlin, 2)
        End If
    Next
End Sub

'*****************************************************************************************************
'* Nombre           : LlenarArrayEnc
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLENA arrayenc CON DATOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub LlenarArrayEnc()
    Dim i_llearr As Long, x_llenarr As Long
    For i_llearr = 1 To FrmConsCompra.Fg1.Cols - 1
        If FrmConsCompra.Fg1.ColWidth(i_llearr) > 0 Then
            vContCol_Usar = vContCol_Usar + 1
            For x_llenarr = 1 To 3
                Select Case x_llenarr
                    Case 1
                        arrayenc(vContCol_Usar, 1) = FrmConsCompra.Fg1.TextMatrix(1, i_llearr)
                    Case 2
                        arrayenc(vContCol_Usar, 2) = FrmConsCompra.Fg1.ColWidth(i_llearr)
                    Case 3
                        arrayenc(vContCol_Usar, 3) = i_llearr
                End Select
            Next
        End If
    Next
End Sub

'*****************************************************************************************************
'* Nombre           : Encabezado
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GENERA EL ENCABEZADO
'* Paranetros       : NOMBRE    |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pEnc      |  BOOLEAN     |  ESPECIFICA SI SE HACE O NO EL ENCABEZADO
'* Devuelve         :
'*****************************************************************************************************
Private Sub Encabezado(pEnc As Boolean)
    Dim vPosLeft As Long, i_enc As Long, vpostop As Long
    Dim vleftitulo As Long
    vPosLeft = Consposinileft
    If pEnc = True Then
        vpostop = consposinittop_enc
        vp.FontSize = 7
        vp.TextBox "Empresa: " & NomEmp, Consposinileft, vpostop, 3000, 250
        
        vp.FontSize = 7
        vp.TextBox "R.U.C.:  " & NumRUC, Consposinileft, vpostop + 200, 2000, 250
        If vOrientPag = 1 Then
            vp.TextAlign = taRightMiddle
            vp.TextBox NomSIS, Consposinileft, vpostop, vLonLinea - Consposinileft, 250
            vp.TextBox "Fecha: " & Date & "", Consposinileft, vpostop + 200, vLonLinea - Consposinileft, 250
            vp.TextAlign = taLeftMiddle
        Else
            vp.TextAlign = taRightMiddle
            vp.TextBox NomSIS, Consposinileft, vpostop, vLonLinea - Consposinileft, 250
            vp.TextBox "Fecha: " & Date & "", Consposinileft, vpostop + 200, vLonLinea - Consposinileft, 250
            vp.TextAlign = taLeftMiddle
        End If
        
        vp.FontSize = 8
        vp.TextAlign = taCenterMiddle
        vp.TextBox m_Titulo1, Consposinileft, vpostop + 400, vLonLinea - Consposinileft, 300
        vp.TextAlign = taLeftMiddle
        
        vp.TextAlign = taCenterMiddle
        vp.TextBox m_titulo2, Consposinileft, vpostop + 600, vLonLinea - Consposinileft, 300
        vp.TextAlign = taLeftMiddle
        ' ESTOS ES CUANDO LE ASIGNO DESDE EL LOAD DEL FORM
        vpostop = Consposinitop
    Else
        ' ESTO ES CUANDO NO VA EL ENCABEZADO Y LE ASIGNO EL VALOR DE LA CONSTANTE
        vpostop = consposinittop_enc
    End If
    ' LINEA SUPERIOR DEL ENCABEZADO
    vp.DrawLine Consposinileft, vpostop - 100, vLonLinea, vpostop - 100
    ' LINEA INFERIOR DEL ENCABEZADO
    vp.DrawLine Consposinileft, vpostop + 300, vLonLinea, vpostop + 300
    vp.FontSize = 7
    vp.FontBold = True
    For i_enc = LBound(arrayenc, 1) To UBound(arrayenc, 1)
        If Trim(arrayenc(i_enc, 1)) <> "" Then
            vp.TextBox arrayenc(i_enc, 1), vPosLeft, vpostop, arrayenc(i_enc, 2), 300
            vPosLeft = vPosLeft + arrayenc(i_enc, 2) + vsepentrecol
        End If
    Next
    vp.FontBold = False
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIOS
    vp.Top = 0: vp.Left = 0
    inicializar_var_decgen
    
    ReDim arrayenc(1 To FrmConsCompra.Fg1.Cols - 1, 1 To 3) As String
    LlenarArrayEnc
    CalLonLinea
    Consposinitop = 2300
    With vp
        .PaperSize = pprA4
        .Zoom = 100
        If vLonLinea > 10900 Then 'HORIZONTAL
            .Orientation = orLandscape
            vOrientPag = 2
            Consposinileft = (cons_ancho_pagHor - vLonLinea) \ 2 + 500
            vLonLinea = vLonLinea + Consposinileft
        Else 'VERTICAL
            .Orientation = orPortrait
            vOrientPag = 1
            Consposinileft = (cons_ancho_pagVert - vLonLinea) \ 2
            vLonLinea = vLonLinea + Consposinileft
        End If
        
        .PageBorder = pbNone
        .StartDoc
            Encabezado True
            vindenc = True
            Detalle
        .EndDoc
        .ScrollIntoView 0, 0, 0, 0
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vp.Width = Me.Width
    vp.Height = Me.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_Titulo1 = "": m_titulo2 = ""
End Sub
