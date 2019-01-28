VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form FrmPrinter 
   Caption         =   "Contabilidad - Reporte Libro Diario"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7515
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VSPrinter7LibCtl.VSPrinter Vp 
      Height          =   7515
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _cx             =   20955
      _cy             =   13256
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
      MarginRight     =   1080
      MarginBottom    =   1080
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
      Zoom            =   39.8040961709706
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
Attribute VB_Name = "FrmPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean
Dim xNumPagina As Integer
Dim AnchoLinea As Integer
Dim AltoHoja As Integer

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Cargar
    End If
End Sub

Private Sub Form_Load()
    Vp.PaperSize = pprA4
    If Prin_OrientacionHoja = 2 Then
        '0 = horinzontal
        Vp.Orientation = orLandscape
        AnchoLinea = 16000
        AltoHoja = 11000
    Else
        '1 = Vertical
        Vp.Orientation = orPortrait
        AnchoLinea = 11000
        AltoHoja = 16000
    End If
    SeEjecuto = False
    xNumPagina = 0
End Sub

Sub Cargar()
    Dim A, B, xPosX2 As Integer
    Dim xFila As Double
    Dim xAlineacion As Integer
    Dim xTotaliza() As String
    Dim xNunColTotaliza As Integer
    
    'averiguamos los campos que se totalizaran y los escribimos en el array
    RstPrin.Filter = "totalizar = -1"
    ReDim xTotaliza(2, RstPrin.RecordCount)
    
    If RstPrin.RecordCount <> 0 Then
        RstPrin.MoveFirst
        xNunColTotaliza = RstPrin.RecordCount
        For A = 0 To RstPrin.RecordCount
            xTotaliza(0, A) = RstPrin("abrev")
            RstPrin.MoveNext
            If RstPrin.EOF = True Then Exit For
        Next A
    End If
    xNumPagina = 1
    
    RstPrin.Filter = adFilterNone
    With Vp
        ' set up
        .FontName = "Courier New"
        .FontSize = 10
        
        .TextColor = &H80000008 'RGB(200, 200, 200)
        'MUESTRA BORDES DEL REPORTE
        .PageBorder = pbNone

        .StartDoc
            Cabecera
            .TextAlign = taLeftMiddle
            xFila = 2500
            .FontSize = xTamañoFuente

            For B = 1 To UBound(ArrayPrin)
                xPosX2 = 900
                RstPrin.MoveFirst
                For A = 1 To RstPrin.RecordCount
                    
                    xAlineacion = CambiarAlineacionRporte(RstPrin("alineacion"))
                    If A = 1 Then
                        If RstPrin("imprimir") = -1 Then
                            Vp.TextAlign = xAlineacion
                            If Prin_TextoConsiderarAncho <> 0 Or F_NulosC(Prin_TextoConsiderar) <> "" Then
                                If UCase(Mid(ArrayPrin(B, A), 1, Prin_TextoConsiderarAncho)) = UCase(F_NulosC(Prin_TextoConsiderar)) Then
                                    Vp.TextBox ArrayPrin(B, A), xPosX2, xFila, 5000, 140
                                Else
                                    Vp.TextBox ArrayPrin(B, A), xPosX2, xFila, RstPrin("anchoprin"), 140
                                End If
                            Else
                                Vp.TextBox ArrayPrin(B, A), xPosX2, xFila, RstPrin("anchoprin"), 140
                            End If
                            xPosX2 = ((xPosX2 + RstPrin("anchoprin")) + 100)
                        End If
                    Else
                        If RstPrin("imprimir") = -1 Then
                            Vp.TextAlign = xAlineacion
                            If Prin_TextoConsiderarAncho <> 0 Or F_NulosC(Prin_TextoConsiderar) <> "" Then
                                If UCase(Mid(ArrayPrin(B, A), 1, Prin_TextoConsiderarAncho)) <> UCase(F_NulosC(Prin_TextoConsiderar)) Then
                                    Vp.TextBox ArrayPrin(B, A), xPosX2, xFila, RstPrin("anchoprin"), 140
                                End If
                            Else
                                Vp.TextBox ArrayPrin(B, A), xPosX2, xFila, RstPrin("anchoprin"), 140
                            End If
                            
                            xPosX2 = ((xPosX2 + RstPrin("anchoprin")) + 100)
                        End If
                    End If
                    
                    If RstPrin("totalizar") = -1 Then
                        Dim X As Integer
                        For X = 0 To xNunColTotaliza - 1
                            If RstPrin("abrev") = xTotaliza(0, X) Then
                                If F_NulosC(ArrayPrin(B, 1)) <> "" Then
                                    xTotaliza(1, X) = F_NulosN(xTotaliza(1, X)) + F_NulosN(ArrayPrin(B, A))
                                End If
                            End If
                        Next X
                    End If
                    
                    RstPrin.MoveNext
                    If RstPrin.EOF = True Then Exit For
                Next A
                
                If xFila >= (AltoHoja - 200) Then
                    Dim NewAlto As Integer
                    NewAlto = AltoHoja
                    xNumPagina = xNumPagina + 1
                    
                    NewAlto = NewAlto + 100
                    Vp.DrawLine 900, NewAlto, AnchoLinea, NewAlto
                    
                    'imprimimos los totales
                    Dim x1, y2  As Integer
                    Dim xPosX3 As Integer
                    
                    'IMPRIMIMOS EL VAN
                    NewAlto = NewAlto + 50
                    Vp.FontBold = True
                    For x1 = 0 To xNunColTotaliza - 1
                        RstPrin.MoveFirst
                        xPosX3 = 900
                       
                        For y2 = 1 To RstPrin.RecordCount
                            If RstPrin("imprimir") = -1 Then
                                If xTotaliza(0, x1) = RstPrin("abrev") Then
                                    If x1 = 0 Then
                                        Vp.TextBox "VAN  ", xPosX3 - 1000, NewAlto, RstPrin("anchoprin"), 140
                                    End If
                                    xAlineacion = CambiarAlineacionRporte(RstPrin("alineacion"))
                                    Vp.TextAlign = xAlineacion
                                    Vp.TextBox Format(xTotaliza(1, x1), "#,###.00"), xPosX3, NewAlto, RstPrin("anchoprin"), 140
                                End If
                                
                                xPosX3 = ((xPosX3 + F_NulosN(RstPrin("anchoprin"))) + 100)
                            End If
                            RstPrin.MoveNext
                            If RstPrin.EOF = True Then Exit For
                        Next y2
                    Next x1
                    
                    .NewPage
                    Vp.TextAlign = taLeftBottom
                    Vp.FontBold = False
                    Cabecera
                    
                    'IMPRIMIMOS EL VIENEN
                    Vp.FontBold = True
                    For x1 = 0 To xNunColTotaliza - 1
                        RstPrin.MoveFirst
                        xPosX3 = 900
                        For y2 = 1 To RstPrin.RecordCount
                            If RstPrin("imprimir") = -1 Then
                                If xTotaliza(0, x1) = RstPrin("abrev") Then
                                    If x1 = 0 Then
                                        Vp.TextBox "VIENEN  ", xPosX3 - 1000, 2500, RstPrin("anchoprin"), 140
                                    End If
                                    xAlineacion = CambiarAlineacionRporte(RstPrin("alineacion"))
                                    Vp.TextAlign = xAlineacion
                                    Vp.TextBox Format(xTotaliza(1, x1), "#,###.00"), xPosX3, 2500, RstPrin("anchoprin"), 140
                                End If
                                
                                xPosX3 = ((xPosX3 + F_NulosN(RstPrin("anchoprin"))) + 100)
                            End If
                            RstPrin.MoveNext
                            If RstPrin.EOF = True Then Exit For
                        Next y2
                    Next x1
                    Vp.FontBold = False
                    xFila = 2640
                Else
                    xFila = xFila + 140
                End If
             Next B
        .EndDoc
        .ScrollIntoView 0, 0
    End With
End Sub

Sub Cabecera()
    Dim xMes, xMoneda As String
    Dim A As Integer
    
    Vp.FontSize = Prin_TamañoCabecera
    Vp.FontName = Prin_FuenteCabecera
    Vp.CurrentX = 900: Vp.CurrentY = 700: Vp.Paragraph = Prin_Cabecera1
    
    'Vp.CurrentX = 8900: Vp.CurrentY = 700: Vp.Paragraph = "FECHA : " + Format(Prin_Fecha, "dd/mm/yy")
    Vp.CurrentX = AnchoLinea - 2000: Vp.CurrentY = 700: Vp.Paragraph = "FECHA : " + Format(Prin_Fecha, "dd/mm/yy")

    Vp.CurrentX = 900: Vp.CurrentY = 950: Vp.Paragraph = Prin_Cabecera2
    'Vp.CurrentX = 8900: Vp.CurrentY = 950: Vp.Paragraph = "Nº PAGINA : " + Format(xNumPagina, "0000")
    Vp.CurrentX = AnchoLinea - 2000: Vp.CurrentY = 950: Vp.Paragraph = "Nº PAGINA : " + Format(xNumPagina, "0000")
    
    Vp.TextAlign = taCenterMiddle
    Vp.CurrentX = 3600: Vp.CurrentY = 1150:  Vp.Paragraph = Prin_Titulo1
    Vp.CurrentX = 3600: Vp.CurrentY = 1400:  Vp.Paragraph = Prin_Titulo2
    Vp.TextAlign = taLeftBottom
    
    Vp.FontSize = xTamañoFuente
    Vp.FontName = "Courier New"
    
    Vp.DrawLine 900, 1900, AnchoLinea, 1900
    Dim xPosX As Integer
    xPosX = 900
    RstPrin.MoveFirst
    
    Vp.TextAlign = taCenterMiddle
    For A = 1 To RstPrin.RecordCount
        If RstPrin("imprimir") = -1 Then
            Vp.BrushColor = &HC0&
            Vp.TextBox RstPrin("abrev"), xPosX, 2000, RstPrin("anchoprin"), 300
            xPosX = ((xPosX + RstPrin("anchoprin")) + 100)
        End If
        RstPrin.MoveNext
        If RstPrin.EOF = True Then Exit For
    Next A
        
    Vp.DrawLine 900, 2400, AnchoLinea, 2400
End Sub

Function CambiarAlineacionRporte(mAlineacion As Integer) As Integer
    Select Case mAlineacion
        Case 1:   CambiarAlineacionRporte = taCenterMiddle '--CENTRADO
        Case 4:   CambiarAlineacionRporte = taLeftMiddle '--IZQUIERDA
        Case 7:   CambiarAlineacionRporte = taRightMiddle '--DERECHO
        Case Else
            CambiarAlineacionRporte = 7
    End Select

End Function

Private Sub Form_Resize()
    Vp.Height = Me.Height - 200
    Vp.Width = Me.Width - 200
End Sub

