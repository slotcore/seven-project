VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   510
      Left            =   10215
      TabIndex        =   3
      Top             =   2805
      Width           =   1530
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Left            =   10440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2055
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   10305
      TabIndex        =   1
      Top             =   780
      Width           =   1530
   End
   Begin VSPrinter7LibCtl.VSPrinter Vp 
      Height          =   7590
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   10050
      _cx             =   17727
      _cy             =   13388
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
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      Zoom            =   75
      ZoomMode        =   0
      ZoomMax         =   400
      ZoomMin         =   25
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   255
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
      Navigation      =   1
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rst As New ADODB.Recordset

Private Sub Command1_Click()
    
    
    RST_Busq Rst, "SELECT alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.preuni " _
        & " From alm_inventario ORDER BY alm_inventario.descripcion", xCon
    
    With Vp
        ' set up
        .FontName = "Tahoma"
        .FontSize = 9
        .PenColor = RGB(190, 190, 190)
        'MUESTRA BORDES DEL REPORTE
        .PageBorder = pbColTopBottom
        .Header = "LUIS RINCON LA TORRE:" & vbLf & "RUC Nº : 1002003335" & vbLf & "  " & vbLf & " " & "| LISTA DE PRODUCTOS |Fecha :" & Format(Date, "dd/mm/yyyy") & vbLf & "  " & vbLf & "  " & vbLf & "  "
        .Footer = "Ejemplo|Haber si funca|Page %d"
        
        .StartDoc
            ' show title
            .Paragraph = "Lista de Productos"
            .FontBold = True
            .Paragraph = "titulo 1"
            .FontBold = False
            .Paragraph = ""
            '.PageBorder = pbColBottom
            
            ' render recordset (this is the main routine here)
            Crear
        .EndDoc
        .ScrollIntoView 0, 0
    End With
End Sub

Sub Crear()
    Dim TamañoCelda As String
    Dim CabeceraCelda As String
    Dim CamposCelda
    
    Vp.StartTable
        CamposCelda = Rst.GetRows(Rst.RecordCount)
        TamañoCelda = "2000|5000|1000"
        CabeceraCelda = "Codigo|Producto|Precio"
        Vp.AddTableArray TamañoCelda, CabeceraCelda, CamposCelda
        
        'cabecera de la tabla
        Vp.TableCell(tcFontBold, 0) = True
        Vp.TableCell(tcBackColor, 0) = vbYellow
        Vp.TableCell(tcRowHeight, 0) = Vp.TextHeight("Test") * 1.1
        Vp.TableCell(tcAlign, 0) = taLeftMiddle
        
        'alineandoa la derecha las columnas numericas
        Vp.TableCell(tcAlign, 1, 3, Rst.RecordCount, 3) = taRightMiddle
    
    Vp.EndTable
End Sub

Sub RenderRecordset(Vp As VSPrinter, rs As Recordset, ByVal maxh As Double)
'
' renders a recordset as a VSPrinter table
'
' parameters:
'
' vp:   Reference to a VSPrinter control. Make sure you already
'       called the StartDoc method before calling this routine.
'
' rs:   The Recordset containing the data to render.
'       This may be DAO or ADO, just change the declaration.
'
' maxh: Maximum allowable row height (important mostly when
'       rendering memo fields, which may be very long).
'       Set to a value <= 0 to accept any height.
'
' The routine renders the table using the AddTableArray method.
' It also uses the TableCell property to format the table.
' The routine ensures that the table will fit on the page width.
' It sizes and aligns columns based on the data.
'
    Dim arr, i%, j%, wid!
    
    ' read recordset into an array
    rs.MoveLast
    rs.MoveFirst
    i = rs.RecordCount
    If i = 0 Then Exit Sub
    arr = rs.GetRows(i)
    
    ' create table header and dummy format
    Dim fmt$, hdr$
    For i = 0 To rs.Fields.Count - 1
        If i > 0 Then hdr = hdr & "|"
        fmt = fmt & "|"
        hdr = hdr & rs.Fields(i).Name
        fmt = fmt & 500
    Next
    
    ' create table
    Vp.StartTable
    Vp.AddTableArray fmt, hdr, arr
    
    ' format table
    For i = 0 To rs.Fields.Count - 1
    
        ' right-align numbers and dates
        Select Case rs.Fields(i).Type
            Case dbBigInt, dbByte, dbChar, dbCurrency, dbDecimal, _
                 dbDouble, dbFloat, dbInteger, dbLong, dbNumeric, dbSingle, _
                 dbDate
                Vp.TableCell(tcColAlign, , i + 1) = taRightTop
        End Select
        
        ' set column width
        If rs.Fields(i).Type = dbMemo Then
            Vp.TableCell(tcColWidth, , i + 1) = "2.5in"
        Else
            fmt = ""
            For j = 0 To UBound(arr, 2)
                If j > 100 Then Exit For
                If Len(fmt) < Len(arr(i, j)) Then
                    fmt = arr(i, j)
                End If
            Next
            If Len(rs.Fields(i).Name) > Len(fmt) Then fmt = rs.Fields(i).Name
            Vp.TableCell(tcColWidth, , i + 1) = Vp.TextWidth(fmt) * 1.4
        End If
    Next
    
    ' format header row (0)
    Vp.TableCell(tcFontBold, 0) = True
    Vp.TableCell(tcBackColor, 0) = vbYellow
    Vp.TableCell(tcRowHeight, 0) = Vp.TextHeight("Test") * 1.1
    Vp.TableCell(tcAlign, 0) = taLeftMiddle
    
    ' make sure it all fits
    For i = 1 To Vp.TableCell(tcCols)
        wid = wid + Vp.TableCell(tcColWidth, , i)
    Next
    Vp.GetMargins
    If wid > Vp.X2 - Vp.X1 Then
        wid = (Vp.X2 - Vp.X1) / wid * 0.95
        For i = 1 To Vp.TableCell(tcCols)
            Vp.TableCell(tcColWidth, , i) = wid * Vp.TableCell(tcColWidth, , i)
        Next
    End If
    
    ' honor maximum row height
    If maxh > 0 Then
        For i = 1 To Vp.TableCell(tcRows)
            If Vp.TableCell(tcRowHeight, i) > maxh Then
                Vp.TableCell(tcRowHeight, i) = maxh
            End If
        Next
    End If
    
    ' done with table
    Vp.EndTable

End Sub


Private Sub Command2_Click()
    
    
    RST_Busq Rst, "SELECT alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.preuni " _
        & " From alm_inventario ORDER BY alm_inventario.descripcion", xCon
    
    With Vp
        ' set up
        .FontName = "Courier New"
        .FontSize = 10
        .ColorMode = cmColor
        '.TextColor = RGB(150, 190, 100)
        .PenColor = RGB(190, 190, 190)
        
        'MUESTRA BORDES DEL REPORTE
        '.PageBorder = pbNone
        .PageBorder = pbTopBottom
        
        .StartDoc
            Cabecera
            
            Dim a, xFila As Integer
            xFila = 1900
            Rst.MoveFirst
            .FontSize = 8
            For a = 1 To Rst.RecordCount
            
                .TextAlign = taLeftTop
                .CurrentX = 1350: .CurrentY = xFila: .Paragraph = Rst("codpro")
                .CurrentX = 3300: .CurrentY = xFila: .Paragraph = Rst("descripcion")
                .TextAlign = taRightTop
                '.TextColor = RGB(150, 190, 100)
                .TextBox Format(Rst("preuni"), "0.00"), 9000, xFila, 1000, "250"
                '.CurrentX = 9000: .CurrentY = xFila: .Paragraph = Format(Rst("preuni"), "0.00")
                
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
                If xFila >= 15000 Then
                    .NewPage
                    .TextAlign = taLeftTop
                    Cabecera
                    .FontSize = 8
                    xFila = 1900
                Else
                    xFila = xFila + 250
                End If
                
                
            Next a
            
'            ' show title
'            .Paragraph = "Lista de Productos"
'            .FontBold = True
'            .Paragraph = "titulo 1"
'            .FontBold = False
'            .Paragraph = ""
'            '.PageBorder = pbColBottom
'
'            ' render recordset (this is the main routine here)
'            Crear
        .EndDoc
        .ScrollIntoView 0, 0
    End With

End Sub

Sub Cabecera()
    Vp.FontSize = 9
    Vp.CurrentX = 1350: Vp.CurrentY = 350: Vp.Paragraph = "LUIS RINCON LA TORRE"
    Vp.CurrentX = 8300: Vp.CurrentY = 350: Vp.Paragraph = "FECHA : " + Format(Date, "dd/mm/yyyy")
    
    Vp.CurrentX = 1350: Vp.CurrentY = 600: Vp.Paragraph = "R.U.C. Nº : 20100523564"

    Vp.CurrentX = 5500: Vp.CurrentY = 850:  Vp.Paragraph = "LISTA DE ITEMS"  ': '.Align = 3
    Vp.CurrentX = 5500: Vp.CurrentY = 1050: Vp.Paragraph = "==============":
    
    Vp.DrawLine 1350, 1400, 10550, 1400
    Vp.CurrentX = 1350: Vp.CurrentY = 1510: Vp.Paragraph = "CODIGO"
    Vp.CurrentX = 3300: Vp.CurrentY = 1510: Vp.Paragraph = "DESCRIPCION"
    Vp.CurrentX = 9000: Vp.CurrentY = 1510: Vp.Paragraph = "P. UNI."
    Vp.DrawLine 1350, 1850, 10550, 1850
End Sub

