VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "VSPrint8.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "VSRpt8.ocx"
Begin VB.Form FrmRepLetras 
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VSPrinter8LibCtl.VSPrinter Vp 
      Height          =   5355
      Left            =   15
      TabIndex        =   0
      Top             =   30
      Width           =   7485
      _cx             =   13203
      _cy             =   9446
      Appearance      =   0
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
      Zoom            =   27.1593944790739
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
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VSReport8LibCtl.VSReport Vsr 
      Left            =   6975
      Top             =   -15
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
End
Attribute VB_Name = "FrmRepLetras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xSql As String
Public xIdLetra As Integer

Private Sub Form_Load()
    Dim RUTASY As String
    
On Error GoTo LaCague
    Me.Caption = "Tesorería - Emisión de Letras"
    Vp.Left = 0
    Vp.Top = 0
    
    Dim xRstEmp As New ADODB.Recordset
    Dim xRstLetra As New ADODB.Recordset
    
    Dim NumLet As String
    Dim xRutaRpt As String
    
    xSql = "SELECT mae_empresa.* FROM mae_empresa "

    RST_Busq xRstEmp, xSql, xCon
    
    xSql = "SELECT * FROM let_letra WHERE id = " & xIdLetra & ""
    RST_Busq xRstLetra, xSql, xCon
    
   ' load report from XML file
    
    '--indica la ruta del reporte
    RUTASY = LeerLineaINI(Trim(App.Path) + "\siac.ini", "RUTASY", "RUTAS")          ' LEEMOS LA RUTA DEL SISTEA
    xRutaRpt = RUTASY & "reportes\letra.xml"
    
    Vsr.Load xRutaRpt, "letras"
    
    xSql = "SELECT let_letradet.idlet, [let_letradet]![numser] & '-' & [let_letradet]![numdoc] AS numlet, '" & NulosC(xRstEmp("luggir")) & "' AS luggir, " _
        & " let_letradet.fchemi, let_letradet.fchven, mae_moneda.simbolo AS simmon, let_letradet.implet, '" & NulosC(xRstEmp("fiador")) & "' AS cobrador, let_letradet.imptext, " _
        & " [mae_bancos]![descripcion] AS banco, '" & NulosC(xRstLetra("girado")) & "' AS nomcli, '" & NulosC(xRstLetra("numdocgir")) & "' AS clidni, '" & NulosC(xRstLetra("numtel")) & "' AS clitel, " _
        & " [mae_banconumcta]![numcue] AS numcueban, '" & NulosC(xRstLetra("dir")) & "' AS dircli " _
        & " FROM ((let_letra LEFT JOIN mae_moneda ON let_letra.idmon = mae_moneda.id) RIGHT JOIN let_letradet ON let_letra.id = let_letradet.idlet) LEFT JOIN (mae_banconumcta " _
        & " LEFT JOIN mae_bancos ON mae_banconumcta.idban = mae_bancos.id) ON let_letradet.idbcocta = mae_banconumcta.id " _
        & " WHERE (((let_letradet.idlet)=" & xIdLetra & "))"
    
    Vsr.DataSource.ConnectionString = xCon.ConnectionString
    Vsr.DataSource.RecordSource = xSql
    Cargar
    Exit Sub
LaCague:
    MsgBox Err.Description & vbCr & Err.Source, vbInformation, xTitulo
    Err.Clear
End Sub

Sub Cargar()
   
    ' no reentrancy
    If Vsr.IsBusy Then Exit Sub
    
    ' prepare to benchmark
    Dim t
    t = Timer
    MousePointer = 11
    
    ' render report
    On Error Resume Next
    Vsr.Render Vp
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description
    End If
    
    ' done, tell user how long it took
    MousePointer = 0
    Vp.NavBarText = "Done in " & Format(Timer - t, "#.##") & " seconds"
    
End Sub

Private Sub Form_Resize()
    Vp.Width = Me.Width - 110
    Vp.Height = Me.Height - 400
End Sub
