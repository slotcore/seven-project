VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form FrmPrinCtaCtaCli 
   Caption         =   "Caja y Bancos - Impresión Cuenta Corriente"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "FrmPrinCtaCtaCli.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
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
Attribute VB_Name = "FrmPrinCtaCtaCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Cargar
    End If
End Sub

Private Sub Form_Load()
    Vp.PaperSize = pprA4
    SeEjecuto = False
End Sub

Sub Cargar()
    With Vp
        .FontName = "Courier New"
        .FontSize = 10
        .TextColor = &H80000008 'RGB(200, 200, 200)
        
        'MUESTRA BORDES DEL REPORTE
        .PageBorder = pbNone
        
        .StartDoc
            Cabecera
            
            Dim A, xFila As Integer
            xFila = 2300
            
            .FontSize = 6
            For A = FrmCtaCte2.Fg1.FixedRows To FrmCtaCte2.Fg1.Rows - 1

                .TextAlign = taLeftTop
                .CurrentX = 900:  .CurrentY = xFila: .Paragraph = FrmCtaCte2.Fg1.TextMatrix(A, 1)
                .CurrentX = 1800: .CurrentY = xFila: .Paragraph = FrmCtaCte2.Fg1.TextMatrix(A, 2)
                .CurrentX = 3000: .CurrentY = xFila: .Paragraph = FrmCtaCte2.Fg1.TextMatrix(A, 3)
                .CurrentX = 3400: .CurrentY = xFila: .Paragraph = FrmCtaCte2.Fg1.TextMatrix(A, 4)
                
                .CurrentX = 5000: .CurrentY = xFila: .Paragraph = Format(FrmCtaCte2.Fg1.TextMatrix(A, 5), "dd/mm/yy")
                .CurrentX = 5800: .CurrentY = xFila: .Paragraph = Format(FrmCtaCte2.Fg1.TextMatrix(A, 6), "dd/mm/yy")
                .CurrentX = 6700: .CurrentY = xFila: .Paragraph = FrmCtaCte2.Fg1.TextMatrix(A, 7)
                
                .TextAlign = taRightTop
                .TextBox FrmCtaCte2.Fg1.TextMatrix(A, 9), 7300, xFila, 400, "250"
                .TextBox FrmCtaCte2.Fg1.TextMatrix(A, 10), 7800, xFila, 900, "250"
                .TextBox FrmCtaCte2.Fg1.TextMatrix(A, 11), 8800, xFila, 900, "250"
                .TextBox FrmCtaCte2.Fg1.TextMatrix(A, 12), 9800, xFila, 900, "250"


                If xFila >= 15500 Then
                    Vp.DrawLine 900, 15700, 11000, 15700
                    .NewPage
                    .TextAlign = taLeftTop
                    Cabecera
                    .FontSize = 6
                    xFila = 2300
                Else
                    xFila = xFila + 200
                End If
            Next A
            Vp.DrawLine 900, 15700, 11000, 15700
        .EndDoc
        .ScrollIntoView 0, 0
    End With
End Sub

Sub Cabecera()
    
    Vp.FontSize = 10
    Vp.CurrentX = 900: Vp.CurrentY = 700: Vp.Paragraph = NomEmp
    Vp.CurrentX = 8900: Vp.CurrentY = 700: Vp.Paragraph = "FECHA : " + Format(Date, "dd/mm/yy")
    
    Vp.CurrentX = 900: Vp.CurrentY = 950: Vp.Paragraph = "R.U.C. Nº : " + NumRuc

    Vp.CurrentX = 5000: Vp.CurrentY = 1100:  Vp.Paragraph = "CUENTA CORRIENTE"
    Vp.FontSize = 8
    
    If FrmCtaCte2.OptSel1.Value = True Then
        If FrmCtaCte2.OptCliente.Value = True Then
            Vp.CurrentX = 900: Vp.CurrentY = 1500:  Vp.Paragraph = "CLIENTE : " + "TODOS LOS CLIENTES"
        End If
        If FrmCtaCte2.OptProvee.Value = True Then
            Vp.CurrentX = 900: Vp.CurrentY = 1500:  Vp.Paragraph = "PROVEEDOR : " + "TODOS LOS PROVEEDORES"
        End If
    End If
    If FrmCtaCte2.OptSel2.Value = True Then
        If FrmCtaCte2.OptCliente.Value = True Then
            Vp.CurrentX = 900: Vp.CurrentY = 1500:  Vp.Paragraph = "CLIENTE : " + UCase(FrmCtaCte2.TxtCliPro.Text)
        End If
        If FrmCtaCte2.OptProvee.Value = True Then
            Vp.CurrentX = 900: Vp.CurrentY = 1500:  Vp.Paragraph = "PROVEEDOR : " + UCase(FrmCtaCte2.TxtCliPro.Text)
        End If
    End If
    
    Vp.CurrentX = 7000: Vp.CurrentY = 1500:  Vp.Paragraph = "HASTA EL DIA DE : " + Format(FrmCtaCte2.TxtFecha.Valor, "DD/MM/YY")
    
    Vp.FontSize = 6
    Vp.DrawLine 900, 1800, 11000, 1800
    Vp.CurrentX = 900:    Vp.CurrentY = 1900:  Vp.Paragraph = "Nº Registro"
    Vp.CurrentX = 1800:   Vp.CurrentY = 1900:  Vp.Paragraph = "Origen"
    Vp.CurrentX = 3000:   Vp.CurrentY = 1900:  Vp.Paragraph = "T.D."
    Vp.CurrentX = 3400:   Vp.CurrentY = 1900:  Vp.Paragraph = "Nº Documento"
    Vp.CurrentX = 5000:   Vp.CurrentY = 1900:  Vp.Paragraph = "Fch. Emi."
    Vp.CurrentX = 5800:   Vp.CurrentY = 1900:  Vp.Paragraph = "Fch. Ven."
    Vp.CurrentX = 6800:   Vp.CurrentY = 1900:  Vp.Paragraph = "M"
    Vp.CurrentX = 7300:   Vp.CurrentY = 1900:  Vp.Paragraph = "T.C."
    Vp.CurrentX = 8000:   Vp.CurrentY = 1900:  Vp.Paragraph = "--CARGO --"
    Vp.CurrentX = 9000:  Vp.CurrentY = 1900:  Vp.Paragraph = "--ABONO --"
    Vp.CurrentX = 10000:  Vp.CurrentY = 1900:  Vp.Paragraph = "--SALDO --"
    Vp.DrawLine 900, 2200, 11000, 2200
End Sub

