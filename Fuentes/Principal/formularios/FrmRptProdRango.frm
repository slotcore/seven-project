VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmRptProdRango 
   Caption         =   "Produccion - Reporte de Produccion x Periodo"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
      Height          =   300
      Left            =   930
      TabIndex        =   0
      Top             =   525
      Width           =   1260
      _ExtentX        =   2223
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
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar"
      Height          =   360
      Left            =   10065
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1665
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6285
      Left            =   0
      TabIndex        =   3
      Top             =   870
      Width           =   11730
      _cx             =   20690
      _cy             =   11086
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   128
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmRptProdRango.frx":0000
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "2"
         Height          =   885
         Left            =   300
         TabIndex        =   7
         Top             =   510
         Visible         =   0   'False
         Width           =   5250
         Begin MSComctlLib.ProgressBar Pgb1 
            Height          =   285
            Left            =   105
            TabIndex        =   8
            Top             =   480
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label LblDia 
            AutoSize        =   -1  'True
            Caption         =   "Procesando Dia  : "
            Height          =   195
            Left            =   135
            TabIndex        =   9
            Top             =   195
            Width           =   1320
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000005&
            BorderWidth     =   2
            Index           =   1
            X1              =   15
            X2              =   15
            Y1              =   0
            Y2              =   855
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            Index           =   0
            X1              =   5235
            X2              =   5235
            Y1              =   30
            Y2              =   885
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            BorderWidth     =   2
            X1              =   0
            X2              =   5220
            Y1              =   15
            Y2              =   15
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            X1              =   30
            X2              =   5250
            Y1              =   870
            Y2              =   870
         End
      End
   End
   Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
      Height          =   300
      Left            =   3630
      TabIndex        =   1
      Top             =   525
      Width           =   1260
      _ExtentX        =   2223
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4950
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptProdRango.frx":00D6
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptProdRango.frx":061A
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptProdRango.frx":079E
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptProdRango.frx":0BF2
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptProdRango.frx":0D0A
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptProdRango.frx":124E
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptProdRango.frx":1792
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptProdRango.frx":18A6
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptProdRango.frx":19BA
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptProdRango.frx":1E0E
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptProdRango.frx":1F7A
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fch. Final"
      Height          =   195
      Left            =   2805
      TabIndex        =   5
      Top             =   570
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fch. Inicio"
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   570
      Width           =   735
   End
End
Attribute VB_Name = "FrmRptProdRango"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xNumPag As Integer

Private Sub Command1_Click()

End Sub

Private Sub CmdBuscar_Click()
    If TxtFchIni.Valor = "" Then
        MsgBox "No ha especificado la fecha de inicio del periodo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    If TxtFchFin.Valor = "" Then
        MsgBox "No ha especificado la fecha de final del periodo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If CDate(NulosC(TxtFchIni.Valor)) > CDate(NulosC(TxtFchFin.Valor)) Then
        MsgBox "La fecha de inicio no puede ser menor o igual a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    MuestraProduccion
End Sub

Private Sub Form_Load()
    Fg1.Rows = 1
    Fg1.Cols = 2
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
End Sub

Sub MuestraProduccion()
    Frame1.Left = 3285
    Frame1.Top = 2235
    Frame1.Visible = True
    
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    Fg1.Rows = 1
    Fg1.Cols = 2

    RST_Busq Rst, "SELECT MAE_Producto.Cod_Item, MAE_Producto.Descripcion, (SELECT Sum([Cantidad]) AS canpro" _
        & " From PRD_Parte_Produccion WHERE (((PRD_Parte_Produccion.Cod_Item)=mae_producto.cod_item) " _
        & " AND ((PRD_Parte_Produccion.Fecha)>=CDate('" & TxtFchIni.Valor & "') And (PRD_Parte_Produccion.Fecha)<=CDate('" & TxtFchFin.Valor & "')))) AS canpro " _
        & " From MAE_Producto WHERE ((((SELECT Sum([Cantidad]) AS canpro From PRD_Parte_Produccion " _
        & " WHERE (((PRD_Parte_Produccion.Cod_Item)=mae_producto.cod_item) AND " _
        & " ((PRD_Parte_Produccion.Fecha)>=CDate('" & TxtFchIni.Valor & "') And (PRD_Parte_Produccion.Fecha)<=CDate('" & TxtFchFin.Valor & "')))))<>0))" _
        & " ORDER BY MAE_Producto.Descripcion", xCon

    If Rst.EOF = False Then
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 0) = Rst("cod_item")
            Fg1.TextMatrix(A, 1) = Rst("descripcion")
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    Dim xFecIni, xFecFin As Variant
    Dim xFchAct As Date
    Dim xNumDias, B As Integer
    xFecIni = DateValue(CDate(TxtFchIni.Valor))
    xFecFin = DateValue(CDate(TxtFchFin.Valor))
    xNumDias = (xFecFin - xFecIni) + 1
    
    xFchAct = CDate(TxtFchIni.Valor)
    Pgb1.Max = xNumDias
    Frame1.Refresh
    
    For A = 1 To xNumDias
        LblDia.Caption = "Procesando Dia  : " + Format(xFchAct, "dd/mm/yy")
        LblDia.Refresh
        Pgb1.Value = A
        
        'RST_Busq Rst, "SELECT MAE_Producto.Cod_Item, MAE_Producto.Descripcion, (SELECT Sum(PRD_Parte_Produccion.Cantidad) AS canpro " _
            & " From PRD_Parte_Produccion WHERE ((PRD_Parte_Produccion.Cod_Item=mae_producto.cod_item) AND " _
            & " (PRD_Parte_Produccion.Fecha=CDate('" & xFchAct & "')))) AS canpro FROM MAE_Producto", xCon
        
        RST_Busq Rst, "SELECT MAE_Producto.Cod_Item, MAE_Producto.Descripcion, " _
            & " (SELECT Sum([Cantidad]) AS canpro From PRD_Parte_Produccion WHERE (((PRD_Parte_Produccion.Cod_Item)=mae_producto.cod_item) " _
            & " AND ((PRD_Parte_Produccion.Fecha)>=CDate('" & TxtFchIni.Valor & "') And (PRD_Parte_Produccion.Fecha) <= CDate('" & TxtFchFin.Valor & "')))) AS canproper, " _
            & " (SELECT Sum([Cantidad]) AS canpro From PRD_Parte_Produccion WHERE (((PRD_Parte_Produccion.Cod_Item)=mae_producto.cod_item) " _
            & " AND ((PRD_Parte_Produccion.Fecha)=CDate('" & xFchAct & "')))) AS canprodia " _
            & " From MAE_Producto WHERE ((((SELECT Sum([Cantidad]) AS canpro " _
            & " From PRD_Parte_Produccion WHERE (((PRD_Parte_Produccion.Cod_Item)=mae_producto.cod_item) " _
            & " AND ((PRD_Parte_Produccion.Fecha)>=CDate('" & TxtFchIni.Valor & "') And (PRD_Parte_Produccion.Fecha)<=CDate('" & TxtFchFin.Valor & "')))))<>0)) " _
            & " ORDER BY MAE_Producto.Descripcion", xCon
        
        
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            Fg1.Cols = Fg1.Cols + 1
            Fg1.ColWidth(Fg1.Cols - 1) = 800
            Fg1.TextMatrix(0, Fg1.Cols - 1) = Format(xFchAct, "dd/mm/yy")
            
            For B = 1 To Rst.RecordCount
                Fg1.TextMatrix(B, Fg1.Cols - 1) = Format(Rst("canprodia"), "0.00")
                Fg1.TextMatrix(B, 0) = Val(Fg1.TextMatrix(B, 0)) + NulosN(Rst("canprodia"))
                
                Rst.MoveNext
                If Rst.EOF = True Then
                    Exit For
                End If
            Next B
        End If
        
        xFchAct = (xFchAct + 1)
    Next A
    
    Fg1.Cols = Fg1.Cols + 1
    Fg1.TextMatrix(0, Fg1.Cols - 1) = "TOTAL"
    For A = 1 To Fg1.Rows - 1
        Fg1.TextMatrix(A, Fg1.Cols - 1) = Format(Fg1.TextMatrix(A, 0), "0.00")
    Next A
    
    For A = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(A, Fg1.Cols - 1)) = 0 Then
            Fg1.RemoveItem (A)
            If A = Fg1.Rows Then Exit For
            A = A - 1
        End If
    Next A
    Frame1.Visible = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 12 Then
        Imprimir
    End If
    If Button.Index = 14 Then
        Unload Me
    End If
End Sub

Sub Imprimir()
    If Fg1.Rows = 1 Then
        MsgBox "No se ha procesado ningun periodo de produccion", vbInformation + vbOKCancel + vbDefaultButton1, xTitulo
        CmdBuscar.SetFocus
        Exit Sub
    End If
    
    
    xNumPag = 1
    Open Trim(App.Path) + "\00000001.txt" For Output As #1
    Cabecera2
    
    Dim A, B As Integer
    Dim xCad As String
    Dim xFila As Integer
    xFila = 9
    
    For B = 0 To Fg1.Rows - 1
        
        For A = 1 To Fg1.Cols - 1
            If A = 1 Then
                xCad = RellenarBlancos(Mid(Trim(Fg1.TextMatrix(B, A)), 1, 45), 45, 1) + "  "
            Else
                xCad = xCad + RellenarBlancos(Fg1.TextMatrix(B, A), 8, 2) + "  "
            End If
        Next A
        
        If B = 0 Then
            Print #1, Tab(6); xCad
            xFila = xFila + 1
            Print #1, "     ========================================================================================================================="
        Else
            Print #1, Tab(6); xCad
        End If
        
        xFila = xFila + 1
        If xFila = 60 Then
            Pie2
            xFila = 9
            xNumPag = xNumPag + 1
        End If
    Next B
    
    If xFila < 60 Then
        For A = 1 To 100
            Print #1, ' Format(xFila, "00")
            xFila = xFila + 1
            If xFila = 60 Then
                Pie2
                Exit For
            End If
        Next A
    End If
    
    Close #1
    
    'Dim xfrm As New Eps_VisorTexto.VisorTexto
    'xfrm.VerTexto Trim(App.Path) + "\00000001.txt", 62
    
    Dim xfrm As New Eps_VisorTexto.VisorTexto
    xfrm.VerTexto Trim(App.Path) + "\00000001.txt", 60, xCon
    
    Set xfrm = Nothing
End Sub

Sub Cabecera2()
    Dim xLen, PosX  As Integer
    PosX = (126 - 17)

    Print #1,
    Print #1, "     " + UCase(xNomEmp); Tab(PosX); "FECHA   :"; Format(Date, "dd/mm/yy")
    Print #1, "     RUC No : " + xNumRuc
    Print #1, "                                                   RESUMEN SEMANAL DE PRODUCCION"
    Print #1, "                                                   ============================="
    Print #1, " "  '-------------------------------------------------------           -------------------------------------------------------
    Print #1, "     PERIODO  :  DEL  " + Trim(TxtFchIni.Valor) + "  AL  " + Trim(TxtFchFin.Valor)
    Print #1, "     ========================================================================================================================="
              '1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
              '         1         2         3         4         5         6         7         8         9         10        11        12
End Sub

Sub Pie2()
    Print #1, "     ========================================================================================================================="
    Print #1, "                                                          PAGINA No : " + Format(xNumPag, "0000")
              '     -----------------------------------------------------XXXXXXXXXXXX0000----------------------------------------------------
End Sub
