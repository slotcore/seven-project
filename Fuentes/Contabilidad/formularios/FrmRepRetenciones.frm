VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRepRetenciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Reporte de Retenciones"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   990
      Left            =   2993
      TabIndex        =   13
      Top             =   3030
      Visible         =   0   'False
      Width           =   5145
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Top             =   435
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Procesando Asientos"
         Height          =   180
         Left            =   165
         TabIndex        =   15
         Top             =   165
         Width           =   1650
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   960
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   5130
         X2              =   5130
         Y1              =   15
         Y2              =   945
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   5115
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   5160
         Y1              =   975
         Y2              =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   900
      Left            =   45
      TabIndex        =   4
      Top             =   -75
      Width           =   11625
      Begin VB.CommandButton CmdExpPDT 
         Height          =   570
         Left            =   10140
         Picture         =   "FrmRepRetenciones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Exportar a PDT"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton CmdExp 
         Height          =   570
         Left            =   9465
         Picture         =   "FrmRepRetenciones.frx":0AC2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exportar a Excel"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Height          =   570
         Left            =   10755
         Picture         =   "FrmRepRetenciones.frx":15CC
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Height          =   570
         Left            =   8820
         Picture         =   "FrmRepRetenciones.frx":18D6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir"
         Top             =   210
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sujeto de Retención"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3330
         TabIndex        =   9
         Top             =   525
         Width           =   2250
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Agente de Retención"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3330
         TabIndex        =   8
         Top             =   240
         Width           =   2250
      End
      Begin VB.CommandButton Command1 
         Height          =   570
         Left            =   8175
         Picture         =   "FrmRepRetenciones.frx":1BE0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Mostrar"
         Top             =   210
         Width           =   615
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   1230
         TabIndex        =   0
         Top             =   195
         Width           =   1200
         _ExtentX        =   2117
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
         Valor           =   "19/10/2007"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   1230
         TabIndex        =   1
         Top             =   510
         Width           =   1200
         _ExtentX        =   2117
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
         Valor           =   "19/10/2007"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   225
         Width           =   870
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid2 
      Height          =   255
      Left            =   45
      TabIndex        =   3
      Top             =   855
      Width           =   11610
      _cx             =   20479
      _cy             =   450
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
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
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmRepRetenciones.frx":2022
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
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6270
      Left            =   45
      TabIndex        =   2
      Top             =   1110
      Width           =   11610
      _cx             =   20479
      _cy             =   11060
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   128
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmRepRetenciones.frx":20A2
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
End
Attribute VB_Name = "FrmRepRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SeEjecuto As Boolean

Private Sub CmdExp_Click()
    If Fg1.Rows = 1 Then Exit Sub
    ExportarDiario
End Sub

Private Sub CmdExpPDT_Click()
    Dim Rst As New ADODB.Recordset
    Dim NomArch, xCad As String
    Dim A As Integer
    
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "Falta especificar la fecha de inicio", vbInformation, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    If NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "Falta especificar la fecha final", vbInformation, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    If TxtFchIni.Valor > TxtFchFin.Valor Then
        MsgBox "La fecha inicial no puede ser mayor a la fecha final", vbInformation, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    
    If Fg1.Rows = 1 Then
        MsgBox "No se ha mostrado ninguna retención, haga click en el botón"
    End If
    
    RST_Busq Rst, "SELECT mae_cliente.numruc, con_retencion.numser, con_retencion.numdoc, con_retencion.fchemi, con_retencion.[imp], mae_documento.codsun," _
        & " vta_ventas.numser as numser2, vta_ventas.numdoc as numdoc2, vta_ventas.fchdoc, vta_ventas.imptotdoc, con_retencion.fchreg,con_retenciondet.impcob " _
        & " FROM ((con_retencion LEFT JOIN mae_cliente ON con_retencion.idpro = mae_cliente.id) LEFT JOIN ((con_retenciondet LEFT JOIN vta_ventas ON con_retenciondet.iddoc = vta_ventas.id) " _
        & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON con_retencion.id = con_retenciondet.id) LEFT JOIN con_tc ON con_retencion.fchemi = con_tc.fecha " _
        & " WHERE (((con_retencion.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (con_retencion.fchreg)<=CDate('" & TxtFchFin.Valor & "'))) " _
        & " ORDER BY con_retencion.fchemi", xCon

    If Rst.RecordCount <> 0 Then
        NomArch = "0621" + NumRUC + AnoTra + Format(TxtFchIni.Valor, "mm") + "R.txt"
       
        Open Trim(App.Path) + "\" + NomArch For Output As #1
    
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            xCad = ""
            xCad = xCad + Rst("numruc") + "|"
            xCad = xCad + Rst("numser") + "|"
            xCad = xCad + Mid(Rst("numdoc"), 3, 8) + "|"
            xCad = xCad + Format(Rst("fchemi"), "dd/mm/yyyy") + "|"
            xCad = xCad + Format(Rst("imp"), "0.00") + "|"
            xCad = xCad + Rst("codsun") + "|"
            xCad = xCad + Rst("numser2") + "|"
            xCad = xCad + Mid(Rst("numdoc2"), 3, 8) + "|"
            xCad = xCad + Format(Rst("fchdoc"), "dd/mm/yyyy") + "|"
            xCad = xCad + Format(Abs(NulosN(Rst("impcob"))), "0.00") + "|"
            
            Print #1, Trim(xCad)
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
    
    Close #1
    MsgBox "Las retenciones fueron exportadas al archivo para el PDT con éxito" & vbCr & "Ruta: " & Trim(App.Path) + "\" + NomArch, vbInformation, xTitulo
End Sub

Private Sub Command1_Click()
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "Falta especificar la fecha de inicio", vbInformation, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    If NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "Falta especificar la fecha final", vbInformation, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    If TxtFchIni.Valor > TxtFchFin.Valor Then
        MsgBox "La fecha inicial no puede ser mayor a la fecha final", vbInformation, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If

    If Option1.Value = True Then MostrarRetencionAgente
    If Option2.Value = True Then MostrarRetencionSujeto
End Sub

Sub MostrarRetencionSujeto()
    Dim Rst As New ADODB.Recordset
    Dim Rst1 As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    Dim xTotal As Double
    
    Dim A, B, C, xFila As Integer
    
    RST_Busq Rst, "SELECT DISTINCT con_retencion.idpro, con_retencion.tipo, mae_cliente.numruc, mae_cliente.nombre, con_retencion.fchreg" _
        & " FROM con_retencion LEFT JOIN mae_cliente ON con_retencion.idpro = mae_cliente.id WHERE (((con_retencion.tipo)=2) AND " _
        & " ((con_retencion.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (con_retencion.fchreg)<=CDate('" & TxtFchFin.Valor & "'))) " _
        & " ORDER BY mae_cliente.nombre", xCon

    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        xFila = 1
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            
            'IMPRIMIMOS ELNUMERO DE RUC Y EL NOMBRE DEL PROVEEDOR
            With Fg1
                .MergeCells = flexMergeFree
                .MergeRow(-1) = True
                .Cell(flexcpText, xFila, 1, xFila, 2) = "Nº R.U.C. : " + Trim(Rst("numruc"))
                .Cell(flexcpText, xFila, 3, xFila, 7) = "CLIENTE : " + Trim(Rst("nombre"))
            End With
            
            'HALLAMOS LAS RETENCIONES AFECTUADAS AL PROVEEDOR
            RST_Busq Rst1, "SELECT con_retencion.id, con_retencion.numreg, con_retencion.numser, con_retencion.numdoc, con_retencion.[imp], con_retencion.fchemi, mae_cliente.nombre, " _
                & " mae_cliente.numruc, con_retencion.idpro, con_retencion.tipo FROM con_retencion LEFT JOIN mae_cliente ON con_retencion.idpro = mae_cliente.id " _
                & " WHERE (((con_retencion.idpro)=" & Rst("idpro") & ") AND ((con_retencion.tipo)=2) AND ((con_retencion.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
                & " And (con_retencion.fchreg)<=CDate('" & TxtFchFin.Valor & "')))", xCon

            If Rst1.RecordCount <> 0 Then
                Rst1.MoveFirst
                xTotal = 0
                For B = 1 To Rst1.RecordCount
                    xFila = xFila + 1
                    Fg1.Rows = Fg1.Rows + 1
                    Fg1.TextMatrix(xFila, 1) = Rst1("numreg")
                    Fg1.TextMatrix(xFila, 2) = Rst1("numser")
                    Fg1.TextMatrix(xFila, 3) = Rst1("numdoc")
                    Fg1.TextMatrix(xFila, 4) = Rst1("fchemi")
                    Fg1.TextMatrix(xFila, 6) = Format(Rst1("imp"), "0.00")
                    
                    'HALLAMO LOS DOCUMENTOS INVOLUCRADOS EN LA RETENCION
                    RST_Busq Rst2, "SELECT con_retenciondet.id, mae_documento.codsun, vta_ventas.numser, vta_ventas.numdoc, " _
                        & " vta_ventas.fchdoc, vta_ventas.imptotdoc AS imptot, con_retencion.tipo, con_retenciondet.impret " _
                        & " FROM con_retencion LEFT JOIN ((con_retenciondet LEFT JOIN vta_ventas ON con_retenciondet.iddoc = vta_ventas.id) " _
                        & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON con_retencion.id = con_retenciondet.id " _
                        & " WHERE (((con_retenciondet.id)=" & Rst1("id") & ") AND ((con_retencion.tipo)=2))", xCon
                    
                    If Rst2.RecordCount <> 0 Then
                        Rst2.MoveFirst
                        For C = 1 To Rst2.RecordCount
                            'xFila = xFila + 1
                            Fg1.TextMatrix(xFila, 7) = NulosC(Rst2("codsun"))
                            Fg1.TextMatrix(xFila, 8) = NulosC(Rst2("numser"))
                            Fg1.TextMatrix(xFila, 9) = NulosC(Rst2("numdoc"))
                            Fg1.TextMatrix(xFila, 10) = NulosC(Rst2("fchdoc"))
                            
                            Fg1.TextMatrix(xFila, 11) = Format(NulosN(Rst2("imptot")), "0.00")
                            Fg1.TextMatrix(xFila, 12) = Format(NulosN(Rst2("impret")), "0.00")
                            
                            Rst2.MoveNext
                            If Rst2.EOF = True Then
                                xFila = xFila + 1
                                Fg1.Rows = Fg1.Rows + 1
                                Exit For
                            End If
                            xFila = xFila + 1
                            Fg1.Rows = Fg1.Rows + 1
                        Next C
                    End If
                    
                    xTotal = xTotal + Rst1("imp")
                    Rst1.MoveNext
                    If Rst1.EOF = True Then Exit For
                Next B
                Fg1.Rows = Fg1.Rows + 1
                xFila = xFila + 1
                Fg1.TextMatrix(xFila, 5) = "TOTAL ==>"
                Fg1.TextMatrix(xFila, 6) = Format(xTotal, "0.00")
                Fg1.Rows = Fg1.Rows + 1
                xFila = xFila + 1
            End If
            
            Rst.MoveNext
            
            If Rst.EOF = True Then Exit For
            xFila = xFila + 1
        Next A
    End If
End Sub


Sub MostrarRetencionAgente()
    Dim Rst As New ADODB.Recordset
    Dim Rst1 As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    Dim xTotal As Double
    
    Dim A, B, C, xFila As Integer
    
    RST_Busq Rst, "SELECT DISTINCTROW con_retencion.idpro, mae_prov.nombre, mae_prov.numruc, con_retencion.tipo " _
        & " FROM con_retencion LEFT JOIN mae_prov ON con_retencion.idpro = mae_prov.id " _
        & " WHERE (((con_retencion.fchemi)>=CDate('" & TxtFchIni.Valor & "') And (con_retencion.fchemi)<=CDate('" & TxtFchFin.Valor & "')) " _
        & " AND ((con_retencion.tipo)=1)) ORDER BY con_retencion.fchemi, mae_prov.nombre, mae_prov.numruc", xCon

    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        xFila = 1
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            
            'IMPRIMIMOS ELNUMERO DE RUC Y EL NOMBRE DEL PROVEEDOR
            With Fg1
                .MergeCells = flexMergeFree
                .MergeRow(-1) = True
                .Cell(flexcpText, xFila, 1, xFila, 2) = "Nº R.U.C. : " + Trim(Rst("numruc"))
                .Cell(flexcpText, xFila, 3, xFila, 7) = "PROVEEDOR : " + Trim(Rst("nombre"))
            End With
            
            'HALLAMOS LAS RETENCIONES AFECTUADAS AL PROVEEDOR
            RST_Busq Rst1, "SELECT con_retencion.id, con_retencion.numser, con_retencion.numdoc, con_retencion.[imp], " _
                & " con_retencion.fchemi, mae_prov.nombre, mae_prov.numruc, con_retencion.tipo FROM con_retencion LEFT JOIN " _
                & " mae_prov ON con_retencion.idpro = mae_prov.id Where (((con_retencion.idpro) = " & Rst("idpro") & ") " _
                & " And ((con_retencion.tipo) = 1)) ORDER BY con_retencion.fchemi, mae_prov.nombre, mae_prov.numruc", xCon

            'SELECT con_retencion.id, con_retencion.numser, con_retencion.numdoc, con_retencion.[imp], con_retencion.fchemi, " _
                & " mae_prov.nombre, mae_prov.numruc FROM con_retencion LEFT JOIN mae_prov ON con_retencion.idpro = mae_prov.id " _
                & " Where (((con_retencion.idpro) = " & Rst("idpro") & ")) ORDER BY con_retencion.fchemi, mae_prov.nombre, mae_prov.numruc", xCon

            If Rst1.RecordCount <> 0 Then
                Rst1.MoveFirst
                xTotal = 0
                For B = 1 To Rst1.RecordCount
                    xFila = xFila + 1
                    Fg1.Rows = Fg1.Rows + 1
                    Fg1.TextMatrix(xFila, 1) = Rst1("numser")
                    Fg1.TextMatrix(xFila, 2) = Rst1("numdoc")
                    Fg1.TextMatrix(xFila, 3) = Rst1("fchemi")
                    Fg1.TextMatrix(xFila, 5) = Format(Rst1("imp"), "0.00")
                    
                    'HALLAMO LOS DOCUMENTOS INVOLUCRADOS EN LA RETENCION
                    RST_Busq Rst2, "SELECT con_retenciondet.id, mae_documento.codsun, com_compras.numser, com_compras.numdoc, " _
                        & " com_compras.fchdoc, com_compras.imptot, con_retencion.tipo FROM con_retencion LEFT JOIN " _
                        & " (mae_documento RIGHT JOIN (con_retenciondet LEFT JOIN com_compras ON con_retenciondet.iddoc = com_compras.id) " _
                        & " ON mae_documento.id = com_compras.tipdoc) ON con_retencion.id = con_retenciondet.id " _
                        & " WHERE (((con_retenciondet.id)=" & Rst1("id") & ") AND ((con_retencion.tipo)=1))", xCon
                    
                    If Rst2.RecordCount <> 0 Then
                        Rst2.MoveFirst
                        For C = 1 To Rst2.RecordCount
                            Fg1.TextMatrix(xFila, 6) = Rst2("codsun")
                            Fg1.TextMatrix(xFila, 7) = Rst2("numser")
                            Fg1.TextMatrix(xFila, 8) = Rst2("numdoc")
                            Fg1.TextMatrix(xFila, 9) = Rst2("fchdoc")
                            Fg1.TextMatrix(xFila, 10) = Format(Rst2("imptot"), "0.00")
                            
                            Rst2.MoveNext
                            If Rst2.EOF = True Then Exit For
                            xFila = xFila + 1
                        Next C
                    End If
                    
                    xTotal = xTotal + Rst1("imp")
                    Rst1.MoveNext
                    If Rst1.EOF = True Then Exit For
                Next B
                Fg1.Rows = Fg1.Rows + 1
                xFila = xFila + 1
                Fg1.TextMatrix(xFila, 4) = "TOTAL ==>"
                Fg1.TextMatrix(xFila, 5) = Format(xTotal, "0.00")
                Fg1.Rows = Fg1.Rows + 1
                xFila = xFila + 1
            End If
            
            Rst.MoveNext
            
            If Rst.EOF = True Then Exit For
            xFila = xFila + 1
        Next A
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Option2.Value = True
    End If
End Sub

Private Sub Form_Load()
    'Fg1.ColWidth(11) = 0
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
    SeEjecuto = False
End Sub

Sub ExportarDiario()
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.WindowState = 1
    objExcel.Workbooks.Add  'Trim(App.Path) + "\RegCompras.xls"
    
    Frame5.Left = 2993
    Frame5.Top = 3030
    'Label3.Caption = "Exportando Documentos"
    Frame5.Visible = True
    
    
    ProgressBar1.Max = Fg1.Rows - 1
    
    With objExcel.ActiveSheet
        
        .Cells(1, 2) = NomEmp
        .Cells(1, 13) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        .Cells(4, 2) = "REGISTRO DE RETENCIONES"
        .Cells(5, 2) = "Del " + Trim(TxtFchIni.Valor) + " Al " + Trim(TxtFchFin.Valor)
        
        xFilas = 7
        For B = 1 To Fg1.Cols - 1
            .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(0, B)
        Next B
        
        xFilas = xFilas + 1
        For A = 1 To Fg1.Rows - 1
            ProgressBar1.Value = A
            Frame5.Refresh
            For B = 1 To Fg1.Cols - 1
                If B = 5 Or B = 10 Or B = 11 Then
                    If NulosN(Fg1.TextMatrix(A, B)) = 0 Then
                        .Cells(xFilas, B + 1) = ""
                    Else
                        .Cells(xFilas, B + 1) = Val(Fg1.TextMatrix(A, B))
                    End If
                Else
                    If Mid(Fg1.TextMatrix(A, 1), 1, 2) = "Nº" Then
                        If B = 1 Then
                            .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                        End If
                        If B = 3 Then
                            .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                        End If
                    Else
                        .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                    End If
                End If
            Next B
            xFilas = xFilas + 1
        Next A
    End With
    
    Frame5.Visible = False
    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Registro de Compras y Ventas"
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
End Sub

