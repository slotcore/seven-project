VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultaMayorAuxiliar 
   Caption         =   "Contabilidad - Mayor Auxiliar"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form2"
   ScaleHeight     =   7440
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fraconsctas 
      Caption         =   "Cuentas Seleccionadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3120
      Left            =   6360
      TabIndex        =   23
      Top             =   2400
      Visible         =   0   'False
      Width           =   5985
      Begin VB.CommandButton cmdEliminarOK 
         Height          =   630
         Left            =   960
         Picture         =   "FrmConsultaMayorAuxiliar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2415
         Width           =   765
      End
      Begin VB.CommandButton cmdOK 
         Height          =   630
         Left            =   90
         Picture         =   "FrmConsultaMayorAuxiliar.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2415
         Width           =   750
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   630
         Left            =   5160
         Picture         =   "FrmConsultaMayorAuxiliar.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2415
         Width           =   750
      End
      Begin VSFlex7Ctl.VSFlexGrid fgconsctas 
         Height          =   2070
         Left            =   90
         TabIndex        =   27
         Top             =   285
         Width           =   5805
         _cx             =   10239
         _cy             =   3651
         _ConvInfo       =   1
         Appearance      =   2
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
         Rows            =   50
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsultaMayorAuxiliar.frx":0716
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
   Begin VB.Frame Framonedas 
      Caption         =   "Monedas"
      ForeColor       =   &H00800000&
      Height          =   825
      Left            =   6240
      TabIndex        =   9
      Top             =   600
      Width           =   1455
      Begin VB.OptionButton OptSoles 
         Caption         =   "Soles"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   210
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton OptDolares 
         Caption         =   "Dolares"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Visualización"
      ForeColor       =   &H00800000&
      Height          =   1065
      Left            =   10800
      TabIndex        =   17
      Top             =   375
      Width           =   2370
      Begin VB.OptionButton optresumidoctames 
         Caption         =   "Resumido por Cuenta Mes"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton optresumidocta 
         Caption         =   "Resumido por Cuenta"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optdetallado 
         Caption         =   "Detallado"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   225
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.Frame Framconscta 
      Caption         =   "Consultar por "
      ForeColor       =   &H00800000&
      Height          =   825
      Left            =   7800
      TabIndex        =   12
      Top             =   600
      Width           =   2940
      Begin VB.OptionButton optconscta 
         Caption         =   "Cuenta"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   225
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optconsel 
         Caption         =   "Seleccion"
         Height          =   195
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Optcontod 
         Caption         =   "Todos"
         Height          =   195
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   825
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6270
      Left            =   0
      TabIndex        =   20
      Top             =   1440
      Width           =   13170
      _cx             =   23230
      _cy             =   11060
      _ConvInfo       =   1
      Appearance      =   2
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConsultaMayorAuxiliar.frx":0792
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
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   -15
      TabIndex        =   0
      Top             =   600
      Width           =   6210
      Begin VB.CommandButton CmdMuestra 
         Height          =   600
         Left            =   5220
         Picture         =   "FrmConsultaMayorAuxiliar.frx":0913
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   165
         Width           =   900
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   600
         TabIndex        =   5
         Top             =   480
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
      End
      Begin VB.CommandButton CmdBusCta 
         Height          =   240
         Left            =   4920
         Picture         =   "FrmConsultaMayorAuxiliar.frx":0D55
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   240
      End
      Begin VB.TextBox TxtIdcuenta 
         Height          =   300
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "TxtIdcuenta"
         Top             =   180
         Width           =   4575
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   3960
         TabIndex        =   7
         Top             =   495
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
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   3720
         TabIndex        =   6
         Top             =   525
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Left            =   105
         TabIndex        =   4
         Top             =   525
         Width           =   240
      End
      Begin VB.Label LblIdcuenta 
         AutoSize        =   -1  'True
         Caption         =   "LblIdcuenta"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4245
         TabIndex        =   16
         Top             =   300
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Libro"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   210
         Width           =   345
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13560
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayorAuxiliar.frx":0E87
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayorAuxiliar.frx":13CB
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayorAuxiliar.frx":175D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayorAuxiliar.frx":1AEF
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayorAuxiliar.frx":1C73
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayorAuxiliar.frx":20C7
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayorAuxiliar.frx":21DF
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayorAuxiliar.frx":2723
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayorAuxiliar.frx":2C67
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayorAuxiliar.frx":2D7B
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayorAuxiliar.frx":2E8F
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayorAuxiliar.frx":32E3
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayorAuxiliar.frx":344F
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mostrar Registro"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a Excel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmConsultaMayorAuxiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub BuscaCtaDetallado(Optional ccuenta As String)
    
    Dim Rst As New ADODB.Recordset
    Dim Rsaux As New ADODB.Recordset
    Dim rstmp As New ADODB.Recordset
    
    Dim swaviso As Byte
    Dim totgendebe As Double
    Dim totgenhaber As Double
    Dim totsaldoant As Double
    
    Dim saldoant As Double
    Dim X As Integer
    Dim xMes As Integer
    Dim TotHabSol As Double
    Dim TotDebSol As Double
    Dim nuevacta As Byte
    Dim totdeb As Double
    Dim tothab As Double
    
    Dim A As Integer
                
    xMes = Month(CDate(TxtFchIni.Valor))
With Fg1
            .Cols = 13
            .Rows = 1
            
            .TextMatrix(0, 1) = "Mes"
            .TextMatrix(0, 2) = "Asiento"
            .TextMatrix(0, 3) = "Libro"
            .TextMatrix(0, 4) = "TD"
            .TextMatrix(0, 5) = "Nº Documento"
            .TextMatrix(0, 6) = "Fch. Reg."
            .TextMatrix(0, 7) = "T.C."
            .TextMatrix(0, 8) = "Cuenta"
            .TextMatrix(0, 9) = "Descripcion"
            .TextMatrix(0, 10) = "DEBE S/."
            .TextMatrix(0, 11) = "HABER S/."
            .TextMatrix(0, 12) = "SALDO "
            
            .ColAlignment(7) = flexAlignRightBottom
            .ColWidth(1) = 435
            .ColWidth(2) = 700
            .ColWidth(3) = 1200
            .ColWidth(4) = 495
            .ColWidth(5) = 1500
            .ColWidth(6) = 1000
            .ColWidth(7) = 600
            .ColWidth(8) = 810
            .ColWidth(9) = 2595
            .ColWidth(10) = 960
            .ColWidth(11) = 960
            .ColWidth(12) = 960
        End With
    

If optconscta.Value = True Or Optcontod.Value = True Then
           
                Fg1.Rows = 1
                
                If optconscta.Value = True Then
                    RST_Busq Rsaux, " SELECT con_planctas.id, con_planctas.cuenta, con_planctas.descripcion FROM con_planctas " _
                    & " WHERE con_planctas.cuenta Like '" & ccuenta & "%' ORDER BY con_planctas.cuenta ", xCon
                Else
                    RST_Busq Rsaux, " SELECT con_planctas.id, con_planctas.cuenta, con_planctas.descripcion FROM con_planctas " _
                    & "  ORDER BY con_planctas.cuenta ", xCon
                End If
           
           nuevacta = 1
           Do While Not Rsaux.EOF
           
                        With Fg1
                                          
                             .AddItem ""
                             .AddItem ""
                             .TextMatrix(.Rows - 1, 8) = Rsaux("cuenta")
                             .TextMatrix(.Rows - 1, 9) = Rsaux("descripcion")
                             
                        End With
                  
                  'Obtenemos el Saldo Anterior
                  If nuevacta = 1 Then
                    RST_Busq rstmp, " SELECT Sum(con_diario.impdebsol) AS SumaDeimpdebsol, Sum(con_diario.imphabsol) AS SumaDeimphabsol " _
                                     & " FROM con_diario WHERE con_diario.idmes < " & xMes & " AND con_diario.idcue=" & Rsaux!id & "", xCon
                     
                     
                     With Fg1
                     If rstmp.RecordCount > 0 Then
                         swaviso = 1
                         
                             .AddItem ""
                             
                             .TextMatrix(.Rows - 1, 9) = "Saldo Anterior"
                             .TextMatrix(.Rows - 1, 12) = Format(NulosN(rstmp("Sumadeimpdebsol")) - NulosN(rstmp("Sumadeimphabsol")), "0.00")
                              saldoant = NulosN(rstmp("Sumadeimpdebsol")) - NulosN(rstmp("Sumadeimphabsol"))
                              totsaldoant = totsaldoant + saldoant
                     Else
                            .TextMatrix(.Rows - 1, 10) = "0.00"
                            .TextMatrix(.Rows - 1, 11) = "0.00"
                            .TextMatrix(.Rows - 1, 12) = "0.00"
                            swaviso = 0
                     End If
                     End With
                         nuevacta = 0
                  End If
                    
                   'Unimos todas las cuentas para
                    RST_Busq Rst, " SELECT con_diario.idlib, con_diario.idmes, con_diario.numasi, mae_libros.descripcion, mae_documento.abrev, [com_compras]![numser]+ '-' + [com_compras]![numdoc] AS numdoc, com_compras.fchdoc, con_planctas.cuenta, con_diario.impdebsol, con_diario.imphabsol,con_tc.impven ,con_planctas.descripcion as [Nomcuenta] " _
                                  & " FROM (mae_libros INNER JOIN (con_planctas INNER JOIN ((con_diario INNER JOIN mae_documento ON con_diario.iddoc = mae_documento.id) INNER JOIN com_compras ON con_diario.idmov = com_compras.id) ON con_planctas.id = con_diario.idcue) ON mae_libros.id = con_diario.idlib) INNER JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
                                & " WHERE con_diario.idlib = 1  AND con_diario.idcue = " & Rsaux("id") & "" _
                                & " UNION SELECT con_diario.idlib, con_diario.idmes, con_diario.numasi,  mae_libros.descripcion, mae_documento.abrev, [vta_ventas]![numser]+ '-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.fchdoc,con_planctas.cuenta, con_diario.impdebsol, con_diario.imphabsol, con_tc.impven ,con_planctas.descripcion as [Nomcuenta]" _
                                & " FROM con_tc RIGHT JOIN (mae_libros INNER JOIN (con_planctas INNER JOIN ((con_diario INNER JOIN mae_documento ON con_diario.iddoc = mae_documento.id) LEFT JOIN vta_ventas ON con_diario.idmov = vta_ventas.id) ON con_planctas.id = con_diario.idcue) ON mae_libros.id = con_diario.idlib) ON con_tc.fecha = vta_ventas.fchdoc " _
                                & " WHERE (((con_diario.idlib) = 2) And (Not (con_diario.iddoc) = 7 And Not (con_diario.iddoc) = 8)) And con_diario.idcue =" & Rsaux("id") & " and con_diario.idmes >= " & xMes & "", xCon
                                                         
            
            
            If Rst.RecordCount <> 0 Then
                                                                                                                
                For A = 1 To Rst.RecordCount
                                                            
                    Fg1.Rows = Fg1.Rows + 1
                    
                    Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("idmes")
                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("numasi")
                    Fg1.TextMatrix(Fg1.Rows - 1, 3) = Rst("descripcion")
                    Fg1.TextMatrix(Fg1.Rows - 1, 4) = Rst("abrev")
                    Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(Rst("numdoc"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(Rst("fchdoc"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosN(Rst("impven"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 8) = Rst("cuenta")
                    Fg1.TextMatrix(Fg1.Rows - 1, 9) = Rst("nomcuenta")
                    Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(Rst("impdebsol"), "0.00")
                    Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(Rst("imphabsol"), "0.00")
                    
                    'If Rst("idmon") = 1 Then
                    '    If NulosN(Rst("impven")) <> 0 Then
                    '        Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(Rst("impdebsol") / Rst("impven"), "0.00")
                    '        Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(Rst("imphabsol") / Rst("impven"), "0.00")
                        'End If
                    ' Else
                    '     Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(Rst("impdebdol"), "0.00")
                    '     Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(Rst("imphabdol"), "0.00")
                   
                    ' End If
                    
                    TotDebSol = TotDebSol + NulosN(Rst("impdebsol"))
                    TotHabSol = TotHabSol + NulosN(Rst("imphabsol"))
                    
                    totgendebe = totgendebe + NulosN(Rst("impdebsol"))
                    totgenhaber = totgenhaber + NulosN(Rst("imphabsol"))
                    Rst.MoveNext
                Next A
                   
                   
            End If
            'MOSTRAMOS SUBTOTALES
                            With Fg1
                            .AddItem ""
                            .TextMatrix(.Rows - 1, 9) = "Total por Cta -> " & Rsaux("cuenta")
                            .TextMatrix(.Rows - 1, 10) = Format(TotDebSol, "0.00")
                            .TextMatrix(.Rows - 1, 11) = Format(TotHabSol, "0.00")
                            .AddItem ""
                            .TextMatrix(.Rows - 1, 9) = "Saldo por Cta -> " & Rsaux("cuenta")
                            'Estableciendo el slado por cuenta
                            If totdeb > tothab Then
                            .TextMatrix(.Rows - 1, 12) = (TotDebSol - TotHabSol) + saldoant
                            Else
                            .TextMatrix(.Rows - 1, 12) = (TotHabSol - TotDebSol) + saldoant
                            End If
                            
                            
                            End With
                            TotDebSol = 0
                            TotHabSol = 0
                            saldoant = 0
                nuevacta = 1
                Rsaux.MoveNext
            Loop
                    
                    With Fg1
                        .AddItem ""
                        .AddItem ""
                        .TextMatrix(.Rows - 1, 9) = "Total General -> "
                        .TextMatrix(.Rows - 1, 10) = Format(totgendebe, "0.00")
                        .TextMatrix(.Rows - 1, 11) = Format(totgenhaber, "0.00")
                           
                           .AddItem ""
                          .TextMatrix(.Rows - 1, 9) = "Saldo General -> "
                            'Estableciendo el slado por cuenta
                            If totgendebe > totgenhaber Then
                            .TextMatrix(.Rows - 1, 12) = (totgendebe - totgenhaber) + totsaldoant
                            Else
                            .TextMatrix(.Rows - 1, 12) = (totgenhaber - totgendebe) + totsaldoant
                            End If
                    End With
     End If

    
    Set Rst = Nothing
    Set Rsaux = Nothing
    Set rstmp = Nothing
    
End Sub
Sub BuscaCtaResumen(ccuenta As String)
    Dim Rst As New ADODB.Recordset
    Dim Rsaux As New ADODB.Recordset
    Dim rstmp As New ADODB.Recordset
    Dim A As Integer
        Dim totdebe As Double
    Dim tothaber As Double
    Dim saldoant As Double
    Dim X As Integer
    
    
    If optconscta.Value = True Or Optcontod.Value = True Then
                
     If optresumidocta.Value = True Then
                
        If optconscta.Value = True Then
            RST_Busq Rsaux, " SELECT con_planctas.cuenta, con_planctas.descripcion FROM con_planctas " _
                          & " WHERE con_planctas.cuenta Like '" & ccuenta & "%' ORDER BY con_planctas.cuenta ", xCon
        Else
            RST_Busq Rsaux, " SELECT con_planctas.cuenta, con_planctas.descripcion FROM con_planctas " _
                          & "  ORDER BY con_planctas.cuenta ", xCon
        End If
                
        Do While Not Rsaux.EOF
        
                        
           'MOSTRAMOS TODOS LOS MOVIMIENTOS EN EL RANGO DE LA FECHA
            RST_Busq Rst, " SELECT  con_planctas.cuenta, con_planctas.descripcion, Sum(con_diario.impdebsol) AS SumaDebe, Sum(con_diario.imphabsol) AS SumaHaber" _
                    & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN mae_documento ON con_diario.iddoc=mae_documento.id) ON con_planctas.id=con_diario.idcue " _
                    & " WHERE con_diario.idmes >= " & Month(CDate(TxtFchIni.Valor)) & " and con_diario.idmes <= " & Month(CDate(TxtFchFin.Valor)) & "" _
                    & " GROUP BY con_planctas.cuenta, con_planctas.descripcion " _
                    & " HAVING con_planctas.cuenta ='" & Trim(Rsaux("cuenta")) & "'", xCon
            
            If Rst.RecordCount > 0 Then
              Do While Not Rst.EOF
                    With Fg1
                        .AddItem ""
                        .TextMatrix(.Rows - 1, 1) = "A " & Format(Me.TxtFchFin.Valor, "mmmm-yyyy")
                        .TextMatrix(.Rows - 1, 2) = Rst("cuenta")
                        .TextMatrix(.Rows - 1, 3) = Rst("descripcion")
                        .TextMatrix(.Rows - 1, 5) = Rst("Sumadebe")
                        .TextMatrix(.Rows - 1, 6) = Rst("sumaHaber")
                        
                         totdebe = Rst("sumadebe")
                        tothaber = Rst("sumaHaber")
                       
                       'OBTENEMOS EL SALDO ANTERIOR ACUMULADO
                        RST_Busq rstmp, " SELECT con_planctas.cuenta, con_planctas.descripcion, Sum(con_diario.impdebsol) AS SumaDebe , Sum(con_diario.imphabsol) AS SumaHaber " _
                        & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN mae_documento ON con_diario.iddoc=mae_documento.id) ON con_planctas.id=con_diario.idcue " _
                        & " WHERE con_diario.idmes < " & Month(CDate(TxtFchIni.Valor)) & "" _
                        & " GROUP BY  con_planctas.cuenta, con_planctas.descripcion " _
                        & " HAVING con_planctas.cuenta ='" & Trim(Rsaux("cuenta")) & "'", xCon
                        
                        If rstmp.RecordCount > 0 Then
                            saldoant = Round(rstmp("sumadebe") - rstmp("sumaHaber"), 2)
                        Else
                           saldoant = 0
                        End If
                        .TextMatrix(.Rows - 1, 4) = saldoant
                        'Actualizamos el Saldo Actual
                        .TextMatrix(.Rows - 1, 7) = (totdebe - tothaber) + saldoant
                    End With
                Rst.MoveNext
              Loop
            End If
            Rsaux.MoveNext
       Loop
    
    ElseIf optresumidoctames.Value = True Then
        

       
       RST_Busq Rsaux, " SELECT con_planctas.cuenta, con_planctas.descripcion FROM con_planctas " _
                     & " WHERE con_planctas.cuenta Like '" & ccuenta & "%' ORDER BY con_planctas.cuenta ", xCon
    
        Do While Not Rsaux.EOF
        

                                   
           'MOSTRAMOS TODOS LOS MOVIMIENTOS EN EL RANGO DE LA FECHA
            RST_Busq Rst, " SELECT con_diario.idmes, con_planctas.cuenta, con_planctas.descripcion, Sum(con_diario.impdebsol) AS SumaDebe, Sum(con_diario.imphabsol) AS SumaHaber " _
                          & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN mae_documento ON con_diario.iddoc = mae_documento.id) ON con_planctas.id = con_diario.idcue " _
                          & " GROUP BY con_diario.idmes, con_planctas.cuenta, con_planctas.descripcion " _
                          & " HAVING con_planctas.cuenta ='" & Trim(Rsaux("cuenta")) & "'", xCon
            
            If Rst.RecordCount > 0 Then
                        
                        'OBTENEMOS EL SALDO ANTERIOR ACUMULADO POR CUENTA UNA VEZ
                        RST_Busq rstmp, " SELECT con_planctas.cuenta, con_planctas.descripcion, Sum(con_diario.impdebsol) AS SumaDebe , Sum(con_diario.imphabsol) AS SumaHaber " _
                        & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN mae_documento ON con_diario.iddoc=mae_documento.id) ON con_planctas.id=con_diario.idcue " _
                        & " WHERE con_diario.idmes < " & Month(CDate(TxtFchIni.Valor)) & "" _
                        & " GROUP BY  con_planctas.cuenta, con_planctas.descripcion " _
                        & " HAVING con_planctas.cuenta ='" & Trim(Rsaux("cuenta")) & "'", xCon
                        
                        If rstmp.RecordCount > 0 Then
                            saldoant = Round(rstmp("sumadebe") - rstmp("sumaHaber"), 2)
                        Else
                           saldoant = 0
                        End If
              
              Do While Not Rst.EOF
                    With Fg1
                        .AddItem ""
                        '.TextMatrix(.Rows - 1, 1) = NomMes(Rst("idmes"))
                        .TextMatrix(.Rows - 1, 2) = Rst("cuenta")
                        .TextMatrix(.Rows - 1, 3) = Rst("descripcion")
                        
                        .TextMatrix(.Rows - 1, 4) = saldoant
                        .TextMatrix(.Rows - 1, 5) = Rst("Sumadebe")
                        .TextMatrix(.Rows - 1, 6) = Rst("sumaHaber")
                        
                         totdebe = Rst("sumadebe")
                        tothaber = Rst("sumaHaber")
                       
                       'Actualizamos el Saldo Actual
                         .TextMatrix(.Rows - 1, 7) = (totdebe - tothaber) + saldoant
                         'El nuevo saldo anterior
                          saldoant = Val(.TextMatrix(.Rows - 1, 7))
                    End With
                    Rst.MoveNext
              Loop
            End If
            Rsaux.MoveNext
       Loop
       
    End If
    ElseIf optconsel.Value = True Then
        'Clasificar por seleccion de cuentas
    End If
Set Rst = Nothing
Set Rsaux = Nothing
Set rstmp = Nothing

End Sub


Private Sub CmdBusCta_Click()
Dim xform As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Id":           xCampos(0, 1) = "id":                   xCampos(0, 2) = "1000":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Cuenta":       xCampos(1, 1) = "cuenta":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Descripcion":  xCampos(2, 1) = "descripcion":          xCampos(2, 2) = "4000":         xCampos(2, 3) = "C"
    
    xform.SqlCad = "SELECT * FROM con_planctas ORDER BY cuenta"
    
    xform.Titulo = "Buscando Cuentas"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdcuenta.Text = xRs("cuenta") + " - " + xRs("descripcion")
        LblIdcuenta.Caption = xRs("id")
    End If
    
    TxtFchIni.SetFocus
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdMuestra_Click()
    If TxtIdcuenta.Text = "" Then
        MsgBox "No ha especificado la cuenta a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdcuenta.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtFchIni.Valor) = "" Or NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "El rango de fechas del periodo a consultar es invalido", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "La fecha de inicio del periodo no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    Dim pos As Integer
    Dim ccuenta As String
    pos = InStr(TxtIdcuenta, "-")
    ccuenta = Trim(Mid(TxtIdcuenta, 1, pos - 1))
    
    
    If optresumidocta.Value = True Or optresumidoctames.Value = True Then
        With Fg1
            .Cols = 8
            .Rows = 1
            
            .TextMatrix(0, 1) = "Mes"
            .TextMatrix(0, 2) = "Cuenta"
            .TextMatrix(0, 3) = "Descripcion"
            .TextMatrix(0, 4) = "Saldo Ant"
            .TextMatrix(0, 5) = "Debe"
            .TextMatrix(0, 6) = "Haber"
            .TextMatrix(0, 7) = "Saldo Act"
            .ColWidth(1) = 1500
            .ColWidth(2) = 1000
            .ColWidth(3) = 3000
            .ColWidth(4) = 1300
            .ColWidth(5) = 1300
            .ColWidth(6) = 1300
            .ColWidth(7) = 1300
            .ColAlignment(7) = flexAlignRightBottom
        End With
    End If
    
    
    
    If (optconscta.Value = True Or optconsel.Value = True Or Optcontod.Value = True) And optdetallado.Value = True Then
        'Modulo Detallado por Cuenta, Cuentas Seleccionadas y Todas
        BuscaCtaDetallado (ccuenta)
    Else
        BuscaCtaResumen (ccuenta)
    End If
     
End Sub


Private Sub Form_Load()
    Blanquea
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    Fg1.ColWidth(11) = 0
    Fg1.ColWidth(12) = 0
    OptSoles.Value = True
End Sub

Sub Blanquea()
    
    TxtIdcuenta.Text = ""
    LblIdcuenta.Caption = ""
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    
End Sub


Private Sub OptDolares_Click()
    If OptDolares.Value = True Then
        Fg1.ColWidth(9) = 0
        Fg1.ColWidth(10) = 0
        Fg1.ColWidth(11) = 1000
        Fg1.ColWidth(12) = 1000
    End If
End Sub

Private Sub OptSoles_Click()
    If OptSoles.Value = True Then
        Fg1.ColWidth(9) = 1000
        Fg1.ColWidth(10) = 1000
        Fg1.ColWidth(11) = 0
        Fg1.ColWidth(12) = 0
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   If Button.Index = 6 Then
        Unload Me
    End If
    If Button.Index = 1 Then
        Call CmdMuestra_Click
    End If
    
    If Button.Index = 3 Then 'IMPRESION
        
    End If
    If Button.Index = 4 Then 'ENVIAR EXCEL
        
    End If
End Sub



Private Sub TxtIdcuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtFchIni.SetFocus
End Sub

Private Sub TxtIdcuenta_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
         CmdBusCta_Click
    End If

End Sub

