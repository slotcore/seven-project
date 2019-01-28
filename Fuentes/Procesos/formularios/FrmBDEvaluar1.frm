VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmBDEvaluar1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herramientas - Evaluar Base de Datos"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8670
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   0
      TabIndex        =   10
      Top             =   390
      Width           =   8085
      Begin VB.CommandButton CmdBusArch2 
         Height          =   240
         Left            =   7050
         Picture         =   "FrmBDEvaluar1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   540
         Width           =   240
      End
      Begin VB.CommandButton CmdBusArch 
         Height          =   240
         Left            =   7050
         Picture         =   "FrmBDEvaluar1.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   225
         Width           =   240
      End
      Begin VB.TextBox TxtArchivo2 
         Height          =   300
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "TxtArchivo2"
         Top             =   510
         Width           =   6015
      End
      Begin VB.TextBox TxtArchivo 
         Height          =   300
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "TxtArchivo"
         Top             =   195
         Width           =   6015
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Data Destino"
         Height          =   195
         Left            =   135
         TabIndex        =   16
         Top             =   540
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Origen"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   225
         Width           =   855
      End
   End
   Begin VB.Frame FraProgreso 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   2310
      TabIndex        =   2
      Top             =   3165
      Visible         =   0   'False
      Width           =   5760
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   345
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   3
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   -60
         X2              =   6360
         Y1              =   675
         Y2              =   690
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   5895
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5745
         X2              =   5745
         Y1              =   -90
         Y2              =   4800
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Interrumpir = ESC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   4140
         TabIndex        =   6
         Top             =   75
         Width           =   1530
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   75
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base de Datos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   1185
         TabIndex        =   4
         Top             =   75
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4860
         Top             =   90
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":0264
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":07A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":0B3A
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":0CBE
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":1112
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":122A
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":176E
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":1CB2
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":1DC6
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":1EDA
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":232E
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":249A
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":29E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar1.frx":2CFC
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Seleccionar ]"
      Height          =   915
      Left            =   8145
      TabIndex        =   7
      Top             =   390
      Width           =   1665
      Begin VB.OptionButton OptTipo 
         Caption         =   "Pendiente"
         Height          =   240
         Index           =   1
         Left            =   165
         TabIndex        =   9
         Top             =   570
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Toda la Lista"
         Height          =   240
         Index           =   0
         Left            =   165
         TabIndex        =   8
         Top             =   270
         Width           =   1335
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   4575
      Left            =   30
      TabIndex        =   0
      Top             =   1365
      Width           =   9780
      _cx             =   17251
      _cy             =   8070
      _ConvInfo       =   1
      Appearance      =   1
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
      ForeColorSel    =   16777215
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmBDEvaluar1.frx":308E
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
Attribute VB_Name = "FrmBDEvaluar1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConexOri As New ADODB.Connection
Dim ConexDest As New ADODB.Connection

Dim BAND_INTERRUMPIR As Boolean


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    End If
End Sub

Private Sub Form_Load()

    CentrarFrm Me
    
    TxtArchivo.Text = ""
    TxtArchivo2.Text = ""
    pConfigurarGrilla
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BAND_INTERRUMPIR = True '--interrumpir
    
    If ConexOri.State = 1 Then ConexOri.Close
    If ConexDest.State = 1 Then ConexDest.Close
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 5 Then pExportarExcel
    If Button.Index = 6 Then pImprimir
    If Button.Index = 8 Then
        BAND_INTERRUMPIR = True
        Unload Me
    End If
End Sub


'------------

Private Sub pConsultar()
    On Error GoTo error
    
    Dim rstOri As New ADODB.Recordset
    Dim rstDest As New ADODB.Recordset
    
    Dim rstOriCampo As New ADODB.Recordset
    Dim rstDestCampo As New ADODB.Recordset
    
    Dim mCampo&
    
    Dim nSQL As String
    
    Fg1.Rows = Fg1.FixedRows
    
    '*********************************************************************
    If NulosC(TxtArchivo.Text) = "" Then
        MsgBox "Falta especificar la Base de Datos Origen", vbExclamation, xTitulo
        TxtArchivo.SetFocus
        Exit Sub
    ElseIf NulosC(TxtArchivo2.Text) = "" Then
        MsgBox "Falta especificar la Base de Datos Origen", vbExclamation, xTitulo
        TxtArchivo2.SetFocus
        Exit Sub
    End If
    '*********************************************************************
    
    '--establecer la consulta
    nSQL = "SELECT MSysObjects.Name as descripcion FROM MSysObjects WHERE (((MSysObjects.Type)=1) AND ((MSysObjects.Flags)=0)) OR (((MSysObjects.Database) Is Not Null)) ORDER BY MSysObjects.Name;"
    
    RST_Busq rstOri, nSQL, ConexOri '--lista de tablas origen
    RST_Busq rstDest, nSQL, ConexDest '--lista de tablas destino
    
    BAND_INTERRUMPIR = False
    If rstOri.RecordCount = 0 Then
        MsgBox "No hay registros en la Base Origen", vbInformation, xTitulo
        Exit Sub
    End If
    PosicionarProgBar
    PgBar.Min = 0
    PgBar.Max = rstOri.RecordCount
    
    '*********************************************************************
    Do While Not rstOri.EOF
        DoEvents
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        
        PgBar.Value = CLng(rstOri.Bookmark)
        
        Fg1.Rows = Fg1.Rows + 1
        
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = rstOri("descripcion")
        
        rstDest.MoveFirst
        rstDest.Find "descripcion='" & rstOri("descripcion") & "'"
        
        If rstDest.EOF = False And rstDest.BOF = False Then
            
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = rstDest("descripcion")
            
            nSQL = "select TOP 1 * FROM " & rstOri("descripcion")
            RST_Busq rstOriCampo, nSQL, ConexOri  '--lista de campos origen
            RST_Busq rstDestCampo, nSQL, ConexDest  '--lista de campos destino
            
            For mCampo = 0 To rstOriCampo.Fields.Count - 1
            
                DoEvents
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(Fg1.Rows - 1, 2) = rstOriCampo.Fields(mCampo).Name
                
                If RstRegistroBuscaCampo(rstDestCampo, rstOriCampo.Fields(mCampo).Name) = True Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 5) = rstOriCampo.Fields(mCampo).Name

                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = "!!!!!!OK"
                
                    If rstOriCampo(mCampo).Type <> rstDestCampo(rstOriCampo.Fields(mCampo).Name).Type Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 6) = "TIPO DATO!!!!!!"
                    ElseIf rstOriCampo(mCampo).DefinedSize <> rstDestCampo(rstOriCampo.Fields(mCampo).Name).DefinedSize Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 6) = "LONGITUD!!!!!!"
                    Else
                        '--eliminando las filas cuando tipo sea pendientes
                        If OptTipo(1).Value = True Then
                            Fg1.Rows = Fg1.Rows - 1
                        End If
                    End If
                    
                Else
                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = "FALTA!!!!!!"
                End If
                
            Next mCampo
            
        Else
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = "FALTA!!!!!!"
        
        End If
        '--eliminando las filas cuando tipo sea pendientes
        If Fg1.TextMatrix(Fg1.Rows - 1, 6) = "" And OptTipo(1).Value = True Then
            Fg1.Rows = Fg1.Rows - 1
        End If
        
        rstOri.MoveNext
    Loop
    '*********************************************************************
   '
SALIR:
    FraProgreso.Visible = False
    Set rstOri = Nothing:       Set rstOriCampo = Nothing
    Set rstDest = Nothing:      Set rstDestCampo = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
'Resume
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    
    Set rstOri = Nothing:       Set rstOriCampo = Nothing
    Set rstDest = Nothing:      Set rstDestCampo = Nothing
    
    SHOW_ERROR Me.Name, "pConsultar"
    
End Sub

Private Sub pExportarExcel()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    X_PRINT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "EVALUAR BASE DE DATOS", "Origen: " & TxtArchivo.Text, "Destino: " & TxtArchivo2.Text, "Evaluar Base de Datos"
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
End Sub


Private Sub pImprimir()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, "EVALUAR BASE DE DATOS", "Origen: " & TxtArchivo.Text, "Destino: " & TxtArchivo2.Text, True, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"

End Sub

Private Sub pConfigurarGrilla()
    '===================================================================================================
    'Propósito: Establecer los encabezados del grid
    '
    'Entradas:  Ninguna
    '
    'Resultados: Grilla con Encabezado
    '===================================================================================================
    
    With Fg1
        '-----
        .Cols = 7
                 
        '.FrozenCols = Q_POS_MES_INICIO - 1
        .ColWidth(0) = 200
        .FrozenCols = 0
        .Rows = 2
        .FixedRows = 2
        
        UNIR_CELDAS Fg1, 0, 1, 0, 2, "Origen", flexAlignCenterCenter
        UNIR_CELDAS Fg1, 0, 4, 0, 5, "Destino", flexAlignCenterCenter
        UNIR_CELDAS Fg1, 0, 6, 0, 6, "Observación", flexAlignCenterCenter
        
        .TextMatrix(1, 1) = "Tabla":    .ColWidth(1) = 2500:   .ColAlignment(1) = flexAlignLeftBottom:        .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 2) = "Campo":    .ColWidth(2) = 1500:   .ColAlignment(2) = flexAlignLeftBottom:        .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(1, 3) = " ":        .ColWidth(3) = 150:
        
        .TextMatrix(1, 4) = "Tabla":    .ColWidth(4) = 2500:   .ColAlignment(4) = flexAlignLeftBottom:        .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 5) = "Campo":    .ColWidth(5) = 1500:   .ColAlignment(5) = flexAlignLeftBottom:        .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(1, 6) = " ":      .ColWidth(6) = 1000:   .ColAlignment(6) = flexAlignLeftBottom:        .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftBottom
                
    End With
    DoEvents
End Sub

Private Sub PosicionarProgBar()
    '--POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
    '    FraProgreso.Width = 5820
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    Me.PgBar.Value = 1
    FraProgreso.Visible = True
End Sub


Private Sub CmdBusArch_Click()
    'CommonDialog1.CancelError = True
    'Especificar las extensiones a usar
    CommonDialog1.DefaultExt = "*.mdb"
    CommonDialog1.Filter = "Microsoft Oficce Access (*.mdb)|*.mdb"
    CommonDialog1.ShowOpen
    If Err Then
        'Cancelada la operación de abrir
    Else
        TxtArchivo.Text = CommonDialog1.FileName
    End If
    
    '--si la base de datos principal existe
    If ArchivoExiste(TxtArchivo.Text) = False Then
        MsgBox "No existe la ruta a la Base de Datos Origen", vbCritical, "Mensaje..."
        Exit Sub
    End If
    
    OPEN_CONEX_TMP ConexOri, TxtArchivo.Text
    If ConexOri.State = 0 Then
        MsgBox "Error en la Conexión a la Data Origen", vbCritical, xTitulo
        TxtArchivo.SetFocus
        Exit Sub
    End If

    TxtArchivo2.SetFocus
    
End Sub

Private Sub CmdBusArch2_Click()
    'CommonDialog1.CancelError = True
    'Especificar las extensiones a usar
    CommonDialog1.DefaultExt = "*.mdb"
    CommonDialog1.Filter = "Microsoft Oficce Access (*.mdb)|*.mdb"
    CommonDialog1.ShowOpen
    If Err Then
        'Cancelada la operación de abrir
    Else
        TxtArchivo2.Text = CommonDialog1.FileName
    End If
    '--si la base de datos principal existe
    If ArchivoExiste(TxtArchivo2.Text) = False Then
        MsgBox "No existe la ruta a la Base de Datos Destino", vbCritical, "Mensaje..."
        Exit Sub
    End If
    
    OPEN_CONEX_TMP ConexDest, TxtArchivo2.Text
    If ConexDest.State = 0 Then
        MsgBox "Error en la Conexión a la Data Destino", vbCritical, xTitulo
        TxtArchivo2.Text = ""
        Exit Sub
    End If
    
End Sub
