VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCronoLineaTiempo 
   Caption         =   "Produccion - Linea de Tiempo"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15600
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   15600
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.ElasticOne Eo1 
      Height          =   5760
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   15480
      _cx             =   27305
      _cy             =   10160
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   8
      BorderWidth     =   2
      ChildSpacing    =   2
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   3
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmCronoLineaTiempo.frx":0000
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   30
         TabIndex        =   6
         Top             =   5070
         Width           =   15420
         Begin VB.Label LblProducto 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProducto"
            Height          =   300
            Left            =   1695
            TabIndex        =   10
            Top             =   15
            Width           =   9900
         End
         Begin VB.Label Label7 
            Caption         =   "MP / Producto =>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   1560
         End
         Begin VB.Label LblTarea 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTarea"
            Height          =   300
            Left            =   1695
            TabIndex        =   8
            Top             =   330
            Width           =   9900
         End
         Begin VB.Label Label4 
            Caption         =   "Tarea =>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   7
            Top             =   375
            Width           =   1560
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   15420
         Begin VB.ComboBox CboFecha 
            Height          =   315
            Left            =   1410
            TabIndex        =   5
            Text            =   "CboFecha"
            Top             =   30
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Dia de Trabajo"
            Height          =   195
            Left            =   135
            TabIndex        =   4
            Top             =   90
            Width           =   1050
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg1 
         Height          =   4605
         Left            =   30
         TabIndex        =   1
         Top             =   435
         Width           =   15420
         _cx             =   27199
         _cy             =   8123
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
         BackColor       =   14220798
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14220798
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   8
         FixedRows       =   3
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCronoLineaTiempo.frx":0050
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   30
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
            Picture         =   "FrmCronoLineaTiempo.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoLineaTiempo.frx":068E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoLineaTiempo.frx":0812
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoLineaTiempo.frx":0C66
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoLineaTiempo.frx":0D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoLineaTiempo.frx":12C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoLineaTiempo.frx":1806
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoLineaTiempo.frx":191A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoLineaTiempo.frx":1A2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoLineaTiempo.frx":1E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoLineaTiempo.frx":1FEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15600
      _ExtentX        =   27517
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Recetas del producto"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Productos "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmCronoLineaTiempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean

Sub PreparaGrid()
    'Grid.WordWrap = True
    Fg1.Rows = 3
    Fg1.Cols = 8
    Fg1.RowHeight(1) = 500
    Fg1.RowHeight(2) = 100
    
    Fg1.WordWrap = True   ' para poner en multilineas los caption de cada columna
    
    GRID_COMBINAR Fg1, 0, 1, 2, 1, "Materia Prima", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 0, 2, 2, 2, "Producto", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 0, 3, 2, 3, "Tarea", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 0, 4, 2, 4, "Cantidad", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 0, 5, 2, 5, "Nº Per.", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 0, 6, 2, 6, "Total Horas", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 0, 7, 2, 7, "Hora Inicio", flexAlignCenterCenter, False, , , &H8000000F, False
    
    Fg1.ColWidth(5) = 400
    Fg1.ColWidth(6) = 600
    Fg1.ColWidth(7) = 600
End Sub

Sub Muestradatos()
    If CboFecha.Text = "" Then
        MsgBox "No ha especificado la fecha de consulta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        CboFecha.SetFocus
        Exit Sub
    End If
    
    Fg1.Rows = 3
    Fg1.Cols = 8
    
    ' xNumTurnos = almacenara el numero de turnos por dia
    ' xHorTurnos = almacenara el numero de horas por turno
    Dim xNumTurnos, xHorTurnos, xHoras, xHora, xNumCols As Integer
    Dim xTiempoMax, xHorIni, xHorIniTur, xHorFinTur As String
    Dim xCad As String
    Dim xNumDias As Double
    Dim RstLis As New ADODB.Recordset
    Dim RstLis2 As New ADODB.Recordset
    
    xCad = "SELECT pro_cronogramatarea.fchini, alm_inventario.descripcion AS matpri, alm_inventario_1.descripcion AS despro, pro_cronogramatarea.horpro, " _
        & " pro_cronogramatarea.orden, pro_tareas.descripcion AS destar, pro_cronogramatarea.fchpro, pro_cronogramatarea.iditem, pro_cronogramatarea.idpro, " _
        & " pro_cronogramatarea.idtar, pro_cronogramatarea.numper, 0 AS canpro, pro_cronogramatarea.horinitar, pro_cronogramatarea.aplpor, " _
        & " IIf([pro_cronogramatarea].[aplpor]=0,[pro_cronogramatarea].[factor]*[pro_cronogramatarea].[cantidad],([pro_cronogramatarea].[factor]*[pro_cronogramatarea].[cantidad])*([pro_cronogramatarea].[aplpor]/100)) AS tiempoesttotal, " _
        & " [tiempoesttotal]/[numper] AS tiempoesttotalper, Format(Int(([tiempoesttotalper]*60)/60),'00') & ':' & Format(([tiempoesttotalper]*60) Mod 60,'00') AS tiempoesttotalhorper, " _
        & " IIf([pro_cronogramatarea].[aplpor]=0,[pro_cronogramatarea].[cantidad],([pro_cronogramatarea].[cantidad]*([pro_cronogramatarea].[aplpor]/100))) AS cantotpro, " _
        & " pro_cronogramatarea.id FROM ((pro_cronogramatarea LEFT JOIN alm_inventario ON pro_cronogramatarea.iditem = alm_inventario.id) LEFT JOIN alm_inventario AS alm_inventario_1 " _
        & " ON pro_cronogramatarea.idpro = alm_inventario_1.id) INNER JOIN pro_tareas ON pro_cronogramatarea.idtar = pro_tareas.id " _
        & " WHERE (((pro_cronogramatarea.fchini)=CDate('" & CboFecha.Text & "'))) " _
        & " ORDER BY pro_cronogramatarea.fchini, alm_inventario.descripcion, alm_inventario_1.descripcion, pro_cronogramatarea.horpro, pro_cronogramatarea.orden"

    RST_Busq RstLis, xCad, xCon
    RST_Busq RstLis2, xCad, xCon
    
    xNumTurnos = 1
    xHorTurnos = 12
    
    xHorIniTur = "07:00"
    xHorFinTur = "19:00"
    
    If RstLis.RecordCount <> 0 Then
        RstLis2.Sort = "tiempoesttotalhorper"
        RstLis2.MoveLast
        
        ' determinamos el tiempo maximo y su respectiva hora de inicio
        xTiempoMax = RstLis2("tiempoesttotalhorper")
        xHorIni = RstLis2("horinitar")
        xHora = Val(Mid(xTiempoMax, 1, 2))
        If xNumTurnos = 1 Then
            xNumDias = xHora / xHorTurnos
            If (Int(xNumDias) - xNumDias) <> 0 Then
                xNumDias = Int(xNumDias) + 2
            Else
                xNumDias = 1
            End If
            xNumCols = ((xNumDias * xHorTurnos) * 2)
        End If

        
        If xNumTurnos = 2 Then
            ' ESTO ESTA POR HACER
        End If
        
        ' xIntervalo = variable para controlar el intervalo entre las horas '7:00 7:30 8:00 8:30 etc
        Dim A, B, C, xNumHor, xIntervalo As Integer
        Dim xNumHoras, xSaldo As Double
        Dim xDiaAct, xHoraAct As Date
        
        xDiaAct = CDate(CboFecha.Text)
        xNumHor = 0
        xIntervalo = 0
        xHoraAct = CDate("7:00")
        
        For A = 1 To xNumCols
            xNumHor = xNumHor + 1
            xHoraAct = xHoraAct + CDate("00:30")
            xIntervalo = xIntervalo + 1
            Fg1.Cols = Fg1.Cols + 1
            Fg1.ColWidth(Fg1.Cols - 1) = 125
            If xIntervalo = 2 Then
                GRID_COMBINAR Fg1, 1, Fg1.Cols - 2, 1, Fg1.Cols - 1, Format(xHoraAct - CDate("1:00"), "HH:MM"), flexAlignCenterCenter, True, , , &H8000000F, False
                xIntervalo = 0
                Fg1.FixedAlignment(8) = flexAlignLeftBottom
            End If
            If xNumHor = 24 Then
                GRID_COMBINAR Fg1, 0, Fg1.Cols - 24, 0, Fg1.Cols - 1, Format(xDiaAct, "dd/mm/yy"), flexAlignCenterCenter, True, , , &H8000000F, False
                xNumHor = 0
                xDiaAct = xDiaAct + 1
                xHoraAct = CDate("7:00")
            End If
        Next A
        
        'determinamos el numero de dias a mostrar en funcion a la horas maxima de la tareas
        RstLis.MoveFirst
        For A = 1 To RstLis.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstLis("matpri"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = RstLis("despro")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = RstLis("destar")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(RstLis("cantotpro"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(RstLis("numper"), "00")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = RstLis("tiempoesttotalhorper")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(RstLis("horinitar"), "HH:MM")
            
            xHoraAct = CDate("7:00")
            For B = 1 To xHorTurnos * 2
                If Format(CDate(xHoraAct), "HH:MM") >= Format(CDate(RstLis("horinitar")), "HH:MM") Then
                    xNumHoras = 0
                    C = 0
                    xNumHoras = HoraDecimal(RstLis("tiempoesttotalhorper"))
                    
                    xSaldo = xNumHoras - Int(xNumHoras)
                    xNumHoras = (Int(xNumHoras) * 2)         ' HALLAMOS EL TOTAL DE HORAS Y LO MULTIPLICAMOS POR 2, PORQUE EL INTERVALO DE LA LINEA DE TIEMPO ES DE 30 MINUTOS
                    
                    If xSaldo <= 0.5 Then
                        xNumHoras = xNumHoras + 1
                    Else
                        xNumHoras = xNumHoras + 2
                    End If
                    
                    For C = 1 To xNumHoras
                        GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, (B + 6) + C, Fg1.Rows - 1, (B + 6) + C, &H40&, flexFillRepeat
                    Next C
                    Exit For
                Else
                    Fg1.TextMatrix(Fg1.Rows - 1, B + 7) = ""
                End If
                
                xHoraAct = xHoraAct + CDate("00:30")
                'End If
            Next B
            
            
            
            RstLis.MoveNext
            If RstLis.EOF = True Then Exit For
        Next A
    End If
    
    Fg1.Select 1, 1
    LblProducto.Caption = Fg1.TextMatrix(Fg1.Row, 2)
    LblTarea.Caption = Fg1.TextMatrix(Fg1.Row, 3)
End Sub

'Private Sub CboFecha_KeyDown(KeyCode As Integer, Shift As Integer)
'    CboFecha.Text = ""
'End Sub

Private Sub CboFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        CboFecha.Text = ""
    End If
End Sub

Private Sub CboFecha_KeyUp(KeyCode As Integer, Shift As Integer)
    CboFecha.Text = ""
End Sub

Private Sub Fg1_RowColChange()
    LblProducto.Caption = NulosC(Fg1.TextMatrix(Fg1.Row, 2))
    LblTarea.Caption = NulosC(Fg1.TextMatrix(Fg1.Row, 3))
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Dim Rst As New ADODB.Recordset
        Dim A As Integer
        
        RST_Busq Rst, "SELECT DISTINCT pro_cronogramatarea.fchini FROM pro_cronogramatarea", xCon
        
        CboFecha.Text = ""
        'CboFecha.Locked = True
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                CboFecha.AddItem Format(Rst("fchini"), "dd/mm/yy")
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    PreparaGrid
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Eo1.Left = 10
    Eo1.Top = 375
End Sub

Private Sub Form_Resize()
    Dim TopEO As Integer
    
    TopEO = 400
    
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.Height <= (TopEO + 2375) Then
        Me.Height = (TopEO + 2375)
    Else
        'EO.Height = (Me.Height - (TopEO + 400))
    End If
    
    Eo1.Height = (Me.Height - 790)
    Eo1.Width = (Me.Width - 125)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        Muestradatos
    End If
    
    If Button.Index = 5 Then
        Unload Me
    End If
End Sub
