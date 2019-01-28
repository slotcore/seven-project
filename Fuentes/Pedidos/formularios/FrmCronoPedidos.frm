VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#12.0#0"; "Codejock.Calendar.v12.0.0.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmCronoPedidos 
   Caption         =   "Ventas - Cronograma de Entregas"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   -585
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ilToolBar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Dia"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Semana de Trabajo"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Semana"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Mes"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.ElasticOne EO1 
      Height          =   5265
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   420
      Width           =   9600
      _cx             =   16933
      _cy             =   9287
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
      GridRows        =   2
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmCronoPedidos.frx":0000
      Begin XtremeCalendarControl.CalendarControl CalendarControl 
         Height          =   4890
         Left            =   30
         TabIndex        =   3
         Top             =   345
         Width           =   6915
         _Version        =   786432
         _ExtentX        =   12197
         _ExtentY        =   8625
         _StockProps     =   64
      End
      Begin XtremeCalendarControl.DatePicker wndDatePicker 
         Height          =   4890
         Left            =   6975
         TabIndex        =   2
         Top             =   345
         Width           =   2595
         _Version        =   786432
         _ExtentX        =   4577
         _ExtentY        =   8625
         _StockProps     =   64
         Show3DBorder    =   0
         RowCount        =   2
      End
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         Caption         =   " Calendario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   6915
      End
   End
   Begin ComctlLib.ImageList ilToolBar 
      Left            =   9660
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCronoPedidos.frx":004F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCronoPedidos.frx":05A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCronoPedidos.frx":0AF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCronoPedidos.frx":1045
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCronoPedidos.frx":1597
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCronoPedidos.frx":16A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCronoPedidos.frx":3503
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCronoPedidos.frx":360D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCronoPedidos.frx":3717
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCronoPedidos.frx":3821
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCronoPedidos.frx":39FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCronoPedidos.frx":3BD5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu menu_01 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu_01_01 
         Caption         =   "Ver Entregas"
      End
   End
   Begin VB.Menu menu_02 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu menu_02_01 
         Caption         =   "Ver Esta Entrega"
      End
   End
End
Attribute VB_Name = "FrmCronoPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ContextEvent As CalendarEvent

Private Sub CalendarControl_ContextMenu(ByVal x As Single, ByVal y As Single)
    Dim HitTest As CalendarHitTestInfo
    Set HitTest = CalendarControl.ActiveView.HitTest
    
    If Not HitTest.ViewEvent Is Nothing Then
        Set ContextEvent = HitTest.ViewEvent.Event
        Me.PopupMenu menu_02
        Set ContextEvent = Nothing
    ElseIf (HitTest.HitCode = xtpCalendarHitTestDayViewTimeScale) Then
        MsgBox "Sin menu"
    Else
        Me.PopupMenu menu_01
    End If
End Sub

Private Sub Form_Load()
    Me.Height = 8500
    Me.Width = 12000
    CentrarFrm Me
    'Me.WindowState = 2
    EO1.Left = 0
    EO1.Top = 435
    EO1.Width = Me.Width - 125
    EO1.Height = Me.Height - 835
    
    'ponemos a activo el boton numero 2
    Toolbar.Buttons(2).Value = tbrPressed

    LlenarDatos
    CargarData
End Sub

Sub LlenarDatos()
    Dim Rst As New ADODB.Recordset
    Dim RstEve As New ADODB.Recordset
    Dim RstPro As New ADODB.Recordset
    Dim A As Integer
    Dim B As Integer
    Dim xBody As String
    Dim xDia As Date
    
    xCon.Execute "DELETE * FROM event"
    RST_Busq RstEve, "SELECT *FROM event", xCon
    
    RST_Busq Rst, "SELECT DISTINCT ped_pedido.id, mae_cliente.nombre, mae_cliente.dir, ped_pedido.idtipped, ped_pedidodetent.fchent, " _
        & " ped_pedidodetent.estado FROM mae_cliente RIGHT JOIN (alm_inventario RIGHT JOIN (ped_pedido LEFT JOIN ped_pedidodetent " _
        & " ON ped_pedido.id = ped_pedidodetent.idped) ON alm_inventario.id = ped_pedidodetent.iditem) ON mae_cliente.id = ped_pedido.idcli " _
        & " WHERE (((ped_pedido.idtipped)=2) AND ((ped_pedidodetent.estado)=2))" _
        & " Union " _
        & " SELECT DISTINCT ped_pedido.id, mae_cliente.nombre, mae_cliente.dir, ped_pedido.idtipped, ped_pedidodetent.fchent, ped_pedidodetent.estado " _
        & " FROM (mae_cliente RIGHT JOIN ped_pedido ON mae_cliente.id = ped_pedido.idcli) LEFT JOIN (alm_inventario RIGHT JOIN ped_pedidodetent " _
        & " ON alm_inventario.id = ped_pedidodetent.iditem) ON ped_pedido.id = ped_pedidodetent.idped WHERE (((ped_pedido.idtipped)=1) " _
        & " AND ((ped_pedidodetent.estado)=2))", xCon

    
    Rst.Sort = "fchent, nombre"
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Dim xHorIni, HorFin As String
        
        xDia = Rst("fchent")
        xHorIni = "00:00:00"
            
        For A = 1 To Rst.RecordCount
            Set RstPro = Nothing
            
            RstEve.AddNew
            RstEve("EventID") = A
            
            RstEve("StartDateTime") = Format(Rst("fchent"), "dd/mm/yyyy") & " " & Format(xHorIni, "hh:mm:ss")        ' "8:00:00"
            
            HorFin = ConvertHora(ConvertSeg(Format(xHorIni, "hh:mm:ss")) + ConvertSeg("02:00:00"))
            
            RstEve("EndDateTime") = Format(Rst("fchent"), "dd/mm/yyyy") & " " & Format(HorFin, "hh:mm:ss")
            RstEve("RecurrenceState") = 0
            RstEve("IsAllDayEvent") = 0
            RstEve("Subject") = Trim(Rst("nombre"))
            RstEve("Location") = Trim(Rst("dir"))
            RstEve("RemainderSoundFile") = ""
            RstEve("Created") = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
            RstEve("Modified") = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
            RstEve("BusyStatus") = 2
            RstEve("ImportanceLevel") = 1
            RstEve("LabelID") = 2
            RstEve("RecurrencePatternID") = 0
            RstEve("ScheduleID") = 0
            RstEve("ISRecurrenceExceptionDeleted") = 0
            RstEve("RExceptionStartTimeOrig") = "00:00:00"
            RstEve("RExceptionEndTimeOrig") = "00:00:00"
            RstEve("IsMeeting") = 0
            RstEve("IsPrivate") = 0
            RstEve("IsReminder") = 0
            RstEve("ReminderMinutesBeforeStart") = 15
            RstEve("CustomPropertiesXMLData") = "<Calendar CompactMode='1'/>"
            
            If Rst("idtipped") = 1 Then
                RST_Busq RstPro, "SELECT DISTINCT ped_pedido.id, mae_cliente.nombre, mae_cliente.dir, ped_pedido.idtipped, ped_pedidodetent.fchent, " _
                    & " ped_pedidodetent.estado, alm_inventario.descripcion, ped_pedidodetent.canpro, mae_unidades.abrev " _
                    & " FROM mae_unidades RIGHT JOIN (mae_cliente RIGHT JOIN (alm_inventario RIGHT JOIN (ped_pedido LEFT JOIN ped_pedidodetent " _
                    & " ON ped_pedido.id = ped_pedidodetent.idped) ON alm_inventario.id = ped_pedidodetent.iditem) ON mae_cliente.id = ped_pedido.idcli) " _
                    & " ON mae_unidades.id = ped_pedidodetent.idunimed WHERE (((ped_pedido.id)=" & Rst("id") & ") AND ((ped_pedido.idtipped)=1) " _
                    & " AND ((ped_pedidodetent.fchent)=CDate('" & Rst("fchent") & "')) AND ((ped_pedidodetent.estado)=2))", xCon
            Else
                RST_Busq RstPro, "SELECT DISTINCT ped_pedido.id, mae_cliente.nombre, mae_cliente.dir, ped_pedido.idtipped, ped_pedidodetent.fchent, " _
                    & " ped_pedidodetent.estado, alm_inventario.descripcion, ped_pedidodetent.canpro, mae_unidades.abrev " _
                    & " FROM mae_unidades RIGHT JOIN (mae_cliente RIGHT JOIN (alm_inventario RIGHT JOIN (ped_pedido LEFT JOIN ped_pedidodetent " _
                    & " ON ped_pedido.id = ped_pedidodetent.idped) ON alm_inventario.id = ped_pedidodetent.iditem) ON mae_cliente.id = ped_pedido.idcli) " _
                    & " ON mae_unidades.id = ped_pedidodetent.idunimed WHERE (((ped_pedido.id)=" & Rst("id") & ") AND ((ped_pedido.idtipped)=2) " _
                    & " AND ((ped_pedidodetent.fchent)=CDate('" & Rst("fchent") & "')) AND ((ped_pedidodetent.estado)=2))", xCon
            End If
            
            B = 0
            xBody = ""
            If RstPro.RecordCount <> 0 Then
                RstPro.MoveFirst
                For B = 1 To RstPro.RecordCount
                    xBody = xBody + Trim(RstPro("descripcion")) & " | " & Trim(RstPro("abrev")) & " | " & Format(RstPro("canpro"), "0.00")
                    RstPro.MoveNext
                    
                    If RstPro.EOF = True Then Exit For
                    xBody = xBody & Chr(13)
                Next B
            End If
            'escribimos el contenido del cuerrpo
            RstEve("Body") = xBody
            
            RstEve.Update
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
            
            If xDia = Rst("fchent") Then
                xHorIni = HorFin
            Else
                xHorIni = "00:00:00"
            End If
        Next A
        
    End If
End Sub

Sub CargarData()
    CalendarControl.SetDataProvider xCon.ConnectionString

    CalendarControl.DataProvider.Create

    If Not CalendarControl.DataProvider.Open Then
        CalendarControl.DataProvider.Create
    End If
    CalendarControl.DataProvider.Open

    CalendarControl.ActiveView.ShowDay Date
    CalendarControl.ViewType = xtpCalendarWorkWeekView
    CalendarControl.DayView.ScrollToWorkDayBegin

    wndDatePicker.AttachToCalendar CalendarControl
    CalendarControl.Populate
    CalendarControl.RedrawControl
End Sub

Private Sub Form_Resize()
    EO1.Width = Me.Width - 125
    EO1.Height = Me.Height - 835
End Sub

Private Sub menu_01_01_Click()
    VerEntregas
End Sub

Sub VerEntregas()

End Sub

Private Sub menu_02_01_Click()
    FrmAsunto.Show vbModal
End Sub


Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Dim BeginSelection As Date, EndSelection As Date, AllDay As Boolean

    Select Case (Button.Index)
        Case 1:
            CalendarControl.ViewType = xtpCalendarDayView
        Case 2:
            CalendarControl.ViewType = xtpCalendarWorkWeekView
        Case 3:
            CalendarControl.ViewType = xtpCalendarWeekView
        Case 4:
            CalendarControl.ViewType = xtpCalendarMonthView
        
'        Case 6:
'            CalendarControl.ActiveView.Cut
'        Case 7:
'            CalendarControl.ActiveView.Copy
'        Case 8:
'            CalendarControl.ActiveView.Paste
'
'        Case 12:
'            'mnuOpenDataProvider_Click
'
'        Case 14:
'            'mnuPageSetup_Click
'        Case 15:
'            'mnuPrintPreview_Click
'        Case 16:
'            'mnuPrintCalendar_Click
    End Select
    
    UpdateToolbar
End Sub

Private Sub UpdateToolbar()
    
    'Toolbar.Buttons(6).Enabled = CalendarControl.ActiveView.CanCut
    'Toolbar.Buttons(7).Enabled = CalendarControl.ActiveView.CanCopy
    'Toolbar.Buttons(8).Enabled = CalendarControl.ActiveView.CanPaste
    
End Sub
