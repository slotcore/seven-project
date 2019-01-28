VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FrmNavegador 
   Caption         =   "Navegador Web "
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   Icon            =   "FrmNavegador.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1185
      TabIndex        =   10
      Top             =   5265
      Visible         =   0   'False
      Width           =   7230
   End
   Begin SizerOneLibCtl.ElasticOne Eo1 
      Height          =   4890
      Left            =   165
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   45
      Width           =   8415
      _cx             =   14843
      _cy             =   8625
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
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmNavegador.frx":0442
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   8355
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            Height          =   600
            Left            =   30
            Picture         =   "FrmNavegador.frx":0484
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            Height          =   600
            Left            =   870
            Picture         =   "FrmNavegador.frx":1DC6
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Appearance      =   0  'Flat
            Height          =   600
            Left            =   1710
            Picture         =   "FrmNavegador.frx":3704
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            Appearance      =   0  'Flat
            Height          =   600
            Left            =   2550
            Picture         =   "FrmNavegador.frx":5046
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton Command5 
            Appearance      =   0  'Flat
            Height          =   600
            Left            =   3390
            Picture         =   "FrmNavegador.frx":6A28
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   4485
            TabIndex        =   9
            Top             =   15
            Visible         =   0   'False
            Width           =   3810
         End
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   4200
         Left            =   30
         TabIndex        =   2
         Top             =   660
         Width           =   8355
         ExtentX         =   14737
         ExtentY         =   7408
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dirección:"
      Height          =   195
      Left            =   285
      TabIndex        =   11
      Top             =   5295
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "FrmNavegador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean

Private Sub Command1_Click()
    ' Atrás
    On Error Resume Next
    WebBrowser1.GoBack
    Text1.Text = WebBrowser1.LocationURL ' depurar
End Sub

Private Sub Command2_Click()
    ' Adelante
    On Error Resume Next
    WebBrowser1.GoForward
    Text1.Text = WebBrowser1.LocationURL ' depurar
End Sub

Private Sub Command3_Click()
'    ' Inicio al disco C
'    Text1.Text = "file://C:"
'    WebBrowser1.Navigate Text1.Text
End Sub

Private Sub Command4_Click()
    ' Actualizar
    WebBrowser1.Refresh
End Sub

Private Sub Command5_Click()
    ' Detener
    WebBrowser1.Stop
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        WebBrowser1.Navigate Text1.Text
    End If
End Sub

Private Sub Form_Load()
    Frame1.BackColor = &H8000000F
    Eo1.Left = 0
    Eo1.Top = 0
    SeEjecuto = False
End Sub

Private Sub Form_Resize()
    'Redimensionamos los controles
    Eo1.Width = Me.Width - 120
    Eo1.Height = Me.Height - 400
'    On Error Resume Next
'    Move (Screen.Width - Width) \ 29, (Screen.Height - Height) \ 29
'
'    Frame1.Left = (Me.Width - Frame1.Width) / 2
'    Frame1.Top = Me.Height - (Frame1.Height + 480)
'
'    WebBrowser1.Width = Me.Width - 360
'    WebBrowser1.Height = Frame1.Top - 180

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text1.Text <> "" Then WebBrowser1.Navigate Text1.Text
    ' KeyAscii= 13 Equivale a la tecla Enter
    ' <> "" Equivale a: no-vacío
End Sub

Private Sub WebBrowser1_DownloadBegin()
    Me.Caption = "Navegador Web: " & WebBrowser1.LocationName
    App.Title = "Navegador Web: " & WebBrowser1.LocationName
    Label2.Caption = "Cargando Página..."
End Sub

Private Sub WebBrowser1_DownloadComplete()
    Me.Caption = "Navegador Web: " & WebBrowser1.LocationName
    App.Title = "Navegador Web: " & WebBrowser1.LocationName
    Label2.Caption = "Listo"
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    'Mostramos la url que se está cargando en el combo
    'Agregamos la url al combo
    Text1.Text = WebBrowser1.LocationURL
    Text2.Text = WebBrowser1.LocationURL
    'Mostramos en el la barra de titulo del formulario el title _
     de la página con la propiedad LocationName
    
    Me.Caption = "Navegador Web: " & WebBrowser1.LocationName
    App.Title = "Navegador Web: " & WebBrowser1.LocationName
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Dim frm_web As FrmNavegador
    'Nueva instancia del formulario
    Set frm_web = New FrmNavegador
    
    'Posición
    With frm_web
        .WindowState = vbNormal
        .Top = frm_web.Top + 1200
        .Left = frm_web.Left + 1200
    'Destino
    Set ppDisp = frm_web.WebBrowser1.Object
    
    'Mostramos el nuevo form creado
    frm_web.Show
    End With
End Sub
