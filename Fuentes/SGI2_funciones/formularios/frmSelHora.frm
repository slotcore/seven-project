VERSION 5.00
Begin VB.Form frmSelHora 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hora"
   ClientHeight    =   825
   ClientLeft      =   4545
   ClientTop       =   2460
   ClientWidth     =   1530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   1530
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtMin 
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
      Height          =   315
      Left            =   540
      MaxLength       =   2
      TabIndex        =   1
      ToolTipText     =   "Ingrese Valores desde 00 -> 60"
      Top             =   465
      Width           =   420
   End
   Begin VB.TextBox txtHora 
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
      Height          =   315
      Left            =   30
      MaxLength       =   2
      TabIndex        =   0
      ToolTipText     =   "Ingrese Valores desde 00 -> 23"
      Top             =   465
      Width           =   420
   End
   Begin VB.TextBox txtSeg 
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
      Height          =   315
      Left            =   1065
      MaxLength       =   2
      TabIndex        =   2
      ToolTipText     =   "Ingrese Valores desde 00 -> 60"
      Top             =   465
      Width           =   420
   End
   Begin VB.Shape sha 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   45
      Index           =   3
      Left            =   990
      Top             =   675
      Width           =   45
   End
   Begin VB.Shape sha 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   45
      Index           =   2
      Left            =   990
      Top             =   555
      Width           =   45
   End
   Begin VB.Shape sha 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   45
      Index           =   1
      Left            =   480
      Top             =   675
      Width           =   45
   End
   Begin VB.Shape sha 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   45
      Index           =   0
      Left            =   480
      Top             =   555
      Width           =   45
   End
   Begin VB.Label lblCol 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "lblCol"
      Height          =   180
      Left            =   705
      TabIndex        =   8
      Top             =   915
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblRow 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "lblRow"
      Height          =   180
      Left            =   60
      TabIndex        =   7
      Top             =   855
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblDate 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   15
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   540
      TabIndex        =   4
      Top             =   270
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seg."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   1065
      TabIndex        =   5
      Top             =   270
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   30
      TabIndex        =   3
      Top             =   270
      Width           =   345
   End
End
Attribute VB_Name = "frmSelHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim SeEjecuto As Boolean
Dim dt As Variant

Dim xObject As Object
Dim SGI_JC As New SGI2_funciones.JC_Varios

Public Sub pRecibeLink(xGrid As Object)
    '---------------------------
    Set xObject = xGrid
    DoEvents
End Sub

Private Sub pActualizarHora()
    If Not Enabled Then Exit Sub
    If Not IsDate(txtHora.Text & ":" & TxtMin.Text & ":" & txtSeg.Text) Then Beep: Exit Sub
    dt = Format(txtHora.Text, "00") & ":" & Format(TxtMin.Text, "00") & ":" & Format(txtSeg.Text, "00")
    lblDate = Format(dt, "hh:mm:ss am/pm")
    Tag = lblDate
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    dt = Tag
    Enabled = False
    If dt = "" Then dt = Time()
    If IsDate(dt) = False Then dt = Time()
    Tag = dt
    lblDate = Format(dt, "hh:mm:ss AM/PM")
    
    txtHora.Text = Format(Hour(dt), "00")
    TxtMin.Text = Format(Minute(dt), "00")
    txtSeg.Text = Format(Second(dt), "00")
    
    Enabled = True
    txtHora.SetFocus
    SeEjecuto = True
End Sub

Private Sub Form_Deactivate()
    
    On Error Resume Next
    
    ' update grid value
    If IsDate(Tag) Then
        If TypeName(xObject) = "VSFlexGrid" Then
            xObject.Cell(flexcpText, lblRow.Caption, lblCol.Caption) = Tag
        ElseIf TypeName(xObject) = "TextBox" Then
            xObject.Text = Tag
        ElseIf TypeName(xObject) = "DTPicker" Then
            xObject.Value = Tag
        End If
    End If
    Err.Clear
    
    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Tag = "": Unload Me
End Sub

Private Sub Form_Load()
    SeEjecuto = False
End Sub

Private Sub txtHora_Change()
    If NulosN(txtHora.Text) > 23 Or NulosN(txtHora.Text) < 0 Then txtHora.Text = ""
    pActualizarHora
End Sub

Private Sub txtHora_GotFocus()
    txtHora.SelStart = 0
    txtHora.SelLength = 30000
    txtHora.MaxLength = 2
End Sub


Private Sub txtHora_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TxtMin.SetFocus
End Sub

Private Sub txtHora_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If SGI_JC.validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txtHora_Validate(Cancel As Boolean)
    If NulosN(txtHora.Text) > 23 Or NulosN(txtHora.Text) < 0 Then txtHora.Text = 0
    txtHora.Text = Format(txtHora.Text, "00")
End Sub

Private Sub TxtMin_Change()
    If NulosN(TxtMin.Text) > 60 Or NulosN(TxtMin.Text) < 0 Then TxtMin.Text = ""
    pActualizarHora
End Sub

Private Sub TxtMin_GotFocus()
    TxtMin.SelStart = 0
    TxtMin.SelLength = 30000
    TxtMin.MaxLength = 4
End Sub

Private Sub TxtMin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtSeg.SetFocus
End Sub

Private Sub TxtMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If SGI_JC.validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtMin_Validate(Cancel As Boolean)
    If NulosN(TxtMin.Text) > 60 Or NulosN(TxtMin.Text) < 0 Then TxtMin.Text = 0
    TxtMin.Text = Format(TxtMin.Text, "00")
End Sub

Private Sub txtSeg_Change()
    If NulosN(txtSeg.Text) > 60 Or NulosN(txtSeg.Text) < 0 Then txtSeg.Text = ""
    pActualizarHora
End Sub


Private Sub txtSeg_GotFocus()
    txtSeg.SelStart = 0
    txtSeg.SelLength = 30000
    txtSeg.MaxLength = 4
End Sub


Private Sub txtSeg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Form_Deactivate
End Sub

Private Sub txtSeg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If SGI_JC.validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txtSeg_Validate(Cancel As Boolean)
    If NulosN(txtSeg.Text) > 60 Or NulosN(txtSeg.Text) < 0 Then txtSeg.Text = 0
    txtSeg.Text = Format(txtSeg.Text, "00")
End Sub
