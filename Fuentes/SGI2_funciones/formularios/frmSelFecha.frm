VERSION 5.00
Begin VB.Form frmSelFecha 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fecha"
   ClientHeight    =   810
   ClientLeft      =   4545
   ClientTop       =   2460
   ClientWidth     =   2430
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
   ScaleHeight     =   810
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDia 
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
      Left            =   45
      MaxLength       =   2
      TabIndex        =   1
      Top             =   465
      Width           =   375
   End
   Begin VB.ComboBox cbMes 
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
      ItemData        =   "frmSelFecha.frx":0000
      Left            =   450
      List            =   "frmSelFecha.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   465
      Width           =   1320
   End
   Begin VB.TextBox txtAno 
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
      Left            =   1785
      MaxLength       =   4
      TabIndex        =   5
      Top             =   465
      Width           =   615
   End
   Begin VB.Label lblCol 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "lblCol"
      Height          =   180
      Left            =   2640
      TabIndex        =   8
      Top             =   1245
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblRow 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "lblRow"
      Height          =   180
      Left            =   2640
      TabIndex        =   7
      Top             =   990
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
      Left            =   15
      TabIndex        =   6
      Top             =   15
      Width           =   2370
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   450
      TabIndex        =   2
      Top             =   270
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1785
      TabIndex        =   4
      Top             =   270
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   270
      Width           =   225
   End
End
Attribute VB_Name = "frmSelFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dt As Variant
Dim SeEjecuto As Boolean

Dim xObject As Object
Dim SGI_JC As New SGI2_funciones.JC_Varios


Public Sub pRecibeLink(xGrid As Object)
    '---------------------------
    Set xObject = xGrid
    DoEvents
End Sub

Private Sub pActualizarFecha()
    If Not Enabled Then Exit Sub
    If Not IsDate(txtDia.Text & "/" & Format(cbMes.ListIndex + 1, "00") & "/" & txtAno.Text) Then Beep: Exit Sub
    dt = txtDia.Text & "/" & Format(CStr(cbMes.ListIndex + 1), "00") & "/" & txtAno.Text
    lblDate = Format(dt, "Long Date")
    Tag = dt
End Sub

Private Sub cbMes_GotFocus()
    SendKeys "{F4}"
End Sub

Private Sub cbMes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pActualizarFecha
        txtAno.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    dt = Tag
    Enabled = False
    If dt = "" Then dt = Date
    If IsDate(dt) = False Then dt = Date
    Tag = dt
    lblDate = Format(dt, "Long Date")
    
    txtAno.Text = Format(Year(dt), "0000")
    txtDia.Text = Format(Day(dt), "00")
    cbMes.ListIndex = Month(dt) - 1
    
    Enabled = True
    txtDia.SetFocus
    SeEjecuto = True
    
End Sub

Private Sub Form_Deactivate()
    
    On Error Resume Next
    
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
    SGI_JC.Llenar_Mes cbMes
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SGI_JC = Nothing
End Sub

Private Sub txtAno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Form_Deactivate
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If SGI_JC.validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txtDia_Change()
    pActualizarFecha
End Sub


Private Sub txtDia_GotFocus()
    txtDia.SelStart = 0
    txtDia.SelLength = 30000
    txtDia.MaxLength = 2
End Sub


Private Sub cbMes_Change()
    pActualizarFecha
End Sub

Private Sub cbMes_Click()
    pActualizarFecha
End Sub

Private Sub txtAno_Change()
    pActualizarFecha
End Sub


Private Sub txtAno_GotFocus()
    txtAno.SelStart = 0
    txtAno.SelLength = 30000
    txtAno.MaxLength = 4
End Sub


Private Sub txtDia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cbMes.SetFocus
End Sub

Private Sub txtDia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If SGI_JC.validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub
