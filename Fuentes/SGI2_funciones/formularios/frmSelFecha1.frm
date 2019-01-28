VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSelFecha1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccionar Fecha"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView mv 
      Height          =   2370
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   17104897
      CurrentDate     =   39711
   End
   Begin VB.Label lblRow 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "lblRow"
      Height          =   180
      Left            =   0
      TabIndex        =   2
      Top             =   2490
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblCol 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "lblCol"
      Height          =   180
      Left            =   0
      TabIndex        =   1
      Top             =   2745
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmSelFecha1"
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
    mv.Value = Date
    Set xObject = xGrid
    DoEvents
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    dt = Tag
    Enabled = False
    If dt = "" Then dt = Date
    If IsDate(dt) = False Then dt = Date
    Tag = dt
    
    mv.Value = dt
    
    Enabled = True
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SGI_JC = Nothing
End Sub

Private Sub mv_DateDblClick(ByVal DateDblClicked As Date)
    If IsDate(mv.Day & "/" & mv.Month & "/" & mv.Year) = True Then
        Tag = CDate(mv.Day & "/" & mv.Month & "/" & mv.Year)
    End If
    
    Form_Deactivate
    Unload Me
End Sub

Private Sub mv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If IsDate(mv.Day & "/" & mv.Month & "/" & mv.Year) = True Then
            Tag = CDate(mv.Day & "/" & mv.Month & "/" & mv.Year)
        End If
        Form_Deactivate
        Unload Me
    ElseIf KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub
