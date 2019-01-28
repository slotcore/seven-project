VERSION 5.00
Begin VB.Form FrmManConceptoFormulaProbar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Concepto - Probar Fórmula"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3660
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   -30
      Width           =   4950
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "txt(0)"
         Top             =   210
         Width           =   3855
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   870
         TabIndex        =   6
         Text            =   "txt(1)"
         Top             =   570
         Width           =   2145
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   5
         Top             =   960
         Width           =   4650
      End
      Begin VB.CommandButton cmdAddVar 
         Caption         =   "Modificar Valor"
         Height          =   315
         Left            =   3195
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Modificar Valor"
         Top             =   570
         Width           =   1545
      End
      Begin VB.Label lbl 
         Caption         =   "Nombre:"
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   9
         Top             =   255
         Width           =   645
      End
      Begin VB.Label lbl 
         Caption         =   "Valor:"
         Height          =   255
         Index           =   1
         Left            =   165
         TabIndex        =   8
         Top             =   615
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "Probar fórmula"
      Height          =   345
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Probar fórmula"
      Top             =   4665
      Width           =   1515
   End
   Begin VB.TextBox txt_formula 
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
      Height          =   720
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "FrmManConceptoFormulaProbar.frx":0000
      Top             =   3885
      Width           =   5010
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   2
      Left            =   885
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Txt(2)"
      Top             =   4665
      Width           =   2550
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Expresión a evaluar:"
      Height          =   255
      Index           =   2
      Left            =   45
      TabIndex        =   11
      Top             =   3660
      Width           =   1455
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado"
      Height          =   195
      Index           =   3
      Left            =   75
      TabIndex        =   10
      Top             =   4770
      Width           =   720
   End
End
Attribute VB_Name = "FrmManConceptoFormulaProbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim i As Integer

Const cTxtNombreVar     As Long = 0
Const cTxtValorVar      As Long = 1

Private Sub cmdAddVar_Click()
    Dim sTmp        As String
    Dim i           As Long
    Dim sVar        As String
    Dim Hallado     As Long
    
    If txt(cTxtNombreVar) = "" Then Exit Sub
    
        sVar = Trim$(txt(cTxtNombreVar)) & " = "
        sTmp = sVar & txt(cTxtValorVar)
        ' Comprobar si está esa variable
        Hallado = -1
        With List1
            For i = 0 To .ListCount - 1
                If InStr(.list(i), sVar) Then
                    Hallado = i
                    Exit For
                End If
            Next
            If Hallado = -1 Then
                .AddItem sTmp
            Else
                .list(Hallado) = sTmp
            End If
        End With
   
End Sub
Private Sub cmdCalcular_Click()
    Dim i           As Long
    Dim J           As Long
    Dim sFormula    As String
    Dim TFormulas As CProcessor
    Set TFormulas = New CProcessor
    On Error GoTo error
    If txt(2).Tag = "" Then
        Set TFormulas = Nothing
        Exit Sub
    End If
    
    TFormulas.BaseCalculation = 1
    With List1
        For i = 0 To .ListCount - 1
            J = InStr(.list(i), "=")
            TFormulas.DeclareConstant(Left$(.list(i), J - 1)) = Mid$(.list(i), J + 1)
        Next
    End With
    
     sFormula = Trim(txt(2).Tag)
    
     txt(2).Text = " " & Format(TFormulas.Calculate(sFormula), FORMAT_MONTO)
     Set TFormulas = Nothing
     Exit Sub
error:
    Set TFormulas = Nothing
    MsgBox Err.Description & vbCr & Err.Source, vbCritical, "Error..."
     
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    LimpiaText txt
    With FrmManConceptoFormula
        txt_formula.Text = .txt_formula.Text
        txt(2).Tag = .txt_formula.Text
        For i = 0 To .List1.ListCount - 1
            List1.AddItem .List1.list(i) & " = "
        Next
    End With
End Sub

Private Sub List1_Click()
    Dim sTmp    As String
    Dim i       As Integer

    With List1
        If .ListIndex > -1 Then
            sTmp = .list(.ListIndex)
            i = InStr(sTmp, "=")
            If i Then
                txt(cTxtNombreVar) = Trim$(Left$(sTmp, i - 1))
                txt(cTxtValorVar) = Trim$(Mid$(sTmp, i + 1))
            End If
        End If
    End With
End Sub

Private Sub txt_Change(Index As Integer)
    If Index = 1 Then
        If txt(Index).Text = "" Then Exit Sub
        If IsNumeric(txt(Index).Text) = False Then
            MsgBox "No es valor numérico para este campo", vbCritical
            txt(Index).Text = ""
        End If
    End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdAddVar_Click
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If validar_numero(KeyAscii) = False And Chr(KeyAscii) <> "." Then
           MsgBox "No es valor numérico para este campo", vbCritical
           KeyAscii = 0
        ElseIf KeyAscii = 13 Then
           cmdAddVar_Click
        End If
    End If
End Sub

