VERSION 5.00
Begin VB.Form FrmGeneraKey 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar Llave"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   2730
      TabIndex        =   7
      Top             =   990
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   345
      Left            =   3945
      TabIndex        =   4
      Top             =   2400
      Width           =   2130
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Halla Numero de Serie"
      Height          =   450
      Left            =   2730
      TabIndex        =   3
      Top             =   300
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   165
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1125
      Width           =   2325
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   690
      Left            =   900
      TabIndex        =   1
      Top             =   2295
      Width           =   2130
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   165
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   390
      Width           =   2325
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Numero de Serie"
      Height          =   195
      Left            =   165
      TabIndex        =   6
      Top             =   900
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Clave"
      Height          =   195
      Left            =   165
      TabIndex        =   5
      Top             =   165
      Width           =   405
   End
End
Attribute VB_Name = "FrmGeneraKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub Command1_Click()
'    'Text1.Text = CadOrigen
'End Sub

Private Sub Command2_Click()
    Text2.Text = HallaContra(Text1.Text)
End Sub

'Private Sub Command3_Click()
'    If (EvaluaClave(Text1.Text, Text2.Text) = True) Then
'        MsgBox ("la contraseña es correcta")
'    Else
'        MsgBox ("maldito piratero te voy a destruir")
'    End If
'End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Text1.Text = ""
    Text2.Text = ""
    'Text1.Text = CadOrigen
End Sub

'Function EvaluaClave(clave As String, password As String) As Boolean
'    Dim Num1 As String
'    Dim Num2 As String
'    Dim Num3 As String
'    Dim Num4 As String
'    Dim xNumero1 As Double
'    Dim xNumero2 As Double
'
'    Dim Valor As Integer
'    Valor = Val(Mid(Trim(password), Len(password) - 1, 2))
'
'    Num1 = Trim(Mid(clave, 1, 4))
'    Num2 = Trim(Mid(clave, 6, 4))
'    Num3 = Trim(Mid(clave, 11, 4))
'    Num4 = Trim(Mid(clave, 16, 5))
'
'    xNumero1 = Val(Num1) + Val(Num2)
'    xNumero2 = Val(Num3) + Val(Num4)
'
'    xNumero1 = xNumero1 * Valor
'    xNumero2 = xNumero2 * Valor
'
'    Dim xCad1 As String
'    xCad1 = Trim(Str(xNumero1)) + Trim(Str(xNumero2)) + Trim(Str(Valor))
'    xCad1 = Mid(xCad1, 1, 4) + "-" + Trim(Mid(xCad1, 5, 20))
'
'    If xCad1 = password Then
'        EvaluaClave = True
'    Else
'        EvaluaClave = False
'    End If
'End Function

Function HallaContra(Cadena As String) As String
    Dim Num1 As String
    Dim Num2 As String
    Dim Num3 As String
    Dim Num4 As String
    
    ' 1110-0111-0011-100
     '1234567890123456789
    
    Num1 = Trim(Mid(Cadena, 1, 4))
    Num2 = Trim(Mid(Cadena, 6, 4))
    Num3 = Trim(Mid(Cadena, 11, 4))
    Num4 = Trim(Mid(Cadena, 16, 5))
    
    
    Dim xNumero1 As Double
    Dim xNumero2 As Double
    Dim A As Double
    
    '(1110 +0111) (0011 +100)
    xNumero1 = Val(Num1) + Val(Num2)
    xNumero2 = Val(Num3) + Val(Num4)
    
    Dim MiValor
    Randomize 10                       ' Inicializa el generador de números aleatorios.
    
    While MiValor < 10
        MiValor = Int((99 * Rnd) + 1)   ' Genera valores aleatorios entre 1 y 10.
    Wend
    
''    'cambiar el 2 para generar nueva clave
'    xNumero1 = xNumero1 * (MiValor + 2)
'    xNumero2 = xNumero2 * (MiValor + 2)

    'cambiar el 4 para generar nueva clave ' esto se modifico el 02-03-2011
    xNumero1 = xNumero1 * (MiValor + 4)
    xNumero2 = xNumero2 * (MiValor + 4)
    
    HallaContra = Trim(Str(xNumero1)) + Trim(Str(xNumero2)) + Trim(Str(MiValor))
    HallaContra = Mid(HallaContra, 1, 4) + "-" + Trim(Mid(HallaContra, 5, 20))
End Function

