VERSION 5.00
Begin VB.Form FrmAcceso 
   Caption         =   "Ingreso de Usuarios"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   1845
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmndCan 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   2355
      TabIndex        =   3
      Top             =   1275
      Width           =   1035
   End
   Begin VB.CommandButton CmdAcep 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1275
      TabIndex        =   2
      Top             =   1275
      Width           =   1035
   End
   Begin VB.TextBox TxtPas 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1860
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "TxtPas"
      Top             =   750
      Width           =   1260
   End
   Begin VB.TextBox TxtUsu 
      Height          =   300
      Left            =   1860
      TabIndex        =   0
      Text            =   "TxtUsu"
      Top             =   315
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   180
      Index           =   1
      Left            =   855
      TabIndex        =   5
      Top             =   810
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   180
      Index           =   0
      Left            =   855
      TabIndex        =   4
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "FrmAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAcep_Click()
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT * FROM usuarios WHERE usuario = '" & NulosC(TxtUsu.Text) & "' and password = '" & NulosC(TxtPas.Text) & "'", xCon
    If Rst.RecordCount <> 0 Then
        MsgBox "Bievenido usuario " & Rst("nombre"), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Unload Me
        xNivelUsuario = Rst("nivel")
        FrmMenuRapido.Show
    Else
        MsgBox "El usuario especificado no existe", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtUsu.Text = ""
        TxtPas.Text = ""
        TxtUsu.SetFocus
    End If
    Set Rst = Nothing
End Sub

Private Sub CmndCan_Click()
    Set xCon = Nothing
    Unload Me
    End
End Sub

Private Sub Form_Activate()
    TxtUsu.SetFocus
End Sub

Private Sub Form_Load()
    TxtUsu.Text = ""
    TxtPas.Text = ""

End Sub

Private Sub TxtPas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtUsu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
