VERSION 5.00
Begin VB.Form FrmFiltro 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtro de Registros"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmNumerico 
      BackColor       =   &H00C0C000&
      Caption         =   " [  Opciones de Filtrado  ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1440
      Left            =   60
      TabIndex        =   11
      Top             =   1590
      Visible         =   0   'False
      Width           =   3570
      Begin VB.OptionButton OptNum4 
         BackColor       =   &H00C0C000&
         Caption         =   "Diferente que"
         Height          =   285
         Left            =   255
         TabIndex        =   15
         Top             =   1125
         Width           =   2760
      End
      Begin VB.OptionButton OptNum2 
         BackColor       =   &H00C0C000&
         Caption         =   "Mayor que"
         Height          =   285
         Left            =   255
         TabIndex        =   14
         Top             =   555
         Width           =   2760
      End
      Begin VB.OptionButton OptNum1 
         BackColor       =   &H00C0C000&
         Caption         =   "Igual que"
         Height          =   285
         Left            =   255
         TabIndex        =   13
         Top             =   270
         Width           =   2760
      End
      Begin VB.OptionButton OptNum3 
         BackColor       =   &H00C0C000&
         Caption         =   "Menor que"
         Height          =   285
         Left            =   255
         TabIndex        =   12
         Top             =   840
         Width           =   2760
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Height          =   1440
      Left            =   3750
      TabIndex        =   10
      Top             =   1485
      Width           =   1785
      Begin VB.CommandButton CmdCan 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   435
         Left            =   135
         TabIndex        =   6
         Top             =   795
         Width           =   1500
      End
      Begin VB.CommandButton CmdAcep 
         Caption         =   "&Aceptar"
         Height          =   435
         Left            =   135
         TabIndex        =   5
         Top             =   345
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   " [  Opciones de Filtrado  ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1440
      Left            =   60
      TabIndex        =   9
      Top             =   1485
      Width           =   3570
      Begin VB.OptionButton Opt4 
         BackColor       =   &H00C0C000&
         Caption         =   "Coincidir Todo"
         Height          =   285
         Left            =   255
         TabIndex        =   4
         Top             =   1005
         Width           =   2760
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00C0C000&
         Caption         =   "Cualquier Parte"
         Height          =   285
         Left            =   255
         TabIndex        =   2
         Top             =   330
         Width           =   2760
      End
      Begin VB.OptionButton Opt2 
         BackColor       =   &H00C0C000&
         Caption         =   "Principio"
         Height          =   285
         Left            =   255
         TabIndex        =   3
         Top             =   660
         Width           =   2760
      End
   End
   Begin VB.ComboBox CboCampos 
      Height          =   315
      Left            =   1740
      TabIndex        =   1
      Text            =   "CboCampos"
      Top             =   900
      Width           =   3780
   End
   Begin VB.TextBox TxtCrit 
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Text            =   "TxtCrit"
      Top             =   360
      Width           =   5475
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "En el Campo"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   945
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   105
      Width           =   495
   End
End
Attribute VB_Name = "FrmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RstFil As New ADODB.Recordset
Dim xCampoSel As String

Private Sub CboCampos_Click()
    xCampoSel = BuscaCampoLista(CboCampos.Text, 0, 1, xCampos)

    If Trim(BuscaCampoLista(CboCampos.Text, 0, 2, xCampos)) = "N" Or Trim(BuscaCampoLista(CboCampos.Text, 0, 2, xCampos)) = "F" Then
        'CUANDO EL CAMPO SEA NUMERICO
        FrmNumerico.Left = 60
        FrmNumerico.Top = 1485
        FrmNumerico.Visible = True
        OptNum1.Value = True
    Else
        'CUANDO EL CAMPO SEA CARACTER
        FrmNumerico.Visible = False
    End If
End Sub

Private Sub CmdAcep_Click()
    If TxtCrit.Text = "" Then Exit Sub
    
    'CUANDO EL CAMPO SEA CARACTER
    If Trim(BuscaCampoLista(CboCampos.Text, 0, 2, xCampos)) = "C" Then
        If Opt1.Value = True Then
            RstFil.Filter = "" & xCampoSel & " like '*" & Trim(TxtCrit.Text) & "*'"
        End If
        If Opt2.Value = True Then
            RstFil.Filter = "" & xCampoSel & " like '" & Trim(TxtCrit.Text) & "*'"
        End If
        If Opt4.Value = True Then
            RstFil.Filter = "" & xCampoSel & " = '" & Trim(TxtCrit.Text) & "'"
        End If
    End If
    
    'CUANDO EL CAMPO SEA NUMERICO
    If Trim(BuscaCampoLista(CboCampos.Text, 0, 2, xCampos)) = "N" Then
        If OptNum1.Value = True Then
            RstFil.Filter = "" & xCampoSel & " = " & Val(TxtCrit.Text) & ""
        End If
        If OptNum2.Value = True Then
            RstFil.Filter = "" & xCampoSel & " > " & Val(TxtCrit.Text) & ""
        End If
        If OptNum3.Value = True Then
            RstFil.Filter = "" & xCampoSel & " < '" & Val(TxtCrit.Text) & "'"
        End If
        If OptNum4.Value = True Then
            RstFil.Filter = "" & xCampoSel & " <> '" & Val(TxtCrit.Text) & "'"
        End If
    End If
    
    'CUANDO EL CAMPO SEA FECHA
    If Trim(BuscaCampoLista(CboCampos.Text, 0, 2, xCampos)) = "F" Then
        If OptNum1.Value = True Then
            RstFil.Filter = "" & xCampoSel & " = ('" & TxtCrit.Text & "')"
        End If
        If OptNum2.Value = True Then
            RstFil.Filter = "" & xCampoSel & " > ('" & TxtCrit.Text & "')"
        End If
        If OptNum3.Value = True Then
            RstFil.Filter = "" & xCampoSel & " < ('" & TxtCrit.Text & "')"
        End If
        If OptNum4.Value = True Then
            RstFil.Filter = "" & xCampoSel & " <> ('" & TxtCrit.Text & "')"
        End If
    End If
    
    Me.Hide
End Sub

Private Sub CmdCan_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    Blanquea
    LLenarCombo
    Opt1.Value = True
    
    xCampoSel = BuscaCampoLista(CboCampos.Text, 0, 1, xCampos)
End Sub

Sub Blanquea()
    TxtCrit.Text = ""
    CboCampos.Text = ""
End Sub

Sub LLenarCombo()
    Dim A As Integer
    For A = LBound(xCampos) To UBound(xCampos)
        CboCampos.AddItem xCampos(A, 0)  'muestra los titulos de los campos en el menu
        If A = UBound(xCampos) - 1 Then
            Exit For
        End If
    Next A
    CboCampos.SelText = xCampos(0, 0)
End Sub

Private Sub TxtCrit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub
