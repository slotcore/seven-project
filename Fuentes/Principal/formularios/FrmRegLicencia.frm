VERSION 5.00
Begin VB.Form FrmRegLicencia 
   Caption         =   "Registro de Licencias"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1410
      Left            =   6630
      TabIndex        =   8
      Top             =   -900
      Visible         =   0   'False
      Width           =   3945
      Begin VB.CommandButton Command4 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1395
         TabIndex        =   11
         Top             =   1080
         Width           =   1170
      End
      Begin VB.TextBox TxtNumSer 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "TxtNumSer"
         Top             =   75
         Width           =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   3930
         Y1              =   1395
         Y2              =   1395
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   1395
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   3930
         X2              =   3930
         Y1              =   15
         Y2              =   1395
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   3945
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label Label3 
         Caption         =   "Envie este numero a eps_76@hotmail.com o comuniquese a 782-9997 para que le entreguen el numero de licencia correspondiente"
         Height          =   585
         Left            =   105
         TabIndex        =   9
         Top             =   420
         Width           =   3750
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   405
      Left            =   4200
      TabIndex        =   4
      Top             =   1125
      Width           =   1110
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Height          =   405
      Left            =   3030
      TabIndex        =   3
      Top             =   1125
      Width           =   1110
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar Clave"
      Height          =   405
      Left            =   570
      TabIndex        =   2
      Top             =   1125
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   15
      TabIndex        =   5
      Top             =   -45
      Width           =   5370
      Begin VB.TextBox TxtLic 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1695
         TabIndex        =   1
         Text            =   "TxtLic"
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox TxtNum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1695
         TabIndex        =   0
         Text            =   "TxtNum"
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Numero de Licencia"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   630
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clave de Registro"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   285
         Width           =   1260
      End
   End
End
Attribute VB_Name = "FrmRegLicencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Con As New ADODB.Connection

Private Sub Command1_Click()
    TxtNumSer.Text = ""
    TxtNumSer.Text = FrmPrimeraVez.CadOrigen
    Frame2.Left = 825
    Frame2.Top = 60
    Frame2.Visible = True
End Sub


Private Sub Command2_Click()
    If NulosC(TxtNum.Text) = "" Then
        MsgBox "No ha especificado la clave de registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNum.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtLic.Text) = "" Then
        MsgBox "No ha especificado el numero de licencia para esta PC", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtLic.SetFocus
        Exit Sub
    End If
    
    Dim NumSerDis As String
    Dim Rst As New ADODB.Recordset
    
    NumSerDis = Trim(Str(LeerNumeroDisco("c:")))
    
    RST_Busq Rst, "SELECT * FROM mae_pc WHERE serdis = '" & NumSerDis & "'", Con
    
    If Rst.RecordCount = 1 Then
        If (FrmPrimeraVez.EvaluaClave(TxtNum.Text, TxtLic.Text) = True) Then
            Rst("registro") = 1
            Rst.Update
            MsgBox "Numero de licencia se registro con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set Rst = Nothing
            Unload Me
            Exit Sub
        Else
            MsgBox "Numero de licencia no valido, por favor ingrese otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtLic.SetFocus
            Set Rst = Nothing
            Exit Sub
        End If
    End If
    
    
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    Frame2.Visible = False
End Sub

Private Sub Form_Activate()
    Dim NumSerDis As String
    Dim Rst As New ADODB.Recordset
    
    Dim xFun As New eps_librerias.FuncionesData
    
    xFun.F_BASEDATOS = AP_RUTABD + "data.mdb"                                           ' PASAMOS LA RUTA DE LA BASE DE DATOS PARA ABRIR LA CONECCION
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"                                       ' PASAMOS LA RUTA DEL ARCHIVO DE TRABJO DE LA BASE DE DATOS
    xFun.F_PASSWORD = Eps_Pass                                                          ' PASAMOS EL PASWORD DE LA BASE DE DATOS
    xFun.F_USUARIO = Eps_User                                                           ' PASAMOS EL USUARIO DE LA BASE DE DATOS
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"                                        ' PASAMOS EL NOMBRE DEL PROVEEDORE DE DATOS PARA ADO 2.5
    
    Set Con = xFun.AbrirConeccion                                                       ' ABRIMOS LA CONECCION DE DATOS
    Set xFun = Nothing
    
    TxtNum.Text = ""
    TxtLic.Text = ""
    
    NumSerDis = Trim(Str(LeerNumeroDisco("c:")))
    
    RST_Busq Rst, "SELECT * FROM mae_pc WHERE serdis = '" & NumSerDis & "'  AND registro = 1", Con
    If Rst.State = 1 Then
        If Rst.RecordCount = 1 Then
            MsgBox "Esta Pc ya fue registrada, No se puede volver a registrar esta Pc", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Unload Me
            Exit Sub
        End If
    End If
    
End Sub


Private Sub TxtLic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
