VERSION 5.00
Begin VB.Form FrmPrimeraVez 
   BorderStyle     =   0  'None
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNumDias 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFDFC&
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   2475
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "TxtNumDias"
      Top             =   2550
      Width           =   825
   End
   Begin VB.OptionButton Opt2 
      BackColor       =   &H00000080&
      Caption         =   "Demo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   3450
      TabIndex        =   6
      Top             =   2280
      Width           =   885
   End
   Begin VB.OptionButton Opt1 
      BackColor       =   &H00000080&
      Caption         =   "Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   2490
      TabIndex        =   5
      Top             =   2280
      Width           =   885
   End
   Begin VB.TextBox TxtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFDFC&
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   2475
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   1890
      Width           =   3345
   End
   Begin VB.TextBox txtserie4 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFDFC&
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   4995
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1530
      Width           =   825
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   285
      Left            =   2865
      TabIndex        =   8
      Top             =   3330
      Width           =   1200
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancelar"
      Height          =   285
      Left            =   4095
      TabIndex        =   9
      Top             =   3330
      Width           =   1200
   End
   Begin VB.TextBox txtserie3 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFDFC&
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1530
      Width           =   825
   End
   Begin VB.TextBox txtserie2 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFDFC&
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   3315
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1530
      Width           =   825
   End
   Begin VB.TextBox txtserie1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFDFC&
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   2475
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1530
      Width           =   825
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Dias"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   885
      TabIndex        =   15
      Top             =   2580
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Instalacion"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   885
      TabIndex        =   14
      Top             =   2250
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   6330
      Picture         =   "FrmPrimeraVez.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   870
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Serie"
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   885
      TabIndex        =   12
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave"
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   885
      TabIndex        =   11
      Top             =   1560
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   330
      X2              =   7785
      Y1              =   3090
      Y2              =   3090
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1620
      Index           =   1
      Left            =   405
      Top             =   1380
      Width           =   7305
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Height          =   1740
      Index           =   0
      Left            =   330
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J y L. Software System S.A.C."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   330
      TabIndex        =   13
      Top             =   945
      Width           =   4110
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DBFDFC&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2430
      Left            =   135
      Top             =   885
      Width           =   7860
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro  de Licencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   855
      TabIndex        =   10
      Top             =   90
      Width           =   2925
   End
End
Attribute VB_Name = "FrmPrimeraVez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo    : FRMPRIMERAVEZ
'* Tipo              : FORMULARIO
'* Descripcion       : FORMULARIO PARA VALIDAR LA LICENCIA DEL SISTEMA EN LA PC, SE VALIDA LEYENDO EL
'*                     NUMERO DE SERIE DEL DISCO Y GENERANDO UN NUMERO DE SERIE, EL CUAL SERA VALIDADO
'*                     EN EL PROGRAMA KEYGEN.EXE EL CUAL DEVOLVERA UN NUMERO DE LICENCIA PARA ACTIVAR
'*                     EL PROGRAMA EN LA PC
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 03/09/09
'* VERSION           : 1.0
'*****************************************************************************************************

Option Explicit
Dim NumVeces As Integer

Private Sub CmdAceptar_Click()
    ' VALIDAMOS QUE LOS DATOS ESTEN INGRESADO CORRECTAMENTE
    If Opt1.Value = False And Opt2.Value = False Then
        MsgBox "No ha especificado si la instalación es definitiva o temporal", vbInformation + vbOKOnly + vbDefaultButton1, "SEVEN Soft"
        Opt1.SetFocus
        Exit Sub
    End If
    
    If Opt1.Value = True Then
        If NulosC(TxtPass.Text) = "" Then
            MsgBox "No ha especificado el número de serie", vbInformation + vbOKOnly + vbDefaultButton1, "SEVEN Soft"
            Exit Sub
        End If
    End If
    
    
    If Opt2.Value = True Then
        TxtNumDias.Text = 30
        If TxtNumDias.Text = "" Then
            MsgBox "No ha especificado el número de dias para la demostración", vbInformation + vbOKOnly + vbDefaultButton1, "SEVEN Soft"
            TxtNumDias.SetFocus
            Exit Sub
        End If
    End If

    Dim xCad As String
    
    ' UNE EN UNA SOLA CADENA EL NUMERO DE SERIE GENERADO POR EL SISTEMA
    xCad = Trim(txtserie1.Text) + "-" + Trim(txtserie2.Text) + "-" + Trim(txtserie3.Text) + "-" + Trim(txtserie4.Text)
    
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT * FROM mae_pc", xCon
    
    If Opt1.Value = True Then
        ' EVALUAMOS SI EL NUMERO DE SERIE REGISTRADO POR EL USUARIO ES VALIDO PARA EN NUMERO DE SERIE GENERADO POR EL SISTEMA
        If (EvaluaClave(xCad, TxtPass.Text) = True) Then
            ' SI EL NUMERO DE SERIE ES CORRECTO REGISTRAMOS LA LICENCIA EN LA BASE DE DATOS PARA QUE NO SE VUELVA A PEDIR
            ' EL NUMERO DE LICENCIA
           
            Rst.AddNew
            Rst("id") = HallaCodigoTabla("mae_pc", xCon, "id")
            Rst("serdis") = Trim(Str(LeerNumeroDisco("c:")))
            Rst("registro") = 1
            Rst.Update
           
        Else
            ' SI EL NUMERO DE SERIE INGRESADO POR EL USUARIO NO ES VALIDO PARA EL NUMERO DE SERIE DEL SISTEMA SE EMITE UNA
            ' SEÑAL DE ALERTA AVISANDO QUE EL NUMERO DE SERIE DIGITADO NO CORRESPONDE, ASI MISMO SE INCREMENTA EN UNO EL
            ' NUMERO DE CHANCES PARA EL IGRESO DEL NUMERO DE SERIE CORRECTO, CUANDO EL CONTADOR LLEGUE A 3 EL SISTEMA SE
            ' CERRARA Y NO PODRA INGRESAR AL SISTEMA HASTA INGRESAR EL NUMERO DE SERIE CORRECTO
            MsgBox ("El numero de serie ingresado no es valido, ingreselo de nuevo o solicite un nuevo numero de serie a Enrique Pollongo escribiendo un correo a eps_76@hotmail.com"), vbInformation + vbOKOnly + vbDefaultButton1, "SEVEN Soft"
            NumVeces = NumVeces + 1
            If NumVeces = 3 Then
                Unload Me
            End If
            Exit Sub
        End If
    End If
    
    If Opt2.Value = True Then
        Rst.AddNew
        Rst("id") = HallaCodigoTabla("mae_pc", xCon, "id")
        Rst("serdis") = Trim(Str(LeerNumeroDisco("c:")))
        Rst("registro") = 0
        Rst.Update
    End If
    
    Set Rst = Nothing
    Unload Me
    MDIPrincipal.Show
End Sub

Private Sub CmdCancel_Click()
    ' CIERRA EL FORMULARIO
    Unload Me
End Sub

Private Sub Form_Activate()
    ' PRIMER EVENTO QUE EJECUTARA EL FORMULARIO, AQUI SE CARGA EL NUMERO DE SERIE QUE GENERA EL SISTEMA PARA LA VALIDACION, Y
    ' CARGA LOS GRAFICOS NECESARIO PARA EL FORMULARIO
    Dim xCad As String
    xCad = CadOrigen
    On Error Resume Next
    Me.Picture = LoadPicture(AP_RUTABM + "marco4.bmp")
    Err.Clear
    Blanquea
        
    txtserie1.Text = Mid(xCad, 1, 4)
    txtserie2.Text = Mid(xCad, 6, 4)
    txtserie3.Text = Mid(xCad, 11, 4)
    txtserie4.Text = Mid(xCad, 16, 5)
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Blanquea()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : BLANQUEA LOS CONTROLES DEL FORMULARIO, PARA EL INGRESO DE DATOS
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Blanquea()
    txtserie1.Text = ""
    txtserie2.Text = ""
    txtserie3.Text = ""
    txtserie4.Text = ""
    TxtPass.Text = ""
    TxtNumDias.Text = ""
    Opt1.Value = False
    Opt2.Value = False
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : CadOrigen()
'* Tipo           : FUNCCION
'* Descripcion    : GENERA EL NUMERO DE SERIE PARA LA UNIDAD DE DISCO DONDE SE ESTA EJECUTANDO EL
'*                  SISTEMA
'* Paranetros     : NULL
'* Retorna        : STRING
'*****************************************************************************************************
Function CadOrigen() As String
    Dim Num1 As String
    Dim Num2 As String
    Dim Num3 As String
    Dim Num4 As String

    Dim xCadena As String
    Dim NumSerie As String
    
    ' CARGAMOS EL NUMERO DE SERIE DE LA UNIDAD DE DISCO POR DEFECTO C:\
    NumSerie = LeerNumeroDisco("C:")
    
    ' GENERAMOS EL NUMERO DE SERIE PARA EL LA UNIDAD DE DISCO ACTUAL
    xCadena = Trim(Mid(NumSerie, 1, 3)) + (Mid(Format(Date, "dd/mm/yy"), 1, 2))
    xCadena = xCadena + Trim(Mid(NumSerie, 4, 3)) + (Mid(Format(Date, "dd/mm/yy"), 4, 2))
    xCadena = xCadena + Trim(Mid(NumSerie, 7, 5)) + (Mid(Format(Date, "dd/mm/yy"), 7, 2))
        
    Num1 = Trim(Mid(xCadena, 1, 4))
    Num2 = Trim(Mid(xCadena, 5, 4))
    Num3 = Trim(Mid(xCadena, 9, 4))
    Num4 = Trim(Mid(xCadena, 14, 5))
    
    ' RETORNA EL NUMERO DE SERIE PARA LA UNIDAD DE DISCO
    CadOrigen = Num1 + "-" + Num2 + "-" + Num3 + "-" + Num4
End Function

'*****************************************************************************************************
'* Nombre Modulo  : EvaluaClave()
'* Tipo           : FUNCCION
'* Descripcion    : EVALUA LA NUMERO DE SERIE GENERADA POR EL SISTEMA Y EL NUMERO DE LICIENCIA
'*                  GENERADO POR EL PROGRAMA KEYGEN.EXE, DEVEUELVE VERDADERO SI EL PASSOWORD INGRESO
'*                  ES EL CORRECTO
'* Paranetros     : NOMBRE   |  TIPO    |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  clave    | STRING   | NUMERO DE SERIE GENERADO POR EL SISTEMA
'*                  password | STRING   | PASWORD GENERADO POR EL PROGRAMA KEYGEN
'* Retorna        : LOGICO
'*****************************************************************************************************
Function EvaluaClave(clave As String, password As String) As Boolean
    Dim Num1 As String
    Dim Num2 As String
    Dim Num3 As String
    Dim Num4 As String
    Dim xNumero1 As Double
    Dim xNumero2 As Double
    Dim Valor As Integer
    
    ' OBTENEMOS EL DIGITO COMBINABLE PARA GENERAR EL NUMERO DE SERIE, ESTO LO SACAMOS DEL ULTIMO DIGITO DEL PASSWORD,
    ' ES TO ES PARA OBTENER EL NUMERO DE SERIE QUE GENERARIA EL PROGRAMA KEYGEN.EXE
    Valor = Val(Mid(Trim(password), Len(password) - 1, 2))
    
    Num1 = Trim(Mid(clave, 1, 4))
    Num2 = Trim(Mid(clave, 6, 4))
    Num3 = Trim(Mid(clave, 11, 4))
    Num4 = Trim(Mid(clave, 16, 5))
    
    xNumero1 = Val(Num1) + Val(Num2)
    xNumero2 = Val(Num3) + Val(Num4)
    
    ' GENERAMOS EL NUMERO DE SERIE QUE GENERARIA KEYGEN.EXE
    ' CAMBIAR EL 6 PARA GENERAR UNA NUEVA SECUENCIA DEL NUMERO DE SERIE
    xNumero1 = xNumero1 * (Valor + 6)
    xNumero2 = xNumero2 * (Valor + 6)
    
    Dim xCad1 As String
    xCad1 = Trim(Str(xNumero1)) + Trim(Str(xNumero2)) + Trim(Str(Valor))
    xCad1 = Mid(xCad1, 1, 4) + "-" + Trim(Mid(xCad1, 5, 20))

    ' EVALUAMOS LA CADENA GENERADA CON EL PASSWORD INGRESO
    If xCad1 = password Then
        ' SI LA CADENA GENERADA COINCIDE CON EL PASSWORD LA FUNCION RETORNA VERDADERO
        EvaluaClave = True
    Else
        ' SI LA CADENA GENERADA NO COINCIDE CON EL PASSWORD LA FUNCION RETORNA FALSO
        EvaluaClave = False
    End If
End Function

Private Sub Opt1_Click()
    TxtNumDias.Locked = True
    TxtNumDias.Text = ""
End Sub

Private Sub Opt2_Click()
    TxtNumDias.Locked = False
End Sub

Private Sub TxtNumDias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtserie1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtserie2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtserie3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtserie4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub
