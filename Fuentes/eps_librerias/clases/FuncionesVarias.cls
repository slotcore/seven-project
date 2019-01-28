VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuncionesVarias"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit  'NO SE PONE OPTION ESPLICIT POR SI NO CORRE LA FUNCION PARA CONVERTIRNUMEROS A LETRAS
Dim o_Script As MSScriptControl.ScriptControl

Sub EjecutarCodigo(Texto As String, Formulario As Object, NombreFormulario As String)
    ' Agrega el código del text1, es decir lo ejecuta
    Set o_Script = New MSScriptControl.ScriptControl
       
    ' Establece como lenguaje VBScript
    o_Script.Language = "VBScript"
       
    'o_Script.AddObject "Form1", Form1
    o_Script.AddObject NombreFormulario, Formulario
    On Error GoTo ErrSub:
    o_Script.AddCode Texto
    Exit Sub

'Error
ErrSub:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Function NumeroLetra(Numero As Double, Moneda As Integer) As String
    Dim xCad As String
    Dim Letra As String
    Dim Entero As String
    Dim Deximal As String
    Dim xLongText As Integer
    
    xCad = Format(Trim(Str(Numero)), "0.00")
    Entero = Mid(Trim(xCad), 1, Len(Trim(xCad)) - 3)
    Deximal = Mid(Trim(xCad), Len(Trim(xCad)) - 1, 2)
    
    Letra = DeNumALetra(Val(Entero)) + " CON " + Deximal + "/100 "
    Dim Cad As String
    Cad = ""
    
    If Numero >= 1000 And Numero <= 1999 Then
        Cad = "Un "
    End If
    If Moneda = 1 Then
        'Letra = Cad + Trim(Letra) + " NUEVOS SOLES"
        ' ACTUALIZADO DEBIDO A CAMBIO EN LA NORMA DE NUEVOS SOLES A SOLES
        Letra = Cad + Trim(Letra) + " SOLES"
    Else
        Letra = Cad + Trim(Letra) + " US $ DOLARES AMERICANOS"
    End If
    NumeroLetra = UCase(Letra)
    
'    If Val(Moneda) = 1 Then
'        Letra = Trim(Letra) + " Nuevos Soles"
'    Else
'        Letra = Trim(Letra) + " US $ Dolares Americanos"
'    End If
'    NumeroLetra = UCase(Letra)
End Function

Function DeNumALetra(ByVal Numero As Variant, Optional TipoCambioSingular As String = "", Optional TipoCambioPlural As String, Optional subTipoCambioSingular As String, Optional subTipoCambioPlural As String, Optional xInternal As Long = 0) As String
         Dim sNum As String, vNum() As String, x As Long, Y As Long, Z As Long, sTmp As String
         Dim D1 As String, D2 As String, D3 As String, D4 As String, DFinal As String
         Dim tNum As String, B1 As Boolean, B2 As Boolean, B3 As Boolean
         Dim wNum() As String, xNums As String, xWords As String, Nombres() As String
         
         '***********************************************************************************************
         '* Esta función convierte números en palabras, sin importar el contexto donde se encuentren    *
         '* La presición (por limitancia del lenguaje) es de 28B, Ej: 9999999999999999999999999999 max. *
         '***********************************************************************************************
         
         'Convierte el valor en un string
         sNum = Trim(CStr(Numero))
                
         'Procesa cada número que exista en la variable por separado
         If xInternal = 0 Then
            'Separa los números limpios de las palabras y los procesa por separado (no incluye números con letras)
            wNum = Split(sNum, " ")
            For x = 0 To UBound(wNum)
                'Concatena los strings o números según corresponda
                If IsNumeric(wNum(x)) Then
                   'Separa los enteros de los decimales para procesarlos por separado
                   If Int(Val(wNum(x))) < wNum(x) Then
                      D1 = Int(Val(wNum(x)))
                      D2 = Mid(CStr(wNum(x)), Len(D1) + 2)
                      DFinal = DFinal & IIf(D1 < 0, "menos ", "") & DeNumALetra(D1, TipoCambioSingular, TipoCambioPlural, 1) & " con "
                      DFinal = DFinal & DeNumALetra(D2, subTipoCambioSingular, subTipoCambioPlural, , , 1) & " "
                   Else
                      DFinal = DFinal & IIf(wNum(x) < 0, "menos ", "") & DeNumALetra(wNum(x), TipoCambioSingular, TipoCambioPlural, subTipoCambioSingular, subTipoCambioPlural, 1) & " "
                   End If
                Else
                   DFinal = DFinal & wNum(x) & " "
                End If
            Next
         Else
            
            'ELimina el signo
            If Not IsNumeric(Left(sNum, 1)) Then
               sNum = Mid(sNum, 2)
            End If
         
            'Elimina cualquier formato posible (incluye valores científicos)
            sNum = Format(sNum, "0")
            
            'Completa con ceros a la izquierda hasta obtener una longitud múltiplo de 3
            Do While Len(sNum) Mod 3 <> 0
               sNum = "0" & sNum
            Loop
         
            'Dimenciona un arreglo con espacio para cada una de las centenas
            ReDim vNum(Len(sNum) / 3 - 1)
            
            'Carga el arreglo con las centenas que corresponda
            For x = 0 To UBound(vNum, 1)
                vNum(x) = Mid(sNum, (x + 1) * 3 - 2, 3)
            Next
             
            'Si el arreglo contiene una sola centena, la convierte en palabras
            If UBound(vNum, 1) = 0 Then
                'Asigna los dígitos de la centena y recuerda si son mayores que cero
                D3 = Left(sNum, 1): B3 = Val(D3) > 0
                D2 = Mid(sNum, 2, 1): B2 = Val(D2) > 0
                D1 = Right(sNum, 1): B1 = Val(D1) > 0
                
                'Procesa las unidades
                Select Case D1
                       Case "1": DFinal = "un"
                       Case "2": DFinal = "dos"
                       Case "3": DFinal = "tres"
                       Case "4": DFinal = "cuatro"
                       Case "5": DFinal = "cinco"
                       Case "6": DFinal = "seis"
                       Case "7": DFinal = "siete"
                       Case "8": DFinal = "ocho"
                       Case "9": DFinal = "nueve"
                End Select
                
                'Procesa las decenas
                Select Case D2
                       Case "1"
                            'Maneja lógica del retrasado mental que puso nombres ilógicos a algunos números.
                            Select Case D1
                                   Case "0": DFinal = "diez"
                                   Case "1": DFinal = "once"
                                   Case "2": DFinal = "doce"
                                   Case "3": DFinal = "trece"
                                   Case "4": DFinal = "catorce"
                                   Case "5": DFinal = "quince"
                                   Case "6": DFinal = "dieciséis"
                                   Case Else
                                        DFinal = "dieci" & DFinal
                            End Select
                       Case "2"
                            If B1 Then
                               If D1 = "2" Then DFinal = "dós"
                               If D1 = "3" Then DFinal = "trés"
                               DFinal = "veinti" & DFinal
                            Else
                               DFinal = "veinte"
                            End If
                       Case "3": If B1 Then DFinal = "treinta y " & DFinal Else DFinal = "treinta"
                       Case "4": If B1 Then DFinal = "cuarenta y " & DFinal Else DFinal = "cuarenta"
                       Case "5": If B1 Then DFinal = "cincuenta y " & DFinal Else DFinal = "cincuenta"
                       Case "6": If B1 Then DFinal = "sesenta y " & DFinal Else DFinal = "sesenta"
                       Case "7": If B1 Then DFinal = "setenta y " & DFinal Else DFinal = "setenta"
                       Case "8": If B1 Then DFinal = "ochenta y " & DFinal Else DFinal = "ochenta"
                       Case "9": If B1 Then DFinal = "noventa y " & DFinal Else DFinal = "noventa"
                End Select
                
                'Procesa las centenas
                Select Case D3
                       Case "1": If B1 Or B2 Then DFinal = "ciento " & DFinal Else DFinal = "cien"
                       Case "2": If B1 Or B2 Then DFinal = "doscientos " & DFinal Else DFinal = "doscientos"
                       Case "3": If B1 Or B2 Then DFinal = "trescientos " & DFinal Else DFinal = "trescientos"
                       Case "4": If B1 Or B2 Then DFinal = "cuatrocientos " & DFinal Else DFinal = "cuatrocientos"
                       Case "5": If B1 Or B2 Then DFinal = "quinientos " & DFinal Else DFinal = "quinientos"
                       Case "6": If B1 Or B2 Then DFinal = "seiscientos " & DFinal Else DFinal = "seiscientos"
                       Case "7": If B1 Or B2 Then DFinal = "setecientos " & DFinal Else DFinal = "setecientos"
                       Case "8": If B1 Or B2 Then DFinal = "ochocientos " & DFinal Else DFinal = "ochocientos"
                       Case "9": If B1 Or B2 Then DFinal = "novecientos " & DFinal Else DFinal = "novecientos"
                End Select
                
                'Si es la ejecución principal efectua algunos arreglines
                If xInternal = 1 Then
                   'Validación del cero
                   If DFinal = "" Then DFinal = "cero"
                   'Validación de terminados en "un"
                   If Right(DFinal, 2) = "un" And TipoCambioSingular = "" Then DFinal = DFinal & "o"
                End If
                
            Else 'Si es más de una centena, las separa y procesa independientemente
                Y = -1
                Z = 1
                For x = UBound(vNum) To 0 Step -1
                    Y = Y + 1
                    
                    'Convierte la centena en palabras
                    tNum = DeNumALetra(vNum(x), xInternal:=2)
                    
                    'Arregla la terminación "uno" cuando corresponde
                    If Y = 0 And Right(tNum, 2) = "un" And TipoCambioSingular & TipoCambioPlural = "" Then tNum = tNum + "o"
                    
                    'Genera un valor temporal para poder modificar
                    sTmp = tNum
                    
                    'Asigna los nombres genéricos principales
                    Nombres = Split(" mil , millón , millones , billón , billones , trillón , trillones , cuatrillón , cuatrillones , quintillón , quintillones , sextillón , sextillones , septillón , septillones , octillón , octillones, nonillón , nonillones , decillón , decillones , undecillón , undecillones , duodecillón , duodecillones , tredecillón , tredecillones , cuatordecillón , cuatordecillones , quindecillón , quindecillones , sexdecillón , sexdecillones , septendecillón , septendecillones , octodecillón , octodecillones , novendecillón , novendecillones , vigintillón , vigintillones ", ",")
                    
                    'Controla que el índice de nombres no salga de los límites
                    If Y > UBound(Nombres) Then
                       WordNum = "?"
                       Exit Function
                    End If
                    
                    'Asigna los nombres correspondientes
                    If Y Mod 2 > 0 Then
                       D1 = Nombres(0)
                       D2 = Nombres(Y - 1)
                    ElseIf Y > 0 Then
                       D1 = Nombres(Y - 1)
                       D2 = Nombres(Y)
                    Else
                       D1 = "": D2 = ""
                    End If
                    
                    'Actualiza el nombre del número
                    Select Case Y Mod 2
                           Case 0: If sTmp = "un" Then sTmp = sTmp & D1 Else sTmp = sTmp & IIf(tNum = "", "", D2)
                           Case Else
                                If sTmp = "un" Then sTmp = ""
                                sTmp = sTmp & IIf(tNum = "", "", D1)
                                If x = 0 And Y > 1 Then
                                   If InStr(1, DFinal, D2, vbTextCompare) = 0 Then sTmp = sTmp & Mid(D2, 2)
                                End If
                    End Select
                    DFinal = sTmp & DFinal
                Next
            End If
         End If
         
         'Aplica el tipo de moneda cuando corresponda
         If xInternal = 1 Then DFinal = DFinal & " " & IIf(Format(sNum, "#0") = "1", TipoCambioSingular, TipoCambioPlural)
    
         'Asigna el número en palabras
          DeNumALetra = Trim(DFinal)
End Function

Function BuscaCampoLista(Texto As String, IndiceBusca As String, IndiceDevuelve As String, Campos() As String) As String
    Dim A As Integer
    
    For A = LBound(Campos) To UBound(Campos)
        If Campos(A, IndiceBusca) = Texto Then
            BuscaCampoLista = Campos(A, IndiceDevuelve)
            Exit For
        End If
        If A = UBound(Campos) - 1 Then
            Exit For
        End If
    Next A
End Function

Sub CerrarAplicacionExterna(ByVal processName As String)
    On Error GoTo ErrHandler
    Dim oWMI
    Dim ret
    Dim sService
    Dim oWMIServices
    Dim oWMIService
    Dim oServices
    Dim oService
    Dim servicename

    Set oWMI = GetObject("winmgmts:")
    Set oServices = oWMI.InstancesOf("win32_process")

    For Each oService In oServices
        servicename = _
            LCase(Trim(CStr(oService.Name) & ""))

        If InStr(1, servicename, _
            LCase(processName), vbTextCompare) > 0 Then
            ret = oService.Terminate
        End If
    Next

    Set oServices = Nothing
    Set oWMI = Nothing
    Exit Sub

ErrHandler:
    Err.Clear
End Sub

Public Sub CentrarFormulario(Frm As Object)
    '--frm formulario
    On Error Resume Next
    If Frm.WindowState <> 2 Then
    Frm.Left = (Screen.Width - Frm.Width) / 2
    Frm.Top = (Screen.Height - Frm.Height) / 2
    ' frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 2
    End If
    Err.Clear
End Sub

Public Function ConvertSeg(dHora As String) As Long
    '--esta funcion devuelve los segundos segun la hora como parametro
    Dim mHora&
    Dim mMin&
    Dim mSeg&
    Dim nTmp As String
    
    Dim PosSeparador As Integer '--indica la separacion ":"
    
    If IsDate(dHora) = False Then
        ConvertSeg = -1
        '--se asume que el valor ingresado es de la sig. forma
        '-- "1008:15:20" hora:minuto:segundo
        
'        nTmp = Right(dHora, 2)
'        mSeg = NulosN(Right(dHora, 2))
'
'        dHora = mID(dHora, 1, Len(dHora) - 3)
'
'        mMin = NulosN(Right(dHora, 2)) * 60
'        nTmp = Right(dHora, 2)
'        dHora = Replace(dHora, ":" & nTmp, "")
'        mHora = NulosN(dHora) * 3600
        
        '**********************************
        '--de la hora
        If dHora <> "" Then
            nTmp = dHora
            PosSeparador = InStr(nTmp, ":")
            
            mHora = F_NulosN(Mid(nTmp, 1, PosSeparador - 1)) * 3600
            
            '--del minuto
            If nTmp <> "" Then
                nTmp = Mid(nTmp, PosSeparador + 1)
                PosSeparador = InStr(nTmp, ":")
                mMin = F_NulosN(Mid(nTmp, 1, PosSeparador - 1)) * 60
            End If
            '--del segundo
            If nTmp <> "" Then
                nTmp = Mid(nTmp, PosSeparador + 1)
                mSeg = F_NulosN(nTmp)
            End If
        End If
        '**********************************
    Else
        mHora = Hour(dHora) * 3600
        mMin = Minute(dHora) * 60
        mSeg = Second(dHora)
    End If
    
    ConvertSeg = mHora + mMin + mSeg
    
End Function

Public Function ConvertHora(mTotalSegundos) As String
    '--esta funcion devuelve la hora segun la cantidad de segundos como parametro
    Dim mTmp&
    Dim mHora As Integer
    Dim mMin As Integer
    Dim mSeg As Integer
    
    mTmp = mTotalSegundos
    mHora = (mTotalSegundos - (mTotalSegundos Mod 3600)) / 3600
    
    mTmp = mTotalSegundos Mod 3600
    
    mMin = (mTmp - (mTmp Mod 60)) / 60
    mSeg = mTmp Mod 60
    
    ConvertHora = Format(mHora, "00") & ":" & Format(mMin, "00") & ":" & Format(mSeg, "00")
End Function

'*****************************************************************************************************
'* Nombre           : HoraDecimal
'* Tipo             : FUNCCION
'* Descripcion      : COMVIERTE LAS HORAS A SU VALOR DECIMAL
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    ----------------------------------------------------------------------------------
'*                    xHora     | String            |  ESPECIFICA LA HORA QUE SE DESEA CONVERTIR
'*
'* Devuelve         : EL VALOR DE LA HORA EN DECIMALES, EL VALOR DEVUELTO ES DE TIPO DOBLE
'* Creado           : 04/05/10
'* Comentarios      : No se usa el tipo date como parametro ya que la funcion convertira mas de 24
'*                    horas a decimales
'*****************************************************************************************************
Public Function HoraDecimal(xHora As String) As Double
    Dim xHoras As Double
    Dim xMinutos As Double
    
    If xHora <> "" Then
        'xHoras = Hour(xHora)
        xHoras = F_NulosN(Mid(xHora, 1, 2))
        'xMinutos = Minute(xHora)
        xMinutos = F_NulosN(Mid(xHora, 4, 2))
        
        If xMinutos <> 0 Then
            xMinutos = (xMinutos / 60)
        End If
        HoraDecimal = xHoras + xMinutos
    Else
        HoraDecimal = ""
    End If
End Function

'Public Function xHoraDecimal(dHoraIni As String, Optional dHoraFin As String) As Double
'
'    Dim mHora As Double
'    Dim mMin As Double
'    Dim mSeg As Double
'    Dim nTmp As String
'    Dim dHora As Date
'
'    If dHoraFin <> "" Then
'        '--verificar que las horas sean correctas
'        If IsDate(dHoraIni) = False Or IsDate(dHoraFin) = False Then
'            ConvertHoraFaccion = 0
'            Exit Function
'        End If
'
'        If CDate(dHoraIni) < CDate(dHoraFin) Then
'            '--restar horas cuando esten en el mismo dia
'            dHora = TimeValue(dHoraIni) - TimeValue(dHoraFin)
'        Else
'            '--si la hora final es del siguiente dia
'            dHora = CDate((TimeValue(dHoraFin) - TimeValue("00:00")) + (TimeValue("23:59") - TimeValue(dHoraIni)) + CDate("00:01"))
'        End If
'    Else
'        '--si la hora inicial es falso entonces salir
'        If IsDate(dHoraIni) = False Then Exit Function
'        dHora = TimeValue(dHoraIni)
'    End If
'
'
'    mHora = Hour(dHora)
'    mMin = Minute(dHora)
'    mMin = IIf(mMin = 0, 0, mMin / 60)
'
'    ConvertHoraFaccion = mHora + mMin
'
'End Function

