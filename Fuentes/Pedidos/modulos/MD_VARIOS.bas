Attribute VB_Name = "MD_VARIOS"
Public Enum e_ARR_XX
'--SE USA ESTOS VALORES CON LA FINALIDAD DE SER UNICOS
    X_MES = 0
    X_TRIMESTRE = 1
    X_SEMESTRE = 2
End Enum


Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Sub Llenar_Anyo(obj As Object, Optional Todos As Boolean = False)
    '--obj:: listbox,combobox
    On Error Resume Next
    With obj
        .Clear
        If Todos = True Then
         .AddItem "Todos los Años"
        End If
        For i = Year(Date) To 2001 Step -1
            .AddItem i
        Next i
        .ListIndex = 0
    End With
End Sub


Public Sub Llenar_Trimestre(obj As Object, Optional Todos As Boolean = False)
    With obj
        .Clear
        If Todos = True Then
         .AddItem "Todos los Trimestres"
        End If
        .AddItem "Enero - Marzo"
        .AddItem "Abril - Junio"
        .AddItem "Julio - Septiembre"
        .AddItem "Octubre - Diciembre"
        .ListIndex = 0
    End With
End Sub

Public Sub Llenar_Semestre(obj As Object, Optional Todos As Boolean = False)
    With obj
        .Clear
        If Todos = True Then
         .AddItem "Todos los Semestres"
        End If
        .AddItem "1er. Semestre"
        .AddItem "2do. Semestre"
        .ListIndex = 0
    End With
End Sub

Public Sub Llenar_Mes(obj As Object, Optional Todos As Boolean = False)
    With obj
        .Clear
        If Todos = True Then
            .AddItem "Todos los Meses"
        End If
        .AddItem "Enero"
        .AddItem "Febrero"
        .AddItem "Marzo"
        .AddItem "Abril"
        .AddItem "Mayo"
        .AddItem "Junio"
        .AddItem "Julio"
        .AddItem "Agosto"
        .AddItem "Septiembre"
        .AddItem "Octubre"
        .AddItem "Noviembre"
        .AddItem "Diciembre"
       ''''''''''''''''
        If Todos = True Then
            .ListIndex = Month(Date)
        Else
            .ListIndex = Month(Date) - 1
        End If
        
        
    End With
End Sub

Public Sub ls_activar_chek(ls As ListBox, Optional N_VALOR As String = "")
    '--SELECCIONARA TODOS LOS ITEM'S DE LA LISTA
    On Error Resume Next
    Dim k As Integer
    For k = ls.ListCount - 1 To 0 Step -1
        ls.ListIndex = k
        If N_VALOR <> "" Then
            If ls.Text = N_VALOR Then
                ls.Selected(k) = True
                Exit For
            End If
        Else
            ls.Selected(k) = True
        End If
        
    Next
    Err.Clear
End Sub


Function ArchivoExiste(PATH As String) As Boolean
    On Error GoTo Fallo
    x = GetAttr(PATH)
    ArchivoExiste = True
    Exit Function
Fallo:
    ArchivoExiste = False
End Function



Function EstaAbiertoPrograma(ClassName As String) As Boolean
    If FindWindow(ClassName, vbNullString) Then
        EstaAbiertoPrograma = True
    End If
End Function


Public Sub abrir_shape(Rst As ADODB.Recordset, sql As String)
    Dim bdshape As New ADODB.Connection
    'Dim Path As String
    Dim dsn As String
    Set bdshape = New ADODB.Connection
    bdshape.Provider = "msdatashape"
    
    
    bdshape.ConnectionString = xCon.ConnectionString
    bdshape.Open
    Rst.CursorLocation = adUseClient
    Rst.Open sql, bdshape, adOpenDynamic, adLockBatchOptimistic
    Set bdshape = Nothing
End Sub


Sub LimpiaText(txt As Variant, Optional Limpiar_Tag As Boolean = False)
    On Error Resume Next
    Dim obj As Variant
    For Each obj In txt
      obj.Text = ""
      obj.ListIndex = -1
      If Limpiar_Tag = True Then
        obj.Tag = ""
      End If
    Next
    Err.Clear
End Sub

Public Function validar_numero(ascii As Integer) As Boolean
    Dim car As String
    car = Chr(ascii)
    If (car >= "0" And car <= "9") Or ascii = 8 Or ascii = 13 Then
     validar_numero = True 'devuelve true si es numero
    Else
     validar_numero = band 'devuelve falso si no es numero
    End If
End Function

Public Function validar_letras(ascii As Integer) As Boolean
    Dim car As String
    car = UCase(Chr(ascii))
    If (car >= "A" And car <= "Z") Or ascii = 13 Or ascii = 32 Or ascii = 8 Or ascii = 241 Or ascii = 209 Then
     validar_letras = True
    Else
     validar_letras = band
    End If
End Function

Sub habilitar(txt As Variant, band As Boolean)
    Dim obj As Variant
    For Each obj In txt
      obj.Enabled = band
    Next
End Sub

Public Sub Ocultar(txt As Variant, band As Boolean)
    On Error Resume Next
    Dim obj As Variant
    For Each obj In txt
        obj.Visible = band
    Next
    Err.Clear
End Sub


Function validar_Combo(txt As Variant) As Integer
    Dim obj As Variant
    
    For Each obj In txt
        If Trim(obj.Text) = "" Then
            validar_Combo = obj.Index
            Exit Function
        End If
    Next
    
    validar_Combo = -1

End Function


Sub LlenarCombo(Rst As ADODB.Recordset, cb, Campo As String, Optional Min As Boolean = False)
    cb.Clear
    If Rst.RecordCount > 0 Then Rst.MoveFirst
    While Not Rst.EOF
     If Min = False Then
        cb.AddItem Rst.Fields(Campo) & ""
     ElseIf Min = True Then
        cb.AddItem StrConv(Rst.Fields(Campo), 3)
     End If
     Rst.MoveNext
    Wend
End Sub




Sub CentrarFrm(frm As Form)
    On Error Resume Next
    If frm.WindowState <> 2 Then
    frm.Left = (Screen.Width - frm.Width) / 2
    frm.Top = (Screen.Height - frm.Height) / 2
    ' frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 2
    End If
    Err.Clear
End Sub


Sub OPEN_CONEX_TMP(X_CONEX As ADODB.Connection, RUTA_BD As String)
    AP_RUTASY = RutaSY
    AP_RUTABD = RutaBD
    AP_RUTABM = RutaBM

    SeEjecutoEmp = False
    Dim xFun As New eps_librerias.FuncionesData
    
    xFun.F_BASEDATOS = RUTA_BD
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "sigpyme.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    
    Set X_CONEX = xFun.AbrirConeccion
    Set xFun = Nothing
   
End Sub


Sub DEFINIR_RST_TMP(RST_TEMPORAL As ADODB.Recordset, RST_ORIGEN As ADODB.Recordset, Optional val As String = "")
    Dim vCampos As Long
    
    vCampos = RST_ORIGEN.Fields.Count
   Set RST_TEMPORAL = Nothing
   Set RST_TEMPORAL = New ADODB.Recordset
   
   For i = 0 To vCampos - 1
       'RST_TEMPORAL.Fields.Append rst_origen.Fields(i).Name, rst_origen.Fields(i).Type, rst_origen.Fields(i).DefinedSize, adFldIsNullable
       RST_TEMPORAL.Fields.Append RST_ORIGEN.Fields(i).Name, RST_ORIGEN.Fields(i).Type, -1, adFldIsNullable
   Next
    '-SI ES RST_TEMPORAL ES PARA SELECIONAR ADICIONAR EL CAMPO VAL,
   If val <> "" Then RST_TEMPORAL.Fields.Append val, adVarChar, 2, adFldIsNullable
   RST_TEMPORAL.Open

End Sub



Sub CARGAR_RST_TMP(RST_TEMPORAL As ADODB.Recordset, _
                RST_ORIGEN As ADODB.Recordset, _
                Optional val As String = "", _
                Optional Valor_Val As Integer = 0, _
                Optional F_UNSOLO_REGISTRO As Boolean = False)
                
    Dim N_CAMPO As String
    
    If F_UNSOLO_REGISTRO = False And RST_ORIGEN.RecordCount > 0 Then RST_ORIGEN.MoveFirst
    While Not RST_ORIGEN.EOF
        RST_TEMPORAL.AddNew
        
        For i = 0 To RST_ORIGEN.Fields.Count - 1
            N_CAMPO = RST_TEMPORAL.Fields(i).Name
            If RST_ORIGEN.Fields(N_CAMPO) & "" <> "" Then
               RST_TEMPORAL.Fields(N_CAMPO) = RST_ORIGEN.Fields(N_CAMPO) & ""
            Else
            End If
        Next
        
        If val <> "" Then
           RST_TEMPORAL.Fields(val) = CStr(Valor_Val)
        End If
        If F_UNSOLO_REGISTRO = True Then Exit Sub
        RST_ORIGEN.MoveNext
    Wend
End Sub


Sub CARGAR_ARR_XX(ARR_XX() As String, T_TIPO As e_ARR_XX)
If T_TIPO = X_MES Then
    ReDim ARR_XX(11, 2)
    ARR_XX(0, 0) = "Enero":        ARR_XX(0, 1) = "Ene":          ARR_XX(0, 2) = "1"
    ARR_XX(1, 0) = "Febrero":      ARR_XX(1, 1) = "Feb":          ARR_XX(1, 2) = "2"
    ARR_XX(2, 0) = "Marzo":        ARR_XX(2, 1) = "Mar":          ARR_XX(2, 2) = "3"
    ARR_XX(3, 0) = "Abril":        ARR_XX(3, 1) = "Abr":          ARR_XX(3, 2) = "4"
    ARR_XX(4, 0) = "Mayo":         ARR_XX(4, 1) = "May":          ARR_XX(4, 2) = "5"
    ARR_XX(5, 0) = "Junio":        ARR_XX(5, 1) = "Jun":          ARR_XX(5, 2) = "6"
    ARR_XX(6, 0) = "Julio":        ARR_XX(6, 1) = "Jul":          ARR_XX(6, 2) = "7"
    ARR_XX(7, 0) = "Agosto":       ARR_XX(7, 1) = "Ago":          ARR_XX(7, 2) = "8"
    ARR_XX(8, 0) = "Septiembre":   ARR_XX(8, 1) = "Sep":          ARR_XX(8, 2) = "9"
    ARR_XX(9, 0) = "Octubre":      ARR_XX(9, 1) = "Oct":          ARR_XX(9, 2) = "10"
    ARR_XX(10, 0) = "Noviembre":   ARR_XX(10, 1) = "Nov":         ARR_XX(10, 2) = "11"
    ARR_XX(11, 0) = "Diciembre":   ARR_XX(11, 1) = "Dic":         ARR_XX(11, 2) = "12"
ElseIf T_TIPO = X_TRIMESTRE Then
    ReDim ARR_XX(3, 2)
    ARR_XX(0, 0) = "Enero - Marzo":         ARR_XX(0, 1) = "Ene-Mar":          ARR_XX(0, 2) = "1"
    ARR_XX(1, 0) = "Abril - Junio":         ARR_XX(1, 1) = "Abr-Jun":          ARR_XX(1, 2) = "2"
    ARR_XX(2, 0) = "Julio - Septiembre":    ARR_XX(2, 1) = "Jul-Sep":          ARR_XX(2, 2) = "3"
    ARR_XX(3, 0) = "Octubre - Diciembre":   ARR_XX(3, 1) = "Oct-Dic":          ARR_XX(3, 2) = "4"
ElseIf T_TIPO = X_SEMESTRE Then
    ReDim ARR_XX(1, 2)
    ARR_XX(0, 0) = "Enero - Junio":         ARR_XX(0, 1) = "1er. Sem":         ARR_XX(0, 2) = "1"
    ARR_XX(1, 0) = "Julio - Diciembre":     ARR_XX(1, 1) = "2do. Sem":         ARR_XX(1, 2) = "2"
End If

End Sub

       
        

