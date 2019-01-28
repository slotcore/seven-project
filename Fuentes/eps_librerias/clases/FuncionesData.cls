VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuncionesData"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public F_PROVEEDOR As String
Public F_BASEDATOS As String
Public F_USUARIO As String
Public F_PASSWORD As String
Public F_GRUPOTRABAJO As String

Public SQL_DATASOURCE As String
Public SQL_CATALOGO As String
Public SQL_DATA As String

Function AbrirConeccion() As ADODB.Connection
    Dim ConeccOpen As New ADODB.Connection
    Dim xCadConeccion As String
    xCadConeccion = "Provider=" + Trim(F_PROVEEDOR) _
                    & ";Password=" & Trim(F_PASSWORD) & ";Persist Security Info=true" _
                    & ";User ID=" & Trim(F_USUARIO) & ";Data Source=" + Trim(F_BASEDATOS) _
                    & ";Jet OLEDB:System database=" + Trim(F_GRUPOTRABAJO)
    
    ConeccOpen.ConnectionString = xCadConeccion
    ConeccOpen.Open
    Set AbrirConeccion = ConeccOpen
End Function

Function AbrirConeccionSQL() As ADODB.Connection
    Dim ConeccOpen As New ADODB.Connection
    Dim xCadConeccion As String
    
'    xCadConeccion = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog='" & SQL_CATALOGO & "';" _
        & " Data Source=" & SQL_DATASOURCE & ";Initial File Name='" & SQL_DATA & "'"
    
    xCadConeccion = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;" _
        & " Data Source=" & SQL_DATASOURCE & ";Initial File Name='" & SQL_DATA & "'"
    
    'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;                    Data Source=PC10\SQLEXPRESS;Initial File Name=F:\proyectos\SAC\data\2009\0002\data.mdf
    'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=''; Data Source='PC10\SQLEXPRESS';Initial File Name='F:\PROYECTOS\SAC\DATA\2009\0002\data.mdb'
    ConeccOpen.ConnectionString = xCadConeccion
    ConeccOpen.Open
    Set AbrirConeccionSQL = ConeccOpen
End Function

Sub RST_Busq(rstBusq As ADODB.Recordset, TxtSQLoTabla As String, xConeccion As Connection)
    F_RST_Busq rstBusq, TxtSQLoTabla, xConeccion
End Sub

Sub RST_Mant(rstMant As ADODB.Recordset, TxtSQLoTabla As String, xConeccion As Connection)
    F_RST_Mant rstMant, TxtSQLoTabla, xConeccion
End Sub

Function Busca_Codigo(valor_busca As Variant, campo_codigo As String, Campo_descripcion As String, Tabla As String, Tipo As String, xConeccion As ADODB.Connection)
    Dim RstBusca As New ADODB.Recordset
    If Tipo = "C" Then
        RST_Busq RstBusca, "Select " & campo_codigo & ", " & Campo_descripcion & "  from  " & Tabla & " where " & campo_codigo & " = '" & valor_busca & "'", xConeccion
    End If
    If Tipo = "N" Then
        RST_Busq RstBusca, "Select " & campo_codigo & ", " & Campo_descripcion & "  from  " & Tabla & " where " & campo_codigo & " = " & valor_busca & "", xConeccion
    End If
    If RstBusca.State = 0 Then
        
    End If
    If RstBusca.RecordCount = 0 Then
        Busca_Codigo = ""
    Else
        Busca_Codigo = F_NulosC(RstBusca(Campo_descripcion))
    End If
End Function

Function HallaCodigoTabla(zTabla As String, xConeccion As ADODB.Connection, Optional CampoId As String) As Double
   '--Modificado 28/10/11 Johan Castro
    Dim xRstTabla As New ADODB.Recordset
'''''    RST_Busq xRstTabla, "SELECT * FROM " & zTabla & " ORDER BY  " & CampoId & " ", xConeccion
    RST_Busq xRstTabla, "SELECT TOP 1 " & CampoId & " FROM " & zTabla & " ORDER BY  " & CampoId & " DESC ", xConeccion
    If xRstTabla.RecordCount = 0 Then
        HallaCodigoTabla = 1
    Else
        xRstTabla.MoveLast
        HallaCodigoTabla = F_NulosN(xRstTabla(CampoId)) + 1
    End If
    Set xRstTabla = Nothing
End Function

Function CrearRstTMP(xCampos() As String) As Recordset
    'xcampos(0,0) = nombre del campo
    'xcampos(0,1) = tipo del campo
    'xcampos(0,2) = tamaño del campo
    
    Dim RercorsetTemporal As New ADODB.Recordset
    Dim xr As Integer
    Dim xc As Integer
    Dim A As Integer
    xc = RercorsetTemporal.Fields.Count - 1
    For xr = xc To 0 Step -1
        RercorsetTemporal.Fields.Delete xr
    Next xr
    
    For A = LBound(xCampos) To UBound(xCampos)
        If xCampos(A, 1) = "C" Then  'Caracter
            RercorsetTemporal.Fields.Append xCampos(A, 0), adVarChar, Val(xCampos(A, 2))
        End If
        If xCampos(A, 1) = "F" Then  'Fecha
            RercorsetTemporal.Fields.Append xCampos(A, 0), adDate ', 50, , adFldIsNullable
        End If
        If xCampos(A, 1) = "N" Then  'Numerico
            RercorsetTemporal.Fields.Append xCampos(A, 0), adInteger ', Val(xCampos(A, 2))
        End If
        If xCampos(A, 1) = "D" Then  'Doble
            RercorsetTemporal.Fields.Append xCampos(A, 0), adDouble ', Val(xCampos(A, 2))
        End If
        If xCampos(A, 1) = "L" Then  'Logico
            RercorsetTemporal.Fields.Append xCampos(A, 0), adBoolean ', Val(xCampos(A, 2))
        End If
    Next A
    
    Set CrearRstTMP = RercorsetTemporal
    Set RercorsetTemporal = Nothing
End Function

Function EscribirNuevoRegistro(xCampos() As String, Tabla As String, Coneccion As ADODB.Connection) As Boolean
    'Columna    | Descripcion
    '------------------------
    '0          | campo
    '1          | Valor
    '2          | requerido
    '3          | tipo
    '4          | msj error
    
On Error GoTo LaCague

    Dim xCadCampos, xCadValores As String
    Dim A As Integer
    Dim B As Integer
    
    'verificamos el los campos esten completos
    For A = LBound(xCampos) To UBound(xCampos)
        ' pregutamos si el capo es requerido
        If F_NulosC(xCampos(A, 2)) = "S" Then
            'If F_NulosC(xCampos(A, 3)) = "N" Then
                If F_NulosC(xCampos(A, 1)) = "" Then
                    MsgBox F_NulosC(xCampos(A, 4)), vbInformation + vbOKOnly + vbDefaultButton1, "Ingreso de Datos"
                    EscribirNuevoRegistro = False
                    Exit Function
                End If
            'End If
        End If
    Next A
    
    xCadCampos = ""
    xCadCampos = xCadCampos + "("
    For A = LBound(xCampos) To UBound(xCampos)
        xCadCampos = xCadCampos & xCampos(A, 0)
        If A <= UBound(xCampos) - 1 Then
            xCadCampos = xCadCampos & ", "
        End If
    Next A
    xCadCampos = xCadCampos & ")"
    
    xCadValores = ""
    
    For A = LBound(xCampos) To UBound(xCampos)
        If UCase(xCampos(A, 3)) = "C" Or xCampos(A, 3) = "M" Then
            xCadValores = xCadValores & "'" & F_NulosC(xCampos(A, 1)) & "'" & " AS x" & xCampos(A, 0)
        End If
        If UCase(xCampos(A, 3)) = "N" Or UCase(xCampos(A, 3)) = "L" Then
            xCadValores = xCadValores & F_NulosN(xCampos(A, 1)) & " AS x" & xCampos(A, 0)
        End If
        
        If UCase(xCampos(A, 3)) = "F" Or UCase(xCampos(A, 3)) = "H" Then
            If F_NulosC(xCampos(A, 1)) = "" Then
                xCadValores = xCadValores & "null AS x" & xCampos(A, 0)
            Else
                xCadValores = xCadValores & "cdate('" & F_NulosC(xCampos(A, 1)) & "')" & " AS x" & xCampos(A, 0)
            End If
        End If
        
        If A <= UBound(xCampos) - 1 Then
            xCadValores = xCadValores & ", "
        End If
    Next A
    
    Coneccion.Execute "INSERT INTO " & Tabla & " " & " " & xCadCampos & " " _
        & " SELECT " & xCadValores
    
    EscribirNuevoRegistro = True
    Exit Function
    
LaCague:
    'Resume
    MsgBox "No se pudo guardar el registro por el siguiente motivo : " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    MsgBox "INSERT INTO " & Tabla & " " & " " & xCadCampos & " " & " SELECT " & xCadValores
    EscribirNuevoRegistro = False
End Function

Function ModificarRegistro(xCampos() As String, Tabla As String, Coneccion As ADODB.Connection) As Boolean
'--Fecha Modificacion: 28/10/11 Enrique Pollongo

    'Columna    | Descripcion
    '------------------------
    '0          | campo
    '1          | Valor
    '2          | requerido
    '3          | tipo
    '4          | msj error
    
On Error GoTo LaCague

    Dim xCadCampos, xCadValores As String
    Dim xCadWhere As String
    Dim A As Integer
    Dim B As Integer
    
    'verificamos el los campos esten completos
    For A = LBound(xCampos) To UBound(xCampos)
        ' pregutamos si el capo es requerido
        If F_NulosC(xCampos(A, 2)) = "S" Then
            'If F_NulosC(xCampos(A, 3)) = "N" Then
                If F_NulosC(xCampos(A, 1)) = "" Then
                    MsgBox F_NulosC(xCampos(A, 4)), vbInformation + vbOKOnly + vbDefaultButton1, "Ingreso de Datos"
                    ModificarRegistro = False
                    Exit Function
                End If
            'End If
        End If
        
    Next A
    
    xCadCampos = ""
    xCadWhere = ""
    xCadCampos = xCadCampos & ""
    xCadWhere = xCadWhere & "("
    For A = LBound(xCampos) To UBound(xCampos)
        If xCampos(A, 5) <> "S" Then
            ' ARMAMOS LA CADENA DE ACTUALIZACION
            If UCase(xCampos(A, 3)) = "C" Or xCampos(A, 3) = "M" Then
                xCadCampos = xCadCampos & xCampos(A, 0) & "='" & xCampos(A, 1) & "'"
            End If
            
            If UCase(xCampos(A, 3)) = "N" Or UCase(xCampos(A, 3)) = "L" Then
                xCadCampos = xCadCampos & xCampos(A, 0) & "=" & F_NulosN(xCampos(A, 1))
            End If
            
            If UCase(xCampos(A, 3)) = "F" Or UCase(xCampos(A, 3)) = "H" Then
                If F_NulosC(xCampos(A, 1)) = "" Then
                    ' SI EL CAMPO ESTA VACIO
                    'xCadValores = xCadValores & "null AS x" & xCampos(A, 0)
                    xCadCampos = xCadCampos & xCampos(A, 0) & "= NULL"
                Else
                    'xCadValores = xCadValores & "cdate('" & F_NulosC(xCampos(A, 1)) & "')" & " AS x" & xCampos(A, 0)
                    xCadCampos = xCadCampos & xCampos(A, 0) & "=CDate('" & F_NulosC(xCampos(A, 1)) & "')"
                End If
            End If
            
            If A <= UBound(xCampos) - 1 Then
                xCadCampos = xCadCampos & ", "
            End If
        Else
            ' ARMAMOS LA CADENA WHERE
            If UCase(xCampos(A, 3)) = "C" Or xCampos(A, 3) = "M" Then
                xCadWhere = xCadWhere & xCampos(A, 0) & "='" & xCampos(A, 1) & "'"
            End If
            
            If UCase(xCampos(A, 3)) = "N" Or UCase(xCampos(A, 3)) = "L" Then
                xCadWhere = xCadWhere & xCampos(A, 0) & "=" & xCampos(A, 1)
            End If
        End If
    Next A
    xCadCampos = xCadCampos & ""
    xCadWhere = xCadWhere & ")"
    
    Coneccion.Execute "UPDATE " & Tabla & " SET " _
        & xCadCampos _
        & " WHERE " & xCadWhere
    
    ModificarRegistro = True
    Exit Function
    
LaCague:
    MsgBox "No se pudo guardar datos en la tabla : " & Tabla & " por el siguiente motivo : " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    ModificarRegistro = False
End Function


Function BuscaConCriterio(CadSql As String, Coneccion As ADODB.Connection) As ADODB.Recordset
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, CadSql, Coneccion
    Set BuscaConCriterio = Rst
    Set Rst = Nothing
End Function

