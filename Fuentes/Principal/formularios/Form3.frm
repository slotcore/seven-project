VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   LinkTopic       =   "Form3"
   ScaleHeight     =   1815
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Left            =   345
      TabIndex        =   1
      Top             =   180
      Width           =   1515
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   435
      TabIndex        =   0
      Top             =   930
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim RstCajCab As New ADODB.Recordset
    Dim RstCajDet As New ADODB.Recordset
    Dim RstCajOri As New ADODB.Recordset
    Dim RstSEVENDia As New ADODB.Recordset
    
    Dim RstSGICab As New ADODB.Recordset
    Dim RstSGIDet As New ADODB.Recordset
    Dim RstSGIDia As New ADODB.Recordset
    Dim RstSGIDiario As New ADODB.Recordset
    
    Dim A, B, C, xId, IdCuenta, xidCuenta, xIdOri, xIdEntGeneradora As Integer
    Dim xImporte As Double
    Dim xNumAsi As String
    
    RST_Busq RstCajCab, "SELECT * FROM con_cajabanco", xCon
    RST_Busq RstCajDet, "SELECT * FROM con_cajabancodet", xCon
    RST_Busq RstCajOri, "SELECT * FROM con_cajabancoorides", xCon
    RST_Busq RstSEVENDia, "SELECT * FROM con_diario", xCon
    'RST_Busq RstSGICab, "SELECT CON_Caja_Banco.* From CON_Caja_Banco WHERE (((CON_Caja_Banco.RegCajaBanco) Like '01%'))", xCon
    'RST_Busq RstSGICab, "SELECT CON_Caja_Banco.* From CON_Caja_Banco WHERE (((CON_Caja_Banco.RegCajaBanco) Like '01%') AND ((CON_Caja_Banco.Operacion)='Depósito'))", xCon
    'RST_Busq RstSGICab, "SELECT CON_Caja_Banco.* From CON_Caja_Banco WHERE (((CON_Caja_Banco.RegCajaBanco) Like '01%') " _
        & " AND ((CON_Caja_Banco.Operacion)='Depósito') AND ((CON_Caja_Banco.Descripcion)='Transaccion con Personal'))", xCon

    ''***********
    ''CAMBIAR MES
    RST_Busq RstSGICab, "SELECT CON_Caja_Banco.* From CON_Caja_Banco WHERE (((CON_Caja_Banco.RegCajaBanco) Like '07%'))", xCon

    If RstSGICab.RecordCount <> 0 Then
        RstSGICab.MoveFirst
        For A = 1 To RstSGICab.RecordCount
            xId = HallaCodigoTabla("con_cajabanco", xCon, "id")
            ''***********
            ''CAMBIAR MES
            xNumAsi = "07" + NuevoNumAsiento(6, 7, xCon)
            
            RstCajCab.AddNew
            RstCajCab("id") = xId
            RstCajCab("numreg") = xNumAsi
            
            If RstSGICab("operacion") = "Retiro" Then
                RstCajCab("tipmov") = "2"    'grabamos que es egreso
                'para cuando sean cuentas de fondo fijo o caja
                If RstSGICab("banco") = "10-2-02" Or RstSGICab("banco") = "10-2-01" Then
                    RstCajCab("tipope") = "1" 'ESPECIFICA QUE ES CAJA
                    If Val(RstSGICab("cod_moneda")) = 1 Then
                        RstCajCab("idori") = 13 '"14" 'le cargamos a fondo fijo mantenimiento SOLES (CAMBIAR PARA OTRA EMPRESA)
                    Else
                        RstCajCab("idori") = 24 '"18" 'le cargamos a fondo fijo mantenimiento Dolares (CAMBIAR PARA OTRA EMPRESA)
                    End If
                    RstCajCab("iddoc") = "1"
                End If
                
                'para cuando sean cuentas de banco
                If RstSGICab("banco") = "10-4-01-01" Then
                    RstCajCab("tipope") = "2" 'ESPECIFICA QUE ES BANCO
                    IdCuenta = Busca_Codigo("10-4-01-01", "cuenta", "id", "con_planctas", "C", xCon)
                    RstCajCab("idcueban") = Busca_Codigo(IdCuenta, "idcuen", "id", "con_bancocuenta", "N", xCon)
                    RstCajCab("iddoc") = "5"
                    RstCajCab("idmedpag") = "7"
                End If
                If RstSGICab("banco") = "10-4-01-02" Then
                    RstCajCab("tipope") = "2"
                    IdCuenta = Busca_Codigo("10-4-01-02", "cuenta", "id", "con_planctas", "C", xCon)
                    RstCajCab("idcueban") = Busca_Codigo(IdCuenta, "idcuen", "id", "con_bancocuenta", "N", xCon)
                    RstCajCab("iddoc") = "5"
                    RstCajCab("idmedpag") = "7"
                End If
            End If
            
            If RstSGICab("operacion") = "Depósito" Then
                RstCajCab("tipmov") = "1"  'grabamos que es ingreso
                
                If RstSGICab("banco") = "10-2-02" Or RstSGICab("banco") = "10-2-01" Then
                    RstCajCab("tipope") = "1" 'ESPECIFICA QUE ES CAJA
                    'el destino se especifica en funcion a la moneda cuandoe es caja
                    If Val(RstSGICab("cod_moneda")) = 1 Then
                        If RstSGICab("descripcion") = "Transaccion con Personal" Then
                            RstCajCab("idori") = 33  'Fondo Fijo Administracion General MN
                        Else
                            RstCajCab("idori") = 1  'Clientes moneda nacional
                        End If
                    Else
                        If RstSGICab("descripcion") = "Transaccion con Personal" Then
                            RstCajCab("idori") = 34  'Fondo Fijo Administracion General ME
                        Else
                            RstCajCab("idori") = 23 'clientes moneda extranjera
                        End If
                    End If
                    RstCajCab("iddoc") = "1"
                Else
                    RstCajCab("tipope") = "2" 'ESPECIFICA QUE ES BANCO
                    If Val(RstSGICab("cod_moneda")) = 1 Then
                        If RstSGICab("descripcion") = "Transaccion con Personal" Then
                            RstCajCab("idori") = 33 'Fondo Fijo Administracion General MN
                        Else
                            RstCajCab("idori") = 1  'Clientes moneda nacional
                        End If
                    Else
                        If RstSGICab("descripcion") = "Transaccion con Personal" Then
                            RstCajCab("idori") = 34 'Fondo Fijo Administracion General MN
                        Else
                            RstCajCab("idori") = 23  'Clientes moneda extranjero
                        End If
                    End If
                    If RstSGICab("banco") = "10-4-01-01" Then
                        IdCuenta = Busca_Codigo("10-4-01-01", "cuenta", "id", "con_planctas", "C", xCon)
                    End If
                    If RstSGICab("banco") = "10-4-01-02" Then
                        IdCuenta = Busca_Codigo("10-4-01-02", "cuenta", "id", "con_planctas", "C", xCon)
                    End If
                    RstCajCab("idcueban") = Busca_Codigo(IdCuenta, "idcuen", "id", "con_bancocuenta", "N", xCon)
                    RstCajCab("iddoc") = "8"
                    RstCajCab("idmedpag") = "1"
                End If
            End If
            
            RstCajCab("idmon") = Val(RstSGICab("cod_moneda"))
            RstCajCab("importe") = RstSGICab("importe")
            RstCajCab("numdoc") = RstSGICab("nro_documento")
            RstCajCab("fchreg") = RstSGICab("fecharegistro")
            RstCajCab("fchope") = RstSGICab("fecha")
            RstCajCab("saldo") = RstSGICab("importe")
            RstCajCab.Update
            
            RST_Busq RstSGIDet, "SELECT CON_Detalle_CajaBanco.* From CON_Detalle_CajaBanco WHERE (((CON_Detalle_CajaBanco.RegCajaBanco)='" & RstSGICab("regcajabanco") & "'))", xCon

            'grabamos el detalle de la operacion de caja y bancos
            If RstSGICab("descripcion") <> "Transaccion con Personal" Then
                If RstSGIDet.RecordCount <> 0 Then
                    RstSGIDet.MoveFirst
                    
                    For B = 1 To RstSGIDet.RecordCount
                        If RstSGICab("operacion") = "Retiro" Then
                            'Proveedores
                            If Val(RstSGICab("cod_moneda")) = "001" Then
                                xIdOri = 11 '12
                            Else
                                xIdOri = 12 '13
                            End If
                            xIdEntGeneradora = 1
                            xidCuenta = 175  'FACTURAS POR PAGAR
                        Else
                            'Clientes
                            If RstSGICab("banco") = "10-4-01-01" Or RstSGICab("banco") = "10-4-01-02" Then
                                If Val(RstSGICab("cod_moneda")) = 1 Then
                                    xIdOri = 24
                                Else
                                    xIdOri = 25
                                End If
                            Else
                                If Val(RstSGICab("cod_moneda")) = 1 Then
                                    xIdOri = 41
                                Else
                                    xIdOri = 42
                                End If
                            End If
                            xIdEntGeneradora = 4
                            xidCuenta = 26 ' facturas por cobrar
                        End If
                        RstCajDet.AddNew
                        RstCajDet("id") = xId
                        RstCajDet("idori") = xIdOri
                        If RstSGICab("operacion") = "Retiro" Then
                            RstCajDet("iddoc") = NulosN(Busca_Codigo(RstSGIDet("nro_documento"), "nro_documento", "iddoc", "con_documentos_pagar", "C", xCon))
                            If RstCajDet("iddoc") = 0 Then
                                MsgBox "La operacion Nº " + Trim(RstSGICab("regcajabanco")) + " Del SGI, tiene una numero de documento que no existe en el SEVEN " + Trim(RstSGIDet("nro_documento")), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                            End If
                        Else
                            RstCajDet("iddoc") = NulosN(Busca_Codigo(RstSGIDet("nro_documento"), "nro_documento", "iddoc", "con_documentos_cobrar", "C", xCon))
                        End If
                        
                        RstCajDet("idorigen") = xIdEntGeneradora
                        RstCajDet("impabo") = RstSGIDet("monto")
                        RstCajDet("salant") = RstSGIDet("monto")
                        RstCajDet("idcue") = xidCuenta
                        RstCajDet.Update
                        
                        RstSGIDet.MoveNext
                        If RstSGIDet.EOF = True Then Exit For
                    Next B
                End If
            End If
                    
            'Grabamos
            RstCajOri.AddNew
            RstCajOri("id") = xId
            If RstSGICab("operacion") = "Retiro" Then
                'destino de egresos
                If RstSGICab("descripcion") = "Transaccion con Personal" Then
                    If Val(RstSGICab("cod_moneda")) = "001" Then
                        RstCajOri("idorides") = 23 '29  'para moneda soles
                    Else
                        RstCajOri("idorides") = 29 '29  'para moneda dolares
                    End If
                Else
                    If Val(RstSGICab("cod_moneda")) = "001" Then
                        RstCajOri("idorides") = 11 '12 facturas a proveedore mone nacional
                    Else
                        RstCajOri("idorides") = 12 '13 facturas a proveedore mone extranjera
                    End If
                End If
            Else
                'destino del ingreso
                If RstSGICab("banco") = "10-4-01-01" Or RstSGICab("banco") = "10-4-01-02" Then
                    If Val(RstSGICab("cod_moneda")) = "001" Then
                        xIdOri = 24 '16 bcp soles
                    Else
                        xIdOri = 25 '17 bcp dolares
                    End If
                Else
                    If Val(RstSGICab("cod_moneda")) = "001" Then
                        xIdOri = 5 '7 fondo fijo mantenimiento MN
                    Else
                        xIdOri = 43 '0 fondo fijo mantenimiento ME
                    End If
                End If
                RstCajOri("idorides") = xIdOri
            End If
            RstCajOri("importe") = RstSGICab("importe")
            RstCajOri.Update
            
            'Grabamos el diario de la operacion
            Set RstSGIDiario = Nothing
            RST_Busq RstSGIDiario, "SELECT CON_Diario_General.*, CON_Diario_General.Libro From CON_Diario_General " _
                & " WHERE (((CON_Diario_General.Nro_Libro)='" & RstSGICab("regcajabanco") & "') AND ((CON_Diario_General.Libro)='Caja Fondo Fijo'))", xCon

            If RstSGIDiario.RecordCount <> 0 Then
                RstSGIDiario.MoveFirst
                For B = 1 To RstSGIDiario.RecordCount
                    RstSEVENDia.AddNew
                    RstSEVENDia("año") = 2007
                    
                    ''***********
                    ''CAMBIAR MES
                    RstSEVENDia("idmes") = 7
                    RstSEVENDia("idlib") = 6
                    RstSEVENDia("idmov") = xId
                    RstSEVENDia("idcue") = Busca_Codigo(RstSGIDiario("cuenta"), "cuenta", "id", "con_planctas", "C", xCon)
                    'RstSEVENDia("iddocpro") = RstSGIDiario()
                    RstSEVENDia("numasi") = Mid(xNumAsi, 3, 4)
                    
                    If RstSGIDiario("cod_moneda") = "001" Then
                        RstSEVENDia("impdebsol") = RstSGIDiario("monto_debe")
                        RstSEVENDia("imphabsol") = RstSGIDiario("monto_haber")
                    Else
                        RstSEVENDia("impdebdol") = RstSGIDiario("monto_debe")
                        RstSEVENDia("imphabdol") = RstSGIDiario("monto_haber")
                    End If
                    
                    ''***********
                    ''CAMBIAR MES
                    RstSEVENDia("fchasi") = CDate("01/07/07")
                    RstSEVENDia("fchdoc") = RstSGICab("fecha")
                    RstSEVENDia.Update
                    
                    RstSGIDiario.MoveNext
                    If RstSGIDiario.EOF = True Then Exit For
                Next B
            End If
            
            RstSGICab.MoveNext
            If RstSGICab.EOF = True Then
                Exit For
            End If
        Next A
    End If
    MsgBox "El proceso de importacion de datos termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Sub

