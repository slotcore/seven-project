VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmLevantar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas - Administrador de Ordenes CEN"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   3975
      TabIndex        =   3
      Top             =   2040
      Width           =   1725
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   2460
      Left            =   120
      TabIndex        =   1
      Top             =   195
      Width           =   3765
      _cx             =   6641
      _cy             =   4339
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   128
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmLevantar.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.CommandButton CmdLoadFile 
      Caption         =   "Cargar Ordenes "
      Height          =   615
      Left            =   3975
      TabIndex        =   0
      Top             =   210
      Width           =   1725
   End
   Begin VB.CommandButton CmdProce 
      Caption         =   "Procesar Ordenes"
      Height          =   615
      Left            =   3975
      TabIndex        =   2
      Top             =   855
      Width           =   1725
   End
End
Attribute VB_Name = "FrmLevantar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstCab As New ADODB.Recordset
Dim RstDet As New ADODB.Recordset
Dim xRutFol As String
Dim xTitulo As String

Private Sub CmdLoadFile_Click()
    Dim File, Folder, FileCollection
    Dim fso As New FileSystemObject
    
    
    xRutFol = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTACE", "RUTAS")
    
    Set Folder = fso.GetFolder(xRutFol)
    Set FileCollection = Folder.Files
    
    Fg1.Rows = 1
    For Each File In FileCollection
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = File.Name
    Next

    If Fg1.Rows = 1 Then
        MsgBox "No se ha descargado ninguna orden de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If

End Sub

Function AbriryGuardar(NomArchv As String) As Boolean
    '===================================================================================================
    'Creado :   /  /  Por: Enrique Pollongo
    'Propósito: Crear nuevo pedido
    '
    'Entradas:  NomArchv = Nombre del archivo del blok de notas
    '
    'Resultados: Pedido registrado en seven
    '
    'Modificado: 13/12/10 Por: Jose Chacon
    '           Aumento de instrucciones para nuevos campos de Base de Datos
    '           Verificacion de carga de Productos por Empresa
    '           19/01/11 Por Johan Castro
    '           Cambiar el tipo de dato a variable xIdPed a Double antes Integer
    '           Agregar lineas de codigo para registrar el historial del pedido
    'Modificado:10/11/11 Johan Castro
    '           Agregar condicion para obtener orden de compra en "ENC" para longitud = 72
    'Modificado:11/11/11 Enrique Pollongo
    '           agregar cadena ENC a un array con funcion Split() para obtener datos exactos de orden de compra
    '===================================================================================================


    Dim crlf$
    Dim file_data$
    crlf$ = Chr(13) & Chr(10)
    Dim xCad As String
    Dim xNumLin As Integer
    Dim xFchEmi, xFchEnt As String
    Dim xIdPed As Double
    Dim RstPunVen As New ADODB.Recordset
    'Se añade una referencia a la tabla ped_pedidodetent
    Dim RstDetEnt As New ADODB.Recordset
    Dim xNonPunVen As String
    
    Dim RstAuxiditem As New ADODB.Recordset
    
    RST_Busq RstCab, "SELECT * FROM ped_pedido", xCon
    RST_Busq RstDet, "SELECT * FROM ped_pedidodet", xCon
         
    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    Open Trim(NomArchv) For Input As #1
    
    xNumLin = 1
    
    RstCab.AddNew
    
    xIdPed = HallaCodigoTabla("ped_pedido", xCon, "id")
    RstCab("id") = xIdPed
    
    Dim xDato() As String
    
    While Not EOF(1)
        Line Input #1, file_data$
        
        xCad = file_data$
        xDato = Split(xCad, ",")
        If xNumLin <= 11 Then
'            If Mid(Trim(xCad), 1, 3) = "ENC" Then
'                If Len(Trim(xCad)) = 74 Then
'                    RstCab("numcen") = Mid(xCad, 45, 21)
'                    'Se llena la orden de compra a partir del numero Cen
'                    RstCab("oc") = Mid((Mid(xCad, 45, 21)), 11, 10)
'                End If
'                If Len(Trim(xCad)) = 75 Then
'                    RstCab("numcen") = Mid(xCad, 46, 21)
'                    'Se llena la orden de compra a partir del numero Cen
'                    RstCab("oc") = Mid((Mid(xCad, 46, 21)), 11, 10)
'                End If
'
'                If Len(Trim(xCad)) = 72 Then
'                    RstCab("numcen") = Mid(xCad, 43, 21)
'                    'Se llena la orden de compra a partir del numero Cen
'                    RstCab("oc") = Mid((Mid(xCad, 43, 21)), 11, 10)
'                End If
'
'            End If
            If Mid(Trim(xCad), 1, 3) = "ENC" Then
                RstCab("numcen") = xDato(5) 'Mid(xCad, 43, 21)
                'Se llena la orden de compra a partir del numero Cen
                RstCab("oc") = Mid(xDato(5), 11, 10)
            End If
            
            If Mid(Trim(xCad), 1, 3) = "DTM" Then
                xFchEmi = Mid(Trim(xCad), 11, 2) + "/" + Mid(Trim(xCad), 9, 2) + "/" + Mid(Trim(xCad), 5, 4)
                xFchEnt = Mid(Trim(xCad), 26, 2) + "/" + Mid(Trim(xCad), 24, 2) + "/" + Mid(Trim(xCad), 20, 4)
                RstCab("fchemi") = CDate(xFchEmi)
                RstCab("fchent") = CDate(xFchEnt)
            End If
            
            If Mid(Trim(xCad), 1, 4) = "BYOC" Then
                RstCab("codcen") = Mid(xCad, 6, 13)
                RST_Busq RstPunVen, "SELECT vta_puntoventa.idcli,vta_puntoventa.descripcion, vta_puntoventa.codcen, vta_puntoventa.id " _
                    & " From vta_puntoventa WHERE (((vta_puntoventa.codcen)='" & Mid(xCad, 22, 13) & "'))", xCon
                RstCab("idpunvecli") = RstPunVen("id")
                RstCab("idcli") = RstPunVen("idcli")
                xNonPunVen = RstPunVen("descripcion")
                Set RstPunVen = Nothing
            End If
            
            If Mid(Trim(xCad), 1, 4) = "DPGR" Then
                RST_Busq RstPunVen, "SELECT vta_puntoventa.descripcion, vta_puntoventa.codcen, vta_puntoventa.id " _
                    & " From vta_puntoventa WHERE (((vta_puntoventa.codcen)='" & Mid(xCad, 6, 13) & "'))", xCon
                RstCab("idlugent") = RstPunVen("id")
                Set RstPunVen = Nothing
            End If
        
            If Mid(Trim(xCad), 1, 6) = "TAXMOA" Then
                RstCab("impigv") = NulosN(ExtraerCad(xCad, ",", 1))
            End If
             
            If Mid(Trim(xCad), 1, 3) = "MOA" Then
                RstCab("imptot") = NulosN(ExtraerCad(xCad, ",", 1))
            End If
            
            Dim Rst As New ADODB.Recordset
            RST_Busq Rst, "SELECT top 1 numdoc AS numero from ped_pedido  WHERE numser ='0001' AND tipdoc =107 ORDER BY numdoc DESC ", xCon
            If Rst.RecordCount = 0 Then
                RstCab("numdoc") = "0000000001"
                RstCab("idprocped") = 2
                RstCab("idtipped") = 1
            Else
                Rst.MoveFirst
                RstCab("numdoc") = Format(NulosN(Rst("numero")) + 1, "0000000000")
                RstCab("idprocped") = 2
                RstCab("idtipped") = 1
                RstCab("tipdoc") = 107
                RstCab("numser") = "0001"
            End If
            Set Rst = Nothing
            
        End If
        xNumLin = xNumLin + 1
    Wend
    Close #1
    RstCab.Update
    
    Open Trim(NomArchv) For Input As #1
    
    'COMENZAMOS A LEER EL DETALLE DE LOS PEDIDOS
    xNumLin = 0
    
    Dim xNumLin2 As Integer
    Dim id_peddet As Integer
    
    id_peddet = HallaCodigoTabla("ped_pedidodet", xCon, "idpeddet") - 1
    
    xNumLin2 = 1
    
    While Not EOF(1)
        Line Input #1, file_data$
        xCad = file_data$
        
        If xNumLin2 >= 11 Then
            If Mid(Trim(xCad), 1, 3) = "LIN" Then
                If xNumLin = 1 Then RstDet.Update
                xNumLin = 1
                RstDet.AddNew
                id_peddet = id_peddet + 1
                RstDet("idped") = xIdPed
                RstDet("idpeddet") = id_peddet
                RstDet("codpro") = ExtraerCad(xCad, ",", 2)   'codigo del producto
                RstDet("fchent") = CDate(xFchEnt)
                
                Dim cadena As String
                cadena = "SELECT mae_productoscen.codcen, mae_productoscen.iditem " _
                        + vbCr + "From mae_productoscen" _
                        + vbCr + "WHERE (((mae_productoscen.codcen)='" & ExtraerCad(xCad, ",", 2) & "'))"
                        
                RST_Busq RstAuxiditem, cadena, xCon
                
                'Se verifica que los productos cargados correspondan a la Empresa actual
                If Not RstAuxiditem.EOF Then
                    'Se carga
                    RstDet("iditem") = RstAuxiditem("iditem")
                Else
                    'Se envia mensaje de error
                    MsgBox "El codigo de los Productos a procesar no concuerdan con la Empresa seleccionada" _
                            + vbCr + "o no se encuentran en la Base de Datos." _
                            + vbCr + "Seleccione otra Empresa y vuelva a intentarlo", vbInformation + vbOKOnly + vbDefaultButton1, "Error de Carga"
                    Resume LaCague
                End If
            End If
            
            If Mid(Trim(xCad), 1, 3) = "QTY" Then
                RstDet("canpro") = ExtraerCad(xCad, ",", 1)   'cantidad del producto
                RstDet("idunimed") = ExtraerCad(xCad, ",", 2)    'unidad de medida del producto
            End If
            
            If Mid(Trim(xCad), 1, 3) = "PRI" Then
                RstDet("impuni") = NulosN(ExtraerCad(xCad, ",", 1))   'importe unitario
                RstDet("impbru") = RstDet("impuni") * RstDet("canpro") 'importe bruto
            End If
            
            If Mid(Trim(xCad), 1, 6) = "TAXMOA" Then
                RstDet("impigv") = NulosN(ExtraerCad(xCad, ",", 2))   'igv del producto
            End If
            
            If Mid(Trim(xCad), 1, 3) = "MOA" Then
                RstDet("imptot") = NulosN(ExtraerCad(xCad, ",", 1))   'importe total del item
            End If
            
            RstDet("estado") = 2
        End If
        xNumLin2 = xNumLin2 + 1
    Wend
    RstDet.Update
    Close #1
    
    'grabamos el movimiento en la tabla var_edicion para pedido
    GrabarOperacion xIdUsuario, 224, 1, Time, Time, Date, xCon, xIdPed
    
    
    xCon.CommitTrans
    
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDetEnt = Nothing
    AbriryGuardar = True
    Exit Function

LaCague:
    'Resume
    If Err.Number = -2147467259 Then
        MsgBox "Se ha encontrado un codigo de producto no registrado : " & Trim(RstDet("codpro")) & Chr(13) _
            & "de la tienda " & Trim(xNonPunVen) & Chr(13) _
            & "se procedera a anular la Orden de Compra Nº " & Trim(RstCab("numcen")), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If

    If Err.Number = 3021 Then
       ' Resume
        MsgBox "No se ha encontrado la tienda con el codigo Nº " & Mid(xCad, 22, 13) & Chr(13) _
            & "se procede a anular la Orden de Compra Nº " & Trim(RstCab("numcen")), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If

    Close #1
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing

    AbriryGuardar = False
    Exit Function

End Function


'Function AbriryGuardar2(NomArchv As String) As Boolean
'    '===================================================================================================
'    'Creado :   /  /  Por: Enrique Pollongo
'    'Propósito: Crear nuevo pedido
'    '
'    'Entradas:  NomArchv = Nombre del archivo del blok de notas
'    '
'    'Resultados: Pedido registrado en seven
'    '
'    'Modificado: 13/12/10 Por: Jose Chacon
'    '           Aumento de instrucciones para nuevos campos de Base de Datos
'    '           Verificacion de carga de Productos por Empresa
'    '           19/01/11 Por Johan Castro
'    '           Cambiar el tipo de dato a variable xIdPed a Double antes Integer
'    '           Agregar lineas de codigo para registrar el historial del pedido
'    '===================================================================================================
'
'
'    Dim crlf$
'    Dim file_data$
'    crlf$ = Chr(13) & Chr(10)
'    Dim xCad As String
'    Dim xNumLin As Integer
'    Dim xFchEmi, xFchEnt As String
'    Dim xIdPed As Double
'    Dim RstPunVen As New ADODB.Recordset
'    'Se añade una referencia a la tabla ped_pedidodetent
'    Dim RstDetEnt As New ADODB.Recordset
'    Dim xNonPunVen As String
'
'    Dim RstAuxiditem As New ADODB.Recordset
'
'    RST_Busq RstCab, "SELECT * FROM ped_pedido", xCon
'    RST_Busq RstDet, "SELECT * FROM ped_pedidodet", xCon
'    RST_Busq RstDetEnt, "SELECT * FROM ped_pedidodetent", xCon
'
'    On Error GoTo LaCague
'
'    xCon.BeginTrans
'
'    Open Trim(NomArchv) For Input As #1
'
'    xNumLin = 1
'
'    RstCab.AddNew
'    RstDetEnt.AddNew
'
'    xIdPed = HallaCodigoTabla("ped_pedido", xCon, "id")
'    RstCab("id") = xIdPed
'
'    While Not EOF(1)
'        Line Input #1, file_data$
'
'        xCad = file_data$
'        If xNumLin <= 11 Then
'            If Mid(Trim(xCad), 1, 3) = "ENC" Then
'                If Len(Trim(xCad)) = 74 Then
'                    RstCab("numcen") = Mid(xCad, 45, 21)
'                    'Se llena la orden de compra a partir del numero Cen
'                    RstCab("oc") = Mid((Mid(xCad, 45, 21)), 11, 10)
'                End If
'                If Len(Trim(xCad)) = 75 Then
'                    RstCab("numcen") = Mid(xCad, 46, 21)
'                    'Se llena la orden de compra a partir del numero Cen
'                    RstCab("oc") = Mid((Mid(xCad, 46, 21)), 11, 10)
'                End If
'            End If
'
'            If Mid(Trim(xCad), 1, 3) = "DTM" Then
'                xFchEmi = Mid(Trim(xCad), 11, 2) + "/" + Mid(Trim(xCad), 9, 2) + "/" + Mid(Trim(xCad), 5, 4)
'                xFchEnt = Mid(Trim(xCad), 26, 2) + "/" + Mid(Trim(xCad), 24, 2) + "/" + Mid(Trim(xCad), 20, 4)
'                RstCab("fchemi") = CDate(xFchEmi)
'
'                RstCab("fchent") = CDate(xFchEnt)
'                RstDetEnt("fchent") = CDate(xFchEnt)
'            End If
'
'            If Mid(Trim(xCad), 1, 4) = "BYOC" Then
'                RstCab("codcen") = Mid(xCad, 6, 13)
'                RST_Busq RstPunVen, "SELECT vta_puntoventa.idcli,vta_puntoventa.descripcion, vta_puntoventa.codcen, vta_puntoventa.id " _
'                    & " From vta_puntoventa WHERE (((vta_puntoventa.codcen)='" & Mid(xCad, 22, 13) & "'))", xCon
'                RstCab("idpunvecli") = RstPunVen("id")
'                RstCab("idcli") = RstPunVen("idcli")
'                xNonPunVen = RstPunVen("descripcion")
'                Set RstPunVen = Nothing
'            End If
'
'            If Mid(Trim(xCad), 1, 4) = "DPGR" Then
'                RST_Busq RstPunVen, "SELECT vta_puntoventa.descripcion, vta_puntoventa.codcen, vta_puntoventa.id " _
'                    & " From vta_puntoventa WHERE (((vta_puntoventa.codcen)='" & Mid(xCad, 6, 13) & "'))", xCon
'                RstCab("idlugent") = RstPunVen("id")
'                Set RstPunVen = Nothing
'            End If
'
'            If Mid(Trim(xCad), 1, 6) = "TAXMOA" Then
'                RstCab("impigv") = NulosN(ExtraerCad(xCad, ",", 1))
'            End If
'
'            If Mid(Trim(xCad), 1, 3) = "MOA" Then
'                RstCab("imptot") = NulosN(ExtraerCad(xCad, ",", 1))
'            End If
'
'            RstCab("idprocped") = 2
'            RstCab("idtipped") = 1
'            RstCab("tipdoc") = 107
'
'            RstCab("numser") = "0001"
'
'            Dim Rst As New ADODB.Recordset
'            RST_Busq Rst, "SELECT top 1 numdoc AS numero from ped_pedido  WHERE numser ='0001' AND tipdoc =107 ORDER BY numdoc DESC ", xCon
'            If Rst.RecordCount = 0 Then
'                RstCab("numdoc") = "0000000001"
'            Else
'                Rst.MoveFirst
'                RstCab("numdoc") = Format(NulosN(Rst("numero")) + 1, "0000000000")
'            End If
'            Set Rst = Nothing
'
'        End If
'
'        xNumLin = xNumLin + 1
'    Wend
'    Close #1
'    RstCab.Update
'
'    Open Trim(NomArchv) For Input As #1
'
'    'COMENZAMOS A LEER EL DETALLE DE LOS PEDIDOS
'    xNumLin = 0
'
'    Dim xNumLin2 As Integer
'    xNumLin2 = 1
'
'    While Not EOF(1)
'        Line Input #1, file_data$
'        xCad = file_data$
'
'        If xNumLin2 >= 11 Then
'            If Mid(Trim(xCad), 1, 3) = "LIN" Then
'                If xNumLin = 1 Then RstDet.Update
'                xNumLin = 1
'                RstDet.AddNew
'                RstDet("idped") = xIdPed
'                'se añade ped_pedidodetent
'                RstDetEnt("idped") = xIdPed
'
'                RstDet("codpro") = ExtraerCad(xCad, ",", 2)   'codigo del producto
'                Dim cadena As String
'                cadena = "SELECT mae_productoscen.codcen, mae_productoscen.iditem " _
'                        + vbCr + "From mae_productoscen" _
'                        + vbCr + "WHERE (((mae_productoscen.codcen)='" & ExtraerCad(xCad, ",", 2) & "'))"
'                RST_Busq RstAuxiditem, cadena, xCon
'
'                'Se verifica que los productos cargados correspondan a la Empresa actual
'                If Not RstAuxiditem.EOF Then
'                    'Se carga
'                    RstDet("iditem") = RstAuxiditem("iditem")
'                    RstDetEnt("iditem") = RstAuxiditem("iditem")
'                Else
'                    'Se envia mensaje de error
'                    MsgBox "El codigo de los Productos a procesar no concuerdan con la Empresa seleccionada" _
'                            + vbCr + "o no se encuentran en la Base de Datos." _
'                            + vbCr + "Seleccione otra Empresa y vuelva a intentarlo", vbInformation + vbOKOnly + vbDefaultButton1, "Error de Carga"
'                    Resume LaCague
'                End If
'            End If
'
'            If Mid(Trim(xCad), 1, 3) = "QTY" Then
'                RstDet("canpro") = ExtraerCad(xCad, ",", 1)   'cantidad del producto
'                RstDetEnt("canpro") = ExtraerCad(xCad, ",", 1)
'
'                RstDet("idunimed") = ExtraerCad(xCad, ",", 2)    'unidad de medida del producto
'                RstDetEnt("idunimed") = ExtraerCad(xCad, ",", 2)
'            End If
'
'            If Mid(Trim(xCad), 1, 3) = "PRI" Then
'                RstDet("impuni") = NulosN(ExtraerCad(xCad, ",", 1))   'importe unitario
'                RstDet("impbru") = RstDet("impuni") * RstDet("canpro") 'importe bruto
'            End If
'
'            If Mid(Trim(xCad), 1, 6) = "TAXMOA" Then
'                RstDet("impigv") = NulosN(ExtraerCad(xCad, ",", 2))   'igv del producto
'            End If
'
'            If Mid(Trim(xCad), 1, 3) = "MOA" Then
'                RstDet("imptot") = NulosN(ExtraerCad(xCad, ",", 1))   'importe total del item
'            End If
'
'            RstDet("estado") = 2
'            RstDetEnt("estado") = 2
'
'        End If
'        xNumLin2 = xNumLin2 + 1
'    Wend
'    RstDet.Update
'    RstDetEnt.Update
'    Close #1
'
'    'grabamos el movimiento en la tabla var_edicion para pedido
'    GrabarOperacion xIdUsuario, 224, 1, Time, Time, Date, xCon, xIdPed
'
'
'    xCon.CommitTrans
'
'    Set RstCab = Nothing
'    Set RstDet = Nothing
'    Set RstDetEnt = Nothing
'    AbriryGuardar = True
'    Exit Function
'
'LaCague:
'    'Resume
''    MsgBox "Entro al Error"
'    If Err.Number = -2147467259 Then
'        MsgBox "Se ha encontrado un codigo de producto no registrado : " & Trim(RstDet("codpro")) & Chr(13) _
'            & "de la tienda " & Trim(xNonPunVen) & Chr(13) _
'            & "se procedera a anular la Orden de Compra Nº " & Trim(RstCab("numcen")), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'    End If
'
'    If Err.Number = 3021 Then
'       ' Resume
'        MsgBox "No se ha encontrado la tienda con el codigo Nº " & Mid(xCad, 22, 13) & Chr(13) _
'            & "se procede a anular la Orden de Compra Nº " & Trim(RstCab("numcen")), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'    End If
'
'    Close #1
'    xCon.RollbackTrans
'    Set RstCab = Nothing
'    Set RstDet = Nothing
'
'    AbriryGuardar = False
'    Exit Function
'End Function

'Sub o Funcion: CmdProce_Click
'Fecha de Modificacion: 13/12/10
'Ultima modificacion hecha por: Jose Chacon
'Modificacion:
'    Eliminacion de mensaje incoherente en la carga de pedidos


Private Sub CmdProce_Click()
    If Fg1.Rows = 1 Then
        MsgBox "No se ha cargado ninguna orden de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim A As Integer
    Dim Fs As New FileSystemObject
    Dim RutaSys As String
    
    RutaSys = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTAPR", "RUTAS")
    For A = 1 To Fg1.Rows - 1
        If AbriryGuardar(xRutFol + Trim(Fg1.TextMatrix(A, 1))) = True Then
            Fs.CopyFile xRutFol + Trim(Fg1.TextMatrix(A, 1)), RutaSys + Trim(Fg1.TextMatrix(A, 1))
            Fs.DeleteFile xRutFol + Trim(Fg1.TextMatrix(A, 1))
            MsgBox "El proceso culmino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    Next A
    Fg1.Rows = 1
End Sub

Private Sub Command3_Click()
    Dim A As String
    
    A = ExtraerCad("QTY,100,NAR", ",", 1)
    MsgBox A
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Fg1.ColWidth(1) = 3000
End Sub

Function ExtraerCad(cadena As String, Caracter As String, ApartirDelCaracter As Integer) As String
'QTY,100,NAR
    Dim xCad As String
    Dim A As Integer
    Dim xNumRepCar As Integer
    
    xNumRepCar = 0
    
    For A = 1 To Len(Trim(cadena))
        If Mid(cadena, A, 1) = Caracter Then
            xNumRepCar = xNumRepCar + 1
        Else
            If ApartirDelCaracter = xNumRepCar Then
                xCad = xCad + Mid(cadena, A, 1)
            End If
        End If
        
        If xNumRepCar > ApartirDelCaracter Then
            Exit For
        End If
    Next A
    If xCad = "NAR" Then xCad = 2
    ExtraerCad = xCad
End Function
