VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000C&
   ClientHeight    =   750
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11775
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   405
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   8643
            MinWidth        =   8643
            Text            =   "Empresa : "
            TextSave        =   "Empresa : "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Usuario : "
            TextSave        =   "Usuario : "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5145
      Top             =   2115
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":07F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0E30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Seleccionar Empresa"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Agregar Usuario"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cambiar mes de Trabajo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda del sistema"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir del sistema"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu almacen 
      Caption         =   "&Almacén"
      Begin VB.Menu almacen_01 
         Caption         =   "Maestros"
         Begin VB.Menu almacen_01_01 
            Caption         =   "Maestro de Almacenes"
         End
         Begin VB.Menu almacen_01_02 
            Caption         =   "Maestro de Series por Almacen"
         End
         Begin VB.Menu almacen_01_03 
            Caption         =   "Maestro de Unidades"
         End
         Begin VB.Menu almacen_01_04 
            Caption         =   "Maestro Tipo de Item"
         End
         Begin VB.Menu almacen_01_05 
            Caption         =   "Maestro de Familias"
         End
         Begin VB.Menu almacen_01_06 
            Caption         =   "Maestro de Clase"
         End
         Begin VB.Menu almacen_01_07 
            Caption         =   "Maestro de Sub Clase"
         End
         Begin VB.Menu almacen_02 
            Caption         =   "Maestro de Items"
         End
      End
      Begin VB.Menu almacen_03 
         Caption         =   "Configuracion"
         Begin VB.Menu almacen_03_01 
            Caption         =   "Almacenaje Automático"
         End
         Begin VB.Menu almacen_03_02 
            Caption         =   "Despacho Automático"
         End
      End
      Begin VB.Menu almacen_05 
         Caption         =   "Ingreso y Salidas de Almacen"
      End
      Begin VB.Menu almacen_04 
         Caption         =   "Devolución"
      End
      Begin VB.Menu almacen_08 
         Caption         =   "Transferencias"
      End
      Begin VB.Menu almacen_09 
         Caption         =   "Ajustes de Inventario"
      End
      Begin VB.Menu almacen_07 
         Caption         =   "Consulta"
         Begin VB.Menu almacen_07_01 
            Caption         =   "Kardex"
         End
         Begin VB.Menu almacen_07_02 
            Caption         =   "Consulta Ingreso/Salidas de Almacen"
         End
      End
   End
   Begin VB.Menu Compras 
      Caption         =   "&Compras"
      Begin VB.Menu Compras_01 
         Caption         =   "Maestros"
         Begin VB.Menu Compras_01_01 
            Caption         =   "Maestro de Proveedores"
         End
         Begin VB.Menu Compras_01_02 
            Caption         =   "Asignar Personal a Compras"
         End
      End
      Begin VB.Menu Compras_02 
         Caption         =   "-"
      End
      Begin VB.Menu Compras_03 
         Caption         =   "Fijar Precios de Compra a Items"
      End
      Begin VB.Menu Compras_04 
         Caption         =   "Orden de Compra"
         Begin VB.Menu Compras_04_01 
            Caption         =   "Orden de Requerimiento"
         End
         Begin VB.Menu Compras_04_04 
            Caption         =   "Orden de Cotización"
         End
         Begin VB.Menu Compras_04_07 
            Caption         =   "Orden de Compra"
         End
      End
      Begin VB.Menu Compras_05 
         Caption         =   "Registro de Compras"
         Begin VB.Menu Compras_05_01 
            Caption         =   "Registrar Compras"
         End
         Begin VB.Menu Compras_05_02 
            Caption         =   "Registrar Renta 4ta Categoria"
         End
         Begin VB.Menu Compras_05_03 
            Caption         =   "Registrar Gastos Reembolsables"
         End
      End
      Begin VB.Menu Compras_07 
         Caption         =   "Consultas"
         Begin VB.Menu Compras_07_01 
            Caption         =   "Consulta de Compras"
         End
         Begin VB.Menu Compras_07_02 
            Caption         =   "Consulta de Honorarios"
         End
      End
   End
   Begin VB.Menu ventas 
      Caption         =   "&Ventas"
      Begin VB.Menu ventas_01 
         Caption         =   "Maestros"
         Begin VB.Menu ventas_01_01 
            Caption         =   "Maestro de Clientes"
         End
         Begin VB.Menu ventas_01_02 
            Caption         =   "Maestro de Puntos de Venta del Cliente"
         End
         Begin VB.Menu ventas_01_03 
            Caption         =   "Maestro de Vendedores"
         End
         Begin VB.Menu ventas_01_04 
            Caption         =   "Maestro de Productos CEN"
         End
         Begin VB.Menu ventas_01_05 
            Caption         =   "Maestro de Conceptos de NC y ND"
         End
         Begin VB.Menu ventas_01_06 
            Caption         =   "Datos de Transporte"
            Begin VB.Menu ventas_01_06_01 
               Caption         =   "Maestro Empresas de Transporte"
            End
            Begin VB.Menu ventas_01_06_02 
               Caption         =   "Maestro de Choferes"
            End
            Begin VB.Menu ventas_01_06_03 
               Caption         =   "Maestro Unidades de Transporte"
            End
            Begin VB.Menu ventas_01_06_04 
               Caption         =   "Maestro Motivos de Traslado"
            End
         End
      End
      Begin VB.Menu ventas_02 
         Caption         =   "-"
      End
      Begin VB.Menu ventas_03 
         Caption         =   "Teleprocesos CEN"
         Begin VB.Menu ventas_03_01 
            Caption         =   "Levantar Pedidos"
         End
         Begin VB.Menu ventas_03_02 
            Caption         =   "Procesar Pedidos"
         End
      End
      Begin VB.Menu ventas_04 
         Caption         =   "Cotizaciones"
      End
      Begin VB.Menu ventas_05 
         Caption         =   "Pedidos"
         Begin VB.Menu ventas_05_01 
            Caption         =   "Pedidos"
         End
         Begin VB.Menu ventas_05_02 
            Caption         =   "Cronograma de Entregas"
         End
         Begin VB.Menu ventas_05_03 
            Caption         =   "Reporte de Pedidos"
         End
      End
      Begin VB.Menu ventas_06 
         Caption         =   "Facturación"
         Begin VB.Menu ventas_06_01 
            Caption         =   "Guias de Remisión"
         End
         Begin VB.Menu ventas_06_02 
            Caption         =   "Registrar Ventas"
         End
         Begin VB.Menu ventas_06_03 
            Caption         =   "Liquidación Gasto Débito"
         End
      End
      Begin VB.Menu ventas_07 
         Caption         =   "-"
      End
      Begin VB.Menu ventas_08 
         Caption         =   "Consultas"
         Begin VB.Menu ventas_08_01 
            Caption         =   "Consulta de Ventas"
         End
         Begin VB.Menu ventas_08_02 
            Caption         =   "Consulta de Devoluciones"
         End
      End
   End
   Begin VB.Menu contabilidad 
      Caption         =   "C&ontabilidad"
      Begin VB.Menu contabilidad_01 
         Caption         =   "Maestros"
         Begin VB.Menu contabilidad_01_01 
            Caption         =   "Maestro de Detracciones"
         End
         Begin VB.Menu contabilidad_01_02 
            Caption         =   "Maestro de Percepciones"
         End
         Begin VB.Menu contabilidad_01_03 
            Caption         =   "Maestro de Retenciones"
         End
         Begin VB.Menu contabilidad_01_04 
            Caption         =   "Maestro de Libros Contables"
         End
         Begin VB.Menu contabilidad_01_05 
            Caption         =   "Maestro de Documentos Contables"
         End
         Begin VB.Menu contabilidad_01_06 
            Caption         =   "Maestro de Impuestos"
         End
         Begin VB.Menu contabilidad_01_07 
            Caption         =   "Maestro de Configuracion de Valorizacion"
         End
      End
      Begin VB.Menu contabilidad_03 
         Caption         =   "Configuración"
         Begin VB.Menu contabilidad_03_01 
            Caption         =   "Plan de Cuentas"
         End
         Begin VB.Menu contabilidad_03_02 
            Caption         =   "Centro de Costos"
            Begin VB.Menu contabilidad_03_02_01 
               Caption         =   "Centro de Costos"
            End
            Begin VB.Menu contabilidad_03_02_02 
               Caption         =   "Asignar Centro de Costos a Areas"
            End
         End
         Begin VB.Menu contabilidad_03_03 
            Caption         =   "Asignar Cuenta Contable a Documentos"
         End
         Begin VB.Menu contabilidad_03_04 
            Caption         =   "Estados Financieros"
            Begin VB.Menu contabilidad_03_04_01 
               Caption         =   "Configuración de Conceptos"
            End
            Begin VB.Menu contabilidad_03_04_02 
               Caption         =   "Diseño de Informes"
            End
         End
         Begin VB.Menu contabilidad_03_05 
            Caption         =   "Corregir Asientos"
         End
         Begin VB.Menu contabilidad_03_06 
            Caption         =   "Cerrar Mes"
         End
         Begin VB.Menu contabilidad_03_07 
            Caption         =   "Tipo de Cambio"
         End
         Begin VB.Menu contabilidad_03_08 
            Caption         =   "Transferir Operaciones"
         End
      End
      Begin VB.Menu contabilidad_04 
         Caption         =   "Operaciones"
         Begin VB.Menu contabilidad_04_01 
            Caption         =   "Detracciones"
            Begin VB.Menu contabilidad_04_01_01 
               Caption         =   "Compras"
            End
            Begin VB.Menu contabilidad_04_01_02 
               Caption         =   "Ventas"
            End
         End
         Begin VB.Menu contabilidad_04_02 
            Caption         =   "Percepción"
         End
         Begin VB.Menu contabilidad_04_03 
            Caption         =   "Retención"
         End
         Begin VB.Menu contabilidad_04_04 
            Caption         =   "Asientos Diversos"
         End
      End
      Begin VB.Menu contabilidad_05 
         Caption         =   "Libros"
         Begin VB.Menu contabilidad_05_01 
            Caption         =   "Registro de Compras"
         End
         Begin VB.Menu contabilidad_05_02 
            Caption         =   "Registro de Honorarios"
         End
         Begin VB.Menu contabilidad_05_03 
            Caption         =   "Registro de Ventas"
         End
         Begin VB.Menu contabilidad_05_04 
            Caption         =   "-"
         End
         Begin VB.Menu contabilidad_05_05 
            Caption         =   "Libro Diario"
         End
         Begin VB.Menu contabilidad_05_06 
            Caption         =   "Libro Mayor"
         End
         Begin VB.Menu contabilidad_05_07 
            Caption         =   "Balance de Comprobación"
         End
         Begin VB.Menu contabilidad_05_08 
            Caption         =   "Estados Financieros"
         End
         Begin VB.Menu contabilidad_05_09 
            Caption         =   "-"
         End
         Begin VB.Menu contabilidad_05_10 
            Caption         =   "Conciliacion Bancaria"
         End
         Begin VB.Menu contabilidad_05_11 
            Caption         =   "Kardex Valorizado"
         End
         Begin VB.Menu contabilidad_05_12 
            Caption         =   "Libro de Costos"
         End
      End
      Begin VB.Menu contabilidad_07 
         Caption         =   "Consultas"
         Begin VB.Menu contabilidad_07_01 
            Caption         =   "DAOT"
         End
         Begin VB.Menu contabilidad_07_02 
            Caption         =   "Centros de Costos"
         End
         Begin VB.Menu contabilidad_07_03 
            Caption         =   "Costos de Movimientos"
         End
         Begin VB.Menu contabilidad_07_04 
            Caption         =   "Costos de Partes de Produccion"
         End
         Begin VB.Menu contabilidad_07_05 
            Caption         =   "Analisis de Costos de Produccion"
         End
         Begin VB.Menu contabilidad_07_06 
            Caption         =   "Kardex Resumen Valorizado"
         End
         Begin VB.Menu contabilidad_07_07 
            Caption         =   "Kardex Valorizado Detallado"
         End
      End
   End
   Begin VB.Menu tesoreria 
      Caption         =   "&Tesorería"
      Begin VB.Menu tesoreria_01 
         Caption         =   "Maestros"
         Begin VB.Menu tesoreria_01_01 
            Caption         =   "Maestro de Origen"
            Begin VB.Menu tesoreria_01_01_01 
               Caption         =   "Ingreso"
            End
            Begin VB.Menu tesoreria_01_01_02 
               Caption         =   "Egreso"
            End
         End
         Begin VB.Menu tesoreria_01_02 
            Caption         =   "Maestro Destino"
            Begin VB.Menu tesoreria_01_02_01 
               Caption         =   "Ingreso"
            End
            Begin VB.Menu tesoreria_01_02_02 
               Caption         =   "Egreso"
            End
         End
         Begin VB.Menu tesoreria_01_03 
            Caption         =   "Maestro de Bancos"
         End
         Begin VB.Menu tesoreria_01_04 
            Caption         =   "Maestro de Cuentas de Banco"
         End
         Begin VB.Menu tesoreria_01_05 
            Caption         =   "Maestro de Documentos de Caja y Bancos"
         End
         Begin VB.Menu tesoreria_01_08 
            Caption         =   "Maestro de Medio de Pago"
         End
         Begin VB.Menu tesoreria_01_09 
            Caption         =   "Asignar Empleados a Tesoreria"
         End
      End
      Begin VB.Menu tesoreria_02 
         Caption         =   "-"
      End
      Begin VB.Menu tesoreria_03 
         Caption         =   "Ingresos"
      End
      Begin VB.Menu tesoreria_04 
         Caption         =   "Egresos"
      End
      Begin VB.Menu tesoreria_05 
         Caption         =   "Canje de Documentos"
      End
      Begin VB.Menu tesoreria_06 
         Caption         =   "Letras"
         Begin VB.Menu tesoreria_06_01 
            Caption         =   "Emision de Letras"
         End
         Begin VB.Menu tesoreria_06_02 
            Caption         =   "Planilla de Cobranzas"
         End
      End
      Begin VB.Menu tesoreria_07 
         Caption         =   "-"
      End
      Begin VB.Menu tesoreria_08 
         Caption         =   "Consultas"
         Begin VB.Menu tesoreria_08_01 
            Caption         =   "Analisis Cliente Proveedor"
         End
         Begin VB.Menu tesoreria_08_02 
            Caption         =   "Analisis Cta. Cte. Cliente"
         End
         Begin VB.Menu tesoreria_08_03 
            Caption         =   "Anticuamiento"
         End
      End
   End
   Begin VB.Menu produccion 
      Caption         =   "&Producción"
      Begin VB.Menu produccion_01 
         Caption         =   "Maestros"
         Begin VB.Menu produccion_01_01 
            Caption         =   "Maestro de Tareas"
         End
         Begin VB.Menu produccion_01_02 
            Caption         =   "Maestro de Recetas"
         End
         Begin VB.Menu produccion_01_03 
            Caption         =   "Maestro de Estacionalidad"
         End
         Begin VB.Menu produccion_01_04 
            Caption         =   "Maestro de Costos de Tareas"
         End
         Begin VB.Menu produccion_01_06 
            Caption         =   "Maestro de Personal x Tareas"
         End
         Begin VB.Menu produccion_01_05 
            Caption         =   "Maestro de Rendimientos"
         End
         Begin VB.Menu produccion_01_07 
            Caption         =   "Maestro de Lineas de Producción"
         End
      End
      Begin VB.Menu produccion_02 
         Caption         =   "-"
      End
      Begin VB.Menu produccion_03 
         Caption         =   "Planeamiento de la Producción"
         Begin VB.Menu produccion_03_01 
            Caption         =   "Cronograma de Producción"
         End
         Begin VB.Menu produccion_03_02 
            Caption         =   "Orden de Producción"
         End
         Begin VB.Menu produccion_03_04 
            Caption         =   "Solicitud de Materiales"
         End
      End
      Begin VB.Menu produccion_04 
         Caption         =   "Registro de Producción"
      End
      Begin VB.Menu produccion_05 
         Caption         =   "Mano de Obra"
         Begin VB.Menu produccion_05_01 
            Caption         =   "Grupos de Trabajo"
         End
         Begin VB.Menu produccion_05_02 
            Caption         =   "Personal de Producción"
         End
         Begin VB.Menu produccion_05_03 
            Caption         =   "Distribuir Tareas por Area"
         End
         Begin VB.Menu produccion_05_04 
            Caption         =   "Registro de Tareas"
         End
      End
      Begin VB.Menu produccion_06 
         Caption         =   "-"
      End
      Begin VB.Menu produccion_07 
         Caption         =   "Consultas"
         Begin VB.Menu produccion_07_01 
            Caption         =   "Consulta de Producción"
         End
         Begin VB.Menu produccion_05_05 
            Caption         =   "Consulta de Tarea"
         End
         Begin VB.Menu produccion_07_02 
            Caption         =   "Consulta de Lineas"
         End
      End
      Begin VB.Menu produccion_08 
         Caption         =   "Analisis Comparativo de Producción"
      End
      Begin VB.Menu produccion_09 
         Caption         =   "Analisis de Planeamiento de Producción"
      End
   End
   Begin VB.Menu planillas 
      Caption         =   "P&lanillas"
      Begin VB.Menu planillas_01 
         Caption         =   "Maestros"
         Begin VB.Menu planillas_01_01 
            Caption         =   "Maestro de Areas"
         End
         Begin VB.Menu planillas_01_02 
            Caption         =   "Maestro de Cargos"
         End
         Begin VB.Menu planillas_01_03 
            Caption         =   "Maestro de Conceptos"
         End
         Begin VB.Menu planillas_01_04 
            Caption         =   "Maestro de Fondo de Pensiones"
         End
         Begin VB.Menu planillas_01_05 
            Caption         =   "Maestro de Empleados"
         End
         Begin VB.Menu planillas_01_06 
            Caption         =   "-"
         End
         Begin VB.Menu planillas_01_07 
            Caption         =   "Maestro de Horarios"
         End
         Begin VB.Menu planillas_01_08 
            Caption         =   "Maestro de Licencias"
         End
         Begin VB.Menu planillas_01_09 
            Caption         =   "Maestro de Permisos"
         End
         Begin VB.Menu planillas_01_10 
            Caption         =   "Maestro de Dias Festivos"
         End
      End
      Begin VB.Menu planillas_02 
         Caption         =   "Controles"
         Begin VB.Menu planillas_02_01 
            Caption         =   "Control de Asistencia"
            Begin VB.Menu planillas_02_01_01 
               Caption         =   "Registro de Asistencia"
            End
            Begin VB.Menu planillas_02_01_02 
               Caption         =   "Importar Asistencia"
            End
         End
         Begin VB.Menu planillas_02_02 
            Caption         =   "Control de Licencias"
         End
         Begin VB.Menu planillas_02_03 
            Caption         =   "Control de Permisos"
         End
         Begin VB.Menu planillas_02_04 
            Caption         =   "Control de Vacaciones"
         End
      End
      Begin VB.Menu planillas_03 
         Caption         =   "-"
      End
      Begin VB.Menu planillas_04 
         Caption         =   "Resumen de Horas"
      End
      Begin VB.Menu planillas_05 
         Caption         =   "Asignar Sueldos"
      End
      Begin VB.Menu planillas_06 
         Caption         =   "Registrar Boletas de Pago"
      End
      Begin VB.Menu planillas_07 
         Caption         =   "Comisión de Vendedores"
      End
      Begin VB.Menu planillas_08 
         Caption         =   "Impresión de Boletas"
      End
      Begin VB.Menu planillas_09 
         Caption         =   "-"
      End
      Begin VB.Menu planillas_10 
         Caption         =   "Planilla de Producción"
      End
      Begin VB.Menu planillas_11 
         Caption         =   "Consulta Planilla de Producción"
      End
      Begin VB.Menu planillas_12 
         Caption         =   "-"
      End
      Begin VB.Menu planillas_13 
         Caption         =   "Consultas"
         Begin VB.Menu planillas_13_01 
            Caption         =   "Consulta de Asistencia"
         End
         Begin VB.Menu planillas_13_02 
            Caption         =   "Consulta de Planillas"
         End
         Begin VB.Menu planillas_13_03 
            Caption         =   "Exportar Datos "
            Begin VB.Menu planillas_13_03_01 
               Caption         =   "Exportar Datos del Trabajador"
            End
         End
         Begin VB.Menu planillas_13_04 
            Caption         =   "Consulta de Rotaciones"
         End
      End
   End
   Begin VB.Menu gestion 
      Caption         =   "&Gestión"
      Begin VB.Menu gestion_01 
         Caption         =   "Análisis de Compras"
      End
      Begin VB.Menu gestion_02 
         Caption         =   "Análisis de Ventas"
      End
      Begin VB.Menu gestion_03 
         Caption         =   "Análisis de Producción"
      End
      Begin VB.Menu gestion_04 
         Caption         =   "Análisis de Tesorería"
      End
      Begin VB.Menu gestion_05 
         Caption         =   "-"
      End
      Begin VB.Menu gestion_06 
         Caption         =   "Planeamiento"
         Begin VB.Menu gestion_06_01 
            Caption         =   "Proyección de Ventas"
         End
         Begin VB.Menu gestion_06_02 
            Caption         =   "Plan de Ventas"
         End
         Begin VB.Menu gestion_06_03 
            Caption         =   "-"
         End
         Begin VB.Menu gestion_06_04 
            Caption         =   "Plan de Producción"
         End
         Begin VB.Menu gestion_06_05 
            Caption         =   "Plan de Producción (Unificado)"
         End
         Begin VB.Menu gestion_06_06 
            Caption         =   "-"
         End
         Begin VB.Menu gestion_06_07 
            Caption         =   "Plan de Abastecimiento"
         End
         Begin VB.Menu gestion_06_08 
            Caption         =   "Plan de Abastecimiento (Unificado)"
         End
         Begin VB.Menu gestion_06_09 
            Caption         =   "-"
         End
         Begin VB.Menu gestion_06_10 
            Caption         =   "Producción Total (unificado)"
         End
      End
      Begin VB.Menu gestion_07 
         Caption         =   "Analisis x Doc Referencia"
      End
      Begin VB.Menu gestion_08 
         Caption         =   "Analisis x Unidad de Negocio"
      End
      Begin VB.Menu gestion_09 
         Caption         =   "Control de Registro x Módulos"
      End
   End
   Begin VB.Menu setup 
      Caption         =   "?"
      Begin VB.Menu setup_01 
         Caption         =   "Mantenimiento de Empresas"
      End
      Begin VB.Menu setup_02 
         Caption         =   "Seleccionar Ruta de Acceso"
      End
      Begin VB.Menu setup_04 
         Caption         =   "Setup"
      End
      Begin VB.Menu setup_05 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu setup17 
         Caption         =   "Mantenimiento de Opciones de Sistema"
      End
      Begin VB.Menu setup_06 
         Caption         =   "Configurar Opciones de Usuario"
      End
      Begin VB.Menu setup_07 
         Caption         =   "Plantillas de Impresión"
      End
      Begin VB.Menu setup_12 
         Caption         =   "Impresión de Etiquetas"
      End
      Begin VB.Menu setup_14 
         Caption         =   "Registrar Licencia"
      End
      Begin VB.Menu setup15 
         Caption         =   "Transferencia"
         Begin VB.Menu setup15_01 
            Caption         =   "Maestro Equivalencia - Cliente/Proveedor"
         End
         Begin VB.Menu setup15_02 
            Caption         =   "Maestro Equivalencia - Items"
         End
         Begin VB.Menu setup15_03 
            Caption         =   "Maestro Equivalencia - T. Documento"
         End
         Begin VB.Menu setup15_04 
            Caption         =   "Transferencia Documentos"
         End
      End
      Begin VB.Menu setup_13 
         Caption         =   "Importación de Datos"
         Begin VB.Menu setup_13_01 
            Caption         =   "Cliente"
         End
         Begin VB.Menu setup_13_02 
            Caption         =   "Proveedores"
         End
         Begin VB.Menu setup_13_03 
            Caption         =   "Registro de Compras"
         End
         Begin VB.Menu setup_13_04 
            Caption         =   "Registro de Honorarios"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu setup_13_05 
            Caption         =   "Registro de Ventas"
         End
      End
      Begin VB.Menu setup_08 
         Caption         =   "-"
      End
      Begin VB.Menu setup_09 
         Caption         =   "Ayuda"
      End
      Begin VB.Menu setup_16 
         Caption         =   "Actualizaciones"
         Enabled         =   0   'False
      End
      Begin VB.Menu setup_10 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo    : MDIPRINCIPAL
'* Tipo              : FORMULARIO
'* Descripcion       :
'*
'*
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 01/09/09
'* VERSION           : 1.0
'*****************************************************************************************************
Option Explicit
Dim SeEjecuto As Boolean

Private Sub almacen_01_01_Click()
    '*****************************************************
    ' modificado 01/05/2012 - Jose Chacon
    '*****************************************************
    Dim xfrm As New SGI2_almacen.almacen
    xfrm.IdMenu = 88
    xfrm.Idusuario = xIdUsuario
    xfrm.ManAlmacen xCon
    Set xfrm = Nothing
End Sub

Private Sub almacen_01_02_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(4, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT alm_numseries.*, alm_almacenes.descripcion AS desalm, mae_documento.descripcion AS destipdoc FROM (alm_numseries " _
        & " LEFT JOIN alm_almacenes ON alm_numseries.idalm = alm_almacenes.id) LEFT JOIN mae_documento ON alm_numseries.idtipdoc = mae_documento.id " _
        & " ORDER BY alm_numseries.numser, mae_documento.descripcion"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":              xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "C"
    xCamposVista(1, 0) = "Almacen":             xCamposVista(1, 1) = "desalm":         xCamposVista(1, 2) = "2000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Tipo Documento":      xCamposVista(2, 1) = "destipdoc":      xCamposVista(2, 2) = "3000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Nº Serie":            xCamposVista(3, 1) = "numser":         xCamposVista(3, 2) = "1500":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "C"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":           xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Almacen":          xCampos(1, 1) = "idalm":        xCampos(1, 2) = "N":    xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Tipo Documento":   xCampos(2, 1) = "idtipdoc":     xCampos(2, 2) = "N":    xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Nº Serie":         xCampos(3, 1) = "numser":       xCampos(3, 2) = "C":    xCampos(3, 3) = "1500"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "alm_almacenes":         xVinculos(0, 1) = "id":            xVinculos(0, 2) = "id,descripcion":
    xVinculos(0, 3) = "Codigo,Descripcion":    xVinculos(0, 4) = "1000,3000":     xVinculos(0, 5) = "N,C":
    xVinculos(0, 6) = "idalm":                 xVinculos(0, 7) = "descripcion":   xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "id"
    
    xVinculos(1, 0) = "mae_documento":         xVinculos(1, 1) = "id":            xVinculos(1, 2) = "id,descripcion":
    xVinculos(1, 3) = "Codigo,Descripcion":    xVinculos(1, 4) = "1000,5000":     xVinculos(1, 5) = "N,C":
    xVinculos(1, 6) = "idtipdoc":              xVinculos(1, 7) = "descripcion":   xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "id"

    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 89
    '-----
    
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "alm_numseries"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Series x Almacen"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub almacen_01_03_Click()
    ' EJECUTA MENU
    Dim xFun As New SGI2_almacen.almacen
    xFun.IdMenu = 5
    xFun.Idusuario = xIdUsuario
    xFun.ManUnidades xCon, 1
    Set xFun = Nothing
End Sub

Private Sub almacen_01_04_Click()
    ' EJECUTA MENU
    Dim xFun As New SGI2_almacen.almacen
    xFun.IdMenu = 6
    xFun.Idusuario = xIdUsuario
    xFun.ManTipoProducto xCon, 1
    Set xFun = Nothing
End Sub

Private Sub almacen_01_05_Click()
    ' EJECUTA MENU
    Dim xFun As New SGI2_almacen.almacen
    xFun.IdMenu = 2
    xFun.Idusuario = xIdUsuario
    xFun.ManFamilia xCon, 1
    Set xFun = Nothing
End Sub

Private Sub almacen_01_06_Click()
    ' EJECUTA MENU
    Dim xfrm As New SGI2_almacen.almacen
    xfrm.IdMenu = 3
    xfrm.Idusuario = xIdUsuario
    xfrm.ManClase xCon, 1
    Set xfrm = Nothing
    
End Sub

Private Sub almacen_01_07_Click()
    ' EJECUTA MENU
    Dim xfrm As New SGI2_almacen.almacen
    xfrm.IdMenu = 4
    xfrm.Idusuario = xIdUsuario
    xfrm.ManSubClase xCon, 1
    Set xfrm = Nothing
End Sub

Private Sub almacen_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New SGI2_almacen.almacen
    xfrm.IdMenu = 7
    xfrm.Idusuario = xIdUsuario
    xfrm.MantItem xCon, 1
    Set xfrm = Nothing
End Sub

Private Sub almacen_03_01_Click()
    Dim xfrm As New SGI2_almacen.almacen
    xfrm.IdMenu = 253
    xfrm.Idusuario = xIdUsuario
    xfrm.ManAlmacenajeAuto xCon
    Set xfrm = Nothing
End Sub

Private Sub almacen_03_02_Click()
    Dim xfrm As New SGI2_almacen.almacen
    xfrm.IdMenu = 253
    xfrm.Idusuario = xIdUsuario
    xfrm.ManDespachoAuto xCon
    Set xfrm = Nothing
End Sub

Private Sub almacen_04_Click()
    Dim xfrm As New SGI2_almacen.almacen
    xfrm.IdMenu = 253
    xfrm.Idusuario = xIdUsuario
    xfrm.DevolucionAlmacen xCon, CInt(Mid(Date, 4, 2))
    Set xfrm = Nothing
End Sub

Private Sub almacen_05_Click()
    Dim xfrm As New SGI2_almacen.almacen
    xfrm.IdMenu = 8
    xfrm.Idusuario = xIdUsuario
    xfrm.IngresoAlmacen xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub almacen_07_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.MostrarStockResumen xCon, False
    Set xfrm = Nothing
End Sub

Private Sub almacen_07_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_compras.Compras
    xfrm.ConsultaIngSalAlm xCon
    Set xfrm = Nothing
End Sub

Private Sub almacen_08_Click()
    Dim xfrm As New SGI2_almacen.almacen
    xfrm.IdMenu = 262
    xfrm.Idusuario = xIdUsuario
    xfrm.TransferenciaAlmacen xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub almacen_09_Click()
    Dim xfrm As New SGI2_almacen.almacen
    xfrm.IdMenu = 263
    xfrm.Idusuario = xIdUsuario
    xfrm.TomaInventario xCon
    Set xfrm = Nothing
End Sub

Private Sub Compras_01_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_compras.Compras
    xfrm.IdMenu = 10
    xfrm.Idusuario = xIdUsuario
    xfrm.ManProveedor xCon, 1
    Set xfrm = Nothing
End Sub

Private Sub Compras_01_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New Sgi2_Procesos.Procesos
    xfrm.PersonalCompras xCon
    Set xfrm = Nothing
End Sub

Private Sub Compras_03_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_compras.Compras
    xfrm.IdMenu = 107
    xfrm.Idusuario = xIdUsuario
    xfrm.AsignaPrecioItem xCon
    Set xfrm = Nothing
End Sub

Private Sub Compras_04_01_Click()
    ' EJECUTA MENU
    Dim xFun As New seven_compras2.Compras
    xFun.IdMenu = 215
    xFun.Idusuario = xIdUsuario
    xFun.ManOrdenrequerimiento xCon
    Set xFun = Nothing
End Sub

Private Sub Compras_04_02_Click()
    ' EJECUTA MENU
    Dim xFun As New seven_compras2.Compras
    xFun.AprobarRequerimiento xIdUsuario, xCon
    Set xFun = Nothing
End Sub

Private Sub Compras_04_03_Click()
    ' EJECUTA MENU
    Dim xFun As New seven_compras2.Compras
    xFun.ManOrdenCotizacion xCon
    Set xFun = Nothing
End Sub

Private Sub Compras_04_04_Click()
    ' EJECUTA MENU
    Dim xFun As New seven_compras2.Compras
    xFun.IdMenu = 216
    xFun.Idusuario = xIdUsuario
    xFun.ManOrdenCotizacion xCon
    Set xFun = Nothing
End Sub

Private Sub Compras_04_05_Click()
    ' EJECUTA MENU
    Dim xFun As New seven_compras2.Compras
    xFun.AprobarCotizacion xIdUsuario, xCon
    Set xFun = Nothing
    
End Sub

Private Sub Compras_04_07_Click()
    ' EJECUTA MENU
    Dim xFun As New seven_compras2.Compras
    xFun.IdMenu = 217
    xFun.Idusuario = xIdUsuario
    xFun.ManOrdenCompra xIdUsuario, xCon
    Set xFun = Nothing
End Sub

Private Sub Compras_04_08_Click()
    ' EJECUTA MENU
    Dim xFun As New seven_compras2.Compras
    xFun.AprobarOrdenCompra xIdUsuario, xCon
    Set xFun = Nothing
End Sub

Private Sub Compras_05_01_Click()
    ' EJECUTA MENU
    Dim xform As New sgi2_compras.Compras
    xform.IdMenu = 218
    xform.Idusuario = xIdUsuario
    xform.RegCompras2 xCon, AP_MESTRA, 0
    Set xform = Nothing
End Sub

Private Sub Compras_05_02_Click()
    ' EJECUTA MENU
    Dim xform As New sgi2_compras.Compras
    xform.IdMenu = 219
    xform.Idusuario = xIdUsuario
    xform.RegHonorarios xCon, AP_MESTRA, 0
    Set xform = Nothing
End Sub

Private Sub Compras_05_03_Click()
    Dim xfrm As New sgi2_compras.Compras
    xfrm.IdMenu = 220
    xfrm.Idusuario = xIdUsuario
    xfrm.Reembolsables xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub Compras_07_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_compras.Compras
    xfrm.RepCompras xCon
    Set xfrm = Nothing
End Sub

Private Sub Compras_07_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_compras.Compras
    xfrm.RepHonorario xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_01_01_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu
    '               22/02/11 Johan Castro
    '               Agregar campo impbase, determinar el calculo de detraccion de compra o venta
    
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(4, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_detraccion.* From mae_detraccion ORDER BY mae_detraccion.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":         xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripción":    xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Tasa":           xCamposVista(2, 1) = "tasa":           xCamposVista(2, 2) = "1100":   xCamposVista(2, 3) = "N":    xCamposVista(2, 4) = "D"
    
    xCamposVista(3, 0) = "Base":           xCamposVista(3, 1) = "impbase":        xCamposVista(3, 2) = "1100":   xCamposVista(2, 3) = "N":    xCamposVista(3, 4) = "D"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "6000"
    xCampos(2, 0) = "Tasa":           xCampos(2, 1) = "tasa":         xCampos(2, 2) = "N":    xCampos(2, 3) = "1100"
    
    xCampos(3, 0) = "Imp Base":       xCampos(3, 1) = "impbase":      xCampos(3, 2) = "N":    xCampos(3, 3) = "1100"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":       xVinculos(0, 1) = "":          xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":       xVinculos(0, 4) = "":          xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":       xVinculos(0, 7) = "":          xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 21
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_detraccion"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Detracciones"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub contabilidad_01_02_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(5, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(5, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_percepcion.id, mae_percepcion.idcuencom, mae_percepcion.idcuenven, mae_percepcion.descripcion, " _
        & " mae_percepcion.tasa, con_planctas.cuenta AS cuentacom, con_planctas_1.cuenta AS cuentaven, " _
        & " con_planctas.descripcion AS descuecom, con_planctas_1.descripcion AS descueven FROM (mae_percepcion LEFT JOIN con_planctas ON " _
        & " mae_percepcion.idcuencom = con_planctas.id) LEFT JOIN con_planctas AS con_planctas_1 ON " _
        & " mae_percepcion.idcuenven = con_planctas_1.id ORDER BY mae_percepcion.descripcion"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":        xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "5000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Tasa":               xCamposVista(2, 1) = "tasa":           xCamposVista(2, 2) = "1100":   xCamposVista(2, 3) = "N":    xCamposVista(2, 4) = "D"
    xCamposVista(3, 0) = "Cta. Compra":        xCamposVista(3, 1) = "cuentacom":      xCamposVista(3, 2) = "1300":   xCamposVista(3, 3) = "N":    xCamposVista(3, 4) = "I"
    xCamposVista(4, 0) = "Cta. Venta":         xCamposVista(4, 1) = "cuentaven":      xCamposVista(4, 2) = "1300":   xCamposVista(4, 3) = "N":    xCamposVista(4, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Tasa":           xCampos(2, 1) = "tasa":         xCampos(2, 2) = "N":    xCampos(2, 3) = "1100"
    xCampos(3, 0) = "Cuenta Compra":  xCampos(3, 1) = "idcuencom":    xCampos(3, 2) = "N":    xCampos(3, 3) = "1100"
    xCampos(4, 0) = "Cuenta Venta":   xCampos(4, 1) = "idcuenven":    xCampos(4, 2) = "N":    xCampos(4, 3) = "1100"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "con_planctas":       xVinculos(0, 1) = "id":          xVinculos(0, 2) = "cuenta,descripcion":
    xVinculos(0, 3) = "Cuenta,Descripcion": xVinculos(0, 4) = "1100,4000":   xVinculos(0, 5) = "C,C":
    xVinculos(0, 6) = "idcuencom":          xVinculos(0, 7) = "descripcion": xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "cuenta"
    
    xVinculos(1, 0) = "con_planctas":       xVinculos(1, 1) = "id":          xVinculos(1, 2) = "cuenta,descripcion":
    xVinculos(1, 3) = "Cuenta,Descripcion": xVinculos(1, 4) = "1100,4000":   xVinculos(1, 5) = "C,C":
    xVinculos(1, 6) = "idcuenven":          xVinculos(1, 7) = "descripcion": xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "cuenta"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 23
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_percepcion"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Percepciones"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub contabilidad_01_03_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu


    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(5, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(5, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_retencion.id, mae_retencion.idcuencom, mae_retencion.idcuenven, mae_retencion.descripcion, " _
        & " mae_retencion.tasa, mae_retencion.defaul, con_planctas.cuenta AS cuentacom, con_planctas_1.cuenta AS cuentaven, " _
        & " con_planctas.descripcion AS descuecom, con_planctas_1.descripcion AS descueven FROM (mae_retencion LEFT JOIN con_planctas ON " _
        & " mae_retencion.idcuencom = con_planctas.id) LEFT JOIN con_planctas AS con_planctas_1 ON " _
        & " mae_retencion.idcuenven = con_planctas_1.id ORDER BY mae_retencion.descripcion"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":        xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "5000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Tasa":               xCamposVista(2, 1) = "tasa":           xCamposVista(2, 2) = "1100":   xCamposVista(2, 3) = "N":    xCamposVista(2, 4) = "D"
    xCamposVista(3, 0) = "Cta. Compra":        xCamposVista(3, 1) = "cuentacom":      xCamposVista(3, 2) = "1300":   xCamposVista(3, 3) = "N":    xCamposVista(3, 4) = "I"
    xCamposVista(4, 0) = "Cta. Venta":         xCamposVista(4, 1) = "cuentaven":      xCamposVista(4, 2) = "1300":   xCamposVista(4, 3) = "N":    xCamposVista(4, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Tasa":           xCampos(2, 1) = "tasa":         xCampos(2, 2) = "N":    xCampos(2, 3) = "1100"
    xCampos(3, 0) = "Cuenta Compra":  xCampos(3, 1) = "idcuencom":    xCampos(3, 2) = "N":    xCampos(3, 3) = "1100"
    xCampos(4, 0) = "Cuenta Venta":   xCampos(4, 1) = "idcuenven":    xCampos(4, 2) = "N":    xCampos(4, 3) = "1100"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "con_planctas":       xVinculos(0, 1) = "id":          xVinculos(0, 2) = "cuenta,descripcion":
    xVinculos(0, 3) = "Cuenta,Descripcion": xVinculos(0, 4) = "1100,4000":   xVinculos(0, 5) = "C,C":
    xVinculos(0, 6) = "idcuencom":          xVinculos(0, 7) = "descripcion": xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "cuenta"
    
    xVinculos(1, 0) = "con_planctas":       xVinculos(1, 1) = "id":          xVinculos(1, 2) = "cuenta,descripcion":
    xVinculos(1, 3) = "Cuenta,Descripcion": xVinculos(1, 4) = "1100,4000":   xVinculos(1, 5) = "C,C":
    xVinculos(1, 6) = "idcuenven":          xVinculos(1, 7) = "descripcion": xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "cuenta"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 22
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_retencion"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Retenciones"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub contabilidad_01_04_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(0, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_libros.* FROM mae_libros"
    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "D"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "5000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    
    'xCamposVista(2, 0) = "CodSun":          xCamposVista(2, 1) = "codsun":    xCamposVista(2, 2) = "800":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    'xCampos(2, 0) = "CodSun":          xCampos(2, 1) = "codsun":        xCampos(2, 3) = "C":    xCampos(2, 2) = "800"
    
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    
'    xVinculos(0, 0) = "mae_impuestos":       xVinculos(0, 1) = "id":            xVinculos(0, 2) = "id,descripcion":
'    xVinculos(0, 3) = "Codigo,Descripcion":  xVinculos(0, 4) = "1000,5000":     xVinculos(0, 5) = "N,C":
'    xVinculos(0, 6) = "idimp":               xVinculos(0, 7) = "descripcion":   xVinculos(0, 8) = "N"
'    xVinculos(0, 9) = "id"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 24
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_libros"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Libros"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub contabilidad_01_05_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(6, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(6, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_documento.*, mae_impuestos.descripcion AS descimp, mae_impuestos.tasa" _
        & " FROM mae_documento LEFT JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id"
    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Cod. Sunat":           xCamposVista(1, 1) = "codsun":         xCamposVista(1, 2) = "1100":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "C"
    xCamposVista(2, 0) = "Descripcion":          xCamposVista(2, 1) = "descripcion":    xCamposVista(2, 2) = "3000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Abreviatura":          xCamposVista(3, 1) = "abrev":          xCamposVista(3, 2) = "1100":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "C"
    xCamposVista(4, 0) = "Impuesto":             xCamposVista(4, 1) = "descimp":        xCamposVista(4, 2) = "3000":   xCamposVista(4, 3) = "C":    xCamposVista(4, 4) = "I"
    xCamposVista(5, 0) = "Tasa":                 xCamposVista(5, 1) = "tasa":           xCamposVista(5, 2) = "1000":   xCamposVista(5, 3) = "C":    xCamposVista(5, 4) = "D"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Abreviatura":    xCampos(2, 1) = "abrev":        xCampos(2, 2) = "C":    xCampos(2, 3) = "1200"
    xCampos(3, 0) = "Cod. Sunat":     xCampos(3, 1) = "codsun":       xCampos(3, 2) = "C":    xCampos(3, 3) = "1200"
    xCampos(4, 0) = "Impuesto":       xCampos(4, 1) = "idimp":        xCampos(4, 2) = "N":    xCampos(4, 3) = "1200"
    xCampos(5, 0) = "Observaciones":  xCampos(5, 1) = "observacion":  xCampos(5, 2) = "M":    xCampos(5, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    
    xVinculos(0, 0) = "mae_impuestos":       xVinculos(0, 1) = "id":            xVinculos(0, 2) = "id,descripcion":
    xVinculos(0, 3) = "Codigo,Descripcion":  xVinculos(0, 4) = "1000,5000":     xVinculos(0, 5) = "N,C":
    xVinculos(0, 6) = "idimp":               xVinculos(0, 7) = "descripcion":   xVinculos(0, 8) = "N"
    xVinculos(0, 9) = "id"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 25
    '-----
        
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_documento"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Documentos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub contabilidad_01_06_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(6, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(6, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_impuestos.*, con_planctas.cuenta, con_planctas.descripcion AS desccuen " _
        & " FROM mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id"

   
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":        xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "3000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Abreviatura":        xCamposVista(2, 1) = "abrev":          xCamposVista(2, 2) = "1100":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "C"
    xCamposVista(3, 0) = "Tasa":               xCamposVista(3, 1) = "tasa":           xCamposVista(3, 2) = "1100":   xCamposVista(3, 3) = "N":    xCamposVista(3, 4) = "R"
    xCamposVista(4, 0) = "Cuenta":             xCamposVista(4, 1) = "cuenta":         xCamposVista(4, 2) = "1100":   xCamposVista(4, 3) = "C":    xCamposVista(4, 4) = "I"
    xCamposVista(5, 0) = "Descripcion Cuenta": xCamposVista(5, 1) = "desccuen":       xCamposVista(5, 2) = "3000":   xCamposVista(5, 3) = "C":    xCamposVista(5, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Abreviatura":    xCampos(2, 1) = "abrev":        xCampos(2, 2) = "C":    xCampos(2, 3) = "1100"
    xCampos(3, 0) = "Tasa":           xCampos(3, 1) = "tasa":         xCampos(3, 2) = "N":    xCampos(3, 3) = "1100"
    xCampos(4, 0) = "Cuenta Compra":  xCampos(4, 1) = "idcuen":       xCampos(4, 2) = "N":    xCampos(4, 3) = "1100"
    xCampos(5, 0) = "Cuenta Venta":   xCampos(5, 1) = "idcuenvta":    xCampos(5, 2) = "N":    xCampos(5, 3) = "1100"
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "con_planctas":       xVinculos(0, 1) = "id":          xVinculos(0, 2) = "cuenta,descripcion":
    xVinculos(0, 3) = "Cuenta,Descripcion": xVinculos(0, 4) = "1100,4000":   xVinculos(0, 5) = "C,C":
    xVinculos(0, 6) = "idcuen":             xVinculos(0, 7) = "descripcion": xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "cuenta"
    
    xVinculos(1, 0) = "con_planctas":       xVinculos(1, 1) = "id":          xVinculos(1, 2) = "cuenta,descripcion":
    xVinculos(1, 3) = "Cuenta,Descripcion": xVinculos(1, 4) = "1100,4000":   xVinculos(1, 5) = "C,C":
    xVinculos(1, 6) = "idcuenvta":          xVinculos(1, 7) = "descripcion": xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "cuenta"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 116
    '-----
        
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_impuestos"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Impuestos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub contabilidad_01_07_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.mantenimiento
    xfrm.IdMenu = 265
    xfrm.Idusuario = xIdUsuario
    xfrm.ManConfigVal xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_03_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.mantenimiento
    xfrm.IdMenu = 26
    xfrm.Idusuario = xIdUsuario
    xfrm.ManPlanCuentas xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_03_02_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.mantenimiento
    xfrm.IdMenu = 36
    xfrm.Idusuario = xIdUsuario
    xfrm.ManCentroCostos xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_03_02_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad2.mantenimientos
    xfrm.ManCentroCostoArea xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_03_03_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.mantenimiento
    xfrm.IdMenu = 117
    xfrm.Idusuario = xIdUsuario
    xfrm.ManCtaDocumento xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_03_04_01_Click()
    Dim xfrm As New sgi2_contabilidad3.mantenimientos
    xfrm.IdMenu = 169
    xfrm.Idusuario = xIdUsuario
    xfrm.ManConcepto xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_03_04_02_Click()
    Dim xfrm As New sgi2_contabilidad3.mantenimientos
    xfrm.IdMenu = 170
    xfrm.Idusuario = xIdUsuario
    xfrm.ManInforme xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_03_05_Click()
    ' EJECUTA MENU
    Dim xfrm As New Sgi2_Procesos.Procesos
    xfrm.CorregirAsiento xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_03_06_Click()
    ' EJECUTA MENU
    FrmCierreMes.Show vbModal
End Sub

Private Sub contabilidad_03_07_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad2.mantenimientos
    xfrm.IdMenu = 205
    xfrm.Idusuario = xIdUsuario
    xfrm.ManTC xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_03_08_Click()
    ' EJECUTA MENU
    Dim xFun As New Sgi2_Procesos.Procesos
    xFun.TransferenciaOperaciones xCon
    Set xFun = Nothing
End Sub

Private Sub contabilidad_04_01_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.mantenimiento
    xfrm.IdMenu = 184
    xfrm.Idusuario = xIdUsuario
    xfrm.ManDetraccion xCon, AP_MESTRA, DET_Compra
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_04_01_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.mantenimiento
    xfrm.IdMenu = 185
    xfrm.Idusuario = xIdUsuario
    xfrm.ManDetraccion xCon, AP_MESTRA, DET_Venta
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_04_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.mantenimiento
    xfrm.IdMenu = 30
    xfrm.Idusuario = xIdUsuario
    xfrm.ManPercepcion xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_04_03_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.mantenimiento
    xfrm.IdMenu = 31
    xfrm.Idusuario = xIdUsuario
    xfrm.ManRetencion xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_04_04_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.mantenimiento
    xfrm.IdMenu = 122
    xfrm.Idusuario = xIdUsuario
    xfrm.ManProviciones xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_05_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.VerRegCompras xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_05_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.VerRegRent4ta xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_05_03_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.VerRegVentas xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_05_05_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.VerDiario xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_05_06_Click()
    ' EJECUTA MENU
    Dim xFor As New sgi2_contabilidad.Consultas
    xFor.Mayor xCon
    Set xFor = Nothing
End Sub

Private Sub contabilidad_05_07_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.HojaTrabajo xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_05_08_Click()
    Dim xfrm As New sgi2_contabilidad3.mantenimientos
    xfrm.VerInforme xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_05_10_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_cajabancos.cajabancos
    xfrm.IdMenu = 123
    xfrm.Idusuario = xIdUsuario
    xfrm.Librobancos xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_05_11_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.MostrarKardexValorizado xCon
    'xfrm.MostrarStockResumen xCon, True
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_05_12_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.mantenimiento
    xfrm.IdMenu = 256
    xfrm.Idusuario = xIdUsuario
    xfrm.verLibroCosto xCon
    Set xfrm = Nothing
End Sub

Private Sub contabilidad_07_01_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad2.Reportes
    xFun.RepDAOT xCon
    Set xFun = Nothing
End Sub

Private Sub contabilidad_07_02_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad2.Reportes
    xFun.CentroCostos xCon
    Set xFun = Nothing
End Sub

Private Sub contabilidad_07_03_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad.Consultas
    xFun.ConsultaCostoMovimiento xCon
    Set xFun = Nothing
End Sub

Private Sub contabilidad_07_04_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad.Consultas
    xFun.ConsultaCostoParte xCon
    Set xFun = Nothing
End Sub

Private Sub contabilidad_07_05_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad.Consultas
    xFun.AnalisisCostoProduccion xCon
    Set xFun = Nothing
End Sub

Private Sub contabilidad_07_06_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad.Consultas
    xFun.ConsultaKardexValResum xCon
    Set xFun = Nothing
End Sub

Private Sub contabilidad_07_07_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad.Consultas
    xFun.InformeKardexVal xCon
    Set xFun = Nothing
End Sub

Private Sub gestion_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_gestion.Compras
    xfrm.AnalizisCompras xCon
    Set xfrm = Nothing
End Sub

Private Sub gestion_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_gestion.ventas
    xfrm.AnalizisVentas xCon
    Set xfrm = Nothing
End Sub

Private Sub gestion_03_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_gestion.produccion
    xfrm.AnalizisProduccion xCon
    Set xfrm = Nothing
End Sub

Private Sub gestion_06_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_gestion.Planeamiento
    xfrm.IdMenu = 193
    xfrm.Idusuario = xIdUsuario
    xfrm.PlanVentasEstimado xCon
    Set xfrm = Nothing
End Sub

Private Sub gestion_06_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_gestion.Planeamiento
    xfrm.IdMenu = 194
    xfrm.Idusuario = xIdUsuario
    xfrm.PlanVentas xCon
    Set xfrm = Nothing
End Sub

Private Sub gestion_06_04_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_gestion.Planeamiento
    xfrm.IdMenu = 195
    xfrm.Idusuario = xIdUsuario
    xfrm.PlanProduccion3 xCon
    Set xfrm = Nothing
End Sub

Private Sub gestion_06_05_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_gestion.Unificados
    xfrm.UnificadoProduccion xCon
    Set xfrm = Nothing
End Sub

Private Sub gestion_06_07_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_gestion.Planeamiento
    xfrm.IdMenu = 196
    xfrm.Idusuario = xIdUsuario
    xfrm.PlanAbastecimiento3 xCon
    Set xfrm = Nothing
End Sub

Private Sub gestion_06_08_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_gestion.Unificados
    xfrm.UnificadoAbastecimiento xCon
    Set xfrm = Nothing
End Sub

Private Sub gestion_06_10_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_gestion.Unificados
    xfrm.UnificadoProducido xCon
    Set xfrm = Nothing
End Sub

Private Sub gestion_07_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_gestion.gestion
    xfrm.ConsultaDocReferencia xCon
    Set xfrm = Nothing
End Sub

Private Sub gestion_08_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_gestion.gestion
    xfrm.ConsultaUnidadNegocio xCon
    Set xfrm = Nothing
End Sub

Private Sub gestion_09_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_gestion.gestion
    xfrm.ConsultaControlRegistro xCon
    Set xfrm = Nothing
End Sub

Private Sub maestros_01_01_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_departamentos.* From mae_departamentos ORDER BY mae_departamentos.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":         xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":    xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "6000"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":       xVinculos(0, 1) = "":          xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":       xVinculos(0, 4) = "":          xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":       xVinculos(0, 7) = "":          xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_departamentos"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Departamentos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub maestros_01_03_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_distritos.* From mae_distritos ORDER BY mae_distritos.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":         xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":    xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "6000"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":       xVinculos(0, 1) = "":          xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":       xVinculos(0, 4) = "":          xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":       xVinculos(0, 7) = "":          xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_distritos"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Distritos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub


Private Sub maestros_02_02_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(4, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_condpago.* From mae_condpago ORDER BY mae_condpago.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "3000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Abreviatura":          xCamposVista(2, 1) = "abrev":          xCamposVista(2, 2) = "3000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Nº Dias":              xCamposVista(3, 1) = "numdia":         xCamposVista(3, 2) = "1000":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "D"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Abreviatura":    xCampos(2, 1) = "abrev":        xCampos(2, 2) = "C":    xCampos(2, 3) = "3000"
    xCampos(3, 0) = "Nº Dias":        xCampos(3, 1) = "numdia":       xCampos(3, 2) = "C":    xCampos(3, 3) = "1000"
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 159
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_condpago"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Condicion de Pago"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub



Private Sub maestros_02_05_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_dociden.* From mae_dociden ORDER BY mae_dociden.id"
    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Cod. Sunat":           xCamposVista(1, 1) = "iddoc":          xCamposVista(1, 2) = "1100":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "C"
    xCamposVista(2, 0) = "Descripcion":          xCamposVista(2, 1) = "descripcion":    xCamposVista(2, 2) = "3000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Cod. Sunat":     xCampos(2, 1) = "iddoc":        xCampos(2, 2) = "C":    xCampos(2, 3) = "1200"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_dociden"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Documentos de Identidad"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub maestros_02_06_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_moneda.* From mae_moneda ORDER BY mae_moneda.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "3000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Abreviatura":          xCamposVista(2, 1) = "simbolo":        xCamposVista(2, 2) = "1100":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "C"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Abreviatura":    xCampos(2, 1) = "simbolo":      xCampos(2, 2) = "C":    xCampos(2, 3) = "1100"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_moneda"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Monedas"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub





Private Sub mantenimiento_01_01_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT man_tareas.* From man_tareas ORDER BY man_tareas.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "id":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Codigo":         xCamposVista(1, 1) = "cod":            xCamposVista(1, 2) = "900":    xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Descripcion":    xCamposVista(2, 1) = "descripcion":    xCamposVista(2, 2) = "6000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Id":             xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "cod":          xCampos(1, 2) = "C":    xCampos(1, 3) = "800"
    xCampos(2, 0) = "Descripcion":    xCampos(2, 1) = "descripcion":  xCampos(2, 2) = "C":    xCampos(2, 3) = "6000"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":       xVinculos(0, 1) = "":          xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":       xVinculos(0, 4) = "":          xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":       xVinculos(0, 7) = "":          xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 154
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "man_tareas"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Id"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Tareas de Mantenimiento"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub mantenimiento_01_02_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu


    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT man_frecuencia.* From man_frecuencia ORDER BY man_frecuencia.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "id":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Codigo":         xCamposVista(1, 1) = "cod":            xCamposVista(1, 2) = "900":    xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Descripcion":    xCamposVista(2, 1) = "descripcion":    xCamposVista(2, 2) = "6000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Id":             xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "cod":          xCampos(1, 2) = "C":    xCampos(1, 3) = "800"
    xCampos(2, 0) = "Descripcion":    xCampos(2, 1) = "descripcion":  xCampos(2, 2) = "C":    xCampos(2, 3) = "6000"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":       xVinculos(0, 1) = "":          xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":       xVinculos(0, 4) = "":          xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":       xVinculos(0, 7) = "":          xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 155
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "man_frecuencia"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Id"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Frecuencia de Mantenimiento"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub mantenimiento_02_Click()
    Dim xfrm As New sgi2_mantenimiento.mantenimiento
    xfrm.IdMenu = 156
    xfrm.Idusuario = xIdUsuario
    xfrm.MantEquipos xCon
    Set xfrm = Nothing
End Sub

Private Sub mantenimiento_03_Click()
    Dim xfrm As New sgi2_mantenimiento.mantenimiento
    xfrm.IdMenu = 166
    xfrm.Idusuario = xIdUsuario
    xfrm.ManPreventivo xCon
    Set xfrm = Nothing
End Sub

Private Sub MDIForm_Activate()
    ' SEGUNDO EVENTO A AJECUTARSE DEL FORMULARIO, AQUI SE VALIDA QUE SE HAYA SELECCIONADO UNA EMPRESA
    If SeEjecutoEmp = False Then
        SeEjecutoEmp = True
        ' INFORMA AL SISTEMA SI HARA LOS PROCESOS CONTABLES
        If CONTABILIZAR = True Then
            AP_MESTRA = Val(Format(Date, "mm"))
            ' CARGA EL STATUS BAR DEL SISTEMA CON EM MES DE TRABAJO ACTUAL
            MDIPrincipal.StatusBar1.Panels(4).Text = "Mes : " + NulosC(Format(Date, "mmmm"))
        Else
            ' MUESTRA EN EN STATUS BAR DEL SISTEMA QUE NO SE REALIZARAN LOS PROCESOS CONTABLES
            MDIPrincipal.StatusBar1.Panels(4).Text = "NO CONTABLE"
        End If
        

    End If
    
    If SeEjecuto = False Then
        SeEjecuto = True
'        '--------------------------------------------
'        '--Ejecutar actualizacion si hay version nueva del seven
'        Dim AP_RUTAPROGRAMA As String
'        AP_RUTAPROGRAMA = AP_RUTASY & "Actualizacion.exe"
'        Call Shell(AP_RUTAPROGRAMA, vbNormalFocus)
'        '--------------------------------------------
        ' CONECTA  A LA BASE DE DATOS DE ENLACE Y NOS PERMITE SELECCIONAR UNA NUEVA EMPRESA DE TRABAJO
        
    End If

    
End Sub

Private Sub MDIForm_Load()
    
    Me.Caption = AP_NOMSIS                                      ' CARGAMOS EN EL CAPTION DEL FORMULARIO EL NOMBRE DEL SISTEMA
    
    ' Se agrega la version al Titulo
    AP_NOMSIS = AP_NOMSIS & " - VERSION: " & AP_VERSION
    
    ActivarMenus
    '***************
    ' Se verifica el inicio de sesion
    AbrirDataEnlace
    FrmSelEmp2.Show vbModal
    '***************
    ' PRIMER EVENTO A EJECUTARSE DEL FORMULARIO, SE DEFINEN EL ATO Y ANCHO DEL FORMULARIO Y SU POSICION INCIAL
    ' EN LA PANTALLA
    SeEjecutoEmp = False
    SeEjecuto = False
    
    Me.Width = 12000
    Me.Height = 8600
    Me.Left = 0
    Me.Top = 0
    
    'Me.Picture = LoadPicture(Trim(AP_RUTABM) + "system.bmp")   ' CARGAMOS EL FONDO DEL FORMULARIO
    On Error Resume Next
    Me.Picture = LoadPicture(Trim(AP_RUTABM) + "system - copia.bmp")
                                              ' ACTIVAMOS LOS MENUS DEL SISTEMA
End Sub

'Private Sub menu10_11_Click()
'    Dim xFun As New Sgi2_Procesos.Procesos
'    xFun.PersonalCompras xCon
'    Set xFun = Nothing
'End Sub

'Private Sub menu10_12_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New Sgi2_Procesos.Procesos
'    xFrm.PersonalProduccion xCon
'    Set xFrm = Nothing
'End Sub

'Private Sub menu10_14_1_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New SGI2_almacen.almacen
'    xFrm.MantItem xCon, 2
'    Set xFrm = Nothing
'End Sub

'Private Sub menu10_14_2_Click()
'    ' EJECUTA MENU
'    Dim xFun As New SGI2_almacen.almacen
'    xFun.ManUnidades xCon, 2
'    Set xFun = Nothing
'End Sub

'Private Sub menu10_14_3_Click()
'    ' EJECUTA MENU
'    Dim xFun As New SGI2_almacen.almacen
'    xFun.ManTipoProducto xCon, 2
'    Set xFun = Nothing
'End Sub

'Private Sub menu10_14_4_Click()
'    ' EJECUTA MENU
'    Dim xFun As New SGI2_almacen.almacen
'    xFun.ManFamilia xCon, 2
'    Set xFun = Nothing
'End Sub

'Private Sub menu10_14_5_Click()
'    ' EJECUTA MENU
'    Dim xFun As New SGI2_almacen.almacen
'    xFun.ManClase xCon, 2
'    Set xFun = Nothing
'End Sub

'Private Sub menu10_14_6_Click()
'    ' EJECUTA MENU
'    Dim xFun As New SGI2_almacen.almacen
'    xFun.ManSubClase xCon, 2
'    Set xFun = Nothing
'End Sub

'Private Sub menu10_14_8_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New sgi2_ventas.ventas
'    xFrm.Clientes xCon, 2
'    Set xFrm = Nothing
'End Sub

'Private Sub menu10_14_9_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New sgi2_compras.Compras
'    xFrm.ManProveedor xCon, 2
'    Set xFrm = Nothing
'End Sub

'Private Sub menu10_8_1_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New Sgi2_Procesos.Procesos
'    xFrm.InventarioCobranza xCon
'    Set xFrm = Nothing
'End Sub
'
'Private Sub menu10_8_2_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New Sgi2_Procesos.Procesos
'    xFrm.InventarioPagos xCon
'    Set xFrm = Nothing
'End Sub
'
'Private Sub menu10_9_1_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New Sgi2_Procesos.Procesos
'    xFrm.CargarClientes xCon
'    Set xFrm = Nothing
'End Sub
'
'Private Sub menu10_9_2_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New Sgi2_Procesos.Procesos
'    xFrm.CargarProveedores xCon
'    Set xFrm = Nothing
'End Sub
'
'Private Sub menu10_9_3_Click()
'    ' EJECUTA MENU
'    Dim xFun As New Sgi2_Procesos.Procesos
'    xFun.CargarPlandeCuentas xCon
'    Set xFun = Nothing
'End Sub
'
'Private Sub menu10_9_6_1_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New Sgi2_Procesos.Procesos
'    xFrm.CargarCompras xCon
'    Set xFrm = Nothing
'End Sub
'
'Private Sub menu10_9_6_2_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New Sgi2_Procesos.Procesos
'    xFrm.CargarCompras2 xCon
'    Set xFrm = Nothing
'End Sub
'
'Private Sub menu10_9_7_1_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New Sgi2_Procesos.Procesos
'    xFrm.CargarVentas xCon
'    Set xFrm = Nothing
'End Sub
'
'Private Sub menu10_9_7_2_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New Sgi2_Procesos.Procesos
'    xFrm.CargarVentasEstudio xCon
'    Set xFrm = Nothing
'End Sub

'Function CargarTabla(IdMantenimiento As Integer) As ADODB.Recordset
'    Dim RstCab As New ADODB.Recordset
'
'    RST_Busq RstCab, "SELECT * FROM mae_manformularios WHERE id = " & IdMantenimiento & "", xCon
'    If RstCab.RecordCount = 0 Then
'        Set CargarTabla = Nothing
'    Else
'        Set CargarTabla = RstCab
'    End If
'End Function
'
'Function CargarTablaCampos(IdMantenimiento As Integer) As ADODB.Recordset
'    Dim RstCab As New ADODB.Recordset
'
'    RST_Busq RstCab, "SELECT * FROM mae_manformulariosdet WHERE id = " & IdMantenimiento & " ORDER BY corr", xCon
'    If RstCab.RecordCount = 0 Then
'        Set CargarTablaCampos = Nothing
'    Else
'        Set CargarTablaCampos = RstCab
'    End If
'End Function

'Private Sub menu44_28_1_Click()
'    ' EJECUTA MENU
'    Dim xform As New Eps_MantTablas.mantenimiento
'    Dim xNivelUsuario As Integer
'    Dim xCampos(4, 4) As String
'    Dim xVinculos(1, 10) As String
'    Dim xCampoBusca(2) As String
'    Dim xCamposVista(5, 4) As String
'    Dim xConsulta As String
'
'    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
'    xConsulta = "SELECT var_codificacion.*, var_formatos.formato FROM var_codificacion LEFT JOIN var_formatos ON var_codificacion.iddato = var_formatos.id" _
'        & " ORDER BY var_codificacion.orden"
'
'    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
'    xCamposVista(0, 0) = "Codigo":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":   xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "C"
'    xCamposVista(1, 0) = "Orden":              xCamposVista(1, 1) = "orden":          xCamposVista(1, 2) = "1000":   xCamposVista(1, 3) = "N":    xCamposVista(1, 4) = "I"
'    xCamposVista(2, 0) = "Descripcion":        xCamposVista(2, 1) = "descripcion":    xCamposVista(2, 2) = "4000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
'    xCamposVista(3, 0) = "Formato":            xCamposVista(3, 1) = "formato":        xCamposVista(3, 2) = "1200":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "I"
'    xCamposVista(4, 0) = "Activo":             xCamposVista(4, 1) = "activo":         xCamposVista(4, 2) = "1200":   xCamposVista(4, 3) = "N":    xCamposVista(4, 4) = "C"
'
'    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
'    xCampos(0, 0) = "Codigo":          xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "900"
'    xCampos(1, 0) = "Descripcion":     xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
'    xCampos(2, 0) = "Formato":         xCampos(2, 1) = "iddato":       xCampos(2, 2) = "N":    xCampos(2, 3) = "1200"
'    xCampos(3, 0) = "Orden":           xCampos(3, 1) = "orden":        xCampos(3, 2) = "N":    xCampos(3, 3) = "1000"
'
'    '0 = nombre de la tabla
'    '1 = nombre del campo con el que se iniciara la busqueda
'    '2 = lista de campos
'    '3 = lista de rotulos para los campos
'    '4 = tamaño de los campos
'    '5 = tipo de los campos
'    '6 = nombre del campo con el que se vincula el array anterior
'    '7 = campo que devolvera la busqueda
'    '8 = tipo del campo que iniciara la busqueda
'    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
''    xVinculos(0, 0) = "mae_doccajabantipo":  xVinculos(0, 1) = "id":          xVinculos(0, 2) = "id,descripcion":
''    xVinculos(0, 3) = "Codigo,Descripcion":  xVinculos(0, 4) = "1100,4000":   xVinculos(0, 5) = "N,C":
''    xVinculos(0, 6) = "tipo":                xVinculos(0, 7) = "descripcion": xVinculos(0, 8) = "N":
''    xVinculos(0, 9) = "id"
'
'    xVinculos(0, 0) = "var_formatos":                xVinculos(0, 1) = "id":               xVinculos(0, 2) = "id,descripcion,formato":
'    xVinculos(0, 3) = "Codigo,Descripcion,Formato":  xVinculos(0, 4) = "1100,4000,1100":   xVinculos(0, 5) = "N,C,C":
'    xVinculos(0, 6) = "iddato":                      xVinculos(0, 7) = "descripcion":      xVinculos(0, 8) = "N":
'    xVinculos(0, 9) = "formato"
'
'    'CAMPOS PARA EFECTUAR LA BUSQUEDA
'    xCampoBusca(0) = "descripcion"
'    xCampoBusca(1) = "id"
'
'    If xNivelUsuario = 0 Then
'        xform.PermiteActualiza = False
'    Else
'        xform.PermiteActualiza = True
'    End If
'
'    xform.CadSQLVista = xConsulta
'    xform.Tabla = "var_codificacion"
'    xform.CampoOrdenado = "orden"
'    xform.CampoClave = "Codigo"
'    xform.PermiteActualiza = True
'    xform.TituloFormulario = "Mantenimiento - Codificacion"
'    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
'End Sub

'Private Sub menu44_28_2_Click()
'    ' EJECUTA MENU
'    Dim xform As New Eps_MantTablas.mantenimiento
'    Dim xNivelUsuario As Integer
'    Dim xCampos(3, 4) As String
'    Dim xVinculos(1, 10) As String
'    Dim xCampoBusca(2) As String
'    Dim xCamposVista(3, 4) As String
'    Dim xConsulta As String
'
'    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
'    xConsulta = "SELECT var_formatos.* FROM var_formatos ORDER BY id"
'
'    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
'    xCamposVista(0, 0) = "Codigo":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":   xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "C"
'    xCamposVista(1, 0) = "Descripcion":        xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "4000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
'    xCamposVista(2, 0) = "Formato":            xCamposVista(2, 1) = "formato":        xCamposVista(2, 2) = "1200":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
'
'    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
'    xCampos(0, 0) = "Codigo":          xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "900"
'    xCampos(1, 0) = "Descripcion":     xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
'    xCampos(2, 0) = "Formato":         xCampos(2, 1) = "formato":      xCampos(2, 2) = "C":    xCampos(2, 3) = "1200"
'
'    '0 = nombre de la tabla
'    '1 = nombre del campo con el que se iniciara la busqueda
'    '2 = lista de campos
'    '3 = lista de rotulos para los campos
'    '4 = tamaño de los campos
'    '5 = tipo de los campos
'    '6 = nombre del campo con el que se vincula el array anterior
'    '7 = campo que devolvera la busqueda
'    '8 = tipo del campo que iniciara la busqueda
'    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
''    xVinculos(0, 0) = "var_formatos":                xVinculos(0, 1) = "id":               xVinculos(0, 2) = "id,descripcion,formato":
''    xVinculos(0, 3) = "Codigo,Descripcion,Formato":  xVinculos(0, 4) = "1100,4000,1100":   xVinculos(0, 5) = "N,C,C":
''    xVinculos(0, 6) = "iddato":                      xVinculos(0, 7) = "descripcion":      xVinculos(0, 8) = "N":
''    xVinculos(0, 9) = "formato"
'
'    'CAMPOS PARA EFECTUAR LA BUSQUEDA
'    xCampoBusca(0) = "descripcion"
'    xCampoBusca(1) = "id"
'
'    If xNivelUsuario = 0 Then
'        xform.PermiteActualiza = False
'    Else
'        xform.PermiteActualiza = True
'    End If
'
'    xform.CadSQLVista = xConsulta
'    xform.Tabla = "var_formatos"
'    xform.CampoOrdenado = "orden"
'    xform.CampoClave = "Codigo"
'    xform.PermiteActualiza = True
'    xform.TituloFormulario = "Mantenimiento - Formato para Codificacion"
'    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
'End Sub

'Private Sub menu44_29_Click()
'    ' EJECUTA MENU
''    Dim xFrm As New sgi2_contabilidad2.estadosfinancieros
''    xFrm.BalanceGeneral xCon
''    Set xFrm = Nothing
'    Dim xFrm As New sgi2_contabilidad3.mantenimientos
'    xFrm.VerInforme xCon
'    Set xFrm = Nothing
'End Sub


'Private Sub menu44_30_1_Click()
'    ' EJECUTA MENU
''    Dim xFrm As New sgi2_contabilidad2.mantenimientos
''    xFrm.ManBalance xCon
''    Set xFrm = Nothing
'
'    Dim xFrm As New sgi2_contabilidad3.mantenimientos
'    xFrm.ManConcepto xCon
'    Set xFrm = Nothing
'End Sub

'Private Sub menu44_30_2_Click()
'    ' EJECUTA MENU
''    Dim xFrm As New sgi2_contabilidad2.mantenimientos
''    xFrm.ManEstados xCon
''    Set xFrm = Nothing
'    Dim xFrm As New sgi2_contabilidad3.mantenimientos
'    xFrm.ManInforme xCon
'    Set xFrm = Nothing
'End Sub

'Private Sub menu55_14_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New sgi2_cajabancos.cajabancos
'    xFrm.ProgramarPagos xCon, xIdUsuario, AP_MESTRA
'    Set xFrm = Nothing
'End Sub
'
'Private Sub menu55_15_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New sgi2_cajabancos.cajabancos
'    xFrm.EmiRendir xCon, xIdUsuario, AP_MESTRA
'    Set xFrm = Nothing
'End Sub
'
'Private Sub menu55_16_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New sgi2_cajabancos.cajabancos
'    xFrm.DevRendir xCon, xIdUsuario, AP_MESTRA
'    Set xFrm = Nothing
'End Sub

'Private Sub menu66_4_Click()
'    ' EJECUTA MENU
'    Dim xFrm As New sgi2_produccion.produccion
'    xFrm.ProgramaProduccion xCon, AP_MESTRA
'    Set xFrm = Nothing
'End Sub

'Private Sub menu88_16_Click()
'    ' EJECUTA MENU
''    Dim xFrm As New sgi2_planillas2.planillas
''    xFrm.Permiso xCon, AP_RUTDATTRA, AP_MESTRA
''    Set xFrm = Nothing
'End Sub

Private Sub planillas_01_01_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_area.* From mae_area ORDER BY mae_area.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":    xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 71
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_area"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Areas"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub planillas_01_02_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_cargo.* From mae_cargo ORDER BY mae_cargo.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":    xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 72
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_cargo"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Cargos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub planillas_01_03_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas3.planillas
    xfrm.IdMenu = 73
    xfrm.Idusuario = xIdUsuario
    xfrm.ManConcepto xCon, AP_RUTDATTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_01_04_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas3.planillas
    xfrm.IdMenu = 58
    xfrm.Idusuario = xIdUsuario
    xfrm.ManRegimenPension xCon, AP_RUTDATTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_01_05_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas.planillas
    xfrm.IdMenu = 59
    xfrm.Idusuario = xIdUsuario
    xfrm.ManNomina xCon, AP_RUTDATTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_01_07_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas2.planillas
    xfrm.IdMenu = 57
    xfrm.Idusuario = xIdUsuario
    xfrm.ManHorario xCon, AP_RUTDATTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_01_08_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas2.planillas
    xfrm.IdMenu = 74
    xfrm.Idusuario = xIdUsuario
    xfrm.ManTipoLicencia xCon, AP_RUTDATTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_01_09_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas2.planillas
    xfrm.IdMenu = 75
    xfrm.Idusuario = xIdUsuario
    xfrm.ManTipoPermiso xCon, AP_RUTDATTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_01_10_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_planillas2.planillas
    xFun.IdMenu = 76
    xFun.Idusuario = xIdUsuario
    xFun.ManDiasFestivos xCon, AP_RUTDATTRA
    Set xFun = Nothing
End Sub

Private Sub planillas_02_01_01_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_planillas2.planillas
    xFun.IdMenu = 250
    xFun.Idusuario = xIdUsuario
    xFun.Asistencia xCon, AP_RUTDATTRA, AP_MESTRA
    Set xFun = Nothing
End Sub

Private Sub planillas_02_01_02_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_produccion.produccion
    xFun.IdMenu = 251
    xFun.Idusuario = xIdUsuario
    xFun.TempusImportarMarcacion xCon
    Set xFun = Nothing
End Sub

Private Sub planillas_02_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas2.planillas
    xfrm.IdMenu = 79
    xfrm.Idusuario = xIdUsuario
    xfrm.Licencia xCon, AP_RUTDATTRA, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_02_03_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas2.planillas
    xfrm.IdMenu = 80
    xfrm.Idusuario = xIdUsuario
    xfrm.Permiso xCon, AP_RUTDATTRA, AP_MESTRA
    Set xfrm = Nothing

End Sub

Private Sub planillas_02_04_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas2.planillas
    xfrm.IdMenu = 81
    xfrm.Idusuario = xIdUsuario
    xfrm.Vacaciones xCon, AP_RUTDATTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_04_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas2.planillas
    xfrm.IdMenu = 85
    xfrm.Idusuario = xIdUsuario
    xfrm.ResumenHoras xCon, AP_RUTDATTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_05_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas3.planillas
    xfrm.IdMenu = 104
    xfrm.Idusuario = xIdUsuario
    xfrm.AsignarSueldo xCon, AP_RUTDATTRA, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_06_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas3.planillas
    xfrm.IdMenu = 105
    xfrm.Idusuario = xIdUsuario
    xfrm.BoletaPago xCon, AP_RUTDATTRA, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_07_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_planillas3.planillas
    xFun.IdMenu = 82
    xFun.Idusuario = xIdUsuario
    xFun.Comision xCon
    Set xFun = Nothing
End Sub

Private Sub planillas_08_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas3.planillas
    xfrm.ImprimirBoletas xCon, AP_RUTDATTRA, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_10_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.CostoProduccion xCon
    Set xfrm = Nothing
End Sub

Private Sub planillas_11_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.RepCosto xCon
    Set xfrm = Nothing
End Sub

Private Sub planillas_13_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas.planillas
    xfrm.ConsAsistencia xCon, AP_RUTDATTRA, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_13_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas3.planillas
    xfrm.ConsPlanilla xCon, AP_RUTDATTRA, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub planillas_13_03_01_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_planillas.planillas
    xFun.ExportarNomina xCon, AP_RUTDATTRA
    Set xFun = Nothing
End Sub

Private Sub planillas_13_04_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_planillas.planillas
    xfrm.Idusuario = xIdUsuario
    xfrm.ConsRotacion xCon, AP_RUTDATTRA
    Set xfrm = Nothing
End Sub

Private Sub produccion_01_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.IdMenu = 47
    xfrm.Idusuario = xIdUsuario
    xfrm.ManTareas xCon
    Set xfrm = Nothing
End Sub

Private Sub produccion_01_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.IdMenu = 48
    xfrm.Idusuario = xIdUsuario
    xfrm.MamRecetas xCon
    Set xfrm = Nothing
End Sub

Private Sub produccion_01_03_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.IdMenu = 49
    xfrm.Idusuario = xIdUsuario
    xfrm.Estacionalidad xCon
    Set xfrm = Nothing
End Sub

Private Sub produccion_01_04_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.IdMenu = 50
    xfrm.Idusuario = xIdUsuario
    xfrm.ConfigurarCosto xCon
    Set xfrm = Nothing
End Sub

Private Sub produccion_01_05_Click()
    Dim xFun As New sgi2_produccion.produccion
    xFun.IdMenu = 149
    xFun.Idusuario = xIdUsuario
    xFun.Rendimiento xCon
    Set xFun = Nothing
End Sub

Private Sub produccion_01_06_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.IdMenu = 241
    xfrm.Idusuario = xIdUsuario
    xfrm.ConfigurarPersonalxTareas xCon
    Set xfrm = Nothing
End Sub

Private Sub produccion_01_07_Click()
    Dim xFun As New sgi2_produccion.produccion
    xFun.IdMenu = 242
    xFun.Idusuario = xIdUsuario
    xFun.CronogramaMantLinea xCon
    Set xFun = Nothing
End Sub

Private Sub produccion_03_01_Click()
'    Dim xFun As New sgi2_produccion.produccion
'    xFun.IdMenu = 51
'    xFun.Idusuario = xIdUsuario
'    xFun.CronogramaProduccion xCon
'    Set xFun = Nothing
End Sub

Private Sub produccion_03_02_Click()
    Dim xFun As New sgi2_produccion.produccion
    xFun.IdMenu = 52
    xFun.Idusuario = xIdUsuario
    xFun.GenOrdenProd xCon, AP_MESTRA
    Set xFun = Nothing
End Sub

Private Sub produccion_03_03_Click()
    Dim xFun As New sgi2_produccion.produccion
    xFun.IdMenu = 53
    xFun.Idusuario = xIdUsuario
    xFun.LineaDeTiempo xCon
    Set xFun = Nothing
End Sub

Private Sub produccion_03_04_Click()
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.IdMenu = 54
    xfrm.Idusuario = xIdUsuario
    xfrm.GenSolicitudMat xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub produccion_04_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.IdMenu = 92
    xfrm.Idusuario = xIdUsuario
    xfrm.OrdenProduccion xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub produccion_05_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.IdMenu = 106
    xfrm.Idusuario = xIdUsuario
    xfrm.ManGrupos xCon
    Set xfrm = Nothing
End Sub

Private Sub produccion_05_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.IdMenu = 118
    xfrm.Idusuario = xIdUsuario
    xfrm.PersonalProduccion xCon
    Set xfrm = Nothing
End Sub

Private Sub produccion_05_03_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.IdMenu = 140
    xfrm.Idusuario = xIdUsuario
    xfrm.DistribucionTareas xCon
    Set xfrm = Nothing
End Sub

Private Sub produccion_05_04_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.IdMenu = 179
    xfrm.Idusuario = xIdUsuario
    xfrm.IngresoTareas xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub produccion_05_05_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.RepTarea xCon
    Set xfrm = Nothing
End Sub

Private Sub produccion_07_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.RepProduccion xCon
    Set xfrm = Nothing
End Sub

Private Sub produccion_07_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.IdMenu = 243
    xfrm.Idusuario = xIdUsuario
    xfrm.RepLinea xCon
    Set xfrm = Nothing
End Sub

Private Sub produccion_08_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.RepCompProduccion xCon
    Set xfrm = Nothing
End Sub

Private Sub produccion_09_Click()
    Dim xfrm As New sgi2_produccion.produccion
    xfrm.RepPlaneacion xCon
    Set xfrm = Nothing
End Sub

Private Sub setup_01_Click()
    ' EJECUTA MENU
    FrmMantEmpresa1.Show vbModal
End Sub

Private Sub setup_02_Click()
    FrmManRutasRutas.Show vbModal
End Sub

''Private Sub setup_03_Click()
''    ' EJECUTA MENU
''    FrmVincularData.Show vbModal
''End Sub

Private Sub setup_04_Click()
    ' EJECUTA MENU
    FrmSetup.Show vbModal
End Sub

Private Sub setup_05_Click()
    ' EJECUTA MENU
    FrmManUsuarios.Show vbModal
End Sub

Private Sub setup_06_Click()
    ' EJECUTA MENU
    FrmManOpcionesUsuario.Show vbModal
End Sub

Private Sub setup_07_Click()
    ' EJECUTA MENU
    Dim xfrm As New Sgi2_Procesos.Procesos
    xfrm.ConfiguraImpresion xCon
    Set xfrm = Nothing
End Sub

Private Sub setup_09_Click()
    Dim xfrm As New eps_librerias.browser
    Dim xDir As String
    xDir = "file:///" & AP_RUTAAR & "0003\Index.htm"
    xfrm.Navegador xDir
    Set xfrm = Nothing
End Sub

Private Sub setup_10_Click()

'    Dim xfrm As New Sgi2_Procesos.Procesos
'    xfrm.BDEvaluar xCon
'    Set xfrm = Nothing
    
    ' EJECUTA MENU
    xCon.Close
    Set xCon = Nothing
    Unload Me
    End
End Sub

Private Sub setup_11_Click()
    Dim xfrm As New eps_librerias.browser
    Dim xDir As String
    xDir = "file:///" & AP_RUTAAR & "0001\Indice.htm"
    xfrm.Navegador xDir
    Set xfrm = Nothing
End Sub

Private Sub setup_12_Click()
    '**************************************
    ' Modificado 29/05/2012 - Jose Chacon
    '**************************************
    ' EJECUTA MENU
    Dim xfrm As New sgi2_etiquetas.etiquetas
    xfrm.IdMenu = 84
    xfrm.Idusuario = xIdUsuario
    xfrm.manEtiquetas xCon
    Set xfrm = Nothing
    
'    FrmEtiquetas.Show vbModal
    'Dim xfrm As New Sgi2_Procesos.Procesos
    'xfrm.BDEvaluar xCon
    'Set xfrm = Nothing
End Sub

Private Sub setup_13_01_Click()
    ' EJECUTA MENU
    Dim xFun As New Sgi2_Procesos.Procesos
    xFun.CargarClientes xCon
    Set xFun = Nothing
End Sub

Private Sub setup_13_02_Click()
    ' EJECUTA MENU
    Dim xFun As New Sgi2_Procesos.Procesos
    xFun.CargarProveedores xCon
    Set xFun = Nothing
End Sub

Private Sub setup_13_03_Click()
    ' EJECUTA MENU
    Dim xFun As New Sgi2_Procesos.Procesos
    xFun.CargarCompras xCon
    Set xFun = Nothing
End Sub

Private Sub setup_13_04_Click()
    ' EJECUTA MENU
'    Dim xFun As New Sgi2_Procesos.Procesos
'    xFun.CargarHonorarios xCon
'    Set xFun = Nothing
End Sub

Private Sub setup_13_05_Click()
    ' EJECUTA MENU
    Dim xFun As New Sgi2_Procesos.Procesos
    xFun.CargarVentas xCon
    Set xFun = Nothing
End Sub

Private Sub setup_14_Click()
    ' EJECUTA MENU
    FrmRegLicencia.Show vbModal
End Sub

Private Sub setup_16_Click()
    ' EJECUTA MENU
    FrmMantActualizar.Show vbModal
End Sub

Private Sub setup15_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_Transferencia.Transferencia
    xfrm.IdMenu = 245
    xfrm.Idusuario = xIdUsuario
    xfrm.ManRuc xCon
    Set xfrm = Nothing
End Sub

Private Sub setup15_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_Transferencia.Transferencia
    xfrm.IdMenu = 246
    xfrm.Idusuario = xIdUsuario
    xfrm.ManItems xCon
    Set xfrm = Nothing
End Sub

Private Sub setup15_03_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_Transferencia.Transferencia
    xfrm.IdMenu = 247
    xfrm.Idusuario = xIdUsuario
    xfrm.ManTipoDoc xCon
    Set xfrm = Nothing
End Sub

Private Sub setup15_04_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_Transferencia.Transferencia
    xfrm.IdMenu = 248
    xfrm.Idusuario = xIdUsuario
    xfrm.Documentos xCon
    Set xfrm = Nothing
End Sub

Private Sub setup17_Click()
    ' EJECUTA MENU
    FrmManMenu.Show vbModal
End Sub

Private Sub tesoreria_01_01_01_Click()
    Dim xFun As New sgi2_contabilidad2.mantenimientos
    xFun.IdMenu = 128
    xFun.Idusuario = xIdUsuario
    xFun.ManOrigenes 1, xCon
    Set xFun = Nothing
End Sub

Private Sub tesoreria_01_01_02_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad2.mantenimientos
    xFun.IdMenu = 129
    xFun.Idusuario = xIdUsuario
    xFun.ManOrigenes 2, xCon
    Set xFun = Nothing
End Sub

Private Sub tesoreria_01_02_01_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad2.mantenimientos
    xFun.IdMenu = 130
    xFun.Idusuario = xIdUsuario
    xFun.ManDestinos 1, xCon
    Set xFun = Nothing
End Sub

Private Sub tesoreria_01_02_02_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad2.mantenimientos
    xFun.IdMenu = 131
    xFun.Idusuario = xIdUsuario
    xFun.ManDestinos 2, xCon
    Set xFun = Nothing
End Sub

Private Sub tesoreria_01_03_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_bancos.* From mae_bancos ORDER BY mae_bancos.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":    xCamposVista(2, 3) = "C":    xCamposVista(1, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 40
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_bancos"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Bancos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub tesoreria_01_04_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(5, 4) As String
    Dim xVinculos(3, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(6, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_banconumcta.*, mae_bancos.descripcion AS descban, con_planctas.cuenta, " _
        & " con_planctas.descripcion AS desccuen, mae_moneda.simbolo AS moneda " _
        & " FROM (con_planctas RIGHT JOIN (mae_bancos RIGHT JOIN mae_banconumcta ON mae_bancos.id = mae_banconumcta.idban) " _
        & " ON con_planctas.id = mae_banconumcta.idcuen) LEFT JOIN mae_moneda ON mae_banconumcta.idmon = mae_moneda.id where mae_banconumcta.id <>0"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":              xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "700":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Banco":               xCamposVista(1, 1) = "descban":        xCamposVista(1, 2) = "2500":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Nº Cuenta":           xCamposVista(2, 1) = "numcue":         xCamposVista(2, 2) = "1500":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "M":                   xCamposVista(3, 1) = "moneda":         xCamposVista(3, 2) = "550":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "C"
    xCamposVista(4, 0) = "Cta. Contable":       xCamposVista(4, 1) = "cuenta":         xCamposVista(4, 2) = "1200":   xCamposVista(4, 3) = "C":    xCamposVista(4, 4) = "I"
    xCamposVista(5, 0) = "Descripcion Cuenta":  xCamposVista(5, 1) = "desccuen":       xCamposVista(5, 2) = "4000":   xCamposVista(5, 3) = "C":    xCamposVista(5, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Banco":          xCampos(1, 1) = "idban":        xCampos(1, 2) = "N":    xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Cuenta":         xCampos(2, 1) = "numcue":       xCampos(2, 2) = "C":    xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Moneda":         xCampos(3, 1) = "idmon":        xCampos(3, 2) = "N":    xCampos(3, 3) = "1000"
    xCampos(4, 0) = "Cuenta":         xCampos(4, 1) = "idcuen":       xCampos(4, 2) = "N":    xCampos(4, 3) = "1500"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "mae_bancos":            xVinculos(0, 1) = "id":            xVinculos(0, 2) = "id,descripcion":
    xVinculos(0, 3) = "Codigo,Descripcion":    xVinculos(0, 4) = "1000,2000":     xVinculos(0, 5) = "N,C":
    xVinculos(0, 6) = "idban":                 xVinculos(0, 7) = "descripcion":   xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "id"
    
    xVinculos(1, 0) = "mae_moneda":            xVinculos(1, 1) = "id":            xVinculos(1, 2) = "id,descripcion":
    xVinculos(1, 3) = "Codigo,Descripcion":    xVinculos(1, 4) = "1000,2000":     xVinculos(1, 5) = "N,C":
    xVinculos(1, 6) = "idmon":                 xVinculos(1, 7) = "descripcion":   xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "id"

    xVinculos(2, 0) = "con_planctas":          xVinculos(2, 1) = "id":            xVinculos(2, 2) = "cuenta,descripcion":
    xVinculos(2, 3) = "Nº Cuenta,Descripcion": xVinculos(2, 4) = "1500,3000":     xVinculos(2, 5) = "C,C":
    xVinculos(2, 6) = "idcuen":                xVinculos(2, 7) = "descripcion":   xVinculos(2, 8) = "N":
    xVinculos(2, 9) = "cuenta"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 41
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_banconumcta"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Cuentas de Banco"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub tesoreria_01_05_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(4, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_doccajaban.*, mae_doccajabantipo.descripcion AS desctipo " _
        & " FROM mae_doccajabantipo INNER JOIN mae_doccajaban ON mae_doccajabantipo.id = mae_doccajaban.tipo"
   
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":   xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":        xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "4000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Abreviatura":        xCamposVista(2, 1) = "abrev":          xCamposVista(2, 2) = "1200":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Tipo":               xCamposVista(3, 1) = "desctipo":       xCamposVista(3, 2) = "1200":   xCamposVista(3, 3) = "N":    xCamposVista(3, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":          xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "900"
    xCampos(1, 0) = "Descripcion":     xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Abreviatura":     xCampos(2, 1) = "abrev":        xCampos(2, 2) = "C":    xCampos(2, 3) = "1300"
    xCampos(3, 0) = "Tipo":            xCampos(3, 1) = "tipo":         xCampos(3, 2) = "N":    xCampos(3, 3) = "1100"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "mae_doccajabantipo":  xVinculos(0, 1) = "id":          xVinculos(0, 2) = "id,descripcion":
    xVinculos(0, 3) = "Codigo,Descripcion":  xVinculos(0, 4) = "1100,4000":   xVinculos(0, 5) = "N,C":
    xVinculos(0, 6) = "tipo":                xVinculos(0, 7) = "descripcion": xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "id"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 42
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_doccajaban"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Documentos Caja y Bancos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub tesoreria_01_06_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Maestro Conceptos de Abonos y Cargos de Banco
    '               Evento que invoca a librería Eps_ManTablas; Da mantenimiento a la tabla  mae_bancoscarabo,
    '               esta tabla actualmente no se utiliza; Razon por la cual se da de baja del menu principal.

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_bancoscarabo.* From mae_bancoscarabo ORDER BY mae_bancoscarabo.id"
    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":    xCamposVista(2, 3) = "C":    xCamposVista(1, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_bancoscarabo"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Conceptos Cargos y Abonos de Bancos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub



Private Sub tesoreria_01_07_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Evento que invoca al mantenimiento de las monedas, Se observa que existe duplicidad con
    '               módulo maestros menu Maestro de monedas se concluye que se elimina el menu de tesoreria
    '               quedando vigente el de modulo de maestros

End Sub

Private Sub tesoreria_01_09_Click()
    ' EJECUTA MENU
    Dim xfrm As New Sgi2_Procesos.Procesos
    xfrm.PersonalTesoreria xCon
    Set xfrm = Nothing
End Sub

Private Sub tesoreria_03_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_cajabancos.cajabancos
    xfrm.IdMenu = 43
    xfrm.Idusuario = xIdUsuario
    xfrm.IngresoCajaBanco2 xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub tesoreria_04_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_cajabancos.cajabancos
    xfrm.IdMenu = 44
    xfrm.Idusuario = xIdUsuario
    xfrm.EgresoCajaBanco2 xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub tesoreria_05_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_cajabancos.cajabancos
    xfrm.IdMenu = 136
    xfrm.Idusuario = xIdUsuario
    xfrm.CanjeDocumentos xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub tesoreria_06_01_Click()
    ' EJECUTA MENU
    Dim xLet As New sgi2_letras.letras
    xLet.IdMenu = 188
    xLet.Idusuario = xIdUsuario
    xLet.ManLetras AP_MESTRA, xCon
    Set xLet = Nothing
End Sub

Private Sub tesoreria_06_02_Click()
    ' EJECUTA MENU
    Dim xLet As New sgi2_letras.letras
    xLet.IdMenu = 137
    xLet.Idusuario = xIdUsuario
    xLet.ManPlanilla AP_MESTRA, xCon
    Set xLet = Nothing
End Sub

Private Sub tesoreria_08_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_cajabancos.cajabancos
    xfrm.ConsultaCtaCte xCon
    Set xfrm = Nothing
End Sub

Private Sub tesoreria_08_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_cajabancos2.analisis
    xfrm.AnalisisCtaCte xCon
    Set xfrm = Nothing
End Sub

Private Sub tesoreria_08_03_Click()
    Dim xfrm As New sgi2_cajabancos.cajabancos
    xfrm.Anticuamiento xCon
    Set xfrm = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    'EJECUTO LAS OPCIONES LDE TOOLBAR EL FORMULARIO
    If Button.Index = 1 Then
        ' CONECTA  A LA BASE DE DATOS DE ENLACE Y NOS PERMITE SELECCIONAR UNA NUEVA EMPRESA DE TRABAJO
        AbrirDataEnlace
        FrmSelEmp2.Show vbModal
        If xIdEmpresa <> 0 Then
        End If
    End If
    
    If Button.Index = 2 Then
        ' LLAMA AL FORMULARIO DE USUARIOS
        'FrmManUsuarios.Show vbModal
    End If
    
    If Button.Index = 3 Then
        ' PERMITE CAMBIAS EL MES DE TRABAJO
        Dim xMesPro As Integer
        xMesPro = AP_MESTRA
        AP_MESTRA = SeleccionaMes(xCon)
        
        If AP_MESTRA = 0 Then
            AP_MESTRA = xMesPro
        End If
        
        ' BUSCAR EL NOMBRE DEL MES DE TRABAJO ACTUAL
        Dim NomMes As String
        NomMes = Busca_Codigo(AP_MESTRA, "id", "descripcion", "con_meses", "N", xCon)
        ' MUESTRA EN EL STATUS BAR DEL FORMULARIO EL NOMBRE DEL MES TRABAJO ACTUAL
        MDIPrincipal.StatusBar1.Panels(4).Text = "Mes : " + Trim(NomMes)
    End If
    
    If Button.Index = 5 Then
        '1 = muestras el contenido
        '2 = muestra el indice
        'PARA MOSTRAR LA AYUDA... DEL SISTEMA
        HtmlHelp ByVal 0&, Trim(App.Path) + "\ayuda\manual.chm", 0, ByVal 0&
        
    End If
    
    If Button.Index = 7 Then
        ' CIERRA LA CONECCION A LA BASE DE DATOS Y PERMITE SALIR DEL SISTEMA
        xCon.Close
        Set xCon = Nothing
        Unload Me
        End
    End If
End Sub

Private Sub ventas_01_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.IdMenu = 14
    xfrm.Idusuario = xIdUsuario
    xfrm.Clientes xCon, 1
    Set xfrm = Nothing
End Sub

Private Sub ventas_01_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.IdMenu = 15
    xfrm.Idusuario = xIdUsuario
    xfrm.PuntosVenta xCon
    Set xfrm = Nothing
End Sub

Private Sub ventas_01_03_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(4, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT vta_vendedores.id, vta_vendedores.basico, vta_vendedores.comision, UCase([pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] +', '+[pla_empleados]![nom]) AS apenom, " _
        & " vta_vendedores.idper FROM vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":              xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Apellidos y Nombres": xCamposVista(1, 1) = "apenom":         xCamposVista(1, 2) = "4000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Basico":              xCamposVista(2, 1) = "basico":         xCamposVista(2, 2) = "1500":   xCamposVista(2, 3) = "N":    xCamposVista(2, 4) = "D"
    xCamposVista(3, 0) = "Comision":            xCamposVista(3, 1) = "comision":       xCamposVista(3, 2) = "1100":   xCamposVista(3, 3) = "N":    xCamposVista(3, 4) = "D"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Vendedor":       xCampos(1, 1) = "idper":        xCampos(1, 2) = "N":    xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Basico":         xCampos(2, 1) = "basico":       xCampos(2, 2) = "N":    xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Comision":       xCampos(3, 1) = "comision":     xCampos(3, 2) = "N":    xCampos(3, 3) = "1000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "pla_empleados":             xVinculos(0, 1) = "id":                xVinculos(0, 2) = "apepat,nom,id":
    xVinculos(0, 3) = "Apellidos,Nombres,Codigo":  xVinculos(0, 4) = "3000,3000,1500":    xVinculos(0, 5) = "C,C,C":
    xVinculos(0, 6) = "idper":                     xVinculos(0, 7) = "apepat":               xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "id"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "apenom"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 93
    '-----
    
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "vta_vendedores"
    xform.CampoOrdenado = "apenom"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Vendedores"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub ventas_01_04_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.IdMenu = 94
    xfrm.Idusuario = xIdUsuario
    xfrm.MantProdCen xCon
    Set xfrm = Nothing
End Sub

Private Sub ventas_01_05_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_ventas.ventas
    xFun.IdMenu = 222
    xFun.Idusuario = xIdUsuario
    xFun.ManConceptoNC_ND xCon
    Set xFun = Nothing
End Sub

Private Sub ventas_01_06_01_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(4, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_emptra.* From mae_emptra ORDER BY mae_emptra.id"

    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Nº R.U.C.":            xCamposVista(1, 1) = "numruc":         xCamposVista(1, 2) = "1200":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Razón Social":         xCamposVista(2, 1) = "nombre":         xCamposVista(2, 2) = "4000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Direccion":            xCamposVista(3, 1) = "direccion":      xCamposVista(3, 2) = "4000":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Nº R.U.C.":      xCampos(1, 1) = "numruc":       xCampos(1, 2) = "C":    xCampos(1, 3) = "1200"
    xCampos(2, 0) = "Razón Social":    xCampos(2, 1) = "nombre":       xCampos(2, 2) = "C":    xCampos(2, 3) = "5000"
    xCampos(3, 0) = "Direccion":      xCampos(3, 1) = "direccion":    xCampos(3, 2) = "C":    xCampos(3, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "nombre"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 110
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_emptra"
    xform.CampoOrdenado = "nombre"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Empresas de Transporte"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub ventas_01_06_02_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(5, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_chofer.id, mae_chofer.idvehiculo, UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom] AS apenom, " _
        & " mae_chofer.numbre, mae_vehiculo.marca, mae_vehiculo.numpla, mae_chofer.categoria, mae_chofer.idper FROM mae_vehiculo RIGHT JOIN (pla_empleados RIGHT JOIN " _
        & " mae_chofer ON pla_empleados.id = mae_chofer.idper) ON mae_vehiculo.id = mae_chofer.idvehiculo " _
        & " ORDER BY UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom]"
    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Apellido y Nombres":   xCamposVista(0, 1) = "apenom":         xCamposVista(0, 2) = "5000":   xCamposVista(0, 3) = "C":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Nº Brevete":           xCamposVista(1, 1) = "numbre":         xCamposVista(1, 2) = "1200":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Marca":                xCamposVista(2, 1) = "marca":          xCamposVista(2, 2) = "1500":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Nº Placa":             xCamposVista(3, 1) = "numpla":         xCamposVista(3, 2) = "1200":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":          xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Nom. Apellidos":  xCampos(1, 1) = "idper":        xCampos(1, 2) = "N":    xCampos(1, 3) = "3000"
    xCampos(2, 0) = "Nº Brevete":      xCampos(2, 1) = "numbre":       xCampos(2, 2) = "C":    xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Categoria":       xCampos(3, 1) = "categoria":    xCampos(3, 2) = "C":    xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Vehiculo":        xCampos(4, 1) = "idvehiculo":   xCampos(4, 2) = "N":    xCampos(4, 3) = "1500"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "pla_empleados":                            xVinculos(0, 1) = "id":                   xVinculos(0, 2) = "apepat,nom,id":
    xVinculos(0, 3) = "Apell. Pat.,Nombres,Codigo":               xVinculos(0, 4) = "2000,2000,1000":       xVinculos(0, 5) = "C,C,N":
    xVinculos(0, 6) = "idper":                                    xVinculos(0, 7) = "nom":                  xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "apepat"

    xVinculos(1, 0) = "mae_vehiculo":               xVinculos(1, 1) = "id":              xVinculos(1, 2) = "marca,numpla,id":
    xVinculos(1, 3) = "Marca,Nº Placa,Codigo":      xVinculos(1, 4) = "2000,2000,1000":  xVinculos(1, 5) = "C,C,N":
    xVinculos(1, 6) = "idvehiculo":                 xVinculos(1, 7) = "numpla":          xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "marca"

    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "idper"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 111
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_chofer"
    xform.CampoOrdenado = "id"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Choferes"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub ventas_01_06_03_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(3) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_vehiculo.*  From mae_vehiculo ORDER BY mae_vehiculo.marca"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":          xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1500":   xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Marca":           xCamposVista(1, 1) = "marca":          xCamposVista(1, 2) = "2000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Nº Placa":        xCamposVista(2, 1) = "numpla":         xCamposVista(2, 2) = "2000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":        xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Marca":         xCampos(1, 1) = "marca":        xCampos(1, 2) = "C":    xCampos(1, 3) = "2000"
    xCampos(2, 0) = "Nº Placa":      xCampos(2, 1) = "numpla":       xCampos(2, 2) = "C":    xCampos(2, 3) = "2000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":      xVinculos(0, 1) = "":      xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":      xVinculos(0, 4) = "":      xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":      xVinculos(0, 7) = "":      xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""

    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "marca"
    xCampoBusca(1) = "numpla"
    xCampoBusca(2) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 112
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_vehiculo"
    xform.CampoOrdenado = "id"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Vehiculos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub ventas_01_06_04_Click()
    '--Modificado:  10/01/11 Johan Castro
    '               Enviar parametro del Código de usuario y código del menu

    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_mottra.* From mae_mottra ORDER BY mae_mottra.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    '--CAMPOS PARA CONTROLAR EL CONTROL DE ACCESOS AL FORMULARIO
    xform.Idusuario = xIdUsuario
    xform.IdMenu = 113
    '-----
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_mottra"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Motivos de Traslado"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub ventas_03_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.IdMenu = 114
    xfrm.Idusuario = xIdUsuario
    xfrm.LevantarPedidos xCon
    Set xfrm = Nothing
End Sub

Private Sub ventas_03_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.IdMenu = 115
    xfrm.Idusuario = xIdUsuario
    xfrm.MuestraPedidos xCon
    Set xfrm = Nothing
End Sub

Private Sub ventas_04_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.IdMenu = 16
    xfrm.Idusuario = xIdUsuario
    xfrm.Cotizaciones xCon
    Set xfrm = Nothing
End Sub

Private Sub ventas_05_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_pedidos.Pedidos
    xfrm.IdMenu = 224
    xfrm.Idusuario = xIdUsuario
    xfrm.Pedidos xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub ventas_05_02_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_pedidos.Pedidos
    xFun.MostrarCronogramaEntregas xCon
    Set xFun = Nothing
End Sub

Private Sub ventas_05_03_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_pedidos.Pedidos
    xfrm.ReportePedidos xCon
    Set xfrm = Nothing
End Sub

Private Sub ventas_06_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.IdMenu = 17
    xfrm.Idusuario = xIdUsuario
    Dim sss As New ADODB.Recordset
    RST_Busq sss, "select * from mae_cliente", xCon
    
    xfrm.GuiasRemision xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub ventas_06_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.IdMenu = 18
    xfrm.Idusuario = xIdUsuario
    xfrm.ventas xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub ventas_06_03_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.IdMenu = 181
    xfrm.Idusuario = xIdUsuario
    xfrm.LiqGasDebito xCon, AP_MESTRA
    Set xfrm = Nothing
End Sub

Private Sub ventas_08_01_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.ReporteVentas xCon
    Set xfrm = Nothing
End Sub

Private Sub ventas_08_02_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.ConDevoluciones xCon
    Set xfrm = Nothing
End Sub
