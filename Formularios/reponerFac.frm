VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form49 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Orden de Compra"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   12840
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form49"
   ScaleHeight     =   9465
   ScaleWidth      =   12840
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   7020
      Picture         =   "reponerFac.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   9060
      Width           =   435
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   6540
      Picture         =   "reponerFac.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   9060
      Width           =   435
   End
   Begin VB.CommandButton btnPrev 
      Height          =   375
      Left            =   5820
      Picture         =   "reponerFac.frx":0F5C
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   9060
      Width           =   435
   End
   Begin VB.CommandButton btnFirst 
      Height          =   375
      Left            =   5340
      Picture         =   "reponerFac.frx":170A
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   9060
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anular"
      Height          =   615
      Left            =   7320
      Picture         =   "reponerFac.frx":1EB8
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   8340
      Width           =   1275
   End
   Begin VB.CommandButton btnCargar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cargar"
      Height          =   615
      Left            =   4740
      Picture         =   "reponerFac.frx":2666
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   8340
      Width           =   1695
   End
   Begin VB.CommandButton btnNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   7260
      TabIndex        =   24
      Top             =   5040
      Width           =   1155
   End
   Begin MSComCtl2.DTPicker dtFechaCompra 
      Height          =   315
      Left            =   3660
      TabIndex        =   18
      Top             =   2220
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Format          =   129761281
      CurrentDate     =   41889
   End
   Begin VB.TextBox txtNumControl 
      Height          =   285
      Left            =   720
      TabIndex        =   17
      Top             =   2220
      Width           =   2295
   End
   Begin VB.CommandButton btnSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   615
      Left            =   10620
      Picture         =   "reponerFac.frx":2E14
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8340
      Width           =   1695
   End
   Begin VB.CommandButton btnCancelar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   8640
      Picture         =   "reponerFac.frx":35C2
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8340
      Width           =   1335
   End
   Begin VB.CommandButton btnGuardar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Guardar"
      Height          =   615
      Left            =   3000
      Picture         =   "reponerFac.frx":3D70
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8340
      Width           =   1695
   End
   Begin VB.CommandButton btnNueva 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo"
      Height          =   615
      Left            =   780
      Picture         =   "reponerFac.frx":451E
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8340
      Width           =   1635
   End
   Begin VB.TextBox txtMontoTotal 
      Height          =   375
      Left            =   11100
      TabIndex        =   33
      Text            =   "Text12"
      Top             =   7920
      Width           =   1635
   End
   Begin VB.ListBox lstTotal 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   11100
      TabIndex        =   32
      Top             =   5400
      Width           =   1635
   End
   Begin VB.ListBox lstMonto 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   9600
      TabIndex        =   30
      Top             =   5400
      Width           =   1515
   End
   Begin VB.ListBox lstCantidad 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   8460
      TabIndex        =   31
      Top             =   5400
      Width           =   1155
   End
   Begin VB.ListBox lstDescripcion 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   720
      TabIndex        =   29
      Top             =   5400
      Width           =   7755
   End
   Begin VB.TextBox txtTotal 
      Height          =   375
      Left            =   11100
      TabIndex        =   28
      Top             =   5040
      Width           =   1635
   End
   Begin VB.TextBox txtMonto 
      Height          =   375
      Left            =   9600
      TabIndex        =   27
      Top             =   5040
      Width           =   1515
   End
   Begin VB.TextBox txtCantidad 
      Height          =   375
      Left            =   8460
      TabIndex        =   26
      Top             =   5040
      Width           =   1155
   End
   Begin VB.TextBox txtProducto 
      Height          =   375
      Left            =   720
      TabIndex        =   25
      Top             =   5040
      Width           =   6555
   End
   Begin VB.TextBox txtTelefono 
      Height          =   285
      Left            =   9180
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   1620
      Width           =   3195
   End
   Begin VB.TextBox txtDireccion 
      Height          =   285
      Left            =   600
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   1620
      Width           =   8475
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   8340
      TabIndex        =   10
      Text            =   "Text5"
      Top             =   1080
      Width           =   4035
   End
   Begin VB.CommandButton btnBuscar 
      Caption         =   "..."
      Height          =   315
      Left            =   7680
      TabIndex        =   9
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtRazonSocial 
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1080
      Width           =   6975
   End
   Begin VB.TextBox txtHora 
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Top             =   420
      Width           =   1695
   End
   Begin VB.TextBox txtFecha 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   420
      Width           =   1695
   End
   Begin VB.TextBox txtReposicion 
      Height          =   285
      Left            =   8400
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   360
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid gridProductos 
      Height          =   2235
      Left            =   720
      TabIndex        =   19
      Top             =   2580
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   3942
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Catalogo de Productos"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   4440
      Visible         =   0   'False
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblAnulada 
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000703C0&
      Height          =   675
      Left            =   4860
      TabIndex        =   45
      Top             =   180
      Width           =   3075
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Compra:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   1980
      Width           =   2775
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Control Externo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   540
      TabIndex        =   15
      Top             =   1980
      Width           =   2775
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   375
      Left            =   10380
      TabIndex        =   34
      Top             =   7980
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   23
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9660
      TabIndex        =   22
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   21
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto a Reponer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   780
      TabIndex        =   20
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   12
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   540
      TabIndex        =   11
      Top             =   1380
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Razon Social del Proveedor:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   540
      TabIndex        =   6
      Top             =   780
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/Codigo del Proveedor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8340
      TabIndex        =   7
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   2
      Top             =   180
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Orden:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8340
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form49"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As Double
Dim contAray As Integer
Dim currentOrden As Long
Dim ArrayOrdenes() As Long
Private Function getTotal() As Double
    Dim i As Integer
    Dim total As Double
    total = 0
    For i = 0 To lstTotal.ListCount - 1
        total = total + Val(lstTotal.List(i))
    Next i
    getTotal = total
End Function
Sub BuscarProveedor()
    Dim filtro As String
    filtro = "RazonSocial='" & txtRazonSocial.Text & "'"
    If Existe("proveedores", filtro) Then
        txtCodigo.Text = Datos.MostrarCampo("proveedores", "rif", filtro)
        txtDireccion.Text = Datos.MostrarCampo("proveedores", "direccion", filtro)
        txtTelefono.Text = Datos.MostrarCampo("proveedores", "telefono", filtro)
    Else
        txtCodigo.Text = ""
        txtDireccion.Text = ""
        txtTelefono.Text = ""
    End If
End Sub

Sub calcularMontoTotal()
    txtTotal.Text = Val(txtCantidad.Text) * Val(txtMonto.Text)
End Sub

Sub mostrarGrilla()
    sqlbase = "select codigo,[descripcion articulo],[stock actual], [Precio de Venta] FROM productosOrdenados2"
    If txtProducto.Text <> "" Then
        sqlbase = sqlbase + " where [descripcion articulo] like '%" & txtProducto.Text & "%'"
    End If
    Adodc1.ConnectionString = Conexion.ConnectionString
    Adodc1.RecordSource = sqlbase
    Adodc1.Refresh
    Set gridProductos.datasource = Adodc1
    gridProductos.Columns(0).Visible = False
    gridProductos.Columns(1).Width = 7000

End Sub

Sub bloquear(st As Boolean)
    txtReposicion.Enabled = st
    txtNumControl.Enabled = st
    dtFechaCompra.Enabled = st
    txtFecha.Enabled = st
    txtHora.Enabled = st
    txtRazonSocial.Enabled = st
    txtCodigo.Enabled = st
    txtDireccion.Enabled = st
    txtTelefono.Enabled = st
    txtProducto.Enabled = st
    txtCantidad.Enabled = st
    txtMonto.Enabled = st
    txtTotal.Enabled = st
    txtMontoTotal.Enabled = st
    lstDescripcion.Enabled = st
    lstCantidad.Enabled = st
    lstMonto.Enabled = st
    lstTotal.Enabled = st
    btnBuscar.Enabled = st
    gridProductos.Enabled = st
End Sub

Sub Limpiar()
    total = 0
    lblAnulada.Visible = False
    txtNumControl.Text = ""
    dtFechaCompra.value = Now
    txtReposicion.Text = ""
    txtFecha.Text = Date
    txtHora.Text = Time
    txtRazonSocial.Text = ""
    txtCodigo.Text = ""
    txtDireccion.Text = ""
    txtTelefono.Text = ""
    txtProducto.Text = ""
    txtCantidad.Text = ""
    txtMonto.Text = ""
    txtTotal.Text = ""
    txtMontoTotal.Text = ""
    lstDescripcion.Clear
    lstCantidad.Clear
    lstMonto.Clear
    lstTotal.Clear
    mostrarGrilla
End Sub

Sub cancelar()
    Limpiar
    bloquear False
End Sub

Sub guardarProveedor()
    Dim filtro As String
    filtro = "RazonSocial='" & txtRazonSocial.Text & "'"
    If Not Existe("proveedores", filtro) Then
        Dim codProveedor As Integer
        codProveedor = Val(Datos.generarCodigo("Proveedores", "codProveedor"))
        
        sql = "Insert Into Proveedores (codProveedor,RazonSocial,rif,Direccion,Telefono) values(" _
        & "" & codProveedor & "," _
        & "'" & txtRazonSocial.Text & "'," _
        & "'" & txtCodigo.Text & "'," _
        & "'" & txtDireccion.Text & "'," _
        & "'" & txtTelefono.Text & "')"
        Conexion.Execute sql
    End If
End Sub

Sub Guardar()
    Dim sql As String
    Dim codProveedor As Integer
    
    If txtRazonSocial.Text = "" Then
        MsgBox "Debe seleccionar un proveedor"
        Exit Sub
    End If
    guardarProveedor
    codProveedor = Val(Datos.MostrarCampo("Proveedores", "codProveedor", "razonSocial='" & txtRazonSocial.Text & "'"))
    
    sql = "insert into reposicionEnc (CodRepo,Fecha,hora,codProveedor,numero_referencia,fecha_compra,total) values(" _
    & "" & Val(txtReposicion.Text) & "," _
    & "'" & Date & "'," _
    & "'" & Time & "'," _
    & "" & codProveedor & "," _
    & "'" & txtNumControl.Text & "'," _
    & "'" & dtFechaCompra.value & "'," _
    & "" & getTotal() & ")"
    Conexion.Execute sql
    
    For i = 0 To lstDescripcion.ListCount - 1
        Dim codReponerPro As Integer
        Dim codProducto As Integer
        codReponerPro = Datos.generarCodigo("reponerPro", "codReponerPro")
        codProducto = Datos.MostrarCampo("productos", "codProducto", "descripcion='" & lstDescripcion.List(i) & "'")
        sql = "insert into reponerPro(codReponerPro,fecha,CodProducto,Cantidad,CodProveedor,precioRep,CodRepo) values(" _
        & "" & codReponerPro & "," _
        & "'" & Date & "'," _
        & "" & codProducto & "," _
        & "" & Val(lstCantidad.List(i)) & "," _
        & "" & codProveedor & "," _
        & "" & Val(lstMonto.List(i)) & "," _
        & "" & Val(txtReposicion.Text) & ")"
        Conexion.Execute sql
        
        sql = "UPDATE productos SET " _
        & "costo=" & Val(lstMonto.List(i)) & "," _
        & "Stock=Stock+" & Val(lstCantidad.List(i)) & " " _
        & "where codProducto=" & codProducto & ""
        Conexion.Execute sql
    Next
End Sub

Sub LimpiarProductos()
    txtProducto.Text = ""
    txtCantidad.Text = ""
    txtMonto.Text = ""
    txtProducto.SetFocus
End Sub

Sub Agregar()
    If Not Existe("Productos", "descripcion='" & txtProducto.Text & "'") Then Exit Sub
    If Val(txtCantidad.Text) <> 0 And Val(txtTotal.Text) <> 0 And Len(txtProducto.Text) <> 0 Then
        lstDescripcion.AddItem txtProducto.Text
        lstCantidad.AddItem txtCantidad.Text
        lstMonto.AddItem txtMonto.Text
        lstTotal.AddItem Val(txtTotal.Text)
        LimpiarProductos
        txtMontoTotal.Text = Format(getTotal(), "currency")
    End If
End Sub

Private Sub loadArray()
    Dim tReposiciones As New ADODB.Recordset
    tReposiciones.Open "SELECT codRepo FROM reposicionEnc", Conexion
    contAray = 0
    While Not tReposiciones.EOF
        ReDim Preserve ArrayOrdenes(contAray)
        ArrayOrdenes(contAray) = tReposiciones("codRepo").value
        tReposiciones.MoveNext
        contAray = contAray + 1
    Wend
End Sub

Private Sub btnCargar_Click()
    frmRepoLoad.Show vbModal
    If frmRepoLoad.numero = -1 Then Exit Sub
    Mostrar frmRepoLoad.numero
End Sub

Private Sub btnFirst_Click()
    currentOrden = 0
    Mostrar ArrayOrdenes(currentOrden)
End Sub

Private Sub btnNuevo_Click()
    Form5.Show vbModal
    mostrarGrilla
End Sub

Private Sub btnBuscar_Click()
    Catalogo.sql = "SELECT *FROM PROVEEDORES"
    Catalogo.Show vbModal
End Sub

Private Sub btnNueva_Click()
    cancelar
    bloquear True
    txtReposicion.Text = Datos.generarCodigo("reposicionEnc", "codRepo")
    txtRazonSocial.SetFocus
End Sub

Private Sub btnGuardar_Click()
    Guardar
    cancelar
    loadArray
    mostrarGrilla
End Sub

Private Sub btnCancelar_Click()
    cancelar
End Sub

Private Sub btnPrev_Click()
    currentOrden = currentOrden - 1
    If currentOrden = -1 Then
        currentOrden = 0
        Exit Sub
    End If
    
    Mostrar ArrayOrdenes(currentOrden)
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub
Private Sub cleanEncabezado()
    txtReposicion.Text = ""
    txtFecha.Text = ""
    txtHora.Text = ""
    txtRazonSocial.Text = ""
    txtDireccion.Text = ""
    txtTelefono.Text = ""
    txtNumControl.Text = ""
    txtTotal.Text = ""
End Sub
Private Sub Mostrar(ByVal codOrden As Long)
    Dim tOrden As New ADODB.Recordset
    cleanEncabezado
    tOrden.Open "SELECT Proveedores.razonsocial,Proveedores.direccion,Proveedores.telefono, " _
    & "reposicionEnc.anulado,reposicionEnc.total, reposicionEnc.codrepo,reposicionEnc.Fecha, reposicionEnc.Hora, reposicionEnc.total,reposicionEnc.numero_referencia,fecha_compra " _
    & "FROM reposicionEnc INNER JOIN Proveedores ON reposicionEnc.codProveedor=Proveedores.codProveedor WHERE codRepo=" & codOrden, Conexion
    If tOrden.EOF Then Exit Sub
    On Error Resume Next
    txtReposicion.Text = tOrden("codRepo").value
    txtFecha.Text = tOrden("fecha").value
    txtHora.Text = tOrden("hora").value
    txtRazonSocial.Text = tOrden("razonsocial").value
    txtDireccion.Text = tOrden("direccion").value
    txtTelefono.Text = tOrden("fecha").value
    txtNumControl.Text = tOrden("numero_referencia").value
    txtMontoTotal.Text = Format(getTotal(), "currency")
    dtFechaCompra.value = tOrden("fecha_compra").value
    lblAnulada.Visible = tOrden("anulado").value
    On Error GoTo 0
    mostrarDetalle codOrden
    loadArray
    mostrarGrilla
End Sub

Private Sub mostrarDetalle(ByVal codOrden As Long)
    Dim tOrdenDet As New ADODB.Recordset
    tOrdenDet.Open "SELECT Productos.Descripcion," _
    & "ReponerPro.Cantidad, ReponerPro.Cantidad, ReponerPro.PrecioRep, ReponerPro.precioRep * ReponerPro.Cantidad as Total " _
    & "FROM reponerPro INNER JOIN Productos ON reponerPro.codProducto=Productos.codProducto WHERE codRepo=" & codOrden, Conexion
    lstDescripcion.Clear
    lstCantidad.Clear
    lstMonto.Clear
    lstTotal.Clear
    If tOrdenDet.EOF Then Exit Sub
    Do While Not tOrdenDet.EOF
        lstDescripcion.AddItem tOrdenDet("descripcion").value
        lstCantidad.AddItem tOrdenDet("cantidad").value
        lstMonto.AddItem tOrdenDet("PrecioRep").value
        lstTotal.AddItem tOrdenDet("total").value
        tOrdenDet.MoveNext
    Loop
End Sub

Private Sub Command1_Click()
    Dim codRepo As Long
    Dim vbresult As VbMsgBoxResult
    Dim tProductos As New ADODB.Recordset
    If lblAnulada.Visible Then Exit Sub
    If MsgBox("¿Esta seguro que desea anular esta Orden de Compra?", vbYesNo) = vbNo Then Exit Sub
    codRepo = Val(txtReposicion.Text)
    If codRepo = 0 Then Exit Sub
    tProductos.Open "SELECT codProducto, Cantidad FROM [reponerPro] WHERE codRepo = " & codRepo, Conexion
    Do While Not tProductos.EOF
        Conexion.Execute "UPDATE [Productos] SET Stock = Stock - " & tProductos("cantidad").value & "  " _
        & " WHERE codProducto =" & tProductos("codProducto").value
        tProductos.MoveNext
    Loop
    Conexion.Execute "UPDATE [reposicionEnc] SET anulado = true WHERE codRepo = " & codRepo
    vbresult = MsgBox("Orden de compra, fue anulada exitosamente!", vbInformation)
    Mostrar codRepo
End Sub

Private Sub Command2_Click()
    currentOrden = currentOrden + 1
    If currentOrden >= contAray Then
        currentOrden = contAray
        Exit Sub
    End If
    Mostrar ArrayOrdenes(currentOrden)
End Sub

Private Sub Command3_Click()
    currentOrden = contAray - 1
    Mostrar ArrayOrdenes(currentOrden)
End Sub

Private Sub gridProductos_DblClick()
    txtProducto.Text = Adodc1.Recordset(1)
    txtCantidad.SetFocus
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    cancelar
    currentOrden = 1
    loadArray
End Sub

Private Sub lstDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
Dim index As Byte
    If KeyCode <> 46 Then Exit Sub
    index = lstDescripcion.ListIndex
    lstDescripcion.RemoveItem index
    lstCantidad.RemoveItem index
    lstMonto.RemoveItem index
    lstTotal.RemoveItem index
            
    
End Sub

Private Sub txtMonto_Change()
    calcularMontoTotal
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Agregar
    End If
End Sub

Private Sub txtRazonSocial_Validate(Cancel As Boolean)
    BuscarProveedor
End Sub

Private Sub txtProducto_Change()
    mostrarGrilla
    Datos.AutoCompletar_TextBox txtProducto
End Sub

Private Sub txtProducto_GotFocus()
    Datos.CargarValores "select *from productos order by descripcion"
End Sub

Private Sub txtProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyBack, vbKeyDelete
            Select Case Len(txtProducto.Text)
                Case Is <> 0
                    bKeyBack = True
            End Select
    End Select
End Sub

Private Sub txtRazonSocial_Change()
    Datos.AutoCompletar_TextBox txtRazonSocial
End Sub

Private Sub txtRazonSocial_GotFocus()
    Datos.CargarValores "select CODPROVEEDOR,RAZONSOCIAL from PROVEEDORES order by RAZONSOCIAL"
End Sub

Private Sub txtRazonSocial_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyBack, vbKeyDelete
            Select Case Len(txtRazonSocial.Text)
                Case Is <> 0
                    bKeyBack = True
            End Select
    End Select
    If KeyCode = 114 Then
        btnBuscar_Click
    End If
End Sub

Private Sub txtCantidad_Change()
    calcularMontoTotal
End Sub

Private Sub txtCantidad_GotFocus()
    txtCantidad.Text = "1"
    txtCantidad.SelStart = 0
    txtCantidad.SelLength = 1
End Sub

