VERSION 5.00
Begin VB.Form Form20 
   Caption         =   "Reponer Productos"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11085
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form20"
   ScaleHeight     =   7245
   ScaleWidth      =   11085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command14 
      Caption         =   "Agregar Producto"
      Height          =   375
      Left            =   9000
      TabIndex        =   34
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   1440
      TabIndex        =   29
      Top             =   4440
      Width           =   4215
      Begin VB.TextBox Text4 
         Height          =   360
         Left            =   1245
         TabIndex        =   30
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Costo:"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reposicion de Productos"
         Height          =   255
         Left            =   3360
         TabIndex        =   32
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ultimo:"
         Height          =   255
         Left            =   2160
         TabIndex        =   31
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "reposicionPro.frx":0000
      Top             =   3600
      Width           =   7455
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Nuevo Proveedor"
      Height          =   375
      Left            =   9000
      TabIndex        =   27
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Left            =   1320
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1680
      Width           =   7575
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Buscar Producto"
      Height          =   375
      Left            =   9000
      TabIndex        =   25
      Top             =   2160
      Width           =   1935
   End
   Begin VB.ComboBox Combo4 
      Height          =   360
      Left            =   1320
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2160
      Width           =   7575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9765
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   1320
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   2640
      Width           =   7575
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   2565
      TabIndex        =   5
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   6840
      TabIndex        =   10
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Observacion:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Categoria"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Reposicion"
      Height          =   255
      Left            =   7365
      TabIndex        =   23
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto :"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reposicion de Productos"
      Height          =   255
      Left            =   4725
      TabIndex        =   18
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual:"
      Height          =   255
      Left            =   3525
      TabIndex        =   19
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String
Sub MostrarStock()
    Dim rs As New ADODB.Recordset
    Dim iSql  As String
    Dim codProducto As Integer
    codProducto = Val(Datos.MostrarCampo("Productos", "codProducto", "Descripcion='" & Combo1.Text & "'"))
    iSql = "select *from Productos where codProducto=" & codProducto
    rs.Open iSql, Conexion
    If rs.EOF Then
        Label6.Caption = ""
        Label9.Caption = ""
    Else
        On Error Resume Next
        Label6.Caption = rs("stock")
        Label9.Caption = rs("costo")
        Text4.Text = rs("costo")
    End If
End Sub
Sub datosActualizacion()
    If NivelEntro = "Administrador" Or NivelEntro = "Cajero" Then
        codProducto = Val(Datos.MostrarCampo("Productos", "codProducto", "Descripcion='" & Combo1.Text & "'"))
        codProveedor = Val(Datos.MostrarCampo("Proveedores", "codproveedor", "razonSocial='" & Combo2.Text & "'"))
        
        Conexion.Execute "update productos set costo=" & Val(Text4.Text) & " where codProducto=" & codProducto
        Conexion.Execute "update reponerPro set precioRep=" & Val(Text4.Text) & " where codReponerPro=" & Text1.Text & ""
    End If

End Sub

Sub Reponer_Productos()
    Dim codProducto As Integer

    Dim Cantidad As Double
    codProducto = Val(Datos.MostrarCampo("Productos", "codProducto", "Descripcion='" & Combo1.Text & "'"))
    codProveedor = Val(Datos.MostrarCampo("Proveedores", "codproveedor", "razonSocial='" & Combo2.Text & "'"))
    Cantidad = Val(Text3.Text)
    Conexion.Execute "Update Productos set stock=stock+" & Cantidad & " where codProducto=" & codProducto
    Conexion.Execute "Update proveedores set observacion='" & Text5.Text & "' where codProveedor=" & codProveedor
    If Combo2.Text <> "" Then
        Conexion.Execute "update reponerPro set codProveedor=" & codProveedor & " where codReponerPro=" & Text1.Text & ""
    End If
    datosActualizacion
End Sub

Sub CargarTablas()
    Tabla = "ReponerPro"
    campoClave = "CodReponerPro"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Text1.Text, Conexion
    If rs.EOF Then
        codProveedor = Val(Datos.MostrarCampo("Proveedores", "codproveedor", "razonSocial='" & Combo2.Text & "'"))
        iSql = "insert into [" & Tabla & "] (" _
        & "CodReponerPro," _
        & "Fecha," _
        & "codProveedor," _
        & "codProducto," _
        & "Cantidad)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "'," _
        & "" & codProveedor & "," _
        & "" & Datos.MostrarCampo("Productos", "codProducto", "Descripcion='" & Combo1.Text & "'") & "," _
        & "'" & Text3.Text & "')"
        Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "fecha='" & Text2.Text & "'" _
            & "Where CodReponerPro=" & Text1.Text & ""
            Reponer_Productos

        datosActualizacion
    End If
    SqlActualizacion = iSql
    
End Function

Sub bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Text4.Enabled = es
    Text5.Enabled = es
    Combo1.Enabled = es
    Combo2.Enabled = es
    Combo4.Enabled = es
    
    Label1.Enabled = es
    Label2.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
    Label5.Enabled = es
    Label7.Enabled = es
    Label8.Enabled = es
    Label9.Caption = ""
    Label10.Enabled = es
    Label11.Enabled = es

End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = Date
    Text3.Text = ""
    Text5.Text = ""
    
    
    Combo1.Text = ""
    Combo2.Text = ""
    Combo4.Text = ""
    Label6.Caption = ""
    Label9.Caption = ""
    'Label9.Caption = ""

End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Combo1.Text = rs(2)
    Text3.Text = rs(3)
End Sub
Private Sub Combo1_Click()
    MostrarStock
End Sub

Private Sub Combo2_Click()
On Error Resume Next
    Text5.Text = Datos.MostrarCampo("Proveedores", "Observacion", "razonSocial='" & Combo2.Text & "'")
    
    
End Sub

Private Sub Combo4_Change()
ProductosPorCategoria

End Sub

Private Sub Combo4_Validate(Cancel As Boolean)
ProductosPorCategoria
End Sub

Private Sub Command1_Click()
    Formularios.Nuevo Me
    Combo1.SetFocus
End Sub

Private Sub Command12_Click()
    MostrarCatalogo "select *from [productosOrdenados]"
    Combo1.Text = Datos.MostrarCampo("Productos", "Descripcion", "codProducto=" & Catalogo.Resultado & "")
    MostrarStock
    
End Sub

Private Sub Command13_Click()
    Form23.Show vbModal
    Datos.llenarCombo "select razonSocial from Proveedores  order by razonSocial", Combo2
    
End Sub

Private Sub Command14_Click()
    Form5.Show vbModal
End Sub

Private Sub Command2_Click()
    If Combo2.Text = "" Then Exit Sub
    If Combo1.Text = "" Then Exit Sub
    Formularios.Guardar Me
    Me.Reponer_Productos
    
End Sub
Private Sub Command3_Click()
    Formularios.cancelar Me
End Sub
Private Sub Command4_Click()
    Formularios.modificar Me
End Sub
Private Sub Command5_Click()
    Formularios.Eliminar Me
End Sub
Private Sub Command6_Click()
    MostrarCatalogo "select *from [" & Tabla & "]"
    Text1.Text = Catalogo.Resultado
    Formularios.Buscar Me
End Sub
Private Sub Command7_Click()
    Unload Me
End Sub
Private Sub Command8_Click()
    Formularios.Primero Me
End Sub
Private Sub Command9_Click()
    Formularios.Anterior Me
End Sub
Private Sub Command10_Click()
    Formularios.Siguiente Me
End Sub
Private Sub Command11_Click()
    Formularios.Ultimo Me
End Sub
Sub ProductosPorCategoria()
    If Combo4.Text = "" Then
        Datos.llenarCombo "select descripcion from productos order by descripcion", Combo1
    Else
        Datos.llenarCombo "select productos.descripcion from productos,tipop where productos.codtipop=tipop.codtipop and tipop.descripcion='" & Combo4.Text & "' order by productos.descripcion", Combo1
    End If
End Sub

Private Sub Form_Load()
    Frame1.Visible = NivelEntro = "Administrador"
    Datos.llenarCombo "select descripcion from tipop order by descripcion", Combo4
    Formularios.ColorLabels ColorLetras, Me
    Frame1.BackColor = Me.BackColor
    bloquear False
    CargarTablas
    Limpiar
    Datos.llenarCombo "select descripcion from Productos where tipo <>'Servicio' order by descripcion", Combo1
    Datos.llenarCombo "select razonSocial from Proveedores  order by razonSocial", Combo2
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Text3.Text = Round(Text3.Text, 2)
End Sub


