VERSION 5.00
Begin VB.Form Form14 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ejecutar Ordenes de Trabajo"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8745
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BackColor       =   &H00CEFFFC&
      Height          =   1185
      Left            =   1200
      TabIndex        =   23
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00CEFFFC&
      Height          =   1185
      Left            =   2280
      TabIndex        =   22
      Top             =   2040
      Width           =   4695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   855
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "ejecutarOrden.frx":0000
      Top             =   3720
      Width           =   8055
   End
   Begin VB.CommandButton Catalogo1 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   1215
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
      Height          =   330
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   1200
      TabIndex        =   25
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   2280
      TabIndex        =   24
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Observacion"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Orden"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Ejecucion"
      Height          =   255
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String
Sub MostrarPedido()
    Dim codPedido As Integer
    Dim NumeroO As Integer
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    Dim rd As New ADODB.Recordset
    
    
    
    NumeroO = Val(Text3.Text)
    iSql = "select *from ordenTEnc where numeroO=" & NumeroO
    rs.Open iSql, Conexion
    If Not rs.EOF Then
        codPedido = rs("codPedido")
    End If
    rs.Close
    
    
    iSql = "select *from pedidoEnc where CodPedido=" & codPedido
    rs.Open iSql, Conexion
    If rs.EOF Then
        Label5.Caption = ""
        List1.Clear
        List2.Clear
    Else
        Label5.Caption = Datos.MostrarCampo("Clientes", "razonsocial", "codCliente=" & rs("codCliente"))
        rd.Open "select *from PedidoDet where CodPedido=" & codPedido, Conexion
        List1.Clear
        List2.Clear
        While Not rd.EOF
            If Datos.MostrarCampo("Productos", "seFabrica", "codProducto=" & rd("codProducto")) = "Si" Then
                List1.AddItem rd("cantidad")
                List2.AddItem Datos.MostrarCampo("productos", "descripcion", "codProducto=" & rd("codProducto"))
            End If
            rd.MoveNext
            
        Wend
        
    End If
    
    
    
End Sub

Sub ActualizarProductos()
    Dim iSql As String
    Dim NumeroO As Integer
    Dim rs As New ADODB.Recordset
    Dim rd As New ADODB.Recordset
    NumeroO = Val(Text3.Text)
    iSql = "select *from ordenTEnc where numeroO=" & NumeroO
    rs.Open iSql, Conexion
    If Not rs.EOF Then
        iSql = "select *from PedidoDet where codPedido=" & rs("codPedido")
        rd.Open iSql, Conexion
        While Not rd.EOF
        'PILAS
            If Datos.MostrarCampo("Productos", "seFabrica", "codProducto=" & rd("codProducto")) = "Si" Then
                iSql = "update Productos set stock=stock+" & rd("cantidad") & " where codProducto=" & rd("codProducto")
                Conexion.Execute iSql
            End If
            rd.MoveNext
            
        Wend
    End If
        
    

End Sub
Sub CambiarStatus()
    Dim iSql As String
    Dim NumeroO As Integer
    Dim rs As New ADODB.Recordset
    NumeroO = Val(Text3.Text)
    iSql = "select *from ordenTEnc where numeroO=" & NumeroO
    rs.Open iSql, Conexion
    If Not rs.EOF Then
        iSql = "update PedidoEnc set status='Terminado' where codPedido=" & rs("codPedido")
        Conexion.Execute iSql
        iSql = "update ordenTEnc set status='Terminado' where numeroO=" & NumeroO
        Conexion.Execute iSql
    End If
End Sub

Sub CargarTablas()
    Tabla = "OrdenEje"
    campoClave = "Codejecucion"
End Sub
Sub MostrarCliente()
    Dim iSql As String
    Dim NumeroO As Integer
    Dim codCliente As Integer
    Dim rs As New ADODB.Recordset
    
    NumeroO = Val(Text3.Text)
    iSql = "select *from ordenTEnc where numeroO=" & NumeroO
    rs.Open iSql, Conexion
    If Not rs.EOF Then
        
        codCliente = Datos.MostrarCampo("PedidoEnc", "CodCliente", "CodPedido=" & rs("codPedido"))
        
        Label5.Caption = Datos.MostrarCampo("Clientes", "razonsocial", "codCliente=" & codCliente)
    Else
        Label5.Caption = ""
    End If
        
    
    
    
    
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Text1.Text, Conexion
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "Codejecucion," _
        & "fecha," _
        & "numeroO," _
        & "observacion)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "" & Date & "," _
        & "" & Text3.Text & "," _
        & "'" & Text4.Text & "')"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "observacion='" & Text4.Text & "' " _
            & "Where Codejecucion=" & Text1.Text & ""
    End If
    SqlActualizacion = iSql
    
End Function

Sub Bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Text4.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
    Label5.Enabled = es
    Label6.Enabled = es
    Catalogo1.Enabled = es
    List1.Enabled = es
    List2.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = Date
    Text3.Text = ""
    Text4.Text = ""
    Label5.Caption = ""
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Text3.Text = rs(2)
    Text4.Text = rs(3)
End Sub

Private Sub Catalogo1_Click()
    MostrarCatalogo "select *from [ordenesPendientes]"
    Text3.Text = Catalogo.Resultado
End Sub

Private Sub Command1_Click()
    Formularios.Nuevo Me
End Sub

Private Sub Command2_Click()
    Formularios.Guardar Me
    CambiarStatus
    ActualizarProductos
End Sub

Private Sub Command3_Click()
    Formularios.Cancelar Me
End Sub

Private Sub Command4_Click()
    Formularios.Modificar Me
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

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    Bloquear False
    CargarTablas
    Limpiar
End Sub

Private Sub Text3_Change()
    MostrarCliente
    MostrarPedido
End Sub
