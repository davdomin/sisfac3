VERSION 5.00
Begin VB.Form Form32 
   Caption         =   "Pagos a Empleados"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form32"
   ScaleHeight     =   7725
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "&Cargar Pagos Pendientes"
      Height          =   495
      Left            =   4800
      TabIndex        =   22
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   21
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   7155
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   20
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   7155
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   19
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   7155
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   18
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   7155
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   600
      TabIndex        =   11
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   3420
      Left            =   5280
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   3420
      Left            =   720
      TabIndex        =   7
      Top             =   2040
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   4320
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   7200
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Total a Pagar"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Montos por Cancelar Segun Empleados"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Hora:"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Control"
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String
Sub CargarTablas()
    Tabla = "pagosEmp"
    campoClave = "CodpagosEmp"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    Dim Monto As Double
    Dim codEmpleado As Integer
    Dim codPagosEmp As Integer
    
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Text1.Text, Conexion
    Dim i As Integer
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "CodpagosEmp," _
        & "fecha," _
        & "hora," _
        & "total)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "'," _
        & "'" & Text3.Text & "'," _
        & "" & Label6.Caption & ")"
        
        For i = 0 To List1.ListCount - 1
            codPagosEmp = Val(Text1.Text)
            codEmpleado = Val(Datos.MostrarCampo("Empleados", "codEmpleado", "Nombre='" & List1.List(i) & "'"))
            Monto = Val(List2.List(i))
            Conexion.Execute "insert into pagosEmpDet (codPagosEmp,codEmpleado,Monto) " _
            & "values(" & codPagosEmp & "," & codEmpleado & "," & Monto & ")"
            Conexion.Execute "Update Entregas set codPagosEmp= " & codPagosEmp & " where codPagosEmp=0 and codEmpleado=" & codEmpleado
        Next
        
        Conexion.Execute "Insert into gastos (codGasto,codProveedor,fecha,hora,concepto,total,pagado) values(" _
        & "" & Datos.generarCodigo("Gastos", "codGasto") & "," _
        & "0," _
        & "'" & Text2.Text & "'," _
        & "'" & Text3.Text & "'," _
        & "'Nomina'," _
        & "" & Label6.Caption & "," _
        & "" & Label6.Caption & ")"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "Fecha='" & Text2.Text & "'," _
            & "hora='" & Text3.Text & "'," _
            & "Total=" & Label6.Caption & "," _
            & "Where CodpagosEmp=" & Text1.Text & ""
    End If
    SqlActualizacion = iSql
    
End Function

Sub Bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
    Label5.Enabled = es
    Label6.Enabled = es
    List1.Enabled = es
    List2.Enabled = es
    Command12.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = Date
    Text3.Text = Time
    List1.Clear
    List2.Clear
    Label6.Caption = ""
End Sub
Sub MostrarDetalle()
    Dim rs As New ADODB.Recordset
    rs.Open "select *from pagosEmpDet where codpagosEmp=" & Val(Text1.Text), Conexion
    While Not rs.EOF
        List1.AddItem Datos.MostrarCampo("Empleados", "nombre", "codEmpleado=" & rs(1))
        List2.AddItem rs(2)
        rs.MoveNext
    Wend
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Text3.Text = rs(2)
    Label6.Caption = rs(3)
    MostrarDetalle
End Sub

Private Sub Command1_Click()
    Formularios.Nuevo Me
End Sub
Sub mostrarPendientes()
Dim Total As Double
    Dim rs As New ADODB.Recordset
    rs.Open "select nombre,sum(monto) from empleados,entregas where empleados.codempleado=entregas.codempleado and entregas.codPagosEmp=0 group by nombre", Conexion
    List1.Clear
    List2.Clear
    Total = 0
    While Not rs.EOF
        List1.AddItem rs(0)
        List2.AddItem rs(1)
        Total = Total + rs(1)
        rs.MoveNext
    Wend
    Label6.Caption = Total
    
End Sub

Private Sub Command12_Click()
    mostrarPendientes
End Sub

Private Sub Command2_Click()
    Formularios.Guardar Me
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
    Formularios.Botones Me, 1
    Bloquear False
    CargarTablas
End Sub




Sub Seleccionar(Indice As Integer)
    List1.Selected(Indice) = True
    List2.Selected(Indice) = True
End Sub
Private Sub List1_Click()
    Seleccionar List1.ListIndex
End Sub
