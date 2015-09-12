VERSION 5.00
Begin VB.Form Form27 
   Caption         =   "Marcas"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form27"
   ScaleHeight     =   5955
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5580
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2940
      TabIndex        =   1
      Top             =   2160
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   4125
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   4125
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   4125
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   4125
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   4125
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   4125
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   4125
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   4620
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   780
      TabIndex        =   13
      Top             =   2160
      Width           =   2055
   End
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String

Sub CargarTablas()
    Tabla = "Marcas"
    campoClave = "CodMarca"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Text1.Text, Conexion
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "CodMarca," _
        & "Nombre)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "')"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "Nombre='" & Text2.Text & "'" _
            & "Where CodMarca=" & Text1.Text & ""
    End If
    SqlActualizacion = iSql
    
End Function

Sub Bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
End Sub

Private Sub Command1_Click()
    Formularios.Nuevo Me
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



