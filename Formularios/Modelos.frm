VERSION 5.00
Begin VB.Form Form28 
   Caption         =   "Modelos"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form28"
   ScaleHeight     =   6585
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   3420
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   2400
      Width           =   4095
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   3885
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   3885
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   3885
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   3885
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   3885
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   3885
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   3885
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3420
      TabIndex        =   1
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Marca"
      Height          =   255
      Left            =   1260
      TabIndex        =   16
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   1260
      TabIndex        =   15
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   5100
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String

Sub CargarTablas()
    Tabla = "Modelos"
    campoClave = "CodModelo"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    Dim codMarca As Integer
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Text1.Text, Conexion
    codMarca = Val(Datos.MostrarCampo("Marcas", "codMarca", "nombre='" & Combo1.Text & "'"))
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "CodModelo," _
        & "Nombre," _
        & "codMarca)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "'," _
        & "" & codMarca & ")"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "Nombre='" & Text2.Text & "'," _
            & "codMarca=" & codMarca & " " _
            & "Where CodModelo=" & Text1.Text & ""
    End If
    SqlActualizacion = iSql
    
End Function

Sub Bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Combo1.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
    Combo1.Text = ""
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Combo1.Text = MostrarCampo("Marcas", "nombre", "codMarca=" & rs(2))
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
    Datos.llenarCombo "select nombre from marcas", Combo1
End Sub





