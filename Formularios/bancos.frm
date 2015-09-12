VERSION 5.00
Begin VB.Form Form24 
   Caption         =   "Bancos"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   LinkTopic       =   "Form24"
   ScaleHeight     =   5310
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   3300
      TabIndex        =   3
      Top             =   3240
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   720
      Left            =   3300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2400
      Width           =   4095
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   4365
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   4365
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   4365
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   4365
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   4365
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   4365
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   4365
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   3300
      TabIndex        =   1
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5820
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
      Height          =   255
      Left            =   1020
      TabIndex        =   18
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      Height          =   255
      Left            =   1020
      TabIndex        =   17
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   1020
      TabIndex        =   16
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   4860
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String

Sub CargarTablas()
    Tabla = "Bancos"
    campoClave = "codBanco"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Val(Text1.Text), Conexion
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "codBanco," _
        & "nombre," _
        & "direccion," _
        & "telefono)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "'," _
        & "'" & Text3.Text & "'," _
        & "'" & Text4.Text & "')"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "nombre='" & Text2.Text & "'," _
            & "direccion='" & Text3.Text & "'," _
            & "telefono='" & Text4.Text & "' " _
            & "Where codBanco=" & Text1.Text & ""
    End If
   ' a = InputBox(iSql, iSql, iSql)
    SqlActualizacion = iSql
    
End Function

Sub bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Text4.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Text3.Text = rs(2)
    Text4.Text = rs(3)
End Sub

Private Sub Command1_Click()
    Formularios.Nuevo Me
End Sub



Private Sub Command2_Click()
    Formularios.Guardar Me
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
    MostrarCatalogo "select codBanco, nombre,direccion,telefono from [" & Tabla & "]"
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
    bloquear False
    CargarTablas
End Sub









