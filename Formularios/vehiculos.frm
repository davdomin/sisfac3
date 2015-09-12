VERSION 5.00
Begin VB.Form Form29 
   Caption         =   "Vehiculos"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8805
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form29"
   ScaleHeight     =   5295
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   3720
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   3720
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Left            =   3720
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2040
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   3720
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   3720
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   3660
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   3660
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   3660
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   3660
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   3660
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   3660
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   3660
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   4215
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   4215
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   4215
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   4215
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Default         =   -1  'True
      Height          =   225
      Left            =   -960
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Año:"
      Height          =   255
      Left            =   1560
      TabIndex        =   23
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Color:"
      Height          =   255
      Left            =   1560
      TabIndex        =   22
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   5340
      TabIndex        =   21
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Placa:"
      Height          =   255
      Left            =   1500
      TabIndex        =   20
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Marca:"
      Height          =   255
      Left            =   1500
      TabIndex        =   19
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo:"
      Height          =   255
      Left            =   1500
      TabIndex        =   18
      Top             =   2040
      Width           =   2055
   End
End
Attribute VB_Name = "Form29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String

Sub CargarTablas()
    Tabla = "Vehiculos"
    campoClave = "codVehiculo"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    Dim codModelo As Integer
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Val(Text1.Text), Conexion
    codModelo = Datos.MostrarCampo("Modelos", "codModelo", "nombre='" & Combo2.Text & "'")
    If rs.EOF Then
        
        iSql = "insert into [" & Tabla & "] (" _
        & "codVehiculo," _
        & "Placa," _
        & "codModelo," _
        & "Color," _
        & "Año) values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "'," _
        & "" & codModelo & "," _
        & "'" & Text3.Text & "'," _
        & "'" & Text4.Text & "')"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "Placa='" & Text2.Text & "'," _
            & "codModelo=" & Text3.Text & "," _
            & "Color='" & Text4.Text & "'," _
            & "Año='" & Text5.Text & "' " _
            & "Where codVehiculo=" & Text1.Text & ""
    End If
    SqlActualizacion = iSql
    
End Function

Sub Bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Combo1.Enabled = es
    Combo2.Enabled = es
    Text3.Enabled = es
    Text4.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
    Label5.Enabled = es
    Label6.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
    Combo1.Text = ""
    Combo2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
Dim codMarca As Integer
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Combo2.Text = Datos.MostrarCampo("Modelos", "nombre", "codModelo=" & rs(2))
    codMarca = Datos.MostrarCampo("Modelos", "codMarca", "codModelo=" & rs(2))
    Combo1.Text = Datos.MostrarCampo("Marcas", "nombre", "codMarca=" & codMarca)
    Text3.Text = rs(3)
    Text4.Text = rs(4)
End Sub

Private Sub Combo1_Click()
Dim codMarca As Integer
    codMarca = Val(Datos.MostrarCampo("Marcas", "CodMarca", "nombre='" & Combo1.Text & "'"))
    Datos.llenarCombo "Select nombre from modelos where codMarca=" & codMarca, Combo2
End Sub

Private Sub Command1_Click()
    Formularios.Nuevo Me
End Sub

Private Sub Command13_Click()
    SendKeys "{TAB}"
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
'    MostrarCatalogo "select codVehiculo, razonsocial,direccion,telefono ,cedrif from [" & Tabla & "]"
    'Text1.Text = Catalogo.Resultado
    'Formularios.Buscar Me
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
    Datos.llenarCombo "Select nombre from marcas", Combo1
    Datos.llenarCombo "Select nombre from Modelos", Combo2
    
End Sub







