VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form34 
   Caption         =   "Trabajos Realizados por Empleados"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form34"
   ScaleHeight     =   4530
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   1080
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ver Reporte"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   3600
      Width           =   1455
   End
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   6480
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   8
      Top             =   2520
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   41025537
      CurrentDate     =   40288
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   41025537
      CurrentDate     =   40288
   End
   Begin VB.Label Label3 
      Caption         =   "Empleado"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim Formulas() As String
    Dim filtro As String
    filtro = "{Entregas.Fecha}>=#" & Format(DTPicker1.value, "mm/dd/yyyy") & "#"
    filtro = filtro & "AND {Entregas.Fecha}<=#" & Format(DTPicker2.value, "mm/dd/yyyy") & "#"
    If Combo1.Text <> "" Then
        filtro = filtro & "AND {Entregas.codEmpleado}=" & Datos.MostrarCampo("Empleados", "codEmpleado", "nombre='" & Combo1.Text & "'") & ""
    End If
    Datos.CargarReporte filtro, App.Path & "\reportes\trabajos.rpt", Formulas()
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    Datos.llenarCombo "select nombre from Empleados", Combo1
    DTPicker1.value = Date
    DTPicker2.value = Date
End Sub
