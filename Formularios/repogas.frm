VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form35 
   Caption         =   "Reporte de Gastos"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form35"
   ScaleHeight     =   3750
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ver Reporte"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   5880
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   2280
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   112590849
      CurrentDate     =   40288
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   112590849
      CurrentDate     =   40288
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Form35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Formulas() As String
Dim filtro As String
    filtro = "{Gastos.Fecha}>=#" & Format(DTPicker1.value, "mm/dd/yyyy") & "#"
    filtro = filtro & "AND {Gastos.Fecha}<=#" & Format(DTPicker2.value, "mm/dd/yyyy") & "#"
    Datos.CargarReporte filtro, App.Path & "\reportes\gastos.rpt", Formulas()
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    DTPicker1.value = Date
    DTPicker2.value = Date
End Sub

