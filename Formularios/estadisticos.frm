VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form53 
   Caption         =   "Reportes Estadisticos"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form53"
   ScaleHeight     =   4740
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Cardex"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   3000
      Width           =   2235
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Mejores 10 Clientes"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   3000
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   8340
      TabIndex        =   5
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresos Mensuales"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   1995
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   129761281
      CurrentDate     =   40288
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   129761281
      CurrentDate     =   40288
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form53"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Formulas() As String
Function filtro() As String
Dim f As String
Dim fechaDesde As String
Dim fechaHasta As String
    fechaDesde = "#" & Format(DTPicker1.value, "mm/dd/yyyy") & "#"
    fechaHasta = "#" & Format(DTPicker2.value, "mm/dd/yyyy") & "#"
    f = "{pedidoenc.fecha}>=" & fechaDesde & " and {pedidoenc.fecha}<=" & fechaHasta
    filtro = f

End Function
Function filtroCardex() As String
Dim f As String
Dim fechaDesde As String
Dim fechaHasta As String
    fechaDesde = "#" & Format(DTPicker1.value, "mm/dd/yyyy") & "#"
    fechaHasta = "#" & Format(DTPicker2.value, "mm/dd/yyyy") & "#"
    f = "{compraventa.fecha}>=" & fechaDesde & " and {compraventa.fecha}<=" & fechaHasta
    filtroCardex = f

End Function

Private Sub Command1_Click()
Dim f As String
Dim Archivo As String
    f = filtro
    Archivo = App.Path & "\reportes\rptVentasMoneda.rpt"
    Datos.CargarReporte f, Archivo, Formulas
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
Dim f As String
Dim Archivo As String
    f = filtro & " AND {clientes.codcliente} <> " & 615
    Archivo = App.Path & "\reportes\rptMejoresClientes.rpt"
    Datos.CargarReporte f, Archivo, Formulas

End Sub

Private Sub Command4_Click()
Dim f As String
Dim Archivo As String
    f = filtroCardex
    Archivo = App.Path & "\reportes\rptMovimientos.rpt"
    Datos.CargarReporte f, Archivo, Formulas

End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    DTPicker1.value = Now
    DTPicker2.value = Now

End Sub
