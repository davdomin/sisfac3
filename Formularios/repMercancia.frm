VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form21 
   Caption         =   "Reponer mercancia"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form21"
   ScaleHeight     =   3885
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   1920
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1920
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reporte"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   110100481
      CurrentDate     =   40288
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   110100481
      CurrentDate     =   40288
   End
   Begin VB.Label Label3 
      Caption         =   "Producto :"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim f As String
    Dim fechaDesde As String
    Dim fechaHasta As String
    Dim Formulas() As String
    Dim Archivo As String
        
    fechaDesde = "#" & Format(DTPicker1.value, "mm/dd/yyyy") & "#"
    fechaHasta = "#" & Format(DTPicker2.value, "mm/dd/yyyy") & "#"
    f = "{ReponerPro.fecha}>=" & fechaDesde & " and {ReponerPro.fecha}<=" & fechaHasta
    If Combo1.Text <> "" Then
        f = f + " and {Productos.descripcion}='" & Combo1.Text & "'"
    Else
        f = ""
    End If

    Archivo = App.Path & "\reportes\reposicion.rpt"
    Datos.CargarReporte f, Archivo, Formulas
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    DTPicker1.value = Date
    DTPicker2.value = Date
    Datos.llenarCombo "select descripcion from Productos where tipo <>'Servicio' order by descripcion", Combo1
End Sub
